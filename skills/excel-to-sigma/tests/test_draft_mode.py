#!/usr/bin/env python3
"""Draft mode for local builds, the real forecast parity check, and audit exit codes."""
from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))

from migration_gate import MigrationGateError, gate_source, plan_review_issues  # noqa: E402


def load_script(filename, name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


FORECAST = load_script("build-forecast-model.py", "draft_mode_forecast")
RESEARCH = load_script("build-research-model.py", "draft_mode_research")
INFER = load_script("infer-canonical-formulas.py", "draft_mode_infer")
DIRECT = load_script("build-research-workbook.py", "draft_mode_direct")
AUDIT = load_script("audit-xlsx.py", "draft_mode_audit")


def run(callable_, *args):
    """Call a CLI main(); return (exit code, stdout, stderr)."""
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = callable_(*args)
    return code, out.getvalue(), err.getvalue()


class Fixtures(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def source(self, name="source.xlsx", *, chart=False, blocked=False):
        wb = Workbook()
        ws = wb.active
        ws.append(["Category", "Value"])
        ws.append(["A", 1])
        if chart:
            bar = BarChart()
            bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=2), titles_from_data=True)
            ws.add_chart(bar, "D2")
        path = self.dir / name
        wb.save(path)
        if blocked:
            with zipfile.ZipFile(path, "a") as archive:
                archive.writestr("xl/activeX/activeX1.xml", "<activeXControl/>")
        return path

    def plan(self, source, *, targets=None, issues=None, section="Revenue"):
        plan = {
            "name": "Draft fixture", "intent": "read-only", "connectionId": "connection",
            "folderId": "folder", "currencyFormat": "$,.0f", "pctFormat": ".1%",
            "sourcePath": str(source),
            "sourceSha256": hashlib.sha256(Path(source).read_bytes()).hexdigest(),
            "actualMonths": [{"month": "Jan-26", "num": 1, "end": "2026-01-31"}],
            "forecastMonths": [{"month": "Feb-26", "num": 2, "end": "2026-02-28"}],
            "lines": [{"line": "Revenue", "section": section, "secOrder": 1, "lineOrder": 1,
                       "kind": "growth", "rate": 0.1, "juneBase": 100.0}],
            "assumptions": [["Revenue growth", 0.1, "monthly"]],
            "actuals": [{"date": "2026-01-10", "month": "Jan-26", "line": "Revenue",
                         "section": section, "secOrder": 1, "lineOrder": 1,
                         "dept": "Sales", "vendor": "Customer", "amt": 100.0}],
        }
        if targets is not None:
            plan["parityTargets"] = targets
        if issues is not None:
            plan["draftReviewIssues"] = issues
        path = self.dir / "plan.json"
        path.write_text(json.dumps(plan), encoding="utf-8")
        return path, plan


FLAT = [{"code": "FLAT_CARRY_POLICY", "severity": "review", "route": "r",
         "message": "m", "evidence": ["Revenue"]}]
# The plan above: actual 100 + forecast 100 * 1.1 = 210 in every section total.
GOOD_TARGETS = {"Revenue": 210.0, "Gross Profit": 210.0, "EBITDA": 210.0, "Net Income": 210.0}


class TestGateDraftMode(Fixtures):
    def test_unresolved_review_is_pending_in_draft_and_raises_in_production(self):
        source = self.source(chart=True)
        report = gate_source(source, draft=True)
        self.assertEqual(report["reviewStatus"], "DRAFT_PENDING")
        self.assertEqual(report["pendingReview"], ["CHARTS_PRESENT"])
        with self.assertRaises(MigrationGateError) as raised:
            gate_source(source)
        self.assertEqual(raised.exception.report["reviewStatus"], "UNRESOLVED")

    def test_draft_still_validates_a_supplied_dispositions_file(self):
        source = self.source(chart=True)
        stale = self.dir / "review.json"
        stale.write_text(json.dumps({"sourceSha256": "0" * 64, "issues": []}), encoding="utf-8")
        with self.assertRaisesRegex(MigrationGateError, "does not match"):
            gate_source(source, str(stale), draft=True)

    def test_draft_never_downgrades_a_blocker(self):
        source = self.source(blocked=True)
        with self.assertRaises(MigrationGateError) as raised:
            gate_source(source, draft=True)
        self.assertEqual(raised.exception.exit_code, 2)

    def test_draft_review_issues_on_a_plan_are_review_issues(self):
        self.assertEqual([i["code"] for i in plan_review_issues({"draftReviewIssues": FLAT})],
                         ["FLAT_CARRY_POLICY"])
        self.assertEqual(plan_review_issues({}), [])


class TestForecastBuilderDraft(Fixtures):
    def test_draft_build_with_unresolved_review_exits_zero_and_writes_specs(self):
        source = self.source(chart=True)
        plan_path, _ = self.plan(source, issues=FLAT)
        out = self.dir / "build"
        code, stdout, _ = run(FORECAST.main, [str(plan_path), "--out-dir", str(out)])
        self.assertEqual(code, 0)
        self.assertIn("DRAFT MODE: review pending for CHARTS_PRESENT, FLAT_CARRY_POLICY"
                      " — fine for a local build; resolve before --post.", stdout)
        for name in ("dm_spec.json", "wb_spec.json", "rate_seed.csv", "manual_seed.csv"):
            self.assertTrue((out / name).is_file(), name)

    def test_post_with_the_same_plan_fails_before_auth(self):
        source = self.source(chart=True)
        plan_path, _ = self.plan(source, issues=FLAT)
        with mock.patch.object(FORECAST, "env", side_effect=AssertionError("auth reached")):
            code, _, stderr = run(FORECAST.main, [str(plan_path), "--post"])
        self.assertEqual(code, 1)
        self.assertIn("unresolved", stderr)

    def test_post_needs_a_disposition_for_the_drafter_issue_then_reaches_auth(self):
        source = self.source()
        plan_path, plan = self.plan(source, issues=FLAT)
        review = self.dir / "review.json"
        review.write_text(json.dumps({"sourceSha256": plan["sourceSha256"], "issues": [
            {"code": "FLAT_CARRY_POLICY", "owner": "Finance", "disposition": "confirmed policy",
             "verified": True}]}), encoding="utf-8")
        with mock.patch.object(FORECAST, "env", side_effect=RuntimeError("auth reached")):
            with self.assertRaisesRegex(RuntimeError, "auth reached"):
                run(FORECAST.main, [str(plan_path), "--post", "--review-dispositions", str(review)])

    def test_blocked_source_still_exits_two_in_draft(self):
        source = self.source(blocked=True)
        plan_path, _ = self.plan(source)
        code, _, stderr = run(FORECAST.main, [str(plan_path), "--out-dir", str(self.dir / "b")])
        self.assertEqual(code, 2)
        self.assertIn("blocked", stderr)
        self.assertFalse((self.dir / "b").exists())

    def test_research_builder_prints_draft_line_and_post_stays_strict(self):
        source = self.source(chart=True)
        plan = {"sourcePath": str(source), "sourceSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                "summary": {}, "parity": {}, "anchors": {}, "lines": []}
        plan_path = self.dir / "research.json"
        plan_path.write_text(json.dumps(plan), encoding="utf-8")
        with mock.patch.object(RESEARCH, "env", side_effect=AssertionError("auth reached")):
            code, _, _ = run(RESEARCH.main, [str(plan_path), "--post"])
        self.assertEqual(code, 1)
        # The dry run passes the gate in draft mode; stop right after it (the stub plan has no lines).
        out = io.StringIO()
        with mock.patch.object(RESEARCH, "build_dm", side_effect=StopIteration), \
                contextlib.redirect_stdout(out), self.assertRaises(StopIteration):
            RESEARCH.main([str(plan_path)])
        self.assertIn("DRAFT MODE: review pending for CHARTS_PRESENT", out.getvalue())


class TestForecastParity(Fixtures):
    def build(self, **kwargs):
        source = self.source()
        plan_path, plan = self.plan(source, **kwargs.pop("plan_kwargs", {}))
        out = self.dir / "build"
        code, stdout, stderr = run(FORECAST.main, [str(plan_path), "--out-dir", str(out), *kwargs.pop("extra", [])])
        return code, stdout, stderr, out

    def test_pass_prints_one_line_per_metric_and_writes_specs(self):
        code, stdout, _, out = self.build(plan_kwargs={"targets": GOOD_TARGETS})
        self.assertEqual(code, 0)
        for metric in ("Revenue", "Gross Profit", "EBITDA", "Net Income"):
            self.assertIn(f"PARITY PASS {metric}:", stdout)
        self.assertTrue((out / "wb_spec.json").is_file())

    def test_fail_exits_one_before_any_spec_is_written(self):
        targets = {**GOOD_TARGETS, "EBITDA": 211.0}
        code, stdout, stderr, out = self.build(plan_kwargs={"targets": targets})
        self.assertEqual(code, 1)
        self.assertIn("PARITY FAIL EBITDA:", stdout)
        self.assertIn("PARITY PASS Revenue:", stdout)
        self.assertIn("LOCAL PARITY FAILED", stderr)
        self.assertFalse(out.exists())

    def test_fail_also_blocks_post_before_auth(self):
        source = self.source()
        plan_path, plan = self.plan(source, targets={**GOOD_TARGETS, "Revenue": 1.0})
        review = self.dir / "review.json"
        review.write_text(json.dumps({"sourceSha256": plan["sourceSha256"], "issues": []}), encoding="utf-8")
        with mock.patch.object(FORECAST, "env", side_effect=AssertionError("auth reached")):
            code, _, _ = run(FORECAST.main, [str(plan_path), "--post", "--review-dispositions", str(review)])
        self.assertEqual(code, 1)

    def test_missing_targets_are_reported_not_failed(self):
        code, stdout, _, _ = self.build()
        self.assertEqual(code, 0)
        self.assertIn("LOCAL PARITY: not checked — plan has no parityTargets", stdout)

    def test_tolerance_is_configurable(self):
        targets = {**GOOD_TARGETS, "Revenue": 210.005}
        code, _, _, _ = self.build(plan_kwargs={"targets": targets})
        self.assertEqual(code, 0)
        code, stdout, _, _ = self.build(plan_kwargs={"targets": targets}, extra=["--parity-tolerance", "0.001"])
        self.assertEqual(code, 1)
        self.assertIn("PARITY FAIL Revenue:", stdout)

    def test_unknown_section_names_the_supported_sections(self):
        code, _, stderr, out = self.build(plan_kwargs={"section": "Capex"})
        self.assertEqual(code, 1)
        self.assertIn("unsupported section 'Capex'", stderr)
        for section in ("Revenue", "Cost of Revenue", "Operating Expenses", "Below the Line"):
            self.assertIn(section, stderr)
        self.assertNotIn("KeyError", stderr)
        self.assertFalse(out.exists())
        with self.assertRaises(FORECAST.PlanError):
            FORECAST.parity(self.plan(self.source("b.xlsx"), section="Capex")[1])


class TestOtherDraftFlags(Fixtures):
    def test_inference_draft_reports_pending_instead_of_failing(self):
        source = self.source(chart=True)
        with mock.patch.object(INFER, "infer", side_effect=AssertionError("inference reached")):
            code, _, stderr = run(INFER.main, [str(source), "--quiet"])
        self.assertEqual(code, 1)  # default stays strict
        with mock.patch.object(INFER, "infer", return_value={"lines": []}):
            code, stdout, _ = run(INFER.main, [str(source), "--quiet", "--draft"])
        self.assertEqual(code, 0)
        self.assertIn("DRAFT MODE: review pending for CHARTS_PRESENT", stdout)

    def test_direct_research_workbook_rejects_draft_with_post(self):
        source = self.source(chart=True)
        with self.assertRaises(SystemExit) as raised:
            run(DIRECT.main, [str(source), "--draft", "--post"])
        self.assertEqual(raised.exception.code, 2)


class TestAuditExitCodes(Fixtures):
    def test_review_required_exits_zero_by_default_and_one_when_strict(self):
        source = self.source(chart=True)
        self.assertEqual(run(AUDIT.main, [str(source)])[0], 0)
        self.assertEqual(run(AUDIT.main, [str(source), "--strict"])[0], 1)
        self.assertEqual(run(AUDIT.main, [str(source), "--allow-review"])[0], 0)

    def test_auto_exits_zero_even_when_strict(self):
        source = self.source()
        self.assertEqual(run(AUDIT.main, [str(source), "--strict"])[0], 0)

    def test_blocked_exits_two_regardless(self):
        source = self.source(blocked=True)
        self.assertEqual(run(AUDIT.main, [str(source)])[0], 2)
        self.assertEqual(run(AUDIT.main, [str(source), "--allow-review"])[0], 2)


if __name__ == "__main__":
    unittest.main()
