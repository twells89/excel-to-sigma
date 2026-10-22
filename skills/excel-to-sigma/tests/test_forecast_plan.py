#!/usr/bin/env python3
"""draft-forecast-plan.py: classification, fail-loud preconditions, verification, builder hand-off.

The synthetic workbook comes from ``forecast_pnl_fixture`` (formulas written by openpyxl,
cached values injected into the saved XML) so the drafter reads real cached values.
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import forecast_pnl_fixture as fixture  # noqa: E402
from forecast_plan import SECTIONS, PlanDraftError, draft_plan, parse_month  # noqa: E402
from migration_gate import plan_review_issues  # noqa: E402


def load_script(filename, name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


DRAFT = load_script("draft-forecast-plan.py", "draft_forecast_plan_cli")
BUILD = load_script("build-forecast-model.py", "draft_forecast_plan_builder")


def find_cell(ws, text, col=1):
    return next(r for r in range(1, ws.max_row + 1) if ws.cell(r, col).value == text)


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        self.path = self.dir / "pnl.xlsx"

    def draft(self, **fixture_kwargs):
        built = fixture.build_pnl(self.path, **fixture_kwargs)
        plan, info = draft_plan(self.path, intent="what-if", connection="conn", folder="fold")
        return plan, info, built

    def assert_fails(self, pattern, **fixture_kwargs):
        fixture.build_pnl(self.path, **fixture_kwargs)
        with self.assertRaisesRegex(PlanDraftError, pattern):
            draft_plan(self.path, intent="what-if", connection="conn", folder="fold")


class TestClassification(Base):
    def test_kinds_values_months_and_order(self):
        plan, info, built = self.draft()
        lines = {line["line"]: line for line in plan["lines"]}
        self.assertEqual({n: l["kind"] for n, l in lines.items()},
                         {"Subscription": "growth", "Services": "manual", "Hosting": "flat",
                          "Support": "growth", "R&D": "manual", "Interest": "flat"})
        self.assertEqual(lines["Subscription"]["rate"], 0.05)
        self.assertEqual(lines["Subscription"]["juneBase"], 1200.0)  # last actual month
        self.assertEqual(lines["Subscription"]["driver"], "Subscription growth")
        self.assertEqual(lines["Services"]["manual"], [300.0, 320.0, 340.0])
        self.assertEqual(lines["Hosting"]["juneBase"], 220.0)
        self.assertEqual([(l["section"], l["secOrder"], l["lineOrder"]) for l in plan["lines"]],
                         [("Revenue", 1, 1), ("Revenue", 1, 2), ("Cost of Revenue", 2, 3),
                          ("Cost of Revenue", 2, 4), ("Operating Expenses", 3, 5), ("Below the Line", 4, 6)])
        self.assertEqual([m["month"] for m in plan["actualMonths"]], ["Jan-26", "Feb-26", "Mar-26"])
        self.assertEqual([m["month"] for m in plan["forecastMonths"]], ["Apr-26", "May-26", "Jun-26"])
        self.assertEqual(plan["forecastMonths"][0], {"month": "Apr-26", "num": 4, "end": "2026-04-30"})
        self.assertEqual(len(plan["actuals"]), 36)
        self.assertEqual(plan["assumptions"], [["Subscription growth", 0.05, "Subscription basis"],
                                               ["Support growth", 0.02, "Support basis"]])
        self.assertEqual(info["kinds"], {"growth": 2, "manual": 2, "flat": 2})

    def test_parity_targets_come_from_the_fy_column(self):
        plan, info, built = self.draft()
        self.assertEqual(info["fyColumn"], "H")
        for metric, value in built["targets"].items():
            self.assertAlmostEqual(plan["parityTargets"][metric], value, places=6)

    def test_parity_targets_fall_back_to_summed_months_without_an_fy_column(self):
        def drop_fy(pnl, gl, drivers):
            pnl["H3"] = None
        plan, info, built = self.draft(mutate=drop_fy)
        self.assertIsNone(info["fyColumn"])
        self.assertAlmostEqual(plan["parityTargets"]["Revenue"], built["targets"]["Revenue"], places=6)

    def test_non_fy_summary_headers_fall_back_to_summed_months(self):
        for header in ("Year to Date", "FY26 Budget", "Total Variance"):
            with self.subTest(header=header):
                def rename_fy(pnl, gl, drivers):
                    pnl["H3"] = header
                plan, info, built = self.draft(mutate=rename_fy)
                self.assertIsNone(info["fyColumn"])
                self.assertAlmostEqual(
                    plan["parityTargets"]["Revenue"], built["targets"]["Revenue"], places=6)

    def test_fy_header_whose_values_are_not_the_month_sum_is_ignored(self):
        # e.g. a "FY26" column that holds budget figures rather than the year's total
        probe = self.dir / "probe.xlsx"
        fixture.build_pnl(probe)
        from openpyxl import load_workbook
        row = find_cell(load_workbook(probe)["P&L"], "Subscription")
        plan, info, built = self.draft(cached_overrides={("P&L", f"H{row}"): 1.0})
        self.assertIsNone(info["fyColumn"])
        self.assertAlmostEqual(plan["parityTargets"]["Revenue"], built["targets"]["Revenue"], places=6)

    def test_source_binding_and_identity_fields(self):
        plan, _, _ = self.draft()
        from migration_gate import source_sha256
        self.assertEqual(plan["sourceSha256"], source_sha256(self.path))
        self.assertEqual(plan["sourcePath"], str(self.path.resolve()))
        self.assertEqual((plan["intent"], plan["connectionId"], plan["folderId"]), ("what-if", "conn", "fold"))
        self.assertEqual((plan["currencyFormat"], plan["pctFormat"]), ("$,.0f", ".1%"))
        self.assertEqual(plan["name"], "pnl")  # no workbook title -> file stem

    def test_report_sheet_is_discovered_and_overridable(self):
        plan, info, _ = self.draft()
        self.assertEqual(info["reportSheet"], "P&L")
        with self.assertRaisesRegex(PlanDraftError, "--sheet 'Nope' not found"):
            draft_plan(self.path, intent="what-if", connection="c", folder="f", sheet="Nope")

    def test_gl_columns_are_mapped_by_name_not_position(self):
        def shuffle(pnl, gl, drivers):
            # Swap the Department and Vendor header positions (data stays put -> labels swap).
            gl["D4"], gl["E4"] = "Memo / Vendor", "Dept"
        plan, _, _ = self.draft(mutate=shuffle)
        self.assertEqual(plan["actuals"][0]["vendor"], "Revenue")  # column D now holds the vendor role
        self.assertEqual(plan["actuals"][0]["dept"], "vendor 0")

    def test_month_headers_accept_datetimes(self):
        from datetime import datetime

        def real_dates(pnl, gl, drivers):
            for index in range(6):
                pnl.cell(3, 2 + index, datetime(2026, index + 1, 1))
        # SUMIFS criteria then compare datetimes with GL text months, so cached values still
        # come from the clean layout; only the header parsing is under test here.
        fixture.build_pnl(self.path, mutate=real_dates)
        plan, _ = draft_plan(self.path, intent="what-if", connection="c", folder="f")
        self.assertEqual(plan["actualMonths"][0]["month"], "Jan-26")
        self.assertEqual(parse_month("Jan-2026"), (2026, 1))
        self.assertIsNone(parse_month("FY2026"))


class TestReviewIssues(Base):
    def test_flat_lines_raise_flat_carry_policy_with_evidence(self):
        plan, info, _ = self.draft()
        (issue,) = plan["draftReviewIssues"]
        self.assertEqual((issue["code"], issue["severity"]), ("FLAT_CARRY_POLICY", "review"))
        self.assertEqual(issue["evidence"], ["Hosting", "Interest"])
        self.assertEqual(info["reviewIssues"], ["FLAT_CARRY_POLICY"])
        self.assertEqual([i["code"] for i in plan_review_issues(plan)], ["FLAT_CARRY_POLICY"])

    def test_no_flat_lines_means_no_drafter_issue(self):
        lines = [l for l in fixture.DEFAULT_LINES if l[2] != "flat"]
        plan, _, _ = self.draft(lines=lines)
        self.assertEqual(plan["draftReviewIssues"], [])


class TestFailLoud(Base):
    def test_unknown_section_lists_what_was_found(self):
        def rename(pnl, gl, drivers):
            pnl.cell(find_cell(pnl, "Below the Line"), 1, "Capex")
        self.assert_fails(r"unsupported section 'Capex'.*Sections found: Revenue, Cost of Revenue, "
                          r"Operating Expenses, Capex", mutate=rename)

    def test_cached_forecast_value_mismatch_names_the_cell(self):
        self.assert_fails(r"(?s)does not reproduce.*P&L!G5 \('Subscription', Jun-26\)",
                          cached_overrides={("P&L", "G5"): 9999.0})

    def test_gl_sum_that_disagrees_with_the_cached_sumifs_fails(self):
        def bump(pnl, gl, drivers):
            gl["F5"] = gl["F5"].value + 50
        self.assert_fails(r"P&L!B5 \('Subscription', Jan-26\): plan 1,050\.00 vs sheet 1,000\.00",
                          mutate=bump)

    def test_mixed_forecast_cells_fail_naming_row_and_cell(self):
        def mix(pnl, gl, drivers):
            pnl["F5"] = 1500  # growth, growth, literal
        self.assert_fails(r"P&L!F5 \('Subscription'\): forecast cells are not uniform", mutate=mix)

    def test_unrecognised_forecast_formula_fails(self):
        def odd(pnl, gl, drivers):
            pnl["E5"] = "=D5*1.05"
        self.assert_fails(r"P&L!E5 \('Subscription'\): unrecognised forecast cell '=D5\*1\.05'",
                          mutate=odd)

    def test_growth_cells_pointing_at_different_drivers_fail(self):
        def split(pnl, gl, drivers):
            pnl["F5"] = "=E5*(1+Drivers!$B$5)"
        self.assert_fails(r"P&L!F5 \('Subscription'\).*different driver", mutate=split)

    def test_growth_driver_on_a_quoted_sheet_name_keeps_its_spaces(self):
        from openpyxl import Workbook
        from forecast_plan import _forecast_cell_kind
        ws = Workbook().active
        kind, detail = _forecast_cell_kind(ws, 5, 7, "= D7 * (1 + 'Driver Assumptions'!$B$4)")
        self.assertEqual((kind, detail), ("growth", ("Driver Assumptions", "B", 4)))

    def test_growth_from_a_non_adjacent_cell_fails(self):
        def skip(pnl, gl, drivers):
            for col in "DEF":
                pnl[f"{col}5"] = f"=C5*(1+Drivers!$B$4)"
        self.assert_fails(r"unrecognised forecast cell", mutate=skip)

    def test_missing_gl_column_fails_naming_it(self):
        def rename(pnl, gl, drivers):
            gl["E4"] = "Notes"
        self.assert_fails(r"no header row maps all of.*missing vendor", mutate=rename)

    def test_gl_line_not_on_the_pnl_fails(self):
        def stray(pnl, gl, drivers):
            gl["C5"] = "Mystery line"
        self.assert_fails(r"GL!row 5: line item 'Mystery line' is not a P&L line", mutate=stray)

    def test_gl_row_in_a_forecast_month_fails(self):
        def late(pnl, gl, drivers):
            gl["B5"] = "May-26"
        self.assert_fails(r"GL!row 5: month 'May-26' is not an actual month", mutate=late)

    def test_unrecognised_row_fails_instead_of_being_skipped(self):
        def extra(pnl, gl, drivers):
            row = pnl.max_row + 1
            pnl.cell(row, 1, "Mystery metric")  # numbers under a section, but no SUMIFS actuals
            for col in range(2, 8):
                pnl.cell(row, col, 5)
        self.assert_fails(r"Mystery metric.*neither a SUMIFS actuals line", mutate=extra)

    def test_workbook_without_cached_values_fails_with_guidance(self):
        # A plain openpyxl save has formulas but no cached values.
        fixture.build_pnl(self.path)
        from openpyxl import load_workbook
        book = load_workbook(self.path)
        book.save(self.path)
        with self.assertRaisesRegex(PlanDraftError, "no numeric cached value"):
            draft_plan(self.path, intent="what-if", connection="c", folder="f")

    def test_not_a_forecast_model_fails(self):
        from openpyxl import Workbook
        book = Workbook()
        book.active["A1"] = "just text"
        book.save(self.path)
        with self.assertRaisesRegex(PlanDraftError, "no sheet has SUMIFS actuals"):
            draft_plan(self.path, intent="what-if", connection="c", folder="f")

    def test_cli_reports_failure_on_stderr_with_nonzero_exit(self):
        fixture.build_pnl(self.path, cached_overrides={("P&L", "G5"): 9999.0})
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = DRAFT.main([str(self.path), "--connection", "c", "--folder", "f",
                               "--out", str(self.dir / "plan.json")])
        self.assertEqual(code, 1)
        self.assertIn("DRAFT FAILED", err.getvalue())
        self.assertFalse((self.dir / "plan.json").exists())


def add_metric_chart(labels, title="Revenue & Gross Profit", min_col=2, max_col=None):
    """mutate hook: a P&L line chart whose series are the given total rows by month."""
    from openpyxl.chart import LineChart, Reference, Series
    from openpyxl.chart.data_source import StrRef
    from openpyxl.chart.series import SeriesLabel

    def mutate(pnl, gl, drivers):
        chart = LineChart()
        chart.title = title
        last_col = max_col or 1 + len(fixture.MONTHS)
        for label in labels:
            row = find_cell(pnl, label)
            series = Series(Reference(pnl, min_col=min_col, max_col=last_col, min_row=row))
            series.tx = SeriesLabel(strRef=StrRef(f"'P&L'!$A${row}"))
            chart.series.append(series)
        chart.set_categories(Reference(pnl, min_col=min_col, max_col=last_col,
                                       min_row=fixture.HEADER_ROW))
        pnl.add_chart(chart, "P2")
    return mutate


class TestSourceCharts(Base):
    def test_metric_line_chart_is_mapped_and_verified(self):
        plan, info, _ = self.draft(mutate=add_metric_chart(["Total Revenue", "Gross Profit"]))
        self.assertEqual(plan["sourceCharts"], [{
            "title": "Revenue & Gross Profit", "kind": "line", "grain": "month",
            "series": ["Revenue", "Gross Profit"], "sourceSheet": "P&L",
            "verifiedPoints": 2 * len(fixture.MONTHS)}])
        self.assertEqual(info["unmappedCharts"], [])

    def test_chart_with_a_non_metric_series_is_left_for_review(self):
        plan, info, _ = self.draft(mutate=add_metric_chart(["Total Revenue", "Subscription"]))
        self.assertEqual(plan["sourceCharts"], [])
        self.assertIn("not Total Revenue/Gross Profit/EBITDA/Net Income", info["unmappedCharts"][0])

    def test_chart_with_mid_span_months_is_left_for_review(self):
        plan, info, _ = self.draft(
            mutate=add_metric_chart(["Total Revenue"], min_col=5, max_col=7))
        self.assertEqual(plan["sourceCharts"], [])
        self.assertIn("month categories", info["unmappedCharts"][0])

    def test_chart_with_literal_categories_is_left_for_review(self):
        from openpyxl.chart.data_source import AxDataSource, StrData, StrVal

        def literal_categories(pnl, gl, drivers):
            add_metric_chart(["Total Revenue"])(pnl, gl, drivers)
            points = [StrVal(idx=i, v=month) for i, month in enumerate(fixture.MONTHS)]
            pnl._charts[-1].series[0].cat = AxDataSource(strLit=StrData(pt=points))

        plan, info, _ = self.draft(mutate=literal_categories)
        self.assertEqual(plan["sourceCharts"], [])
        self.assertEqual(len(info["unmappedCharts"]), 1)

    def test_chart_values_that_disagree_with_the_plan_fail(self):
        # Total rows are not line rows, so only the chart check reads this cache.
        fixture.build_pnl(self.path, mutate=add_metric_chart(["Total Revenue"]),
                          cached_overrides={("P&L", f"B{self._total_row()}"): 999999})
        with self.assertRaisesRegex(PlanDraftError, r"chart .*'Revenue'"):
            draft_plan(self.path, intent="what-if", connection="conn", folder="fold")

    def _total_row(self):
        probe = self.dir / "probe.xlsx"
        fixture.build_pnl(probe)
        from openpyxl import load_workbook
        return find_cell(load_workbook(probe)["P&L"], "Total Revenue")

    def test_builder_emits_a_line_chart_per_source_chart(self):
        fixture.build_pnl(self.path, mutate=add_metric_chart(["Total Revenue", "Gross Profit"]))
        plan, _ = draft_plan(self.path, intent="what-if", connection="conn", folder="fold")
        spec = BUILD.build_wb_spec(plan, "<dm-id>", {"glDetail": "glDetail", "spine": "spine",
                                                     "assumptions": "assumptions"})
        overview = spec["pages"][0]["elements"]
        chart = next(e for e in overview if e["id"] == "sourceChart1")
        self.assertEqual(chart["kind"], "line-chart")
        self.assertEqual(chart["name"], "Revenue & Gross Profit")
        self.assertEqual([c["name"] for c in chart["columns"]], ["Month", "Revenue", "Gross Profit"])
        self.assertIn('elementId="sourceChart1"', spec["layout"])


class TestBuilderHandoff(Base):
    def run_main(self, module, argv):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = module.main(argv)
        return code, out.getvalue(), err.getvalue()

    def test_builder_and_drafter_agree_on_sections(self):
        self.assertEqual(tuple(BUILD.SECTIONS), SECTIONS)

    def test_drafted_plan_passes_the_builder_dry_run_with_parity(self):
        fixture.build_pnl(self.path)
        plan_path = self.dir / "plan.json"
        code, stdout, _ = self.run_main(DRAFT, [str(self.path), "--connection", "conn", "--folder", "fold",
                                                "--out", str(plan_path)])
        self.assertEqual(code, 0)
        self.assertIn("6 lines (flat 2, growth 2, manual 2)", stdout)
        self.assertIn("36 GL rows", stdout)
        build_dir = self.dir / "build"
        code, stdout, _ = self.run_main(BUILD, [str(plan_path), "--out-dir", str(build_dir)])
        self.assertEqual(code, 0)
        for metric in ("Revenue", "Gross Profit", "EBITDA", "Net Income"):
            self.assertIn(f"PARITY PASS {metric}:", stdout)
        self.assertIn("DRAFT MODE: review pending for FLAT_CARRY_POLICY", stdout)
        dm = json.loads((build_dir / "dm_spec.json").read_text())
        self.assertEqual(dm["pages"][0]["elements"][0]["id"], "glDetail")

    def test_drafted_plan_post_needs_the_flat_policy_disposition(self):
        fixture.build_pnl(self.path)
        plan_path = self.dir / "plan.json"
        self.run_main(DRAFT, [str(self.path), "--connection", "c", "--folder", "f", "--out", str(plan_path)])
        with mock.patch.object(BUILD, "env", side_effect=AssertionError("auth reached")):
            code, _, stderr = self.run_main(BUILD, [str(plan_path), "--post"])
        self.assertEqual(code, 1)
        self.assertIn("FLAT_CARRY_POLICY", stderr)


if __name__ == "__main__":
    unittest.main()
