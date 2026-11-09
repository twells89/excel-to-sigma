#!/usr/bin/env python3
"""The demo source workbook must route as a forecast, reproduce its cached values, and carry
the seven demo macros with the expected classes."""
from __future__ import annotations

import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(SCRIPTS))

from forecast_plan import draft_plan  # noqa: E402

# The demo kit (docs/demo-kit: the demo source generator and its VBA module) lives outside the skill;
# a standalone copy of the skill has no demo kit, so the tests that need it skip.
DEMO_KIT = ROOT.parents[1] / "docs" / "demo-kit"
BAS = DEMO_KIT / "demo-forecast-macros.bas"
TABS = ["Instructions", "Actuals", "Plan", "Dimensions", "Assumptions", "Forecast", "Variance",
        "Versions", "Approvals"]
TABLES = ["tblActuals", "tblPlan", "tblRegions", "tblProducts", "tblDrivers", "tblVersions",
          "tblApprovals"]


def load_generator():
    spec = importlib.util.spec_from_file_location("make_demo_forecast", DEMO_KIT / "make-demo-forecast.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_json(*args):
    done = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True)
    return done.returncode, json.loads(done.stdout)


@unittest.skipUnless(DEMO_KIT.is_dir(), "demo kit not present (standalone skill copy)")
class TestDemoForecastSource(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.src = load_generator().build_demo_workbook(Path(cls._tmp.name))

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def test_tabs_and_tables(self):
        from openpyxl import load_workbook
        wb = load_workbook(self.src)
        self.assertEqual(wb.sheetnames, TABS)
        found = sorted(name for ws in wb for name in ws.tables)
        self.assertEqual(found, sorted(TABLES))

    def test_routes_as_forecast(self):
        _code, plan = run_json(SCRIPTS / "migrate.py", "plan", self.src, "--out-dir", self._tmp.name)
        self.assertEqual(plan["archetype"], "forecast", plan.get("routeReason"))
        self.assertNotEqual(plan["auditStatus"], "BLOCKED")

    def test_drafter_reproduces_cached_values(self):
        _plan, shape = draft_plan(self.src, intent="what-if", connection="<c>", folder="<f>")
        self.assertGreaterEqual(shape["lines"], 12)

    def test_drafter_captures_driver_scenarios(self):
        plan, _shape = draft_plan(self.src, intent="what-if", connection="<c>", folder="<f>")
        scn = plan["scenarios"]
        self.assertEqual(scn["names"], ["Downside", "Base", "Upside"])
        self.assertEqual(scn["default"], "Base")      # the scenario whose rates are the live rates
        na = next(ln for ln in plan["lines"] if ln["line"] == "North America Revenue")
        self.assertAlmostEqual(scn["rates"][na["driver"]]["Base"], na["rate"])
        self.assertAlmostEqual(scn["rates"][na["driver"]]["Upside"], round(na["rate"] + 0.004, 4))
        self.assertAlmostEqual(scn["rates"][na["driver"]]["Downside"], round(na["rate"] - 0.004, 4))
        self.assertEqual(len(scn["rates"]), len(plan["assumptions"]))

    def test_macro_classes(self):
        done = subprocess.run([sys.executable, SCRIPTS / "macro-classify.py", "--json", BAS],
                              capture_output=True, text=True, check=True)
        recs = {r["name"]: r for r in json.loads(done.stdout)}
        got = {n: r["class"] for n, r in recs.items()}
        self.assertEqual(sorted(got), sorted([
            "RefreshActuals_Click", "SetScenario", "RollForwardMonth_Click", "SnapshotVersion_Click",
            "SubmitForApproval_Click", "RecalcAll", "FormatReport"]))
        self.assertEqual(got["RefreshActuals_Click"], "EXTERNAL_DATA_REFRESH")
        self.assertEqual(got["SetScenario"], "WHATIF_SCENARIO")
        self.assertEqual(got["SnapshotVersion_Click"], "ACTION_WRITEBACK")
        # Late-bound Outlook (compiles with no references): EXTERNAL_IO outranks the writeback,
        # but both the table write and the email are surfaced.
        submit = recs["SubmitForApproval_Click"]
        self.assertEqual(got["SubmitForApproval_Click"], "EXTERNAL_IO")
        self.assertIn("(also: ACTION_WRITEBACK)", submit["signals"])
        proc = submit["analysis"]["procedure"]
        self.assertIn("email", [e["kind"] for e in proc["externalSideEffects"]])
        self.assertIn("append", [w["kind"] for w in proc["writes"]])
        self.assertIn(got["RollForwardMonth_Click"], ("ACTION_WRITEBACK", "TRANSFORM_ROLLUP"))
        self.assertEqual(got["RecalcAll"], "RECALC")
        self.assertEqual(got["FormatReport"], "FORMAT_COSMETIC")

    def test_module_needs_no_vba_references(self):
        text = BAS.read_text(encoding="utf-8")
        self.assertNotIn("As Outlook.", text)
        self.assertNotIn("New Outlook", text)
        self.assertNotIn("olMailItem", text)

    def test_variance_rows_point_at_matching_forecast_rows(self):
        from openpyxl import load_workbook
        wf, wv = load_workbook(self.src), load_workbook(self.src, data_only=True)
        var, fc, fcv = wf["Variance"], wf["Forecast"], wv["Forecast"]
        checked = 0
        for row in range(4, var.max_row + 1):
            label = var.cell(row, 1).value
            if not label:
                continue
            ref = re.fullmatch(r"=Forecast!([A-Z]+)(\d+)", var.cell(row, 3).value)
            self.assertIsNotNone(ref, var.cell(row, 3).value)
            target = int(ref.group(2))
            self.assertEqual(fc.cell(target, 1).value, label)
            self.assertAlmostEqual(wv["Variance"].cell(row, 3).value,
                                   fcv[f"{ref.group(1)}{target}"].value, places=4)
            checked += 1
        self.assertEqual(checked, 12)

if __name__ == "__main__":
    unittest.main()
