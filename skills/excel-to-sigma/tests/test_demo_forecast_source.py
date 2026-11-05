#!/usr/bin/env python3
"""The demo source workbook must route as a forecast, reproduce its cached values, and carry
the seven demo macros with the expected classes."""
from __future__ import annotations

import importlib.util
import json
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

BAS = SCRIPTS / "demo-forecast-macros.bas"
TABS = ["Instructions", "Actuals", "Plan", "Dimensions", "Assumptions", "Forecast", "Variance",
        "Versions", "Approvals"]
TABLES = ["tblActuals", "tblPlan", "tblRegions", "tblProducts", "tblDrivers", "tblVersions",
          "tblApprovals"]


def load_generator():
    spec = importlib.util.spec_from_file_location("make_demo_forecast", SCRIPTS / "make-demo-forecast.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_json(*args):
    done = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True)
    return done.returncode, json.loads(done.stdout)


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

    def test_macro_classes(self):
        done = subprocess.run([sys.executable, SCRIPTS / "macro-classify.py", "--json", BAS],
                              capture_output=True, text=True, check=True)
        got = {r["name"]: r["class"] for r in json.loads(done.stdout)}
        self.assertEqual(sorted(got), sorted([
            "RefreshActuals_Click", "SetScenario", "RollForwardMonth_Click", "SnapshotVersion_Click",
            "SubmitForApproval_Click", "RecalcAll", "FormatReport"]))
        self.assertEqual(got["RefreshActuals_Click"], "EXTERNAL_DATA_REFRESH")
        self.assertEqual(got["SetScenario"], "WHATIF_SCENARIO")
        self.assertEqual(got["SnapshotVersion_Click"], "ACTION_WRITEBACK")
        self.assertEqual(got["SubmitForApproval_Click"], "ACTION_WRITEBACK")
        self.assertIn(got["RollForwardMonth_Click"], ("ACTION_WRITEBACK", "TRANSFORM_ROLLUP"))
        self.assertEqual(got["RecalcAll"], "RECALC")
        self.assertEqual(got["FormatReport"], "FORMAT_COSMETIC")

    def test_submit_macro_carries_the_email(self):
        # The analyzer only flags email for a late-bound CreateObject("Outlook.Application"), which
        # would also outrank the writeback class. The demo macro uses early binding so it stays
        # ACTION_WRITEBACK; the email is still visible in the source the audience sees.
        text = BAS.read_text(encoding="utf-8")
        body = text[text.index("Sub SubmitForApproval_Click"):text.index("Sub RecalcAll")]
        self.assertIn("Outlook.Application", body)
        self.assertIn("cfo@", body)


if __name__ == "__main__":
    unittest.main()
