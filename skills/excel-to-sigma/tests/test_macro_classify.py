#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATH = ROOT / "scripts" / "macro-classify.py"
SPEC = importlib.util.spec_from_file_location("macro_classify", PATH)
macro = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(macro)


class TestMacroClassify(unittest.TestCase):
    def classify(self, name, body):
        return macro.classify(name, f"Sub {name}()\n{body}\nEnd Sub")

    def test_python_integration_routes_to_reviewed_code_element(self):
        cls, status, target, _note, _signals = self.classify(
            "RunModel", 'RunPython ("forecast.run()")')
        self.assertEqual((cls, status), ("PYTHON_AUTOMATION", "PYTHON_REVIEW"))
        self.assertIn("run-python-element", target)

    def test_explicit_stored_procedure_routes_to_action(self):
        cls, status, target, _note, _signals = self.classify(
            "PostClose",
            'cmd.CommandType = adCmdStoredProc\ncmd.CommandText = "FINANCE.CLOSE_MONTH"',
        )
        self.assertEqual((cls, status), ("STORED_PROCEDURE", "STORED_PROC"))
        self.assertIn("call-stored-procedure", target)

    def test_vba_helper_call_is_not_a_stored_procedure(self):
        cls, status, _target, _note, _signals = self.classify(
            "Workbook_Open", "Call RebuildPnL\nApplication.Calculate")
        self.assertEqual((cls, status), ("RECALC + event-handler (auto-run)", "AUTO"))

    def test_raw_sql_refresh_fails_loud(self):
        cls, status, _target, _note, _signals = self.classify(
            "RefreshData",
            'Set cn = CreateObject("ADODB.Connection")\n'
            'cmd.CommandText = "SELECT * FROM FINANCE.ACTUALS"',
        )
        self.assertEqual((cls, status), ("EXTERNAL_DATA_REFRESH", "DATA_REVIEW"))

    def test_http_can_route_to_connector_action(self):
        cls, status, target, _note, _signals = self.classify(
            "Submit",
            'Set req = CreateObject("MSXML2.XMLHTTP")\n'
            'req.Open "POST", "https://example.invalid", False\nreq.Send payload',
        )
        self.assertEqual(status, "ACTION")
        self.assertEqual(cls, "API_CALL")
        self.assertIn("call-api", target)

    def test_transform_is_not_assumed_automatic(self):
        cls, status, _target, _note, _signals = self.classify(
            "RebuildPnL",
            "For i = 1 To 10\nCells(i, 2).Value = WorksheetFunction.Sum(i)\nNext i",
        )
        self.assertEqual((cls, status), ("TRANSFORM_ROLLUP", "PYTHON_REVIEW"))

    def test_commit_name_cannot_authorize_append(self):
        cls, status, target, *_ = self.classify("CommitExpenseApproval_Click",
            'targetRow = ActiveCell.Row\nCells(targetRow, 9).Value = "Approved"\nCells(targetRow, 16).Value = Now')
        self.assertEqual((cls, status), ("SELECTED_RECORD_UPDATE", "ACTION"))
        self.assertIn("update", target.lower())
        self.assertNotIn("append", target.lower())
        cls, status, *_ = self.classify("CommitExpenseApproval_Click", "Application.CalculateFull")
        self.assertEqual((cls, status), ("RECALC", "AUTO"))

    def test_comments_cannot_drive_classification(self):
        code = "Sub Example()\n' RunPython(\"danger\")\nApplication.Calculate\nEnd Sub"
        _name, _kind, body = next(macro.split_procedures(code))
        cls, status, *_ = macro.classify("Example", body)
        self.assertEqual((cls, status), ("RECALC", "AUTO"))

    def test_json_mode_handles_the_recorded_macro_in_the_inventory_fixture(self):
        # Regression: `macro-classify.py --json` crashed on 08 Inventory Replenishment
        # Plan.xlsm with `MacroBehaviorError: unterminated procedure: Macro1`.
        fixture = ROOT / "tests" / "fixtures" / "broader-eight" / "08 Inventory Replenishment Plan.xlsm"
        try:
            import oletools  # noqa: F401
        except ImportError:
            self.skipTest("oletools is not installed")
        result = subprocess.run([sys.executable, str(PATH), "--json", str(fixture)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        records = json.loads(result.stdout)
        self.assertEqual([r["name"] for r in records], ["Macro1"])
        self.assertEqual((records[0]["class"], records[0]["status"]), ("NAVIGATION", "AUTO"))
        self.assertEqual(records[0]["analysis"]["procedure"]["calls"], [])


if __name__ == "__main__":
    unittest.main()
