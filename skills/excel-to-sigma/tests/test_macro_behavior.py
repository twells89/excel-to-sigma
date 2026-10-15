from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts" / "lib"))

from macro_behavior import (
    MacroBehaviorError,
    analyze_macro_source,
    procedure_by_identity,
    validate_macro_source,
)


class TestMacroBehavior(unittest.TestCase):
    def write_module(self, text: str, name: str = "Workflow.bas") -> Path:
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = Path(folder.name) / name
        path.write_bytes(text.encode("utf-8"))
        return path

    def test_same_name_different_body_changes_provenance_and_write_operation(self):
        update = self.write_module(
            'Attribute VB_Name = "Workflow"\nSub Commit()\nCells(2, 9).Value = "Approved"\nEnd Sub\n',
            "update.bas",
        )
        append = self.write_module(
            'Attribute VB_Name = "Workflow"\nSub Commit()\nListRows.Add\nEnd Sub\n',
            "append.bas",
        )

        left = procedure_by_identity(analyze_macro_source(update), "Workflow.Commit")
        right = procedure_by_identity(analyze_macro_source(append), "Workflow.Commit")

        self.assertNotEqual(left["bodySha256"], right["bodySha256"])
        self.assertNotEqual(left["procedureSha256"], right["procedureSha256"])
        self.assertEqual(left["disposition"]["operation"], "update")
        self.assertEqual(right["disposition"]["operation"], "append")

    def test_comments_never_create_writes_calls_or_external_effects(self):
        source = self.write_module(
            'Attribute VB_Name = "Comments"\nSub SaveAnything()\n'
            "' ListRows.Add: Commit Send RunPython\n"
            'Application.CalculateFull\nEnd Sub\n'
        )

        procedure = analyze_macro_source(source)["procedures"][0]

        self.assertEqual(procedure["disposition"]["operation"], "recalculate")
        self.assertEqual(procedure["writes"], [])
        self.assertEqual(procedure["externalSideEffects"], [])
        self.assertEqual(procedure["calls"], [])

    def test_local_helper_is_expanded_and_missing_helper_blocks_emission(self):
        present = self.write_module(
            'Attribute VB_Name = "Helpers"\nSub Parent()\nCall WriteStatus\nEnd Sub\n'
            'Private Sub WriteStatus()\nCells(2, 9).Value = "Approved"\nEnd Sub\n',
            "present.bas",
        )
        missing = self.write_module(
            'Attribute VB_Name = "Helpers"\nSub Parent()\nCall WriteStatus\nEnd Sub\n',
            "missing.bas",
        )

        expanded = procedure_by_identity(analyze_macro_source(present), "Helpers.Parent")
        unresolved = procedure_by_identity(analyze_macro_source(missing), "Helpers.Parent")

        self.assertEqual(expanded["helperDefinitions"][0]["procedureId"], "Helpers.WriteStatus")
        self.assertEqual(expanded["writes"][0]["column"], 9)
        self.assertEqual(expanded["unresolvedDependencies"], [])
        self.assertEqual(unresolved["unresolvedDependencies"][0]["name"], "WriteStatus")
        self.assertEqual(unresolved["disposition"]["status"], "UNRESOLVED")

    def test_synthetic_approval_is_selected_row_update_with_exact_guards(self):
        source = self.write_module(
            'Attribute VB_Name = "ApprovalWorkflow"\nOption Explicit\n'
            'Public Sub CommitExpenseApproval_Click()\n'
            'Dim targetRow As Long\ntargetRow = ActiveCell.Row\n'
            'If targetRow < 2 Then Exit Sub\n'
            'If ActiveSheet.Name <> "Expense Register" Then Exit Sub\n'
            'Cells(targetRow, 9).Value = "Approved"\n'
            'Cells(targetRow, 16).Value = Now\nEnd Sub\n'
        )

        procedure = analyze_macro_source(source)["procedures"][0]

        self.assertEqual(procedure["disposition"], {"operation": "update", "status": "ACTION"})
        self.assertEqual(procedure["selectedRecord"]["source"], "ActiveCell.Row")
        self.assertEqual([guard["kind"] for guard in procedure["guards"]],
                         ["minimum-row", "active-sheet"])
        self.assertEqual([write["column"] for write in procedure["writes"]], [9, 16])
        self.assertTrue(procedure["writes"][1]["usesNow"])
        self.assertFalse(any(write["kind"] == "append" for write in procedure["writes"]))

    def test_sample_snapshot_records_rectangle_append_and_h7_overwrite(self):
        source = self.write_module(
            'Attribute VB_Name = "FPAModel"\nSub CommitForecast_Click()\n'
            'Set src = Sheets("P&L").Range("H7:M33")\n'
            'nextRow = dst.Cells(dst.Rows.Count, 1).End(xlUp).Row + 1\n'
            'src.Copy\ndst.Cells(nextRow, 1).PasteSpecial Paste:=xlPasteValues\n'
            'dst.Cells(nextRow, 1).Value = Now\nEnd Sub\n'
        )

        procedure = analyze_macro_source(source)["procedures"][0]

        self.assertEqual(procedure["affectedRanges"][0]["rows"], 27)
        self.assertEqual(procedure["affectedRanges"][0]["columns"], 6)
        self.assertEqual(procedure["disposition"]["operation"], "append")
        self.assertEqual(procedure["overwrites"][0]["sourceCell"], "P&L!H7")

    def test_strings_and_fixed_destination_pastes_do_not_authorize_append(self):
        path=self.write_module('Sub Work()\nMsgBox "ListRows.Add"\nEnd Sub\n')
        self.assertEqual(analyze_macro_source(path)['procedures'][0]['writes'],[])
        fixed=self.write_module('Sub Work()\n'
            'Sheets("P&L").Range("H7:M33").Copy\n'
            'dst.Cells(2, 1).PasteSpecial Paste:=xlPasteValues\nEnd Sub\n')
        self.assertNotEqual(analyze_macro_source(fixed)['procedures'][0]['disposition']['operation'],'append')

    def test_dynamic_range_stays_unresolved(self):
        source = self.write_module(
            'Attribute VB_Name = "Dynamic"\nSub Build()\n'
            'Set src = Sheets(sheetName).Range(address)\nsrc.Copy\nEnd Sub\n'
        )

        procedure = analyze_macro_source(source)["procedures"][0]

        self.assertIn("dynamic-sheet-or-range", {
            item["kind"] for item in procedure["unresolvedDependencies"]})
        self.assertEqual(procedure["disposition"]["status"], "UNRESOLVED")

    def test_transitive_helpers_and_bare_calls_cannot_hide_side_effects(self):
        path = self.write_module('Sub Parent()\nMiddle\nEnd Sub\n'
                                 'Sub Middle()\nCall Child\nEnd Sub\n'
                                 'Sub Child()\nListRows.Add\nEnd Sub\n')
        parent = analyze_macro_source(path)["procedures"][0]
        self.assertEqual(parent["disposition"]["operation"], "append")
        self.assertEqual(len(parent["helperDefinitions"]), 2)

    def test_argument_bearing_helpers_expand_transitively_or_remain_missing(self):
        present=self.write_module('Sub Parent()\nDeleteAudit targetRow\nEnd Sub\n'
            'Sub DeleteAudit(r As Long)\nCall WriteAudit(r)\nEnd Sub\n'
            'Sub WriteAudit(r As Long)\nCells(r, 10).Value = "Deleted"\nEnd Sub\n')
        parent=analyze_macro_source(present)['procedures'][0]
        self.assertEqual([c['name'] for c in parent['calls']],['DeleteAudit'])
        self.assertEqual(parent['writes'][0]['column'],10)
        self.assertEqual(len(parent['helperDefinitions']),2)
        missing=self.write_module('Sub Parent()\nMissingHelper targetRow\nEnd Sub\n')
        parent=analyze_macro_source(missing)['procedures'][0]
        self.assertIn('MissingHelper',[d.get('name') for d in parent['unresolvedDependencies']])
        self.assertEqual(parent['disposition']['status'],'UNRESOLVED')

    def test_function_expression_helpers_are_not_silently_discarded(self):
        path=self.write_module('Sub Parent()\nDim result As Boolean\nresult = DeleteAudit(targetRow)\nEnd Sub\n'
            'Function DeleteAudit(r As Long) As Boolean\nCells(r, 10).Value = "Deleted"\nDeleteAudit = True\nEnd Function\n')
        parent=analyze_macro_source(path)['procedures'][0]
        self.assertEqual([c['name'] for c in parent['calls']],['DeleteAudit'])
        self.assertEqual(parent['writes'][0]['column'],10)
        missing=self.write_module('Sub Parent()\nresult = MissingFunction(targetRow)\nEnd Sub\n')
        parent=analyze_macro_source(missing)['procedures'][0]
        self.assertIn('MissingFunction',[d.get('name') for d in parent['unresolvedDependencies']])

    def test_unmodeled_mutations_and_member_calls_are_unresolved(self):
        for statement in ['Range("A1").ClearContents', 'Rows(2).Delete',
                          'Range("A1").Value = 3', 'Cells(2, 10) = "Deleted"',
                          'AuditModule.DeleteAudit targetRow']:
            with self.subTest(statement=statement):
                path=self.write_module('Sub Parent()\n'+statement+'\nEnd Sub\n')
                proc=analyze_macro_source(path)['procedures'][0]
                self.assertTrue(proc['unresolvedDependencies'])
                self.assertEqual(proc['disposition']['status'],'UNRESOLVED')

    def test_early_exit_guard_calls_expand_even_inside_compound_and_nested_conditions(self):
        helper=('Function DeleteAudit(r As Long) As Boolean\nCells(r, 10).Value = "Deleted"\nDeleteAudit = False\nEnd Function\n')
        for condition in ['targetRow < 2 Or DeleteAudit(targetRow)',
                          '(targetRow < 2) Or (Not (DeleteAudit(targetRow)))',
                          'targetRow < 2 Or CBool(DeleteAudit(targetRow))']:
            with self.subTest(condition=condition):
                path=self.write_module('Sub Parent()\nDim targetRow As Long\n'
                    'targetRow = ActiveCell.Row\nIf '+condition+' Then Exit Sub\nEnd Sub\n'+helper)
                proc=analyze_macro_source(path)['procedures'][0]
                self.assertIn('DeleteAudit',[c['name'] for c in proc['calls']])
                self.assertIn(10,[w.get('column') for w in proc['writes']])
                self.assertTrue(proc['helperDefinitions'])
                self.assertTrue(proc['unresolvedDependencies'])
                self.assertEqual(proc['disposition']['status'],'UNRESOLVED')

    def test_guard_missing_functions_identifiers_and_member_behavior_remain_unresolved(self):
        for condition,missing in [('targetRow < 2 Or MissingGuard(targetRow)','MissingGuard'),
                                  ('targetRow < 2 Or MissingGuard','MissingGuard'),
                                  ('targetRow < 2 Or AuditModule.Check(targetRow)',None),
                                  ('targetRow < 2 Or Range("A1").ClearContents',None)]:
            with self.subTest(condition=condition):
                path=self.write_module('Sub Parent()\nDim targetRow As Long\n'
                    'targetRow = ActiveCell.Row\nIf '+condition+' Then Exit Sub\nEnd Sub\n')
                proc=analyze_macro_source(path)['procedures'][0]
                self.assertTrue(proc['unresolvedDependencies'])
                self.assertEqual(proc['disposition']['status'],'UNRESOLVED')
                if missing:
                    text=str(proc['unresolvedDependencies'])+str(proc['calls'])
                    self.assertIn(missing,text)

    def test_forged_behavior_facts_fail_source_validation(self):
        path = self.write_module('Sub Work()\nApplication.Calculate\nEnd Sub\n')
        analysis = analyze_macro_source(path)
        analysis['procedures'][0]['writes'] = [{'kind': 'append'}]
        with self.assertRaises(MacroBehaviorError):
            validate_macro_source(analysis)

    def test_validation_reopens_exact_module_bytes(self):
        source = self.write_module(
            'Attribute VB_Name = "Stable"\nSub Work()\nApplication.Calculate\nEnd Sub\n'
        )
        analysis = analyze_macro_source(source)
        self.assertEqual(analysis["moduleByteSha256"],
                         hashlib.sha256(source.read_bytes()).hexdigest())
        validate_macro_source(analysis)

        source.write_text(source.read_text() + "' changed\n", encoding="utf-8")
        with self.assertRaisesRegex(MacroBehaviorError, "module bytes changed"):
            validate_macro_source(analysis)


if __name__ == "__main__":
    unittest.main()
