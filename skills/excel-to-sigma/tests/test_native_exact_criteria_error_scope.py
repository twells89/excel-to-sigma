"""Criteria errors follow requested physical rows in complete emitted graphs."""
import copy
import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))
from native_exact_plan import build_exact_native_calculation_plan, lower_exact_native_calculations
from native_graph_reference import NativeGraphReference
from source_inputs import build_source_input_plan
import test_native_calculations


class CriteriaErrorScopeTests(unittest.TestCase):
    source = test_native_calculations.NativeCalculationTests.source

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def lower(self, formulas):
        path, old = self.source()
        book = load_workbook(path)
        for row in (2, 3):
            book['Ledger Renamed'][f'D{row}'] = f'=1/(B{row}-4)'
        for address, formula in formulas.items():
            book['Summary Renamed'][address] = formula
        book.save(path)
        decisions = copy.deepcopy(old['decisions'])
        decisions['sourceSha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
        inputs = build_source_input_plan(path, decisions, connection_id='connection', folder_id='folder')
        plan = build_exact_native_calculation_plan(path, inputs)
        rows = [{**row, 'Row ID': f'physical-{i}'} for i, row in
                enumerate(inputs['sharedInputArchitecture']['inputs']['ledger']['rows'])]
        assigned = {'ledger': {'elementId': 'observed-input', 'name': 'tblLedger',
                              'rowIdColumn': 'Row ID',
                              'rows': {str(i): row['Row ID'] for i, row in enumerate(rows)}}}
        return plan, lower_exact_native_calculations(plan, assigned, max_columns=8), rows

    def assert_outputs(self, lowered, rows, expected):
        graph = NativeGraphReference({'observed-input': ('tblLedger', rows)}).materialize(lowered['shards'])
        for address, want in expected.items():
            with self.subTest(address=address):
                result = graph.output(lowered, 'Summary Renamed!' + address)
                if isinstance(want, str):
                    self.assertEqual(result['error'], want)
                else:
                    self.assertEqual(result['error'], '')
                    self.assertEqual(result['kind'], 'n')
                    self.assertEqual(result['number'], want)

    def test_finite_criteria_errors_exclude_unrequested_rows_after_reorder_and_append(self):
        plan, lowered, rows = self.lower({
            'A1': '=SUMIFS(\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!D2:D2,"<0")',
            'A2': '=COUNTIFS(\'Ledger Renamed\'!D2:D2,"<0")',
            'A3': '=SUMIFS(\'Ledger Renamed\'!B3:B3,\'Ledger Renamed\'!D3:D3,"<0")',
            'A4': '=COUNTIFS(\'Ledger Renamed\'!D3:D3,"<0")',
            'A5': '=SUMIFS(\'Ledger Renamed\'!D2:D2,\'Ledger Renamed\'!B2:B2,">0")',
            'A6': '=SUMIFS(\'Ledger Renamed\'!D3:D3,\'Ledger Renamed\'!B3:B3,">0")',
            'B1': '=SUMIFS(tblLedger[Amount],tblLedger[Result],"<0")',
            'B2': '=COUNTIFS(tblLedger[Result],"<0")',
        })
        expected = {'A1': 3, 'A2': 1, 'A3': '#DIV/0!', 'A4': '#DIV/0!',
                    'A5': -1, 'A6': '#DIV/0!', 'B1': '#DIV/0!', 'B2': '#DIV/0!'}
        for address, want in expected.items():
            record = plan['formulas']['Summary Renamed!' + address]
            self.assertEqual(record['localExpectedError'] if isinstance(want, str)
                             else record['localExpected'], want)
        self.assert_outputs(lowered, rows, expected)
        self.assert_outputs(lowered, list(reversed(rows)), expected)
        # Clear the original error, then add a new physical row whose live
        # calculated criterion errors. Only formal domains include that row.
        rows[1]['Amount'] = 6
        cleared = {**expected, 'A3': 0, 'A4': 0, 'A6': .5, 'B1': 3, 'B2': 1}
        self.assert_outputs(lowered, list(reversed(rows)), cleared)
        rows.insert(0, {'Account': 'c', 'Amount': 4, 'Rate': 1, 'Row ID': 'appended'})
        self.assert_outputs(lowered, rows, {**cleared, 'B1': '#DIV/0!', 'B2': '#DIV/0!'})
        # An edit to the finite domain's selected row still propagates errors.
        rows[1]['Amount'] = 4
        self.assert_outputs(lowered, rows, {**cleared, 'A1': '#DIV/0!', 'A2': '#DIV/0!',
                                           'A5': '#DIV/0!', 'B1': '#DIV/0!', 'B2': '#DIV/0!'})

    def test_raw_input_criteria_preserve_scalar_criterion_errors(self):
        _, lowered, rows = self.lower({
            'A1': '=SUMIFS(\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!B2:B2,">0")',
            'A2': '=COUNTIFS(\'Ledger Renamed\'!B2:B2,">0")',
            'A3': '=SUMIFS(\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!B2:B2,"<"&\'Ledger Renamed\'!D2)',
            'A4': '=COUNTIFS(\'Ledger Renamed\'!B2:B2,"<"&\'Ledger Renamed\'!D2)',
            'A5': '=SUMIFS(\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!D3)',
            'A6': '=COUNTIFS(\'Ledger Renamed\'!B2:B2,\'Ledger Renamed\'!D3)',
        })
        expected = {'A1': 3, 'A2': 1, 'A3': 0, 'A4': 0, 'A5': '#DIV/0!', 'A6': '#DIV/0!'}
        self.assert_outputs(lowered, rows, expected)
        rows = [rows[1], {'Account': 'c', 'Amount': 4, 'Rate': 1, 'Row ID': 'appended'}, rows[0]]
        self.assert_outputs(lowered, rows, expected)


if __name__ == '__main__':
    unittest.main()
