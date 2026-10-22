"""Blank date inputs retain Excel's zero comparison semantics in native lanes."""
import copy
import datetime as dt
import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))
from native_exact_plan import build_exact_native_calculation_plan, lower_exact_native_calculations
from source_inputs import build_source_input_plan
from native_graph_reference import NativeGraphReference
from native_formula_reference import evaluate_native
from native_exact_plan import input_scalar
import test_native_calculations


class BlankDateInputTests(unittest.TestCase):
    source = test_native_calculations.NativeCalculationTests.source

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def test_blank_date_numeric_lane_is_zero_and_kind_stays_blank(self):
        for value_type in ('date', 'datetime'):
            with self.subTest(value_type=value_type):
                scalar = input_scalar('Sheet!A1', {'Sheet!A1': {
                    'kind': 'input', 'valueType': value_type}}, lambda _binding: '[Date Input]')
                refs = {'[Date Input]': None}
                self.assertEqual(evaluate_native(scalar.kind, refs), 'z')
                self.assertEqual(evaluate_native(scalar.number, refs), 0)

    def test_blank_datetime_compares_and_counts_as_zero_in_both_epochs(self):
        for epoch, serial in (('1900', 45293), ('1904', 43831)):
            with self.subTest(epoch=epoch):
                path, old = self.source()
                book = load_workbook(path)
                if epoch == '1904':
                    book.epoch = CALENDAR_MAC_1904
                book['Ledger Renamed']['B2'] = None
                book['Ledger Renamed']['B3'] = dt.datetime(2024, 1, 2)
                summary = book['Summary Renamed']
                summary['A1'] = "='Ledger Renamed'!B2=0"
                summary['A2'] = '=COUNTIFS(tblLedger[Amount],0)'
                summary['A3'] = "='Ledger Renamed'!B2<1"
                summary['A4'] = "='Ledger Renamed'!B3"
                book.save(path)
                decisions = copy.deepcopy(old['decisions'])
                decisions['sourceSha256'] = hashlib.sha256(path.read_bytes()).hexdigest()
                feature = decisions['features'][0]
                feature['columnTypes']['Amount'] = 'datetime'
                feature['reviewedBlankColumns'] = ['Amount']
                inputs = build_source_input_plan(path, decisions, connection_id='connection', folder_id='folder')
                plan = build_exact_native_calculation_plan(path, inputs)
                expected = {'A1': True, 'A2': 1, 'A3': True, 'A4': serial}
                for cell, value in expected.items():
                    record = plan['formulas']['Summary Renamed!' + cell]
                    self.assertEqual(record['status'], 'ready-to-bind')
                    self.assertEqual(record['localExpected'], value)
                assigned = {'ledger': {'elementId': 'observed-input', 'name': 'tblLedger',
                                      'rowIdColumn': 'Row ID', 'rows': {'0': 'row-a', '1': 'row-b'}}}
                lowered = lower_exact_native_calculations(plan, assigned, max_columns=8)
                # Native date columns are read back as timestamps. Keep the raw
                # blank intact while exercising the emitted scalar and row paths.
                rows = [{'Account': 'a', 'Amount': None, 'Rate': 2, 'Row ID': 'row-a'},
                        {'Account': 'b', 'Amount': dt.datetime(2024, 1, 2), 'Rate': 1, 'Row ID': 'row-b'}]
                graph = NativeGraphReference({'observed-input': ('tblLedger', rows)}).materialize(lowered['shards'])
                self.assertIsNone(rows[0]['Amount'])
                for cell, value in expected.items():
                    result = graph.output(lowered, 'Summary Renamed!' + cell)
                    self.assertEqual(result['error'], '')
                    lane = 'logical' if type(value) is bool else 'number'
                    self.assertEqual(result[lane], value)


if __name__ == '__main__':
    unittest.main()
