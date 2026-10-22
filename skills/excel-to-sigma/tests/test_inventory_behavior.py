"""Pinned source literals and independent offline behavior expectations."""
import copy
import importlib.util
import json
import math
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts' / 'lib'))
if importlib.util.find_spec('inventory_behavior'):
    from inventory_behavior import (InventoryBehaviorError, adjust_demand,
        smooth_demand, snapshot_forecast, refresh_on_hand, select_purchase_orders)

ORACLE = json.loads((ROOT / 'tests/fixtures/inventory-source-oracle.json').read_text())


class InventoryBehaviorTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(importlib.util.find_spec('inventory_behavior'), 'inventory behavior port missing')
        self.demand = copy.deepcopy(ORACLE['demand'])
        self.skus = copy.deepcopy(ORACLE['skus'])

    def test_adjustment_matches_all_72_pinned_source_rows_and_keeps_other_cells(self):
        out = adjust_demand(self.demand, 100)
        self.assertEqual([r['Forecast Units'] for r in out], ORACLE['oracleAdjustment100'])
        self.assertEqual([r['SKU'] for r in out], [r['SKU'] for r in self.demand])
        self.assertEqual([r['Planned Receipts'] for r in out], [r['Planned Receipts'] for r in self.demand])
        self.assertEqual(self.demand, ORACLE['demand'])

    def test_adjustment_uses_vba_bankers_ties_and_accepts_zero_negative_percent(self):
        self.demand[0]['Forecast Units'] = 5
        self.demand[1]['Forecast Units'] = -5
        self.assertEqual([r['Forecast Units'] for r in adjust_demand(self.demand, -50)[:2]], [2, -2])
        self.assertEqual(adjust_demand(self.demand, 0), self.demand)

    def test_adjustment_rejects_bad_percent_and_overflow_before_any_change(self):
        original = copy.deepcopy(self.demand)
        for pct in ('', 'bad', None, float('nan'), float('inf'), 1e308):
            with self.subTest(pct=pct), self.assertRaises(InventoryBehaviorError):
                adjust_demand(self.demand, pct)
        self.assertEqual(self.demand, original)

    def test_smoothing_matches_all_72_pinned_rows_sequentially(self):
        out = smooth_demand(self.demand)
        self.assertEqual([r['Forecast Units'] for r in out], ORACLE['oracleSmoothing03'])
        for i in (0, 12, 24, 36, 48, 60):
            self.assertEqual(out[i]['Forecast Units'], self.demand[i]['Forecast Units'])
        self.assertEqual(self.demand, ORACLE['demand'])

    def test_smoothing_uses_previous_updated_row_only_for_adjacent_same_sku(self):
        rows = copy.deepcopy(self.demand)
        rows[0]['Forecast Units'], rows[1]['Forecast Units'], rows[2]['Forecast Units'] = 0, 10, 10
        rows[3]['SKU'] = 'OTHER'
        self.assertEqual([r['Forecast Units'] for r in smooth_demand(rows)[:4]], [0, 3, 5, 56])

    def test_snapshot_captures_all_72_once_with_shared_operation_time_and_invoker(self):
        rows = snapshot_forecast(self.demand, 'intent-A', '2026-10-20T12:00:00Z', 'planner@example.com')
        self.assertEqual(len(rows), 72)
        self.assertEqual([r['Forecast Units'] for r in rows], [r['Forecast Units'] for r in self.demand])
        self.assertEqual({r['Operation ID'] for r in rows}, {'intent-A'})
        self.assertEqual({r['Committed At'] for r in rows}, {'2026-10-20T12:00:00Z'})
        self.assertEqual({r['Committed By'] for r in rows}, {'planner@example.com'})

    def test_refresh_updates_only_on_hand_and_rejects_bad_keysets(self):
        erp = [{'SKU': row['SKU'], 'On Hand': 0 if i == 0 else row['On Hand'] + 1}
               for i, row in enumerate(self.skus)]
        out = refresh_on_hand(self.skus, erp)
        self.assertEqual(out[0]['On Hand'], 0)
        self.assertEqual([r['Description'] for r in out], [r['Description'] for r in self.skus])
        self.assertEqual([r['On Hand'] for r in self.skus], [r['On Hand'] for r in ORACLE['skus']])
        for invalid in (erp[:-1], erp + [erp[0]], erp[:-1] + [{'SKU': 'UNKNOWN', 'On Hand': 1}]):
            with self.assertRaises(InventoryBehaviorError): refresh_on_hand(self.skus, invalid)

    def test_po_selection_is_exact_order_only_rows_four_to_nine_and_whole_quantities(self):
        rows = copy.deepcopy(ORACLE['projection'])
        rows[0]['Reorder?'], rows[0]['Suggested Qty'] = 'ORDER', 25
        rows[1]['Reorder?'], rows[1]['Suggested Qty'] = 'order', 50
        rows[2]['Reorder?'], rows[2]['Suggested Qty'] = '', 0
        self.assertEqual(select_purchase_orders(rows), [{'Source Row': 4, 'SKU': 'SKU-200', 'Quantity': 25}])
        rows[0]['Suggested Qty'] = 1.5
        with self.assertRaises(InventoryBehaviorError): select_purchase_orders(rows)
        rows[0]['Suggested Qty'] = 2**53 + 2
        with self.assertRaises(InventoryBehaviorError): select_purchase_orders(rows)
        with self.assertRaises(InventoryBehaviorError): select_purchase_orders(rows[:-1])

    def test_null_numeric_cells_coerce_to_zero_and_nonfinite_or_missing_rows_fail(self):
        rows = copy.deepcopy(self.demand)
        rows[0]['Forecast Units'] = None
        self.assertEqual(adjust_demand(rows, 20)[0]['Forecast Units'], 0)
        self.assertEqual(smooth_demand(rows)[0]['Forecast Units'], None)
        rows[0]['Forecast Units'] = math.inf
        with self.assertRaises(InventoryBehaviorError): adjust_demand(rows, 20)
        rows = copy.deepcopy(self.demand)
        rows.pop(4)
        with self.assertRaises(InventoryBehaviorError): snapshot_forecast(rows, 'x', '2026-10-20T12:00:00Z', 'p')


if __name__ == '__main__':
    unittest.main()
