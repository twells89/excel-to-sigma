"""Offline controller proof. Captures here are synthetic, never staging receipts."""
import copy
import importlib.util
import json
import sys
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts' / 'lib'))
if importlib.util.find_spec('inventory_controller'):
    from inventory_controller import (InventoryControllerError, plan_demand_update,
        plan_refresh, plan_snapshot, plan_purchase_orders, verify_update, verify_append, verify_po_stage,
        audit_po_results, plan_restore_update)

ORACLE = json.loads((ROOT / 'tests/fixtures/inventory-source-oracle.json').read_text())
HASH = ORACLE['sourceSha256']


def capture(rows, table='Demand', version='v1', execution='run1'):
    cols = {'SKU': 'col-sku', 'Source Row': 'col-ordinal', 'Input Row ID': 'col-id'}
    if table == 'Demand': cols.update({'Forecast Units': 'col-units', 'Week Start': 'col-week', 'Planned Receipts': 'col-receipts'})
    if table == 'SKUs': cols.update({'On Hand': 'col-onhand', 'Description': 'col-desc'})
    if table == 'Projection': cols.update({'Reorder?': 'col-order', 'Suggested Qty': 'col-qty'})
    if table == 'Snapshots': cols.update({'Operation ID': 'col-operation', 'Committed At': 'col-at',
        'Week Start': 'col-week', 'Forecast Units': 'col-units', 'Committed By': 'col-by',
        'Payload Fingerprint': 'col-payload'})
    if table == 'PO_STAGE': cols.update({'Operation ID': 'col-operation', 'Quantity': 'col-qty',
        'Batch Fingerprint': 'col-fingerprint'})
    return {'origin': 'sigma-query', 'executionId': execution, 'evidenceId': execution + '-evidence',
            'sourceSha256': HASH, 'documentId': 'synthetic-doc', 'version': version,
            'binding': {'elementId': 'synthetic-' + table, 'rowIdColumn': 'Input Row ID', 'columnIds': cols},
            'rows': [{**copy.deepcopy(r), 'Input Row ID': f'{table.lower()}-{i}'}
                     for i, r in enumerate(rows, 1)]}


class InventoryControllerTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(importlib.util.find_spec('inventory_controller'), 'inventory controller missing')
        self.demand = capture(ORACLE['demand'])
        self.skus = capture(ORACLE['skus'], 'SKUs')

    def test_update_plan_binds_exact_physical_ids_and_requires_fresh_readback(self):
        plan = plan_demand_update(self.demand, 'adjust', 'intent-A', percentage=100)
        self.assertEqual(plan['status'], 'NEEDS_ACTUAL_BINDING_AND_WRITE')
        self.assertEqual(len(plan['updates']), 72)
        self.assertEqual(plan['updates'][0]['rowId'], 'demand-1')
        self.assertEqual(plan['updates'][0]['columnId'], 'col-units')
        self.assertEqual(len(plan['planHash']), 64)
        after = copy.deepcopy(self.demand)
        after['executionId'], after['evidenceId'] = 'run2', 'run2-evidence'
        for row, update in zip(after['rows'], plan['updates']):
            row['Forecast Units'] = update['newValue']
        self.assertEqual(verify_update(plan, after)['status'], 'PERSISTED_COMPONENT_WRITE')
        restore = plan_restore_update(plan, after, 'restore-A')
        self.assertEqual(restore['updates'][0]['newValue'], self.demand['rows'][0]['Forecast Units'])
        restored = copy.deepcopy(self.demand)
        restored['executionId'], restored['evidenceId'] = 'run3', 'run3-evidence'
        self.assertEqual(verify_update(restore, restored)['status'], 'PERSISTED_COMPONENT_WRITE')
        tampered = copy.deepcopy(plan)
        tampered['updates'][0]['newValue'] = -99
        with self.assertRaises(InventoryControllerError): verify_update(tampered, after)
        after['version'] = 'v2'
        with self.assertRaises(InventoryControllerError): verify_update(plan, after)

    def test_duplicate_missing_wrong_row_and_stale_capture_fail_closed(self):
        for mutation in ('duplicate', 'missing', 'wrong', 'stale'):
            c = copy.deepcopy(self.demand)
            if mutation == 'duplicate': c['rows'][1]['Input Row ID'] = c['rows'][0]['Input Row ID']
            if mutation == 'missing': c['rows'][1].pop('Input Row ID')
            if mutation == 'wrong': c['rows'][1]['Source Row'] = 99
            if mutation == 'stale': c['sourceSha256'] = '0' * 64
            with self.subTest(mutation=mutation), self.assertRaises(InventoryControllerError):
                plan_demand_update(c, 'smooth', 'intent-B')

    def test_decimal_percentage_serializes_as_double_in_sealed_plan(self):
        plan=plan_demand_update(self.demand,'adjust','intent-decimal',percentage=Decimal('10'))
        self.assertEqual(plan['percentage'],10.0)
        self.assertEqual(len(plan['planHash']),64)

    def test_refresh_has_exact_inverse_and_description_guard(self):
        erp = [{'SKU': r['SKU'], 'On Hand': r['On Hand'] + 3} for r in ORACLE['skus']]
        erp_capture = {'origin': 'warehouse-query', 'evidenceId': 'synthetic-erp-query',
            'schema': 'SIGMA_WRITE_DB.SIGMA_MIGRATION8_INV_TEST', 'source': 'synthetic-test-schema',
            'snapshotDate': '2026-10-20', 'warehouseCurrentDate': '2026-10-20', 'rows': erp}
        plan = plan_refresh(self.skus, erp_capture, 'intent-ERP')
        self.assertEqual(len(plan['updates']), 6)
        self.assertEqual(plan['updates'][0]['columnId'], 'col-onhand')
        self.assertEqual(plan['inverse'][0]['newValue'], ORACLE['skus'][0]['On Hand'])
        stale=copy.deepcopy(erp_capture);stale['warehouseCurrentDate']='2026-10-21'
        with self.assertRaises(InventoryControllerError): plan_refresh(self.skus, stale, 'intent-ERP')
        after = copy.deepcopy(self.skus)
        after['executionId'], after['evidenceId'] = 'run2', 'run2-evidence'
        for row, update in zip(after['rows'], plan['updates']): row['On Hand'] = update['newValue']
        after['rows'][0]['Description'] = 'corrupt'
        with self.assertRaises(InventoryControllerError): verify_update(plan, after)

    def test_snapshot_same_intent_retry_never_appends_again_new_intent_can(self):
        empty = capture([], 'Snapshots')
        first = plan_snapshot(self.demand, empty, 'intent-S', '2026-10-20T12:00:00Z', 'planner@example.com')
        self.assertEqual(first['status'], 'NEEDS_ACTUAL_BINDING_AND_WRITE')
        self.assertEqual(len(first['appendRows']), 72)
        after = capture(first['appendRows'], 'Snapshots', execution='run2')
        self.assertEqual(verify_append(first, after)['status'], 'PERSISTED_COMPONENT_APPEND')
        retry = plan_snapshot(self.demand, after, 'intent-S', '2026-10-20T12:00:00Z', 'planner@example.com')
        self.assertEqual(retry['status'], 'ALREADY_PRESENT_READBACK_ONLY')
        another = plan_snapshot(self.demand, after, 'intent-T', '2026-10-20T12:05:00Z', 'planner@example.com')
        self.assertEqual(len(another['appendRows']), 72)
        partial = capture(first['appendRows'][:-1], 'Snapshots', execution='partial')
        resumed = plan_snapshot(self.demand, partial, 'intent-S', '2026-10-20T12:00:00Z', 'planner@example.com')
        self.assertEqual(resumed['status'], 'NEEDS_RECOVERY_WRITE')
        self.assertEqual([r['Source Row'] for r in resumed['appendRows']], [73])

    def test_purchase_order_plan_exact_selection_retry_changed_payload_and_partial(self):
        projection = copy.deepcopy(ORACLE['projection'])
        projection[0]['Reorder?'], projection[0]['Suggested Qty'] = 'ORDER', 25
        projection[1]['Reorder?'], projection[1]['Suggested Qty'] = 'ORDER', 0
        current = capture(projection, 'Projection')
        empty = capture([], 'PO_STAGE')
        plan = plan_purchase_orders(current, empty, 'intent-P')
        self.assertEqual([(r['SKU'], r['Quantity']) for r in plan['stageRows']], [('SKU-200', 25), ('SKU-300', 0)])
        results = [{'operation_id':'intent-P','source_row':4,'sku':'SKU-200','quantity':25,
            'status':'STAGED','error_code':''},
            {'operation_id':'intent-P','source_row':5,'sku':'SKU-300','quantity':0,
            'status':'RETRY_PRESENT','error_code':''}]
        self.assertEqual(audit_po_results(plan, results)['status'], 'NEEDS_STAGE_READBACK')
        with self.assertRaises(InventoryControllerError): audit_po_results(plan, results[:1])
        wrong = copy.deepcopy(results); wrong[0]['quantity'] = 25.5
        with self.assertRaises(InventoryControllerError): audit_po_results(plan, wrong)
        failed = copy.deepcopy(results); failed[0]['status'] = 'FAILURE';failed[0]['error_code'] = 'DB_ERROR'
        self.assertEqual(audit_po_results(plan, failed)['status'], 'PROCEDURE_FAILURE')
        staged = capture(plan['stageRows'], 'PO_STAGE', execution='run2')
        self.assertEqual(verify_po_stage(plan, staged)['status'], 'PERSISTED_COMPONENT_PO_STAGE')
        self.assertEqual(plan_purchase_orders(current, staged, 'intent-P')['status'], 'ALREADY_PRESENT_READBACK_ONLY')
        changed = copy.deepcopy(current)
        changed['rows'][0]['Suggested Qty'] = 50
        with self.assertRaises(InventoryControllerError): plan_purchase_orders(changed, staged, 'intent-P')
        partial = capture(plan['stageRows'][:1], 'PO_STAGE', execution='partial')
        resumed = plan_purchase_orders(current, partial, 'intent-P')
        self.assertEqual(resumed['status'], 'NEEDS_RECOVERY_WRITE')
        self.assertEqual([(r['SKU'], r['Quantity']) for r in resumed['stageRows']], [('SKU-300', 0)])
        staged['rows'][0]['Quantity'] = 99
        with self.assertRaises(InventoryControllerError): verify_po_stage(plan, staged)

    def test_no_order_is_a_noop_and_uncertain_write_is_never_success(self):
        projection = capture(ORACLE['projection'], 'Projection')
        plan = plan_purchase_orders(projection, capture([], 'PO_STAGE'), 'intent-0')
        self.assertEqual(plan['status'], 'NO_SELECTION_NO_WRITE')
        update = plan_demand_update(self.demand, 'smooth', 'intent-S')
        with self.assertRaises(InventoryControllerError): verify_update(update, self.demand)


if __name__ == '__main__':
    unittest.main()
