"""The empty append-log exception is narrow and never declares completion."""
import importlib.util
import sys
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'/'lib'))
if importlib.util.find_spec('inventory_completion_gate'):
    from inventory_completion_gate import assess_empty_snapshot_log, InventoryCompletionError


class EmptySnapshotLogGateTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(importlib.util.find_spec('inventory_completion_gate'),'narrow gate missing')
        self.evidence={
            'kind':'inventory-empty-append-log','role':'append-only-snapshots',
            'schema':'SIGMA_WRITE_DB.SIGMA_MIGRATION8_INV_TEST',
            'sourceSha256':'461bee7b71f30822ea4d263ef529723d89dc9f11e7365e6434cefe89619c2f7e',
            'documentId':'actual-doc','version':'v1','elementId':'actual-snapshot-element',
            'rowIdColumnId':'physical-id','columnIds':{'Operation ID':'op-id','Source Row':'row-id',
                'SKU':'sku-id','Forecast Units':'units-id','Week Start':'week-id'},
            'bindingHash':'a'*64,'publishedButtonId':'actual-button',
            'captures':{'baseline':{'origin':'sigma-query','evidenceId':'e0','count':0},
                'first':{'origin':'sigma-query','evidenceId':'e1','count':72},
                'sameIntentRetry':{'origin':'sigma-query','evidenceId':'e2','count':72},
                'newIntent':{'origin':'sigma-query','evidenceId':'e3','count':144},
                'restored':{'origin':'sigma-query','evidenceId':'e4','count':0}},
            'intentIds':{'first':'intent-A','retry':'intent-A','new':'intent-B'},
            'roles':{'planner':{'writeCount':144,'evidenceId':'rp'},
                'viewer':{'writeCount':0,'evidenceId':'rv'}},
            'screenshots':['published','persisted','restored'],
            'productionGates':{'sourceAudit':True,'inputBinding':True,'actionBinding':True,
                'published':True,'persistence':True,'roleSessions':True,'restoration':True}}
        for name,capture in self.evidence['captures'].items():
            capture.update({'documentId':'actual-doc','version':'v1','elementId':'actual-snapshot-element',
                'bindingHash':'a'*64,'rowIdColumnId':'physical-id',
                'rowFingerprint':{'baseline':'0'*64,'first':'1'*64,'sameIntentRetry':'1'*64,
                    'newIntent':'2'*64,'restored':'0'*64}[name]})
            if name not in ('baseline','restored'):
                capture['publishedActionExecutionId']='action-'+name
                capture['writeEvidenceId']='write-'+name

    def test_exact_zero_baseline_and_real_sequence_is_review_eligible_only(self):
        out=assess_empty_snapshot_log(self.evidence)
        self.assertEqual(out['status'],'NARROW_EMPTY_LOG_EVIDENCE_REVIEW_ELIGIBLE')
        self.assertEqual(out['productionStatus'],'BLOCKED_PENDING_INDEPENDENT_LIVE_VERIFICATION')

    def test_any_missing_or_broad_exception_fails(self):
        import copy
        for edit in ('baseline','retry','schema','role','viewer','binding','gate'):
            item=copy.deepcopy(self.evidence)
            if edit=='baseline':item['captures']['baseline']['count']=1
            if edit=='retry':item['captures']['sameIntentRetry']['count']=144
            if edit=='schema':item['schema']='ERP_PROD'
            if edit=='role':item['role']='general-table'
            if edit=='viewer':item['roles']['viewer']['writeCount']=1
            if edit=='binding':item['rowIdColumnId']=''
            if edit=='gate':item['productionGates']['restoration']=False
            with self.subTest(edit=edit),self.assertRaises(InventoryCompletionError):
                assess_empty_snapshot_log(item)


if __name__=='__main__':unittest.main()
