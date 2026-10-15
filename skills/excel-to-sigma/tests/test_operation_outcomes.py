import importlib.util,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
class OutcomeTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('operation_outcomes'),'persisted outcome adapter missing')
        import operation_outcomes
        return operation_outcomes
    def fixture(self):
        op='op1';p=self.module();fields=['Workflow','Context','Source Hash','Plan Hash','State','Outcome','Completed At']
        binding={'elementId':'ops','pageId':'audit','columnIds':{x:x.lower().replace(' ','_') for x in fields},
            'sdkColumns':[{'id':x.lower().replace(' ','_'),'name':x,'type':'text','columnClass':'LiteralEditableColumn'} for x in fields]}
        operation={'complete':True,'rowCount':1,'rowIdColumn':'ID','keyColumns':['ID'],'binding':binding,
            'rows':[{'ID':op,'Workflow':'saveScenario','Context':'forecast','Source Hash':'a'*64,'Plan Hash':'b'*64,'State':'PENDING','Outcome':None,'Completed At':None}]}
        expected={'workflow':'saveScenario','operationId':op,'context':'forecast','sourceHash':'a'*64,'planHash':'b'*64,'values':{'Rate':.1},'documentId':'stage','version':4}
        capture={'operationId':op,'context':'forecast','evidenceId':'actualquery','origin':'sigma-query','executionId':'run',
            'documentId':'stage','version':4,'sourceHash':'a'*64,'planHash':'b'*64,'writeOutcome':'confirmed','before':[],
            'after':[{'ID':'saved','Operation ID':op,'Rate':.1}]}
        return p,operation,expected,capture
    def test_outcome_written_only_after_persisted_business_readback(self):
        p,operations,expected,capture=self.fixture()
        result=p.reconcile_operation(expected,capture,operations,completed_at='2026-10-15T12:00:00')
        self.assertEqual(result['businessAudit']['status'],'PASSED_COMPONENT_READBACK')
        self.assertIn("'CONFIRMED'",result['program']);self.assertIn('table.update_rows',result['program'])
        self.assertIn('actualquery',result['program'])
        for edit in ({'version':3},{'operationId':'other'},{'origin':'local-simulation'},{'after':[]}):
            with self.subTest(edit=edit),self.assertRaises(ValueError):p.reconcile_operation(expected,{**capture,**edit},operations,completed_at='2026-10-15T12:00:00')
    def test_outcome_view_is_readonly_and_pending_is_never_success(self):
        p,operations,expected,capture=self.fixture()
        view=p.build_outcome_view({'elementId':'ops','name':'Workflow Operations'})
        formula=next(c['formula'] for c in view['columns'] if c['name']=='Result')
        self.assertIn('CONFIRMED',formula);self.assertIn('Pending verification',formula)
        self.assertNotIn('actions',view);self.assertEqual(view['source']['elementId'],'ops')
if __name__=='__main__':unittest.main()
