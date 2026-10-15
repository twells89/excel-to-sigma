"""Offline evidence validation; fixture captures never establish real acceptance."""
import copy,importlib.util,json,sys,tempfile,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts/lib'))

class AcceptanceTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('acceptance'),'acceptance adapter missing')
        import acceptance
        return acceptance
    def matrix(self):
        p=self.module()
        matrix={'version':1,'workflows':{'sales':{'sourceHash':'a'*64,'planHash':'b'*64,'documentId':'stage',
            'version':6,'url':'https://example.com/stage','sourceFormulaCount':170,'sourceStatus':'REVIEW_REQUIRED'}},
            'cases':[{'id':'sales-values','workflow':'sales','mode':'query','count':2,'mandatory':True,
                'oracle':{'checks':[{'key':'A','expected':100.01,'comparison':'money'},{'key':'B','expected':'Open','comparison':'exact'}],
                    'reviewer':'independent','origin':'independent-oracle','sourceHash':'a'*64}}]}
        return p.seal_matrix(matrix)
    def capture(self,matrix):
        return {'caseId':'sales-values','evidenceId':'runtime-evidence','origin':'sigma-query','executionId':'query1',
            'sourceHash':'a'*64,'planHash':'b'*64,'documentId':'stage','version':6,
            'actual':[{'key':'A','value':100.01},{'key':'B','value':'Open'}],
            'receipt':{'executionId':'query1','documentId':'stage','version':6,'status':'COMPLETED'},'references':[{'kind':'query','id':'query1'}]}
    def test_append_only_chain_computes_outcome_and_rejects_stale_local_and_incomplete_captures(self):
        p=self.module();matrix=self.matrix();capture=self.capture(matrix)
        ledger=p.append_evidence([],matrix,capture)
        self.assertEqual(ledger[0]['status'],'PASSED');self.assertTrue(p.validate_ledger(matrix,ledger))
        for edit in ({'version':5},{'origin':'local-simulation'},{'actual':[{'key':'A','value':100.01}]},
                {'receipt':{'executionId':'wrong','documentId':'stage','version':6,'status':'COMPLETED'}}):
            with self.subTest(edit=edit),self.assertRaises(ValueError):p.append_evidence(ledger,matrix,{**capture,**edit,'evidenceId':'new'})
        tampered=copy.deepcopy(ledger);tampered[0]['capture']['actual'][0]['value']=99
        with self.assertRaises(ValueError):p.validate_ledger(matrix,tampered)
        with self.assertRaises(ValueError):p.append_evidence(ledger,matrix,capture)
    def test_money_cents_ratios_dates_and_unresolved_denominators(self):
        p=self.module();self.assertTrue(p.compare_value(100.01,100.010000000000005,'money'))
        self.assertFalse(p.compare_value(100.01,100.02,'money'))
        self.assertFalse(p.compare_value(None,0,'money'));self.assertFalse(p.compare_value(float('nan'),0,'ratio',precision=6))
        m=self.matrix();capture=self.capture(m);capture['actual'][0]['value']=100.02
        ledger=p.append_evidence([],m,capture);report=p.build_workflow_reports(m,ledger)['sales']
        self.assertEqual(report['verdict'],'BLOCKED');self.assertEqual(report['passed'],0);self.assertEqual(report['mandatory'],1)
    def test_real_role_session_required_and_admin_bypass_or_service_writes_refused(self):
        p=self.module();m=self.matrix();m.pop('matrixHash')
        m['cases'][0].update(mode='browser',role={'memberId':'viewer','accountType':'view','grant':'view'},
            inputMode='published-explore',editEntitlement=False)
        m=p.seal_matrix(m);capture=self.capture(m)
        capture.update(origin='sigma-browser',published=True,publishedVersion=6,steps=['Open published','Inspect Edit data'],
            session={'memberId':'viewer','accountType':'view','effectiveGrant':'view','adminBypass':False,
                'identityEvidenceId':'banner','editInputTables':False},screenshots=[{'path':'screenshot.png','sha256':'c'*64}],dataEntry={'inputMode':'published-explore','editDataAvailable':False,'writeCount':0})
        self.assertEqual(p.append_evidence([],m,capture)[0]['status'],'PASSED')
        for change in ({'session':{**capture['session'],'memberId':'admin'}},{'session':{**capture['session'],'adminBypass':True}},
                       {'origin':'sigma-query'},{'publishedVersion':5}):
            with self.assertRaises(ValueError):p.append_evidence([],m,{**capture,**change})
    def test_independent_staging_identities_and_secret_exclusion(self):
        p=self.module()
        with self.assertRaises(ValueError):p.assert_isolated_inputs({'stage':{'connectionId':'conn','tableIdentity':'same'}},
            {'original':{'connectionId':'conn','tableIdentity':'same'}})
        self.assertTrue(p.assert_isolated_inputs({'stage':{'connectionId':'conn','tableIdentity':'new'}},
            {'original':{'connectionId':'conn','tableIdentity':'old'}}))
        m=self.matrix();capture=self.capture(m);capture['authorization']='Bearer secret'
        with self.assertRaises(ValueError):p.append_evidence([],m,capture)
    def test_persistence_requires_full_reopen_restore_and_stable_ids(self):
        p=self.module();before=[{'ID':'1','Key':'x','Value':1}];changed=[{'ID':'1','Key':'x','Value':2}]
        self.assertTrue(p.verify_persistence(before,changed,changed,before,row_id='ID',key_columns=['Key'],
            allowed_changes=[{'key':{'Key':'x'},'column':'Value','value':2}]))
        for reopened,restored in [(before,before),(changed,changed),([{'ID':'other','Key':'x','Value':2}],before)]:
            with self.assertRaises(ValueError):p.verify_persistence(before,changed,reopened,restored,row_id='ID',key_columns=['Key'],
                allowed_changes=[{'key':{'Key':'x'},'column':'Value','value':2}])
    def test_entry_permission_uses_entitlement_and_actual_mode(self):
        p=self.module()
        self.assertFalse(p.published_edit_allowed(True,'edit','draft'))
        self.assertFalse(p.published_edit_allowed(False,'edit','published-explore'))
        self.assertFalse(p.published_edit_allowed(True,'view','published-explore'))
        self.assertTrue(p.published_edit_allowed(True,'explore','published-explore'))
        self.assertTrue(p.published_edit_allowed(True,'view','published-all-users'))
        with self.assertRaises(ValueError):p.run_integration(self.matrix(),None,{'synthetic':True},enabled=False)
    def test_cli_appends_atomically_and_failed_replay_preserves_state_and_output(self):
        import subprocess
        p=self.module();m=self.matrix();capture=self.capture(m)
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp);bundle=folder/'bundle.json';state=folder/'ledger.json';out=folder/'result.json'
            bundle.write_text(json.dumps({'matrix':m,'capture':capture}))
            command=[sys.executable,str(ROOT/'scripts/migration-acceptance.py'),'record','--bundle',str(bundle),'--state',str(state),'--out',str(out)]
            first=subprocess.run(command,capture_output=True,text=True);self.assertEqual(first.returncode,0,first.stderr)
            original=state.read_text();out.write_text('preserve output')
            replay=subprocess.run(command,capture_output=True,text=True);self.assertNotEqual(replay.returncode,0)
            self.assertEqual(state.read_text(),original);self.assertEqual(out.read_text(),'preserve output')

if __name__=='__main__':unittest.main()
