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
    def bound_matrix_inputs(self):
        from presentation import extract_presentation,build_business_input_architecture,build_business_input_plan
        from migration_manifest import build_manifest
        from input_sync import build_artifact_records,_digest
        from test_presentation import research_runtime_plan
        names={'sales':'sales-pipeline','inventory':'inventory-reorder','expense':'expense-approval','forecast':'forecast-model','research':'research-model'}
        manifests={};receipts={};targets={}
        for workflow,name in names.items():
            path=ROOT/'tests/fixtures'/('fixture-'+name+'.xlsx');source=extract_presentation(path)
            arch=build_business_input_architecture(source,connection_id='conn',folder_id='synthetic',research_plan=research_runtime_plan() if workflow=='research' else None)
            manifest=build_manifest(path,build_business_input_plan(source,arch));manifest['target']={'documentId':'stage-'+workflow,'version':6}
            bindings={f:{'documentId':'stage-'+workflow,'version':6,'pageId':'page-'+f,'elementId':'assigned-'+workflow+'-'+f,
                'columnIds':{c['name']:'c'+str(i) for i,c in enumerate(spec['columns'])}} for f,spec in arch['inputs'].items()}
            receipt={'status':'APPROVED_STAGING','sourceHash':manifest['sourceHash'],'planHash':manifest['planHash'],'targetBindings':bindings}
            receipt['artifacts']=build_artifact_records(manifest,arch,bindings)
            tables={f:{**{k:b[k] for k in ('documentId','version','pageId','elementId')},'connectionId':'conn','tableIdentity':'warehouse-'+workflow+'-'+f,
                'targetBindingHash':_digest(b),'identityEvidenceId':'actual-'+f} for f,b in bindings.items()}
            manifests[workflow]=manifest;receipts[workflow]=receipt
            targets[workflow]={'documentId':'stage-'+workflow,'version':6,'url':'https://example.com/'+workflow,'sourceStatus':'SUPPORTED','inputTables':tables}
        return manifests,targets,receipts

    def test_reduced_approved_component_report_never_claims_migration_complete(self):
        p=self.module();m=self.matrix();m.pop('matrixHash');m['workflows']['sales']['sourceStatus']='SUPPORTED';m=p.seal_matrix(m)
        ledger=p.append_evidence([],m,self.capture(m))
        report=p.build_workflow_reports(m,ledger,final_review={'status':'APPROVED','reviewer':'independent','matrixHash':m['matrixHash'],'ledgerHead':ledger[-1]['recordHash'],'branchCommit':'commit'})['sales']
        self.assertEqual(report['passed'],1);self.assertNotEqual(report['verdict'],'VERIFIED_COMPLETE')
        self.assertNotEqual(report['verdict'],'STAGING_ACCEPTANCE_PASSED_REQUIRES_FINAL_REVIEW')
        self.assertEqual(report['coverageStatus'],'COMPONENT_ONLY');self.assertTrue(report['completionBlockers'])
        import subprocess
        with tempfile.TemporaryDirectory() as tmp:
            bundle=Path(tmp)/'bundle.json';out=Path(tmp)/'reports.json'
            bundle.write_text(json.dumps({'matrix':m,'ledger':ledger,'final_review':{'status':'APPROVED','reviewer':'independent','matrixHash':m['matrixHash'],'ledgerHead':ledger[-1]['recordHash'],'branchCommit':'commit'}}))
            result=subprocess.run([sys.executable,str(ROOT/'scripts/migration-acceptance.py'),'reports','--bundle',str(bundle),'--out',str(out)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            actual=json.loads(out.read_text());self.assertEqual(actual['reports']['sales']['verdict'],'BLOCKED')
            self.assertIn('Completion blocked:',actual['markdown'])


    def test_noop_and_cancelling_persistence_edits_are_not_actual_saved_changes(self):
        p=self.module();rows=[{'ID':'row1','Key':'a','Value':1}]
        for edits in ([{'key':{'Key':'a'},'column':'Value','value':1}],
            [{'key':{'Key':'a'},'column':'Value','value':2},{'key':{'Key':'a'},'column':'Value','value':1}]):
            with self.subTest(edits=edits),self.assertRaises(ValueError):p.verify_persistence(rows,rows,rows,rows,row_id='ID',key_columns=['Key'],allowed_changes=edits)

    def test_isolation_inventory_cannot_omit_planned_business_inputs(self):
        p=self.module();manifests,targets,receipts=self.bound_matrix_inputs()
        for target in targets.values():target['inputTables']={'unrelated-helper':{'connectionId':'conn','tableIdentity':'isolated-helper-'+target['documentId']}}
        with self.assertRaises(ValueError):p.build_acceptance_matrix(manifests,targets,receipts=receipts)

    def test_complete_assigned_inventory_and_exact_integration_scope(self):
        p=self.module();manifests,targets,receipts=self.bound_matrix_inputs()
        matrix=p.build_acceptance_matrix(manifests,targets,receipts=receipts)
        self.assertTrue(p.validate_production_matrix(matrix))
        workflow='sales';features=list(targets[workflow]['inputTables']);feature=features[0]
        for kind in ('missing','extra','elementId','pageId','documentId','version','connectionId','targetBindingHash','identityEvidenceId','duplicate-storage'):
            changed=copy.deepcopy(targets[workflow])
            if kind=='missing':changed['inputTables'].pop(feature)
            elif kind=='extra':changed['inputTables']['foreign']=copy.deepcopy(changed['inputTables'][feature])
            elif kind=='duplicate-storage':changed['inputTables'][features[1]]['tableIdentity']=changed['inputTables'][feature]['tableIdentity']
            else:changed['inputTables'][feature][kind]=5 if kind=='version' else 'foreign' if kind!='identityEvidenceId' else None
            with self.subTest(kind=kind),self.assertRaises(ValueError):p._bound_inputs(changed,manifests[workflow],receipts[workflow])
        forged=copy.deepcopy(manifests[workflow]);forged['planProjection']['sharedInputArchitecture']['inputs'].pop(feature)
        with self.assertRaises(ValueError):p._bound_inputs(targets[workflow],forged,receipts[workflow])
        receipt=copy.deepcopy(receipts[workflow]);receipt['planHash']='f'*64
        with self.assertRaises(ValueError):p._bound_inputs(targets[workflow],manifests[workflow],receipt)
        changed=copy.deepcopy(matrix);changed.pop('matrixHash');changed['workflows'][workflow]['inputTables'].pop(feature);changed=p.seal_matrix(changed)
        with self.assertRaises(ValueError):p.validate_production_matrix(changed)
        class Adapter:
            calls=0
            def capture(self,request):self.calls+=1;raise RuntimeError('adapter reached after preflight')
        adapter=Adapter();union=p._matrix_inputs(matrix)
        scope={'synthetic':True,'connectionId':'conn','folderId':'synthetic','documentIds':[t['documentId'] for t in targets.values()],
            'identityEvidenceId':'actual-identities','stagingInputs':union,'originalInputs':{'original':{'connectionId':'conn','tableIdentity':'original-only'}}}
        for kind in ('missing','extra','shared-original'):
            bad=copy.deepcopy(scope)
            if kind=='missing':bad['stagingInputs'].pop(next(iter(union)))
            elif kind=='extra':bad['stagingInputs']['unplanned']={'connectionId':'conn','tableIdentity':'unplanned'}
            else:bad['originalInputs']={'original':copy.deepcopy(next(iter(union.values())))}
            with self.subTest(kind=kind),self.assertRaises(ValueError):p.run_integration(matrix,adapter,bad,enabled=True)
        self.assertEqual(adapter.calls,0)
        with self.assertRaisesRegex(RuntimeError,'adapter reached'):p.run_integration(matrix,adapter,scope,enabled=True)
        self.assertEqual(adapter.calls,1)

    def test_persistence_comparison_policy_requires_an_effective_business_delta(self):
        p=self.module();before=[{'ID':'row','Key':'a','Value':1.001}];changed=[{'ID':'row','Key':'a','Value':1.002}]
        with self.assertRaises(ValueError):p.verify_persistence(before,changed,changed,before,row_id='ID',key_columns=['Key'],
            allowed_changes=[{'key':{'Key':'a'},'column':'Value','value':1.002,'comparison':'money'}])
        for value in (False,float('inf'),float('nan')):
            invalid=[{'ID':'row','Key':'a','Value':value}]
            with self.subTest(value=value),self.assertRaises(ValueError):p.verify_persistence(before,invalid,invalid,before,row_id='ID',key_columns=['Key'],
                allowed_changes=[{'key':{'Key':'a'},'column':'Value','value':value}])
        zero=[{'ID':'row','Key':'a','Value':0}]
        self.assertTrue(p.verify_persistence(before,zero,zero,before,row_id='ID',key_columns=['Key'],
            allowed_changes=[{'key':{'Key':'a'},'column':'Value','value':0,'comparison':'money'}]))
        with self.assertRaises(ValueError):p.verify_persistence(before,zero,[{'ID':'row','Key':'a','Value':False}],before,row_id='ID',key_columns=['Key'],
            allowed_changes=[{'key':{'Key':'a'},'column':'Value','value':0}])

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
