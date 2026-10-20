import copy,importlib.util,sys,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts/lib'))
class PromotionTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('promotion'),'promotion adapter missing')
        import promotion
        return promotion
    def actions(self,current):
        # Unit fixture only; production callers must actually inspect the SDK.
        return {'documentId':current['workbookId'],'version':current['documentVersion'],'origin':'sigma-sdk-readonly','complete':True,'evidenceId':'unit-sdk',
            'elements':{e['id']:{'elementId':e['id'],'complete':True,'inspection':'list_action_sequences','sequences':[]} for e in current['contents']['elements']}}
    def fixture(self):
        p=self.module()
        current={'workbookId':'original','documentVersion':7,'url':'https://example.com/original',
            'contents':{'kind':'workbook','pages':[{'id':'p','name':'Existing'}],
            'elements':[{'id':'input','kind':'input-table','name':'Input','source':{'kind':'empty','connectionId':'conn'},
                'columns':[{'id':'key','name':'Key','type':'text'},{'id':'value','name':'Value','type':'number'},{'id':'ID'}]},
                {'id':'foreign','kind':'text','body':'keep'}],'layout':'<Page id="p"><Element elementId="input" /></Page>'}}
        rows={'input':{'complete':True,'rowCount':1,'rowIdColumn':'ID','keyColumns':['Key'],
            'rows':[{'ID':'row1','Key':'a','Value':10}],
            'binding':{'elementId':'input','pageId':'p','columnIds':{'Key':'key','Value':'value'},
                'columnTypes':{'Key':'text','Value':'number'},'columnClasses':{'Key':'LiteralEditableColumn','Value':'LiteralEditableColumn'},
                'sdkColumns':[{'id':'key','name':'Key','columnClass':'LiteralEditableColumn','type':'text'},
                    {'id':'value','name':'Value','columnClass':'LiteralEditableColumn','type':'number'},
                    {'id':'ID','name':'ID','columnClass':'FormulaColumn','formula':'[New Input Table/ID]','type':None}]}}}
        sdk={'input':{'elementId':'input','editMode':'published-explore','columns':[['key',True,None],['value',True,['number',2,None,None]],['ID',False,None]]}}
        scope={'documentId':'original','version':7,'sourceHash':'a'*64,'planHash':'b'*64,
            'affectedTables':['input'],'affectedDataModels':[],'owner':'reviewer'}
        scope['sdkActionInventory']=self.actions(current)
        return p,current,rows,sdk,scope
    def test_backup_integrity_full_rows_and_sdk_visibility_required(self):
        p,current,rows,sdk,scope=self.fixture();backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        self.assertTrue(p.validate_backup(backup))
        for bad in ({**rows,'input':{**rows['input'],'complete':False}}, {**rows,'input':{**rows['input'],'rowCount':2}}):
            with self.assertRaises(ValueError):p.build_promotion_backup(scope,current,bad,sdk,{})
        with self.assertRaises(ValueError):p.build_promotion_backup(scope,current,rows,{}, {})
        tampered=copy.deepcopy(backup);tampered['rows']['input']['rows'][0]['Value']=11
        with self.assertRaises(ValueError):p.validate_backup(tampered)
    def test_rollback_restores_business_rows_document_and_sdk_without_unowned_deletion(self):
        p,current,rows,sdk,scope=self.fixture();backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        fresh={**current,'documentVersion':8};after=copy.deepcopy(rows);after['input']['rows'][0]['Value']=99
        plan=p.prepare_rollback(backup,fresh,after)
        self.assertEqual(plan['documentRequest']['body']['contents'],current['contents'])
        self.assertEqual(plan['documentRequest']['body']['documentVersion'],8)
        self.assertIn('table.update_rows',plan['rowPrograms']['input']);self.assertIn('10',plan['rowPrograms']['input'])
        self.assertIn('set_visible(False)',plan['sdkPrograms']['input'])
        self.assertEqual(plan['order'],['rows','dataModels','document','sdkActions','sdkDisplay','readback'])
        self.assertNotIn('set_edit_mode',plan['sdkPrograms']['input'])
        explicit=p.prepare_rollback(backup,fresh,after,changed_edit_modes={'input':'draft'})
        self.assertIn("table.set_edit_mode('published-explore')",explicit['sdkPrograms']['input'])

        after['input']['rows'].append({'ID':'unrelated','Key':'b','Value':0});after['input']['rowCount']=2
        with self.assertRaises(ValueError):p.prepare_rollback(backup,fresh,after)
        plan=p.prepare_rollback(backup,fresh,after,owned_added_rows={'input':['unrelated']})
        self.assertIn("table.delete_rows(['unrelated'])",plan['rowPrograms']['input'])
    def test_promotion_preserves_url_versions_foreign_contents_and_requires_complete_acceptance(self):
        p,current,rows,sdk,scope=self.fixture();backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        candidate=copy.deepcopy(current['contents']);candidate['elements'].append({'id':'new','kind':'text','body':'Ready'})
        review={'owner':'reviewer','backupHash':backup['backupHash'],'replacementIds':[],
            'candidateHash':p._digest(candidate),'allowNewIds':['new'],'targetDocumentId':'original'}
        with self.assertRaises(ValueError):p.prepare_promotion(backup,current,candidate,review,{'status':'BLOCKED'})
        from acceptance import seal_matrix,append_evidence
        matrix=seal_matrix({'version':1,'workflows':{'sales':{'sourceHash':'a'*64,'planHash':'b'*64,'documentId':'stage','version':6,'url':'https://example.com/stage','sourceFormulaCount':170,'sourceStatus':'SUPPORTED'}},'cases':[{'id':'check','workflow':'sales','mode':'query','count':1,'mandatory':True,'oracle':{'origin':'independent-oracle','reviewer':'independent','sourceHash':'a'*64,'checks':[{'key':'x','expected':1,'comparison':'exact'}]}}]})
        ledger=append_evidence([],matrix,{'caseId':'check','sourceHash':'a'*64,'planHash':'b'*64,'documentId':'stage','version':6,'origin':'sigma-query','executionId':'runtime','evidenceId':'proof','references':['query'],'receipt':{'status':'COMPLETED','executionId':'runtime','documentId':'stage','version':6},'actual':[{'key':'x','value':1}]})
        acceptance={'matrix':matrix,'ledger':ledger,'finalReview':{'status':'APPROVED','reviewer':'independent','matrixHash':matrix['matrixHash'],'ledgerHead':ledger[-1]['recordHash'],'branchCommit':'commit','candidateHash':p._digest(candidate),'targetDocumentId':'original'}}
        with self.assertRaisesRegex(ValueError,'all five workflows'):
            p.prepare_promotion(backup,current,candidate,review,acceptance)
        # A forged status or a reduced one-case ledger can never unlock promotion.
        changed=copy.deepcopy(candidate);changed['elements'][1]['body']='overwrite foreign'
        with self.assertRaises(ValueError):p.prepare_promotion(backup,current,changed,{**review,'candidateHash':p._digest(changed)},acceptance)
    def test_full_mandatory_evidence_unlocks_exact_original_candidate_and_dryrun_only(self):
        from acceptance import build_acceptance_matrix,seal_matrix,append_evidence
        from test_acceptance import AcceptanceTests
        p,current,rows,sdk,scope=self.fixture()
        manifests,targets,receipts=AcceptanceTests().bound_matrix_inputs()
        m=build_acceptance_matrix(manifests,targets,receipts=receipts);m.pop('matrixHash')
        for case in m['cases']:
            keys=case.get('chartKeys') or [str(i) for i in range(case['count'])]
            checks=([{'key':'source-validation-and-cancel','expected':case['validationContract']['expectedOutcomes'],'comparison':'exact'}]
                    if case['id'].endswith('-zero-and-invalid-entry') else
                    [{'key':key,'expected':True,'comparison':'exact'} for key in keys])
            case['oracle']={'origin':'independent-oracle','reviewer':'unit-test-oracle',
                'sourceHash':m['workflows'][case['workflow']]['sourceHash'],'checks':checks}
        m=seal_matrix(m);ledger=[]
        # Synthetic captures exercise only the validator, never count as live acceptance.
        for case in m['cases']:
            target=m['workflows'][case['workflow']];identity=case['id']
            capture={**{k:target[k] for k in ('sourceHash','planHash','documentId','version')},'caseId':identity,'evidenceId':identity,'executionId':identity,'origin':'sigma-'+case['mode'],
                'references':['unit-fixture'],'receipt':{'status':'COMPLETED','executionId':identity,'documentId':target['documentId'],'version':6},
                'actual':[{'key':key,'value':True} for key in (case.get('chartKeys') or [str(i) for i in range(case['count'])])]}
            if identity.endswith('-staging-writeback-isolation'):capture['isolation']={'identityEvidenceId':'unit','staging':target['inputTables'],'original':{'input':{'connectionId':'conn','tableIdentity':'original-'+case['workflow']}}}
            if case['mode']=='browser':
                role=case.get('role',{'memberId':'unit','accountType':'build','grant':'edit'})
                capture.update(published=True,publishedVersion=6,steps=['unit-fixture'],screenshots=['unit-fixture'],session={'memberId':role['memberId'],'accountType':role['accountType'],'effectiveGrant':role['grant'],'identityEvidenceId':'unit','adminBypass':False,'editInputTables':case.get('editEntitlement',True)})
                if 'role' in case:
                    allowed=case['editEntitlement'];capture['dataEntry']={'inputMode':'published-explore','editDataAvailable':allowed,
                        'writeCount':1 if allowed else 0,'validSaved':allowed}
                if identity.endswith('-zero-and-invalid-entry'):
                    from test_acceptance_c2 import synthetic_validation_bundle
                    capture['dataEntry']={'inputMode':'published-explore','editDataAvailable':True}
                    capture['validationInventory'],capture['validationBranches'],refs=synthetic_validation_bundle(case,target,capture['session'])
                    capture['references'] += refs
                    capture['actual']=[{'key':'source-validation-and-cancel',
                        'value':case['validationContract']['expectedOutcomes']}]
            if case.get('operationRequired'):capture['operationId']='unit-operation'
            if case.get('persistence'):
                capture['persistence']={'before':[{'ID':'row','Key':'a','Value':1}],'changed':[{'ID':'row','Key':'a','Value':2}],'reopened':[{'ID':'row','Key':'a','Value':2}],'restored':[{'ID':'row','Key':'a','Value':1}],
                    'rowIdColumn':'ID','keyColumns':['Key'],'allowedChanges':[{'key':{'Key':'a'},'column':'Value','value':2}],'reopenEvidenceId':'unit','restoreEvidenceId':'unit'}
            ledger=append_evidence(ledger,m,capture)
        scope={**scope,'sourceHash':manifests['sales']['sourceHash'],'planHash':manifests['sales']['planHash']}
        backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        candidate=copy.deepcopy(current['contents']);candidate['elements'].append({'id':'new','kind':'text','body':'Ready'})
        review={'owner':'reviewer','backupHash':backup['backupHash'],'candidateHash':p._digest(candidate),'replacementIds':[],'allowNewIds':['new'],'targetDocumentId':'original'}
        final={'status':'APPROVED','reviewer':'independent','matrixHash':m['matrixHash'],'ledgerHead':ledger[-1]['recordHash'],'branchCommit':'commit','candidateHash':p._digest(candidate),'targetDocumentId':'original'}
        acceptance={'matrix':m,'ledger':ledger,'finalReview':final}
        plan=p.prepare_promotion(backup,current,candidate,review,acceptance)
        self.assertEqual(plan['preservedUrl'],current['url']);self.assertEqual(plan['request']['body']['documentVersion'],7)
        self.assertEqual(plan['request']['path'],'/v2/workbooks/original/contents')
        with self.assertRaises(ValueError):p.finalize_mutation(plan,{**current,'documentVersion':8},{})
        request=p.finalize_mutation(plan,current,{'httpStatus':200,'dryRun':True,'evidenceId':'actual','planHash':plan['promotionHash']},fresh_sdk_actions=self.actions(current))
        self.assertEqual(request['consumeBeforeSend'],plan['promotionHash'])
        with self.assertRaises(ValueError):p.prepare_promotion(backup,{**current,'documentVersion':8},candidate,review,acceptance)
        foreign=copy.deepcopy(candidate);foreign['elements'][1]['body']='replace foreign'
        with self.assertRaises(ValueError):p.prepare_promotion(backup,current,foreign,{**review,'candidateHash':p._digest(foreign)},{**acceptance,'finalReview':{**final,'candidateHash':p._digest(foreign)}})

    def test_affected_data_models_require_full_backup_and_fresh_versioned_restore(self):
        p,current,rows,sdk,scope=self.fixture();scope['affectedDataModels']=['dm']
        models={'dm':{'dataModelId':'dm','documentVersion':4,'complete':True,'contents':{'schemaVersion':1,'kind':'data-model','pages':[]}}}
        backup=p.build_promotion_backup(scope,current,rows,sdk,models)
        with self.assertRaises(ValueError):p.prepare_rollback(backup,current,rows)
        plan=p.prepare_rollback(backup,current,rows,fresh_data_models={'dm':{'dataModelId':'dm','documentVersion':5}})
        request=plan['dataModelRequests']['dm'];self.assertEqual(request['path'],'/v2/dataModels/dm/spec');self.assertEqual(request['body']['documentVersion'],5)
        restored={'dm':{**models['dm'],'documentVersion':6}}
        self.assertEqual(p.verify_restoration(backup,current,rows,sdk,restored,sdk_actions=self.actions(current))['status'],'RESTORED_BUSINESS_VALUES_AND_DOCUMENT_DISPLAY')
        with self.assertRaises(ValueError):p.verify_restoration(backup,current,rows,sdk,{})

    def test_controller_ledger_refuses_replay_and_requires_both_row_and_sdk_readback(self):
        p,current,rows,sdk,scope=self.fixture();backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        state=p.checkpoint_promotion([],backup,{'step':'backup','outcome':'verified','evidenceId':'backup-read','hash':backup['backupHash']})
        with self.assertRaises(ValueError):p.checkpoint_promotion(state,backup,{'step':'backup','outcome':'verified','evidenceId':'backup-read','hash':backup['backupHash']})
        with self.assertRaises(ValueError):p.verify_restoration(backup,{**current,'documentVersion':9},rows,{})
        result=p.verify_restoration(backup,{**current,'documentVersion':9},rows,sdk,sdk_actions=self.actions({**current,'documentVersion':9}))
        self.assertEqual(result['status'],'RESTORED_BUSINESS_VALUES_AND_DOCUMENT_DISPLAY')
if __name__=='__main__':unittest.main()
