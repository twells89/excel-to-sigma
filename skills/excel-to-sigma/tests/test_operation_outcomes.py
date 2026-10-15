import importlib.util,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
class OutcomeTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('operation_outcomes'),'persisted outcome adapter missing')
        import operation_outcomes
        return operation_outcomes
    def fixture(self):
        import copy
        from presentation import extract_presentation,build_business_input_architecture,build_business_input_plan
        from migration_manifest import build_manifest
        from input_sync import build_artifact_records,_digest
        from native_actions import OPERATION_FIELDS,_spec
        root=Path(__file__).resolve().parents[1];path=root/'tests/fixtures/fixture-forecast-model.xlsx'
        source=extract_presentation(path);arch=build_business_input_architecture(source,connection_id='conn',folder_id='synthetic')
        arch['inputs']['workflowOperations']=_spec(next(iter(arch['inputs'].values())),'Workflow Operations',OPERATION_FIELDS)
        self.manifest=build_manifest(path,build_business_input_plan(source,arch));self.manifest['target']={'documentId':'stage','version':4}
        bindings={f:{'documentId':'stage','version':4,'pageId':'audit','elementId':f,'rowIdColumnId':'native-row-id',
            'columnIds':{c['name']:'c'+str(i) for i,c in enumerate(spec['columns'])}} for f,spec in arch['inputs'].items()}
        self.receipt={'status':'APPROVED_STAGING','sourceHash':self.manifest['sourceHash'],'planHash':self.manifest['planHash'],'targetBindings':bindings,
            'artifacts':build_artifact_records(self.manifest,arch,bindings)}
        op='op1';p=self.module();spec=arch['inputs']['workflowOperations'];binding=copy.deepcopy(bindings['workflowOperations'])
        binding.update(columnTypes={c['name']:c['type'] for c in spec['columns']},columnClasses={c['name']:'LiteralEditableColumn' for c in spec['columns']},
            sdkColumns=[{'id':binding['columnIds'][c['name']],'name':c['name'],'type':c['type'],'columnClass':'LiteralEditableColumn'} for c in spec['columns']]+[
                {'id':'native-row-id','name':'Row ID','columnClass':'InputTableIDColumn','type':None,'systemKind':'row_id','formula':None,'elementId':'workflowOperations'}])
        meta={**self.manifest['target'],'sourceHash':self.manifest['sourceHash'],'planHash':self.manifest['planHash']}
        operation={**meta,'origin':'sigma-query','executionId':'operations-before','evidenceId':'operations-before','complete':True,'rowCount':1,'rowIdColumn':'ID','keyColumns':['ID'],'binding':binding,
            'rowIdBinding':{'sourceColumnId':'native-row-id','readbackColumn':'ID','targetBindingHash':_digest(bindings['workflowOperations']),'evidenceHash':'c'*64},
            'rows':[{'ID':op,**{c['name']:None for c in spec['columns']},'Workflow':'saveScenario','Context':'forecast','Source Hash':meta['sourceHash'],'Plan Hash':meta['planHash'],'State':'PENDING'}]}
        other=copy.deepcopy(operation['rows'][0]);other.update(ID='unrelated',State='CONFIRMED',Outcome='old outcome')
        operation['rows'].append(other);operation['rowCount']=2
        expected={**meta,'workflow':'saveScenario','operationId':op,'context':'forecast','values':{'Rate':.1}}
        capture={**meta,'operationId':op,'context':'forecast','evidenceId':'actualquery','origin':'sigma-query','executionId':'run',
            'writeOutcome':'confirmed','before':[],'after':[{'ID':'saved','Operation ID':op,'Rate':.1}]}
        return p,operation,expected,capture
    def test_outcome_written_only_after_persisted_business_readback(self):
        p,operations,expected,capture=self.fixture()
        result=p.reconcile_operation(expected,capture,operations,completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=self.receipt)
        self.assertEqual(result['businessAudit']['status'],'PASSED_COMPONENT_READBACK')
        self.assertIn("'CONFIRMED'",result['program']);self.assertIn('table.update_rows',result['program'])
        self.assertIn('actualquery',result['program'])
        for edit in ({'version':3},{'operationId':'other'},{'origin':'local-simulation'},{'after':[]}):
            with self.subTest(edit=edit),self.assertRaises(ValueError):p.reconcile_operation(expected,{**capture,**edit},operations,completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=self.receipt)
    def test_foreign_and_stale_destination_cannot_receive_confirmed_outcome(self):
        import copy
        p,operations,expected,capture=self.fixture()
        foreign=copy.deepcopy(operations);foreign['binding'].update(pageId='foreign-page',elementId='foreign-table',documentId='wrong-document',version=1)
        with self.assertRaises(ValueError):p.reconcile_operation(expected,capture,foreign,completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=self.receipt)

    def test_exact_final_readback_and_all_destination_guards(self):
        import copy
        p,operations,expected,capture=self.fixture()
        plan=p.reconcile_operation(expected,capture,operations,completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=self.receipt)
        final=copy.deepcopy(operations);final.update(executionId='operations-after',evidenceId='operations-after',rows=list(plan['intendedRows'].values()))
        self.assertEqual(p.verify_outcome_row(plan,final)['status'],'CONFIRMED_COMPONENT_OUTCOME')
        edits=[('binding',{'elementId':'foreign-table'}),('binding',{'pageId':'foreign-page'}),('binding',{'version':3}),
            ('capture',{'version':3}),('capture',{'origin':'local-simulation'}),('capture',{'evidenceId':'operations-before'}),
            ('row',{'Outcome':'prefix '+final['rows'][0]['Outcome']}),('row',{'Completed At':'2026-10-16T12:00:00'}),('row',{'Context':'foreign'}),('other-row',{'Outcome':'unrelated changed'})]
        for location,edit in edits:
            changed=copy.deepcopy(final)
            if location=='binding':changed['binding'].update(edit)
            elif location=='capture':changed.update(edit)
            elif location=='other-row':changed['rows'][1].update(edit)
            else:changed['rows'][0].update(edit)
            with self.subTest(edit=edit),self.assertRaises(ValueError):p.verify_outcome_row(plan,changed)
        changed=copy.deepcopy(plan);changed['intendedRows']['op1']['Outcome']='other'
        with self.assertRaises(ValueError):p.verify_outcome_row(changed,final)
        for edit in ({'version':3},{'origin':'local-simulation'},{'complete':False}):
            with self.subTest(edit=edit),self.assertRaises(ValueError):p.reconcile_operation(expected,capture,{**operations,**edit},completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=self.receipt)
        for kind in ('pageId','elementId','version','columnId','columnType','sdkClass','rowId','receipt'):
            bad=copy.deepcopy(operations);receipt=copy.deepcopy(self.receipt)
            if kind in ('pageId','elementId','version'):bad['binding'][kind]='foreign' if kind!='version' else 1
            elif kind=='columnId':bad['binding']['columnIds']['State']='foreign'
            elif kind=='columnType':bad['binding']['columnTypes']['Completed At']='text'
            elif kind=='sdkClass':bad['binding']['sdkColumns'][0]['columnClass']='FormulaColumn'
            elif kind=='rowId':bad['rowIdBinding']['sourceColumnId']='foreign'
            else:receipt['planHash']='a'*64
            with self.subTest(kind=kind),self.assertRaises(ValueError):p.reconcile_operation(expected,capture,bad,completed_at='2026-10-15T12:00:00',manifest=self.manifest,receipt=receipt)
        with self.assertRaises(ValueError):p.reconcile_operation(expected,capture,operations,completed_at='2026-10-15T12:00:00')

    def test_outcome_view_is_readonly_and_pending_is_never_success(self):
        p,operations,expected,capture=self.fixture()
        view=p.build_outcome_view({'elementId':'ops','name':'Workflow Operations'})
        formula=next(c['formula'] for c in view['columns'] if c['name']=='Result')
        self.assertIn('CONFIRMED',formula);self.assertIn('Pending verification',formula)
        self.assertNotIn('actions',view);self.assertEqual(view['source']['elementId'],'ops')
if __name__=='__main__':unittest.main()
