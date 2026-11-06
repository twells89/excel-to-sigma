"""Source/gate/emitted-effect regression tests; these are not published UI proof."""
import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts' / 'lib'))
from migration_manifest import build_manifest, validate_manifest
from input_sync import build_artifact_records, _digest
from presentation import extract_presentation, build_business_input_architecture, build_business_input_plan
from macro_behavior import analyze_macro_source
from native_actions import (NativeActionError, build_native_action_plan, validate_native_action_plan,
    bind_native_action_target, emit_native_action_program, emit_action_schema,
    checkpoint_action_schema, audit_native_action_readback)

APPROVAL = '''Attribute VB_Name = "ApprovalWorkflow"
Sub CommitExpenseApproval_Click()
Dim targetRow As Long
targetRow = ActiveCell.Row
If targetRow < 2 Then Exit Sub
If ActiveSheet.Name <> "Expense Register" Then Exit Sub
Cells(targetRow, 9).Value = "Approved"
Cells(targetRow, 16).Value = Now
End Sub
'''


def staging(manifest, *, partial=False):
    specs=manifest['planProjection']['sharedInputArchitecture']['inputs']
    bindings={}
    for feature,spec in specs.items():
        cols=spec['columns'][:11] if partial and feature=='expenseBusinessInput' else spec['columns']
        bindings[feature]={'documentId':'doc','version':8,'pageId':'host',
            'elementId':feature,'columnIds':{c['name']:'c'+str(i) for i,c in enumerate(cols)}}
    result={'status':'APPROVED_STAGING','sourceHash':manifest['sourceHash'],
        'planHash':manifest['planHash'],'targetBindings':bindings}
    result['artifacts']=build_artifact_records(manifest,manifest['planProjection']['sharedInputArchitecture'],bindings)
    return result


class TestNativeActions(unittest.TestCase):
    def setup_case(self,workflow='expense', choose=None):
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup);folder=Path(tmp.name)
        module=folder/'approval.bas';module.write_text(APPROVAL)
        analyses=[analyze_macro_source(ROOT/'scripts'/'Sample Macros.bas'),analyze_macro_source(module)]
        fixture=ROOT/'tests'/'fixtures'/('fixture-'+{'expense':'expense-approval','forecast':'forecast-model'}[workflow]+'.xlsx')
        extracted=extract_presentation(fixture)
        arch=build_business_input_architecture(extracted,connection_id='conn',folder_id='folder')
        base=build_manifest(fixture,build_business_input_plan(extracted,arch));base['target']={'documentId':'doc','version':8}
        self.base_manifest=copy.deepcopy(base)
        receipt=staging(base)
        rows={f:{'rowIdColumn':'ID','rows':[{'ID':f+'-'+str(i),**r} for i,r in enumerate(s['rows'])]}
              for f,s in arch['inputs'].items()}
        contents={'workbookId':'doc','documentVersion':8,'contents':{'kind':'workbook',
            'pages':[{'id':'host','name':'Host'}], 'elements':[
                {'id':f,'kind':'input-table','columns':[{'id':cid,'name':n} for n,cid in t['columnIds'].items()]}
                for f,t in receipt['targetBindings'].items()], 'layout':'<Page id="host">'+''.join(
                    '<Element id="'+f+'" />' for f in receipt['targetBindings'])+'</Page>'}}
        proc=analyses[1]['procedures'][0]
        decision={'reviewStatus':'approved','reviewer':'reviewer','decisionId':'approval-v1',
            'procedureId':proc['procedureId'],'procedureSha256':proc['procedureSha256'],
            'mappingKind':'reviewed-replacement','page':'expense-edit','target':'expenseBusinessInput',
            'acknowledgements':['additive-operation-storage','selection-context-safeguards']}
        decisions=choose(analyses) if choose else ({'approval':decision} if workflow=='expense' else {})
        presentation={'pages':{workflow+'-edit':'host', 'forecast-overview':'host'}}
        if choose:
            node={'id':'forecast','kind':'table','columns':[
                {'id':'line','name':'Line Item','formula':'[FJ/Line Item]'},
                {'id':'month','name':'Month','formula':'[FJ/Month]'},
                {'id':'amount','name':'Amount','formula':'[Live calculated forecast amount]'}]}
            contents['contents']['elements'].append(node)
            contents['contents']['layout']=contents['contents']['layout'].replace('</Page>','<Element id="forecast" /></Page>')
            presentation['readTargets']={'liveForecast':{'pageId':'host','elementId':'forecast',
                'columnIds':{'Line Item':'line','Month':'month','Amount':'amount'},
                'keyColumns':['Line Item','Month'],'elementHash':_digest(node)}}
            rows['liveForecast']={'rows':[{'Line Item':'Subscription Revenue','Month':'Jul-26','Amount':123.45}]}
        result=build_native_action_plan(workbook_source=fixture,base_manifest=base,base_receipt=receipt,
            presentation_bindings=presentation,
            current_contents=contents,input_readbacks=rows,macro_analyses=analyses,decisions=decisions)
        return result,contents,rows,analyses

    def test_plan_hash_binds_raw_module_and_additive_schema_preserves_source_data(self):
        plan,contents,rows,analyses=self.setup_case()
        validate_manifest(plan);validate_native_action_plan(plan)
        projection=lambda m:[(f['id'],f['mappingStatus'],f.get('targetMapping')) for f in m['formulaCells']]
        self.assertEqual(projection(plan),projection(self.base_manifest))
        specs=plan['planProjection']['sharedInputArchitecture']['inputs']
        self.assertEqual(len(specs['expenseBusinessInput']['rows']),48)
        self.assertEqual(len(specs['expenseBusinessInput']['columns']),14)
        self.assertEqual(specs['expenseBusinessInput']['columns'][-3:],[
            {'name':'ApprovedAt','type':'datetime','provenance':'macro-column16'},
            {'name':'Last Approval Operation ID','type':'text','provenance':'task4-safeguard'},
            {'name':'Applied Approval Operation IDs','type':'text','provenance':'task4-safeguard'}])
        for before,after in zip(rows['expenseBusinessInput']['rows'],specs['expenseBusinessInput']['rows']):
            self.assertEqual({k:v for k,v in after.items() if k in before}, {k:v for k,v in before.items() if k!='ID'})
        self.assertFalse(specs['workflowOperations']['rows'])
        Path(analyses[1]['sourcePath']).write_text(APPROVAL+"' changed\n")
        with self.assertRaisesRegex(ValueError,'module|source'):
            validate_native_action_plan(plan)

    def test_forecast_keeps_12_independent_rows_and_9_authored_columns(self):
        plan,*_=self.setup_case('forecast')
        spec=plan['planProjection']['sharedInputArchitecture']['inputs']['forecastBusinessInput']
        self.assertEqual((len(spec['rows']),len(spec['columns'])),(12,9))
        self.assertIn('PostToGrid', json.dumps(plan['planProjection']['nativeActionArchitecture']['unresolved']))

    def test_schema_extension_requires_original_48_ids_and_values_unchanged(self):
        plan,contents,rows,_=self.setup_case();receipt=staging(plan,partial=True)
        current=[{'id':'c'+str(i),**c} for i,c in enumerate(plan['planProjection']['sharedInputArchitecture']['inputs']['expenseBusinessInput']['columns'][:11])]
        extension=emit_action_schema(plan,receipt,'expenseBusinessInput',current_columns=current)
        self.assertEqual(len(extension['missingColumns']),3)
        after=copy.deepcopy(rows['expenseBusinessInput']['rows'])
        for r in after:r.update(ApprovedAt=None,**{'Last Approval Operation ID':None,'Applied Approval Operation IDs':None})
        created={c['name']:'new'+str(i) for i,c in enumerate(extension['missingColumns'])}
        proof=checkpoint_action_schema(extension,created,result_version=9,before=rows['expenseBusinessInput'],
            after={'rowIdColumn':'ID','rows':after,'readbackId':'read','elementId':'expenseBusinessInput','version':9})
        self.assertEqual(proof['preservedRowCount'],48)
        self.assertEqual(proof['preservedValueCount'],528)
        self.assertEqual(len(proof['originalEditableColumns']),11)
        self.assertEqual(len(proof['preservationReadbacks']['before']['rows'][0]),12)
        after[0]['Amount']=0
        with self.assertRaisesRegex(NativeActionError,'preserv'):
            checkpoint_action_schema(extension,created,result_version=9,before=rows['expenseBusinessInput'],
                after={'rowIdColumn':'ID','rows':after,'readbackId':'read','elementId':'expenseBusinessInput','version':9})

    def test_review_cannot_route_fixed_row_source_as_selected_record_approval(self):
        from native_actions import _contract
        _,_,_,analyses=self.setup_case()
        proc=copy.deepcopy(analyses[1]['procedures'][0]);proc['writes'][0]['row']=2
        with self.assertRaisesRegex(NativeActionError,'selected-record'):
            _contract('approval',proc,{'target':'expenseBusinessInput'})

    def test_approval_rejects_hidden_helper_side_effects_and_unmodeled_mutation(self):
        from native_actions import _contract
        cases=[('DeleteAudit targetRow', 'Sub DeleteAudit(r As Long)\nCells(r, 10).Value = "Deleted"\nEnd Sub\n'),
            ('MissingHelper targetRow',''),
            ('result = DeleteAudit(targetRow)', 'Function DeleteAudit(r As Long) As Boolean\nCells(r, 10).Value = "Deleted"\nDeleteAudit = True\nEnd Function\n'),
            ('Range("A1").ClearContents',''),('newRow.Range(1, 10).Value = "Deleted"',''),
            ('result = MissingFunction','')]
        for statement,helper in cases:
            with self.subTest(statement=statement),tempfile.TemporaryDirectory() as directory:
                path=Path(directory)/'approval.bas'
                path.write_text(APPROVAL.replace('End Sub',statement+'\nEnd Sub')+helper)
                proc=analyze_macro_source(path)['procedures'][0]
                with self.assertRaises(NativeActionError):
                    _contract('approval',proc,{'target':'expenseBusinessInput'})

    def test_approval_guard_expressions_cannot_hide_local_or_missing_side_effects(self):
        from native_actions import _contract
        helper=('Function DeleteAudit(r As Long) As Boolean\nCells(r, 10).Value = "Deleted"\nDeleteAudit = False\nEnd Function\n')
        cases=[('targetRow < 2 Or DeleteAudit(targetRow)',helper),
            ('(targetRow < 2) Or (Not (DeleteAudit(targetRow)))',helper),
            ('targetRow < 2 Or CBool(DeleteAudit(targetRow))',helper),
            ('targetRow < 2 Or MissingGuard(targetRow)',''),
            ('targetRow < 2 Or MissingGuard',''),
            ('targetRow < 2 Or AuditModule.Check(targetRow)',''),
            ('targetRow < 2 Or Range("A1").ClearContents',''),
            ('targetRow < 2 Or False','')]
        for condition,definition in cases:
            with self.subTest(condition=condition),tempfile.TemporaryDirectory() as directory:
                path=Path(directory)/'approval.bas'
                path.write_text(APPROVAL.replace('If targetRow < 2 Then Exit Sub',
                    'If '+condition+' Then Exit Sub')+definition)
                proc=analyze_macro_source(path)['procedures'][0]
                with self.assertRaises(NativeActionError):
                    _contract('approval',proc,{'target':'expenseBusinessInput'})

    def test_approval_write_worksheet_must_resolve_to_guarded_expense_register(self):
        from native_actions import _contract
        for owner,allowed in [('',True),('ActiveSheet.',True),
                             ('Sheets("Expense Register").',True),
                             ('Worksheets("Expense Register").',True),
                             ('Sheets("Other").',False),('ws.',False),
                             ('OtherWorkbook.Sheets("Expense Register").',False)]:
            with self.subTest(owner=owner),tempfile.TemporaryDirectory() as directory:
                path=Path(directory)/'approval.bas'
                path.write_text(APPROVAL.replace('Cells(targetRow',owner+'Cells(targetRow'))
                proc=analyze_macro_source(path)['procedures'][0]
                if allowed:_contract('approval',proc,{'target':'expenseBusinessInput'})
                else:
                    with self.assertRaises(NativeActionError):
                        _contract('approval',proc,{'target':'expenseBusinessInput'})

    def test_plan_preparation_rejects_narrow_expense_snapshot_for_each_original_field(self):
        plan,contents,readbacks,analyses=self.setup_case()
        base=self.base_manifest;schema=base['planProjection']['sharedInputArchitecture']['inputs']['expenseBusinessInput']['columns']
        for column in schema:
            with self.subTest(omitted=column['name']):
                narrow=copy.deepcopy(readbacks)
                for row in narrow['expenseBusinessInput']['rows']:row.pop(column['name'])
                with self.assertRaisesRegex(NativeActionError,'complete|schema|field'):
                    build_native_action_plan(workbook_source=base['source']['path'],base_manifest=base,
                        base_receipt=staging(base),presentation_bindings={'pages':{'expense-edit':'host'}},
                        current_contents=contents,input_readbacks=narrow,macro_analyses=analyses,
                        decisions=plan['planProjection']['nativeActionArchitecture']['reviewedDecisions'])

    def test_extension_checkpoint_rejects_narrow_before_or_after_even_with_matching_hash(self):
        plan,_,rows,_=self.setup_case();schema=plan['planProjection']['sharedInputArchitecture']['inputs']['expenseBusinessInput']['columns']
        extension=emit_action_schema(plan,staging(plan,partial=True),'expenseBusinessInput',
            current_columns=[{'id':'c'+str(i),**c} for i,c in enumerate(schema[:11])])
        created={c['name']:'new'+str(i) for i,c in enumerate(schema[-3:])}
        before=rows['expenseBusinessInput'];after=copy.deepcopy(before)
        after.update(readbackId='read',elementId='expenseBusinessInput',version=9)
        for row in after['rows']:row.update({c['name']:None for c in schema[-3:]})
        for column in schema[:11]:
            for side in ['before','after','both']:
                with self.subTest(omitted=column['name'],side=side):
                    old,new,ext=copy.deepcopy(before),copy.deepcopy(after),copy.deepcopy(extension)
                    if side in ['before','both']:
                        for row in old['rows']:row.pop(column['name'])
                        ext['requiredBeforeHash']=_digest(old)
                    if side in ['after','both']:
                        for row in new['rows']:row.pop(column['name'])
                    with self.assertRaisesRegex(NativeActionError,'complete|schema|field|preserv'):
                        checkpoint_action_schema(ext,created,result_version=9,before=old,after=new)

    def test_action_binding_and_emitter_reject_narrow_preservation_evidence(self):
        plan,receipt,binding,contents=self.bound_approval()
        _,_,readbacks,_=self.setup_case()
        for row in readbacks['expenseBusinessInput']['rows']:
            row.update({c['name']:None for c in plan['planProjection']['sharedInputArchitecture']['inputs']['expenseBusinessInput']['columns'][-3:]})
            row.pop('Amount')
        readbacks['workflowOperations']={'rowIdColumn':'ID','rows':[]}
        with self.assertRaisesRegex(NativeActionError,'complete|schema|field'):
            bind_native_action_target(plan,receipt,current_contents=contents,input_readbacks=readbacks)
        incomplete_binding=copy.deepcopy(binding)
        snapshot=incomplete_binding['tables']['expenseBusinessInput']['inputReadback']
        for row in snapshot['rows']:row.pop('Amount')
        incomplete_binding['tables']['expenseBusinessInput']['readbackHash']=_digest(snapshot)
        incomplete_binding['bindingHash']=_digest({k:v for k,v in incomplete_binding.items() if k!='bindingHash'})
        incomplete_receipt=copy.deepcopy(receipt);incomplete_receipt['nativeActionBinding']=incomplete_binding
        with self.assertRaisesRegex(NativeActionError,'complete|schema|field'):
            emit_native_action_program(plan,incomplete_receipt,incomplete_binding,current_contents=contents)
        bad=copy.deepcopy(receipt);proof=bad['nativeActionSchemaProofs']['expenseBusinessInput']
        # A legacy proof that merely claims 48 rows cannot authorize emission.
        for key in ['originalEditableColumns','preservationReadbacks','preservedValueCount']:
            proof.pop(key,None)
        proof['preservationHash']=_digest({k:v for k,v in proof.items() if k!='preservationHash'})
        with self.assertRaisesRegex(NativeActionError,'preserv|schema|complete'):
            emit_native_action_program(plan,bad,binding,current_contents=contents)

    def test_binding_rejects_wrong_document_missing_layout_and_ambiguous_ids(self):
        plan,contents,rows,_=self.setup_case();receipt=staging(plan)
        with self.assertRaises(NativeActionError):
            bind_native_action_target(plan,receipt,current_contents={**contents,'workbookId':'other'},input_readbacks=rows)
        bad=copy.deepcopy(contents);bad['contents']['layout']=''
        with self.assertRaises(NativeActionError):
            bind_native_action_target(plan,receipt,current_contents=bad,input_readbacks=rows)

    def bound_approval(self):
        plan,contents,rows,_=self.setup_case();receipt=staging(plan)
        before=copy.deepcopy(rows['expenseBusinessInput'])
        specs=plan['planProjection']['sharedInputArchitecture']['inputs']
        for f,s in specs.items():
            if f in rows:
                for row in rows[f]['rows']:
                    for c in s['columns']:row.setdefault(c['name'],None)
                node=next(e for e in contents['contents']['elements'] if e['id']==f)
                node['columns']=[{'id':cid,'name':n} for n,cid in receipt['targetBindings'][f]['columnIds'].items()]
            if f not in rows:rows[f]={'rowIdColumn':'ID','rows':[]}
            if f not in {e['id'] for e in contents['contents']['elements']}:
                contents['contents']['elements'].append({'id':f,'kind':'input-table','columns':[
                    {'id':cid,'name':n} for n,cid in receipt['targetBindings'][f]['columnIds'].items()]})
                contents['contents']['layout']=contents['contents']['layout'].replace('</Page>','<Element id="'+f+'" /></Page>')
        alias={'id':'ID','name':'ID','columnClass':'FormulaColumn','type':None,'systemKind':None,
            'formula':'[New Input Table/ID]','elementId':'expenseBusinessInput',
            'rowIdBinding':{'kind':'row_id','sourceColumn':'ID','valueType':'text','evidenceHash':'a'*64}}
        receipt['targetBindings']['expenseBusinessInput']['systemColumn']=alias
        receipt['artifacts']=build_artifact_records(plan,plan['planProjection']['sharedInputArchitecture'],receipt['targetBindings'])
        binding=bind_native_action_target(plan,receipt,current_contents=contents,input_readbacks=rows)
        receipt['nativeActionBinding']=binding
        current=[{'id':'c'+str(i),**c} for i,c in enumerate(specs['expenseBusinessInput']['columns'][:11])]
        extension=emit_action_schema(plan,staging(plan,partial=True),'expenseBusinessInput',current_columns=current)
        created={c['name']:receipt['targetBindings']['expenseBusinessInput']['columnIds'][c['name']] for c in specs['expenseBusinessInput']['columns'][-3:]}
        after={**rows['expenseBusinessInput'],'readbackId':'read','elementId':'expenseBusinessInput','version':8}
        receipt['nativeActionSchemaProofs']={'expenseBusinessInput':checkpoint_action_schema(extension,created,
            result_version=8,before=before,after=after)}
        return plan,receipt,binding,contents

    def test_emission_constructs_real_update_only_and_explicit_prepare_retry(self):
        plan,receipt,binding,contents=self.bound_approval()
        result=emit_native_action_program(plan,receipt,binding,current_contents=contents)
        program=result['program'];compile(program,'candidate','exec')
        self.assertIn('insert_row(t_workflowOperations',program)
        self.assertIn('.returns()',program)
        self.assertNotIn('update_rows(t_expenseBusinessInput',program)
        self.assertEqual(result['constructionStatus'],'PARTIAL_REQUIRES_PUBLIC_ACTIONS')
        self.assertTrue(any(r['role']=='confirm' for r in result['publicActionRequirements']))
        self.assertNotIn('insert_row(t_expenseBusinessInput',program)
        self.assertIn('Applied Approval Operation IDs',program)
        self.assertNotIn('count_distinct',program)
        self.assertIn('set_control_value',program)
        self.assertNotIn('on_select',program)
        self.assertTrue(any(r['role']=='selection' for r in result['publicActionRequirements']))
        self.assertIn('pending_intent_count',program)
        self.assertIn('SUPERSEDED_NEW_INTENT',program)
        self.assertNotIn('.execute(',program)
        self.assertNotIn('set_success_toast',program)
        self.assertEqual(result['productionStatus'],'BLOCKED')
        self.assertIn('set_disabled(True)',program)
        enabled=emit_native_action_program(plan,receipt,binding,current_contents=contents,
            staging_review={'status':'APPROVED_STAGING','owner':'reviewer','bindingHash':binding['bindingHash'],
                            'documentId':'doc','version':8,'sourceHash':plan['sourceHash'],'planHash':plan['planHash'],
                            'candidateProgramHash':result['candidateProgramHash']})
        self.assertIn('set_disabled(True)',enabled['program'])
        self.assertNotIn('set_disabled(False)',enabled['program'])
        self.assertEqual(enabled['productionStatus'],'BLOCKED')

    def test_readback_never_reports_success_for_zero_write_wrong_row_or_partial_failure(self):
        expected={'workflow':'approval','operationId':'op-A','rowIdColumn':'ID','keyColumn':'Expense ID',
            'selectedRowId':'physical-2','selectedBusinessKey':'EXP-2','context':'expense-edit',
            'approvedAt':'2026-10-15T12:00:00Z',
            'sourceHash':'a'*64,'planHash':'b'*64,'writeFields':['Approval Status','ApprovedAt',
            'Last Approval Operation ID','Applied Approval Operation IDs']}
        row={'ID':'physical-2','Expense ID':'EXP-2','Approval Status':'Pending','ApprovedAt':None,
             'Last Approval Operation ID':None,'Applied Approval Operation IDs':None,'Amount':100}
        capture={'documentId':'doc','version':8,'context':'expense-edit','selectionCount':1,
            'before':[row], 'after':[dict(row,**{'Approval Status':'Approved','ApprovedAt':'2026-10-15T12:00:00Z',
                'Last Approval Operation ID':'op-A','Applied Approval Operation IDs':'|op-A|'})],
            'writeOutcome':'confirmed','operationId':'op-A','evidenceId':'read-1'}
        report=audit_native_action_readback(expected,capture)
        self.assertEqual(report['status'],'PASSED_COMPONENT_READBACK')
        self.assertEqual(report['productionStatus'],'BLOCKED')
        for change in ({'selectionCount':0},{'selectionCount':2},{'context':'other'},
                       {'after':[row]},{'after':[]},{'writeOutcome':'failed'}):
            bad={**capture,**change}
            self.assertEqual(audit_native_action_readback(expected,bad)['status'],'FAILED')
        bad=copy.deepcopy(capture);bad['after'][0]['Amount']=999
        self.assertEqual(audit_native_action_readback(expected,bad)['status'],'FAILED')

    def test_source_bound_projection_exposes_physical_identity_without_touching_inputs(self):
        from native_actions import emit_input_identity_projection
        plan,receipt,binding,contents=self.bound_approval()
        result=emit_input_identity_projection(plan,receipt,'expenseBusinessInput',current_contents=contents,
            destination_page='host',name='Expense approval selection')
        self.assertIn('.to_source()',result['program'])
        self.assertIn('text(source_row_id)',result['program'])
        self.assertIn('add_table(',result['program'])
        self.assertNotIn('update_rows',result['program'])
        self.assertNotIn('set_visible_columns',result['program'])
        self.assertEqual(result['status'],'PENDING_ASSIGNED_PROJECTION_READBACK')
        with self.assertRaisesRegex(NativeActionError,'page'):
            emit_input_identity_projection(plan,receipt,'expenseBusinessInput',current_contents=contents,
                destination_page='absent',name='bad')

    def test_native_approval_accepts_only_source_bound_single_select_status_conversion(self):
        from input_display import _hash
        plan,receipt,binding,contents=self.bound_approval()
        feature='expenseBusinessInput';target=receipt['targetBindings'][feature]
        source=plan['source']['path'];extracted=extract_presentation(source)
        rule=next(r for r in extracted['validations'] if r['field']=='Approval Status')
        target['columnContracts']={'Approval Status':{'id':target['columnIds']['Approval Status'],
            'columnClass':'SingleSelectEditableColumn','type':'text','allowedValues':rule['emission']['values'],
            'sourceValidationHash':_hash(rule),'sdkEvidenceHash':'a'*64}}
        binding['tables'][feature]['target']=copy.deepcopy(target)
        binding['bindingHash']=_digest({k:v for k,v in binding.items() if k!='bindingHash'})
        receipt['nativeActionBinding']=binding
        receipt['artifacts']=build_artifact_records(plan,plan['planProjection']['sharedInputArchitecture'],receipt['targetBindings'])
        result=emit_native_action_program(plan,receipt,binding,current_contents=contents)
        from native_actions import _c
        self.assertIn('isinstance('+_c(feature,'Approval Status')+', SingleSelectEditableColumn)',result['program'])
        self.assertIn(_c(feature,'Approval Status')+'.allowed_values()',result['program'])
        self.assertIn('isinstance('+_c(feature,'Amount')+', LiteralEditableColumn)',result['program'])

    def test_emission_refuses_missing_extension_preservation_proof_and_stale_current(self):
        plan,receipt,binding,contents=self.bound_approval()
        bad=copy.deepcopy(receipt);bad.pop('nativeActionSchemaProofs',None)
        with self.assertRaisesRegex(NativeActionError,'extension|preserv'):
            emit_native_action_program(plan,bad,binding,current_contents=contents)
        stale=copy.deepcopy(contents);stale['contents']['elements'][0]['columns'][0]['name']='other'
        with self.assertRaisesRegex(NativeActionError,'stale'):
            emit_native_action_program(plan,receipt,binding,current_contents=stale)

    def test_sdk_column_type_narrowing_uses_unique_locals_for_all_effect_keys(self):
        import ast
        plan,staged,binding,contents=self.bound_approval()
        program=emit_native_action_program(plan,staged,binding,current_contents=contents)['program'];tree=ast.parse(program)
        editable=set();physical=set()
        for node in ast.walk(tree):
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id=='isinstance' and isinstance(node.args[1],ast.Name) and node.args[1].id in {'LiteralEditableColumn','SingleSelectEditableColumn','MultiSelectEditableColumn'}:
                self.assertIsInstance(node.args[0],ast.Name,'SDK narrowing cannot persist on a dictionary subscript')
                editable.add(node.args[0].id)
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id=='isinstance' and isinstance(node.args[1],ast.Name) and node.args[1].id=='InputTableIDColumn':physical.add(node.args[0].id)
            if isinstance(node,ast.Assign) and isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Attribute) and node.value.func.attr=='get_column':
                self.assertTrue(all(isinstance(t,ast.Name) for t in node.targets),'get_column must assign a unique narrowed local')
        self.assertTrue(editable)
        for node in ast.walk(tree):
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id in {'insert_row','update_rows'}:
                dictionaries=[arg for arg in node.args if isinstance(arg,ast.Dict)]
                self.assertTrue(dictionaries)
                self.assertTrue(all(isinstance(key,ast.Name) and key.id in editable for key in dictionaries[-1].keys))
                if node.func.id=='update_rows':self.assertTrue(all(isinstance(key,ast.Name) and key.id in physical for key in node.args[1].keys))

    def test_forecast_workflows_emit_effects_and_snapshot_replacement_is_explicit(self):
        def decisions(analyses):
            procs={p['name']:p for p in analyses[0]['procedures']}
            def d(name):
                p=procs[name]
                return {'reviewStatus':'approved','reviewer':'reviewer','decisionId':name+'-v1',
                    'procedureId':p['procedureId'],'procedureSha256':p['procedureSha256'],
                    'mappingKind':'reviewed-replacement','page':'forecast-edit'}
            assumptions=[{'sourceCell':'Assumptions!B'+str(5+i),'feature':'forecastBusinessInput',
                'keyColumn':'Line Item','key':line,'valueColumn':'Monthly Growth','name':'Rate '+str(i)}
                for i,line in enumerate(['Subscription Revenue','Usage Revenue'])]
            apply={**d('ApplyScenario'),'assumptions':assumptions}
            save={**d('SaveScenario_Click'),'assumptions':assumptions}
            snap={**d('CommitForecast_Click'),'overwriteDisposition':'retain H7 as a reviewed replacement',
                'rangeMismatchDisposition':'source 27×6 vs fixture 12 independent lines; use named live projection',
                'projection':[{'line':'Subscription Revenue','period':'Jul-26','name':'Subscription Revenue Jul-26',
                    'feature':'liveForecast','keyColumn':'Line Item','key':'Subscription Revenue','valueColumn':'Amount',
                    'extraKeys':[{'column':'Month','value':'Jul-26'}]}]}
            return {'navigation':{**d('GoToForecast_Click'),'destination':'forecast-overview'},
                    'applyScenario':apply,'saveScenario':save,'snapshot':snap}
        plan,contents,rows,_=self.setup_case('forecast',decisions)
        receipt=staging(plan)
        for f,spec in plan['planProjection']['sharedInputArchitecture']['inputs'].items():
            if f not in rows:rows[f]={'rowIdColumn':'ID','rows':[]}
            node=next((e for e in contents['contents']['elements'] if e['id']==f),None)
            if node is None:
                contents['contents']['elements'].append({'id':f,'kind':'input-table','columns':[
                    {'id':cid,'name':n} for n,cid in receipt['targetBindings'][f]['columnIds'].items()]})
                contents['contents']['layout']=contents['contents']['layout'].replace('</Page>','<Element id="'+f+'" /></Page>')
        binding=bind_native_action_target(plan,receipt,current_contents=contents,input_readbacks=rows)
        receipt['nativeActionBinding']=binding
        result=emit_native_action_program(plan,receipt,binding,current_contents=contents)
        program=result['program'];compile(program,'forecast-candidate','exec')
        self.assertIn('navigate_to(',program)
        self.assertIn('c_liveForecast',program)
        from native_actions import _c
        self.assertIn("const_text('Jul-26'), "+_c('liveForecast','Month'),program)
        self.assertIn('update_rows(t_forecastBusinessInput',program)
        self.assertIn('insert_row(t_savedScenarios',program)
        self.assertIn('insert_row(t_forecastSnapshots',program)
        self.assertIn('Start new saveScenario',program)
        self.assertIn('Captured:Subscription Revenue Jul-26',program)
        self.assertIn(f'eq(as_num(coalesce(count_if(eq({_c("savedScenarios","Operation ID")}, saveScenario_operation)), const_num(0))), const_num(0))',program)
        self.assertIn(f'eq(as_num(coalesce(count_if(eq({_c("forecastSnapshots","Operation ID")}, snapshot_operation)), const_num(0))), const_num(0))',program)
        ap,ar,ab,ac=self.bound_approval()
        self.assertIn('Before:',emit_native_action_program(ap,ar,ab,current_contents=ac)['program'])
        self.assertIn('sourceOverwrites',plan['planProjection']['nativeActionArchitecture']['workflows']['snapshot'])

    def test_append_audit_distinguishes_retry_new_intent_duplicate_and_period_mismatch(self):
        expected={'workflow':'snapshot','operationId':'op1','context':'forecast-edit',
                  'values':{'lineA/Jul':12,'lineA/Aug':13}}
        row={'Operation ID':'op1','lineA/Jul':12,'lineA/Aug':13}
        capture={'evidenceId':'query1','operationId':'op1','context':'forecast-edit',
                 'writeOutcome':'confirmed','before':[],'after':[row]}
        self.assertEqual(audit_native_action_readback(expected,capture)['status'],'PASSED_COMPONENT_READBACK')
        retry={**capture,'before':[row],'writeOutcome':'confirmed-noop'}
        self.assertEqual(audit_native_action_readback(expected,retry)['status'],'PASSED_COMPONENT_READBACK')
        intentional={**expected,'operationId':'op2'}
        capture2={**capture,'operationId':'op2','before':[row], 'after':[row,{**row,'Operation ID':'op2'}]}
        self.assertEqual(audit_native_action_readback(intentional,capture2)['status'],'PASSED_COMPONENT_READBACK')
        for bad in ({**capture,'after':[row,row]}, {**capture,'after':[{**row,'lineA/Jul':13}]},
                    {**capture,'writeOutcome':'failed'}):
            self.assertEqual(audit_native_action_readback(expected,bad)['status'],'FAILED')

    def test_new_approval_timestamp_requires_typed_expectation_or_controller_window(self):
        import datetime as dt
        expected={'workflow':'approval','operationId':'fresh','rowIdColumn':'ID','keyColumn':'Expense ID',
            'selectedRowId':'physical','selectedBusinessKey':'E','context':'expense-edit',
            'writeFields':['Approval Status','ApprovedAt','Last Approval Operation ID','Applied Approval Operation IDs']}
        old={'ID':'physical','Expense ID':'E','Approval Status':'Approved','ApprovedAt':'2000-01-01T00:00:00Z',
            'Last Approval Operation ID':'old','Applied Approval Operation IDs':'|old|'}
        capture={'context':'expense-edit','selectionCount':1,'before':[old],
            'after':[{**old,'Last Approval Operation ID':'fresh','Applied Approval Operation IDs':'|old||fresh|'}],
            'writeOutcome':'confirmed','operationId':'fresh','evidenceId':'read',
            'executionWindow':{'startedAt':'2026-10-15T12:00:00Z','completedAt':'2026-10-15T12:00:02Z'}}
        self.assertEqual(audit_native_action_readback(expected,capture)['status'],'FAILED')
        for value in ['not-a-timestamp',1,True,{},'2026-10-15','2026-10-15T12:00:01',
                      '2026-10-15T11:59:59Z','2026-10-15T12:00:03Z']:
            with self.subTest(timestamp=value):
                capture['after'][0]['ApprovedAt']=value
                self.assertEqual(audit_native_action_readback(expected,capture)['status'],'FAILED')
        for value in ['2026-10-15T12:00:01Z',dt.datetime(2026,10,15,12,0,1,tzinfo=dt.timezone.utc)]:
            capture['after'][0]['ApprovedAt']=value
            self.assertEqual(audit_native_action_readback(expected,capture)['status'],'PASSED_COMPONENT_READBACK')
        capture.pop('executionWindow')
        self.assertEqual(audit_native_action_readback(expected,capture)['status'],'FAILED')
        exact={**expected,'approvedAt':'2026-10-15T12:00:01Z'}
        self.assertEqual(audit_native_action_readback(exact,capture)['status'],'PASSED_COMPONENT_READBACK')
        self.assertEqual(audit_native_action_readback({**exact,'approvedAt':'2026-10-15T12:00:02Z'},capture)['status'],'FAILED')

    def test_old_retry_preserves_timestamp_and_intent_but_new_intent_can_change_it(self):
        expected={'workflow':'approval','operationId':'old','rowIdColumn':'ID','keyColumn':'Expense ID',
            'selectedRowId':'physical','selectedBusinessKey':'E','context':'expense-edit',
            'sourceHash':'a'*64,'planHash':'b'*64,'writeFields':['Approval Status','ApprovedAt',
            'Last Approval Operation ID','Applied Approval Operation IDs']}
        before={'ID':'physical','Expense ID':'E','Approval Status':'Approved','ApprovedAt':'later',
                'Last Approval Operation ID':'new','Applied Approval Operation IDs':'|old||new|'}
        capture={'context':'expense-edit','selectionCount':1,'before':[before],'after':[before],
            'writeOutcome':'confirmed-noop','operationId':'old','evidenceId':'retry'}
        self.assertEqual(audit_native_action_readback(expected,capture)['status'],'PASSED_COMPONENT_READBACK')
        bad=copy.deepcopy(capture);bad['after']=[{**before,'ApprovedAt':'refreshed'}]
        self.assertEqual(audit_native_action_readback(expected,bad)['status'],'FAILED')


    def test_public_elementId_layout_binds_and_ambiguous_or_wrong_owner_refuses(self):
        plan,receipt,_,contents=self.bound_approval()
        rows={f:{'rowIdColumn':'ID','rows':[{'ID':f+'-'+str(i),**r} for i,r in enumerate(spec['rows'])]}
              for f,spec in plan['planProjection']['sharedInputArchitecture']['inputs'].items()}
        contents['contents']['layout']=contents['contents']['layout'].replace('<Element id=', '<Element elementId=')
        binding=bind_native_action_target(plan,receipt,current_contents=contents,input_readbacks=rows)
        self.assertIn('expenseBusinessInput',binding['tables'])
        for layout in ['<Page id="wrong"><Element elementId="expenseBusinessInput" /></Page>',
                       '<Page id="host"><Element id="foreign" elementId="expenseBusinessInput" /></Page>',
                       '<Page id="host"><Element elementId="expenseBusinessInput" /><Element elementId="expenseBusinessInput" /></Page>']:
            bad=copy.deepcopy(contents);bad['contents']['layout']=layout
            with self.assertRaises(ValueError):
                bind_native_action_target(plan,receipt,current_contents=bad,input_readbacks=rows)

    def test_action_controls_and_buttons_have_disjoint_frames_on_shared_and_separate_hosts(self):
        import ast
        import re
        class Frame:
            def __init__(self,page,y=0,height=0):self.page=page;self.top=y;self.size=height
            def y(self):return self.top
            def height(self):return self.size
            def parent_grid_width(self):return 24
            def set_frame(self,x,y,width,height):self.top=y;self.size=height
        class Host:
            def __init__(self):self.items=[Frame(self,30,12)]
            def list_elements(self,flat=False):return list(self.items)
            def add_button(self,label):
                frame=Frame(self);self.items.append(frame);return frame
            def add_text(self,label,**kwargs):return self.add_button(label)
        def control(host,**kwargs):return host.add_button('control')
        for shared in (True,False):
            plan,receipt,_,contents=self.bound_approval()
            if not shared:
                receipt['targetBindings']['workflowOperations']['pageId']='audit'
                receipt['artifacts']=build_artifact_records(plan,plan['planProjection']['sharedInputArchitecture'],receipt['targetBindings'])
                contents['contents']['pages'].append({'id':'audit','name':'Audit'})
                contents['contents']['layout']=contents['contents']['layout'].replace('<Element id="workflowOperations" />','')+'<Page id="audit"><Element elementId="workflowOperations" /></Page>'
                rows={f:{'rowIdColumn':'ID','rows':[{'ID':f+'-'+str(i),**r} for i,r in enumerate(spec['rows'])]}
                      for f,spec in plan['planProjection']['sharedInputArchitecture']['inputs'].items()}
                receipt['nativeActionBinding']=bind_native_action_target(plan,receipt,current_contents=contents,input_readbacks=rows)
            program=emit_native_action_program(plan,receipt,receipt['nativeActionBinding'],current_contents=contents)['program']
            tree=ast.parse(program[program.index('audit_cursor = 1'):]);keep=[]
            for node in tree.body:
                if isinstance(node,ast.For) and isinstance(node.iter,ast.Call) and isinstance(node.iter.func,ast.Attribute) and node.iter.func.attr=='list_elements':keep.append(node)
                elif isinstance(node,ast.Assign):
                    names=[t.id for t in node.targets if isinstance(t,ast.Name)]
                    call=node.value if isinstance(node.value,ast.Call) else None
                    if any(n in ('cursor','audit_cursor') for n in names) or (call and ((isinstance(call.func,ast.Name) and call.func.id in ('add_text_control','add_number_control')) or (isinstance(call.func,ast.Attribute) and call.func.attr in ('add_button','add_text')))):
                        if call and isinstance(call.func,ast.Attribute) and call.func.attr=='add_text':call.keywords=[]
                        keep.append(node)
                elif isinstance(node,ast.Expr) and isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Attribute) and node.value.func.attr=='set_frame':keep.append(node)
            workflow_host=Host();audit_host=workflow_host if shared else Host()
            namespace={'GridFramed':Frame,'add_text_control':control,'add_number_control':control,'h_workflowOperations':audit_host}
            for name in re.findall(r'(\w+)\.(?:add_button|add_text|list_elements)',program):namespace.setdefault(name,workflow_host)
            exec(compile(ast.Module(body=keep,type_ignores=[]),'actual-emitted-placements','exec'),namespace)
            for host in {workflow_host,audit_host}:
                intervals=sorted((e.y(),e.y()+e.height()) for e in host.items)
                self.assertTrue(all(a[1]<=b[0] for a,b in zip(intervals,intervals[1:])),(shared,intervals))


import importlib.util
import unittest.mock
import io
import contextlib
import csv
import types

SCRIPTS=ROOT/'scripts'
DEMO_BAS=SCRIPTS/'demo-forecast-macros.bas'


def _load(filename,name):
    spec=importlib.util.spec_from_file_location(name,SCRIPTS/filename)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    return module


def _walk(value):
    if isinstance(value,dict):
        yield value
        for child in value.values():yield from _walk(child)
    elif isinstance(value,list):
        for child in value:yield from _walk(child)


class TestDemoWorkflows(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import demo_workflows
        cls.dw=demo_workflows
        cls.build=_load('build-forecast-model.py','wf_forecast_builder')
        cls.manifest=demo_workflows.build_workflow_manifest(DEMO_BAS)
        cls.plan={'name':'Forecast fixture','intent':'what-if','connectionId':'conn-1','folderId':'folder',
            'currencyFormat':'$,.0f','pctFormat':'.1%','fyLabel':'FY27',
            'actualMonths':[{'month':'Jan-26','num':1,'end':'2026-01-31'}],
            'forecastMonths':[{'month':'Feb-26','num':2,'end':'2026-02-28'}],
            'lines':[{'line':'Revenue','section':'Revenue','secOrder':1,'lineOrder':1,'kind':'growth','rate':0.1,'juneBase':100.0}],
            'assumptions':[['Revenue growth',0.1,'monthly']],
            'actuals':[{'date':'2026-01-10','month':'Jan-26','line':'Revenue','section':'Revenue','secOrder':1,
                'lineOrder':1,'dept':'Sales','vendor':'Customer','amt':100.0}]}
        cls.els={'glDetail':'gl','spine':'sp','assumptions':'as','metricMap':'mm','planDetail':'pd'}

    def spec(self,workflows=None):
        return self.build.build_wb_spec(copy.deepcopy(self.plan),'dm',self.els,workflows=workflows)

    def test_demo_manifest_routes(self):
        flows={w['macro']:w for w in self.manifest['workflows']}
        self.assertEqual(sorted(flows),['RollForwardMonth_Click','SetScenario','SnapshotVersion_Click','SubmitForApproval_Click'])
        self.assertEqual({m:w['writes'] for m,w in flows.items()},{'RollForwardMonth_Click':True,
            'SetScenario':False,'SnapshotVersion_Click':True,'SubmitForApproval_Click':True})
        self.assertEqual(flows['RollForwardMonth_Click']['effect'],'update-rows')
        self.assertEqual(flows['RollForwardMonth_Click']['target'],{'elementId':'parameters','columns':['Forecast Start Month']})
        self.assertEqual(flows['SnapshotVersion_Click']['effect'],'insert-rows')
        self.assertEqual(flows['SnapshotVersion_Click']['target']['elementId'],'versions')
        self.assertEqual(flows['SubmitForApproval_Click']['effect'],'insert-rows')
        self.assertEqual(flows['SubmitForApproval_Click']['class'],'EXTERNAL_IO')
        self.assertEqual(flows['SubmitForApproval_Click']['target']['elementId'],'approvals')
        self.assertEqual(flows['SetScenario']['effect'],'control')
        self.assertEqual(flows['SetScenario']['target'],{'elementId':'parameters','columns':['Selected Scenario']})
        for w in flows.values():
            self.assertEqual(set(w)>={'macro','class','actionId','label','effect','target','writes'},True)
        disp={d['macro']:d for d in self.manifest['dispositions']}
        self.assertEqual({m for m in disp if m!='SubmitForApproval_Click'},{'RefreshActuals_Click','RecalcAll','FormatReport'})
        self.assertEqual(disp['RefreshActuals_Click']['class'],'EXTERNAL_DATA_REFRESH')
        self.assertEqual(disp['RecalcAll']['class'],'RECALC')
        self.assertEqual(disp['FormatReport']['class'],'FORMAT_COSMETIC')
        self.assertEqual(disp['SubmitForApproval_Click']['disposition'],'email-replaced-never-sent')
        mapped={w['macro'] for w in self.manifest['workflows']}|set(disp)
        self.assertEqual(len(mapped),7)

    def test_submit_never_emails(self):
        for w in self.manifest['workflows']:
            self.assertIn(w['effect'],('update-rows','insert-rows','control'))
            self.assertNotIn('call-api',json.dumps(w).lower())
            self.assertNotIn('email',json.dumps(w['effect']).lower())
        wf=self.dw.workflow_elements(self.manifest,'conn-1','FY27 forecast')
        submit=next(b for b in wf['buttons'] if b['id']=='btnSubmit')
        effects=[e for a in submit['actions'] for e in a['effects']]
        self.assertEqual([e['effect'] for e in effects],['insert-rows'])
        self.assertEqual(effects[0]['tableElementId'],'approvals')
        self.assertNotIn('call-api',json.dumps(wf))
        self.assertNotIn('send-email',json.dumps(wf))

    def test_unrouted_macro_is_dispositioned_not_dropped(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'m.bas'
            p.write_text('Sub Mystery()\n    Dim x As Long\n    x = 1\nEnd Sub\n')
            m=self.dw.build_workflow_manifest(p)
        self.assertEqual(m['workflows'],[])
        self.assertEqual(m['dispositions'][0]['macro'],'Mystery')

    def test_without_workflows_unchanged(self):
        base=self.spec(); again=self.spec(None)
        self.assertEqual(base,again)
        ids={e['id'] for e in _walk(base) if 'id' in e}
        self.assertFalse(ids&{'parameters','versions','approvals','btnRoll','btnSnapshot','btnSubmit','approvalStatus'})

    def test_workflows_added_to_workbook(self):
        spec=self.spec(self.manifest)
        by_id={e['id']:e for pg in spec['pages'] for e in pg['elements']}
        for tid,names in (('parameters',['Selected Scenario','Forecast Start Month']),
                ('versions',['Version','Scenario','Forecast Start Month','FY Revenue','FY Gross Margin','FY EBITDA','Saved By','Saved At']),
                ('approvals',['Request','Submitted By','Approver','Status','Submitted At'])):
            t=by_id[tid]
            self.assertEqual(t['kind'],'input-table')
            self.assertEqual(t['source'],{'kind':'empty','connectionId':'conn-1'})
            self.assertEqual([c['name'] for c in t['columns'] if 'type' in c],names)
            self.assertFalse({c['id'] for c in t['columns']}&{'CREATED_BY','CREATED_AT'})
        self.assertEqual(sorted(k for k in ('btnRoll','btnSnapshot','btnSubmit') if k in by_id),['btnRoll','btnSnapshot','btnSubmit'])
        roll=by_id['btnRoll']['actions'][0]['effects'][0]
        self.assertEqual(roll['effect'],'update-rows')
        self.assertEqual(roll['whichRows'],{'type':'formula','formula':'true'})
        self.assertEqual(roll['values']['pa_start']['formula'],'DateAdd("month", 1, Max([Parameters/Forecast Start Month]))')
        approve=by_id['btnSubmit']['actions'][0]['effects'][0]
        self.assertEqual(approve['values']['ap_status']['value']['value'],'Submitted')
        self.assertIn('FY27',approve['values']['ap_req']['value']['value'])
        self.assertEqual(by_id['approvalStatus']['kind'],'table')
        self.assertEqual(by_id['approvalStatus']['source'],{'kind':'table','elementId':'parameters'})
        wire=self.build.workbook_wire.wire_workbook(spec)
        layout=wire['contents']['layout']
        for eid in ('parameters','versions','approvals','btnRoll','btnSnapshot','btnSubmit','approvalStatus','hdrTitle','hdrBand'):
            self.assertIn(f'elementId="{eid}"',layout)

    def test_snapshot_freezes_headline_figures_from_kpi_formulas(self):
        spec=self.spec(self.manifest)
        by_id={e['id']:e for pg in spec['pages'] for e in pg['elements']}
        eff=by_id['btnSnapshot']['actions'][0]['effects'][0]
        self.assertEqual(eff['effect'],'insert-rows'); self.assertEqual(eff['tableElementId'],'versions')
        cols={c['name']:c['id'] for c in by_id['versions']['columns'] if 'type' in c}
        self.assertEqual(set(eff['values']),set(cols.values()))  # one row covers every Versions column
        f=lambda name:eff['values'][cols[name]]['formula']
        plain=self.spec()   # the plain dashboard's KPI tiles carry the same headline formulas
        tiles={e['id']:e['columns'][0]['formula'] for pg in plain['pages'] for e in pg['elements'] if e['kind']=='kpi-chart'}
        self.assertEqual(f('FY Revenue'),tiles['kpiRev'])
        self.assertIn(f('FY EBITDA'),set(tiles.values()))  # the EBITDA row's own KPI formula, not recomputed
        self.assertTrue(tiles['kpiGM'].startswith('('+f('FY Gross Margin')+')'))
        self.assertNotEqual(f('FY EBITDA'),f('FY Gross Margin'))
        self.assertEqual(f('Scenario'),'Max([Parameters/Selected Scenario])')
        self.assertIn('CurrentUserEmail()',f('Saved By')); self.assertEqual(f('Saved At'),'Now()')

    def test_latest_status_is_row_with_max_submitted_at(self):
        spec=self.spec(self.manifest)
        by_id={e['id']:e for pg in spec['pages'] for e in pg['elements']}
        self.assertEqual(by_id['approvalStatus']['columns'][0]['formula'],
            'Coalesce(Lookup([Approvals/Status], Max([Approvals/Submitted At]), [Approvals/Submitted At]), "Draft")')

    def test_manual_commit_program_final_mode_and_seed(self):
        spec=self.build.workbook_wire.wire_workbook(self.spec(self.manifest))
        seed={'scenario':'Base','start':'2026-02-01'}
        program=self.dw.emit_manual_commit_program(spec,seed)
        modes=[l.split("'")[1] for l in program.splitlines() if l.startswith('t.set_edit_mode(')]
        self.assertEqual(modes.count('draft'),modes.count('published-all-users'))
        self.assertEqual(set(modes),{'draft','published-all-users'})
        pairs=[l for l in program.splitlines() if l.startswith('t.set_edit_mode(')]
        self.assertTrue(all(pairs[i+1]=="t.set_edit_mode('published-all-users')" for i in range(0,len(pairs),2)))
        self.assertIn("c_scn = t.get_column('pa_scn')",program)
        self.assertIn("c_start = t.get_column('pa_start')",program)
        self.assertIn("isinstance(c_scn, LiteralEditableColumn)",program)
        self.assertIn("t.upsert_rows([{c_scn: 'Base', c_start: '2026-02-01'}])",program)
        self.assertGreater(program.index('upsert_rows'),program.rindex("set_edit_mode('published-all-users')"))
        compile(program,'program','exec')

    def test_manual_commit_program_covers_every_input_table(self):
        spec=self.build.workbook_wire.wire_workbook(self.spec(self.manifest))
        program=self.dw.emit_manual_commit_program(spec)
        tables=[e['id'] for e in spec['contents']['elements'] if e['kind']=='input-table']
        self.assertGreaterEqual(len(tables),5)  # rate + manual inputs plus the three workflow tables
        for tid in tables:
            self.assertIn(f"get_element('{tid}')",program)
        self.assertEqual(program.count("t.set_edit_mode('draft')"),len(tables))
        self.assertEqual(program.count("t.set_edit_mode('published-all-users')"),len(tables))
        self.assertEqual(program.count('isinstance(t, InputTable)'),len(tables))
        self.assertIn("get_page('model').get_element('versions')",program)   # plumbing lives on the hidden page
        self.assertIn("get_page('model').get_element('parameters')",program)
        compile(program,'program','exec')

    def test_cli_writes_program_dry_run(self):
        build=self.build
        with tempfile.TemporaryDirectory() as d:
            plan=Path(d)/'plan.json';plan.write_text(json.dumps(self.plan))
            wf=Path(d)/'workflows.json';wf.write_text(json.dumps(self.manifest))
            out=Path(d)/'out';prog=Path(d)/'program.py'
            args=[str(plan),'--out-dir',str(out),'--workflows',str(wf),'--emit-manual-commit-program',str(prog)]
            gate=(Path(d)/'source.xlsx',{'status':'DRAFT','sourceSha256':'0'*64})
            with contextlib.redirect_stdout(io.StringIO()),unittest.mock.patch.object(build,'gate_plan_source',return_value=gate):
                rc=build.main(args)
            self.assertEqual(rc,0)
            self.assertIn("set_edit_mode('published-all-users')",prog.read_text())
            wired=json.loads((out/'wb_spec.json').read_text())
            self.assertIn('btnRoll',{e['id'] for e in wired['contents']['elements']})
            self.assertIn("c_start: '2026-02-01'",prog.read_text())  # plan's first forecast month
            self.assertFalse((out/'parameters_seed.csv').exists())

    @staticmethod
    def run_program(program):
        """Execute the emitted program against a fake Code Mode host; return
        ``{elementId: [[row dict keyed by column id], ...]}`` per upsert_rows call."""
        class LiteralEditableColumn:
            def __init__(self,cid):self.cid=cid
            def set_format_pattern(self,pattern):self.pattern=pattern
            def set_visible(self,visible):self.visible=visible
        class InputTableStyle:   # the SDK's InputTable.table_style() (sdk/sigma/tableStyle.py)
            def set_preset(self,preset):self.preset=preset
            def set_grid_lines(self,lines):self.lines=lines
            def set_autofit_columns(self,show):self.autofit=show
        class InputTable:
            def __init__(self,eid,calls):self.eid=eid;self.calls=calls
            def get_column(self,cid):return LiteralEditableColumn(cid)
            def table_style(self):return InputTableStyle()
            def set_edit_mode(self,mode):pass
            def upsert_rows(self,rows):
                self.calls.setdefault(self.eid,[]).append([{c.cid:v for c,v in r.items()} for r in rows])
        calls={}
        class Page:
            def get_element(self,eid):return InputTable(eid,calls)
        sigma=types.SimpleNamespace(workbook=types.SimpleNamespace(get_page=lambda pid:Page()))
        exec(compile(program,'program','exec'),{'sigma':sigma,'InputTable':InputTable,
                                                 'LiteralEditableColumn':LiteralEditableColumn})
        return calls

    def test_cli_program_seeds_rate_and_manual_tables_from_the_seed_csv_rows(self):
        build=self.build
        plan=json.loads((SCRIPTS/'sample-model-plan.json').read_text())
        with tempfile.TemporaryDirectory() as d:
            planp=Path(d)/'plan.json';planp.write_text(json.dumps(plan))
            wf=Path(d)/'workflows.json';wf.write_text(json.dumps(self.manifest))
            out=Path(d)/'out';prog=Path(d)/'program.py'
            gate=(Path(d)/'source.xlsx',{'status':'DRAFT','sourceSha256':'0'*64})
            with contextlib.redirect_stdout(io.StringIO()),unittest.mock.patch.object(build,'gate_plan_source',return_value=gate):
                rc=build.main([str(planp),'--out-dir',str(out),'--workflows',str(wf),'--emit-manual-commit-program',str(prog)])
            self.assertEqual(rc,0)
            program=prog.read_text()
            with (out/'rate_seed.csv').open(newline='') as f:rate_csv=list(csv.DictReader(f))
            with (out/'manual_seed.csv').open(newline='') as f:man_csv=list(csv.DictReader(f))
        self.assertEqual((len(rate_csv),len(man_csv)),(5,24))
        calls=self.run_program(program)
        # one upsert per editable table, every seed row in that one batch
        self.assertEqual({k:len(v) for k,v in calls.items()},{'parameters':1,'rateInput':1,'manualInput':1})
        rate,manual=calls['rateInput'][0],calls['manualInput'][0]
        self.assertEqual(len(rate),len(rate_csv)); self.assertEqual(len(manual),len(man_csv))
        self.assertEqual([{'LINE_ITEM':r['LINE_ITEM'],'RATE':str(r['RATE'])} for r in rate],rate_csv)
        self.assertEqual([{k:(f'{v:.2f}' if k=='AMOUNT' else str(v)) for k,v in r.items()} for r in manual],man_csv)
        self.assertEqual(rate[0],{'LINE_ITEM':'Subscription Revenue','RATE':float(rate_csv[0]['RATE'])})
        self.assertEqual(manual[0]['LINE_ITEM'],'Professional Services Rev')
        self.assertEqual(manual[0]['AMOUNT'],70000.0)
        self.assertIsInstance(manual[0]['MONTH_NUM'],int)
        # typed checks guard every seeded table and column, and seeds follow every mode switch
        for eid in ('rateInput','manualInput'):
            self.assertIn(f"raise RuntimeError('{eid} is not an InputTable')",program)
        self.assertIn("raise RuntimeError('rateInput.RATE is not an editable column')",program)
        self.assertGreater(program.index('upsert_rows'),program.rindex("set_edit_mode('published-all-users')"))

    def test_cli_program_without_workflows_seeds_only_rate_and_manual(self):
        build=self.build
        plan=json.loads((SCRIPTS/'sample-model-plan.json').read_text())
        with tempfile.TemporaryDirectory() as d:
            planp=Path(d)/'plan.json';planp.write_text(json.dumps(plan))
            out=Path(d)/'out';prog=Path(d)/'program.py'
            gate=(Path(d)/'source.xlsx',{'status':'DRAFT','sourceSha256':'0'*64})
            with contextlib.redirect_stdout(io.StringIO()),unittest.mock.patch.object(build,'gate_plan_source',return_value=gate):
                rc=build.main([str(planp),'--out-dir',str(out),'--emit-manual-commit-program',str(prog)])
            self.assertEqual(rc,0)
            calls=self.run_program(prog.read_text())
        self.assertEqual({k:len(v[0]) for k,v in calls.items()},{'rateInput':5,'manualInput':24})

    def test_program_seed_rejects_more_than_2000_rows_and_unknown_tables(self):
        spec=self.build.workbook_wire.wire_workbook(self.spec(self.manifest))
        rows=[['L',i] for i in range(2001)]
        with self.assertRaisesRegex(self.dw.WorkflowError,'2,000'):
            self.dw.emit_manual_commit_program(spec,table_seeds=[('rateInput',['LINE_ITEM','RATE'],rows)])
        with self.assertRaisesRegex(self.dw.WorkflowError,'nope'):
            self.dw.emit_manual_commit_program(spec,table_seeds=[('nope',['A'],[['x']])])
        program=self.dw.emit_manual_commit_program(spec,table_seeds=[('rateInput',['LINE_ITEM','RATE'],rows[:2000])])
        self.assertEqual(len(self.run_program(program)['rateInput'][0]),2000)

    def test_demo_manifest_cli(self):
        mod=_load('macro-actions.py','wf_macro_actions')
        with tempfile.TemporaryDirectory() as d:
            out=Path(d)/'workflows.json'
            self.assertEqual(mod.main(['demo-manifest',str(DEMO_BAS),'--out',str(out)]),0)
            self.assertEqual(len(json.loads(out.read_text())['workflows']),4)

if __name__=='__main__':unittest.main()
