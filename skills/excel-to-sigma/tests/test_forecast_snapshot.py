"""Corrected snapshot contract; synthetic rows are unit data, never live evidence."""
import copy
import importlib.util
from pathlib import Path
import sys
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))


class SnapshotTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('forecast_snapshot'), 'corrected snapshot adapter missing')
        import forecast_snapshot
        return forecast_snapshot

    def definition(self):
        p = self.module()
        return p.build_snapshot_definition(
            Path(__file__).parent / 'fixtures/fixture-forecast-model.xlsx',
            plan_hash='a' * 64, procedure_hash='b' * 64,
            decision={'decisionId': 'reviewed-jul-dec', 'status': 'APPROVED_BUSINESS_SEMANTICS_ONLY',
                      'scope': {'periods': list(p.PERIODS), 'financialLineCount': 18, 'valueCount': 108,
                                'timestamp': 'separate metadata; no financial value overwritten'}})

    def operation(self):
        d = self.definition()
        row = {'nativeRowId': 'operation-A', 'Workflow': 'snapshot', 'Context': 'forecast-edit',
               'State': 'PREPARED', 'Created At': '2026-10-16T12:00:00Z',
               'Source Hash': d['sourceHash'], 'Plan Hash': d['planHash'],
               'Source Procedure Hash': d['procedureHash'], 'Replacement Decision ID': d['decisionId'],
               'Snapshot Definition Hash': d['definitionHash']}
        row.update({c['operationFrozenColumn']: i + .25 for i, c in enumerate(d['cells'])})
        return d, row

    def test_definition_covers_all_18_lines_i_to_n_and_keeps_timestamp_separate(self):
        p = self.module(); d = self.definition()
        self.assertEqual(len(d['cells']), 108)
        self.assertEqual(len(d['inputFields']),96)
        self.assertEqual({c['period'] for c in d['cells']}, set(p.PERIODS))
        self.assertEqual(len({c['sourceRow'] for c in d['cells']}), 18)
        self.assertEqual([c['sourceCell'] for c in d['cells'][:6]], ['I8','J8','K8','L8','M8','N8'])
        self.assertEqual(d['cells'][0]['lineItem'], 'Subscription Revenue')
        self.assertEqual(d['cells'][-1]['lineItem'], 'Net Income')
        self.assertNotIn('Captured At', {c['operationFrozenColumn'] for c in d['cells']})

    def test_partial_retry_uses_same_operation_and_new_intent_retains_old_snapshot(self):
        p = self.module(); d, op = self.operation()
        first = p.reconcile_snapshot_rows(d, op, [])
        self.assertEqual(len(first['missingRows']),108)
        persisted = [{**r, 'nativeRowId': 'row-'+str(i)} for i,r in enumerate(first['missingRows'][:37])]
        retry = p.reconcile_snapshot_rows(d, op, persisted)
        self.assertEqual(len(retry['missingRows']),71)
        persisted += [{**r, 'nativeRowId': 'row-'+str(i+37)} for i,r in enumerate(retry['missingRows'])]
        self.assertEqual(p.reconcile_snapshot_rows(d, op, persisted)['missingRows'],[])
        second = {**op, 'nativeRowId':'operation-B'}
        second[d['cells'][0]['operationFrozenColumn']] = 999
        new = p.reconcile_snapshot_rows(d, second, persisted)
        self.assertEqual(len(new['missingRows']),108)
        self.assertEqual(new['missingRows'][0]['Operation ID'],'operation-B')
        self.assertEqual(persisted[0]['Value'],.25)
        self.assertEqual(first['status'],'COMPONENT_RECONCILIATION_ONLY')
        self.assertEqual(first['productionStatus'],'BLOCKED')

    def test_conflicts_duplicates_foreign_keys_and_metadata_fail_before_any_plan(self):
        p = self.module(); d,op = self.operation()
        rows = [{**r,'nativeRowId':str(i)} for i,r in enumerate(p.reconcile_snapshot_rows(d,op,[])['missingRows'])]
        for field,value in [('Value',999),('Captured At','2026-10-17T12:00:00Z'),('Source Hash','c'*64),
                            ('Source Cell','H8'),('Period','Jun-26'),('Snapshot Key','foreign')]:
            changed = copy.deepcopy(rows);changed[0][field]=value
            with self.subTest(field=field),self.assertRaises(ValueError):p.reconcile_snapshot_rows(d,op,changed)
        with self.assertRaises(ValueError):p.reconcile_snapshot_rows(d,op,rows+[rows[0]])
        with self.assertRaises(ValueError):p.reconcile_snapshot_rows(d,op,[{k:v for k,v in rows[0].items() if k!='nativeRowId'}])
        for value in [None, float('nan'), float('inf'), True, '0.25']:
            changed={**op,d['cells'][0]['operationFrozenColumn']:value}
            with self.subTest(value=value),self.assertRaises(ValueError):p.reconcile_snapshot_rows(d,changed,[])

    def test_persisted_numeric_and_timestamp_representations_preserve_values(self):
        p=self.module()
        d,op=self.operation();op[d['cells'][0]['operationFrozenColumn']]=12.0
        row=p.reconcile_snapshot_rows(d,op,[])['missingRows'][0]
        persisted={**row,'nativeRowId':'stored','Value':12,'Source Row':8.0,
                   'Captured At':'2026-10-16T08:00:00-04:00'}
        self.assertEqual(p.reconcile_snapshot_rows(d,op,[persisted])['existingCount'],1)

    def test_reconciliation_requires_fresh_actual_target_bound_complete_captures(self):
        p = self.module(); d,op = self.operation()
        meta={'origin':'sigma-query','documentId':'stage','version':19,'sourceHash':d['sourceHash'],
              'planHash':d['planHash'],'complete':True,'rowCount':1,'evidenceId':'ops-read','executionId':'ops-run',
              'elementId':'operations','rowIdColumnId':'native-ops-id'}
        operations={**meta,'rows':[op]}
        snapshots={**meta,'rowCount':0,'rows':[],'elementId':'snapshots','rowIdColumnId':'native-snapshot-id',
                   'evidenceId':'snapshot-read','executionId':'snapshot-run'}
        target={'documentId':'stage','version':19,'operationsElementId':'operations','operationsRowIdColumnId':'native-ops-id',
                'snapshotsElementId':'snapshots','snapshotsRowIdColumnId':'native-snapshot-id'}
        result=p.plan_snapshot_reconciliation(d,'operation-A',operations,snapshots,target)
        self.assertEqual(result['operationId'],'operation-A')
        self.assertEqual(result['operationReadbackHash'],p.digest(operations))
        for delta in [{'complete':False},{'version':18},{'elementId':'foreign'},{'origin':'local-simulation'},
                      {'rowCount':1},{'rowIdColumnId':'alias-id'}]:
            with self.subTest(delta=delta),self.assertRaises(ValueError):
                p.plan_snapshot_reconciliation(d,'operation-A',operations,{**snapshots,**delta},target)

    def test_reconciliation_rejects_unbound_targets_even_when_capture_headers_match(self):
        p=self.module();d,op=self.operation()
        target={'documentId':'stage','version':20,'operationsElementId':'operations','operationsRowIdColumnId':'native-ops-id',
                'snapshotsElementId':'snapshots','snapshotsRowIdColumnId':'native-snapshot-id'}
        meta={'origin':'sigma-query','documentId':'stage','version':20,'sourceHash':d['sourceHash'],
              'planHash':d['planHash'],'complete':True,'evidenceId':'actual-read','executionId':'actual-run'}
        operations={**meta,'rowCount':1,'rows':[op],'elementId':'operations','rowIdColumnId':'native-ops-id'}
        snapshots={**meta,'rowCount':0,'rows':[],'elementId':'snapshots','rowIdColumnId':'native-snapshot-id'}
        missing=object()
        cases=[('documentId',v) for v in (missing,None,'','  ',False,42)]
        cases += [('version',v) for v in (missing,None,False,True,0,-1,1.5,20.0,'20')]
        for field in ('operationsElementId','operationsRowIdColumnId','snapshotsElementId','snapshotsRowIdColumnId'):
            cases += [(field,v) for v in (missing,None,'','  ',False,42)]
        for field,value in cases:
            changed=copy.deepcopy(target);ops=copy.deepcopy(operations);snaps=copy.deepcopy(snapshots)
            capture_field='elementId' if field.endswith('ElementId') else 'rowIdColumnId' if field.endswith('RowIdColumnId') else field
            affected=[ops] if field.startswith('operations') else [snaps] if field.startswith('snapshots') else [ops,snaps]
            if value is missing:
                changed.pop(field)
                for capture in affected:capture.pop(capture_field)
            else:
                changed[field]=value
                for capture in affected:capture[capture_field]=value
            with self.subTest(field=field,value=value),self.assertRaises(ValueError):
                p.plan_snapshot_reconciliation(d,'operation-A',ops,snaps,changed)

    def test_capture_versions_and_row_counts_are_strict_integer_contracts(self):
        p=self.module();d,op=self.operation()
        target={'documentId':'stage','version':1,'operationsElementId':'operations','operationsRowIdColumnId':'native-ops-id',
                'snapshotsElementId':'snapshots','snapshotsRowIdColumnId':'native-snapshot-id'}
        meta={'origin':'sigma-query','documentId':'stage','version':1,'sourceHash':d['sourceHash'],
              'planHash':d['planHash'],'complete':True,'evidenceId':'actual-read','executionId':'actual-run'}
        operations={**meta,'rowCount':1,'rows':[op],'elementId':'operations','rowIdColumnId':'native-ops-id'}
        snapshots={**meta,'rowCount':0,'rows':[],'elementId':'snapshots','rowIdColumnId':'native-snapshot-id'}
        for role in ('operations','snapshots'):
            for field,values in [('version',[None,True,False,0,-1,1.0,1.5]),
                                 ('rowCount',[None,True,False,-1,0.0,1.0,1.5])]:
                for value in values:
                    ops=copy.deepcopy(operations);snaps=copy.deepcopy(snapshots)
                    (ops if role=='operations' else snaps)[field]=value
                    with self.subTest(role=role,field=field,value=value),self.assertRaises(ValueError):
                        p.plan_snapshot_reconciliation(d,'operation-A',ops,snaps,target)

    def candidate_fixture(self):
        p=self.module();d=self.definition()
        import xml.etree.ElementTree as ET
        primary=[{'id':'p'+str(i),'name':n,'type':t,'class':cl} for i,(n,t,cl) in enumerate([
            ('Line Item','text','LiteralEditableColumn'),('Method','text','SingleSelectEditableColumn'),
            ('Monthly Growth','number','LiteralEditableColumn'),*[(n,'number','LiteralEditableColumn') for n in p.PERIODS],
            ('ID',None,'FormulaColumn')])]
        primary[1]['allowedValues']=['growth','manual','flat'];primary[-1]['formula']='[Row ID]'
        ops={'Workflow':'text','Context':'text','State':'text','Created At':'datetime',
             'Source Hash':'text','Plan Hash':'text','Source Procedure Hash':'text','Replacement Decision ID':'text',
             'Snapshot Definition Hash':'text','Snapshot Capture Contract Hash':'text',
             'Authoring Baseline Document Version':'number','Authoring Baseline Contents Hash':'text',
             'Source Graph Hash':'text','Primary Identity Contract Hash':'text','Snapshot Schema Contract Hash':'text'}
        ops.update({c['operationFrozenColumn']:'number' for c in d['cells']})
        ops.update({c['operationColumn']:c['type'] for c in d['inputFields']})
        schemas={'primary':primary,'operations':[{'id':'o'+str(i),'name':n,'type':t,'class':'LiteralEditableColumn'} for i,(n,t) in enumerate(ops.items())],
                 'snapshots':[{'id':'s'+str(i),'name':n,'type':t,'class':'LiteralEditableColumn'} for i,(n,t) in enumerate(p.SNAPSHOT_FIELDS.items())],
                 'finance':[{'id':'f'+str(i),'name':n,'class':'FormulaColumn','formula':'[Source/'+n+']'} for i,n in enumerate(['Source Row','Section','Line Item',*p.PERIODS])]}
        for role in ['operations','snapshots']:schemas[role].append({'id':'ID','name':'Row ID','class':'InputTableIDColumn','kind':'row_id'})
        bindings={role:{'elementId':role,'pageId':'technical' if role in ('operations','snapshots') else 'forecast-overview' if role=='finance' else 'forecast-edit',
                       'name':role,'columns':cols} for role,cols in schemas.items()}
        bindings['primary']['identities']=[{'sourceRow':r,'nativeRowId':'input-'+str(r),'lineItem':next(c['lineItem'] for c in d['inputFields'] if c['sourceRow']==r)} for r in range(6,18)]
        elements=[{'id':role,'name':role,'kind':'table' if role=='finance' else 'input-table',
            'columns':[{k:v for k,v in c.items() if k in ('id','name','type','formula')} for c in cols]} for role,cols in schemas.items()]
        current={'workbookId':'stage','documentVersion':19,'contents':{'elements':elements,
                 'pages':[{'id':'forecast-edit'},{'id':'forecast-overview'},{'id':'technical','visibility':'hidden'}],
                 'layout':'<Page id="forecast-edit"><Element elementId="primary" /></Page><Page id="forecast-overview"><Element elementId="finance" /></Page><Page id="technical"><Element elementId="operations" /><Element elementId="snapshots" /></Page>'}}
        return p,d,current,bindings

    def test_prepare_and_confirm_are_disabled_and_freeze_live_204_values(self):
        p,d,current,bindings=self.candidate_fixture()
        prepared=p.render_snapshot_candidate(d,current,bindings,phase='prepare')
        compile(prepared['program'],'prepare','exec')
        self.assertEqual(prepared['frozenValueCount'],204)
        self.assertIn('set_disabled(True)',prepared['program'])
        self.assertIn('.returns()',prepared['program'])
        self.assertNotIn('update_rows(',prepared['program'])
        self.assertNotIn('CONFIRMED',prepared['program'])
        self.assertIn("Frozen Snapshot / 8:Jul-26",prepared['program'])
        current['contents']['elements'].append({'id':'op-control','kind':'control','controlType':'text','controlId':'forecast-snapshot-operation','value':''})
        current['contents']['layout']=current['contents']['layout'].replace('<Element elementId="operations" />','<Element elementId="op-control" /><Element elementId="operations" />')
        confirmed=p.render_snapshot_candidate(d,current,bindings,phase='confirm',operation_control_id='op-control',prepared_contract=prepared['preparedContract'])
        compile(confirmed['program'],'confirm','exec')
        self.assertEqual(confirmed['insertEffectCount'],108)
        self.assertIn('if_else(',confirmed['program'])
        self.assertIn('set_disabled(True)',confirmed['program'])
        self.assertNotIn('now()',confirmed['program'])
        self.assertNotIn('update_rows(',confirmed['program'])
        self.assertNotIn('CONFIRMED',confirmed['program'])
        self.assertEqual(confirmed['productionStatus'],'BLOCKED')

    def test_confirm_empty_snapshot_counts_default_to_zero(self):
        import ast
        p,d,current,b=self.candidate_fixture()
        prepared=p.render_snapshot_candidate(d,current,b,phase='prepare')
        current['contents']['elements'].append({'id':'op-control','kind':'control','controlType':'text','controlId':'forecast-snapshot-operation','value':''})
        current['contents']['layout']=current['contents']['layout'].replace('<Element elementId="operations" />','<Element elementId="op-control" /><Element elementId="operations" />')
        tree=ast.parse(p.render_snapshot_candidate(d,current,b,phase='confirm',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'])['program'])
        zero_checks=[n for n in ast.walk(tree) if isinstance(n,ast.Call)
            and isinstance(n.func,ast.Name) and n.func.id=='eq'
            and ast.unparse(n.args[1])=='const_num(0)'
            and any(isinstance(v,ast.Call) and isinstance(v.func,ast.Name)
                    and v.func.id=='count_if' for v in ast.walk(n.args[0]))]
        self.assertEqual(len(zero_checks),3)
        def evaluate(node):
            if isinstance(node,ast.Constant):return node.value
            name=node.func.id
            if name=='count_if':return None
            args=[evaluate(arg) for arg in node.args]
            if name in ('const_num','as_num'):return args[0]
            if name=='coalesce':return next((value for value in args if value is not None),None)
            if name=='eq':return None if None in args else args[0]==args[1]
            raise AssertionError('unexpected zero-check operation '+name)
        for check in zero_checks:
            self.assertTrue(ast.unparse(check.args[0]).startswith('as_num(coalesce(count_if('))
            self.assertIs(evaluate(check),True)

    def test_prepare_pending_count_defaults_only_missing_aggregate_to_zero(self):
        import ast
        p,d,current,b=self.candidate_fixture()
        tree=ast.parse(p.render_snapshot_candidate(d,current,b,phase='prepare')['program'])
        condition=next(n.value.args[0] for n in tree.body if isinstance(n,ast.Expr)
                       and isinstance(n.value,ast.Call) and isinstance(n.value.func,ast.Attribute)
                       and n.value.func.attr=='set_condition')
        def leaves(node):
            if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id=='and_':
                return [leaf for arg in node.args for leaf in leaves(arg)]
            return [node]
        guards=leaves(condition)
        self.assertEqual(len(guards),375)
        pending=next(n for n in guards if isinstance(n,ast.Call) and n.func.id=='eq'
                     and ast.unparse(n.args[1])=='const_num(0)'
                     and any(isinstance(v,ast.Call) and isinstance(v.func,ast.Name)
                             and v.func.id=='count_if' for v in ast.walk(n)))
        def evaluate(node, count):
            if isinstance(node,ast.Constant): return node.value
            name=node.func.id
            if name=='count_if': return count
            args=[evaluate(arg,count) for arg in node.args]
            if name in ('const_num','as_num'): return args[0]
            if name=='coalesce': return next((v for v in args if v is not None),None)
            if name=='eq': return None if None in args else args[0]==args[1]
            raise AssertionError('unexpected pending-count operation '+name)
        # Characterize the emitted formula, using the observed empty-Ops NULL.
        for count,expected in [(None,True),(0,True),(1,False),(2,False),(-1,False)]:
            with self.subTest(count=count): self.assertIs(evaluate(pending,count),expected)
        count_call=next(n for n in ast.walk(pending) if isinstance(n,ast.Call)
                        and isinstance(n.func,ast.Name) and n.func.id=='count_if')
        expected=ast.parse('eq(as_num(coalesce(count_if('+ast.unparse(count_call.args[0])+'), const_num(0))), const_num(0))',mode='eval').body
        self.assertEqual(ast.dump(pending),ast.dump(expected))
        self.assertIn("const_text('snapshot')",ast.unparse(count_call))
        self.assertIn("const_text('PREPARED')",ast.unparse(count_call))
        self.assertEqual(sum(isinstance(n,ast.Call) and isinstance(n.func,ast.Name)
                             and n.func.id=='coalesce' for n in ast.walk(condition)),1)
        insert=next(n for n in ast.walk(tree) if isinstance(n,ast.Call)
                    and isinstance(n.func,ast.Name) and n.func.id=='insert_row')
        self.assertEqual(len(insert.args[1].keys),219)

    def test_sdk_imports_and_numeric_lookup_types_match_documented_modules(self):
        import ast
        p,d,current,b=self.candidate_fixture()
        program=p.render_snapshot_candidate(d,current,b,phase='prepare')['program']
        tree=ast.parse(program)
        imports={alias.name:node.module for node in tree.body if isinstance(node,ast.ImportFrom) for alias in node.names}
        self.assertEqual(imports['TextControl'],'sigma.controlHandle')
        self.assertEqual(imports['add_text_control'],'sigma.control')
        # lookup returns generic Formula; numeric arithmetic requires NumberFormula.
        self.assertIn('v0 = as_num(lookup(',program)
        self.assertIn('a1 = as_num(lookup(',program)
        self.assertIn('a0 = as_text(lookup(',program)

    def test_new_sequence_is_disabled_before_condition_or_toast_work(self):
        import ast
        p,d,current,b=self.candidate_fixture()
        program=p.render_snapshot_candidate(d,current,b,phase='prepare')['program']
        nodes=ast.parse(program).body
        created=next(i for i,n in enumerate(nodes) if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='sequence' for t in n.targets))
        first=nodes[created+1]
        self.assertIsInstance(first,ast.Expr)
        self.assertEqual(first.value.func.attr,'set_disabled')
        self.assertEqual(first.value.args[0].value,True)
        last_disable=max(i for i,n in enumerate(nodes) if isinstance(n,ast.Expr) and isinstance(n.value,ast.Call) and isinstance(n.value.func,ast.Attribute) and n.value.func.attr=='set_disabled')
        self.assertGreater(last_disable,created+1)
        self.assertIn('if not sequence.is_disabled():',program)

    def test_public_id_placeholder_is_replaced_only_by_actual_native_id_class(self):
        p,d,current,b=self.candidate_fixture()
        for role in ('operations','snapshots'):
            b[role]['columns'][-1]['id']='native-'+role
            e=next(x for x in current['contents']['elements'] if x['id']==role)
            e['columns'][-1]={'id':'ID'}
        candidate=p.render_snapshot_candidate(d,current,b,phase='prepare')
        self.assertIn("get_column('native-operations')",candidate['program'])
        bad=copy.deepcopy(b);bad['operations']['columns'][-1]['class']='FormulaColumn'
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,bad,phase='prepare')

    def test_renderer_blocks_wrong_ids_alias_selector_and_incomplete_schema(self):
        p,d,current,b=self.candidate_fixture()
        for role,field,value in [('operations','elementId','other'),('snapshots','pageId','forecast-edit')]:
            bad=copy.deepcopy(b);bad[role][field]=value
            with self.subTest(role=role,field=field),self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,bad,phase='prepare')
        for change in ('class','missing','identity'):
            bad=copy.deepcopy(b)
            if change=='class':bad['operations']['columns'][-1]['class']='FormulaColumn'
            elif change=='missing':bad['operations']['columns'].pop(0)
            else:bad['primary']['identities'][1]['nativeRowId']=bad['primary']['identities'][0]['nativeRowId']
            with self.subTest(change=change),self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,bad,phase='prepare')
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='confirm',operation_control_id='missing')

    def test_changed_definition_or_wrong_approval_is_rejected(self):
        p=self.module();d,op=self.operation()
        changed=copy.deepcopy(d);changed['cells'][0]['sourceCell']='H8'
        with self.assertRaises(ValueError):p.reconcile_snapshot_rows(changed,op,[])
        with self.assertRaises(ValueError):p.build_snapshot_definition(Path(__file__).parent/'fixtures/fixture-forecast-model.xlsx',
            plan_hash='a'*64,procedure_hash='b'*64,decision={'decisionId':'old','status':'PENDING','scope':{}})

    def completion_fixture(self):
        p=self.module();d,op=self.operation()
        rows=[{**row,'nativeRowId':'snapshot-'+str(i)} for i,row in enumerate(p.reconcile_snapshot_rows(d,op,[])['missingRows'])]
        base={'origin':'sigma-query','documentId':'stage','version':28,'sourceHash':d['sourceHash'],
              'planHash':d['planHash'],'complete':True,'evidenceId':'read','executionId':'run'}
        operations={**base,'elementId':'operations','rowIdColumnId':'native-ops-id',
                    'rowCount':1,'rows':[op]}
        snapshots={**base,'elementId':'snapshots','rowIdColumnId':'native-snapshot-id',
                   'rowCount':108,'rows':rows}
        target={'documentId':'stage','version':28,'operationsElementId':'operations',
                'operationsRowIdColumnId':'native-ops-id','snapshotsElementId':'snapshots',
                'snapshotsRowIdColumnId':'native-snapshot-id'}
        return p,d,op,operations,snapshots,target

    def test_completion_marks_only_exact_prepared_operation_then_resumes_clear(self):
        p,d,op,operations,snapshots,target=self.completion_fixture()
        planned=p.plan_snapshot_completion(d,'operation-A',operations,snapshots,target,control_value='operation-A')
        self.assertEqual(planned['nextAction'],'MARK_TERMINAL')
        self.assertEqual(planned['observedOperationState'],'PREPARED')
        self.assertEqual(planned['operationReadbackHash'],p.digest(operations))
        self.assertEqual(planned['snapshotReadbackHash'],p.digest(snapshots))
        self.assertEqual(planned['reconciliationHash'],planned['reconciliation']['reconciliationHash'])
        self.assertEqual(planned['completionPlanHash'],p.digest({k:v for k,v in planned.items() if k!='completionPlanHash'}))
        self.assertEqual(planned['operationId'],'operation-A')
        self.assertEqual(planned['terminalState'],'SNAPSHOT_COMPONENT_COMPLETE')
        self.assertEqual(planned['reconciliation']['existingCount'],108)
        self.assertEqual(planned['productionStatus'],'BLOCKED')
        terminal=copy.deepcopy(operations);terminal['rows'][0]['State']='SNAPSHOT_COMPONENT_COMPLETE'
        resumed=p.plan_snapshot_completion(d,'operation-A',terminal,snapshots,target,control_value='operation-A')
        self.assertEqual(resumed['nextAction'],'CLEAR_CONTROL')
        self.assertEqual(resumed['observedOperationState'],p.SNAPSHOT_TERMINAL_STATE)
        self.assertNotEqual(resumed['operationReadbackHash'],planned['operationReadbackHash'])
        self.assertEqual(p.plan_snapshot_completion(d,'operation-A',terminal,snapshots,target,control_value='')['nextAction'],'NONE')
        self.assertEqual(operations['rows'][0]['State'],'PREPARED')
        self.assertEqual(len(snapshots['rows']),108)

    def test_deliberate_second_intent_has_distinct_id_and_retains_old_snapshots(self):
        p,d,op,operations,snapshots,target=self.completion_fixture()
        terminal=copy.deepcopy(operations);terminal['rows'][0]['State']=p.SNAPSHOT_TERMINAL_STATE
        self.assertEqual(p.plan_snapshot_completion(d,'operation-A',terminal,snapshots,target,
                                                    control_value='')['nextAction'],'NONE')
        new_op=copy.deepcopy(op)
        new_op['nativeRowId']='operation-B'
        new_op['Created At']='2026-10-16T00:00:00+00:00'
        retry=p.reconcile_snapshot_rows(d,new_op,snapshots['rows'])
        self.assertEqual(retry['operationId'],'operation-B')
        self.assertEqual(retry['existingCount'],0)
        self.assertEqual(len(retry['missingRows']),108)
        self.assertEqual({row['Operation ID'] for row in retry['missingRows']},{'operation-B'})
        self.assertEqual({row['Operation ID'] for row in snapshots['rows']},{'operation-A'})

    def test_completion_blocks_incomplete_conflicting_or_foreign_snapshot(self):
        p,d,op,operations,snapshots,target=self.completion_fixture()
        incomplete=copy.deepcopy(snapshots);incomplete['rows'].pop();incomplete['rowCount']=107
        conflicting=copy.deepcopy(snapshots);conflicting['rows'][0]['Value']=999
        duplicate=copy.deepcopy(snapshots);duplicate['rows'].append({**duplicate['rows'][0],'nativeRowId':'extra'});duplicate['rowCount']=109
        foreign=copy.deepcopy(snapshots);foreign['rows'][0]['Snapshot Key']='foreign'
        for name,rows in [('incomplete',incomplete),('conflicting',conflicting),('duplicate',duplicate),('foreign',foreign)]:
            with self.subTest(name=name),self.assertRaises(ValueError):
                p.plan_snapshot_completion(d,'operation-A',operations,rows,target,control_value='operation-A')

    def test_public_insert_reconciliation_remains_prepared_only(self):
        import inspect
        p,d,op,operations,snapshots,target=self.completion_fixture()
        self.assertNotIn('allowed_states',inspect.signature(p.reconcile_snapshot_rows).parameters)
        self.assertNotIn('allowed_states',inspect.signature(p.plan_snapshot_reconciliation).parameters)
        empty=copy.deepcopy(snapshots);empty['rows']=[];empty['rowCount']=0
        for state in ('FAILED',p.SNAPSHOT_TERMINAL_STATE):
            changed=copy.deepcopy(operations);changed['rows'][0]['State']=state
            with self.subTest(state=state),self.assertRaises(ValueError):
                p.reconcile_snapshot_rows(d,changed['rows'][0],[])
            with self.subTest(state=state),self.assertRaises(ValueError):
                p.plan_snapshot_reconciliation(d,'operation-A',changed,empty,target)
            with self.subTest(state=state),self.assertRaises(ValueError):
                p.plan_snapshot_completion(d,'operation-A',changed,empty,target,control_value='operation-A')

    def test_completion_blocks_other_pending_intent_wrong_control_and_stale_capture(self):
        p,d,op,operations,snapshots,target=self.completion_fixture()
        other=copy.deepcopy(operations);other['rows'].append({**op,'nativeRowId':'operation-B'});other['rowCount']=2
        with self.assertRaises(ValueError):p.plan_snapshot_completion(d,'operation-A',other,snapshots,target,control_value='operation-A')
        for control in ('', 'operation-B', None):
            with self.subTest(control=control),self.assertRaises(ValueError):
                p.plan_snapshot_completion(d,'operation-A',operations,snapshots,target,control_value=control)
        stale=copy.deepcopy(snapshots);stale['version']=27
        with self.assertRaises(ValueError):p.plan_snapshot_completion(d,'operation-A',operations,stale,target,control_value='operation-A')

    def test_completion_renderer_separates_exact_id_mark_from_verified_clear(self):
        import ast
        p,d,current,b=self.candidate_fixture()
        prepared=p.render_snapshot_candidate(d,current,b,phase='prepare')
        current['contents']['elements'].append({'id':'op-control','kind':'control','controlType':'text',
                                                  'controlId':'forecast-snapshot-operation','value':''})
        current['contents']['layout']=current['contents']['layout'].replace(
            '<Element elementId="operations" />','<Element elementId="op-control" /><Element elementId="operations" />')
        _,_,_,operations,snapshots,target=self.completion_fixture()
        for capture in (operations,snapshots):capture['version']=19
        target['version']=19
        for role,capture in [('operations',operations),('snapshots',snapshots)]:
            native_id=next(col['id'] for col in b[role]['columns'] if col['class']=='InputTableIDColumn')
            capture['rowIdColumnId']=native_id
            target[role+'RowIdColumnId']=native_id
        mark_plan=p.plan_snapshot_completion(d,'operation-A',operations,snapshots,target,control_value='operation-A')
        mark=p.render_snapshot_candidate(d,current,b,phase='completion-mark',operation_control_id='op-control',
                                         prepared_contract=prepared['preparedContract'],completion_plan=mark_plan)
        tree=ast.parse(mark['program'])
        updates=[n for n in ast.walk(tree) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='update_rows']
        self.assertEqual(len(updates),1)
        self.assertIsInstance(updates[0].args[1],ast.Dict)
        self.assertEqual(len(updates[0].args[1].keys),1)
        self.assertEqual(ast.unparse(updates[0].args[1].values[0]),'op')
        self.assertNotIn('set_control_value(op, "")',mark['program'])
        self.assertIn("const_text('operation-A')",mark['program'])
        self.assertIn("const_text('SNAPSHOT_COMPONENT_COMPLETE')",mark['program'])
        self.assertIn('for i in range(108):',mark['program'])
        self.assertIn('eq(count_if(matching), const_num(1))',mark['program'])
        self.assertIn('not_(allowed_parts[0])',mark['program'])
        self.assertEqual(mark['insertEffectCount'],0)
        self.assertEqual(mark['productionStatus'],'BLOCKED')
        terminal=copy.deepcopy(operations);terminal['rows'][0]['State']='SNAPSHOT_COMPONENT_COMPLETE'
        clear_plan=p.plan_snapshot_completion(d,'operation-A',terminal,snapshots,target,control_value='operation-A')
        cleared=p.render_snapshot_candidate(d,current,b,phase='completion-clear',operation_control_id='op-control',
                                            prepared_contract=prepared['preparedContract'],completion_plan=clear_plan)
        clear_tree=ast.parse(cleared['program'])
        self.assertFalse(any(isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='update_rows'
                             for n in ast.walk(clear_tree)))
        self.assertIn('set_control_value(op, "")',cleared['program'])
        self.assertIn("const_text('operation-A')",cleared['program'])
        self.assertIn("const_text('SNAPSHOT_COMPONENT_COMPLETE')",cleared['program'])
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-clear',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=mark_plan)
        forged_action=copy.deepcopy(mark_plan);forged_action['nextAction']='CLEAR_CONTROL'
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-clear',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=forged_action)
        forged_action['completionPlanHash']=p.digest({k:v for k,v in forged_action.items() if k!='completionPlanHash'})
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-clear',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=forged_action)
        forged_readback=copy.deepcopy(clear_plan);forged_readback['operationReadbackHash']='f'*64
        forged_readback['completionPlanHash']=p.digest({k:v for k,v in forged_readback.items() if k!='completionPlanHash'})
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-clear',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=forged_readback)
        tampered=copy.deepcopy(mark_plan);tampered['reconciliation']['existingCount']=107
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-mark',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=tampered)
        malformed=copy.deepcopy(mark_plan);malformed['reconciliation']['target']=None
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-mark',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=malformed)
        wrong_native_target=copy.deepcopy(mark_plan)
        wrong_native_target['reconciliation']['target']['operationsRowIdColumnId']='unrelated-native-id'
        wrong_native_target['reconciliation']['reconciliationHash']=p.digest({
            k:v for k,v in wrong_native_target['reconciliation'].items() if k!='reconciliationHash'})
        with self.assertRaises(ValueError):p.render_snapshot_candidate(d,current,b,phase='completion-mark',
            operation_control_id='op-control',prepared_contract=prepared['preparedContract'],completion_plan=wrong_native_target)

if __name__ == '__main__':unittest.main()
