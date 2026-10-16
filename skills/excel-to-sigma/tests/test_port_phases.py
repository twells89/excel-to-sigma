"""Synthetic offline phase evidence; never actual Sigma output/runtime proof."""
import copy,sys,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts/lib'))
import test_custom_ports as fixtures
import custom_ports as p

class PortPhaseTests(unittest.TestCase):
    def test_same_type_output_semantic_permutations_fail_even_with_fresh_review_hash(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        for left,right in [('Target','LowerBound'),('OperationId','ErrorCode')]:
            ev=outputs(m,r,c,first,('allocation','goalSeek'));mapping=ev['outputs']['goalSeek']['columnMap']
            mapping[left],mapping[right]=mapping[right],mapping[left];reviewed(ev)
            with self.subTest(fields=(left,right)),self.assertRaisesRegex(ValueError,'SEMANTIC_OUTPUT'):
                p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=context(m,r,c),output_evidence=ev)

    def test_post_phase_and_next_phase_refuse_lost_rest_omitted_sdk_action(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        from action_inventory import digest
        from test_action_inventory import ActionInventoryTests
        _,m,r,c,s=fixtures.PortGateTests().case()
        c['contents']['elements'].append({'id':'old-navigation','kind':'button','text':'Existing REST-omitted action'})
        c['contents']['layout']=c['contents']['layout'].replace('</Page>','<Element elementId="old-navigation" gridRow="100 / 102"/></Page>',1)
        before=context(m,r,c);_,old=ActionInventoryTests().fixture();old=ActionInventoryTests().reviewed(old)
        e=copy.deepcopy(old['elements']['button']);e['elementId']='old-navigation';artifact=e.pop('restorationArtifact')
        artifact.update(elementId='old-navigation',sourceObservationHash=digest(e));e['restorationArtifact']=artifact;before['sdkActions']['elements']['old-navigation']=e;reviewed(before)
        first=p.prepare_port_update(m,r,c,s,phase_context=before);c=installed(c,first);m,r=restage(m,r,c);after=context(m,r,c)
        with self.assertRaisesRegex(ValueError,'POST_PHASE'):p.audit_port_update(first,c)
        with self.assertRaises(ValueError):p.audit_port_update(first,c,post_phase_context=after)
        with self.assertRaises(ValueError):p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=after,output_evidence=outputs(m,r,c,first,('allocation','goalSeek')))

    def test_post_phase_requires_exact_action_identity_or_reviewed_actual_restoration(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        from action_inventory import digest
        from test_action_inventory import ActionInventoryTests
        _,m,r,c,s=fixtures.PortGateTests().case()
        c['contents']['elements'].append({'id':'old-action','kind':'button','text':'Old omitted action'})
        c['contents']['layout']=c['contents']['layout'].replace('</Page>','<Element elementId="old-action" gridRow="100 / 102"/></Page>',1)
        before=context(m,r,c);_,template=ActionInventoryTests().fixture();template=ActionInventoryTests().reviewed(template)
        e=copy.deepcopy(template['elements']['button']);e['elementId']='old-action';artifact=e.pop('restorationArtifact')
        artifact.update(elementId='old-action',sourceObservationHash=digest(e));e['restorationArtifact']=artifact;before['sdkActions']['elements']['old-action']=e;reviewed(before)
        first=p.prepare_port_update(m,r,c,s,phase_context=before);c=installed(c,first);m,r=restage(m,r,c);after=context(m,r,c)
        after['sdkActions']['elements']['old-action']=copy.deepcopy(e);reviewed(after)
        proof=p.audit_port_update(first,c,post_phase_context=after);self.assertEqual(proof['status'],'VERIFIED_COMPONENT_CONSTRUCTION');self.assertIn('actionPreservation',proof)
        altered=copy.deepcopy(after);element=altered['sdkActions']['elements']['old-action'];element['sequences'][0]['id']='restored-assigned-ID'
        artifact=element.pop('restorationArtifact');artifact['sourceObservationHash']=digest(element);element['restorationArtifact']=artifact;reviewed(altered)
        with self.assertRaisesRegex(ValueError,'RESTORED_SEQUENCE_ID'):p.audit_port_update(first,c,post_phase_context=altered)
        alteration={'status':'APPROVED_STAGING','owner':'unit-reviewer','evidenceId':'SYNTHETIC_RESTORATION_ONLY','documentId':'doc','version':c['documentVersion'],
            'beforeInventoryHash':digest(before['sdkActions']),'afterInventoryHash':digest(altered['sdkActions']),
            'assignedSequenceIds':{'old-action':{e['sequences'][0]['id']:'restored-assigned-ID'}},
            'restorationProofHashes':{'old-action':digest(artifact['restoreProof'])}}
        altered['actionRestorationReview']=alteration;reviewed(altered)
        self.assertEqual(p.audit_port_update(first,c,post_phase_context=altered)['status'],'VERIFIED_COMPONENT_CONSTRUCTION')
        for change in ({'version':8},{'assignedSequenceIds':{}},{'restorationProofHashes':{}},{'owner':'PENDING'}):
            bad=copy.deepcopy(altered);bad['actionRestorationReview'].update(change);reviewed(bad)
            with self.subTest(change=change),self.assertRaises(ValueError):p.audit_port_update(first,c,post_phase_context=bad)
        for field in ('inputValuesHash','schemaDisplayHash'):
            bad=copy.deepcopy(after);bad[field]='f'*64;reviewed(bad)
            with self.subTest(field=field),self.assertRaisesRegex(ValueError,'INPUT_OR_DISPLAY_CHANGED'):p.audit_port_update(first,c,post_phase_context=bad)
        bad=copy.deepcopy(after);bad['receiptHash']=None;reviewed(bad)
        with self.assertRaises(ValueError):p.audit_port_update(first,c,post_phase_context=bad)
        self.assertEqual(p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=after,output_evidence=outputs(m,r,c,first,('allocation','goalSeek')))['phase'],'solver-outputs')
        self.assertEqual(p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=altered,output_evidence=outputs(m,r,c,first,('allocation','goalSeek')))['phase'],'solver-outputs')

    def test_new_phase_hosts_and_run_effects_must_match_declared_port_constructor(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        from action_inventory import digest,configuration
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        after=context(m,r,c);button=after['sdkActions']['elements']['task5AllocationButton']
        button['sequences'][0]['blocks'][0]['fields']['codeElementId']='foreign-code'
        artifact=button.pop('restorationArtifact');button['semanticHash']=digest(configuration(button));artifact['sourceObservationHash']=digest(button);artifact['semanticHash']=button['semanticHash'];artifact['restoreProof']['semanticHash']=button['semanticHash'];button['restorationArtifact']=artifact;reviewed(after)
        with self.assertRaisesRegex(ValueError,'NEW_PORT_ACTION'):p.audit_port_update(first,c,post_phase_context=after)
        c=copy.deepcopy(c);c['contents']['elements'].append({'id':'foreign-new-host','kind':'text','body':'unrelated'});m,r=restage(m,r,c)
        with self.assertRaisesRegex(ValueError,'UNDECLARED_POST_PHASE'):p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=context(m,r,c),output_evidence=outputs(m,r,c,first,('allocation','goalSeek')))
    def test_emitted_runtime_marshals_decimal_target_success_and_failure_to_real_snowpark_types(self):
        from decimal import Decimal
        from types import SimpleNamespace
        import json,hashlib,math
        model=json.loads((ROOT/'tests/fixtures/task5-controlled-model.json').read_text())
        request={'RequestKey':'goal-seek','OperationId':'unit-A','State':'UNPREPARED','Target':Decimal('2000000.0000000000000000000'),
            'LowerBound':Decimal('0'),'UpperBound':Decimal('.25'),'ToleranceCents':Decimal('1'),'MaxIterations':64,'ModelBinding':'unit-binding'}
        class Frame:
            def __init__(self,rows):self.rows=rows;self.columns=list(rows[0])
            def collect(self):return [SimpleNamespace(as_dict=lambda row=row:row) for row in self.rows]
        data={'Actuals':model['actuals'],'Forecast Inputs View':model['inputs'],'Forecast Line Definitions':model['definitions'],'Goal Seek Invocation':[request]}
        output=[]
        def create(rows,schema):
            # Snowpark DOUBLE to_sql rejects Decimal: enforce primitive schema
            # types here rather than accepting any Python row as older doubles did.
            fields=[part.rsplit(' ',1)[1] for part in schema.split(', ')]
            for row in rows:
                for value,kind in zip(row,fields):
                    if value is None:continue
                    expected={'DOUBLE':float,'INTEGER':int,'BOOLEAN':bool,'STRING':str}[kind]
                    self.assertIs(type(value),expected,(kind,value))
                    if kind=='DOUBLE':self.assertTrue(math.isfinite(value))
            return rows
        sigma=SimpleNamespace(get_element=lambda name:Frame(data[name]),output=lambda name,rows:output.append((name,rows)))
        session=SimpleNamespace(create_dataframe=create)
        program=p.emit_python_programs('unit-binding')['goalSeek']
        for state,expected in [('UNPREPARED','FAILURE'),('PENDING','SUCCESS')]:
            request['State']=state;exec(program['code'],{'sigma':sigma,'session':session})
            result=dict(zip(program['columns'],output[-1][1][0]));self.assertEqual(result['Status'],expected)
            self.assertEqual(result['Target'],2000000.0)
        request['Target']=True;exec(program['code'],{'sigma':sigma,'session':session})
        result=dict(zip(program['columns'],output[-1][1][0]));self.assertEqual(result['Status'],'FAILURE');self.assertIsNone(result['Target'])
        request['Target']=Decimal('NaN');exec(program['code'],{'sigma':sigma,'session':session})
        result=dict(zip(program['columns'],output[-1][1][0]));self.assertEqual(result['Status'],'FAILURE');self.assertIsNone(result['Target'])
        data.update({'Task 5 Headcount Inputs':[{'Department':d,'Headcount':Decimal(n)} for d,n in zip(p.DEPARTMENTS,[40,30,20,10])],
            'Task 5 Allocation Assumptions':[{'Key':'Overhead','Value':Decimal('100000')}]})
        allocation=p.emit_python_programs('unit-binding')['allocation'];exec(allocation['code'],{'sigma':sigma,'session':session})
        self.assertEqual(len(output[-1][1]),4)
        self.assertEqual([row[allocation['columns'].index('SourceIndex')] for row in output[-1][1]],[1,2,3,4])
        data['Task 5 Allocation Assumptions'][0]['Value']=Decimal('Infinity');exec(allocation['code'],{'sigma':sigma,'session':session})
        self.assertEqual(len(output[-1][1]),1)
        self.assertEqual(hashlib.sha256((ROOT/'scripts/lib/custom_port_math.py').read_bytes()).hexdigest(),'067a4d5e5cdb83f703a5cd2a3d543a04893d295b85b55f8d1bc61c00c37cec4c')
    def test_default_phase_never_references_unmaterialized_code_outputs(self):
        _,m,r,c,s=fixtures.PortGateTests().case()
        proposal=p.prepare_port_update(m,r,c,s)
        added=[e for e in proposal['dryRunRequest']['contents']['elements'] if e['id'] in proposal['addedElementIds']]
        self.assertEqual(len(added),4)
        self.assertEqual(sum(e['kind']=='code' for e in added),2)
        self.assertFalse(any(e.get('source',{}).get('kind')=='code-output' for e in added),'actual backend400: output not found until Python has run')
        self.assertFalse(any('ApplyGuard' in e['id'] for e in added))
        self.assertEqual(proposal['phase'],'code-first');self.assertEqual(proposal['runtimeStatus'],'UNRESOLVED')
    def test_later_phase_needs_actual_output_metadata_not_expected_schema(self):
        _,m,r,c,s=fixtures.PortGateTests().case()
        for phase in ('solver-outputs','guard-output','complete'):
            with self.subTest(phase=phase),self.assertRaises(ValueError):p.prepare_port_update(m,r,c,s,phase=phase)
    def test_low_level_fragment_default_is_safe_and_atomic_mode_is_refused(self):
        fragment=p.build_port_fragments('binding','conn')
        self.assertEqual([e['kind'] for e in fragment['elements']],['code','button','code','button'])
        with self.assertRaises(ValueError):p.build_port_fragments('binding','conn',phase='complete')
    def test_three_phases_preserve_actual_pairs_and_immutable_owned_graph_without_business_success(self):
        from port_phase_fixtures import complete
        _,m,r,c,s=fixtures.PortGateTests().case();original=copy.deepcopy(c)
        m,r,c,last=complete(m,r,c,s)
        self.assertEqual(last['phase'],'guard-output');self.assertEqual(len(last['ownedPortElementIds']),9)
        self.assertEqual(last['addedElementIds'],['task5GoalSeekApplyGuard'])
        self.assertEqual(c['contents']['elements'][:len(original['contents']['elements'])],original['contents']['elements'])
        self.assertEqual(last['runtimeColumnMaps']['Goal Seek Result']['Target'],{'identifier':'"Target"','rowKey':'Target'})
        for e in c['contents']['elements']:
            if e.get('source',{}).get('kind')=='code-output':
                self.assertTrue(all(col['formula'].startswith('[Code: ') for col in e['columns']))
                self.assertFalse(any('Code Output/' in col['formula'] for col in e['columns']))
        guard=next(e for e in c['contents']['elements'] if e['id']=='task5GoalSeekApplyGuardPython')
        self.assertIn("'identifier': '\"Target\"'",guard['code'])
        self.assertEqual(last['runtimeStatus'],'UNRESOLVED');self.assertFalse(last['workflowComplete']);self.assertEqual(last['productionStatus'],'BLOCKED')
    def test_materialization_requires_actual_reviewed_current_schema_execution_and_recovery(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        capture=outputs(m,r,c,first,('allocation','goalSeek'));current_context=context(m,r,c)
        def prepare(ev=capture,**changes):return p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=current_context,output_evidence=ev,**changes)
        candidate=prepare();self.assertEqual(set(candidate['addedElementIds']),{'task5AllocationResult','task5GoalSeekResult','task5GoalSeekApplyGuardPython','task5GoalSeekApplyGuardButton'})
        for mode in ('expected-only','stale','failed-run','wrong-code','wrong-type','missing-pair','duplicate-field','missing-install','pending'):
            ev=copy.deepcopy(capture);record=ev['outputs']['goalSeek']
            if mode=='expected-only':record['metadata']['origin']='EXPECTED_SCHEMA'
            elif mode=='stale':ev['version']-=1
            elif mode=='failed-run':record['execution']['status']='FAILED'
            elif mode=='wrong-code':record['execution']['codeHash']='a'*64
            elif mode=='wrong-type':record['metadata']['observedFields'][0]['type']='boolean'
            elif mode=='missing-pair':record['columnMap'].pop('Target')
            elif mode=='duplicate-field':record['metadata']['observedFields'][1]=record['metadata']['observedFields'][0]
            elif mode=='missing-install':ev['installation']={}
            elif mode=='pending':ev['review']['owner']='PENDING'
            record['metadataHash']=p._digest(record['metadata']);reviewed(ev)
            if mode=='pending':ev['review']['owner']='PENDING'
            with self.subTest(mode=mode),self.assertRaises(ValueError):prepare(ev)
        with self.assertRaises(ValueError):prepare(consumed_phases=[candidate['phaseId']])
        bad=copy.deepcopy(current_context);bad['sdkActions']['elements'].pop(next(iter(bad['sdkActions']['elements'])));reviewed(bad)
        with self.assertRaises(ValueError):p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=bad,output_evidence=capture)
        # A recorded prior consumed phase is expected; only the new phase may not replay.
        self.assertEqual(prepare(consumed_phases=[first['phaseId']])['phaseId'],candidate['phaseId'])
        dry={'httpStatus':400,'dryRun':True,'proposalHash':candidate['proposalHash'],'evidenceId':'actual-scratch-output-not-found'}
        with self.assertRaises(ValueError):p.finalize_port_update(candidate,c,dry,fresh_phase_context=current_context)
    def test_default_partial_cannot_finalize_without_actual_current_context(self):
        _,m,r,c,s=fixtures.PortGateTests().case();proposal=p.prepare_port_update(m,r,c,s)
        with self.assertRaises(ValueError):p.finalize_port_update(proposal,c,{'httpStatus':200,'dryRun':True,'proposalHash':proposal['proposalHash'],'evidenceId':'unit-only'})
    def test_phase_order_code_source_freshness_context_and_unknown_collisions_fail_closed(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        def candidate(cc=c,mm=m,rr=r,ss=s,phase='solver-outputs',**kw):
            return p.prepare_port_update(mm,rr,cc,ss,phase=phase,previous_phase=first,phase_context=kw.pop('phase_context',context(mm,rr,cc)),
                output_evidence=outputs(mm,rr,cc,first,('allocation','goalSeek')),**kw)
        for phase in ('guard-output','complete'):
            with self.subTest(phase=phase),self.assertRaises(ValueError):candidate(phase=phase)
        for change in ('code','source','collision','duplicate-title','stale-version'):
            cc=copy.deepcopy(c);ss=copy.deepcopy(s)
            if change=='code':next(e for e in cc['contents']['elements'] if e['id']=='task5GoalSeekPython')['code']+='\n# changed\n'
            elif change=='source':
                cc['contents']['elements'][0]['source']['statement']='different objective';ss['Actuals']['elementHash']=p._digest(cc['contents']['elements'][0])
            elif change=='collision':cc['contents']['elements'].append({'id':'task5GoalSeekResult','kind':'text','body':'foreign'})
            elif change=='duplicate-title':cc['contents']['elements'].append({'id':'foreign','kind':'table','name':'Goal Seek Result'})
            else:cc['documentVersion']+=1
            with self.subTest(change=change),self.assertRaises(ValueError):candidate(cc=cc,ss=ss)
        missing=context(m,r,c);missing['codeObservations'].pop(next(iter(missing['codeObservations'])));reviewed(missing)
        with self.assertRaises(ValueError):candidate(phase_context=missing)
        incomplete=context(m,r,c);button=next(e for e in incomplete['sdkActions']['elements'].values() if e['sequences']);button.pop('restorationArtifact');reviewed(incomplete)
        with self.assertRaises(ValueError):candidate(phase_context=incomplete)
        forged=copy.deepcopy(first);cc=copy.deepcopy(c)
        for target in (forged['dryRunRequest']['contents'],cc['contents']):next(e for e in target['elements'] if e['id']=='task5GoalSeekPython')['code']+='\n# foreign implementation\n'
        forged['proposalHash']=p._digest({k:v for k,v in forged.items() if k!='proposalHash'})
        evidence=outputs(m,r,cc,forged,('allocation','goalSeek'))
        with self.assertRaisesRegex(ValueError,'REVIEWED_PORT_SOURCE'):p.prepare_port_update(m,r,cc,s,phase='solver-outputs',previous_phase=forged,phase_context=context(m,r,cc),output_evidence=evidence)
        proposal=candidate();dry={'httpStatus':200,'dryRun':True,'proposalHash':proposal['proposalHash'],'evidenceId':'SYNTHETIC_ONLY'}
        current=context(m,r,c)
        self.assertEqual(p.finalize_port_update(proposal,c,dry,fresh_phase_context=current)['body']['documentVersion'],9)
        with self.assertRaises(ValueError):p.finalize_port_update(proposal,c,dry,fresh_phase_context=current,consumed_proposals=[proposal['phaseId']])
        current['inputValuesHash']='e'*64;reviewed(current)
        with self.assertRaises(ValueError):p.finalize_port_update(proposal,c,dry,fresh_phase_context=current)
        with self.assertRaises(ValueError):p.audit_port_update(proposal,{**c,'documentVersion':8,'contents':proposal['dryRunRequest']['contents']})
    def test_native_children_keep_actual_assigned_ids_formulas_types_and_no_send_final_adoption(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed,native_children
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        m,r,c,ev=native_children(m,r,c,outputs(m,r,c,first,('allocation','goalSeek')))
        def prepare(evidence=ev):return p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=first,phase_context=context(m,r,c),output_evidence=evidence)
        second=prepare();self.assertEqual(second['addedElementIds'],['task5GoalSeekApplyGuardPython','task5GoalSeekApplyGuardButton'])
        self.assertEqual(second['adoptedElementIds'],['actual-native-allocation','actual-native-goalSeek'])
        self.assertEqual(second['dryRunRequest']['contents']['elements'][:-2],c['contents']['elements'])
        for mode in ('missing','foreign','hash','type','formula','column-id'):
            bad=copy.deepcopy(ev);record=bad['outputs']['goalSeek'];native=record['nativeChildObservation']
            if mode=='missing':record.pop('nativeChildObservation')
            elif mode=='foreign':native['elementId']='foreign'
            elif mode=='hash':native['elementHash']='a'*64
            elif mode=='type':native['typedColumns']['entries'][0]['type']={'type':'boolean'}
            elif mode=='formula':native['typedColumns']['entries'][0]['formula']='[Wrong/OperationId]'
            else:native['columnIds']['OperationId']='foreign-column'
            native['typedColumnsHash']=p._digest(native['typedColumns']);reviewed(bad)
            with self.subTest(mode=mode),self.assertRaises(ValueError):prepare(bad)
        c=installed(c,second);m,r=restage(m,r,c)
        m,r,c,ev=native_children(m,r,c,outputs(m,r,c,second,('applyGuard',)))
        final=p.prepare_port_update(m,r,c,s,phase='guard-output',previous_phase=second,phase_context=context(m,r,c),output_evidence=ev)
        self.assertEqual(final['addedElementIds'],[]);self.assertEqual(final['adoptedElementIds'],['actual-native-applyGuard']);self.assertFalse(final['restMutationRequired'])
        self.assertEqual(final['dryRunRequest']['contents'],c['contents']);self.assertEqual(len(final['ownedPortElementIds']),9)
        self.assertEqual(p.audit_port_update(final,c,post_phase_context=context(m,r,c))['status'],'VERIFIED_COMPONENT_NATIVE_ADOPTION')
        with self.assertRaises(ValueError):p.finalize_port_update(final,c,{'httpStatus':200,'dryRun':True,'evidenceId':'unit'},fresh_phase_context=context(m,r,c))
    def test_reviewed_current_code_adoption_is_distinct_from_an_unexecuted_new_put(self):
        from port_phase_fixtures import context,installed,restage,outputs,reviewed
        _,m,r,c,s=fixtures.PortGateTests().case();first=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c));c=installed(c,first);m,r=restage(m,r,c)
        fragments=p.build_port_fragments(first['modelBinding'],'conn')
        ev=reviewed({'origin':'sigma-code-first-native-adoption-actual','evidenceId':'SYNTHETIC_UNIT_ONLY','documentId':c['workbookId'],'version':c['documentVersion'],
            'sourceHash':m['sourceHash'],'planHash':m['planHash'],'modelBinding':first['modelBinding'],'contentsHash':p._digest(c['contents']),
            'receiptHash':p._digest(r),'elementHashes':{e['id']:p._digest(e) for e in fragments['elements']}})
        adopted=p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c),adoption_evidence=ev)
        self.assertFalse(adopted['restMutationRequired']);self.assertEqual(adopted['addedElementIds'],[]);self.assertEqual(len(adopted['ownedPortElementIds']),4)
        proof=p.audit_port_update(adopted,c,post_phase_context=context(m,r,c));self.assertEqual(proof['status'],'VERIFIED_COMPONENT_NATIVE_ADOPTION')
        capture=outputs(m,r,c,adopted,('allocation','goalSeek'));capture['installation']=dict(proof,contentsHash=p._digest(c['contents']),evidenceId='SYNTHETIC_UNIT_ONLY');reviewed(capture)
        second=p.prepare_port_update(m,r,c,s,phase='solver-outputs',previous_phase=adopted,phase_context=context(m,r,c),output_evidence=capture)
        self.assertEqual(second['previousProposalHash'],adopted['proposalHash'])
        with self.assertRaises(ValueError):p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c))
        bad=copy.deepcopy(ev);bad['elementHashes']['task5GoalSeekPython']='f'*64;reviewed(bad)
        with self.assertRaises(ValueError):p.prepare_port_update(m,r,c,s,phase_context=context(m,r,c),adoption_evidence=bad)
if __name__=='__main__':unittest.main()
