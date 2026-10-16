"""Offline regression contracts for actual SDK runtime gaps; never live evidence."""
import ast,copy,re,sys,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts/lib'))
import test_native_actions as helpers
import native_actions as native

class PublicNativeActionsTests(unittest.TestCase):
    APPROVAL_FIELDS=('Expense ID','Date','Employee','Department','Category','Vendor',
                     'Amount','Receipt Status','Approval Status','Card Amount','Ledger Amount')

    def formula_arguments(self,formula,function):
        """Read nested public function arguments without mistaking quoted text for syntax."""
        result=[]
        for match in re.finditer(r'\b'+re.escape(function)+r'\(',formula):
            start=match.end();depth=1;quoted=False;escaped=False
            for i in range(start,len(formula)):
                char=formula[i]
                if quoted:
                    if escaped:escaped=False
                    elif char=='\\':escaped=True
                    elif char=='"':quoted=False
                elif char=='"':quoted=True
                elif char=='(':depth+=1
                elif char==')':
                    depth-=1
                    if not depth:result.append(formula[start:i]);break
            else:self.fail('unclosed generated public formula')
        return result

    def fixture(self):
        helper=helpers.TestNativeActions();self.addCleanup(helper.doCleanups)
        return helper.bound_approval()
    def probe(self,m,r,b,c,cap,a,disabled):
        from action_inventory import digest
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':c['documentVersion'],
            **{k:disabled[k] for k in ('sourceHash','planHash','bindingHash','fragmentHash','contentsHash','assignedComponentsHash','sdkInventoryHash')},
            'activation':'selection-probe','probeOnly':True,'businessActionsDisabledInventoryHash':digest(cap)}
        return native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a,staging_review=review)

    def test_reviewed_control_only_probe_breaks_disabled_cardinality_bootstrap_without_writes(self):
        m,r,b,c,cap,a,disabled=self.public_fixture();probe=self.probe(m,r,b,c,cap,a,disabled)
        self.assertEqual(probe['activation'],'selection-probe')
        roles={e['role']:e for e in probe['actionEdits']};self.assertIn('clear-context',roles)
        for role,edit in roles.items():
            if role in {'selection','recovery','clear-context'}:
                self.assertEqual(edit['afterAction']['state'],'enabled')
                self.assertTrue(all(effect['effect']=='set-control-value' for effect in edit['afterAction']['effects']))
            else:self.assertEqual(edit['afterAction']['state'],'disabled')
        self.assertEqual(probe['productionStatus'],'BLOCKED')

    def test_public_formulas_reject_renamed_or_ambiguous_actual_input_even_freshly_rebound(self):
        from input_sync import _digest
        for attack in ('rename','duplicate'):
            m,r,b,c,cap,a,f=self.public_fixture();eid=b['tables']['expenseBusinessInput']['target']['elementId']
            source=next(e for e in c['contents']['elements'] if e['id']==eid)
            if attack=='rename':source['name']='Renamed Expense Source'
            else:next(e for e in c['contents']['elements'] if e['id']!=eid)['name']=source['name'].swapcase()
            b['contentsHash']=_digest(c['contents']);b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
            a.update(bindingHash=b['bindingHash'],contentsHash=b['contentsHash'])
            with self.subTest(attack=attack),self.assertRaisesRegex(ValueError,'name'):native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a)

    def test_selection_probe_rejects_enabled_omitted_business_sequence_and_incomplete_review(self):
        from action_inventory import configuration,digest
        m,r,b,c,cap,a,f=self.public_fixture();probe=self.probe(m,r,b,c,cap,a,f)
        for change in ({'probeOnly':False},{'businessActionsDisabledInventoryHash':'stale'},{'enable':True},{'activation':'unknown'}):
            with self.subTest(change=change),self.assertRaises(ValueError):native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a,staging_review={**probe['constructionReview'],**change})
        # A valid independently reviewed restoration artifact does not authorize
        # leaving an enabled native Prepare/other write running during a probe.
        from test_action_inventory import ActionInventoryTests
        _,old=ActionInventoryTests().fixture();old=ActionInventoryTests().reviewed(old)
        host=a['buttons']['approval:prepare_button']['elementId'];e=copy.deepcopy(old['elements']['button']);e['elementId']=host
        artifact=e.pop('restorationArtifact');artifact.update(elementId=host,sourceObservationHash=digest(e));e['restorationArtifact']=artifact;cap['elements'][host]=e
        f=native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a)
        with self.assertRaisesRegex(ValueError,'every other business action disabled'):self.probe(m,r,b,c,cap,a,f)

    def test_disabled_construction_cannot_supply_published_selection_proof(self):
        from action_inventory import qualify_public_selection_capture,qualify_public_selection_construction,digest
        from input_sync import _digest
        m,r,b,c,cap,a,f=self.public_fixture();fresh,after=self.post_actions(c,cap,f)
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':fresh['documentVersion'],
            'contentsHash':_digest(fresh['contents']),'captureHash':digest(after),**{k:f[k] for k in ('fragmentHash','sourceHash','planHash','bindingHash')},
            'publishedSelectionEvidence':self.selection_proofs(f,fresh['documentVersion'])}
        constructed=qualify_public_selection_construction(after,fresh,f,review)
        self.assertFalse(any(e.get('publicSelectionAuthority') for e in constructed['elements'].values()))
        with self.assertRaisesRegex(ValueError,'enabled'):qualify_public_selection_capture(after,fresh,f,review)
    def test_sdk_never_emits_global_selection_count_or_boolean_row_selector(self):
        m,r,b,c=self.fixture();result=native.emit_native_action_program(m,r,b,current_contents=c);tree=ast.parse(result['program'])
        calls=[n for n in ast.walk(tree) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name)]
        self.assertFalse(any(n.func.id=='count_distinct' for n in calls),'global CountDistinct is proven to count48 for one selected row')
        for n in calls:
            if n.func.id=='update_rows':self.assertIsInstance(n.args[1],ast.Dict,'Boolean row selectors are proven invalid_formula')
        self.assertEqual(result['constructionStatus'],'PARTIAL_REQUIRES_PUBLIC_ACTIONS')
        self.assertTrue(result['publicActionRequirements'])
        self.assertIn('Clear selection context',result['program']);self.assertIn('Captured expense context',result['program'])
        self.assertIn('.returns()',result['program'],'known Prepare ActionVariable must remain native')
    def public_fixture(self,normalized_ops=False,native_expense=False):
        from input_sync import _digest
        from action_inventory import empty_action_capture
        m,r,b,c=self.fixture()
        if native_expense:
            from input_sync import build_artifact_records
            from input_display import _system_column_contract
            target=r['targetBindings']['expenseBusinessInput'];target.pop('systemColumn')
            system={**_system_column_contract(target),'id':'ID'}
            b['tables']['expenseBusinessInput'].update(target=copy.deepcopy(target),systemColumn=system)
            r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings'])
            b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
        if normalized_ops:
            from input_sync import build_artifact_records
            target=r['targetBindings']['workflowOperations'];alias={'id':'ID','name':'ID','columnClass':'FormulaColumn','type':None,'systemKind':None,'formula':'[New Input Table/ID]','elementId':target['elementId'],
                'rowIdBinding':{'kind':'row_id','sourceColumn':'ID','valueType':'text','evidenceHash':'a'*64}}
            target['systemColumn']=alias;b['tables']['workflowOperations'].update(target=copy.deepcopy(target),systemColumn=alias)
            r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings'])
            b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
        for feature,record in b['tables'].items():
            next(e for e in c['contents']['elements'] if e['id']==record['target']['elementId'])['name']=m['planProjection']['sharedInputArchitecture']['inputs'][feature]['name']
        b['contentsHash']=_digest(c['contents']);b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
        routes=native.emit_native_action_program(m,r,b,current_contents=c)
        assigned={'controls':{},'buttons':{},'actions':{}}
        for alias,record in routes['componentControls'].items():
            eid='assigned-'+alias;node={'id':eid,'kind':'control','controlId':record['variableId'],'controlType':record['type'],'value':0 if record['type']=='number' else ''}
            c['contents']['elements'].append(node);c['contents']['layout']=c['contents']['layout'].replace('</Page>','<Element elementId="'+eid+'" /></Page>')
            assigned['controls'][alias]={'elementId':eid,'pageId':'host','elementHash':_digest(node)}
        for key,record in routes['componentButtons'].items():
            eid='assigned-'+key.replace(':','-');node={'id':eid,'kind':'button','text':record['text'],'appearance':'filled'}
            c['contents']['elements'].append(node);c['contents']['layout']=c['contents']['layout'].replace('</Page>','<Element elementId="'+eid+'" /></Page>')
            assigned['buttons'][key]={'elementId':eid,'pageId':'host','elementHash':_digest(node)}
        for route in routes['publicActionRequirements']:assigned['actions'][route['workflow']+':'+route['role']]={'sequenceId':'public-'+route['workflow']+'-'+route['role'],'beforeActionHash':None}
        b['contentsHash']=_digest(c['contents']);b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
        assigned.update(target=m['target'],bindingHash=b['bindingHash'],contentsHash=b['contentsHash'])
        cap=empty_action_capture(c,evidence_id='offline-fixture');cap.update(origin='sigma-sdk-readonly',complete=True)
        for e in cap['elements'].values():e.update(complete=True,inspection='sdk-type-no-action-events')
        fragments=native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=assigned)
        return m,r,b,c,cap,assigned,fragments
    def post_actions(self,current,capture,fragments):
        from action_inventory import digest,configuration
        fresh=copy.deepcopy(current);fresh['documentVersion']+=1
        cap=copy.deepcopy(capture);cap.update(version=fresh['documentVersion'],evidenceId='offline-post')
        for edit in fragments['actionEdits']:
            host=next(e for e in fresh['contents']['elements'] if e['id']==edit['hostElementId']);host.setdefault('actions',[]).append(copy.deepcopy(edit['afterAction']))
            action=edit['afterAction'];blocks=[]
            for effect in action['effects']:
                if effect['effect']=='set-control-value':
                    count=effect['control'].endswith('selection-count');value=effect['value']
                    if value.get('formula','').startswith('CountDistinct([Selection/'):value={'type':'unavailable','reason':'SDK_SELECTION_NAMESPACE_NO_GETTER'}
                    control_element=next(e['id'] for e in current['contents']['elements'] if e.get('kind')=='control' and e.get('controlId')==effect['control'])
                    blocks.append({'kind':'FloatSetControlValueEffect' if count else 'TextSetControlValueEffect','fields':{'controlId':effect['control'],'controlElementId':control_element,'value':value}})
                else:blocks.append({'kind':'UpdateRowsEffect','fields':{'tableId':effect['tableElementId'],'whichRows':effect['whichRows'],'values':effect['values']}})
            element=cap['elements'][host['id']];element['sequences'].append({'id':action['id'],'name':action['name'],'trigger':action['trigger'] if isinstance(action['trigger'],str) else action['trigger']['on'],
                'disabled':action['state']=='disabled','condition':action['trigger'].get('condition') if isinstance(action['trigger'],dict) else None,'blocks':blocks,'successToast':None})
            element['semanticHash']=digest(configuration(element))
        return fresh,cap
    def selection_proofs(self,fragments,version):
        return {a['hostElementId']+':'+a['controlElementId']:{name:{'evidenceId':'offline-'+name,'origin':'sigma-browser','published':True,
            'documentId':'doc','version':version,'hostElementId':a['hostElementId'],'controlElementId':a['controlElementId'],
            'memberId':'fixture-member','identityEvidenceId':'offline-banner','adminBypass':False,'actualCount':count}
            for name,count in [('zero',0),('one',1),('multiple',2),('clear',0)]} for a in fragments['selectionAuthorities']}
    def test_selection_authority_is_exact_source_bound_actual_proof_not_missing_node_repr(self):
        from action_inventory import qualify_public_selection_capture,digest
        from input_sync import _digest
        m,r,b,c,cap,assigned,f=self.public_fixture();f=self.probe(m,r,b,c,cap,assigned,f);fresh,after=self.post_actions(c,cap,f)
        runtime=self.selection_proofs(f,9);key=next(iter(runtime))
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':9,'contentsHash':_digest(fresh['contents']),
            'captureHash':digest(after),**{k:f[k] for k in ('fragmentHash','sourceHash','planHash','bindingHash')},'publishedSelectionEvidence':runtime}
        result=qualify_public_selection_capture(after,fresh,f,review)
        for authority in f['selectionAuthorities']:
            e=result['elements'][authority['hostElementId']];seq=next(s for s in e['sequences'] if s['id']==authority['sequenceId'])
            self.assertEqual(seq['blocks'][authority['effectIndex']]['fields']['value'],{'type':'formula','formula':authority['formula']})
            self.assertIn('publicSelectionAuthority',e);self.assertNotIn('restorationArtifact',e)
        changes=[('pending',{'owner':'PENDING'}),('missing',{'publishedSelectionEvidence':{}}),('stale',{'version':8}),
            ('missing-recovery',{'publishedSelectionEvidence':{key:runtime[key]}})]
        for field,value in [('actualCount',48),('adminBypass',True),('hostElementId','foreign'),('controlElementId','foreign')]:
            altered=copy.deepcopy(runtime);altered[key]['one'][field]=value;changes.append((field,{'publishedSelectionEvidence':altered}))
        for name,change in changes:
            with self.subTest(change=name),self.assertRaises(ValueError):qualify_public_selection_capture(after,fresh,f,{**review,**change})
        bad=copy.deepcopy(after);a=f['selectionAuthorities'][0];e=bad['elements'][a['hostElementId']];e['sequences'][0]['blocks'][a['effectIndex']]['fields']['value']='Missing node'
        from action_inventory import configuration
        e['semanticHash']=digest(configuration(e))
        with self.assertRaises(ValueError):qualify_public_selection_capture(bad,fresh,f,{**review,'captureHash':digest(bad)})
        bad=copy.deepcopy(after);bad['elements'][a['hostElementId']]['sequences'][0]['trigger']='on-click'
        bad['elements'][a['hostElementId']]['semanticHash']=digest(configuration(bad['elements'][a['hostElementId']]))
        with self.assertRaises(ValueError):qualify_public_selection_capture(bad,fresh,f,{**review,'captureHash':digest(bad)})
        wrong=copy.deepcopy(fresh);host=next(e for e in wrong['contents']['elements'] if e['id']==a['hostElementId']);host['actions'][0]['effects'][a['effectIndex']]['value']['formula']='CountDistinct([Wrong/ID])'
        with self.assertRaises(ValueError):qualify_public_selection_capture(after,wrong,f,{**review,'contentsHash':_digest(wrong['contents'])})

    def test_public_routes_preserve_complete_frozen_predicate_and_exact_selection_scope(self):
        m,r,b,c,cap,assigned,f=self.public_fixture()
        by={e['role']:e for e in f['actionEdits']};self.assertEqual(set(by),{'selection','recovery','confirm','clear-context'})
        selection=by['selection']['afterAction'];self.assertEqual(selection['trigger'],'on-select')
        self.assertEqual(selection['effects'][2]['value'],{'type':'formula','formula':'CountDistinct([Selection/ID])'})
        confirm=by['confirm']['afterAction'];update=confirm['effects'][0]
        self.assertEqual(update['whichRows']['type'],'formula')
        self.assertIn('([ID] = [approval-physical-row])',update['whichRows']['formula'])
        guard=confirm['trigger']['condition']['formula'];self.assertIn('CountIf(',guard);self.assertIn('Applied Approval Operation IDs',guard)
        for column in m['planProjection']['sharedInputArchitecture']['inputs']['expenseBusinessInput']['columns'][:-3]:
            self.assertIn('Expense Business Inputs/'+column['name'],guard);self.assertIn('Before:'+column['name'],guard)
        for token in (m['sourceHash'],m['planHash'],'approval-context','approval-selection-count','approval-operation','approval-record','approval-physical-row'):
            self.assertIn(token,guard)
        self.assertEqual(update['values'][b['tables']['expenseBusinessInput']['target']['columnIds']['ApprovedAt']],{'type':'formula','formula':'Now()'})
        self.assertNotIn('successToast',confirm);self.assertEqual(confirm['state'],'disabled')
        self.assertTrue(f['selectionAuthorities']);self.assertNotIn('Missing node',str(f));self.assertNotIn('returns',str(f['actionEdits']))

    def test_approval_event_freeze_is_scalar_not_lookup_inside_countif(self):
        # Native published evidence: all11 scalar equalities are true, while
        # CountIf(rowField = Lookup(...)) returns0 for same/cross-source lookups.
        # This checks our emitted contract, not a simulated Sigma execution.
        m,r,b,c,cap,a,f=self.public_fixture()
        action=next(e['afterAction'] for e in f['actionEdits'] if e['role']=='confirm')
        guard=action['trigger']['condition']['formula']
        counts=self.formula_arguments(guard,'CountIf')
        self.assertTrue(counts)
        self.assertTrue(all('Lookup(' not in expression for expression in counts),
                        'approval must not hoist frozen Lookup comparisons into CountIf')
        for field in self.APPROVAL_FIELDS:
            current=f'Lookup([Expense Business Inputs/{field}], [approval-physical-row], [Expense Business Inputs/ID])'
            before=f'Lookup([Workflow Operations/Before:{field}], [approval-operation], [Workflow Operations/Row ID])'
            self.assertIn(f'((IsNull({current}) and IsNull({before})) or ({current} = {before}))',guard,field)
        self.assertIn('PREPARED',guard)
        self.assertIn('Lookup([Expense Business Inputs/Applied Approval Operation IDs]',guard)
        self.assertEqual(action['state'],'disabled');self.assertEqual(f['productionStatus'],'BLOCKED')

    def test_approval_public_write_selector_rechecks_all_frozen_fields(self):
        m,r,b,c,cap,a,f=self.public_fixture()
        action=next(e['afterAction'] for e in f['actionEdits'] if e['role']=='confirm')
        selector=action['effects'][0]['whichRows']
        self.assertEqual(selector['type'],'formula','physical-ID-only writes discard the write-boundary freeze')
        formula=selector['formula']
        self.assertNotIn('CountIf(',formula)
        self.assertIn('([ID] = [approval-physical-row])',formula)
        self.assertIn('([Expense ID] = [approval-record])',formula)
        self.assertIn('[Applied Approval Operation IDs]',formula)
        for field in self.APPROVAL_FIELDS:
            before=f'Lookup([Workflow Operations/Before:{field}], [approval-operation], [Workflow Operations/Row ID])'
            self.assertIn(f'((IsNull([{field}]) and IsNull({before})) or ([{field}] = {before}))',formula,field)
        self.assertEqual(set(action['effects'][0]['values']),
                         {b['tables']['expenseBusinessInput']['target']['columnIds'][field] for field in
                          ('Approval Status','ApprovedAt','Last Approval Operation ID','Applied Approval Operation IDs')})

    def test_history_value_uses_native_ui_bound_selected_lookup(self):
        # Native UI saved Lookup with target-owned physical column IDs inside
        # the value. Definition validation and actual effects are separate gates.
        for native_expense in (False,True):
            with self.subTest(native_expense=native_expense):
                m,r,b,c,cap,a,f=self.public_fixture(native_expense=native_expense)
                action=next(e['afterAction'] for e in f['actionEdits'] if e['role']=='confirm')
                expense=b['tables']['expenseBusinessInput']
                history=expense['target']['columnIds']['Applied Approval Operation IDs']
                rowid=expense['systemColumn']['id']
                selected_history=f'Lookup([{history}], [approval-physical-row], [{rowid}])'
                self.assertEqual(action['effects'][0]['values'][history],{'type':'formula','formula':
                    '(Coalesce('+selected_history+', "") & (("|" & [approval-operation]) & "|"))'})
                self.assertIn('Lookup([Expense Business Inputs/Applied Approval Operation IDs]',action['trigger']['condition']['formula'])
                self.assertIn('[Applied Approval Operation IDs]',action['effects'][0]['whichRows']['formula'])
                self.assertEqual(action['state'],'disabled');self.assertNotIn('successToast',action)

    def test_value_id_context_preserves_cross_source_and_selector_names(self):
        columns={'history':{'elementId':'expense','tableName':'Expense','name':'History','id':'history-id'},
                 'rowid':{'elementId':'expense','tableName':'Expense','name':'ID','id':'row-id'},
                 'before':{'elementId':'ops','tableName':'Operations','name':'Before:History','id':'before-id'},
                 'opsid':{'elementId':'ops','tableName':'Operations','name':'ID','id':'ops-id'}}
        controls={'selected':{'variableId':'selected-row'},'operation':{'variableId':'operation'}}
        node=ast.parse('eq(lookup(history, selected, rowid), lookup(before, operation, opsid))',mode='eval').body
        self.assertEqual(native._public_formula(node,columns,controls,local_element_id='expense',local_column_ids=True),
                         '(Lookup([history-id], [selected-row], [row-id]) = Lookup([Operations/Before:History], [operation], [Operations/ID]))')
        self.assertEqual(native._public_formula(node,columns,controls,local_element_id='expense'),
                         '(Lookup([History], [selected-row], [ID]) = Lookup([Operations/Before:History], [operation], [Operations/ID]))')
        self.assertEqual(native._public_formula(node,columns,controls),
                         '(Lookup([Expense/History], [selected-row], [Expense/ID]) = Lookup([Operations/Before:History], [operation], [Operations/ID]))')
        for malformed in ('',None,'history] & True & [other'):
            with self.subTest(malformed=malformed),self.assertRaisesRegex(ValueError,'identifier'):
                native._public_formula(node,{**columns,'history':{**columns['history'],'id':malformed}},controls,
                                       local_element_id='expense',local_column_ids=True)

    def test_native_id_approval_also_defers_full_formula_selector_instead_of_sdk_hoisting(self):
        m,r,b,c,cap,a,f=self.public_fixture(native_expense=True)
        emitted=native.emit_native_action_program(m,r,b,current_contents=c)
        self.assertNotIn('update_rows(t_expenseBusinessInput',emitted['program'])
        self.assertTrue(any(q['workflow']=='approval' and q['role']=='confirm' for q in emitted['publicActionRequirements']))
        action=next(e['afterAction'] for e in f['actionEdits'] if e['role']=='confirm')
        self.assertEqual(action['effects'][0]['whichRows']['type'],'formula')
        self.assertTrue(all('Lookup(' not in expression for expression in
                            self.formula_arguments(action['trigger']['condition']['formula'],'CountIf')))
    def test_public_routing_rejects_foreign_components_stale_state_and_missing_recovery(self):
        from input_sync import _digest
        m,r,b,c,cap,assigned,f=self.public_fixture()
        for change in ('control','contents','binding','schema','inventory'):
            cc,aa,ss=copy.deepcopy(c),copy.deepcopy(assigned),copy.deepcopy(cap)
            if change=='control':aa['controls']['approval_count']['elementId']='foreign'
            elif change=='contents':cc['documentVersion']=7
            elif change=='binding':aa['bindingHash']='a'*64
            elif change=='schema':aa['controls']['approval_count']['elementHash']='f'*64
            else:ss['elements'].pop(next(iter(ss['elements'])))
            with self.subTest(change=change),self.assertRaises(ValueError):native.emit_public_native_action_fragments(m,r,b,current_contents=cc,sdk_actions=ss,assigned_components=aa)
        # A fresh hash and still-identical approved column-ID binding cannot
        # authorize a same-ID source whose actual schema has changed.
        bad=copy.deepcopy(c);bad['contents']['elements'][0]['columns'][0]['name']='Foreign field'
        fresh_binding=copy.deepcopy(b);fresh_binding['contentsHash']=_digest(bad['contents']);fresh_binding['bindingHash']=_digest({k:v for k,v in fresh_binding.items() if k!='bindingHash'})
        rr=copy.deepcopy(r);rr['nativeActionBinding']=fresh_binding;aa=copy.deepcopy(assigned);aa.update(bindingHash=fresh_binding['bindingHash'],contentsHash=fresh_binding['contentsHash'])
        with self.assertRaisesRegex(ValueError,'source kind/page/column'):native.emit_public_native_action_fragments(m,rr,fresh_binding,current_contents=bad,sdk_actions=cap,assigned_components=aa)
        eid=assigned['buttons']['approval:prepare_button']['elementId'];cap['elements'][eid]['sequences']=[{'id':'omitted-prepare','name':'Prepare','trigger':'on-click','condition':None,'blocks':[],'disabled':True,'successToast':None}]
        from action_inventory import digest,configuration
        cap['elements'][eid]['semanticHash']=digest(configuration(cap['elements'][eid]))
        with self.assertRaisesRegex(ValueError,'restoration'):native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=assigned)
    def test_exact_action_only_adapter_preserves_unrelated_content_and_guards_finalize(self):
        import workflow_repairs as repair
        from input_sync import _digest
        m,r,b,c,cap,assigned,f=self.public_fixture()
        scope={'documentId':'doc','version':8,'folderId':'folder','folderEvidenceId':'actual-fixture','connectionId':'conn'}
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','beforeHash':_digest(c['contents']),
            **{k:f[k] for k in ('sourceHash','planHash','bindingHash','fragmentHash','assignedComponentsHash','sdkInventoryHash')},
            'allowedActionEdits':f['allowedEditPaths'],'readbackEvidenceHashes':dict.fromkeys(['fullRows','schemaDisplay','sourceGraph'],'a'*64)}
        from action_inventory import configuration,digest
        after,post_capture=self.post_actions(c,cap,f)
        raw_capture=copy.deepcopy(post_capture)
        from action_inventory import qualify_public_selection_construction
        selection_review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':9,'contentsHash':_digest(after['contents']),
            'captureHash':digest(post_capture),**{k:f[k] for k in ('fragmentHash','sourceHash','planHash','bindingHash')},
            'publishedSelectionEvidence':self.selection_proofs(f,9)}
        post_capture=qualify_public_selection_construction(post_capture,after,f,selection_review)

        review['sdkAllowedChanges']={host:configuration(post_capture['elements'][host]) for host in {e['hostElementId'] for e in f['actionEdits']}}
        proposal=repair.prepare_native_action_update(m,r,b,f,c,scope=scope,review=review,sdk_actions=cap,assigned_components=assigned)
        self.assertEqual(proposal['dryRunRequest']['contents']['layout'],c['contents']['layout'])
        for old in c['contents']['elements']:
            new=next(e for e in proposal['dryRunRequest']['contents']['elements'] if e['id']==old['id'])
            self.assertEqual({k:v for k,v in old.items() if k!='actions'},{k:v for k,v in new.items() if k!='actions'})
        for change in ({'allowedActionEdits':[]},{'readbackEvidenceHashes':{}},{'fragmentHash':'stale'},{'sdkAllowedChanges':{}}):
            with self.assertRaises(ValueError):repair.prepare_native_action_update(m,r,b,f,c,scope=scope,review={**review,**change},sdk_actions=cap,assigned_components=assigned)
        bad=copy.deepcopy(f);bad['actionEdits'][0]['afterAction']['effects'][2]['value']['formula']='CountDistinct([Whole table/ID])'
        with self.assertRaises(ValueError):repair.prepare_native_action_update(m,r,b,bad,c,scope=scope,review=review,sdk_actions=cap,assigned_components=assigned)
        post_review={'status':'APPROVED_STAGING','owner':'offline-reviewer','proposalHash':proposal['proposalHash'],'contentsHash':_digest(after['contents']),
            'actionInventoryHash':digest(raw_capture),'allowedSemanticChangesHash':digest(review['sdkAllowedChanges'])}
        self.assertEqual(repair.verify_native_action_update_readback(proposal,after,raw_capture,post_review)['status'],'VERIFIED_NATIVE_ACTION_CONSTRUCTION_RUNTIME_UNRESOLVED')
        with self.assertRaises(ValueError):repair.verify_native_action_update_readback(proposal,after,raw_capture,{**post_review,'allowedSemanticChangesHash':'stale'})
        bad=copy.deepcopy(raw_capture);authority=f['selectionAuthorities'][0];element=bad['elements'][authority['hostElementId']]
        element['sequences'][0]['blocks'][authority['effectIndex']]['fields']['controlElementId']='foreign'
        element['semanticHash']=digest(configuration(element))
        with self.assertRaisesRegex(ValueError,'arbitrary opaque'):repair.verify_native_action_update_readback(proposal,after,bad,{**post_review,'actionInventoryHash':digest(bad)})
        dry={'httpStatus':200,'dryRun':True,'evidenceId':'offline-only','proposalHash':proposal['proposalHash']}
        self.assertEqual(repair.finalize_repair_update(proposal,c,dry,fresh_sdk_actions=cap)['body']['documentVersion'],8)
        with self.assertRaises(ValueError):repair.finalize_repair_update(proposal,{**c,'documentVersion':9},dry,fresh_sdk_actions=cap)
        with self.assertRaises(ValueError):repair.finalize_repair_update(proposal,c,dry,consumed_proposals=[proposal['proposalHash']],fresh_sdk_actions=cap)

    def test_normalized_operations_id_uses_public_single_row_and_never_sdk_cast(self):
        m,r,b,c,cap,assigned,f=self.public_fixture(normalized_ops=True)
        result=native.emit_native_action_program(m,r,b,current_contents=c)
        self.assertNotIn('update_rows(t_workflowOperations',result['program'])
        self.assertNotIn('add_system_columns',result['program'])
        action=next(e['afterAction'] for e in f['actionEdits'] if e['role']=='new-intent')
        self.assertEqual(action['effects'][0]['whichRows']['primaryKeys'],{'ID':{'type':'control','control':'approval-operation'}})
        self.assertIn('CountIf(',action['trigger']['condition']['formula'])
    def test_edited_host_cannot_use_semantic_review_to_drop_unrelated_omitted_sdk_sequence(self):
        import workflow_repairs as repair
        from action_inventory import digest,configuration
        from input_sync import _digest
        from test_action_inventory import ActionInventoryTests
        m,r,b,c,cap,a,f=self.public_fixture();host=f['actionEdits'][0]['hostElementId']
        # Reuse the independent recovery fixture for an unrelated REST-omitted sequence.
        _,old=ActionInventoryTests().fixture();old=ActionInventoryTests().reviewed(old);e=copy.deepcopy(old['elements']['button']);e['elementId']=host
        artifact=e.pop('restorationArtifact');artifact.update(elementId=host,sourceObservationHash=digest(e));e['restorationArtifact']=artifact;cap['elements'][host]=e
        f=native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a)
        after,post=self.post_actions(c,cap,f)
        allowed={h:configuration(post['elements'][h]) for h in {ed['hostElementId'] for ed in f['actionEdits']}}
        allowed[host]['sequences'][0]['disabled']=True  # Same count, unauthorized semantic change.
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','beforeHash':_digest(c['contents']),
            **{k:f[k] for k in ('sourceHash','planHash','bindingHash','fragmentHash','assignedComponentsHash','sdkInventoryHash')},
            'allowedActionEdits':f['allowedEditPaths'],'sdkAllowedChanges':allowed,'readbackEvidenceHashes':dict.fromkeys(['fullRows','schemaDisplay','sourceGraph'],'a'*64)}
        scope={'documentId':'doc','version':8,'folderId':'folder','folderEvidenceId':'offline','connectionId':'conn'}
        with self.assertRaisesRegex(ValueError,'unrelated SDK sequence'):repair.prepare_native_action_update(m,r,b,f,c,scope=scope,review=review,sdk_actions=cap,assigned_components=a)
    def test_sdk_enabling_requires_exact_disabled_program_hash_and_cli_public_route(self):
        import json,subprocess,tempfile
        from input_sync import _json_value
        m,r,b,c,cap,assigned,f=self.public_fixture();candidate=native.emit_native_action_program(m,r,b,current_contents=c)
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','sourceHash':m['sourceHash'],'planHash':m['planHash'],'documentId':'doc','version':8,'bindingHash':b['bindingHash'],'candidateProgramHash':candidate['candidateProgramHash']}
        self.assertEqual(native.emit_native_action_program(m,r,b,current_contents=c,staging_review=review)['constructionStatus'],'PARTIAL_REQUIRES_PUBLIC_ACTIONS')
        for change in ({'candidateProgramHash':'stale'},{'owner':'PENDING'},{'sourceHash':'foreign'},{'enable':True}):
            with self.assertRaises(ValueError):native.emit_native_action_program(m,r,b,current_contents=c,staging_review={**review,**change})
        with tempfile.TemporaryDirectory() as tmp:
            bundle=Path(tmp)/'bundle.json';out=Path(tmp)/'out.json'
            bundle.write_text(json.dumps(_json_value({'manifest':m,'receipt':r,'action_bindings':b,'current_contents':c,'sdk_actions':cap,'assigned_components':assigned})))
            result=subprocess.run([sys.executable,str(ROOT/'scripts/migration-acceptance.py'),'public-native-actions','--bundle',str(bundle),'--out',str(out)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr);self.assertEqual(json.loads(out.read_text())['fragmentHash'],f['fragmentHash'])

    def test_public_enabling_requires_exact_prior_disabled_actions_and_fresh_sdk_proof(self):
        # All evidence here is synthetic unit data, never actual runtime proof.
        import hashlib
        from input_sync import _digest
        from action_inventory import digest
        m,r,b,c,cap,a,f=self.public_fixture()
        fields=('sourceHash','planHash','bindingHash','fragmentHash','contentsHash','assignedComponentsHash','sdkInventoryHash')
        review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':8,
            **{k:f[k] for k in fields},'enable':True,'disabledConstructionProofHash':digest(cap)}
        with self.assertRaises(ValueError):native.emit_public_native_action_fragments(m,r,b,current_contents=c,sdk_actions=cap,assigned_components=a,staging_review=review)
        f=self.probe(m,r,b,c,cap,a,f)
        fresh,after=self.post_actions(c,cap,f);f=copy.deepcopy(f)
        # The new probe version is observed, never assigned onto old proof.
        m['target']['version']=fresh['documentVersion'];b['target']['version']=fresh['documentVersion'];a['target']['version']=fresh['documentVersion']
        from input_sync import build_artifact_records
        for target in r['targetBindings'].values():target['version']=fresh['documentVersion']
        for record in b['tables'].values():record['target']['version']=fresh['documentVersion']
        r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings'])
        from action_inventory import qualify_public_selection_capture
        selection_review={'status':'APPROVED_STAGING','owner':'offline-reviewer','documentId':'doc','version':fresh['documentVersion'],'contentsHash':_digest(fresh['contents']),
            'captureHash':digest(after),**{k:f[k] for k in ('fragmentHash','sourceHash','planHash','bindingHash')},
            'publishedSelectionEvidence':self.selection_proofs(f,fresh['documentVersion'])}
        after=qualify_public_selection_capture(after,fresh,f,selection_review)
        # This fixture exercises the existing recovery contract, not a recovery shortcut.
        for eid,e in after['elements'].items():
            if not e['sequences']:continue
            program='# offline missing-only test artifact\n';ph=hashlib.sha256(program.encode()).hexdigest()
            e['restorationArtifact']={'elementId':eid,'documentId':'doc','reviewer':'offline-reviewer','sourceObservationHash':digest(e),
                'semanticHash':e['semanticHash'],'program':program,'programHash':ph,'constructionArtifactHash':'offline','bindingHash':b['bindingHash'],
                'guardPolicy':'missing-sequences-only; exact-current-target-and-bindings; no-constructor-replay',
                'restoreProof':{'status':'PASSED_ISOLATED_RESTORATION','evidenceId':'offline-only','documentId':'isolated-fixture',
                    'semanticHash':e['semanticHash'],'sdkReadbackEvidenceId':'offline-only','programHash':ph,'bindingHash':b['bindingHash'],'beforeMissingSequences':True}}
        b['contentsHash']=_digest(fresh['contents']);b['bindingHash']=_digest({k:v for k,v in b.items() if k!='bindingHash'});r['nativeActionBinding']=b
        a.update(bindingHash=b['bindingHash'],contentsHash=b['contentsHash'])
        for record in [*a['controls'].values(),*a['buttons'].values()]:record['elementHash']=_digest(next(e for e in fresh['contents']['elements'] if e['id']==record['elementId']))
        for edit in f['actionEdits']:a['actions'][edit['workflow']+':'+edit['role']]['beforeActionHash']=edit['afterActionHash']
        disabled=native.emit_public_native_action_fragments(m,r,b,current_contents=fresh,sdk_actions=after,assigned_components=a)
        review.update({k:disabled[k] for k in fields});review['version']=fresh['documentVersion'];review['disabledConstructionProofHash']=digest(after)
        enabled=native.emit_public_native_action_fragments(m,r,b,current_contents=fresh,sdk_actions=after,assigned_components=a,staging_review=review)
        self.assertTrue(all(e['afterAction']['state']=='enabled' for e in enabled['actionEdits']))
        self.assertEqual(enabled['status'],'ENABLED_REVIEWED_STAGING_ONLY_REQUIRES_RUNTIME')
        for change in ({'disabledConstructionProofHash':'stale'},{'owner':'PENDING'},{'fragmentHash':f['fragmentHash']}):
            with self.subTest(change=change),self.assertRaises(ValueError):native.emit_public_native_action_fragments(m,r,b,current_contents=fresh,sdk_actions=after,assigned_components=a,staging_review={**review,**change})
        stale=copy.deepcopy(after)
        for element in stale['elements'].values():
            if element.get('publicSelectionAuthority'):
                element['publicSelectionAuthority'][0]['review']['version']-=1
                artifact=element.pop('restorationArtifact');artifact['sourceObservationHash']=digest(element);element['restorationArtifact']=artifact
                break
        candidate=native.emit_public_native_action_fragments(m,r,b,current_contents=fresh,sdk_actions=stale,assigned_components=a)
        updated={**review,**{k:candidate[k] for k in fields},'disabledConstructionProofHash':digest(stale)}
        with self.assertRaisesRegex(ValueError,'fresh qualified'):native.emit_public_native_action_fragments(m,r,b,current_contents=fresh,sdk_actions=stale,assigned_components=a,staging_review=updated)

    def test_public_fragment_and_narrow_adapter_entrypoints_exist(self):
        self.assertTrue(callable(getattr(native,'emit_public_native_action_fragments',None)))
        import workflow_repairs
        self.assertTrue(callable(getattr(workflow_repairs,'prepare_native_action_update',None)))
    def test_research_declared_sort_uses_and_hides_existing_source_order_column(self):
        from test_workflow_repairs import WorkflowRepairsTests
        h=WorkflowRepairsTests();self.addCleanup(h.doCleanups);p=h.module();contract=h.contract('research')
        from presentation import extract_presentation,build_business_input_architecture
        from test_presentation import research_runtime_plan
        arch=build_business_input_architecture(extract_presentation(h.fixture('research')),connection_id='conn',folder_id='folder',research_plan=research_runtime_plan())
        spec=arch['inputs'][arch['presentationInputFeature']]
        target={'elementId':'input','systemColumnId':'ID','systemEvidenceHash':'a'*64,'columnIds':{c['name']:'c'+str(i) for i,c in enumerate(spec['columns'])}}
        fragment=p.build_normalized_input_audit(contract,spec,target,connection_id='conn');audit=next(e for e in fragment['elements'] if e['id']=='research-normalized-input-audit')
        declared={c['id'] for c in audit['columns']}
        self.assertTrue(all(s['columnId'] in declared for s in audit['sort']),'actual API rejects undeclared source_order')
        existing=[c for c in audit['columns'] if c['name']=='SourceOrder'];self.assertEqual(len(existing),1);self.assertTrue(existing[0]['hidden'])
        self.assertIn('\n\nThe source inputs audit',next(e['body'] for e in fragment['elements'] if e['id']=='research-input-guidance'))

    def test_research_edit_grid_has_visible_rows_and_guide_does_not_overlap(self):
        import workflow_repairs as repair
        from input_sync import _digest
        c={'elements':[{'id':'input','kind':'input-table'},{'id':'guide','kind':'text','body':'guide'},{'id':'title','kind':'text'}],
            'layout':'<Page id="research-edit"><Element elementId="title" gridRow="1 / 4"/><Element elementId="input" gridRow="4 / 10" gridColumn="1 / 25"/><Element elementId="guide" gridRow="11 / 25"/></Page>'}
        args={'input_element_id':'input','guidance_element_id':'guide','page_id':'research-edit','input_element_hash':_digest(c['elements'][0]),'guidance_element_hash':_digest(c['elements'][1]),'layout_hash':_digest(c['layout'])}
        result=repair.repair_research_edit_layout(c,**args)
        self.assertEqual(result['elements'],c['elements']);self.assertIn('gridRow="4 / 18"',result['layout']);self.assertIn('gridRow="20 / 34"',result['layout'])
        with self.assertRaises(ValueError):repair.repair_research_edit_layout(c,**{**args,'layout_hash':'stale'})
        crowded=copy.deepcopy(c);crowded['layout']=crowded['layout'].replace('</Page>','<Element elementId="foreign" gridRow="18 / 23"/></Page>')
        with self.assertRaises(ValueError):repair.repair_research_edit_layout(crowded,**{**args,'layout_hash':_digest(crowded['layout'])})
        taller=copy.deepcopy(c);taller['layout']=taller['layout'].replace('4 / 10','4 / 28')
        with self.assertRaises(ValueError):repair.shrink_input_layout(taller,{'input':{'elementHash':args['input_element_hash'],'pageId':'research-edit','fromGridRow':'4 / 28','toGridRow':'4 / 10'}})
if __name__=='__main__':unittest.main()
