"""Approved input-display contracts and independent SDK metadata observations."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT/'scripts'
sys.path.insert(0, str(SCRIPTS/'lib'))
from test_presentation import research_runtime_plan
from test_approved_wiring import receipt
import presentation
from migration_manifest import build_manifest


class TestInputDisplay(unittest.TestCase):
    def prepared(self, workflow):
        names={'research':'research-model','forecast':'forecast-model','sales':'sales-pipeline',
               'inventory':'inventory-reorder','expense':'expense-approval'}
        source=ROOT/'tests'/'fixtures'/('fixture-'+names[workflow]+'.xlsx')
        plan=presentation.extract_presentation(source)
        architecture=presentation.build_business_input_architecture(plan,connection_id='conn',folder_id='folder',
            research_plan=research_runtime_plan() if workflow=='research' else None)
        logical=presentation.build_business_input_plan(plan,architecture)
        manifest=build_manifest(source,logical)
        manifest['target']={'documentId':'doc','version':8}
        return source,plan,manifest,receipt(manifest)

    def test_public_display_emission_covers_five_assigned_schemas_without_data_or_validation_writes(self):
        for workflow in ['research','forecast','sales','inventory','expense']:
            with self.subTest(workflow=workflow), tempfile.TemporaryDirectory() as directory:
                source,plan,manifest,staging=self.prepared(workflow)
                tmp=Path(directory)
                for name,data in [('manifest',manifest),('receipt',staging)]:
                    (tmp/(name+'.json')).write_text(json.dumps(data))
                result=subprocess.run([sys.executable,str(SCRIPTS/'presentation-inputs.py'),'display',str(source),
                    '--manifest',str(tmp/'manifest.json'),'--receipt',str(tmp/'receipt.json'),
                    '--out',str(tmp/'display.json')],capture_output=True,text=True)
                self.assertEqual(result.returncode,0,result.stderr)
                display=json.loads((tmp/'display.json').read_text())
                self.assertEqual(display['status'],'PENDING_SDK_READBACK')
                feature=manifest['planProjection']['sharedInputArchitecture']['presentationInputFeature']
                self.assertEqual({c['id'] for c in display['columns']},set(staging['targetBindings'][feature]['columnIds'].values()))
                self.assertIn('set_visible_columns(',display['program'])
                self.assertIn('set_format',display['program'])
                for forbidden in ('update_rows','upsert_rows','add_editable_column','set_validation','set_filter'):
                    self.assertNotIn(forbidden,display['program'])
                self.assertIn('.is_visible()',display['readbackProgram'])
                for method in ('.kind()', '.decimal_places()', '.currency()', '.pattern()'):
                    self.assertIn(method,display['readbackProgram'])
                self.assertNotIn('set_format',display['readbackProgram'])
                self.assertIn(manifest['planHash'],display['program'])
                if workflow=='research':
                    self.assertEqual(len(display['columns']),146)
                    self.assertEqual([c['name'] for c in display['columns'] if c['visible']],
                                     ['Scope',*(f'i_model_inputs_c{i}' for i in range(2,8))])
                    money=next(c for c in display['columns'] if c['name']=='i_fy_results_b56')
                    self.assertEqual(money['format']['pattern'],'($,.1f')
                    self.assertEqual(money['sourceFormat'],'$#,##0.0;($#,##0.0);-')
                    self.assertEqual(money['fidelity'],'PARTIAL')
                    self.assertIn('zero',money['unresolvedReason'])
                    growth=next(c for c in display['columns'] if c['name']=='i_model_inputs_c2')
                    self.assertEqual(growth['format'],{'pattern':'.1%'})

    def captures(self, display, columns=None):
        base=display['target']['version']
        columns=columns or [dict(id=c['id'],name=c['name'],visible=c['visible'],
            columnClass=c['columnClass'],type=c['type'],systemKind=None,elementId=display['target']['elementId'],formula=None,
            format=({'kind':'number',**c['format']} if c['format'] else None)) for c in display['columns']]
        system=display['systemColumn']
        system_id=system.get('id','system-row')
        if not any(c['id']==system_id for c in columns):
            columns.append(dict(id=system_id,name=system['name'],visible=False,format=None,
                columnClass=system['columnClass'],type=system['type'],systemKind=system['systemKind'],
                elementId=system['elementId'],formula=system.get('formula')))
        execution={'request':{'workbookId':'doc','documentVersion':base,'code':display['program']},
            'result':{'structuredContent':{'type':'success','persisted':True,'workbookId':'doc',
                'documentVersion':base+1,'output':repr({'source':'sdk-display-execution',
                    'contractHash':display['contractHash'],'elementId':display['target']['elementId']})}}}
        observation={'request':{'workbookId':'doc','documentVersion':base+1,'code':display['readbackProgram']},
            'result':{'structuredContent':{'type':'success-no-edits','persisted':False,'output':repr({
                'source':'sdk-column-readback','contractHash':display['contractHash'],
                'elementId':display['target']['elementId'],'columns':columns})}}}
        current={'request':{'workbookId':'doc'},'result':{'workbookId':'doc','documentVersion':base+1}}
        return execution,observation,current

    def test_durable_public_checkpoint_consumes_raw_capture_once_and_rejects_stale_foreign_evidence(self):
        from input_display import build_input_display, checkpoint_input_display, verify_input_display
        source,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        execution,observation,current=self.captures(display)
        state=checkpoint_input_display(display,execution,observation,current)
        result=verify_input_display(display,state,current)
        self.assertEqual(result['supportedDisplayStatus'],'PASSED')
        self.assertEqual(result['status'],'PARTIAL')
        with self.assertRaisesRegex(ValueError,'consumed|pending'):
            verify_input_display(display,state,current)
        with self.assertRaisesRegex(ValueError,'already|replay'):
            checkpoint_input_display(display,execution,observation,current,state)
        with self.assertRaises(ValueError):verify_input_display(display,{'readbackId':'invented'},current)
        for capture,route in [(execution,('request','workbookId')),(execution,('request','code')),
                              (observation,('request','documentVersion')),(current,('result','documentVersion'))]:
            bad=copy.deepcopy(capture);bad[route[0]][route[1]]='foreign'
            args=[display,execution,observation,current]
            args[1 if capture is execution else 2 if capture is observation else 3]=bad
            with self.assertRaises(ValueError):checkpoint_input_display(*args)
        invalid=copy.deepcopy(execution);invalid['request']['documentVersion']=8.0
        with self.assertRaises(ValueError):checkpoint_input_display(display,invalid,observation,current)
        fresh=checkpoint_input_display(display,execution,observation,current)
        stale=copy.deepcopy(current);stale['result']['documentVersion']=10
        with self.assertRaisesRegex(ValueError,'current|version'):verify_input_display(display,fresh,stale)
        tampered=copy.deepcopy(fresh);tampered['attempts'][0]['observation']['result']['structuredContent']['output']='{}'
        with self.assertRaisesRegex(ValueError,'hash|capture'):verify_input_display(display,tampered,current)
        with tempfile.TemporaryDirectory() as directory:
            tmp=Path(directory)
            for name,data in [('manifest',manifest),('receipt',staging),('execution',execution),
                              ('readback',observation),('current',current)]:
                (tmp/(name+'.json')).write_text(json.dumps(data))
            base=[sys.executable,str(SCRIPTS/'presentation-inputs.py')]
            common=[str(source),'--manifest',str(tmp/'manifest.json'),'--receipt',str(tmp/'receipt.json'),
                    '--state',str(tmp/'ledger.json'),'--current-version',str(tmp/'current.json')]
            recorded=subprocess.run(base+['checkpoint-display']+common+['--execution',str(tmp/'execution.json'),
                '--readback',str(tmp/'readback.json'),'--out',str(tmp/'checkpoint.json')],capture_output=True,text=True)
            self.assertEqual(recorded.returncode,0,recorded.stderr)
            persisted=json.loads((tmp/'ledger.json').read_text())
            self.assertEqual(persisted['attempts'][0]['observation'],observation)
            verify=base+['verify-display']+common+['--out',str(tmp/'verified.json')]
            stale=copy.deepcopy(current);stale['result']['documentVersion']=10
            (tmp/'current.json').write_text(json.dumps(stale))
            rejected=subprocess.run(verify,capture_output=True,text=True)
            self.assertNotEqual(rejected.returncode,0)
            self.assertEqual(json.loads((tmp/'ledger.json').read_text())['attempts'][0]['status'],'PENDING')
            (tmp/'current.json').write_text(json.dumps(current))
            checked=subprocess.run(verify,capture_output=True,text=True)
            self.assertEqual(checked.returncode,0,checked.stderr)
            self.assertEqual(json.loads((tmp/'ledger.json').read_text())['attempts'][0]['status'],'CONSUMED')
            replay=subprocess.run(verify,capture_output=True,text=True)
            self.assertNotEqual(replay.returncode,0)
            rerecord=subprocess.run(base+['checkpoint-display']+common+['--execution',str(tmp/'execution.json'),
                '--readback',str(tmp/'readback.json'),'--out',str(tmp/'checkpoint.json')],capture_output=True,text=True)
            self.assertNotEqual(rerecord.returncode,0)
        for field in ['sourceHash','planHash']:
            foreign=copy.deepcopy(manifest);foreign[field]='foreign'
            with self.assertRaises(ValueError):build_input_display(plan,foreign,staging)

    def test_new_approved_execution_supersedes_stale_pending_capture_without_resetting_ledger(self):
        from input_display import build_input_display,checkpoint_input_display,verify_input_display
        from input_sync import build_artifact_records
        _,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        old=self.captures(display)
        pending=checkpoint_input_display(display,*old)
        manifest['target']['version']=9
        for target in staging['targetBindings'].values():target['version']=9
        staging['artifacts']=build_artifact_records(manifest,
            manifest['planProjection']['sharedInputArchitecture'],staging['targetBindings'])
        later=build_input_display(plan,manifest,staging)
        fresh=self.captures(later)
        ledger=checkpoint_input_display(later,*fresh,pending)
        self.assertEqual(ledger['attempts'][0]['status'],'SUPERSEDED')
        self.assertEqual(ledger['attempts'][0]['observation'],old[1])
        self.assertEqual(verify_input_display(later,ledger,fresh[2])['supportedDisplayStatus'],'PASSED')
        with self.assertRaises(ValueError):verify_input_display(display,ledger,old[2])

    def test_verifier_checks_literal_types_class_and_actual_system_kind_and_consumes_failed_attempt(self):
        from input_display import build_input_display, checkpoint_input_display, verify_input_display
        _,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        for field,value,index in [('type','boolean',0),('columnClass','SingleSelectEditableColumn',0),
                                  ('columnClass','LiteralEditableColumn',-1),('systemKind','created_at',-1),
                                  ('type','number',-1),('visible',True,-1),('format',None,1)]:
            with self.subTest(field=field,value=value,index=index):
                execution,observation,current=self.captures(display)
                import ast
                output=ast.literal_eval(observation['result']['structuredContent']['output'])
                output['columns'][index][field]=value
                observation['result']['structuredContent']['output']=repr(output)
                state=checkpoint_input_display(display,execution,observation,current)
                self.assertEqual(verify_input_display(display,state,current)['status'],'FAILED')
                with self.assertRaises(ValueError):verify_input_display(display,state,current)

    def test_readonly_currency_patterns_contain_actual_currency_marker(self):
        _,plan,_,_=self.prepared('research')
        artifact=presentation.emit_presentation(plan,presentation.fixture_bindings(plan),logical_fixture=True)
        main=next(e for e in artifact['elements'] if e['id']=='research-main-view')
        by_name={c['name']:c for c in main['columns']}
        self.assertEqual(by_name['Revenue']['format']['formatString'],'($,.1f')
        self.assertEqual(by_name['EPS']['format']['formatString'],'($.2f')
        self.assertTrue(artifact['formatLimitations'])

    def sdk_adapter(self,display,system=None):
        state={c['id']:{'name':c['name'],'type':c['type'],'visible':True,'format':None} for c in display['columns']}
        system=system or display['systemColumn']
        system_id=system.get('id','system-row')
        state[system_id]={**system,'visible':True,'format':None}
        class Format:
            def __init__(self,data):self.data=data
            def kind(self):return self.data.get('kind')
            def decimal_places(self):return self.data.get('decimals')
            def currency(self):return self.data.get('currency')
            def pattern(self):return self.data.get('pattern')
        class Column:
            def __init__(self,column_id):self.column_id=column_id
            def id(self):return self.column_id
            def name(self):return state[self.id()]['name']
            def element_id(self):return state[self.id()].get('elementId',display['target']['elementId'])
            def is_visible(self):return state[self.id()]['visible']
            def format(self):return Format(state[self.id()]['format']) if state[self.id()]['format'] else None
            def set_format(self,kind,decimals=None,currency=None):
                state[self.id()]['format']={'kind':kind,'decimals':decimals,'currency':currency,'pattern':None}
            def set_format_pattern(self,pattern):
                state[self.id()]['format']={'kind':'number','pattern':pattern}
        class LiteralColumn(Column):
            def type(self):return state[self.id()]['type']
        class SystemColumn(Column):
            def kind(self):return state[self.id()].get('kind','row_id')
        class SelectColumn(Column):pass
        class FormulaColumn(Column):
            def formula(self):return state[self.id()]['formula']
        class Table:
            def name(self):return display['tableName']
            def id(self):return display['target']['elementId']
            def get_column(self,column_id):return SelectColumn(column_id) if state[column_id].get('select') else LiteralColumn(column_id)
            def list_columns(self):return [LiteralColumn(i) if i!=system_id else (FormulaColumn(i) if state[i]['columnClass']=='FormulaColumn' else SystemColumn(i)) for i in state]
            def set_visible_columns(self,columns):
                selected={c.id() for c in columns}
                for i in state:state[i]['visible']=i in selected
        class Page:
            def get_element(self,element_id):
                self.assert_id=element_id
                assert element_id==display['target']['elementId']
                return Table()
        class Workbook:
            def get_page(self,page_id):
                assert page_id==display['target']['pageId']
                return Page()
        class Sigma:workbook=Workbook()
        outputs=[]
        context={'sigma':Sigma(),'Page':Page,'InputTable':Table,'LiteralEditableColumn':LiteralColumn,
                 'FormulaColumn':FormulaColumn,'InputTableIDColumn':SystemColumn,'SingleSelectEditableColumn':SelectColumn,'MultiSelectEditableColumn':SelectColumn,
                 'print':lambda value:outputs.append(value)}
        return context,state,outputs

    def test_generated_program_and_fresh_observation_contract_detect_metadata_noops(self):
        from input_display import build_input_display,checkpoint_input_display,verify_input_display
        _,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        context,state,outputs=self.sdk_adapter(display)
        # Every schema guard runs before visibility/format setters.
        first=display['columns'][0]['id']
        for changes in [{'type':'boolean'},{'select':True}]:
            original=copy.deepcopy(state[first]);state[first].update(changes)
            with self.assertRaises(RuntimeError):exec(display['program'],context)
            self.assertTrue(all(c['visible'] and c['format'] is None for c in state.values()))
            with self.assertRaises(RuntimeError):exec(display['readbackProgram'],context)
            state[first]=original
        state['system-row']['kind']='created_at'
        with self.assertRaises(RuntimeError):exec(display['program'],context)
        self.assertTrue(all(c['visible'] for c in state.values()))
        state['system-row']['kind']='row_id'
        exec(display['program'],context)
        self.assertEqual(sum(value['visible'] for value in state.values()),7)
        exec(display['readbackProgram'],context)
        observed=outputs[-1]
        self.assertEqual(observed['columns'][0]['columnClass'],'LiteralEditableColumn')
        self.assertEqual(observed['columns'][-1]['systemKind'],'row_id')
        execution,observation,current=self.captures(display)
        observation['result']['structuredContent']['output']=repr(observed)
        ledger=checkpoint_input_display(display,execution,observation,current)
        self.assertEqual(verify_input_display(display,ledger,current)['supportedDisplayStatus'],'PASSED')
        for value in state.values():value.update(visible=True,format=None)
        exec(display['readbackProgram'],context)
        observation['result']['structuredContent']['output']=repr(outputs[-1])
        ledger=checkpoint_input_display(display,execution,observation,current)
        self.assertEqual(verify_input_display(display,ledger,current)['status'],'FAILED')

    def approved_alias(self,manifest,staging,feature=None):
        from input_sync import build_artifact_records
        feature=feature or manifest['planProjection']['sharedInputArchitecture']['presentationInputFeature']
        target=staging['targetBindings'][feature]
        # Exact representation from authentic post-REST SDK captures; the proof
        # hash is a test-only reviewed identity-evidence fixture.
        target['systemColumn']={'id':'ID','name':'ID','columnClass':'FormulaColumn',
            'elementId':target['elementId'],'formula':'[New Input Table/ID]',
            'type':None,'systemKind':None,'rowIdBinding':{'kind':'row_id',
                'sourceColumn':'ID','valueType':'text','evidenceHash':'a'*64}}
        staging['artifacts']=build_artifact_records(manifest,
            manifest['planProjection']['sharedInputArchitecture'],staging['targetBindings'])
        return target['systemColumn']

    def test_public_display_binds_reviewed_normalized_system_alias_without_schema_writes(self):
        from input_display import build_input_display,checkpoint_input_display,verify_input_display
        source,plan,manifest,staging=self.prepared('forecast')
        original=build_input_display(plan,manifest,staging)
        alias=self.approved_alias(manifest,staging)
        # Merely seeing an ID-named formula column never widens the native route.
        context,state,_=self.sdk_adapter(original,alias)
        with self.assertRaises(RuntimeError):exec(original['program'],context)
        self.assertTrue(all(c['visible'] and c['format'] is None for c in state.values()))
        display=build_input_display(plan,manifest,staging)
        self.assertEqual(display['sourceHash'],original['sourceHash'])
        self.assertEqual(display['planHash'],original['planHash'])
        self.assertNotEqual(display['contractHash'],original['contractHash'])
        context,state,outputs=self.sdk_adapter(display)
        for field,value in [('formula','[Foreign Table/ID]'),('formula','[New Input Table/ID] + 1'),
                            ('elementId','foreign'),('name','Other ID')]:
            with self.subTest(field=field,value=value):
                before=copy.deepcopy(state['ID']);state['ID'][field]=value
                for program in ['program','readbackProgram']:
                    with self.assertRaises(RuntimeError):exec(display[program],context)
                self.assertTrue(all(c['visible'] and c['format'] is None for c in state.values()))
                state['ID']=before
        state['foreign-id']=state.pop('ID')
        with self.assertRaises(RuntimeError):exec(display['program'],context)
        state['ID']=state.pop('foreign-id')
        exec(display['program'],context)
        self.assertFalse(state['ID']['visible'])
        exec(display['readbackProgram'],context)
        observed=outputs[-1]
        system=observed['columns'][-1]
        self.assertEqual(system['columnClass'],'FormulaColumn')
        self.assertEqual(system['formula'],'[New Input Table/ID]')
        self.assertIsNone(system['type'])
        captures=self.captures(display,observed['columns'])
        ledger=checkpoint_input_display(display,*captures)
        self.assertEqual(verify_input_display(display,ledger,captures[2])['supportedDisplayStatus'],'PASSED')
        for field,value in [('id','foreign'),('formula','[Foreign Table/ID]'),('elementId','foreign'),
                            ('type','text'),('columnClass','LiteralEditableColumn')]:
            bad=copy.deepcopy(observed['columns']);bad[-1][field]=value
            captures=self.captures(display,bad)
            ledger=checkpoint_input_display(display,*captures)
            self.assertEqual(verify_input_display(display,ledger,captures[2])['status'],'FAILED')
        with tempfile.TemporaryDirectory() as directory:
            tmp=Path(directory)
            for name,value in [('manifest',manifest),('receipt',staging)]:
                (tmp/(name+'.json')).write_text(json.dumps(value))
            checked=subprocess.run([sys.executable,str(SCRIPTS/'presentation-inputs.py'),'display',str(source),
                '--manifest',str(tmp/'manifest.json'),'--receipt',str(tmp/'receipt.json'),
                '--out',str(tmp/'display.json')],capture_output=True,text=True)
            self.assertEqual(checked.returncode,0,checked.stderr)
            self.assertEqual(json.loads((tmp/'display.json').read_text())['systemColumn'],alias)
        for forbidden in ['set_formula','add_system_columns','add_formula_column','update_rows']:
            self.assertNotIn(forbidden,display['program'])

    def test_explicit_alias_route_executes_all_seven_approved_table_schemas(self):
        from input_display import build_input_display,checkpoint_input_display,verify_input_display
        features=[]
        for workflow in ['research','forecast','sales','inventory','expense']:
            _,plan,manifest,staging=self.prepared(workflow)
            for feature in staging['targetBindings']:
                with self.subTest(feature=feature):
                    self.approved_alias(manifest,staging,feature)
                    display=build_input_display(plan,manifest,staging,feature)
                    context,state,outputs=self.sdk_adapter(display)
                    exec(display['program'],context)
                    exec(display['readbackProgram'],context)
                    captures=self.captures(display,outputs[-1]['columns'])
                    ledger=checkpoint_input_display(display,*captures)
                    self.assertEqual(verify_input_display(display,ledger,captures[2])['supportedDisplayStatus'],'PASSED')
                    self.assertEqual(state['ID']['formula'],'[New Input Table/ID]')
                    self.assertFalse(state['ID']['visible'])
                    features.append(feature)
        self.assertEqual(len(features),7)

    def test_alias_binding_rejects_foreign_unapproved_or_incomplete_identity_descriptor(self):
        from input_display import build_input_display
        from input_sync import build_artifact_records,InputSyncError
        _,plan,manifest,staging=self.prepared('forecast')
        self.approved_alias(manifest,staging)
        feature=manifest['planProjection']['sharedInputArchitecture']['presentationInputFeature']
        unapproved=copy.deepcopy(staging)
        unapproved['targetBindings'][feature]['systemColumn']['rowIdBinding']['evidenceHash']='b'*64
        with self.assertRaises(InputSyncError):build_input_display(plan,manifest,unapproved)
        for field,value in [('elementId','foreign'),('id','foreign'),('formula','[Foreign/ID]'),
                            ('type','text'),('rowIdBinding',{})]:
            bad=copy.deepcopy(staging);bad['targetBindings'][feature]['systemColumn'][field]=value
            bad['artifacts']=build_artifact_records(manifest,
                manifest['planProjection']['sharedInputArchitecture'],bad['targetBindings'])
            with self.assertRaises(ValueError):build_input_display(plan,manifest,bad)

    def test_supported_semantics_remain_explicit_for_partial_or_unknown_source_patterns(self):
        from input_display import source_format_intent
        self.assertEqual(source_format_intent('0.0%','number')[0],{'pattern':'.1%'})
        self.assertEqual(source_format_intent('0.0%;(0.0%);-','number')[0],{'pattern':'(.1%'})
        self.assertEqual(source_format_intent('0.00E+00','number')[1],'UNRESOLVED')
        self.assertIsNone(source_format_intent('0.00E+00','number')[0])
