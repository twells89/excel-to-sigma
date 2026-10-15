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

    def test_guards_and_sdk_readback_verifier_distinguish_dropped_rest_metadata(self):
        from input_display import build_input_display,verify_input_display
        source,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        observed=[]
        for c in display['columns']:
            fmt=c['format']
            actual=({'kind':'number','pattern':fmt['pattern'],'decimals':None,'currency':None}
                    if 'pattern' in fmt else {**fmt,'pattern':None,'currency':None})
            observed.append({'id':c['id'],'name':c['name'],'visible':c['visible'],'format':actual})
        observed.append({'id':'system-row','name':'Row ID','visible':False,'format':None})
        readback={'source':'sdk-column-readback','readbackId':'fresh-sdk-call','documentId':'doc',
                  'documentVersion':9,'elementId':display['target']['elementId'],
                  'contractHash':display['contractHash'],'columns':observed}
        good=verify_input_display(display,readback)
        self.assertEqual(good['supportedDisplayStatus'],'PASSED')
        self.assertEqual(good['status'],'PARTIAL')
        with tempfile.TemporaryDirectory() as directory:
            tmp=Path(directory)
            for name,data in [('manifest',manifest),('receipt',staging),('readback',readback)]:
                (tmp/(name+'.json')).write_text(json.dumps(data))
            checked=subprocess.run([sys.executable,str(SCRIPTS/'presentation-inputs.py'),'verify-display',str(source),
                '--manifest',str(tmp/'manifest.json'),'--receipt',str(tmp/'receipt.json'),
                '--readback',str(tmp/'readback.json'),'--out',str(tmp/'verified.json')],capture_output=True,text=True)
            self.assertEqual(checked.returncode,0,checked.stderr)
            self.assertEqual(json.loads((tmp/'verified.json').read_text())['status'],'PARTIAL')
            bad=copy.deepcopy(readback)
            bad['columns'][0]['visible']=False
            (tmp/'readback.json').write_text(json.dumps(bad))
            rejected=subprocess.run([sys.executable,str(SCRIPTS/'presentation-inputs.py'),'verify-display',str(source),
                '--manifest',str(tmp/'manifest.json'),'--receipt',str(tmp/'receipt.json'),
                '--readback',str(tmp/'readback.json'),'--out',str(tmp/'verified.json')],capture_output=True,text=True)
            self.assertNotEqual(rejected.returncode,0)
            self.assertEqual(json.loads((tmp/'verified.json').read_text())['status'],'FAILED')


        missing=copy.deepcopy(readback);missing['columns'].pop()
        self.assertEqual(verify_input_display(display,missing)['status'],'FAILED')
        dropped=copy.deepcopy(readback)
        for c in dropped['columns']: c.update(visible=True,format=None)
        failed=verify_input_display(display,dropped)
        self.assertEqual(failed['status'],'FAILED')
        self.assertTrue(failed['mismatches'])
        foreign=copy.deepcopy(readback);foreign['source']='rest-code-rep'
        with self.assertRaisesRegex(ValueError,'SDK|sdk'):
            verify_input_display(display,foreign)
        for field,value in [('documentId','foreign'),('contractHash','foreign'),('elementId','foreign')]:
            foreign=copy.deepcopy(readback);foreign[field]=value
            with self.assertRaises(ValueError):verify_input_display(display,foreign)
        for field,value in [('sourceHash','foreign'),('planHash','foreign')]:
            foreign=copy.deepcopy(manifest);foreign[field]=value
            with self.assertRaises(ValueError):build_input_display(plan,foreign,staging)
        foreign=copy.deepcopy(manifest);foreign['target']['version']=7
        with self.assertRaises(ValueError):build_input_display(plan,foreign,staging)

    def test_readonly_currency_patterns_contain_actual_currency_marker(self):
        _,plan,_,_=self.prepared('research')
        artifact=presentation.emit_presentation(plan,presentation.fixture_bindings(plan),logical_fixture=True)
        main=next(e for e in artifact['elements'] if e['id']=='research-main-view')
        by_name={c['name']:c for c in main['columns']}
        self.assertEqual(by_name['Revenue']['format']['formatString'],'($,.1f')
        self.assertEqual(by_name['EPS']['format']['formatString'],'($.2f')
        self.assertTrue(artifact['formatLimitations'])

    def test_generated_program_and_fresh_observation_contract_detect_metadata_noops(self):
        from input_display import build_input_display,verify_input_display
        _,plan,manifest,staging=self.prepared('research')
        display=build_input_display(plan,manifest,staging)
        state={c['id']:{'name':c['name'],'visible':True,'format':None} for c in display['columns']}
        state['system-row']={'name':'Row ID','visible':True,'format':None}
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
            def is_visible(self):return state[self.id()]['visible']
            def format(self):return Format(state[self.id()]['format']) if state[self.id()]['format'] else None
            def set_format(self,kind,decimals=None,currency=None):
                state[self.id()]['format']={'kind':kind,'decimals':decimals,'currency':currency,'pattern':None}
            def set_format_pattern(self,pattern):
                state[self.id()]['format']={'kind':'number','pattern':pattern}
        class SystemColumn(Column):pass
        class Table:
            def name(self):return display['tableName']
            def id(self):return display['target']['elementId']
            def get_column(self,column_id):return Column(column_id)
            def list_columns(self):return [Column(i) if i!='system-row' else SystemColumn(i) for i in state]
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
        context={'sigma':Sigma(),'Page':Page,'InputTable':Table,'LiteralEditableColumn':Column,
                 'InputTableIDColumn':SystemColumn,'SingleSelectEditableColumn':Column,'MultiSelectEditableColumn':Column,
                 'print':lambda value:outputs.append(value)}
        exec(display['program'],context)
        self.assertEqual(sum(value['visible'] for value in state.values()),7)
        # A separate fetch-only execution must observe the resulting backing state.
        exec(display['readbackProgram'],context)
        observed=outputs[-1]
        observed.update(documentId='doc',documentVersion=9,readbackId='fresh-sdk-session')
        self.assertEqual(verify_input_display(display,observed)['supportedDisplayStatus'],'PASSED')
        for value in state.values():value.update(visible=True,format=None)
        exec(display['readbackProgram'],context)
        dropped=outputs[-1]
        dropped.update(documentId='doc',documentVersion=9,readbackId='different-sdk-session')
        self.assertEqual(verify_input_display(display,dropped)['status'],'FAILED')

    def test_supported_semantics_remain_explicit_for_partial_or_unknown_source_patterns(self):
        from input_display import source_format_intent
        self.assertEqual(source_format_intent('0.0%','number')[0],{'pattern':'.1%'})
        self.assertEqual(source_format_intent('0.0%;(0.0%);-','number')[0],{'pattern':'(.1%'})
        self.assertEqual(source_format_intent('0.00E+00','number')[1],'UNRESOLVED')
        self.assertIsNone(source_format_intent('0.00E+00','number')[0])
