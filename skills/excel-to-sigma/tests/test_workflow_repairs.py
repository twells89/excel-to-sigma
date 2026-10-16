"""Offline construction and independent source-data proofs, not UI acceptance."""
import importlib.util
import sys
import unittest
from decimal import Decimal
from openpyxl import load_workbook
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts/lib'))

class WorkflowRepairsTests(unittest.TestCase):
    def module(self):
        self.assertIsNotNone(importlib.util.find_spec('workflow_repairs'),'workflow repair adapter is missing')
        import workflow_repairs
        return workflow_repairs
    def fixture(self,workflow):
        return ROOT/'tests/fixtures'/('fixture-'+{'forecast':'forecast-model','sales':'sales-pipeline','inventory':'inventory-reorder','expense':'expense-approval','research':'research-model'}[workflow]+'.xlsx')
    def contract(self,workflow):return self.module().extract_workflow_repairs(self.fixture(workflow))
    def test_workflow_completion_cannot_pass_without_actual_mandatory_runtime(self):
        p=self.module();contract=self.contract('expense')
        with self.assertRaisesRegex(ValueError,'actual workflow acceptance'):
            p.validate_workflow_completion({'sourceHash':contract['source']['sha256'],'planHash':'b'*64,'planProjection':{'workflowRepairArchitecture':{'contract':contract}},'target':{'documentId':'stage','version':7}})

    def test_all_formulas_accounted_without_cached_values(self):
        for workflow,count in [('forecast',278),('sales',170),('inventory',442),('expense',632),('research',343)]:
            with self.subTest(workflow=workflow):
                result=self.contract(workflow)
                self.assertEqual(len(result['formulaCells']),count)
                self.assertEqual(len({r['id'] for r in result['formulaCells']}),count)
                self.assertFalse(any('cachedValue' in r for r in result['formulaCells']))
                self.assertTrue(all(r['originalFormula'].startswith('=') and r['sourceHash']==result['source']['sha256'] for r in result['formulaCells']))
                self.assertEqual(result['productionStatus'],'BLOCKED')
    def test_emitted_cent_comparison_uses_native_round_and_preserves_priority(self):
        from presentation import extract_presentation,build_business_input_architecture
        from input_sync import build_operational_input_elements
        from native_graph_oracle import project
        arch=build_business_input_architecture(extract_presentation(self.fixture('expense')),connection_id='conn',folder_id='folder')
        element=next(e for e in build_operational_input_elements(arch)['elements'] if e['id']=='expensesWithPolicy')
        base={'Policy/Category':'Meals','Policy/Per-item Limit':200,'Policy/Receipt Threshold':25,
            'ExpenseJoin/Amount':100,'ExpenseJoin/Ledger Amount':100,'ExpenseJoin/Receipt Status':'Attached','ExpenseJoin/Approval Status':'Approved'}
        for delta,want in [(.01,'CLEAR'),(-.01,'CLEAR'),(.02,'RECON'),(-.02,'RECON')]:
            got=project(element['columns'],{**base,'ExpenseJoin/Card Amount':100+delta})
            self.assertEqual(got['Exception'],want)
        got=project(element['columns'],{**base,'ExpenseJoin/Amount':300,'ExpenseJoin/Card Amount':100.02})
        self.assertEqual(got['Exception'],'RECON')
        self.assertEqual(arch['inputs']['expensePolicy']['editMode'],'published-explore')

    def test_secondary_literals_and_formula_provenance_exact(self):
        sales=self.contract('sales');spec=sales['additionalInputs']['salesRepTargets']
        self.assertEqual([r['Annual Target'] for r in spec['rows']],[1250000,1400000,1550000,1700000])
        self.assertEqual(spec['editMode'],'published-explore');self.assertEqual(spec['keyColumns'],['Rep'])
        inventory=self.contract('inventory');self.assertEqual(len(inventory['referenceTables']['Suppliers']['rows']),4)
        self.assertEqual(len(inventory['secondaryRows']),36);self.assertEqual(inventory['secondaryRows'][0]['Buyer Review'],'No action')
        expense=self.contract('expense');self.assertEqual(len(expense['secondaryRows']),48)
        self.assertIn('Match card and ledger',next(r for r in expense['formulaCells'] if r['id']=='Reconciliation!I5')['originalFormula'])
    def test_forecast_view_has_live_months_derived_rows_fy_and_source_order(self):
        p=self.module();contract=self.contract('forecast')
        fragment=p.build_workflow_repair_fragments(contract,{'calculation':{'elementId':'pnl','name':'PnL Combined'}},connection_id='conn')
        wide=next(e for e in fragment['elements'] if e['id']=='forecast-monthly-pnl')
        self.assertEqual([c['name'] for c in wide['columns'] if c['name'].endswith('-26')],['Jan-26','Feb-26','Mar-26','Apr-26','May-26','Jun-26','Jul-26','Aug-26','Sep-26','Oct-26','Nov-26','Dec-26'])
        self.assertIn('FY Total',[c['name'] for c in wide['columns']]);self.assertIn('Net Income',[r['Line Item'] for r in contract['displayRows']])
        self.assertEqual([r['Source Row'] for r in contract['displayRows']],[8,9,10,11,12,13,14,15,16,17,18,19,22,23,24,25,27,28])
        self.assertTrue(all('SumIf(' in c['formula'] for c in wide['columns'] if c['name'].endswith('-26')))
        self.assertFalse(any('Forecast Inputs View/' in c['formula'] for c in wide['columns']))
    def test_reconciliation_cent_boundary_and_precedence(self):
        from input_sync import evaluate_expense_exception
        base={'Amount':100,'Card Amount':100,'Ledger Amount':100,'Receipt Status':'Attached','Approval Status':'Approved'}
        for delta,want in [(0.01,'CLEAR'),(-.01,'CLEAR'),(.02,'RECON'),(-.02,'RECON')]:
            self.assertEqual(evaluate_expense_exception({**base,'Card Amount':100+delta},200)['exception'],want)
        self.assertEqual(evaluate_expense_exception({**base,'Amount':300,'Receipt Status':'Missing'},200)['exception'],'OVER LIMIT')
        p=self.module();self.assertFalse(p.reconciliation_required(Decimal('100.01'),Decimal('100.00')))
        self.assertTrue(p.reconciliation_required(Decimal('100.02'),Decimal('100.00')))

class AuditAndPresentationRepairTests(unittest.TestCase):
    module=WorkflowRepairsTests.module
    fixture=WorkflowRepairsTests.fixture
    contract=WorkflowRepairsTests.contract
    def test_unnamed_owned_expense_markers_require_exact_source_bound_direct_references(self):
        import copy
        p=self.module();contract=self.contract('expense');markers=('Last Approval Operation ID','Applied Approval Operation IDs')
        source={'id':'expenseRegister','kind':'table','name':'Expense Register','columns':[],
            'description':{'text':'excel-to-sigma:native:expense:'+contract['source']['sha256']}}
        node={'id':'expense-main-view','kind':'table','source':{'kind':'table','elementId':source['id']},
            'columns':[{'id':cid,'formula':formula} for cid,formula in [('amount','[Expense Register/Amount]'),
                ('9a02f594f2312fb55654bd326f5be5ef','[Expense Register/'+markers[0]+']'),
                ('4f3851f66779254208edd88e0f0e786d','[Expense Register/'+markers[1]+']'),
                ('approved','[Expense Register/ApprovedAt]'),('foreign','[Foreign Register/'+markers[0]+']'),
                ('concat','Concat([Expense Register/'+markers[0]+'], " note")'),('comment','[Expense Register/'+markers[1]+'] /* comment */')]]}
        binding={'element':source,'elementHash':p._digest(source),'viewElementHash':p._digest(node)}
        result=p.propagate_readonly_presentation(contract,node,{'amount':'Expense Register!G2'},technical_source=binding)
        for i,col in enumerate(result['columns']):
            expected=copy.deepcopy(node['columns'][i])
            if i in (1,2):expected['hidden']=True
            elif i==0:expected['format']=contract['semanticFormats']['Expense Register!G2']['emission']
            self.assertEqual(col,expected)
        self.assertFalse(any('hidden' in c for c in node['columns']))
        unbound=p.propagate_readonly_presentation(contract,node,{})
        self.assertFalse(any('hidden' in c for c in unbound['columns']))
        named=copy.deepcopy(node);named['columns'][1]['name']=markers[0]
        self.assertTrue(p.propagate_readonly_presentation(contract,named,{})['columns'][1]['hidden'])
        conflict=copy.deepcopy(node);conflict['columns'][1]['name']='Business narrative'
        with self.assertRaisesRegex(ValueError,'conflicting'):p.propagate_readonly_presentation(contract,conflict,{},technical_source={**binding,'viewElementHash':p._digest(conflict)})
        for patch in ({'elementHash':'stale'},{'viewElementHash':'stale'},{'element':{**source,'id':'foreign'}}):
            with self.subTest(patch=patch),self.assertRaises(ValueError):p.propagate_readonly_presentation(contract,node,{},technical_source={**binding,**patch})
    def test_normalized_research_audit_reads_every_live_input_and_physical_id(self):
        p=self.module();self.assertTrue(hasattr(p,'build_normalized_input_audit'),'normalized audit builder missing')
        spec={'name':'Research Inputs','columns':[{'name':'Scope','type':'text'},
            {'name':'i_model_inputs_c2','type':'number','provenance':'Model Inputs!C2'},
            {'name':'i_fy_results_b8','type':'number','provenance':'FY results!B8'}]}
        target={'elementId':'wide','columnIds':{'Scope':'scope','i_model_inputs_c2':'growth','i_fy_results_b8':'rev'},
            'systemColumnId':'ID','systemEvidenceHash':'a'*64}
        result=p.build_normalized_input_audit(self.contract('research'),spec,target,connection_id='conn')
        audit=next(e for e in result['elements'] if e['id']=='research-normalized-input-audit')
        raw=next(c for c in audit['columns'] if c['name']=='RawValue')
        self.assertIn('[Detail/i_model_inputs_c2]',raw['formula']);self.assertIn('[Detail/i_fy_results_b8]',raw['formula'])
        self.assertFalse(any('defaultValue' in r for r in result['metadataRows']))
        self.assertEqual([r['SourceCell'] for r in result['metadataRows']],['Model Inputs!C2','FY results!B8'])
        self.assertEqual(next(c for c in audit['columns'] if c['name']=='StoredRowID')['formula'],'[Detail/ID]')
    def test_normalized_research_145_values_follow_live_edits_and_keep_real_id(self):
        from test_input_display import TestInputDisplay
        from native_graph_oracle import project
        helper=TestInputDisplay();_,_,manifest,receipt=helper.prepared('research')
        spec=manifest['planProjection']['sharedInputArchitecture']['inputs']['researchModelInputs']
        target={**receipt['targetBindings']['researchModelInputs'],'systemColumnId':'ID','systemEvidenceHash':'a'*64}
        result=self.module().build_normalized_input_audit(self.contract('research'),spec,target,connection_id='conn')
        self.assertEqual(len(result['metadataRows']),145)
        equity_label=next(r['Label'] for r in result['metadataRows'] if r['SourceCell']=='FY results!M56')
        self.assertIn('2030',equity_label);self.assertNotIn('520',equity_label)
        audit=next(e for e in result['elements'] if e['id']=='research-normalized-input-audit')
        values=dict(spec['rows'][0]);values['i_model_inputs_c2']=.1;values['i_fy_results_m56']=999
        for meta in result['metadataRows']:
            joined={**{'Base/'+k:v for k,v in meta.items()},**{'Detail/'+k:v for k,v in values.items()},'Detail/ID':'stored-row'}
            actual=project(audit['columns'],joined)
            self.assertEqual(actual['RawValue'],values[meta['LogicalInputId']]);self.assertEqual(actual['StoredRowID'],'stored-row')

    def test_readonly_formats_apply_to_actual_view_columns_and_cf_priority(self):
        p=self.module();self.assertTrue(hasattr(p,'propagate_readonly_presentation'),'readonly format propagation missing')
        contract=self.contract('sales')
        node={'id':'view','kind':'table','columns':[{'id':'prob','name':'Probability','formula':'[Pipeline/Probability]'},
            {'id':'acv','name':'ACV','formula':'[Pipeline/ACV]'}]}
        result=p.propagate_readonly_presentation(contract,node,{'prob':'Pipeline!I2','acv':'Pipeline!J2'})
        self.assertEqual(next(c for c in result['columns'] if c['id']=='prob')['format']['formatString'],'.0%')
        self.assertEqual(next(c for c in result['columns'] if c['id']=='acv')['format']['formatString'],'$,.0f')
        self.assertNotIn('format',node['columns'][0])
    def test_research_one_row_input_height_shrink_preserves_unrelated_layout_bytes(self):
        p=self.module();node={'id':'input','kind':'input-table','name':'Research Inputs'}
        layout='<Page id="edit"><Element elementId="input" gridColumn="1 / 25" gridRow="4 / 28" /><Element elementId="other" gridRow="30 / 34" unknown="keep" /></Page>'
        contents={'elements':[node,{'id':'other','kind':'text','body':'keep'}],'pages':[{'id':'edit'}],'layout':layout}
        patch={'input':{'pageId':'edit','elementHash':p._digest(node),'fromGridRow':'4 / 28','toGridRow':'4 / 18'}}
        result=p.shrink_input_layout(contents,patch)
        self.assertEqual(result['layout'],layout.replace('4 / 28','4 / 18'))
        self.assertEqual(result['elements'],contents['elements'])
        with self.assertRaises(ValueError):p.shrink_input_layout(contents,{'input':{**patch['input'],'toGridRow':'1 / 99'}})

    def test_owned_scaffold_cleanup_preserves_foreign_and_refuses_unknown_placements(self):
        p=self.module();self.assertTrue(hasattr(p,'cleanup_owned_scaffolds'),'scaffold cleanup missing')
        current={'pages':[{'id':'temp','name':'Host'},{'id':'foreign','name':'Foreign'},{'id':'overview','name':'Overview'}],
            'elements':[{'id':'pending','kind':'text','body':'pending'},{'id':'foreign-element','kind':'text','body':'keep'}],
            'layout':'<Page id="temp"><Element elementId="pending" /></Page><Page id="foreign"><Element elementId="foreign-element" /></Page><Page id="overview" />'}
        review={'pages':{'temp':p._digest(current['pages'][0])},'elements':{'pending':p._digest(current['elements'][0])},'landingPageId':'overview','owner':'reviewer'}
        result=p.cleanup_owned_scaffolds(current,review)
        self.assertEqual(result['pages'][0]['id'],'overview')
        self.assertEqual(next(v for v in result['pages'] if v['id']=='temp')['visibility'],'hidden')
        self.assertEqual(next(v for v in result['elements'] if v['id']=='foreign-element'),current['elements'][1])
        bad={**current,'layout':current['layout'].replace('elementId="pending"','elementId="foreign-element"')}
        with self.assertRaises(ValueError):p.cleanup_owned_scaffolds(bad,review)

class RepairIntegrationTests(unittest.TestCase):
    module=WorkflowRepairsTests.module
    fixture=WorkflowRepairsTests.fixture
    contract=WorkflowRepairsTests.contract
    def test_sales_dashboard_uses_live_weighted_values_and_preserves_static_source_formulas(self):
        import re,json
        from native_graph_oracle import scalar
        p=self.module();c=self.contract('sales')
        result=p.build_dashboard_view(c,{'elementId':'pipeline','name':'Pipeline'},connection_id='conn')
        formula=next(col['formula'] for col in result['elements'][-1]['columns'] if col['name']=='Weighted pipeline')
        wb=load_workbook(self.fixture('sales'));rows={r[0].value:float(r[8].value)*float(r[9].value) for r in list(wb['Pipeline'])[1:]};wb.close()
        rendered=re.sub(r'Lookup\(\[Pipeline/Weighted ACV\], ("[^"]+"), \[Pipeline/Opportunity ID\]\)',lambda m:str(rows[json.loads(m.group(1))]),formula)
        self.assertAlmostEqual(scalar(rendered,{}),sum(rows.values()),places=6)
        first=next(iter(rows));rows[first]=0
        rendered=re.sub(r'Lookup\(\[Pipeline/Weighted ACV\], ("[^"]+"), \[Pipeline/Opportunity ID\]\)',lambda m:str(rows[json.loads(m.group(1))]),formula)
        self.assertAlmostEqual(scalar(rendered,{}),sum(rows.values()),places=6)
        expense=p.build_dashboard_view(self.contract('expense'),{'elementId':'expense','name':'Expense Register'},connection_id='conn')
        static=next(r for r in expense['provenance'] if r['sourceCell']=='Dashboard!B7')
        self.assertEqual(static['kind'],'source-static-expression');self.assertEqual(static['formula'],'=SUM(7,31,1,5)')
    def test_emitted_pnl_months_and_fy_match_independent_source_equations_after_edits(self):
        import copy,re,json
        from presentation import extract_presentation,build_business_input_architecture
        from input_sync import build_operational_input_elements
        from native_graph_oracle import run_forecast,project
        import fixture_regression
        p=self.module();path=self.fixture('forecast');contract=self.contract('forecast')
        plan=extract_presentation(path);arch=build_business_input_architecture(plan,connection_id='conn',folder_id='folder')
        native=build_operational_input_elements(arch)['elements']
        view=next(e for e in p.build_workflow_repair_fragments(contract,{'calculation':{'elementId':'pnlCombined','name':'PnL Combined'}},connection_id='conn')['elements'] if e['id']=='forecast-monthly-pnl')
        for edited in (False,True):
            wide=copy.deepcopy(arch['inputs']['forecastBusinessInput']['rows'])
            edits={'Forecast Inputs!C6':.05,'Forecast Inputs!D8':0} if edited else {}
            if edited:wide[0]['Monthly Growth']=.05;wide[2]['Jul-26']=0
            actual=run_forecast(native,wide);expected=fixture_regression.forecast(path,edits)['rows']
            for spine,want in zip(contract['displayRows'],expected):
                columns=copy.deepcopy(view['columns'])
                for col in columns:
                    def aggregate(m):
                        section,month=map(json.loads,m.groups())
                        return str(sum(r['Amount'] for r in actual if r['Section']==section and r['Month']==month))
                    col['formula']=re.sub(r'SumIf\(\[PnL Combined/Amount\], \[PnL Combined/Section\] = ("[^"]+") and \[PnL Combined/Month\] = ("[^"]+")\)',aggregate,col['formula'])
                    col['formula']=re.sub(r'Lookup\(\[Forecast Monthly Line Results/([^\]]+)\], \[P&L Source Rows/Line Item\], \[Forecast Monthly Line Results/Line Item\]\)',lambda m:str(sum(r['Amount'] for r in actual if r['Line Item']==spine['Line Item'] and r['Month']==m.group(1))),col['formula'])
                got=project(columns,{'P&L Source Rows/'+k:v for k,v in spine.items()})
                for period in [*contract['months'],'FY Total']:self.assertAlmostEqual(got[period],want[period],delta=.000001)

    def test_navigation_uses_current_openapi_page_target_and_hidden_toast(self):
        p=self.module();nav=p.build_navigation([{'id':'overview','name':'Overview'}])
        self.assertEqual(nav[0]['actions'][0]['effects'][0],{'effect':'navigate','target':{'type':'page','page':'overview'}})

class GuardedRepairProposalTests(unittest.TestCase):
    def actions(self,current):
        from test_promotion import PromotionTests
        return PromotionTests().actions(current)
    module=WorkflowRepairsTests.module
    fixture=WorkflowRepairsTests.fixture
    contract=WorkflowRepairsTests.contract
    def test_repair_update_binds_actual_source_graph_and_preserves_foreign_objects(self):
        p=self.module();self.assertTrue(hasattr(p,'prepare_repair_update'),'guarded repair authoring missing')
        # Exercise the reusable pure merge guard independently of receipt binding.
        current={'workbookId':'stage','documentVersion':6,'contents':{'kind':'workbook',
            'pages':[{'id':'old','name':'Old'}],'elements':[{'id':'foreign','kind':'text','body':'keep'}],
            'layout':'<Page id="old"><Element elementId="foreign" /></Page>'}}
        fragment={'pages':[{'id':'new','name':'Workflow'}],'elements':[{'id':'newtitle','kind':'text','body':'Ready'}],
            'layout':'<Page id="new"><Element elementId="newtitle" /></Page>'}
        scope={'documentId':'stage','version':6,'folderId':'synthetic','folderEvidenceId':'folder-read','connectionId':'conn'}
        proposal=p.prepare_fragment_update(current,[fragment],scope=scope,source_hash='a'*64,plan_hash='b'*64,review={'owner':'reviewer','beforeHash':p._digest(current['contents']),'sdkActionInventory':self.actions(current)})
        self.assertEqual(proposal['dryRunRequest']['contents']['elements'][0],current['contents']['elements'][0])
        dry={'httpStatus':200,'dryRun':True,'evidenceId':'actual','proposalHash':proposal['proposalHash']}
        request=p.finalize_repair_update(proposal,current,dry,fresh_sdk_actions=self.actions(current))
        self.assertEqual(request['body']['documentVersion'],6)
        with self.assertRaises(ValueError):p.finalize_repair_update(proposal,{**current,'documentVersion':7},dry)
        with self.assertRaises(ValueError):p.finalize_repair_update(proposal,current,dry,consumed_proposals=[proposal['proposalHash']])
        with self.assertRaises(ValueError):p.prepare_fragment_update(current,[{'pages':[],'elements':[current['contents']['elements'][0]],'layout':''}],scope=scope,source_hash='a'*64,plan_hash='b'*64,review={'owner':'reviewer','beforeHash':p._digest(current['contents']),'sdkActionInventory':self.actions(current)})

    def test_public_repair_adapter_consumes_exact_receipt_and_rejects_foreign_graph(self):
        import copy
        from presentation import extract_presentation,build_business_input_architecture,build_business_input_plan
        from migration_manifest import build_manifest
        from input_sync import build_artifact_records,build_operational_input_elements
        p=self.module();contract=self.contract('sales');plan=extract_presentation(self.fixture('sales'))
        arch=build_business_input_architecture(plan,connection_id='conn',folder_id='folder')
        arch['inputs'].update({f:{**copy.deepcopy(spec),'connectionId':'conn'} for f,spec in contract['additionalInputs'].items()})
        manifest=build_manifest(self.fixture('sales'),build_business_input_plan(plan,arch));manifest['target']={'documentId':'stage','version':6}
        targets={f:{'documentId':'stage','version':6,'pageId':'host','elementId':f,'columnIds':{c['name']:'c'+str(i) for i,c in enumerate(spec['columns'])}} for f,spec in arch['inputs'].items()}
        receipt={'status':'APPROVED_STAGING','sourceHash':manifest['sourceHash'],'planHash':manifest['planHash'],'targetBindings':targets}
        receipt['artifacts']=build_artifact_records(manifest,arch,targets)
        inputs=[{'id':f,'name':spec['name'],'kind':'input-table','source':{'kind':'empty','connectionId':'conn'},'columns':[{'id':cid,'name':n} for n,cid in targets[f]['columnIds'].items()]} for f,spec in arch['inputs'].items()]
        native=build_operational_input_elements(arch,targets)['elements']
        current={'workbookId':'stage','documentVersion':6,'contents':{'kind':'workbook','pages':[{'id':'host','name':'Inputs'}], 'elements':inputs+native,'layout':'<Page id="host">'+''.join('<Element elementId="'+e['id']+'" />' for e in inputs+native)+'</Page>'}}
        calc=next(e for e in native if e['name']=='Pipeline');rep=next(e for e in inputs if e['id']=='salesRepTargets')
        bindings={k:{'elementId':e['id'],'name':e['name'],'elementHash':p._digest(e)} for k,e in [('calculation',calc),('repTargets',rep)]}
        scope={'documentId':'stage','version':6,'connectionId':'conn','folderId':'folder','folderEvidenceId':'read'}
        review={'status':'APPROVED_STAGING','owner':'reviewer','planHash':manifest['planHash'],'contractHash':contract['contractHash'], 'beforeHash':p._digest(current['contents']),'sourceBindings':bindings,'sdkActionInventory':self.actions(current),
            'inputSystems':{f:{'targetBindingHash':p._digest(t),'columnId':'ID','evidenceHash':'a'*64} for f,t in targets.items()}}
        proposal=p.prepare_repair_update(manifest,receipt,contract,current,bindings,scope=scope,review=review)
        self.assertIn('sales-rep-coverage',proposal['addedElementIds'])
        readback={'workbookId':'stage','documentVersion':7,'contents':proposal['dryRunRequest']['contents']}
        tables={f:{'elementId':f,'columns':[{'id':t['columnIds'][c['name']],'name':c['name'],'type':c['type'],'columnClass':'LiteralEditableColumn'} for c in arch['inputs'][f]['columns']]+[{'id':'fresh-native-'+f,'name':'Row ID','columnClass':'InputTableIDColumn','systemKind':'row_id'}]} for f,t in targets.items()}
        sdk={'documentId':'stage','version':7,'evidenceId':'actual-sdk','tables':tables}
        post_review={'status':'APPROVED_STAGING','owner':'reviewer','contentsHash':p._digest(readback['contents']),'sdkHash':p._digest(sdk),
            'systemColumns':{f:{'observationHash':p._digest(table['columns'][-1]),'rowIdEvidenceHash':'a'*64} for f,table in tables.items()}}
        rebound=p.rebind_after_repair(manifest,receipt,proposal,readback,sdk,post_review,sdk_actions=self.actions(readback))
        self.assertTrue(all(t['version']==7 and t['rowIdColumnId'].startswith('fresh-native-') for t in rebound['receipt']['targetBindings'].values()))
        with self.assertRaises(ValueError):p.rebind_after_repair(manifest,receipt,proposal,readback,{**sdk,'version':6},post_review,sdk_actions=self.actions(readback))
        foreign_sdk=copy.deepcopy(sdk);feature='salesSettings'
        observed={'id':'ID','name':'ID','columnClass':'FormulaColumn','type':None,'systemKind':None,'formula':'[Foreign/ID]','elementId':feature}
        foreign_sdk['tables'][feature]['columns'][-1]=observed
        alias={**observed,'formula':'[New Input Table/ID]','rowIdBinding':{'kind':'row_id','sourceColumn':'ID','valueType':'text','evidenceHash':'a'*64}}
        foreign_review=copy.deepcopy(post_review);foreign_review['sdkHash']=p._digest(foreign_sdk)
        foreign_review['systemColumns'][feature]={'observationHash':p._digest(observed),'rowIdEvidenceHash':'a'*64,'approvedAlias':alias}
        with self.assertRaisesRegex(ValueError,'alias.*observation'):
            p.rebind_after_repair(manifest,receipt,proposal,readback,foreign_sdk,foreign_review,sdk_actions=self.actions(readback))


        wrong=copy.deepcopy(current);next(e for e in wrong['contents']['elements'] if e['id']=='salesBusinessInput')['source']['connectionId']='other'
        with self.assertRaises(ValueError):p.prepare_repair_update(manifest,receipt,contract,wrong,bindings,scope=scope,review={**review,'beforeHash':p._digest(wrong['contents'])})

if __name__=='__main__':unittest.main()
