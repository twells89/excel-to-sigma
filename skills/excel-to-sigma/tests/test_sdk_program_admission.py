"""Known actual SDK syntax exclusions, not a claim of live SDK typechecking."""
import ast,copy,datetime as dt,sys,types,unittest
from unittest.mock import patch
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts/lib'))
import test_promotion as rollback_fixtures
import test_operation_outcomes as outcome_fixtures
import test_input_display as display_fixtures
import test_input_sync as sync_fixtures


def unsupported(program):
    failures=[]
    for node in ast.walk(ast.parse(program)):
        if isinstance(node,(ast.Set,ast.SetComp,ast.GeneratorExp,ast.AnnAssign)):failures.append(type(node).__name__)
        if isinstance(node,ast.Call):
            if isinstance(node.func,ast.Name) and node.func.id=='set':failures.append('set()')
            if isinstance(node.func,ast.Attribute) and node.func.attr=='issubset':failures.append('issubset()')
            if isinstance(node.func,ast.Name) and node.func.id=='isinstance' and len(node.args)>1 and isinstance(node.args[1],ast.Tuple):failures.append('tuple isinstance')
    return failures


def fake_sdk(binding,*,table_name='Input'):
    """Strict effect recorder; never stands in for actual MCP admission."""
    effects=[]
    class Column:
        def __init__(self,meta):self.meta=meta
        def id(self):return self.meta['id']
        def name(self):return self.meta['name']
        def element_id(self):return self.meta.get('elementId',binding['elementId'])
        def format(self):return None
        def is_visible(self):return True
        def set_visible(self,value):effects.append(('visible',self.id(),value))
        def set_format(self,*args,**kwargs):effects.append(('format',self.id(),args,kwargs))
        def set_format_pattern(self,value):effects.append(('pattern',self.id(),value))
    class LiteralEditableColumn(Column):
        def type(self):return self.meta['type']
    class SingleSelectEditableColumn(LiteralEditableColumn):
        def allowed_values(self):return self.meta['allowedValues']
    class MultiSelectEditableColumn(SingleSelectEditableColumn):pass
    class FormulaColumn(Column):
        def formula(self):return self.meta['formula']
    class InputTableIDColumn(Column):
        def kind(self):return self.meta.get('systemKind','row_id')
    classes={c.__name__:c for c in (LiteralEditableColumn,SingleSelectEditableColumn,MultiSelectEditableColumn,FormulaColumn,InputTableIDColumn)}
    class InputTable:
        def __init__(self):self.columns=[classes[c['columnClass']](copy.deepcopy(c)) for c in binding['sdkColumns']]
        def id(self):return binding['elementId']
        def name(self):return table_name
        def list_columns(self):return self.columns
        def get_column(self,identity):return next((c for c in self.columns if c.id()==identity),None)
        def update_rows(self,values):effects.append(('update',values))
        def delete_rows(self,values):effects.append(('delete',values))
        def set_visible_columns(self,values):effects.append(('visible-columns',[c.id() for c in values]))
    table=InputTable()
    class Page:
        def get_element(self,identity):return table if identity==binding['elementId'] else None
    page=Page();sigma=types.SimpleNamespace(workbook=types.SimpleNamespace(get_page=lambda identity:page if identity==binding['pageId'] else None))
    modules={'sigma':types.ModuleType('sigma'),'sigma.page':types.ModuleType('sigma.page'),
        'sigma.inputTable':types.ModuleType('sigma.inputTable'),'sigma.column':types.ModuleType('sigma.column')}
    modules['sigma.page'].Page=Page;modules['sigma.inputTable'].InputTable=InputTable
    for name,cls in classes.items():setattr(modules['sigma.column' if name=='FormulaColumn' else 'sigma.inputTable'],name,cls)
    return sigma,table,effects,modules


class SdkProgramAdmissionTests(unittest.TestCase):
    def test_fractional_datetime_literals_roundtrip_utc_offsets_and_seed_retry_exactly(self):
        from input_sync import _sdk_literal,plan_seed
        for wire,expected in [('2026-10-15T08:45:12.123Z',dt.datetime(2026,10,15,8,45,12,123000)),
                ('2026-10-15T04:45:12.123456-04:00',dt.datetime(2026,10,15,8,45,12,123456)),
                ('2026-10-15T08:45:12.000001+00:00',dt.datetime(2026,10,15,8,45,12,1))]:
            literal=_sdk_literal(wire,'datetime');readback=eval(literal,{'datetime':dt.datetime})
            with self.subTest(wire=wire):
                self.assertEqual(readback,expected)
                source=[{'Key':'asof','At':wire}];prior=plan_seed(source,[],key_columns=['Key'],row_id_column='ID',column_types={'Key':'text','At':'datetime'})
                result=plan_seed(source,[{'ID':'stored-id','Key':'asof','At':readback}],key_columns=['Key'],row_id_column='ID',prior_seed=prior,column_types={'Key':'text','At':'datetime'})
                self.assertEqual(result['inserts'],[]);self.assertEqual(result['updates'],[]);self.assertEqual(result['conflicts'],[])
        self.assertEqual(_sdk_literal('2026-10-15','datetime'),'datetime(2026, 10, 15, 12)')

    def test_construction_and_resumable_seed_emit_exact_fractional_datetime_rows(self):
        from input_sync import build_artifact_records,emit_construction_program,emit_write_programs,plan_seed
        h=sync_fixtures.TestMcpPrograms();h.setUp();spec=h.architecture['inputs']['salesSettings']
        rows=[{'Key':'utc','At':'2026-10-15T08:45:12.123Z'},{'Key':'offset','At':'2026-10-15T04:45:12.123456-04:00'}]
        spec.update(rows=rows,keyColumns=['Key'],columns=[{'name':'Key','type':'text','provenance':'source'},{'name':'At','type':'datetime','provenance':'source'}])
        h.bindings['salesSettings'].update(elementId='assigned-settings',columnIds={'Key':'key-id','At':'at-id'})
        h.receipt['artifacts']=build_artifact_records(h.manifest,h.architecture,h.bindings)
        types={'Key':'text','At':'datetime'};seed=plan_seed(rows,[],['Key'],'ID',column_types=types)
        programs=[emit_construction_program(h.manifest,h.receipt,'salesSettings'),emit_write_programs(seed,h.manifest,h.receipt,'salesSettings')[0]['program']]
        expected=[dt.datetime(2026,10,15,8,45,12,123000),dt.datetime(2026,10,15,8,45,12,123456)]
        for program in programs:
            self.assertEqual(unsupported(program),[])
            calls=[n for n in ast.walk(ast.parse(program)) if isinstance(n,ast.Call) and isinstance(n.func,ast.Name) and n.func.id=='datetime']
            values=[eval(compile(ast.Expression(n),'emitted-date','eval'),{'datetime':dt.datetime}) for n in calls]
            self.assertCountEqual(values,expected)
        actual=[{**row,'At':value,'ID':'server-'+row['Key']} for row,value in zip(rows,expected)]
        rerun=plan_seed(rows,actual,['Key'],'ID',prior_seed=seed,column_types=types)
        self.assertEqual(rerun['conflicts'],[]);self.assertEqual(rerun['updates'],[]);self.assertEqual(len(rerun['unchanged']),2)

    def test_rollback_schema_identity_type_and_options_checks_run_before_any_effect(self):
        p,current,rows,sdk,scope=rollback_fixtures.PromotionTests().fixture()
        binding=rows['input']['binding'];binding['columnClasses']['Key']='SingleSelectEditableColumn';binding['sdkColumns'][0].update(columnClass='SingleSelectEditableColumn',allowedValues=['a','b']);sdk['input']['columns'][0].append(['a','b'])
        backup=p.build_promotion_backup(scope,current,rows,sdk,{});changed=copy.deepcopy(rows);changed['input']['rows'][0]['Value']=99
        plan=p.prepare_rollback(backup,{**current,'documentVersion':8},changed)
        for family,program in [('rows',plan['rowPrograms']['input']),('display',plan['sdkPrograms']['input'])]:
            self.assertEqual(unsupported(program),[])
            for attack in ('duplicate','missing','extra','owner','name','type','options','formula'):
                sigma,table,effects,modules=fake_sdk(binding)
                if attack=='duplicate':table.columns.append(table.columns[0])
                elif attack=='missing':table.columns.pop(0)
                elif attack=='extra':table.columns.append(type(table.columns[0])({**table.columns[0].meta,'id':'foreign'}))
                elif attack=='owner':table.columns[0].meta['elementId']='foreign-owner'
                elif attack=='name':table.columns[0].meta['name']='foreign-name'
                elif attack=='type':table.columns[1].meta['type']='text'
                elif attack=='options':table.columns[0].meta['allowedValues']=['other']
                else:table.columns[-1].meta['formula']='[Foreign/ID]'
                with self.subTest(family=family,attack=attack),patch.dict(sys.modules,modules),self.assertRaises(RuntimeError):exec(program,{'sigma':sigma})
                self.assertEqual(effects,[],(family,attack))
            sigma,table,effects,modules=fake_sdk(binding)
            with patch.dict(sys.modules,modules):exec(program,{'sigma':sigma})
            self.assertTrue(effects)

    def test_display_select_readback_is_typed_and_never_changes_options(self):
        from input_display import build_input_display,_hash
        from input_sync import build_artifact_records
        _,plan,m,r=display_fixtures.TestInputDisplay().prepared('expense');target=r['targetBindings']['expenseBusinessInput'];rule=next(v for v in plan['validations'] if v['field']=='Approval Status')
        target['columnContracts']={'Approval Status':{'id':target['columnIds']['Approval Status'],'columnClass':'SingleSelectEditableColumn',
            'type':'text','allowedValues':rule['emission']['values'],'sourceValidationHash':_hash(rule),'sdkEvidenceHash':'a'*64}}
        r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings']);display=build_input_display(plan,m,r)
        columns=[{**c,'elementId':target['elementId']} for c in display['columns']];columns.append({**display['systemColumn'],'id':'system-id'})
        binding={'elementId':target['elementId'],'pageId':target['pageId'],'sdkColumns':columns};sigma,table,effects,modules=fake_sdk(binding,table_name=display['tableName']);observed=[]
        with patch.dict(sys.modules,modules):exec(display['readbackProgram'],{'sigma':sigma,'print':lambda value:observed.append(value)})
        self.assertEqual(effects,[]);actual=next(c for c in observed[0]['columns'] if c['name']=='Approval Status')
        self.assertEqual(actual['columnClass'],'SingleSelectEditableColumn');self.assertEqual(actual['allowedValues'],rule['emission']['values']);self.assertEqual(actual['type'],'text')
        for family in ('program','readbackProgram'):
            for attack in ('duplicate','extra','owner','name','type','options','system-kind'):
                sigma,table,effects,modules=fake_sdk(binding,table_name=display['tableName'])
                selected=next(c for c in table.columns if c.name()=='Approval Status')
                if attack=='duplicate':table.columns.append(table.columns[0])
                elif attack=='extra':table.columns.append(type(selected)({**selected.meta,'id':'foreign-extra'}))
                elif attack=='owner':selected.meta['elementId']='foreign-owner'
                elif attack=='name':selected.meta['name']='foreign-name'
                elif attack=='type':selected.meta['type']='number'
                elif attack=='options':selected.meta['allowedValues']=['unapproved']
                else:table.columns[-1].meta['systemKind']='created_at'
                with self.subTest(family=family,attack=attack),patch.dict(sys.modules,modules),self.assertRaises(RuntimeError):exec(display[family],{'sigma':sigma})
                self.assertEqual(effects,[],(family,attack))
    def test_rollback_and_outcome_programs_avoid_actual_unsupported_sdk_constructs(self):
        p,current,rows,sdk,scope=rollback_fixtures.PromotionTests().fixture()
        backup=p.build_promotion_backup(scope,current,rows,sdk,{})
        changed=copy.deepcopy(rows);changed['input']['rows'][0]['Value']=99
        plan=p.prepare_rollback(backup,{**current,'documentVersion':8},changed)
        h=outcome_fixtures.OutcomeTests();outcomes,operations,expected,capture=h.fixture()
        outcome=outcomes.reconcile_operation(expected,capture,operations,completed_at='2026-10-15T12:00:00',manifest=h.manifest,receipt=h.receipt)
        for family,program in [('rollback-rows',plan['rowPrograms']['input']),('rollback-display',plan['sdkPrograms']['input']),('outcome',outcome['program'])]:
            with self.subTest(family=family):self.assertEqual(unsupported(program),[])

    def test_five_display_families_and_select_readback_avoid_actual_unsupported_syntax(self):
        from input_display import build_input_display,_hash
        from input_sync import build_artifact_records
        helper=display_fixtures.TestInputDisplay()
        for workflow in ('forecast','sales','inventory','expense','research'):
            _,plan,m,r=helper.prepared(workflow)
            if workflow=='expense':
                target=r['targetBindings']['expenseBusinessInput'];rule=next(v for v in plan['validations'] if v['field']=='Approval Status')
                target['columnContracts']={'Approval Status':{'id':target['columnIds']['Approval Status'],'columnClass':'SingleSelectEditableColumn',
                    'type':'text','allowedValues':rule['emission']['values'],'sourceValidationHash':_hash(rule),'sdkEvidenceHash':'a'*64}}
                r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings'])
            result=build_input_display(plan,m,r)
            for family in ('program','readbackProgram'):
                with self.subTest(workflow=workflow,family=family):self.assertEqual(unsupported(result[family]),[])

    def test_admission_filter_rejects_archived_failure_shapes(self):
        for code in ('a = set(c.id() for c in table.list_columns())','a = {"id"}',
                'if set(approved).issubset(set(observed)): pass','if isinstance(column, (LiteralEditableColumn, SingleSelectEditableColumn)): pass'):
            with self.subTest(code=code):self.assertTrue(unsupported(code))


if __name__=='__main__':unittest.main()
