"""Two-phase input staging retains formula obligations and cannot write formulas."""
import copy
import sys
import tempfile
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))
from input_orchestration import gate_input_only_staging
from input_sync import emit_construction_program, emit_write_programs, plan_seed
from migration_manifest import build_manifest, validate_staging_approval
from native_calculations import _observed_bindings, build_native_calculation_backup
from native_exact_plan import build_exact_native_calculation_plan, prepare_exact_native_scaffold
import test_native_calculations


class InputOnlyStagingTests(unittest.TestCase):
    source = test_native_calculations.NativeCalculationTests.source

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def bundle(self):
        source, inputs = self.source()
        plan = {'sourceSha256': inputs['sourceSha256'], 'sourcePath': str(source),
                'inputPlanHash': inputs['inputPlanHash'],
                'sharedInputArchitecture': inputs['sharedInputArchitecture'],
                'executionScope': 'input-tables-only'}
        manifest = build_manifest(source, plan)
        manifest['target'] = {'documentId': 'observed-doc', 'version': 7}
        binding = {'documentId': 'observed-doc', 'version': 7, 'pageId': 'observed-page'}
        approval = {'status':'approved_for_staging', 'scope': 'input-tables-only', 'owner': 'root owner', 'reviewer': 'root reviewer',
                    'sourceHash': manifest['sourceHash'], 'planHash': manifest['planHash'],
                    'reviewItems': [r['id'] for r in manifest['formulaCells']],
                    'pendingFormulaIds': [r['id'] for r in manifest['formulaCells']],
                    'mappedFeatures': ['ledger'], 'targetBindings': {'ledger': binding}}
        gate = {'status': 'AUTO', 'reviewStatus': 'NOT_REQUIRED',
                'sourceSha256': manifest['sourceHash'], 'issues': []}
        return source, inputs, plan, manifest, approval, gate

    def test_inputs_seed_from_approved_literals_while_every_formula_stays_pending(self):
        source, inputs, plan, manifest, approval, gate = self.bundle()
        before = copy.deepcopy(manifest)
        receipt = gate_input_only_staging(manifest, approval, inputs, gate, source_path=source)
        self.assertEqual(receipt['status'], 'APPROVED_INPUT_STAGING')
        self.assertEqual(set(receipt['pendingFormulaIds']), {r['id'] for r in manifest['formulaCells']})
        self.assertEqual(before, manifest)
        self.assertTrue(any(r['mappingStatus'] == 'unresolved' for r in manifest['formulaCells']))
        program = emit_construction_program(manifest, receipt, 'ledger')
        self.assertIn('add_editable_column', program)
        self.assertIn('upsert_rows', program)
        self.assertNotIn('Result', program)
        with self.assertRaisesRegex(ValueError, 'unresolved'):
            validate_staging_approval(manifest, approval)
        binding=approval['targetBindings']['ledger']
        binding.update(elementId='observed-input',columnIds={'Account':'ca','Amount':'cb','Rate':'cc'})
        assigned=gate_input_only_staging(manifest,approval,inputs,gate,source_path=source)
        spec=inputs['sharedInputArchitecture']['inputs']['ledger']
        seed=plan_seed(spec['rows'],[],spec['keyColumns'],'Row ID',
            snapshot={'id':'observed-query','ageSeconds':0},column_types={c['name']:c['type']for c in spec['columns']})
        self.assertEqual(len(emit_write_programs(seed,manifest,assigned,'ledger')),1)

    def test_input_only_receipt_cannot_authorize_native_formula_scaffold(self):
        source, inputs, plan, manifest, approval, gate = self.bundle()
        receipt = gate_input_only_staging(manifest, approval, inputs, gate, source_path=source)
        native = build_exact_native_calculation_plan(source, inputs)
        current = {'workbookId': 'observed-doc', 'documentVersion': 7,
                   'contents': {'kind': 'workbook', 'elements': [], 'pages': []}}
        rows = {'ledger': []}
        with self.assertRaises(ValueError):
            prepare_exact_native_scaffold(native, inputs, manifest, receipt, current, rows,
                backup=build_native_calculation_backup(current, rows), seed_binding={}, page_id='observed-page')
        with self.assertRaisesRegex(ValueError, 'APPROVED_STAGING'):
            _observed_bindings(inputs, manifest, receipt, current['contents'], rows)

    def test_no_scope_hash_source_formula_or_named_review_can_be_omitted(self):
        for field in ('status', 'scope', 'owner', 'reviewer', 'pendingFormulaIds', 'reviewItems', 'mappedFeatures'):
            source, inputs, plan, manifest, approval, gate = self.bundle()
            approval.pop(field)
            with self.subTest(field=field), self.assertRaises(ValueError):
                gate_input_only_staging(manifest, approval, inputs, gate, source_path=source)
        for change in ('plan', 'source', 'input', 'pending', 'blocked', 'foreign-binding'):
            source, inputs, plan, manifest, approval, gate = self.bundle()
            if change == 'plan': approval['planHash'] = '0' * 64
            if change == 'source': gate['sourceSha256'] = '0' * 64
            if change == 'input': inputs['inputPlanHash'] = '0' * 64
            if change == 'pending': approval['pendingFormulaIds'].pop()
            if change == 'blocked': gate['status'] = 'BLOCKED'
            if change == 'foreign-binding': approval['targetBindings']['ledger']['documentId'] = 'foreign'
            with self.subTest(change=change), self.assertRaises(ValueError):
                gate_input_only_staging(manifest, approval, inputs, gate, source_path=source)

    def test_input_only_is_not_a_whole_migration_or_native_review_alias(self):
        source, inputs, plan, manifest, approval, gate = self.bundle()
        plan['executionScope'] = 'native-formulas'
        other = build_manifest(source, plan); other['target'] = manifest['target']
        approval['planHash'] = other['planHash']
        with self.assertRaisesRegex(ValueError, 'input-tables-only'):
            gate_input_only_staging(other, approval, inputs, gate, source_path=source)
        plan['executionScope'] = 'input-tables-only'; plan['nativeReviewApproval'] = {'status':'pending'}
        other=build_manifest(source,plan);other['target']=manifest['target'];approval['planHash']=other['planHash']
        with self.assertRaisesRegex(ValueError, 'input-tables-only'):
            gate_input_only_staging(other, approval, inputs, gate, source_path=source)


if __name__ == '__main__': unittest.main()
