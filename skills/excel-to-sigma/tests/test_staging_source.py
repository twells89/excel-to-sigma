"""Real research-fixture coverage for staging versus verified completion."""
import copy
import importlib.util
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / 'scripts'), str(ROOT / 'scripts/lib')]
from migration_gate import MigrationGateError, gate_plan_source, source_sha256, source_issue_hash
from migration_manifest import ManifestError, validate_completion
from xlsx_audit import audit_workbook


def script(name):
    spec = importlib.util.spec_from_file_location(name.replace('-', '_'), ROOT / 'scripts' / name)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class TestResearchSourceStaging(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = ROOT / 'tests/fixtures/fixture-research-model.xlsx'
        cls.plan = script('infer-canonical-formulas.py').infer(cls.source, sheet='FY results')
        cls.cli = script('input-sync.py')
        cls.issues = audit_workbook(cls.source)['issues']
        assert {i['code'] for i in cls.issues} == {
            'CHARTS_PRESENT', 'DATA_VALIDATION_PRESENT', 'CONDITIONAL_FORMATTING_PRESENT'}

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.plan_path = self.folder / 'plan.json'
        self.plan_path.write_text(json.dumps(self.plan))
        self.manifest_path = self.folder / 'manifest.json'
        self.cli.main(['architecture', '--plan', str(self.plan_path), '--source', str(self.source),
                       '--connection', 'test-connection', '--folder', 'test-folder',
                       '--workflow', 'research', '--out-plan', str(self.plan_path),
                       '--out-manifest', str(self.manifest_path)])
        self.manifest = json.loads(self.manifest_path.read_text())
        architecture = json.loads(self.plan_path.read_text())['sharedInputArchitecture']
        self.approval = {
            'sourceHash': self.manifest['sourceHash'], 'planHash': self.manifest['planHash'],
            'owner': 'Fixture Owner', 'reviewer': 'Fixture Reviewer',
            'reviewItems': [f['id'] for f in self.manifest['formulaCells']],
            'mappedFeatures': architecture['mappedFeatures'],
            'targetBindings': {name: {'documentId': 'isolated-doc', 'version': 1, 'pageId': 'test-page'}
                               for name in architecture['inputs']},
            'sourceReviewApproval': {'status': 'approved_for_staging', 'issues': self.issues},
        }

    def stage(self, approval=None, source=None):
        path = self.folder / 'approval.json'
        path.write_text(json.dumps(self.approval if approval is None else approval))
        receipt = self.folder / 'receipt.json'
        self.cli.main(['stage', '--manifest', str(self.manifest_path), '--plan', str(self.plan_path),
                       '--source', str(source or self.source), '--approval', str(path),
                       '--out', str(receipt)])
        return json.loads(receipt.read_text())

    def test_actual_fixture_stages_with_explicit_pending_scope_and_emits_program(self):
        receipt = self.stage()
        self.assertEqual(receipt['status'], 'APPROVED_STAGING')
        self.assertEqual(receipt['sourceReviewStatus'], 'PENDING_STAGING')
        self.assertEqual(receipt['pendingSourceIssues'], self.issues)
        self.assertNotIn('verified', json.dumps(receipt['pendingSourceIssues']))
        output = self.folder / 'create.py'
        self.cli.main(['emit-create', '--manifest', str(self.manifest_path),
                       '--receipt', str(self.folder / 'receipt.json'),
                       '--feature', 'researchModelInputs', '--out', str(output)])
        self.assertIn('upsert_rows', output.read_text())

    def test_approval_rejects_stale_swapped_omitted_scope_or_missing_named_reviewer(self):
        for mutation in ('stale', 'swapped', 'omitted', 'reviewer', 'owner'):
            with self.subTest(mutation=mutation):
                approval = copy.deepcopy(self.approval)
                if mutation == 'stale': approval['planHash'] = 'f' * 64
                elif mutation == 'swapped': approval['sourceReviewApproval']['issues'][0]['evidence'] = ['wrong']
                elif mutation == 'omitted': approval['sourceReviewApproval']['issues'].pop()
                else: approval[mutation] = ' '
                with self.assertRaises((ValueError, MigrationGateError)):
                    self.stage(approval)

    def test_swapped_source_bytes_refused_before_staging(self):
        swapped = self.folder / 'swapped.xlsx'
        wb = load_workbook(self.source); wb['Model Inputs']['C2'] = .99; wb.save(swapped); wb.close()
        with self.assertRaisesRegex(MigrationGateError, 'SHA256'):
            self.stage(source=swapped)

    def test_blocked_real_source_refused_even_with_explicit_approval(self):
        source = self.folder / 'blocked.xlsx'
        source.write_bytes(self.source.read_bytes())
        with zipfile.ZipFile(source, "a") as archive:
            archive.writestr("xl/activeX/activeX1.xml", "<activeX/>")
        plan = json.loads(self.plan_path.read_text()); plan['sourceSha256'] = source_sha256(source)
        self.plan_path.write_text(json.dumps(plan))
        with self.assertRaisesRegex(MigrationGateError, 'blocked'):
            self.stage(source=source)

    def test_same_stage_approval_does_not_satisfy_production_source_gate(self):
        path = self.folder / 'stage-approval.json'; path.write_text(json.dumps(self.approval))
        with self.assertRaises(MigrationGateError):
            gate_plan_source(self.plan, source_arg=self.source, dispositions_path=path)
        with self.assertRaisesRegex(MigrationGateError, 'unresolved'):
            gate_plan_source(self.plan, source_arg=self.source)

    def completion_manifest(self):
        manifest = copy.deepcopy(self.manifest)
        manifest['target'] = {'documentId': 'isolated-doc', 'version': 1}
        for item in manifest['formulaCells']:
            item['targetBinding'] = {'elementId': 'calc', 'columnId': item['id'],
                                     'documentId': 'isolated-doc', 'version': 1}
            item['verification'] = {
                'status': 'passed', 'evidenceId': 'formula-' + item['id'],
                'sourceCellId': item['id'], 'targetId': item['id'],
                'sourceHash': manifest['sourceHash'], 'planHash': manifest['planHash'],
                'targetDocumentId': 'isolated-doc', 'targetVersion': 1}
        return manifest

    def test_global_completion_refuses_pending_features_after_formula_checks_pass(self):
        with self.assertRaisesRegex(ManifestError, 'source issue'):
            validate_completion(self.completion_manifest())

    def test_source_issue_completion_requires_exact_scope_and_bound_runtime_evidence(self):
        manifest = self.completion_manifest()
        manifest['sourceReviewEvidence'] = [
            {'code': issue['code'], 'issueHash': source_issue_hash(issue), 'status': 'passed',
             'evidenceId': 'runtime-' + issue['code'], 'targetId': 'actual-element',
             'targetBinding': {'documentId': 'isolated-doc', 'version': 1, 'elementId': 'actual-element'},
             'sourceHash': manifest['sourceHash'], 'planHash': manifest['planHash'],
             'targetDocumentId': 'isolated-doc', 'targetVersion': 1}
            for issue in self.issues]
        self.assertTrue(validate_completion(manifest))
        for field in ('sourceHash', 'planHash', 'targetVersion', 'targetId', 'issueHash', 'status', 'evidenceId'):
            with self.subTest(field=field):
                changed = copy.deepcopy(manifest)
                changed['sourceReviewEvidence'][0][field] = '' if field == 'evidenceId' else 'stale'
                with self.assertRaisesRegex(ManifestError, 'source issue'):
                    validate_completion(changed)
        manifest['sourceReviewEvidence'].pop()
        with self.assertRaisesRegex(ManifestError, 'source issue'):
            validate_completion(manifest)

    def test_stage_still_refuses_cached_or_unresolved_formula_mappings(self):
        for status in ('cached', 'unresolved'):
            manifest = copy.deepcopy(self.manifest)
            manifest['formulaCells'][0]['mappingStatus'] = status
            # Rebuild a valid manifest hash so rejection tests the formula gate.
            from migration_manifest import _plan_hash
            manifest['planHash'] = _plan_hash(manifest['sourceHash'], manifest['planProjection'],
                                              manifest['formulaCells'])
            from migration_manifest import validate_staging_approval
            approval = {**self.approval, 'planHash': manifest['planHash']}
            with self.assertRaisesRegex(ManifestError, 'unresolved formula'):
                validate_staging_approval(manifest, approval)

    def test_stage_refuses_plan_changed_without_rebuilding_approval(self):
        plan = json.loads(self.plan_path.read_text())
        plan['sharedInputArchitecture']['inputs']['researchModelInputs']['editMode'] = 'published-explore'
        self.plan_path.write_text(json.dumps(plan))
        with self.assertRaisesRegex(ValueError, 'planHash'):
            self.stage()

