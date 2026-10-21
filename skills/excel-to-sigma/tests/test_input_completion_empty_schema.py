"""Completion evidence for explicitly reviewed, initially empty input tables."""
from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

from input_orchestration import gate_input_completion  # noqa: E402
from input_sync import build_artifact_records, build_shared_input_architecture  # noqa: E402


@mock.patch("input_orchestration.gate_completion_manifest", return_value=True)
class TestEmptySchemaCompletion(unittest.TestCase):
    def completion_records(self, *, empty_schema=True):
        architecture = build_shared_input_architecture(
            {"sourceInputs": []}, connection_id="conn", folder_id="folder",
            workflows=["research"])
        spec = architecture["inputs"]["researchModelInputs"]
        spec.update({"rows": [], "emptySchema": empty_schema,
                     "typeReviewEvidence": "Reviewed source table header defines Scope as text"})
        target = {"documentId": "doc", "version": 3}
        bindings = {"researchModelInputs": {
            **target, "pageId": "page", "elementId": "input",
            "columnIds": {"Scope": "scope-column"}}}
        manifest = {"sourceHash": "a" * 64, "planHash": "b" * 64,
                    "target": target,
                    "planProjection": {"sharedInputArchitecture": architecture}}
        receipt = {"status": "APPROVED_STAGING", "sourceHash": manifest["sourceHash"],
                   "planHash": manifest["planHash"], "targetBindings": bindings,
                   "mappedFeatures": architecture["mappedFeatures"],
                   "artifacts": build_artifact_records(manifest, architecture, bindings)}
        evidence = {"sourceHash": manifest["sourceHash"], "planHash": manifest["planHash"],
                    "status": "passed", "target": target, "targetBindings": bindings,
                    "mappedFeatures": architecture["mappedFeatures"],
                    "tables": {"researchModelInputs": {
                        "querySnapshotId": "query", "readbackId": "readback",
                        "persistenceReadbackId": "persisted", "rows": 0, "distinct": 0,
                        "keyColumns": ["Scope"], "naturalKeys": []}},
                    "checkpointBindings": [{
                        "checkpointId": "checkpoint", "feature": "researchModelInputs",
                        "artifactHash": receipt["artifacts"]["researchModelInputs"]["artifactHash"],
                        **target, "elementId": "input"}],
                    "acceptanceChecks": [{
                        "id": check, "status": "passed", "restored": True,
                        "evidenceId": "evidence-" + check, **target,
                        "sourceHash": manifest["sourceHash"], "planHash": manifest["planHash"]}
                        for check in architecture["acceptanceChecks"]]}
        return manifest, receipt, evidence

    def test_reviewed_empty_schema_completes_with_zero_counts_and_empty_keys(self, gate):
        manifest, receipt, evidence = self.completion_records()
        result = gate_input_completion(manifest, receipt, evidence)
        self.assertEqual(result["status"], "VERIFIED_COMPLETE")
        self.assertEqual(result["tableReadbackIds"], {"researchModelInputs": "readback"})
        gate.assert_called_once_with(manifest)

    def test_empty_rows_require_explicit_boolean_empty_schema_approval(self, gate):
        for marker in (False, None, 1, "true"):
            with self.subTest(marker=marker):
                records = self.completion_records(empty_schema=marker)
                with self.assertRaisesRegex(ValueError, "approved row count"):
                    gate_input_completion(*records)
        manifest, receipt, evidence = self.completion_records()
        manifest["planProjection"]["sharedInputArchitecture"]["inputs"]["researchModelInputs"].pop("emptySchema")
        with self.assertRaisesRegex(ValueError, "approved row count"):
            gate_input_completion(manifest, receipt, evidence)

    def test_reviewed_empty_schema_rejects_nonzero_or_untyped_counts(self, gate):
        for rows, distinct in ((1, 1), (1, 0), (0, 1), (False, 0), (0, False), (0.0, 0)):
            with self.subTest(rows=rows, distinct=distinct):
                manifest, receipt, evidence = self.completion_records()
                evidence["tables"]["researchModelInputs"].update(rows=rows, distinct=distinct)
                with self.assertRaisesRegex(ValueError, "approved row count"):
                    gate_input_completion(manifest, receipt, evidence)

    def test_reviewed_empty_schema_requires_empty_natural_key_evidence(self, gate):
        for keys in (None, [{"Scope": "unexpected"}], [{}]):
            with self.subTest(keys=keys):
                manifest, receipt, evidence = self.completion_records()
                evidence["tables"]["researchModelInputs"]["naturalKeys"] = keys
                with self.assertRaisesRegex(ValueError, "natural keys"):
                    gate_input_completion(manifest, receipt, evidence)

    def test_reviewed_empty_schema_preserves_source_and_plan_hash_checks(self, gate):
        for record_index in (1, 2):
            for field in ("sourceHash", "planHash"):
                with self.subTest(record_index=record_index, field=field):
                    records = self.completion_records()
                    records[record_index][field] = "c" * 64
                    with self.assertRaisesRegex(ValueError, field + " is stale"):
                        gate_input_completion(*records)

    def test_reviewed_empty_schema_preserves_approved_schema_and_key_columns(self, gate):
        manifest, receipt, evidence = self.completion_records()
        wrong_keys = copy.deepcopy(evidence)
        wrong_keys["tables"]["researchModelInputs"]["keyColumns"] = ["Other"]
        with self.assertRaisesRegex(ValueError, "key columns are stale"):
            gate_input_completion(manifest, receipt, wrong_keys)
        manifest["planProjection"]["sharedInputArchitecture"]["inputs"]["researchModelInputs"]["columns"][0]["type"] = "number"
        with self.assertRaisesRegex(ValueError, "artifacts are stale"):
            gate_input_completion(manifest, receipt, evidence)


if __name__ == "__main__":
    unittest.main()
