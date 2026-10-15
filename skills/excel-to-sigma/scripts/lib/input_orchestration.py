"""Source/plan/target-bound gates for shared-input staging and completion."""
from __future__ import annotations

try:
    from .input_sync import build_artifact_records
    from .migration_gate import gate_completion_manifest, gate_staging_manifest
except ImportError:
    from input_sync import build_artifact_records
    from migration_gate import gate_completion_manifest, gate_staging_manifest


def _architecture(value):
    if not isinstance(value, dict):
        raise ValueError("approved architecture must be an object")
    architecture = value.get("sharedInputArchitecture", value)
    if not isinstance(architecture, dict) or not isinstance(architecture.get("inputs"), dict):
        raise ValueError("approved architecture is missing sharedInputArchitecture.inputs")
    return architecture


def _same_hashes(manifest, value, label):
    if value.get("sourceHash") != manifest.get("sourceHash"):
        raise ValueError(f"{label} sourceHash is stale")
    if value.get("planHash") != manifest.get("planHash"):
        raise ValueError(f"{label} planHash is stale")


def gate_input_staging(manifest, approval, plan_or_architecture, source_gate_evidence):
    """Run PR1 review, preserve the real source gate, and bind exact artifacts."""
    gate_staging_manifest(manifest, approval)
    architecture = _architecture(plan_or_architecture)
    for field in ("sourceHash", "planHash", "reviewer", "mappedFeatures", "targetBindings"):
        if not approval.get(field):
            raise ValueError(f"approved staging requires {field}")
    _same_hashes(manifest, approval, "approval")
    if set(approval["mappedFeatures"]) != set(architecture.get("mappedFeatures") or []):
        raise ValueError("mappedFeatures must exactly match the approved architecture")
    if not isinstance(source_gate_evidence, dict):
        raise ValueError("source gate evidence is required")
    if source_gate_evidence.get("sourceSha256") != manifest.get("sourceHash"):
        raise ValueError("source gate evidence sourceHash is stale")
    source_status = source_gate_evidence.get("status")
    review_status = source_gate_evidence.get("reviewStatus")
    if source_status == "BLOCKED" or review_status == "BLOCKED":
        raise ValueError("source audit is BLOCKED")
    if review_status not in {"NOT_REQUIRED", "DISPOSITIONED"}:
        raise ValueError("source audit review is unresolved")
    artifacts = build_artifact_records(manifest, architecture, approval["targetBindings"])
    return {
        "status": "APPROVED_STAGING", "sourceHash": approval["sourceHash"],
        "planHash": approval["planHash"], "reviewer": approval["reviewer"],
        "mappedFeatures": sorted(approval["mappedFeatures"]),
        "sourceStatus": source_status, "sourceReviewStatus": review_status,
        "sourceAuditEvidenceId": source_gate_evidence.get("evidenceId"),
        "targetBindings": approval["targetBindings"], "artifacts": artifacts,
    }


def _validate_target_bindings(manifest, receipt, runtime_evidence, architecture):
    receipt_bindings = receipt.get("targetBindings") or {}
    if runtime_evidence.get("targetBindings") != receipt_bindings:
        raise ValueError("runtime target bindings do not match approved target bindings")
    expected_tables = set(architecture["inputs"])
    if set(receipt_bindings) != expected_tables:
        raise ValueError("approved target bindings must exactly cover input tables")
    target = runtime_evidence.get("target") or {}
    manifest_target = manifest.get("target") or {}
    if target != manifest_target or not target.get("documentId") or target.get("version") is None:
        raise ValueError("runtime target document/version does not match manifest")
    for name, binding in receipt_bindings.items():
        if (binding.get("documentId") != target["documentId"]
                or binding.get("version") != target["version"]
                or not binding.get("pageId") or not binding.get("elementId")):
            raise ValueError(f"target binding is incomplete or stale for {name}")
        expected_columns = {column["name"] for column in architecture["inputs"][name]["columns"]}
        if set(binding.get("columnIds") or {}) != expected_columns:
            raise ValueError(f"target columns are incomplete or stale for {name}")


def gate_input_completion(manifest, staging_receipt, runtime_evidence):
    """Require complete feature/table/readback coverage for the approved live target."""
    gate_completion_manifest(manifest)
    architecture = _architecture((manifest.get("planProjection") or {}).get(
        "sharedInputArchitecture", {}))
    if staging_receipt.get("status") != "APPROVED_STAGING":
        raise ValueError("completion requires an approved staging receipt")
    _same_hashes(manifest, staging_receipt, "staging receipt")
    _same_hashes(manifest, runtime_evidence, "runtime evidence")
    if runtime_evidence.get("status") != "passed":
        raise ValueError("runtime evidence has not passed")
    expected_features = set(architecture.get("mappedFeatures") or [])
    if (set(staging_receipt.get("mappedFeatures") or []) != expected_features
            or set(runtime_evidence.get("mappedFeatures") or []) != expected_features):
        raise ValueError("runtime mapped features must exactly cover approved feature scope")
    expected_artifacts = build_artifact_records(
        manifest, architecture, staging_receipt.get("targetBindings") or {})
    if staging_receipt.get("artifacts") != expected_artifacts:
        raise ValueError("staging receipt artifacts are stale")
    _validate_target_bindings(manifest, staging_receipt, runtime_evidence, architecture)

    expected_tables = set(architecture["inputs"])
    tables = runtime_evidence.get("tables")
    if not isinstance(tables, dict) or set(tables) != expected_tables:
        raise ValueError("runtime table evidence must exactly cover approved input tables")
    for name, spec in architecture["inputs"].items():
        item = tables[name]
        for field in ("querySnapshotId", "readbackId", "persistenceReadbackId"):
            if not item.get(field):
                raise ValueError(f"runtime table evidence for {name} requires {field}")
        if item.get("keyColumns") != spec.get("keyColumns"):
            raise ValueError(f"runtime key columns are stale for {name}")
        if (not isinstance(item.get("rows"), int) or item["rows"] < 1
                or item.get("distinct") != item["rows"]):
            raise ValueError(f"runtime evidence contains duplicate or empty business keys for {name}")

    checkpoints = runtime_evidence.get("checkpointBindings")
    if not isinstance(checkpoints, list) or not checkpoints:
        raise ValueError("checkpoint bindings are required")
    covered = set()
    target = runtime_evidence["target"]
    for item in checkpoints:
        feature = item.get("feature")
        if feature not in expected_tables or not item.get("checkpointId"):
            raise ValueError("checkpoint binding has an unknown feature")
        binding = staging_receipt["targetBindings"][feature]
        if (item.get("artifactHash") != expected_artifacts[feature]["artifactHash"]
                or item.get("documentId") != target["documentId"]
                or item.get("version") != target["version"]
                or item.get("elementId") != binding["elementId"]):
            raise ValueError(f"checkpoint binding is stale for {feature}")
        covered.add(feature)
    if covered != expected_tables:
        raise ValueError("checkpoint bindings must exactly cover approved input tables")

    checks = runtime_evidence.get("acceptanceChecks")
    expected_checks = set(architecture.get("acceptanceChecks") or [])
    if not isinstance(checks, list) or {check.get("id") for check in checks} != expected_checks:
        raise ValueError("acceptance checks must exactly cover approved feature checks")
    for check in checks:
        if (check.get("status") != "passed" or check.get("restored") is not True
                or not check.get("evidenceId")
                or check.get("documentId") != target["documentId"]
                or check.get("version") != target["version"]
                or check.get("sourceHash") != manifest["sourceHash"]
                or check.get("planHash") != manifest["planHash"]):
            raise ValueError(f"acceptance check is incomplete or stale: {check.get('id')}")
    return {
        "status": "VERIFIED_COMPLETE", "target": target,
        "tableReadbackIds": {name: item["readbackId"] for name, item in tables.items()},
        "checkpointIds": [item["checkpointId"] for item in checkpoints],
        "acceptanceEvidenceIds": [item["evidenceId"] for item in checks],
    }
