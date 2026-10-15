"""Source/plan-bound gates for shared-input staging and completion."""
from __future__ import annotations

try:
    from .migration_gate import gate_completion_manifest, gate_staging_manifest
except ImportError:
    from migration_gate import gate_completion_manifest, gate_staging_manifest


def gate_input_staging(manifest, approval, architecture):
    """Call the PR1 staging gate, then bind the approved input feature scope."""
    gate_staging_manifest(manifest, approval)
    for field in ("sourceHash", "planHash", "reviewer", "mappedFeatures"):
        if not approval.get(field):
            raise ValueError(f"approved staging requires {field}")
    if approval["sourceHash"] != manifest.get("sourceHash"):
        raise ValueError("sourceHash does not match the manifest")
    if approval["planHash"] != manifest.get("planHash"):
        raise ValueError("planHash does not match the manifest")
    if set(approval["mappedFeatures"]) != set(architecture.get("mappedFeatures") or []):
        raise ValueError("mappedFeatures must exactly match the approved architecture")
    return {"status": "APPROVED_STAGING", "sourceHash": approval["sourceHash"],
            "planHash": approval["planHash"], "reviewer": approval["reviewer"],
            "mappedFeatures": sorted(approval["mappedFeatures"]),
            "sourceStatus": "PRESERVED"}


def gate_input_completion(manifest, runtime_evidence):
    """Call the PR1 completion gate and require persisted runtime checkpoints."""
    gate_completion_manifest(manifest)
    required = ("sourceHash", "planHash", "querySnapshotId", "readbackId",
                "checkpointIds", "uniqueKeyCounts", "persistenceReadbackId",
                "acceptanceChecks", "status")
    missing = [field for field in required if not runtime_evidence.get(field)]
    if missing:
        raise ValueError("runtime evidence missing: " + ", ".join(missing))
    if runtime_evidence["sourceHash"] != manifest.get("sourceHash"):
        raise ValueError("runtime evidence sourceHash is stale")
    if runtime_evidence["planHash"] != manifest.get("planHash"):
        raise ValueError("runtime evidence planHash is stale")
    if runtime_evidence["status"] != "passed":
        raise ValueError("runtime evidence has not passed")
    counts = runtime_evidence["uniqueKeyCounts"]
    if not isinstance(counts, dict) or any(
            item.get("rows") != item.get("distinct") for item in counts.values()):
        raise ValueError("runtime evidence contains duplicate business keys")
    checks = runtime_evidence["acceptanceChecks"]
    if not isinstance(checks, list) or not checks or any(
            check.get("status") != "passed" or check.get("restored") is not True
            for check in checks):
        raise ValueError("runtime acceptance edits must pass and be restored")
    return {"status": "VERIFIED_COMPLETE", "readbackId": runtime_evidence["readbackId"],
            "checkpointIds": list(runtime_evidence["checkpointIds"])}
