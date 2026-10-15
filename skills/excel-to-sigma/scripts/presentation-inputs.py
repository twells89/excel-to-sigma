#!/usr/bin/env python3
"""Prepare and emit approved editable inputs used by presentation migration.

This command performs no Sigma calls. Construction and extension output is
derived only from an immutable manifest plus an APPROVED_STAGING receipt.
Extension seeding additionally requires a checkpoint made from assigned column
IDs and an immediate readback of those exact columns.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

LIB = Path(__file__).resolve().parent / "lib"
sys.path.insert(0, str(LIB))

from input_sync import (  # noqa: E402
    checkpoint_input_extension,
    emit_write_programs,
    plan_extension_seed,
)
from migration_manifest import build_manifest, validate_manifest  # noqa: E402
from presentation import (  # noqa: E402
    build_business_input_architecture,
    build_business_input_plan,
    emit_approved_business_input_program,
    extract_presentation,
)


def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _write(path, value):
    Path(path).write_text(json.dumps(value, indent=2, default=str) + "\n", encoding="utf-8")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    prepare = commands.add_parser("prepare", help="build the source-bound plan and manifest")
    prepare.add_argument("source")
    prepare.add_argument("--connection", required=True)
    prepare.add_argument("--folder", required=True)
    prepare.add_argument("--research-plan")
    prepare.add_argument("--reviewed-formulas",
                         help="approved replacement records for parser-unsupported source formulas")
    prepare.add_argument("--out-plan", required=True)
    prepare.add_argument("--out-manifest", required=True)

    target = commands.add_parser(
        "bind-target", help="bind a reviewed manifest to the current document version")
    target.add_argument("--manifest", required=True)
    target.add_argument("--document-id", required=True)
    target.add_argument("--document-version", required=True, type=int)
    target.add_argument("--out", required=True)

    emit = commands.add_parser("emit", help="emit approved construction or missing-column extension")
    emit.add_argument("--manifest", required=True)
    emit.add_argument("--receipt", required=True)
    emit.add_argument("--feature", required=True)
    emit.add_argument("--current-columns")
    emit.add_argument("--out", required=True)

    checkpoint = commands.add_parser(
        "checkpoint-extension", help="bind created IDs to the exact post-extension readback")
    checkpoint.add_argument("--extension-plan", required=True)
    checkpoint.add_argument("--created-columns", required=True)
    checkpoint.add_argument("--result-version", required=True, type=int)
    checkpoint.add_argument("--readback", required=True)
    checkpoint.add_argument("--out", required=True)

    seed = commands.add_parser(
        "seed-extension", help="plan first writes for checkpointed, still-null extension fields")
    seed.add_argument("--manifest", required=True)
    seed.add_argument("--receipt", required=True)
    seed.add_argument("--feature", required=True)
    seed.add_argument("--current-rows", required=True)
    seed.add_argument("--row-id", required=True)
    seed.add_argument("--prior-seed", required=True)
    seed.add_argument("--extension-proof", required=True)
    seed.add_argument("--snapshot-id", required=True)
    seed.add_argument("--snapshot-age-seconds", required=True, type=int)
    seed.add_argument("--out", required=True)
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if args.command == "prepare":
        extracted = extract_presentation(args.source)
        architecture = build_business_input_architecture(
            extracted, connection_id=args.connection, folder_id=args.folder,
            research_plan=_read(args.research_plan) if args.research_plan else None)
        reviewed_payload = _read(args.reviewed_formulas) if args.reviewed_formulas else {}
        reviewed = reviewed_payload.get("formulaCells", reviewed_payload if isinstance(reviewed_payload, list) else [])
        plan = build_business_input_plan(extracted, architecture, reviewed)
        manifest = build_manifest(args.source, plan)
        _write(args.out_plan, plan)
        _write(args.out_manifest, manifest)
    elif args.command == "bind-target":
        manifest = _read(args.manifest)
        validate_manifest(manifest)
        manifest["target"] = {"documentId": args.document_id,
                              "version": args.document_version}
        _write(args.out, manifest)
    elif args.command == "emit":
        result = emit_approved_business_input_program(
            _read(args.manifest), _read(args.receipt), args.feature,
            current_columns=_read(args.current_columns) if args.current_columns else None)
        _write(args.out, result)
    elif args.command == "checkpoint-extension":
        proof = checkpoint_input_extension(
            _read(args.extension_plan), _read(args.created_columns),
            result_version=args.result_version, readback=_read(args.readback))
        _write(args.out, proof)
    elif args.command == "seed-extension":
        manifest, receipt = _read(args.manifest), _read(args.receipt)
        feature = args.feature
        spec = manifest["planProjection"]["sharedInputArchitecture"]["inputs"][feature]
        result = plan_extension_seed(
            spec["rows"], _read(args.current_rows), spec["keyColumns"], args.row_id,
            prior_seed=_read(args.prior_seed), manifest=manifest,
            staging_receipt=receipt, feature=feature,
            extension_proof=_read(args.extension_proof),
            snapshot={"id": args.snapshot_id, "ageSeconds": args.snapshot_age_seconds},
            column_types={column["name"]: column["type"] for column in spec["columns"]})
        _write(args.out, {"plan": result,
                          "programs": emit_write_programs(result, manifest, receipt, feature)})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
