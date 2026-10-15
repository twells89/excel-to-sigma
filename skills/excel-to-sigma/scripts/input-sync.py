#!/usr/bin/env python3
"""Plan resumable input-table writes and emit reviewable Sigma MCP programs.

This CLI performs no MCP query or mutation.  Query current rows with the Sigma
MCP companion tools, run ``plan``, review the JSON/program bundle, execute the
program through Code Mode, then query again and retain the checkpoint evidence.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))

from input_orchestration import gate_input_completion, gate_input_staging  # noqa: E402
from input_sync import (  # noqa: E402
    build_shared_input_architecture,
    emit_construction_program,
    emit_write_programs,
    plan_seed,
)
from migration_manifest import build_manifest  # noqa: E402


def _load(path):
    with Path(path).open(encoding="utf-8") as stream:
        return json.load(stream)


def _default(value):
    if isinstance(value, (dt.date, dt.datetime)):
        return value.isoformat()
    raise TypeError(type(value).__name__)


def _write(path, value):
    Path(path).write_text(json.dumps(value, indent=2, default=_default) + "\n", encoding="utf-8")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    plan = commands.add_parser("plan", help="reconcile queried current rows and emit write programs")
    plan.add_argument("--source-rows", required=True)
    plan.add_argument("--current-rows", required=True)
    plan.add_argument("--keys", required=True, help="comma-separated business key columns")
    plan.add_argument("--row-id", required=True, help="queried Sigma system row-ID column")
    plan.add_argument("--prior-seed")
    plan.add_argument("--snapshot-id", required=True)
    plan.add_argument("--snapshot-age-seconds", required=True, type=int)
    plan.add_argument("--target", help="assigned page/element/column IDs; emits write chunks when supplied")
    plan.add_argument("--out", required=True)

    create = commands.add_parser("emit-create", help="emit an independent ordinary input-table program")
    create.add_argument("--spec", required=True)
    create.add_argument("--seed-rows", required=True)
    create.add_argument("--source-hash", required=True)
    create.add_argument("--plan-hash", required=True)
    create.add_argument("--out", required=True)

    architecture = commands.add_parser("architecture", help="add shared-input scope before manifest hashing")
    architecture.add_argument("--plan", required=True)
    architecture.add_argument("--source", required=True)
    architecture.add_argument("--connection", required=True)
    architecture.add_argument("--folder", required=True)
    architecture.add_argument("--out-plan", required=True)
    architecture.add_argument("--out-manifest", required=True)

    stage = commands.add_parser("stage", help="run the source/plan staging gate")
    stage.add_argument("--manifest", required=True)
    stage.add_argument("--approval", required=True)
    stage.add_argument("--architecture", required=True)
    stage.add_argument("--out", required=True)

    complete = commands.add_parser("complete", help="run completion gate from recorded runtime evidence")
    complete.add_argument("--manifest", required=True)
    complete.add_argument("--runtime-evidence", required=True)
    complete.add_argument("--out", required=True)
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if args.command == "plan":
        result = plan_seed(
            _load(args.source_rows), _load(args.current_rows),
            [key.strip() for key in args.keys.split(",") if key.strip()], args.row_id,
            prior_seed=_load(args.prior_seed) if args.prior_seed else None,
            snapshot={"id": args.snapshot_id, "ageSeconds": args.snapshot_age_seconds},
        )
        bundle = {"plan": result, "programs": emit_write_programs(result, _load(args.target))
                  if args.target else []}
        _write(args.out, bundle)
    elif args.command == "emit-create":
        program = emit_construction_program(
            _load(args.spec), _load(args.seed_rows), source_hash=args.source_hash,
            plan_hash=args.plan_hash)
        Path(args.out).write_text(program, encoding="utf-8")
    elif args.command == "architecture":
        plan = _load(args.plan)
        shared = build_shared_input_architecture(
            plan, connection_id=args.connection, folder_id=args.folder)
        plan["targetScope"] = shared["targetScope"]
        plan["sharedInputArchitecture"] = shared
        plan.pop("migrationManifest", None)
        manifest = build_manifest(args.source, plan)
        plan["migrationManifest"] = manifest
        _write(args.out_plan, plan)
        _write(args.out_manifest, manifest)
    elif args.command == "stage":
        receipt = gate_input_staging(_load(args.manifest), _load(args.approval),
                                     _load(args.architecture))
        _write(args.out, receipt)
    elif args.command == "complete":
        receipt = gate_input_completion(_load(args.manifest), _load(args.runtime_evidence))
        _write(args.out, receipt)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
