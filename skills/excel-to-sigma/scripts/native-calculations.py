#!/usr/bin/env python3
"""Read-only source binding and guarded native calculation artifact emission."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
from native_calculations import (build_native_calculation_plan,
    lower_native_calculations, prepare_native_calculation_update,
    confirm_native_calculation_readback, build_native_calculation_backup)


from native_exact_plan import (prepare_exact_native_scaffold, confirm_exact_native_scaffold,
    prepare_exact_native_update, confirm_exact_native_readback)


def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    plan = commands.add_parser("plan", help="account for exact source formulas")
    plan.add_argument("--source", required=True)
    plan.add_argument("--input-plan")
    plan.add_argument("--exact", action="store_true", help="typed error-preserving native v2 compiler")
    lower = commands.add_parser("lower-candidate", help="lower using observed assigned row IDs")
    lower.add_argument("--native-plan", required=True)
    lower.add_argument("--assigned-rows", required=True)
    lower.add_argument("--max-columns", type=int, default=64)
    backup = commands.add_parser("backup", help="persist complete document and queried rows")
    backup.add_argument("--current", required=True)
    backup.add_argument("--observed-rows", required=True)
    update = commands.add_parser("prepare-update", help="prepare one guarded REST contents candidate")
    for name in ("native-plan", "input-plan", "manifest", "staging-receipt", "current", "observed-rows"):
        update.add_argument("--" + name, required=True)
    update.add_argument("--shard-index", type=int, default=0)
    update.add_argument("--max-columns", type=int, default=64)
    update.add_argument("--prior-readbacks")
    update.add_argument("--backup", required=True)
    update.add_argument("--native-bindings", help="all actual assigned v2 scaffold receipts")
    scaffold = commands.add_parser("prepare-scaffold", help="prepare one receipt-bound MCP structural program")
    for name in ("native-plan", "input-plan", "manifest", "staging-receipt", "current", "observed-rows", "backup", "seed-binding", "page-id"):
        scaffold.add_argument("--" + name, required=True)
    scaffold.add_argument("--shard-index", type=int, default=0)
    scaffold_confirm = commands.add_parser("confirm-scaffold", help="record actual MCP scaffold element and column IDs")
    scaffold_confirm.add_argument("--proposal", required=True)
    scaffold_confirm.add_argument("--readback", required=True)
    confirm = commands.add_parser("confirm-readback", help="check saved native formula structure")
    confirm.add_argument("--proposal", required=True)
    confirm.add_argument("--readback", required=True)
    for command in (plan, lower, backup, update, confirm, scaffold, scaffold_confirm):
        command.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    if args.command == "plan":
        result = build_native_calculation_plan(args.source, _read(args.input_plan) if args.input_plan else None, exact=args.exact)
    elif args.command == "lower-candidate":
        result = lower_native_calculations(_read(args.native_plan), _read(args.assigned_rows),
                                           max_columns=args.max_columns)
    elif args.command == "backup":
        result = build_native_calculation_backup(_read(args.current), _read(args.observed_rows))
    elif args.command == "prepare-scaffold":
        result = prepare_exact_native_scaffold(_read(args.native_plan), _read(args.input_plan), _read(args.manifest),
            _read(args.staging_receipt), _read(args.current), _read(args.observed_rows), backup=_read(args.backup),
            seed_binding=_read(args.seed_binding), page_id=args.page_id, shard_index=args.shard_index)
    elif args.command == "confirm-scaffold":
        result = confirm_exact_native_scaffold(_read(args.proposal), _read(args.readback))
    elif args.command == "prepare-update" and _read(args.native_plan).get("nativePlanVersion") == 2:
        if not args.native_bindings:
            raise ValueError("exact native update requires --native-bindings from confirmed MCP scaffold readback")
        result = prepare_exact_native_update(_read(args.native_plan), _read(args.input_plan), _read(args.manifest),
            _read(args.staging_receipt), _read(args.current), _read(args.observed_rows), backup=_read(args.backup),
            native_bindings=_read(args.native_bindings), shard_index=args.shard_index,
            prior_readbacks=_read(args.prior_readbacks) if args.prior_readbacks else ())
    elif args.command == "prepare-update":
        result = prepare_native_calculation_update(
            _read(args.native_plan), _read(args.input_plan), _read(args.manifest),
            _read(args.staging_receipt), _read(args.current), _read(args.observed_rows),
            backup=_read(args.backup),
            shard_index=args.shard_index, max_columns=args.max_columns,
            prior_readbacks=_read(args.prior_readbacks) if args.prior_readbacks else ())
    else:
        proposal = _read(args.proposal)
        confirmer = confirm_exact_native_readback if proposal.get("status") == "GUARDED_EXACT_REST_CANDIDATE_PENDING_READBACK" else confirm_native_calculation_readback
        result = confirmer(proposal, _read(args.readback))
    destination = Path(args.out)
    if args.command == "plan" and destination.resolve() == Path(args.source).resolve():
        raise ValueError("native plan output cannot overwrite source workbook")
    destination.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
