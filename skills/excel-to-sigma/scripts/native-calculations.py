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


def _read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    plan = commands.add_parser("plan", help="account for exact source formulas")
    plan.add_argument("--source", required=True)
    plan.add_argument("--input-plan")
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
    confirm = commands.add_parser("confirm-readback", help="check saved native formula structure")
    confirm.add_argument("--proposal", required=True)
    confirm.add_argument("--readback", required=True)
    for command in (plan, lower, backup, update, confirm):
        command.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    if args.command == "plan":
        result = build_native_calculation_plan(args.source, _read(args.input_plan) if args.input_plan else None)
    elif args.command == "lower-candidate":
        result = lower_native_calculations(_read(args.native_plan), _read(args.assigned_rows),
                                           max_columns=args.max_columns)
    elif args.command == "backup":
        result = build_native_calculation_backup(_read(args.current), _read(args.observed_rows))
    elif args.command == "prepare-update":
        result = prepare_native_calculation_update(
            _read(args.native_plan), _read(args.input_plan), _read(args.manifest),
            _read(args.staging_receipt), _read(args.current), _read(args.observed_rows),
            backup=_read(args.backup),
            shard_index=args.shard_index, max_columns=args.max_columns,
            prior_readbacks=_read(args.prior_readbacks) if args.prior_readbacks else ())
    else:
        result = confirm_native_calculation_readback(_read(args.proposal), _read(args.readback))
    destination = Path(args.out)
    if args.command == "plan" and destination.resolve() == Path(args.source).resolve():
        raise ValueError("native plan output cannot overwrite source workbook")
    destination.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
