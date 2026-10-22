#!/usr/bin/env python3
"""Emit Sigma MCP Code Mode programs that create, seed and edit a LINKED input table.

Linked input tables cannot be authored through the REST workbook spec (inherited
columns show "multiple values").  Through Sigma MCP Code Mode they work: validated
live 2026-10-22.  This CLI performs no MCP or REST call; it writes three programs
(``call1.py`` base + one-row-per-key grouped table, ``call2.py`` linked table and
columns, ``call3.py`` rows by key), readback SQL, and a manifest.  Run each program
with the ``build_workbook`` tool in order, capture the printed ids with ``ids``,
re-emit with ``--ids`` so calls 2 and 3 carry real ids, then run the readback SQL
with the ``query`` tool (type ``workbook``) using the ids ``describe`` prints.

    linked-input-program.py emit --spec spec.json --out-dir out/
    linked-input-program.py ids --from-output call1-output.txt --out ids.json
    linked-input-program.py emit --spec spec.json --ids ids.json --out-dir out/

See refs/input-tables.md ("Linked input tables") for the spec fields.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))

from linked_inputs import LinkedInputError, emit, parse_printed_ids  # noqa: E402


def _load(path):
    with Path(path).open(encoding="utf-8") as stream:
        return json.load(stream)


def _cmd_emit(args):
    spec = _load(args.spec)
    ids = _load(args.ids) if args.ids else {}
    result = emit(spec, ids)
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    for name, text in result["programs"].items():
        (out / (name + ".py")).write_text(text, encoding="utf-8")
    (out / "readback.json").write_text(json.dumps(result["readback"], indent=2) + "\n", encoding="utf-8")
    (out / "manifest.json").write_text(json.dumps(result["manifest"], indent=2) + "\n", encoding="utf-8")
    for call in result["manifest"]["calls"]:
        pending = call["unresolvedPlaceholders"]
        print(call["file"] + ": target=" + call["target"]["type"]
              + (" (placeholders to fill via --ids: " + ", ".join(pending) + ")" if pending else ""))
    for warning in result["warnings"]:
        print("warning: " + warning, file=sys.stderr)
    print("wrote " + str(out))


def _cmd_ids(args):
    ids = _load(args.out) if args.out and Path(args.out).exists() else {}
    ids.update(parse_printed_ids(Path(args.from_output).read_text(encoding="utf-8")))
    if args.workbook_id:
        ids["workbookId"] = args.workbook_id
    text = json.dumps(ids, indent=2) + "\n"
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    emit_cmd = commands.add_parser("emit", help="validate a spec and write the programs + readback SQL")
    emit_cmd.add_argument("--spec", required=True)
    emit_cmd.add_argument("--out-dir", required=True)
    emit_cmd.add_argument("--ids", help="ids.json from earlier calls (and describe) to fill placeholders")
    ids_cmd = commands.add_parser("ids", help="merge ids printed by a program run into ids.json")
    ids_cmd.add_argument("--from-output", required=True, help="text file holding the program's print output")
    ids_cmd.add_argument("--out", help="ids.json to create or update (default: stdout)")
    ids_cmd.add_argument("--workbook-id", help="workbook id returned by call 1")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    try:
        {"emit": _cmd_emit, "ids": _cmd_ids}[args.command](args)
    except LinkedInputError as exc:
        print("error: " + str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
