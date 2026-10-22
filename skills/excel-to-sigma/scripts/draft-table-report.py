#!/usr/bin/env python3
"""Draft a build-table-report.py plan from a workbook of formal Tables + formula-driven grids.

Reads the tables (literal columns = data, calculated columns = row formulas) and the report
grids (summary / expansion / keyed), lowers every formula to a small IR, recomputes every
recognised cell from the table data, and compares it with Excel's cached value (tolerance
0.01; text flags exact). It writes a source-bound plan JSON. Any formula it does not
recognise, and any mismatch, fails with the exact sheet!cell - nothing is coerced and no
cached value is frozen into the plan as a "conversion".

usage: draft-table-report.py <xlsx|xlsm> --intent what-if --connection <id> --folder <id> --out plan.json
Next:  build-table-report.py plan.json  (dry run), then  ... --post --review-dispositions review.json
"""
import argparse
import json
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
from lib.table_report import INTENTS, PlanDraftError, draft_table_report  # noqa: E402


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("workbook")
    parser.add_argument("--intent", default="what-if", choices=INTENTS,
                        help="read-only: every table is a data-model VALUES element; editable | what-if: "
                             "tables the formulas read become input tables (default: what-if)")
    parser.add_argument("--connection", required=True, help="write-enabled Sigma connectionId")
    parser.add_argument("--folder", required=True, help="target Sigma folderId")
    parser.add_argument("--name", help="workbook name (default: workbook title, else file stem)")
    parser.add_argument("--currency-format", default="$,.0f")
    parser.add_argument("--pct-format", default=".1%")
    parser.add_argument("--tolerance", type=float, default=0.01,
                        help="absolute tolerance when re-deriving cached values (default: 0.01)")
    parser.add_argument("--out", required=True, help="plan JSON to write")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    try:
        plan, info = draft_table_report(
            args.workbook, intent=args.intent, connection=args.connection, folder=args.folder,
            name=args.name, currency_format=args.currency_format, pct_format=args.pct_format,
            tolerance=args.tolerance)
    except PlanDraftError as error:
        print(f"DRAFT FAILED: {error}", file=sys.stderr)
        return 1
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, indent=1), encoding="utf-8")
    print(f"plan -> {out}")
    print(f"  {len(info['tables'])} tables ({len(info['editable'])} editable, {len(info['readOnly'])} read-only) "
          f"| {len(info['frames'])} frames | {info['params']} parameters | intent={plan['intent']}")
    for kind, title in info["frames"]:
        print(f"  frame {kind}: {title}")
    print(f"  verified: {info['verifiedCells']} source cells reproduced from the table data "
          f"({info['formulaCells']} formula cells in the workbook, every one covered)")
    print(f"  parity targets: {info['parityTargets']} summary cells (re-checked by build-table-report.py)")
    if info["editable"]:
        print("  editable (input tables, seeded afterwards): " + ", ".join(info["editable"]))
    if info["readOnly"]:
        print("  read-only (data-model VALUES): " + ", ".join(info["readOnly"]))
    for chart in info["unmappedCharts"]:
        print(f"  source chart NOT mapped (rebuild by hand): {chart}")
    print("  review issues: " + (", ".join(info["reviewIssues"]) or "none from the drafter")
          + " (plus any workbook-audit issues; see migrate.py plan)")
    print("  ask the user: " + "; ".join(plan["askUser"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
