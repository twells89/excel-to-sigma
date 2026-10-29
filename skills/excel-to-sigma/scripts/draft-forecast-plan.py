#!/usr/bin/env python3
"""Draft a build-forecast-model.py plan from a cell-drawn actuals + driver P&L workbook.

Reads the P&L (formulas + Excel's cached values), the GL actuals tab, and the drivers
tab; discovers the sections and their sign roles from the derived rows; classifies each
line's forecast as growth / year-over-year / percent-of-total / flat / manual (an
external-link cell is kept as its cached value and flagged for review); re-derives every
line, total and derived row for every month or quarter from the drafted plan and compares
it with the workbook's cached values; and writes a source-bound plan JSON. Any shape it
does not recognise fails with the offending cell instead of being coerced.

usage: draft-forecast-plan.py <xlsx> --intent what-if --connection <id> --folder <id> --out plan.json [--look look.json]
Next:  build-forecast-model.py plan.json  (dry run), then  ... --post --review-dispositions review.json
"""
import argparse
import json
import sys
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
from lib.forecast_plan import INTENTS, PlanDraftError, draft_plan  # noqa: E402


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("workbook")
    parser.add_argument("--intent", default="what-if", choices=INTENTS)
    parser.add_argument("--connection", required=True, help="write-enabled Sigma connectionId")
    parser.add_argument("--folder", required=True, help="target Sigma folderId")
    parser.add_argument("--name", help="workbook name (default: workbook title, else file stem)")
    parser.add_argument("--sheet", help="report sheet (default: most SUMIFS/driver-growth cells)")
    parser.add_argument("--currency-format", default="$,.0f")
    parser.add_argument("--pct-format", default=".1%")
    parser.add_argument("--look", help="look.json from style-reference.py: the workbook takes that workbook's theme, "
                                       "card styles and KPI type (Excel's formats, titles and order still win)")
    parser.add_argument("--tolerance", type=float, default=None,
                        help="absolute tolerance when re-deriving cached values (default: float noise "
                             "only, max(1e-6, 1e-9 * |cached value|))")
    parser.add_argument("--out", required=True, help="plan JSON to write")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    look = json.loads(Path(args.look).read_text(encoding="utf-8")) if args.look else None
    try:
        plan, info = draft_plan(
            args.workbook, intent=args.intent, connection=args.connection, folder=args.folder,
            name=args.name, sheet=args.sheet, currency_format=args.currency_format,
            pct_format=args.pct_format, tolerance=args.tolerance, look=look)
    except PlanDraftError as error:
        print(f"DRAFT FAILED: {error}", file=sys.stderr)
        return 1
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, indent=1), encoding="utf-8")
    kinds = ", ".join(f"{k} {v}" for k, v in sorted(info["kinds"].items()))
    print(f"plan -> {out}")
    print(f"  {info['lines']} lines ({kinds}) | {info['glRows']} GL rows ({info['glSheet']}) "
          f"| {info['drivers']} drivers | intent={plan['intent']}")
    unit = "months" if info["grain"] == "month" else "periods"
    print(f"  {unit}: {info['actualRange'][0]}-{info['actualRange'][1]} actual, "
          f"{info['forecastRange'][0]}-{info['forecastRange'][1]} forecast "
          f"(report sheet {info['reportSheet']!r})")
    print("  sections: " + "; ".join(info["sections"]))
    print("  derived rows: " + (", ".join(info["derivedRows"]) or "none"))
    covered = (f"FY column {info['fyColumn']}" + (f" over {len(info['fyPeriods'])} of the {unit}"
                                                    if info["fyPeriods"] else "")
               if info["fyColumn"] else f"summed {unit}")
    print(f"  parity targets ({covered}): "
          + ("  ".join(f"{k}={v:,.2f}" for k, v in info["parityTargets"].items()) or "none"))
    if info["missingTargets"] and not info["parityTargets"]:
        print("  no target found for: " + ", ".join(info["missingTargets"]))
    for note in info["skippedTargets"]:
        print(f"  no parity target: {note}")
    if info["ignoredColumns"]:
        print("  columns not carried (outside the periods and FY total): " + ", ".join(info["ignoredColumns"]))
    print(f"  verified: {info['verifiedCells']} line / total / derived-row cells and every GL "
          f"{unit[:-1]} sum match the workbook's cached values")
    shown = info["presentation"]
    print(f"  presentation: title {shown['title']!r}; money {shown['money']}; "
          f"{shown['percent'] or 'no percent row (--pct-format)'}")
    for note in shown["notes"]:
        print(f"  note: {note}")
    if look:
        print(f"  look: the workbook takes the look of {(look.get('source') or {}).get('name') or args.look} "
              "(theme, card styles, KPI type); Excel's formats, titles and order still win")
    for chart in info["charts"]:
        print(f"  source chart mapped + values verified: {chart}")
    for chart in info["unmappedCharts"]:
        print(f"  source chart NOT mapped (rebuild by hand): {chart}")
    print("  review issues: " + (", ".join(info["reviewIssues"]) or "none from the drafter")
          + " (plus any workbook-audit issues; see migrate.py plan)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
