#!/usr/bin/env python3
"""Phase 0 — discover an .xlsx workbook.

Inventories formal Tables (the migration's Rosetta Stone), pivots/charts, and a
formula census (flagging untranslatable / report-style sheets). Prints a summary
and, for each formal Table, its columns + inferred grain.

Usage:
    python xlsx-discover.py "Sample Forecast.xlsx"
"""
import sys
import re
from collections import Counter
from openpyxl import load_workbook

# functions we translate today (see refs/excel-translation.md)
TRANSLATABLE = {
    "IF", "IFS", "SWITCH", "IFERROR", "AND", "OR", "NOT",
    "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MIN", "MAX", "MEDIAN",
    "VLOOKUP", "XLOOKUP", "HLOOKUP", "INDEX", "MATCH", "LOOKUP",
    "LEFT", "RIGHT", "MID", "LEN", "CONCAT", "CONCATENATE", "TRIM",
    "UPPER", "LOWER", "SUBSTITUTE", "FIND", "TEXT",
    "TODAY", "NOW", "YEAR", "MONTH", "DAY", "DATEDIF", "EDATE",
    "EOMONTH", "WEEKDAY", "DATE",
    "ROUND", "ROUNDUP", "ROUNDDOWN", "CEILING", "FLOOR", "ABS", "POWER", "MOD",
    "SUMPRODUCT",
}
# fail-loud: these usually mark a "report drawn in cells" needing manual rebuild
VOLATILE = {"OFFSET", "INDIRECT", "INDEX_VOLATILE"}

FUNC_RE = re.compile(r"\b([A-Z][A-Z0-9_.]+)\s*\(")


def main(path):
    wb = load_workbook(path, data_only=False)
    print(f"# Discovery: {path}\n")
    print(f"Sheets ({len(wb.sheetnames)}): {', '.join(wb.sheetnames)}\n")

    # ---- formal Tables ----
    tables = []
    for ws in wb.worksheets:
        for name in ws.tables:
            tbl = ws.tables[name]
            ref = tbl.ref
            # column headers from the first row of the ref
            cells = ws[ref]
            headers = [c.value for c in cells[0]] if cells else []
            nrows = len(cells) - 1 if cells else 0
            tables.append((name, ws.title, ref, headers, nrows))

    print(f"## Formal Tables ({len(tables)}) — DM-element / input-table candidates")
    for name, sheet, ref, headers, nrows in tables:
        print(f"\n  • {name}  (sheet '{sheet}', {ref}, ~{nrows} rows)")
        print(f"    columns: {', '.join(str(h) for h in headers)}")
    if not tables:
        print("  (none — workbook is likely 'report drawn in cells'; expect manual rebuild)")

    # ---- pivots / charts ----
    n_charts = sum(len(getattr(ws, "_charts", [])) for ws in wb.worksheets)
    n_pivots = sum(len(getattr(ws, "_pivots", [])) for ws in wb.worksheets)
    print(f"\n## Visuals: {n_charts} chart(s), {n_pivots} pivot table(s)")

    # ---- formula census ----
    func_counts = Counter()
    formula_cells = 0
    report_sheets = Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formula_cells += 1
                    for m in FUNC_RE.findall(c.value):
                        func_counts[m] += 1
                        if m in VOLATILE:
                            report_sheets[ws.title] += 1
    print(f"\n## Formula census: {formula_cells} formula cells, "
          f"{len(func_counts)} distinct functions")
    untranslatable = {f: n for f, n in func_counts.items()
                      if f not in TRANSLATABLE and f in VOLATILE}
    unknown = {f: n for f, n in func_counts.items()
               if f not in TRANSLATABLE and f not in VOLATILE}
    top = func_counts.most_common(15)
    print("  top functions:", ", ".join(f"{f}×{n}" for f, n in top))
    if untranslatable:
        print(f"  ⚠️ volatile/untranslatable (manual rebuild): "
              + ", ".join(f"{f}×{n}" for f, n in untranslatable.items()))
        print(f"     report-style sheets: {', '.join(report_sheets)}")
    if unknown:
        print(f"  ❓ not yet in the rules table (verify): "
              + ", ".join(f"{f}×{n}" for f, n in sorted(unknown.items())))

    print("\nNext: confirm the fact Table + grain, then xlsx-to-input-csv.py")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("usage: xlsx-discover.py <file.xlsx>")
    main(sys.argv[1])
