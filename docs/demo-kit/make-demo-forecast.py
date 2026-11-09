#!/usr/bin/env python3
"""Build the demo source workbook: a macro-era FP&A rolling forecast (fictional Vantora Software).

This is the "before" file for the Excel -> Sigma demo. It is deliberately shaped for the
skill's *forecast* path (see lib/forecast_plan.py):

  Forecast   P&L-in-cells: sections Revenue / Cost of Revenue / Operating Expenses, one line
             per region (4), month columns Jan-26..Dec-27 (18 actual, 6 forecast), actual
             months = SUMIFS into Actuals, forecast months = prior*(1+Assumptions!rate)
             (plus one flat and one typed-literal opex line), derived rows Gross Margin / EBITDA.
  Actuals    tblActuals - GL postings, region x product line x section x month.
  Plan, Dimensions, Assumptions, Variance, Versions, Approvals, Instructions - the supporting
  tabs a real model of this kind carries, with formal Tables tblPlan, tblRegions, tblProducts,
  tblDrivers, tblVersions, tblApprovals.

openpyxl writes formulas without cached values, but the drafter (and Excel viewers) read the
cached results, so every formula cell's value is computed here in Python and injected into the
saved sheet XML. Macro buttons are drawn as button-styled cells (openpyxl cannot bind macros);
the VBA lives in demo-forecast-macros.bas and is imported by hand to produce the .xlsm.

Demo kit only (not part of the skill's migration path).

Usage: docs/demo-kit/make-demo-forecast.py --out <dir>   -> <dir>/Vantora_FY27_Rolling_Forecast_v7_FINAL.xlsx
"""
from __future__ import annotations

import argparse
import random
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

COMPANY = "Vantora Software"
FILE_NAME = "Vantora_FY27_Rolling_Forecast_v7_FINAL.xlsx"

MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTHS = [f"{MON[i % 12]}-{26 + i // 12}" for i in range(24)]      # Jan-26 .. Dec-27
N_ACT = 18                                                          # Jan-26 .. Jun-27 are actuals
FY_FIRST = 12                                                       # FY27 = Jan-27 .. Dec-27
HEADER_ROW = 4

REGIONS = [  # name, head, currency, base monthly revenue ($), monthly growth
    ("North America", "Dana Whitfield", "USD", 4_200_000, 0.012),
    ("EMEA", "Lukas Brandt", "EUR", 2_600_000, 0.011),
    ("APAC", "Mei Tanaka", "JPY", 1_700_000, 0.016),
    ("LATAM", "Camila Duarte", "BRL", 900_000, 0.014),
]
PRODUCTS = [  # name, category, revenue share, cost-of-revenue ratio, margin target
    ("Platform Subscription", "Recurring", 0.58, 0.20, 0.80),
    ("Professional Services", "Services", 0.27, 0.64, 0.36),
    ("Support & Maintenance", "Recurring", 0.15, 0.38, 0.62),
]
SEASON = [0.93, 0.94, 1.00, 1.00, 1.02, 1.05, 0.97, 0.96, 1.04, 1.06, 1.10, 1.18]
OPEX_RATIO = {"North America": 0.36, "EMEA": 0.40, "APAC": 0.44, "LATAM": 0.47}

# forecast behaviour: (kind, monthly rate or manual list); rates are the Base scenario
REV_RATE = {"North America": 0.010, "EMEA": 0.009, "APAC": 0.014, "LATAM": 0.012}
COGS_RATE = {"North America": 0.008, "EMEA": 0.008, "APAC": 0.011, "LATAM": 0.010}
OPEX_PLAN = {"North America": ("growth", 0.006), "EMEA": ("growth", 0.005),
             "APAC": ("flat", None),
             "LATAM": ("manual", [610_000.0, 615_000.0, 622_000.0, 630_000.0, 640_000.0, 655_000.0])}

SECTIONS = ["Revenue", "Cost of Revenue", "Operating Expenses"]
SUFFIX = {"Revenue": "Revenue", "Cost of Revenue": "Cost of Revenue", "Operating Expenses": "Opex"}

NAVY = "1F3864"
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
HEAD_FONT = Font(bold=True, color="FFFFFF")
ACT_FILL = PatternFill("solid", fgColor="EAF1FB")
FC_FILL = PatternFill("solid", fgColor="FFF4D6")
BTN_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="8EA9DB")
MONEY = '#,##0;(#,##0);"-"'
PCT = "0.0%"


def line_label(region: str, section: str) -> str:
    return f"{region} {SUFFIX[section]}"


def _inject_cached_values(path: Path, cache: dict, sheet_names: list[str]) -> None:
    temp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            match = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if match:
                sheet = sheet_names[int(match.group(1)) - 1]

                def fill(found, sheet=sheet):
                    value = cache.get((sheet, found.group(2)))
                    if value is None:
                        return found.group(0)
                    return f"{found.group(1)}{found.group(3)}<v>{value!r}</v>"

                data = re.sub(r'(<c r="([A-Z]+\d+)"[^>]*>)(<f>.*?</f>)<v />',
                              fill, data.decode("utf-8")).encode("utf-8")
            target.writestr(item, data)
    shutil.move(temp, path)


def _table(ws, name: str, ref: str, style: str = "TableStyleMedium2") -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)


def _button(ws, cell: str, caption: str, width_cols: int = 2) -> None:
    """Button look-alike (a macro shape in the real file): bordered, filled, bold, centred."""
    first = ws[cell]
    first.value = caption
    first.fill = BTN_FILL
    first.font = Font(bold=True, color=NAVY)
    first.alignment = Alignment(horizontal="center", vertical="center")
    box = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    first.border = box
    if width_cols > 1:
        end = ws.cell(first.row, first.column + width_cols - 1)
        ws.merge_cells(start_row=first.row, start_column=first.column,
                       end_row=end.row, end_column=end.column)
        for col in range(first.column, end.column + 1):
            ws.cell(first.row, col).border = box
            ws.cell(first.row, col).fill = BTN_FILL


def build_demo_workbook(out: Path, seed: int = 27) -> Path:
    """Write the demo ``.xlsx``; ``out`` is a directory (default file name) or a ``.xlsx`` path."""
    out = Path(out)
    path = out if out.suffix.lower() == ".xlsx" else out / FILE_NAME
    path.parent.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)

    wb = Workbook()
    ins = wb.active
    ins.title = "Instructions"
    act_ws = wb.create_sheet("Actuals")
    plan_ws = wb.create_sheet("Plan")
    dim = wb.create_sheet("Dimensions")
    asm = wb.create_sheet("Assumptions")
    fc = wb.create_sheet("Forecast")
    var = wb.create_sheet("Variance")
    ver = wb.create_sheet("Versions")
    apr = wb.create_sheet("Approvals")
    cache: dict[tuple[str, str], float] = {}

    def put(ws, coord, formula, value):
        ws[coord] = formula
        cache[(ws.title, coord)] = float(value)

    # ------------------------------------------------------------------ Dimensions
    dim.append(["Region", "Region Head", "Currency"])
    for name, head, ccy, *_ in REGIONS:
        dim.append([name, head, ccy])
    _table(dim, "tblRegions", f"A1:C{len(REGIONS) + 1}")
    for col, head in zip("EFG", ["Product Line", "Category", "Margin Target"]):
        dim[f"{col}1"] = head
    for i, (name, category, _, _, target) in enumerate(PRODUCTS, 2):
        dim[f"E{i}"], dim[f"F{i}"], dim[f"G{i}"] = name, category, target
        dim[f"G{i}"].number_format = "0%"
    _table(dim, "tblProducts", f"E1:G{len(PRODUCTS) + 1}", "TableStyleMedium9")
    for col, width in zip("ABCDEFG", (18, 18, 12, 4, 24, 14, 14)):
        dim.column_dimensions[col].width = width

    # ------------------------------------------------------------------ Actuals (GL)
    headers = ["Posting Date", "Month", "Line Item", "Region", "Product Line", "Amount"]
    act_ws.append(headers)
    amounts: dict[tuple[str, int], float] = {}     # (line label, month idx) -> sum of postings
    rows = []
    for m in range(N_ACT):
        season = SEASON[m % 12]
        for region, _head, _ccy, base, growth in REGIONS:
            region_rev = base * (1 + growth) ** m * season * rng.uniform(0.97, 1.03)
            region_cost_noise = rng.uniform(0.98, 1.02)
            region_rev_total = 0.0
            for product, _cat, share, cost_ratio, _t in PRODUCTS:
                rev = round(region_rev * share * rng.uniform(0.96, 1.04), 2)
                region_rev_total += rev
                cogs = round(rev * cost_ratio * region_cost_noise, 2)
                rows.append((m, region, product, "Revenue", rev))
                rows.append((m, region, product, "Cost of Revenue", cogs))
            opex_total = region_rev_total * OPEX_RATIO[region] * rng.uniform(0.97, 1.03)
            for product, _cat, share, *_ in PRODUCTS:
                rows.append((m, region, product, "Operating Expenses", round(opex_total * share, 2)))
    for m, region, product, section, value in rows:
        label = line_label(region, section)
        act_ws.append([datetime(2026 + m // 12, m % 12 + 1, 1), MONTHS[m], label, region, product, value])
        amounts[(label, m)] = round(amounts.get((label, m), 0.0) + value, 2)
    act_last = len(rows) + 1
    for r in range(2, act_last + 1):
        act_ws.cell(r, 1).number_format = "yyyy-mm-dd"
        act_ws.cell(r, 6).number_format = "#,##0.00"
    _table(act_ws, "tblActuals", f"A1:F{act_last}")
    for col, width in zip("ABCDEF", (13, 9, 28, 15, 24, 14)):
        act_ws.column_dimensions[col].width = width
    act_ws.freeze_panes = "A2"

    # ------------------------------------------------------------------ Assumptions
    asm["A1"] = f"{COMPANY} - FY27 forecast drivers"
    asm["A1"].font = Font(bold=True, size=14, color=NAVY)
    asm["A2"], asm["B2"] = "Active scenario", "Base"
    asm["A2"].font = Font(bold=True)
    asm.append([])
    for col, head in zip("ABCDEF", ["Driver", "Monthly Rate", "Basis", "Downside", "Base", "Upside"]):
        asm[f"{col}4"] = head
    driver_row: dict[str, int] = {}
    drivers: list[tuple[str, str, float]] = []   # (line label, basis text, base rate)
    for region, *_ in REGIONS:
        drivers.append((line_label(region, "Revenue"), "Net new ARR + expansion", REV_RATE[region]))
    for region, *_ in REGIONS:
        drivers.append((line_label(region, "Cost of Revenue"), "Hosting + delivery cost growth",
                        COGS_RATE[region]))
    for region, *_ in REGIONS:
        kind, param = OPEX_PLAN[region]
        if kind == "growth":
            drivers.append((line_label(region, "Operating Expenses"), "Headcount + programme spend", param))
    for i, (label, basis, rate) in enumerate(drivers, 5):
        driver_row[label] = i
        asm[f"A{i}"] = f"{label} growth"
        asm[f"B{i}"] = rate
        asm[f"C{i}"] = basis
        asm[f"D{i}"] = round(rate - 0.004, 4)
        asm[f"E{i}"] = rate
        asm[f"F{i}"] = round(rate + 0.004, 4)
        for col in "BDEF":
            asm[f"{col}{i}"].number_format = "0.0%"
    _table(asm, "tblDrivers", f"A4:F{4 + len(drivers)}")
    for col, width in zip("ABCDEF", (36, 14, 32, 12, 12, 12)):
        asm.column_dimensions[col].width = width

    # ------------------------------------------------------------------ Forecast
    fc["A1"] = f"{COMPANY} - FY27 Rolling Forecast (v7 FINAL)"
    fc["A1"].font = Font(bold=True, size=14, color=NAVY)
    fc["A2"] = "USD, reporting currency. Actuals via Refresh Actuals; forecast driven by Assumptions."
    fc["A2"].font = Font(italic=True, color="666666")
    _button(fc, "D2", "Refresh Actuals", 2)
    _button(fc, "G2", "Roll Forward Month", 2)
    _button(fc, "J2", "Snapshot Version", 2)
    _button(fc, "M2", "Submit for Approval", 2)
    for i in range(24):
        c = fc.cell(3, 2 + i, "Actual" if i < N_ACT else "Forecast")
        c.font = Font(italic=True, size=9, color="666666")
        c.alignment = Alignment(horizontal="center")
        h = fc.cell(HEADER_ROW, 2 + i, MONTHS[i])
        h.fill, h.font, h.alignment = HEAD_FILL, HEAD_FONT, Alignment(horizontal="center")
    fy_col = 2 + 24
    fy = fc.cell(HEADER_ROW, fy_col, "FY27")
    fy.fill, fy.font, fy.alignment = HEAD_FILL, HEAD_FONT, Alignment(horizontal="center")
    fc.cell(HEADER_ROW, 1, "USD").fill = HEAD_FILL
    fc.cell(HEADER_ROW, 1).font = HEAD_FONT
    fy_cols = f"{get_column_letter(2 + FY_FIRST)}{{r}}:{get_column_letter(25)}{{r}}"

    series: dict[str, list[float]] = {}
    fc_row_of: dict[str, int] = {}
    row = HEADER_ROW + 1
    total_rows: dict[str, int] = {}
    total_vals: dict[str, list[float]] = {}
    act_rng = lambda c: f"Actuals!${c}$2:${c}${act_last}"
    for section in SECTIONS:
        fc.cell(row, 1, section).font = Font(bold=True, color=NAVY)
        row += 1
        first = row
        for region, *_ in REGIONS:
            label = line_label(region, section)
            fc.cell(row, 1, label)
            vals: list[float] = []
            if section == "Operating Expenses":
                kind, param = OPEX_PLAN[region]
            else:
                kind, param = "growth", None
            for m in range(24):
                col = get_column_letter(2 + m)
                coord = f"{col}{row}"
                if m < N_ACT:
                    formula = (f"=SUMIFS({act_rng('F')},{act_rng('C')},$A{row},"
                               f"{act_rng('B')},{col}${HEADER_ROW})")
                    value = amounts[(label, m)]
                    put(fc, coord, formula, value)
                    fc[coord].fill = ACT_FILL
                else:
                    prev = get_column_letter(1 + m)
                    if kind == "growth":
                        value = vals[-1] * (1 + driver_value(label, drivers))
                        put(fc, coord, f"={prev}{row}*(1+Assumptions!$B${driver_row[label]})", value)
                    elif kind == "flat":
                        value = vals[-1]
                        put(fc, coord, f"={prev}{row}", value)
                    else:
                        value = param[m - N_ACT]
                        fc[coord] = value
                    fc[coord].fill = FC_FILL
                vals.append(value)
            put(fc, f"{get_column_letter(fy_col)}{row}", f"=SUM({fy_cols.format(r=row)})",
                sum(vals[FY_FIRST:]))
            series[label] = vals
            fc_row_of[label] = row
            row += 1
        last = row - 1
        fc.cell(row, 1, f"Total {section}").font = Font(bold=True)
        total = [sum(series[line_label(r[0], section)][m] for r in REGIONS) for m in range(24)]
        for m in range(24):
            col = get_column_letter(2 + m)
            put(fc, f"{col}{row}", f"=SUM({col}{first}:{col}{last})", total[m])
            fc[f"{col}{row}"].font = Font(bold=True)
        put(fc, f"{get_column_letter(fy_col)}{row}", f"=SUM({fy_cols.format(r=row)})", sum(total[FY_FIRST:]))
        total_rows[section], total_vals[section] = row, total
        row += 2

    def derived(label, formula_for, values, fmt=MONEY, bold=True):
        nonlocal row
        fc.cell(row, 1, label).font = Font(bold=bold)
        for m in range(24):
            col = get_column_letter(2 + m)
            put(fc, f"{col}{row}", formula_for(col), values[m])
            fc[f"{col}{row}"].number_format = fmt
            fc[f"{col}{row}"].font = Font(bold=bold)
        if fmt == MONEY:
            put(fc, f"{get_column_letter(fy_col)}{row}", f"=SUM({fy_cols.format(r=row)})",
                sum(values[FY_FIRST:]))
        else:
            cache_fy = None
        here = row
        row += 1
        return here

    rev, cogs, opex = (total_vals[s] for s in SECTIONS)
    rr, cr, orr = (total_rows[s] for s in SECTIONS)
    gm = [a - b for a, b in zip(rev, cogs)]
    gm_row = derived("Gross Margin", lambda c: f"={c}{rr}-{c}{cr}", gm)
    ebitda = [g - o for g, o in zip(gm, opex)]
    eb_row = derived("EBITDA", lambda c: f"={c}{gm_row}-{c}{orr}", ebitda)
    row += 1
    derived("Gross Margin %", lambda c: f"=IF({c}{rr}=0,0,{c}{gm_row}/{c}{rr})",
            [g / r if r else 0.0 for g, r in zip(gm, rev)], PCT, bold=False)
    derived("EBITDA %", lambda c: f"=IF({c}{rr}=0,0,{c}{eb_row}/{c}{rr})",
            [e / r if r else 0.0 for e, r in zip(ebitda, rev)], PCT, bold=False)
    forecast_last_row = row - 1

    for r in range(HEADER_ROW + 1, forecast_last_row + 1):
        for c in range(2, fy_col + 1):
            cell = fc.cell(r, c)
            if cell.number_format == "General":
                cell.number_format = MONEY
    fc.column_dimensions["A"].width = 30
    for c in range(2, fy_col + 1):
        fc.column_dimensions[get_column_letter(c)].width = 12
    fc.freeze_panes = fc.cell(HEADER_ROW + 1, 2)

    # ------------------------------------------------------------------ Plan (board budget)
    plan_ws.append(["Month", "Line Item", "Region", "Plan Amount"])
    plan_amounts: dict[str, float] = {}
    for m in range(FY_FIRST, 24):
        for section in SECTIONS:
            for region, *_ in REGIONS:
                label = line_label(region, section)
                value = round(series[label][m] * rng.uniform(0.93, 1.07), 2) if m >= N_ACT else \
                    round(series[label][m] * rng.uniform(0.95, 1.05), 2)
                plan_ws.append([MONTHS[m], label, region, value])
                plan_amounts[label] = round(plan_amounts.get(label, 0.0) + value, 2)
    plan_last = plan_ws.max_row
    for r in range(2, plan_last + 1):
        plan_ws.cell(r, 4).number_format = "#,##0.00"
    _table(plan_ws, "tblPlan", f"A1:D{plan_last}")
    for col, width in zip("ABCD", (10, 28, 15, 14)):
        plan_ws.column_dimensions[col].width = width

    # ------------------------------------------------------------------ Variance (vs board plan)
    var["A1"] = "FY27 forecast vs board plan"
    var["A1"].font = Font(bold=True, size=14, color=NAVY)
    for col, head in zip("ABCDE", ["Line", "FY27 Plan", "FY27 Forecast", "Variance", "Variance %"]):
        var[f"{col}3"] = head
        var[f"{col}3"].fill, var[f"{col}3"].font = HEAD_FILL, HEAD_FONT
    vrow = 4
    for label, frow in fc_row_of.items():
        var[f"A{vrow}"] = label
        plan_val = plan_amounts[label]
        fc_val = sum(series[label][FY_FIRST:])
        put(var, f"B{vrow}", f"=SUMIFS(Plan!$D$2:$D${plan_last},Plan!$B$2:$B${plan_last},$A{vrow})", plan_val)
        put(var, f"C{vrow}", f"=Forecast!{get_column_letter(fy_col)}{frow}", fc_val)
        put(var, f"D{vrow}", f"=C{vrow}-B{vrow}", fc_val - plan_val)
        put(var, f"E{vrow}", f"=IF(B{vrow}=0,0,D{vrow}/B{vrow})", (fc_val - plan_val) / plan_val if plan_val else 0.0)
        for col in "BCD":
            var[f"{col}{vrow}"].number_format = MONEY
        var[f"E{vrow}"].number_format = PCT
        vrow += 1
    var.column_dimensions["A"].width = 30
    for col in "BCDE":
        var.column_dimensions[col].width = 15

    # ------------------------------------------------------------------ Versions / Approvals
    ver.append(["Version", "Saved On", "Scenario", "Saved By", "Comment"])
    for row_values in [
        ("v1", datetime(2026, 11, 3), "Base", "Priya Raman", "First cut after Q3 close"),
        ("v2", datetime(2026, 11, 10), "Base", "Priya Raman", "NA pipeline refresh"),
        ("v3", datetime(2026, 11, 17), "Downside", "Priya Raman", "FX headwind case for CFO"),
        ("v4", datetime(2026, 11, 24), "Base", "Marcus Hale", "APAC hiring freeze applied"),
        ("v5", datetime(2026, 12, 8), "Base", "Marcus Hale", "Board pre-read"),
        ("v6", datetime(2026, 12, 15), "Upside", "Priya Raman", "Channel partner upside"),
        ("v7", datetime(2027, 1, 12), "Base", "Priya Raman", "FINAL (really) - actuals through Dec"),
    ]:
        ver.append(list(row_values))
    for r in range(2, ver.max_row + 1):
        ver.cell(r, 2).number_format = "yyyy-mm-dd"
    _table(ver, "tblVersions", f"A1:E{ver.max_row}")
    for col, width in zip("ABCDE", (9, 12, 12, 16, 40)):
        ver.column_dimensions[col].width = width

    apr.append(["Version", "Submitted By", "Approver", "Status", "Date", "Comment"])
    for row_values in [
        ("FY27 v3", "Priya Raman", "CFO", "Returned", datetime(2026, 11, 18), "Rework FX assumptions"),
        ("FY27 v5", "Marcus Hale", "CFO", "Approved", datetime(2026, 12, 9), "Approved for board"),
        ("FY27 v7", "Priya Raman", "CFO", "Pending", datetime(2027, 1, 13), ""),
    ]:
        apr.append(list(row_values))
    for r in range(2, apr.max_row + 1):
        apr.cell(r, 5).number_format = "yyyy-mm-dd"
    _table(apr, "tblApprovals", f"A1:F{apr.max_row}")
    for col, width in zip("ABCDEF", (10, 16, 10, 10, 12, 30)):
        apr.column_dimensions[col].width = width

    # ------------------------------------------------------------------ Instructions
    ins["A1"] = f"{COMPANY} - FY27 Rolling Forecast"
    ins["A1"].font = Font(bold=True, size=16, color=NAVY)
    ins["A2"] = "Owner: FP&A (Priya Raman). Version 7 FINAL. Enable macros when prompted."
    steps = [
        "How to use this model",
        "1. Click Refresh Actuals first - this reloads the Actuals tab from the GL warehouse.",
        "2. Pick a scenario on the Assumptions tab (Downside / Base / Upside); the rates update.",
        "3. Review the Forecast tab. Actuals are blue, forecast months are yellow.",
        "4. When a month closes, click Roll Forward Month to lock the first forecast month in as actuals.",
        "5. Click Snapshot Version to save a frozen copy to the Versions tab BEFORE changing drivers.",
        "6. Click Submit for Approval - this locks the Forecast sheet and emails the CFO.",
        "7. Check the Variance tab against the board plan before submitting.",
        "",
        "Do not insert rows or columns on the Forecast tab - the macros address cells by position.",
        "If numbers look stale, run RecalcAll from the Macros dialog (Alt+F8).",
    ]
    for i, text in enumerate(steps, 4):
        ins[f"A{i}"] = text
    ins["A4"].font = Font(bold=True, size=12, color=NAVY)
    _button(ins, "A17", "Refresh Actuals", 1)
    _button(ins, "B17", "Roll Forward Month", 1)
    _button(ins, "C17", "Snapshot Version", 1)
    _button(ins, "D17", "Submit for Approval", 1)
    ins.column_dimensions["A"].width = 100
    for col in "BCD":
        ins.column_dimensions[col].width = 22

    wb.save(path)
    _inject_cached_values(path, cache, wb.sheetnames)
    return path


def driver_value(label: str, drivers) -> float:
    return next(rate for name, _basis, rate in drivers if name == label)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", required=True, help="output directory (or a .xlsx path)")
    parser.add_argument("--seed", type=int, default=27)
    args = parser.parse_args(argv)
    path = build_demo_workbook(Path(args.out), seed=args.seed)
    print(f"wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
