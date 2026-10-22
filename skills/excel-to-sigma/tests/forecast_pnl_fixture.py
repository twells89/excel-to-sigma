"""Synthetic cell-drawn P&L workbook (SUMIFS actuals + driver growth + flat + manual).

openpyxl writes formulas without cached values, but the drafter reads Excel's
cached results. ``build_pnl`` therefore computes every formula cell's value in
Python while writing it, saves the workbook, then injects those values as ``<v>``
elements into the saved sheet XML. The result is a normal workbook whose cached
values are exactly what Excel would have stored, so the tests exercise the real
``load_workbook(data_only=True)`` path with no mocking.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

MONTHS = ["Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26"]
N_ACT = 3          # Jan-Mar actual, Apr-Jun forecast
HEADER_ROW = 3

# (section, label, kind, parameter, actual amounts per actual month)
DEFAULT_LINES = [
    ("Revenue", "Subscription", "growth", 0.05, [1000.0, 1100.0, 1200.0]),
    ("Revenue", "Services", "manual", [300.0, 320.0, 340.0], [250.0, 260.0, 270.0]),
    ("Cost of Revenue", "Hosting", "flat", None, [200.0, 210.0, 220.0]),
    ("Cost of Revenue", "Support", "growth", 0.02, [100.0, 105.0, 110.0]),
    ("Operating Expenses", "R&D", "manual", [400.0, 410.0, 420.0], [380.0, 385.0, 390.0]),
    ("Below the Line", "Interest", "flat", None, [20.0, 20.0, 21.0]),
]
SECTION_ORDER = ["Revenue", "Cost of Revenue", "Operating Expenses", "Below the Line"]


def _inject_cached_values(path: Path, cache: dict[tuple[str, str], float], sheet_names: list[str]):
    temp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            match = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if match:
                sheet = sheet_names[int(match.group(1)) - 1]

                def fill(found):
                    value = cache.get((sheet, found.group(2)))
                    if value is None:
                        return found.group(0)
                    return f"{found.group(1)}{found.group(3)}<v>{value!r}</v>"

                data = re.sub(r'(<c r="([A-Z]+\d+)"[^>]*>)(<f>.*?</f>)<v />',
                              fill, data.decode("utf-8")).encode("utf-8")
            target.writestr(item, data)
    shutil.move(temp, path)


def build_pnl(path, *, lines=None, mutate=None, cached_overrides=None, gl_sheet="GL"):
    """Write the fixture to ``path``; returns the expected values for assertions.

    ``mutate(pnl, gl, drivers)`` edits cells after the clean layout is written (cached
    values stay as computed for the clean layout). ``cached_overrides`` replaces cached
    values by ``(sheet, coord)`` to simulate a workbook whose caches disagree.
    """
    path = Path(path)
    lines = lines or DEFAULT_LINES
    wb = Workbook()
    pnl = wb.active
    pnl.title = "P&L"
    gl = wb.create_sheet(gl_sheet)
    drivers = wb.create_sheet("Drivers")
    cache: dict[tuple[str, str], float] = {}

    def put(ws, coord, formula, value):
        ws[coord] = formula
        cache[(ws.title, coord)] = float(value)

    # ---- drivers
    drivers["A1"] = "Forecast drivers"
    drivers["A3"] = "Driver"
    drivers["B3"] = "Monthly Rate"
    drivers["C3"] = "Basis"
    driver_row = {}
    for label, kind, param in ((l[1], l[2], l[3]) for l in lines):
        if kind == "growth":
            row = 4 + len(driver_row)
            driver_row[label] = row
            drivers.cell(row, 1, f"{label} growth")
            drivers.cell(row, 2, param)
            drivers.cell(row, 3, f"{label} basis")

    # ---- GL: two postings per actual line-month
    gl["A1"] = "GL detail"
    for col, head in enumerate(["Txn Date", "Month", "Line Item", "Department", "Memo / Vendor", "Amount"], 1):
        gl.cell(4, col, head)
    row = 5
    actual = {}
    for section, label, kind, param, amounts in lines:
        for month_index, amount in enumerate(amounts):
            half = round(amount / 2, 2)
            for part, value in enumerate((half, amount - half)):
                gl.cell(row, 1, datetime(2026, month_index + 1, 5 + part))
                gl.cell(row, 2, MONTHS[month_index])
                gl.cell(row, 3, label)
                gl.cell(row, 4, section)
                gl.cell(row, 5, f"vendor {part}")
                gl.cell(row, 6, value)
                row += 1
            actual[(label, month_index)] = amount
    gl_last = row - 1

    # ---- P&L
    pnl["A1"] = "Synthetic Co - rolling forecast"
    for index, month in enumerate(MONTHS):
        pnl.cell(HEADER_ROW, 2 + index, month)
    fy_col = 2 + len(MONTHS)
    pnl.cell(HEADER_ROW, fy_col, "FY2026")
    values = {}
    totals = {}
    row = HEADER_ROW + 1
    section_rows = {}
    for section in SECTION_ORDER:
        members = [l for l in lines if l[0] == section]
        if not members:
            continue
        pnl.cell(row, 1, section)
        row += 1
        first = row
        for _, label, kind, param, amounts in members:
            pnl.cell(row, 1, label)
            series = []
            for month_index in range(len(MONTHS)):
                col = get_column_letter(2 + month_index)
                coord = f"{col}{row}"
                if month_index < N_ACT:
                    formula = (f"=SUMIFS({gl_sheet}!$F$5:$F${gl_last},{gl_sheet}!$C$5:$C${gl_last},"
                               f"$A{row},{gl_sheet}!$B$5:$B${gl_last},{col}${HEADER_ROW})")
                    value = actual[(label, month_index)]
                    put(pnl, coord, formula, value)
                else:
                    prev = get_column_letter(1 + month_index)
                    if kind == "growth":
                        value = series[-1] * (1 + param)
                        put(pnl, coord, f"={prev}{row}*(1+Drivers!$B${driver_row[label]})", value)
                    elif kind == "flat":
                        value = series[-1]
                        put(pnl, coord, f"={prev}{row}", value)
                    else:
                        value = param[month_index - N_ACT]
                        pnl[coord] = value
                series.append(value)
            put(pnl, f"{get_column_letter(fy_col)}{row}",
                f"=SUM(B{row}:{get_column_letter(fy_col - 1)}{row})", sum(series))
            values[label] = series
            row += 1
        last = row - 1
        total = [sum(values[l[1]][m] for l in members) for m in range(len(MONTHS))]
        pnl.cell(row, 1, f"Total {section}")
        for month_index in range(len(MONTHS)):
            col = get_column_letter(2 + month_index)
            put(pnl, f"{col}{row}", f"=SUM({col}{first}:{col}{last})", total[month_index])
        put(pnl, f"{get_column_letter(fy_col)}{row}", f"=SUM(B{row}:{get_column_letter(fy_col - 1)}{row})", sum(total))
        totals[section] = (row, total)
        row += 2

    def derived(label, expr_rows, signs, series):
        nonlocal row
        pnl.cell(row, 1, label)
        for month_index in range(len(MONTHS)):
            col = get_column_letter(2 + month_index)
            formula = "=" + "".join(f"{s}{col}{r}" for s, r in zip(signs, expr_rows)).lstrip("+")
            put(pnl, f"{col}{row}", formula, series[month_index])
        put(pnl, f"{get_column_letter(fy_col)}{row}", f"=SUM(B{row}:{get_column_letter(fy_col - 1)}{row})", sum(series))
        row_here = row
        row += 1
        return row_here

    zero = [0.0] * len(MONTHS)
    rev = totals["Revenue"]
    cogs = totals.get("Cost of Revenue", (None, zero))
    opex = totals.get("Operating Expenses", (None, zero))
    below = totals.get("Below the Line", (None, zero))
    gp = [r - c for r, c in zip(rev[1], cogs[1])]
    gp_row = derived("Gross Profit", [rev[0], cogs[0]], ["+", "-"], gp)
    ebitda = [g - o for g, o in zip(gp, opex[1])]
    ebitda_row = derived("EBITDA", [gp_row, opex[0]], ["+", "-"], ebitda)
    net = [e - b for e, b in zip(ebitda, below[1])]
    derived("Net Income", [ebitda_row, below[0]], ["+", "-"], net)
    expected = {"Revenue": sum(rev[1]), "Gross Profit": sum(gp), "EBITDA": sum(ebitda),
                "Net Income": sum(net)}

    if mutate:
        mutate(pnl, gl, drivers)
    for key, value in (cached_overrides or {}).items():
        cache[key] = float(value)
    wb.save(path)
    _inject_cached_values(path, cache, wb.sheetnames)
    return {"targets": expected, "values": values, "gl_last": gl_last, "driver_row": driver_row}
