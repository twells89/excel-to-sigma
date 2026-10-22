"""Synthetic table-report workbooks with Excel-style cached values.

openpyxl writes formulas without cached results, but the drafter reads Excel's cached
values. Each builder computes every formula cell's value in Python while writing it,
saves the workbook, then injects those values into the sheet XML (numbers as ``<v>``,
text as ``t="str"``), so the tests exercise the real ``load_workbook(data_only=True)``
path with no mocking. ``patch_cell`` rewrites one cell's cached value or formula in an
existing file (openpyxl cannot round-trip cached values) for failure tests.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.table import Table

DATE_FORMAT = "yyyy-mm-dd"


def _serial(value):
    return to_excel(datetime(value.year, value.month, value.day))


def _sheet_files(zf):
    workbook = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
    targets.update({m.group(2): m.group(1) for m in re.finditer(
        r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    out = {}
    for match in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        target = targets[match.group(2)].lstrip("/")
        out[match.group(1)] = target if target.startswith("xl/") else f"xl/{target}"
    return out


def inject_cached_values(path, cache):
    """``cache``: {(sheet, coord): number | str | date | bool}."""
    path = Path(path)
    temp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        files = {v: k for k, v in _sheet_files(source).items()}
        for item in source.infolist():
            data = source.read(item.filename)
            sheet = files.get(item.filename)
            if sheet is not None:
                def fill(found):
                    value = cache.get((sheet, found.group(1)))
                    if value is None and (sheet, found.group(1)) not in cache:
                        return found.group(0)
                    formula = found.group(3)
                    if isinstance(value, bool):
                        return f'<c r="{found.group(1)}"{found.group(2)} t="b">{formula}<v>{int(value)}</v></c>'
                    if isinstance(value, (date, datetime)):
                        value = _serial(value)
                    if isinstance(value, str) or value is None:
                        text = "" if value is None else escape(value)
                        return f'<c r="{found.group(1)}"{found.group(2)} t="str">{formula}<v>{text}</v></c>'
                    return f'<c r="{found.group(1)}"{found.group(2)}>{formula}<v>{value!r}</v></c>'

                data = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)>(<f>.*?</f>)<v ?/></c>', fill,
                              data.decode("utf-8")).encode("utf-8")
            target.writestr(item, data)
    shutil.move(temp, path)


def patch_cell(src, dst, sheet, coord, *, value=None, formula=None):
    """Copy ``src`` to ``dst`` changing one cell's cached value and/or formula text."""
    with zipfile.ZipFile(src) as source, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as target:
        files = _sheet_files(source)
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == files[sheet]:
                text = data.decode("utf-8")
                found = re.search(rf'<c r="{coord}"[^>]*?(?:/>|>.*?</c>)', text, re.S)
                if not found:
                    raise KeyError(f"{sheet}!{coord} not found in the sheet XML")
                cell = found.group(0)
                if formula is not None:
                    cell = re.sub(r"<f>.*?</f>", lambda m: f"<f>{escape(formula)}</f>", cell, flags=re.S)
                if value is not None:
                    if re.search(r"<v>.*?</v>|<v ?/>", cell, re.S):
                        cell = re.sub(r"<v>.*?</v>|<v ?/>", f"<v>{value!r}</v>", cell, flags=re.S)
                    else:
                        cell = cell.replace("</c>", f"<v>{value!r}</v></c>")
                text = text[:found.start()] + cell + text[found.end():]
                data = text.encode("utf-8")
            target.writestr(item, data)


class _Book:
    def __init__(self):
        self.wb = Workbook()
        self.first = True
        self.cache = {}

    def sheet(self, title):
        if self.first:
            ws = self.wb.active
            ws.title = title
            self.first = False
            return ws
        return self.wb.create_sheet(title)

    def put(self, ws, coord, formula, value, date_format=False):
        ws[coord] = formula
        self.cache[(ws.title, coord)] = value
        if date_format:
            ws[coord].number_format = DATE_FORMAT

    def table(self, ws, name, header, rows, first_row=1):
        for col, head in enumerate(header, 1):
            ws.cell(first_row, col, head)
        for r, row in enumerate(rows, first_row + 1):
            for col, value in enumerate(row, 1):
                if value is not None:
                    ws.cell(r, col, value)
                    if isinstance(value, (date, datetime)):
                        ws.cell(r, col).number_format = DATE_FORMAT
        ref = f"A{first_row}:{get_column_letter(len(header))}{first_row + len(rows)}"
        ws.add_table(Table(displayName=name, ref=ref))

    def save(self, path):
        path = Path(path)
        self.wb.save(path)
        inject_cached_values(path, self.cache)
        return path


def _eomonth(d):
    first_next = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return date.fromordinal(first_next.toordinal() - 1)


# ----------------------------------------------------------------------------- orders
ORDERS = [  # order, region, product, ship date, qty, price
    ("O1", "East", "Widget", date(2026, 1, 5), 2, 60.0),
    ("O2", "East", "Gadget", date(2026, 1, 20), 1, 150.0),
    ("O3", "East", "Widget", date(2026, 2, 3), 3, 70.0),
    ("O4", "West", "Gizmo", date(2026, 1, 11), 4, 40.0),
    ("O5", "West", "Widget", date(2026, 2, 14), 2, 90.0),
    ("O6", "West", "Gadget", date(2026, 2, 27), 5, 30.0),
    ("O7", "North", "Widget", date(2026, 1, 30), 1, 200.0),
    ("O8", "North", "Gizmo", date(2026, 2, 2), 6, 25.0),
]
RATES = {"Widget": 0.1, "Gadget": 0.2}
TARGETS = {"East": 400.0, "West": 0.0, "North": 300.0}


def build_orders(path, ghost=False, mixed=False, totals=None):
    """Tables with calculated columns + one summary grid (date windows, COUNTIFS, AVERAGEIFS...).

    ``ghost`` lists a region with no orders; ``mixed`` types a literal into a calculated column;
    ``totals`` adds a Total row (``"sum"`` = SUM of the rows above, ``"bad"`` = a different formula).
    """
    book = _Book()
    data = book.sheet("Orders")
    T = "tblOrders"
    book.table(data, T, ["Order", "Region", "Product", "Ship Date", "Qty", "Price", "Rate", "Revenue", "Bucket",
                         "Month End", "Flag", "Cap", "Boxes"],
               [list(r) + [None] * 7 for r in ORDERS])
    rates = book.sheet("Rates")
    book.table(rates, "tblRates", ["Product", "Rate", "Pct"], [[p, r, None] for p, r in RATES.items()])
    for i, (p, r) in enumerate(RATES.items(), 2):
        book.put(rates, f"C{i}", "=INDEX(tblRates[Rate],MATCH(tblRates[[#This Row],[Product]],tblRates[Product],0))*100",
                 r * 100)
    this = lambda c: f"{T}[[#This Row],[{c}]]"  # noqa: E731
    revenue = {}
    for i, (order, region, product, shipped, qty, price) in enumerate(ORDERS, 2):
        rate = RATES.get(product, 0)
        rev = qty * price * (1 - rate)
        revenue[order] = rev
        book.put(data, f"G{i}", f"=_xlfn.XLOOKUP({this('Product')},tblRates[Product],tblRates[Rate],0)", rate)
        book.put(data, f"H{i}", f"={this('Qty')}*{this('Price')}*(1-{this('Rate')})", rev)
        book.put(data, f"I{i}", f'=IF({this("Revenue")}>=100,"Big","Small")', "Big" if rev >= 100 else "Small")
        book.put(data, f"J{i}", f"=EOMONTH({this('Ship Date')},0)", _eomonth(shipped), date_format=True)
        book.put(data, f"K{i}", f'=IF(AND({this("Qty")}>1,MONTH({this("Ship Date")})>=2),"Y","N")',
                 "Y" if qty > 1 and shipped.month >= 2 else "N")
        book.put(data, f"L{i}", f"=MIN({this('Qty')},3)+MAX({this('Price')},50)", min(qty, 3) + max(price, 50))
        boxes = -(-qty // 5) * 5
        book.put(data, f"M{i}", f"=CEILING({this('Qty')},5)", boxes)
    summary = book.sheet("Summary")
    summary["A1"] = "Revenue by region"
    for col, head in enumerate(["Region", "Jan", "Feb", "Total", "Big", "Avg Rev", "Target", "Attain"], 1):
        summary.cell(3, col, head)
    targets = dict(TARGETS, **({"South": 100.0} if ghost else {}))
    columns_seen = {c: [] for c in "BCDEG"}
    for r, region in enumerate(targets, 4):
        summary.cell(r, 1, region)
        month_values = []
        for col, month in ((2, 1), (3, 2)):
            formula = (f"=SUMIFS({T}[Revenue],{T}[Region],$A{r},{T}[Ship Date],\">=\"&DATE(2026,{month},1),"
                       f"{T}[Ship Date],\"<=\"&EOMONTH(DATE(2026,{month},1),0))")
            value = sum(revenue[o[0]] for o in ORDERS if o[1] == region and o[3].month == month)
            month_values.append(value)
            book.put(summary, f"{get_column_letter(col)}{r}", formula, value)
            columns_seen[get_column_letter(col)].append(value)
        total = sum(month_values)
        columns_seen["D"].append(total)
        book.put(summary, f"D{r}", f"=SUM(B{r}:C{r})", total)
        rows = [o for o in ORDERS if o[1] == region]
        columns_seen["E"].append(sum(1 for o in rows if revenue[o[0]] >= 100))
        columns_seen["G"].append(targets[region])
        book.put(summary, f"E{r}", f'=COUNTIFS({T}[Region],$A{r},{T}[Bucket],"Big")',
                 sum(1 for o in rows if revenue[o[0]] >= 100))
        book.put(summary, f"F{r}", f"=AVERAGEIFS({T}[Revenue],{T}[Region],$A{r})",
                 sum(revenue[o[0]] for o in rows) / len(rows) if rows else 0)
        summary.cell(r, 7, targets[region])
        book.put(summary, f"H{r}", f"=IFERROR(D{r}/G{r},0)", total / targets[region] if targets[region] else 0)
    if totals:
        last = 3 + len(targets)
        summary.cell(last + 1, 1, "Total")
        for col, values in columns_seen.items():
            formula = f"=SUM({col}4:{col}{last})" if totals == "sum" else f"=SUM({col}5:{col}{last})"
            book.put(summary, f"{col}{last + 1}", formula, sum(values))
    if mixed:
        data["M3"] = 99  # a typed value among the calculated Boxes cells
    book.save(path)
    return {"revenue": revenue}


# ----------------------------------------------------------------------------- stock
ITEMS = {"A": 100, "B": 50}
WEEKS = [date(2026, 3, 2), date(2026, 3, 9), date(2026, 3, 16)]
FLOWS = {  # (item, week index) -> (in, out)
    ("A", 0): (0, 30), ("A", 1): (40, 90), ("A", 2): (0, 10),
    ("B", 0): (10, 20), ("B", 1): (0, 15), ("B", 2): (30, 5),
}


def build_stock(path, running=True, sparse=False):
    """Keyed grid: item x week, running balance (or a plain pick), MIN and a text flag.

    ``sparse`` omits one (item, week) row from the flows table (the grid cell is then 0 in Excel).
    """
    book = _Book()
    items = book.sheet("Items")
    book.table(items, "tblItem", ["Item", "Opening"], [[k, v] for k, v in ITEMS.items()])
    flows = book.sheet("Flows")
    book.table(flows, "tblFlow", ["Item", "Week", "In", "Out"],
               [[item, WEEKS[w], FLOWS[(item, w)][0], FLOWS[(item, w)][1]] for item in ITEMS for w in range(3)
                if not (sparse and (item, w) == ("B", 1))])
    bal = book.sheet("Balance")
    bal["A1"] = "Projected balance"
    bal.cell(3, 1, "Item")
    for c, week in enumerate(WEEKS, 2):
        bal.cell(3, c, week)
        bal.cell(3, c).number_format = DATE_FORMAT
    bal.cell(3, 5, "Low")
    bal.cell(3, 6, "Short?")
    for r, item in enumerate(ITEMS, 4):
        bal.cell(r, 1, item)
        balances, prev = [], None
        for w in range(3):
            col = get_column_letter(2 + w)
            sums = (f"SUMIFS(tblFlow[Out],tblFlow[Item],$A{r},tblFlow[Week],{col}$3)",
                    f"SUMIFS(tblFlow[In],tblFlow[Item],$A{r},tblFlow[Week],{col}$3)")
            inflow, outflow = (0, 0) if sparse and (item, w) == ("B", 1) else FLOWS[(item, w)]
            if running:
                if w == 0:
                    formula = f"=INDEX(tblItem[Opening],MATCH($A{r},tblItem[Item],0))-{sums[0]}+{sums[1]}"
                    value = ITEMS[item] - outflow + inflow
                else:
                    formula = f"={get_column_letter(1 + w)}{r}-{sums[0]}+{sums[1]}"
                    value = prev - outflow + inflow
            else:
                formula = f"={sums[1]}"
                value = inflow
            book.put(bal, f"{col}{r}", formula, value)
            balances.append(value)
            prev = value
        book.put(bal, f"E{r}", f"=MIN(B{r}:D{r})", min(balances))
        book.put(bal, f"F{r}", f'=IF(E{r}<20,"LOW","")', "LOW" if min(balances) < 20 else "")
    book.save(path)
    return {"balances": {item: None for item in ITEMS}}


# ----------------------------------------------------------------------------- staff
PEOPLE = [  # name, dept, salary, start
    ("Ann", "Eng", 120000.0, date(2026, 1, 1)),
    ("Bob", "Eng", 96000.0, date(2026, 2, 1)),
    ("Cy", "Ops", 60000.0, date(2026, 3, 1)),
    ("Di", "Ops", 72000.0, date(2025, 6, 1)),
]
MONTHS = [date(2026, 1, 1), date(2026, 2, 1), date(2026, 3, 1)]
UPLIFT = 0.1


def build_staff(path, short=False):
    """Ordinal expansion (people x months) with a parameter, a row total and a dept summary.

    ``short`` stops the ordinals one row early (the grid then misses a table row).
    """
    book = _Book()
    people = book.sheet("People")
    book.table(people, "tblPeople", ["Name", "Dept", "Salary", "Start"], [list(p) for p in PEOPLE])
    rates = book.sheet("Rates")
    rates["A1"] = "Uplift"
    rates["B1"] = UPLIFT
    cost = book.sheet("Cost")
    cost["A1"] = "Cost by person"
    for c, head in enumerate(["#", "Name", "Dept"], 1):
        cost.cell(3, c, head)
    for c, month in enumerate(MONTHS, 4):
        cost.cell(3, c, month)
        cost.cell(3, c).number_format = DATE_FORMAT
    cost.cell(3, 7, "Total")
    grid = {}
    for i, (name, dept, salary, start) in enumerate(PEOPLE[:-1] if short else PEOPLE, 1):
        r = 3 + i
        cost.cell(r, 1, i)
        book.put(cost, f"B{r}", f"=INDEX(tblPeople[Name],$A{r})", name)
        book.put(cost, f"C{r}", f"=INDEX(tblPeople[Dept],$A{r})", dept)
        values = []
        for c, month in enumerate(MONTHS, 4):
            col = get_column_letter(c)
            formula = (f"=IF(INDEX(tblPeople[Start],$A{r})<={col}$3,"
                       f"INDEX(tblPeople[Salary],$A{r})/12*(1+Rates!$B$1),0)")
            value = salary / 12 * (1 + UPLIFT) if start <= month else 0
            book.put(cost, f"{col}{r}", formula, value)
            values.append(value)
        book.put(cost, f"G{r}", f"=SUM(D{r}:F{r})", sum(values))
        grid[name] = values
    dept = book.sheet("Dept")
    dept["A1"] = "Cost by department"
    for c, head in enumerate(["Dept", "Cost", "Active"], 1):
        dept.cell(3, c, head)
    for r, d in enumerate(["Eng", "Ops"], 4):
        dept.cell(r, 1, d)
        members = [p[0] for p in PEOPLE if p[1] == d and p[0] in grid]
        book.put(dept, f"B{r}", f"=SUMIFS(Cost!$G:$G,Cost!$C:$C,$A{r})", sum(sum(grid[m]) for m in members))
        book.put(dept, f"C{r}", f'=COUNTIFS(Cost!$C:$C,$A{r},Cost!$F:$F,">0")',
                 sum(1 for m in members if grid[m][2] > 0))
    book.save(path)
    return {"grid": grid}


# ----------------------------------------------------------------------------- A1 table
LINES = [("food", 2, 5.0), ("food", 3, 4.0), ("gear", 1, 20.0)]


def build_a1_table(path):
    """A table whose calculated column uses plain same-row A1 references (``=B2*C2``)."""
    book = _Book()
    ws = book.sheet("Lines")
    book.table(ws, "tblLines", ["Kind", "Qty", "Price", "Amount"], [[k, q, p, None] for k, q, p in LINES])
    for r, (_, qty, price) in enumerate(LINES, 2):
        book.put(ws, f"D{r}", f"=B{r}*C{r}", qty * price)
    summary = book.sheet("Totals")
    for c, head in enumerate(["Kind", "Amount"], 1):
        summary.cell(3, c, head)
    for r, kind in enumerate(["food", "gear"], 4):
        summary.cell(r, 1, kind)
        book.put(summary, f"B{r}", f"=SUMIFS(tblLines[Amount],tblLines[Kind],$A{r})",
                 sum(q * p for k, q, p in LINES if k == kind))
    book.save(path)


# ----------------------------------------------------------------------------- cash ledger
CASH_AR = [("INV-1", 100.0, 1), ("INV-2", 200.0, 1), ("INV-3", 50.0, 3), ("INV-4", 400.0, 5), ("INV-5", 70.0, 5)]
CASH_AP = [("Acme", 30.0, 1), ("Acme", 20.0, 1), ("Bolt", 40.0, 4)]   # no unique vendor: a keyless table
CASH_RATES = {"Base": 0.9, "Low": 0.5, "High": 1.0}
CASH_RENT = [10, 0, 0, 0, 10]
CASH_OPENING, CASH_MIN, CASH_SEED, CASH_GROWTH, CASH_PAYROLL = 1000.0, 300.0, 80.0, 0.1, 60.0
CASH_FIRST_DAY = date(2026, 1, 2)
CASH_ROWS = ["Week Ending", "Beginning Cash", "Collections", "New Sales", "Total Receipts", "Payroll", "Rent",
             "Payments", "Total Out", "Net", "Ending Cash", "Cushion", "Lowest Ending"]


def cash_expected(scenario="Base", weeks=5):
    """Independent recomputation of the ledger (no drafter code) for ``scenario``."""
    out = {name: [] for name in CASH_ROWS}
    prev_end = None
    for i in range(weeks):
        week = i + 1
        collections = sum(a for _, a, w in CASH_AR if w == week) * CASH_RATES[scenario]
        new_sales = CASH_SEED * (1 + CASH_GROWTH) ** i
        receipts = collections + new_sales
        payroll = CASH_PAYROLL if week % 2 == 0 else 0
        payments = sum(a for _, a, w in CASH_AP if w == week)
        total_out = payroll + CASH_RENT[i] + payments
        net = receipts - total_out
        beginning = CASH_OPENING if prev_end is None else prev_end
        ending = beginning + net
        prev_end = ending
        row = {"Week Ending": date.fromordinal(CASH_FIRST_DAY.toordinal() + 7 * i), "Beginning Cash": beginning,
               "Collections": collections, "New Sales": new_sales, "Total Receipts": receipts,
               "Payroll": payroll, "Rent": CASH_RENT[i], "Payments": payments, "Total Out": total_out,
               "Net": net, "Ending Cash": ending, "Cushion": ending - CASH_MIN}
        for name, value in row.items():
            out[name].append(value)
        out["Lowest Ending"].append(min(out["Ending Cash"]))
    return out


def build_cash(path, scenario="Base", sections=True, selected=None, stray=False):
    """A 5-week cash ledger: SUMIFS rows x a scenario factor (INDIRECT over named cells), a growth recursion,
    a MOD cadence, typed rent, a date row from a date parameter, a carried running balance, a running MIN,
    SUM totals and a caption. ``Payments`` reads a keyless table (``Acme`` twice).

    ``selected`` types a different scenario into the selector cell than the one the values were computed
    for; ``stray`` adds a formula far below the ledger (outside any grid)."""
    from openpyxl.workbook.defined_name import DefinedName

    book = _Book()
    cash = book.sheet("Cash")
    inputs = book.sheet("Inputs")
    book.table(book.sheet("Invoices"), "tblIn", ["Invoice", "Amount", "Week"], [list(r) for r in CASH_AR])
    book.table(book.sheet("Bills"), "tblOut", ["Vendor", "Amount", "Pay Week"], [list(r) for r in CASH_AP])
    for r, (label, value) in enumerate([("Scenario", selected or scenario), ("Opening cash", CASH_OPENING),
                                        ("Minimum cash", CASH_MIN), ("Week 1 ends", CASH_FIRST_DAY),
                                        ("Seed receipts", CASH_SEED), ("Weekly growth", CASH_GROWTH),
                                        ("Payroll", CASH_PAYROLL)], 1):
        inputs.cell(r, 1, label)
        inputs.cell(r, 2, value)
    inputs["B4"].number_format = DATE_FORMAT
    inputs["A10"], inputs["B10"] = "Scenario", "Collection rate"
    for r, (name, rate) in enumerate(CASH_RATES.items(), 11):
        inputs.cell(r, 1, name)
        inputs.cell(r, 2, rate)
    names = {"Scenario": "Inputs!$B$1", "OpeningCash": "Inputs!$B$2", "MinCash": "Inputs!$B$3",
             "Rate_Base": "Inputs!$B$11", "Rate_Low": "Inputs!$B$12", "Rate_High": "Inputs!$B$13"}
    for name, ref in names.items():
        book.wb.defined_names[name] = DefinedName(name, attr_text=ref)
    weeks = 5
    expected = cash_expected(scenario, weeks)
    book.put(cash, "A1", '="Case: "&Scenario', f"Case: {scenario}")
    cash["A3"] = "Wk"
    for i in range(weeks):
        cash.cell(3, 2 + i, i + 1)
    cash.cell(3, 2 + weeks, "Total")
    letters = [get_column_letter(2 + i) for i in range(weeks)]
    rows, r = {}, 4
    for name in CASH_ROWS:
        if sections and name in ("Collections", "Payroll"):
            cash.cell(r, 1, "Receipts" if name == "Collections" else "Disbursements")
            r += 1
        rows[name] = r
        cash.cell(r, 1, name)
        r += 1
    rr = rows
    for i, col in enumerate(letters):
        prev = letters[i - 1] if i else None
        formulas = {
            "Week Ending": ("=Inputs!$B$4" if i == 0 else f"={prev}{rr['Week Ending']}+7"),
            "Beginning Cash": ("=OpeningCash" if i == 0 else f"={prev}{rr['Ending Cash']}"),
            "Collections": f"=SUMIFS(tblIn[Amount],tblIn[Week],{col}$3)*INDIRECT(\"Rate_\"&Scenario)",
            "New Sales": ("=Inputs!$B$5" if i == 0 else f"={prev}{rr['New Sales']}*(1+Inputs!$B$6)"),
            "Total Receipts": f"={col}{rr['Collections']}+{col}{rr['New Sales']}",
            "Payroll": f"=IF(MOD({col}$3,2)=0,Inputs!$B$7,0)",
            "Payments": f"=SUMIFS(tblOut[Amount],tblOut[Pay Week],{col}$3)",
            "Total Out": f"=SUM({col}{rr['Payroll']}:{col}{rr['Payments']})",
            "Net": f"={col}{rr['Total Receipts']}-{col}{rr['Total Out']}",
            "Ending Cash": f"={col}{rr['Beginning Cash']}+{col}{rr['Net']}",
            "Cushion": f"={col}{rr['Ending Cash']}-MinCash",
            "Lowest Ending": f"=MIN($B${rr['Ending Cash']}:{col}{rr['Ending Cash']})",
        }
        for name, formula in formulas.items():
            book.put(cash, f"{col}{rr[name]}", formula, expected[name][i], date_format=name == "Week Ending")
        cash[f"{col}{rr['Rent']}"] = CASH_RENT[i]
    total_col = get_column_letter(2 + weeks)
    for name in ("Collections", "New Sales", "Total Receipts", "Payroll", "Rent", "Payments", "Total Out", "Net"):
        book.put(cash, f"{total_col}{rr[name]}", f"=SUM(B{rr[name]}:{letters[-1]}{rr[name]})", sum(expected[name]))
    if stray:
        cash["A30"] = "Scratch"
        book.put(cash, "B30", "=B5+1", CASH_OPENING + 1)
    book.save(path)
    return {"rows": rr, "expected": expected}
