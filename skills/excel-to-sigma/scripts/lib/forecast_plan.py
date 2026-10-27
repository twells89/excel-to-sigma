"""Draft a ``build-forecast-model.py`` plan from a cell-drawn P&L workbook.

The drafter reads the report sheet's formulas and Excel's cached values, the GL
actuals tab its SUMIFS read, and the driver tab its growth formulas read. It then
re-derives every line-period, section total and derived row from the drafted plan
and compares each with the workbook's cached values, so a plan that leaves this
module already reproduces the spreadsheet.

Sections are discovered from the sheet (any names, any count); each gets a sign role
read from the derived rows' formulas (``Gross Margin = Net Sales - COGS``) and only
falls back to a name table when no derived row decides it. Periods are months or
quarters. Forecast cells may be driver growth, year-over-year growth, a percentage of
a section total, a carried-forward value, a typed literal, or an external-link value
(kept as the cached number and flagged for review).

Every precondition fails loudly with ``PlanDraftError`` and a precise cell
reference. An unknown shape is never coerced (SKILL.md Phase 0c).
"""
from __future__ import annotations

import calendar
import importlib.util
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

try:
    from .migration_gate import source_sha256
    from . import forecast_calc as calc
    from . import forecast_presentation as fpres
except ImportError:  # direct import when scripts/lib itself is on sys.path
    from migration_gate import source_sha256
    import forecast_calc as calc
    import forecast_presentation as fpres

SECTIONS = ("Revenue", "Cost of Revenue", "Operating Expenses", "Below the Line")
LEGACY_SIGNS = {"Revenue": 1, "Cost of Revenue": -1, "Operating Expenses": -1, "Below the Line": -1}
TARGET_LABELS = calc.TARGET_LABELS
INTENTS = ("read-only", "editable", "what-if")
MONTH_ABBR = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")

MONTH_RE = re.compile(r"^\s*([A-Za-z]{3,9})[-\s/']*(\d{2}|\d{4})\s*$")
QUARTER_PATTERNS = (
    re.compile(r"^\s*Q([1-4])[\s\-'/]*(FY)?[\s\-'/]*(\d{4}|\d{2})\b(.*)$", re.I),
    re.compile(r"^\s*(FY)?[\s\-']*(\d{4}|\d{2})[\s\-']*Q([1-4])\b(.*)$", re.I),
)
WEEK_ORDINAL_RE = re.compile(r"^\s*(?:wk|week|w)[\s\-#]*(\d{1,2})\s*$", re.I)
SUMIFS_RE = re.compile(r"\bSUMIFS\b", re.I)
SHEET_REF_RE = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!")
UNQUOTED_SPACE_RE = re.compile(r"\s+(?=(?:[^']*'[^']*')*[^']*$)")
FY_HEADER_RE = re.compile(
    r"^(?:fy(?:\s*\d{2,4})?(?:\s+(?:total|year))?|total|full[ -]?year|annual|year)$", re.I)
ZERO_FORMULA_RE = re.compile(r"^=\(?0+(?:\.0*)?\)?$")

_DRV = r"(?:'(?P<s1>(?:[^']|'')+)'|(?P<s2>[A-Za-z0-9_.]+))!(?P<dc>[A-Z]{1,3})(?P<dr>\d+)"
_GROWTH_FORMS = (
    re.compile(rf"^=(?P<rc>[A-Z]{{1,3}})(?P<rr>\d+)\*\(1\+{_DRV}\)$"),
    re.compile(rf"^=(?P<rc>[A-Z]{{1,3}})(?P<rr>\d+)\*\({_DRV}\+1\)$"),
    re.compile(rf"^=(?P<rc>[A-Z]{{1,3}})(?P<rr>\d+)\+(?P=rc)(?P=rr)\*{_DRV}$"),
)
_PCT_FORMS = (
    re.compile(rf"^=(?P<rc>[A-Z]{{1,3}})(?P<rr>\d+)\*{_DRV}$"),
    re.compile(rf"^={_DRV}\*(?P<rc>[A-Z]{{1,3}})(?P<rr>\d+)$"),
)
FLAT_RE = re.compile(r"^=(?P<rc>[A-Z]{1,3})(?P<rr>\d+)$")
EXTERNAL_RE = re.compile(r"^='?\[(?P<book>[^\]]+)\][^'!]*'?!(?P<c>[A-Z]{1,3})(?P<r>\d+)$")

POSITIVE_WORDS = ("revenue", "sales", "income", "gain")
NEGATIVE_WORDS = ("cost", "cogs", "expense", "opex", "sg&a", "sga", "overhead", "payroll", "tax")


class PlanDraftError(RuntimeError):
    """The workbook does not have a shape this drafter supports."""


_WHERE = {"cell": None}


def _at(where):
    """Remember the cell being read, so an unexpected failure can name it."""
    _WHERE["cell"] = where


# ---------------------------------------------------------------- helpers
def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _text(value):
    return value.strip() if isinstance(value, str) and value.strip() else None


def _formula(value):
    return value if isinstance(value, str) and value.startswith("=") else None


def parse_month(value):
    """Return ``(year, month)`` for a month header (datetime or ``Jan-26`` text), else None."""
    if isinstance(value, (datetime, date)):
        return value.year, value.month
    if isinstance(value, str):
        found = MONTH_RE.match(value)
        if found and found.group(1)[:3].lower() in MONTH_ABBR:
            year = int(found.group(2))
            return (year + 2000 if year < 100 else year), MONTH_ABBR.index(found.group(1)[:3].lower()) + 1
    return None


def parse_quarter(value):
    """``(year, quarter, fiscal)`` for ``Q3 FY25`` / ``Q2 2026`` / ``2026 Q1`` text, else None.

    A trailing word (``Q3 FY26 Bud``) is allowed; ``fiscal`` is True for FY labels, whose
    calendar end date is not knowable from the label alone.
    """
    if not isinstance(value, str):
        return None
    for index, pattern in enumerate(QUARTER_PATTERNS):
        found = pattern.match(value)
        if found:
            if index == 0:
                quarter, fy, year, rest = found.group(1), found.group(2), found.group(3), found.group(4)
            else:
                fy, year, quarter, rest = found.group(1), found.group(2), found.group(3), found.group(4)
            if rest.strip() and not re.fullmatch(r"[\s\-_]*[A-Za-z][A-Za-z ]*", rest):
                return None
            year = int(year)
            return (year + 2000 if year < 100 else year), int(quarter), bool(fy)
    return None


def month_meta(year_month, num):
    year, month = year_month
    end = date(year, month, calendar.monthrange(year, month)[1])
    return {"month": f"{MONTH_ABBR[month - 1].capitalize()}-{year % 100:02d}", "num": num,
            "end": end.isoformat()}


def _ref(ws, col, row):
    return f"{ws.title}!{get_column_letter(col)}{row}"


def _load_discover():
    """Import xlsx-discover.py so report-sheet selection matches discovery exactly."""
    scripts = Path(__file__).resolve().parents[1]
    if str(scripts) not in sys.path:
        sys.path.insert(0, str(scripts))
    spec = importlib.util.spec_from_file_location("xlsx_discover_for_forecast_plan",
                                                  scripts / "xlsx-discover.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---------------------------------------------------------------- locating
def find_report_sheet(wf, sheet=None):
    if sheet:
        if sheet not in wf.sheetnames:
            raise PlanDraftError(f"--sheet {sheet!r} not found; sheets: {', '.join(wf.sheetnames)}")
        return wf[sheet]
    sheetmap, _, _ = _load_discover().classify_report_cells(wf)
    scored = sorted(((counts.get("ACTUAL", 0) + counts.get("DRIVER_GROWTH", 0), name)
                     for name, counts in sheetmap.items()), reverse=True)
    if not scored or scored[0][0] == 0:
        raise PlanDraftError(
            "no sheet has SUMIFS actuals or driver-growth formulas; this is not an "
            "actuals + driver forecast model (pass --sheet if you know the report sheet)")
    if len(scored) > 1 and scored[1][0] == scored[0][0]:
        raise PlanDraftError(
            f"sheets {scored[0][1]!r} and {scored[1][1]!r} score equally as the report sheet; "
            "pass --sheet")
    return wf[scored[0][1]]


def _header_period(value):
    """``(grain, order_key, label, end_iso_or_None)`` for a period header cell, else None."""
    month = parse_month(value)
    if month:
        meta = month_meta(month, 0)
        return "month", month, meta["month"], meta["end"]
    quarter = parse_quarter(value)
    if quarter:
        year, number, fiscal = quarter
        end = None if fiscal else date(year, number * 3, calendar.monthrange(year, number * 3)[1]).isoformat()
        return "quarter", (year, number), str(value).strip(), end
    if isinstance(value, str):
        week = WEEK_ORDINAL_RE.match(value)
        if week:
            return "week", (0, int(week.group(1))), value.strip(), None
    return None


def _week_date_run(ws, row):
    """Header cells that are dates exactly seven days apart (ISO-style week starts)."""
    run, best = [], []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row, col).value
        stamp = value.date() if isinstance(value, datetime) else value if isinstance(value, date) else None
        if stamp is not None and run and (stamp - run[-1][1]).days == 7:
            run.append((col, stamp))
            continue
        best = max(best, run, key=len)
        run = [(col, stamp)] if stamp is not None else []
    best = max(best, run, key=len)
    if len(best) < 3:
        return None
    return [{"col": col, "grain": "week", "key": (0, i + 1), "label": stamp.isoformat(),
             "end": date.fromordinal(stamp.toordinal() + 6).isoformat()}
            for i, (col, stamp) in enumerate(best)]


def find_period_header(ws):
    """First row with >=3 consecutive month (or quarter) labels.

    Returns ``(row, grain, [{"col", "key", "label", "end"}, ...])``.
    """
    for row in range(1, ws.max_row + 1):
        weeks = _week_date_run(ws, row)
        if weeks:
            return row, "week", weeks
        run, best = [], []
        for col in range(1, ws.max_column + 1):
            parsed = _header_period(ws.cell(row, col).value)
            if parsed and (not run or run[-1]["grain"] == parsed[0]):
                run.append({"col": col, "grain": parsed[0], "key": parsed[1], "label": parsed[2],
                            "end": parsed[3]})
                continue
            best = max(best, run, key=len)
            run = []
            if parsed:
                run.append({"col": col, "grain": parsed[0], "key": parsed[1], "label": parsed[2],
                            "end": parsed[3]})
        best = max(best, run, key=len)
        if len(best) < 3:
            continue
        grain = best[0]["grain"]
        for prev, cur in zip(best, best[1:]):
            if grain == "month":
                expected = (prev["key"][0] + (prev["key"][1] == 12), prev["key"][1] % 12 + 1)
                what = "month headers are not consecutive calendar months"
                shown = lambda k: f"{k[0]}-{k[1]:02d}"
            elif grain == "week":
                expected = (0, prev["key"][1] + 1)
                what = "week headers are not consecutive weeks"
                shown = lambda k: f"week {k[1]}"
            else:
                expected = (prev["key"][0] + (prev["key"][1] == 4), prev["key"][1] % 4 + 1)
                what = "quarter headers are not consecutive quarters"
                shown = lambda k: f"{k[0]}-Q{k[1]}"
            if cur["key"] != expected:
                raise PlanDraftError(
                    f"{_ref(ws, cur['col'], row)}: {what} "
                    f"({shown(prev['key'])} then {shown(cur['key'])})")
        return row, grain, best
    raise PlanDraftError(
        f"{ws.title!r}: no header row with 3 or more consecutive month, quarter or week labels")


def _walk_rows(ws, header_row, period_cols, fy_col):
    """Yield ('section', row, label) / ('data', row, label) in sheet order."""
    value_cols = list(period_cols) + ([fy_col] if fy_col else [])
    rows = []
    for row in range(header_row + 1, ws.max_row + 1):
        label = ws.cell(row, 1).value
        has_values = any(ws.cell(row, col).value is not None for col in value_cols)
        if label is None and not has_values:
            continue
        if label is not None and not _text(label) and not has_values:
            continue
        rows.append((row, label, has_values))
    for index, (row, label, has_values) in enumerate(rows):
        if has_values:
            yield "data", row, label
        elif _text(label) and index + 1 < len(rows) and rows[index + 1][2]:
            yield "section", row, _text(label)
        # else: a note / legend row with nothing under it


def _compact(formula):
    """Drop whitespace outside quoted sheet names and ``$`` anchors."""
    return UNQUOTED_SPACE_RE.sub("", formula).replace("$", "")


def _forecast_cell_kind(ws, col, row, value):
    """Return (kind, detail) for one forecast cell, or (None, repr) when unrecognised.

    Kinds: growth ``(sheet, col, row)``, yoy ``(sheet, col, row, lag)``, flat,
    pct_of_total ``(sheet, col, row, total_row)``, external ``(link text)``, manual.
    """
    formula = _formula(value)
    if formula:
        compact = _compact(formula)
        this_col = col
        for form in _GROWTH_FORMS:
            grown = form.match(compact)
            if grown and int(grown.group("rr")) == row:
                lag = this_col - column_index_from_string(grown.group("rc"))
                sheet = (grown.group("s1") or grown.group("s2")).replace("''", "'")
                target = (sheet, grown.group("dc"), int(grown.group("dr")))
                if lag == 1:
                    return "growth", target
                if lag >= 2:
                    return "yoy", target + (lag,)
        flat = FLAT_RE.match(compact)
        if flat and column_index_from_string(flat.group("rc")) == this_col - 1 \
                and int(flat.group("rr")) == row:
            return "flat", None
        for form in _PCT_FORMS:
            share = form.match(compact)
            if (share and column_index_from_string(share.group("rc")) == this_col
                    and int(share.group("rr")) != row):
                sheet = (share.group("s1") or share.group("s2")).replace("''", "'")
                return "pct_of_total", (sheet, share.group("dc"), int(share.group("dr")),
                                        int(share.group("rr")))
        if EXTERNAL_RE.match(compact):
            return "external", formula
        return None, repr(value)
    if _is_number(value):
        return "manual", None
    return None, repr(value)


# ---------------------------------------------------------------- row classification
def _classify_row(ws, row, text, period_cols):
    """``"line"`` (actuals + forecast) or ``"derived"`` (same-column total / margin / profit)."""
    first = ws.cell(row, period_cols[0]).value
    where = _ref(ws, period_cols[0], row)
    _at(f"{where} ({text!r})")
    neither = (f"{where} ({text!r}): row is neither a SUMIFS actuals line nor a recognised "
               "same-column total/derived row")
    if _is_number(first):
        return "line"
    formula = _formula(first)
    if not formula:
        raise PlanDraftError(f"{neither} (first period cell is {first!r})")
    if SUMIFS_RE.search(formula):
        return "line"
    try:
        tree = calc.parse_formula(formula)
    except calc.ExprError as error:
        raise PlanDraftError(f"{neither} ({formula!r}: {error})") from None
    refs = _refs(tree)
    if not refs:
        return "line"  # =0 style literal
    here = get_column_letter(period_cols[0])
    for sheet, col, _ in refs:
        if sheet not in (None, ws.title) or col != here:
            raise PlanDraftError(f"{neither} ({formula!r} reads outside its own column)")
    return "derived"


def _refs(tree):
    kind = tree[0]
    if kind == "ref":
        return [(tree[1], tree[2], tree[3])]
    if kind == "range":
        return [(tree[1], tree[2], tree[3]), (tree[1], tree[4], tree[5])]
    if kind in ("bin", "cmp"):
        return _refs(tree[2]) + _refs(tree[3])
    if kind == "neg":
        return _refs(tree[1])
    if kind == "call":
        return [r for arg in tree[2] for r in _refs(arg)]
    return []


def _flat_sum_rows(tree):
    """Rows summed by ``SUM(range, ...)`` or ``a+b+c`` of plain references; else None."""
    if tree[0] == "call" and tree[1] == "SUM":
        rows = []
        for arg in tree[2]:
            if arg[0] == "ref":
                rows.append(arg[3])
            elif arg[0] == "range" and arg[2] == arg[4]:
                rows.extend(range(min(arg[3], arg[5]), max(arg[3], arg[5]) + 1))
            else:
                return None
        return rows
    if tree[0] == "ref":
        return [tree[3]]
    if tree[0] == "bin" and tree[1] == "+":
        left, right = _flat_sum_rows(tree[2]), _flat_sum_rows(tree[3])
        return None if left is None or right is None else left + right
    return None


class _Resolver:
    def __init__(self, row_entity):
        self.row_entity = row_entity

    def __call__(self, sheet, col, row):
        if row not in self.row_entity:
            raise calc.ExprError(f"row {row} is not a P&L line, section total or derived row")
        return list(self.row_entity[row])

    def range(self, sheet, col, row_a, row_b):
        lo, hi = min(row_a, row_b), max(row_a, row_b)
        return [list(self.row_entity[r]) for r in range(lo, hi + 1) if r in self.row_entity]


def _derive_rows(ws, vs, period_cols, line_rows, sections, derived_candidates):
    """Split derived candidates into section totals and generic derived rows.

    Returns ``(totals, derived)``: ``totals[section] = (row, label)`` and ``derived`` a
    list of plan derived-row dicts (without cached values).
    """
    line_by_row = {row: text for row, text, _ in line_rows}
    section_rows = {name: [row for row, _, sec in line_rows if sec == name] for name in sections}
    trees = {}
    for row, text in derived_candidates:
        shapes = {}
        for col in period_cols:
            cell = ws.cell(row, col).value
            formula = _formula(cell)
            _at(f"{_ref(ws, col, row)} ({text!r})")
            if not formula:
                raise PlanDraftError(
                    f"{_ref(ws, col, row)} ({text!r}): derived row has {cell!r} where a formula "
                    "is expected in every period")
            try:
                tree = calc.parse_formula(formula)
            except calc.ExprError as error:
                raise PlanDraftError(f"{_ref(ws, col, row)} ({text!r}): {formula!r}: {error}") from None
            shapes.setdefault(calc.relative_shape(tree, col), (col, formula))
            if col == period_cols[0]:
                trees[row] = tree
        if len(shapes) > 1:
            (_, first_formula), (second_col, second_formula) = list(shapes.values())[:2]
            raise PlanDraftError(
                f"{_ref(ws, second_col, row)} ({text!r}): formula {second_formula!r} differs from "
                f"{first_formula!r} in the first period; derived rows must use one formula")
    totals, other = {}, []
    for row, text in derived_candidates:
        rows = _flat_sum_rows(trees[row])
        owner = None
        if rows and all(r in line_by_row for r in rows):
            owner = next(sec for r_, _, sec in line_rows if r_ == rows[0])
            if sorted(rows) != sorted(section_rows[owner]) or owner in totals:
                owner = None
        if owner:
            totals[owner] = (row, text)
        else:
            other.append((row, text))
    names = Counter(text for _, text in other)
    dup = [n for n, c in names.items() if c > 1]
    if dup:
        raise PlanDraftError(f"duplicate derived-row labels {dup}: plan rows are keyed by label")
    row_entity = {row: ["line", text] for row, text in line_by_row.items()}
    row_entity.update({row: ["sec", sec] for sec, (row, _) in totals.items()})
    row_entity.update({row: ["drv", text] for row, text in other})
    resolver = _Resolver(row_entity)
    derived = []
    for row, text in other:
        try:
            expr = calc.to_entities(trees[row], resolver)
        except calc.ExprError as error:
            raise PlanDraftError(f"{_ref(ws, period_cols[0], row)} ({text!r}): "
                                 f"{ws.cell(row, period_cols[0]).value!r}: {error}") from None
        derived.append({"label": text, "sourceRow": row, "expr": expr})
    exprs = {d["label"]: d["expr"] for d in derived}
    section_lines = {sec: [text for _, text, owner in line_rows if owner == sec] for sec in sections}
    row_of = {text: row for row, text, _ in line_rows}
    clean = lambda v: int(v) if float(v).is_integer() else v
    for d in derived:
        terms = calc.linear_terms(d["expr"], exprs)
        section_coeffs, line_coeffs = calc.fold_terms(terms, section_lines) if terms else ({}, {})
        d["kind"] = "amount" if (section_coeffs or line_coeffs) else "ratio"
        d["formula"] = calc.render_expr(d["expr"])
        if d["kind"] == "amount":
            d["coefficients"] = {k: clean(v) for k, v in section_coeffs.items()}
            if line_coeffs:
                # Individual lines (``=Subscription - COGS``) enter the KPIs, the trend and the
                # charts as line-level terms, never widened to their whole section.
                for label in line_coeffs:
                    if re.search(r'["\\]', label):
                        raise PlanDraftError(
                            f"{_ref(ws, period_cols[0], d['sourceRow'])} ({d['label']!r}): reads line "
                            f"{label!r} (row {row_of[label]}), whose name contains a quote or backslash "
                            "the Sigma formulas cannot carry")
                d["lineCoefficients"] = {k: clean(v) for k, v in line_coeffs.items()}
    return totals, derived


def _infer_signs(sections, derived):
    """Sign role per section: ``(signs, sources)``; raise naming the section when undecidable.

    Only the bottom-line rows vote (amount rows no other amount row reads), with their
    coefficients expanded through the rows they read: in ``Total Costs = COGS + Opex`` and
    ``Profit = Revenue - Total Costs`` the sections enter ``Profit`` as -1, and the helper
    row's own +1 is not a second opinion. A section that only enters rows through an
    individual line has no section-level vote and falls back to its name (flagged).
    """
    amount = [d for d in derived if d.get("kind") == "amount"]
    read_by_others = set()
    for d in amount:
        read_by_others |= calc.derived_references(d["expr"])
    seen = {name: [] for name in sections}
    for d in amount:
        if d["label"] in read_by_others:
            continue
        for name, coefficient in (d.get("coefficients") or {}).items():
            if abs(abs(coefficient) - 1) > 1e-9:
                raise PlanDraftError(
                    f"section {name!r}: derived row {d['label']!r} weights it by {coefficient:g}; "
                    "sections must enter derived rows as +1 or -1")
            seen[name].append((1 if coefficient > 0 else -1, d["label"]))
    signs, sources, needs_name = {}, {}, []
    for name in sections:
        found = {sign for sign, _ in seen[name]}
        if len(found) > 1:
            rows = ", ".join(f"{label!r} ({'+' if sign > 0 else '-'})" for sign, label in seen[name])
            raise PlanDraftError(f"section {name!r} has ambiguous sign: derived rows disagree ({rows})")
        if found:
            signs[name], sources[name] = found.pop(), "derived-rows"
        else:
            needs_name.append(name)
    for name in needs_name:
        lowered = name.lower()
        positive = any(word in lowered for word in POSITIVE_WORDS)
        negative = any(word in lowered for word in NEGATIVE_WORDS)
        if positive == negative:
            raise PlanDraftError(
                f"section {name!r}: no derived row uses it and its name reads as "
                f"{'both revenue and cost' if positive else 'neither revenue nor cost'}; "
                "cannot assign a sign role (add a derived row such as Profit = Revenue - this section)")
        signs[name], sources[name] = (1 if positive else -1), "name"
    return signs, sources


# ---------------------------------------------------------------- drafting
def draft_plan(source, *, intent, connection, folder, name=None, sheet=None,
               currency_format="$,.0f", pct_format=".1%", tolerance=None, look=None):
    """Return ``(plan, summary)`` for ``source``; raise ``PlanDraftError`` on any unknown shape.

    ``tolerance`` is an absolute bound on cached-value comparisons; by default only float
    noise is allowed (see ``forecast_calc.close``). ``look`` is a reference look
    (``style-reference.py``) the built workbook takes instead of the default presentation. Any failure that is not already a
    ``PlanDraftError`` (an unexpected cell value, a defect in this module) is reported as
    one naming the last cell read, so a caller that routes on ``PlanDraftError`` never
    crashes on a workbook.
    """
    _at(None)
    try:
        return _draft_plan(source, intent=intent, connection=connection, folder=folder, name=name,
                           sheet=sheet, currency_format=currency_format, pct_format=pct_format,
                           tolerance=tolerance, look=look)
    except PlanDraftError:
        raise
    except Exception as error:  # noqa: BLE001 - the router must never crash on a workbook
        where = _WHERE["cell"] or str(source)
        raise PlanDraftError(
            f"{where}: unexpected {type(error).__name__}: {error}; the drafter cannot read this shape") from error


def _draft_plan(source, *, intent, connection, folder, name, sheet, currency_format, pct_format,
                tolerance, look=None):
    if intent not in INTENTS:
        raise PlanDraftError(f"intent must be one of {', '.join(INTENTS)}")
    if fpres.rl.problem(look):
        raise PlanDraftError(fpres.rl.problem(look).replace("presentation.look", "--look"))
    src = Path(source).expanduser().resolve()
    if not src.is_file():
        raise PlanDraftError(f"source workbook not found: {src}")
    wf = load_workbook(src)
    wv = load_workbook(src, data_only=True)

    ws = find_report_sheet(wf, sheet)
    vs = wv[ws.title]
    header_row, grain, headers = find_period_header(ws)
    period_cols = [h["col"] for h in headers]
    next_header = ws.cell(header_row, period_cols[-1] + 1).value
    fy_col = period_cols[-1] + 1 if _text(next_header) and FY_HEADER_RE.match(_text(next_header)) else None
    extra_headers = [str(ws.cell(header_row, c).value) for c in range(period_cols[-1] + 1 + (1 if fy_col else 0),
                                                                      ws.max_column + 1)
                     if _text(ws.cell(header_row, c).value)]

    # ---- sections, line rows, derived-row candidates
    sections, line_rows, derived_candidates = [], [], []
    pending_label, current = None, None
    for kind, row, label in _walk_rows(ws, header_row, period_cols, fy_col):
        if kind == "section":
            pending_label, current = label, None
            continue
        text = _text(label)
        if text is None:
            raise PlanDraftError(f"{_ref(ws, 1, row)}: row has values but no text label")
        if _classify_row(ws, row, text, period_cols) == "derived":
            derived_candidates.append((row, text))
            continue
        if pending_label is None:
            raise PlanDraftError(f"{_ref(ws, 1, row)} ({text!r}): line appears before any section label")
        if current is None:
            if pending_label in sections:
                raise PlanDraftError(f"{_ref(ws, 1, row)}: section {pending_label!r} appears twice")
            sections.append(pending_label)
            current = pending_label
        line_rows.append((row, text, current))
    if not line_rows:
        raise PlanDraftError(f"{ws.title!r}: no SUMIFS actuals lines found under a section")
    names = Counter(text for _, text, _ in line_rows)
    dup = [n for n, c in names.items() if c > 1]
    if dup:
        raise PlanDraftError(f"duplicate line labels {dup}: SUMIFS keys on the label, so it must be unique")

    # ---- actual vs forecast columns: leading periods where a line carries SUMIFS
    has_actual = [any(_formula(ws.cell(row, col).value) and SUMIFS_RE.search(ws.cell(row, col).value)
                      for row, _, _ in line_rows) for col in period_cols]
    n_act = 0
    while n_act < len(period_cols) and has_actual[n_act]:
        n_act += 1
    unit = "months" if grain == "month" else "periods"
    if n_act == 0 or n_act == len(period_cols) or any(has_actual[n_act:]):
        raise PlanDraftError(
            f"{ws.title!r}: expected SUMIFS actual {unit} followed by forecast {unit}; "
            f"found {n_act} actual of {len(period_cols)} {unit}")
    act_cols, fc_cols = period_cols[:n_act], period_cols[n_act:]
    periods = []
    for i, header in enumerate(headers):
        periods.append({"label": header["label"], "num": i + 1, "grain": grain, "end": header["end"],
                        "role": "actual" if i < n_act else "forecast"})
    actual_periods, forecast_periods = periods[:n_act], periods[n_act:]
    last_act = act_cols[-1]
    sumifs_sheets = set()
    for row, text, sec in line_rows:
        for col in act_cols:
            cell = ws.cell(row, col).value
            if _formula(cell) and SUMIFS_RE.search(cell):
                for found in SHEET_REF_RE.finditer(cell):
                    sumifs_sheets.add((found.group(1) or found.group(2)).replace("''", "'"))
                continue
            zero = (cell == 0 and _is_number(cell)) or (_formula(cell) and ZERO_FORMULA_RE.match(_compact(cell)))
            if not zero:
                raise PlanDraftError(
                    f"{_ref(ws, col, row)} ({text!r}): row is neither a SUMIFS actuals line nor a "
                    f"recognised same-column total/derived row (actual-{unit[:-1]} cell {cell!r} is "
                    "not a SUMIFS formula or a zero)")
            if _is_number(vs.cell(row, col).value) and abs(vs.cell(row, col).value) > (
                    1e-6 if tolerance is None else tolerance):
                raise PlanDraftError(f"{_ref(ws, col, row)} ({text!r}): zero literal has cached value "
                                     f"{vs.cell(row, col).value!r}")

    bad = [s for s in sections if re.search(r'["\[\]]', s)]
    if bad:
        raise PlanDraftError(f"section name {bad[0]!r} contains a quote or bracket the Sigma formulas cannot carry")
    totals, derived = _derive_rows(ws, vs, period_cols, line_rows, sections, derived_candidates)
    signs, sign_sources = _infer_signs(sections, derived)
    section_of_total_row = {row: sec for sec, (row, _) in totals.items()}

    # ---- drivers
    lines = []
    driver_ranges = []
    flat_lines, external_evidence = [], []
    frozen_drivers = {}   # "Sheet!B4" -> (formula, cached rate, [lines that read it])

    def note_driver(dsheet, dcol, drow, rate, user=None):
        """Remember a driver cell that holds a formula: the plan keeps only its cached rate."""
        cell = wf[dsheet][f"{dcol}{drow}"].value
        if _formula(cell):
            entry = frozen_drivers.setdefault(f"{dsheet}!{dcol}{drow}", (cell, rate, []))
            if user and user not in entry[2]:
                entry[2].append(user)
    for line_order, (row, text, sec) in enumerate(line_rows, 1):
        kinds = []
        for col in fc_cols:
            _at(f"{_ref(ws, col, row)} ({text!r})")
            kind, detail = _forecast_cell_kind(ws, col, row, ws.cell(row, col).value)
            if kind is None:
                raise PlanDraftError(
                    f"{_ref(ws, col, row)} ({text!r}): unrecognised forecast cell {detail}; "
                    "expected =prev*(1+Drivers!$X$n), =prev+prev*Drivers!$X$n, =<k periods back>*(1+rate), "
                    "=<section total>*rate, =prev, an external-link cell, or a numeric literal")
            kinds.append((col, kind, detail))
        grouped = [(col, "manual" if kind == "external" else kind, None if kind == "external" else detail)
                   for col, kind, detail in kinds]
        first_col, first_kind, first_detail = grouped[0]
        for col, kind, detail in grouped[1:]:
            if kind != first_kind or detail != first_detail:
                raise PlanDraftError(
                    f"{_ref(ws, col, row)} ({text!r}): forecast cells are not uniform "
                    f"({get_column_letter(first_col)}{row} is {first_kind}, "
                    f"{get_column_letter(col)}{row} is {kind}"
                    f"{'' if detail == first_detail else ' with a different driver'})")
        base = vs.cell(row, last_act).value
        entry = {"line": text, "section": sec, "secOrder": sections.index(sec) + 1,
                 "lineOrder": line_order, "sourceRow": row, "kind": first_kind}
        if first_kind in ("growth", "yoy", "pct_of_total"):
            dsheet, dcol, drow = first_detail[:3]
            if dsheet not in wv.sheetnames:
                raise PlanDraftError(f"{_ref(ws, first_col, row)}: driver sheet {dsheet!r} not found")
            rate = wv[dsheet][f"{dcol}{drow}"].value
            if not _is_number(rate):
                raise PlanDraftError(
                    f"{dsheet}!{dcol}{drow}: driver for {text!r} is not a numeric rate ({rate!r})")
            note_driver(dsheet, dcol, drow, float(rate), text)
            if (dsheet, dcol) not in driver_ranges:
                driver_ranges.append((dsheet, dcol))
            drv_name = next((_text(c.value) for c in wv[dsheet][drow] if _text(c.value)), None)
            if first_kind == "growth":
                entry.update(rate=float(rate), juneBase=_number(base, ws, last_act, row, text),
                             driver=drv_name, driverCell=f"{dsheet}!{dcol}{drow}")
            elif first_kind == "yoy":
                if first_detail[3] > len(act_cols):
                    raise PlanDraftError(
                        f"{_ref(ws, first_col, row)} ({text!r}): reads {first_detail[3]} periods back "
                        f"but only {len(act_cols)} actual periods exist")
                entry.update(rate=float(rate), lag=first_detail[3], driver=drv_name,
                             driverCell=f"{dsheet}!{dcol}{drow}")
            else:
                owner = section_of_total_row.get(first_detail[3])
                if owner is None or owner == sec:
                    raise PlanDraftError(
                        f"{_ref(ws, first_col, row)} ({text!r}): percentage base row {first_detail[3]} "
                        "is not the total of another section")
                entry.update(rate=float(rate), of=owner, driver=drv_name,
                             driverCell=f"{dsheet}!{dcol}{drow}")
        elif first_kind == "flat":
            entry.update(juneBase=_number(base, ws, last_act, row, text))
            flat_lines.append(text)
        else:
            values = []
            for col in fc_cols:
                cached = vs.cell(row, col).value
                if not _is_number(cached):
                    raise PlanDraftError(
                        f"{_ref(ws, col, row)} ({text!r}): no numeric cached value ({cached!r}); open the "
                        "workbook in Excel, recalculate and save so cached values exist")
                values.append(float(cached))
            entry["manual"] = values
            links = [(col, ws.cell(row, col).value) for col, kind, _ in kinds if kind == "external"]
            if links:
                external_evidence.append(
                    f"{text}: " + "; ".join(f"{_ref(ws, col, row)} = {formula}" for col, formula in links))
        lines.append(entry)
    pct_sections = {ln["section"] for ln in lines if ln["kind"] == "pct_of_total"}
    for ln in lines:
        if ln["kind"] == "pct_of_total" and ln["of"] in pct_sections:
            raise PlanDraftError(
                f"{_ref(ws, 1, ln['sourceRow'])} ({ln['line']!r}): percentage base section "
                f"{ln['of']!r} itself holds percentage lines (circular)")

    assumptions = []
    first_driver = None     # (sheet, first driver row, rate column): the drivers table's source cells
    for dsheet, dcol in driver_ranges:
        dws = wv[dsheet]
        rate_col = column_index_from_string(dcol)
        for drow in range(1, dws.max_row + 1):
            rate = dws.cell(drow, rate_col).value
            drv_name = next((_text(c.value) for c in dws[drow][:rate_col - 1] if _text(c.value)), None)
            if drv_name and _is_number(rate):
                first_driver = first_driver or (dsheet, drow, rate_col)
                note_driver(dsheet, dcol, drow, float(rate))
                basis = next((_text(c.value) for c in dws[drow][rate_col:] if _text(c.value)), "")
                assumptions.append([drv_name, float(rate), basis])

    # ---- actuals (GL)
    if len(sumifs_sheets - {ws.title}) != 1:
        raise PlanDraftError(
            f"expected the actual-{unit[:-1]} SUMIFS to read exactly one GL sheet, found "
            f"{sorted(sumifs_sheets) or 'none'}")
    gl_name = next(iter(sumifs_sheets - {ws.title}))
    actuals = read_actuals(wv[gl_name], lines, actual_periods, grain)

    # ---- plan core + verification: every line-period, section total and derived row
    section_meta = []
    for order, sec in enumerate(sections, 1):
        meta = {"name": sec, "order": order, "sign": signs[sec], "signSource": sign_sources[sec]}
        if sec in totals:
            meta.update(totalLabel=totals[sec][1], totalRow=totals[sec][0])
        section_meta.append(meta)
    core = {"periods": periods, "sections": section_meta, "lines": lines, "actuals": actuals,
            "derivedRows": derived}
    _at(f"{ws.title}: re-deriving the drafted plan")
    try:
        ev = calc.Evaluation(core)
    except calc.RowEvalError as error:
        owner = next((d for d in derived if d["label"] == error.label), None)
        where = (_ref(ws, period_cols[error.index], owner["sourceRow"])
                 if owner is not None and error.index is not None else ws.title)
        raise PlanDraftError(f"{where} ({error.label!r}): drafted plan cannot be evaluated: {error}; "
                             "a derived row that divides by a total must guard it with IFERROR or "
                             "IF(...=0,...)") from None
    except (calc.ExprError, ZeroDivisionError) as error:
        raise PlanDraftError(f"{ws.title!r}: drafted plan cannot be evaluated: {error}") from None
    problems = verify(ws, vs, ev, lines, totals, derived, period_cols, tolerance)
    if problems:
        shown = "\n  ".join(problems[:10])
        more = f"\n  ... and {len(problems) - 10} more" if len(problems) > 10 else ""
        raise PlanDraftError(
            f"drafted plan does not reproduce the workbook's cached values "
            f"({len(problems)} mismatch(es)):\n  {shown}{more}")

    # A header alone does not prove a full-year total: every line row's cached FY
    # value must equal the sum of the cells it covers, else targets come from the periods.
    fy_cols = _fy_coverage(ws, vs, line_rows, period_cols, fy_col, tolerance) if fy_col else None
    if fy_cols is None:
        fy_col = None
    fy_labels = [headers[period_cols.index(c)]["label"] for c in fy_cols] if fy_cols else None
    fy_idx = [period_cols.index(c) for c in fy_cols] if fy_cols else list(range(len(period_cols)))
    targets, skipped = parity_targets(ws, vs, ev, totals, derived, period_cols, fy_col, fy_idx,
                                      sections, signs, tolerance)
    for d in derived:
        d["values"] = {periods[i]["label"]: (float(vs.cell(d["sourceRow"], col).value)
                                             if _is_number(vs.cell(d["sourceRow"], col).value) else None)
                       for i, col in enumerate(period_cols)}
        if fy_col and _is_number(vs.cell(d["sourceRow"], fy_col).value):
            d["fyTotal"] = float(vs.cell(d["sourceRow"], fy_col).value)
    for meta in section_meta:
        if "totalRow" in meta:
            meta["values"] = {periods[i]["label"]: float(vs.cell(meta["totalRow"], col).value)
                              for i, col in enumerate(period_cols)}
            if fy_col and _is_number(vs.cell(meta["totalRow"], fy_col).value):
                meta["fyTotal"] = float(vs.cell(meta["totalRow"], fy_col).value)
    charts, unmapped_charts = source_charts(wf, wv, ev, tolerance)
    gl_row, gl_cols = gl_header(wf[gl_name])
    presentation, format_notes = fpres.propose(
        src=src, wf=wf, ws=ws, header_row=header_row, period_cols=period_cols, line_rows=line_rows,
        totals=totals, derived=derived, gl=wf[gl_name], gl_header=gl_row, gl_cols=gl_cols,
        driver=first_driver, look=look)

    issues = []
    if flat_lines:
        issues.append({
            "code": "FLAT_CARRY_POLICY",
            "severity": "review",
            "route": "forecast-carry-forward-policy-review",
            "message": (f"{len(flat_lines)} forecast line(s) carry the prior {grain} forward "
                        "(=prior); that syntax alone does not prove a flat carry-forward policy."),
            "evidence": flat_lines,
        })
    if external_evidence:
        issues.append({
            "code": "EXTERNAL_LINK_VALUE",
            "severity": "review",
            "route": "forecast-external-link-review",
            "message": (f"{len(external_evidence)} forecast line(s) read another workbook "
                        "(external link). The plan carries the cached value Excel last stored as a "
                        "manual input; the link itself is not rebuilt, so confirm the numbers are current."),
            "evidence": external_evidence,
        })
    if frozen_drivers:
        issues.append({
            "code": "DRIVER_FORMULA_FROZEN",
            "severity": "review",
            "route": "forecast-driver-formula-review",
            "message": (f"{len(frozen_drivers)} driver cell(s) hold a formula, not a typed rate. The plan "
                        "carries the rate Excel last cached for each as an editable input; the formula "
                        "is not rebuilt, so a change to its inputs in the source will not flow through. "
                        "Confirm the cached rates are current, or rebuild the formula."),
            "evidence": [f"{cell}: {formula} (cached {rate:g}"
                         + (f"; read by {', '.join(users)}" if users else "") + ")"
                         for cell, (formula, rate, users) in frozen_drivers.items()],
        })
    if format_notes.codes:
        issues.append({
            "code": "NUMBER_FORMAT_APPROXIMATED",
            "severity": "review",
            "route": "number-format-review",
            "message": ("These Excel number formats have no exact Sigma format; the closest d3 / strftime "
                        "format was used. Confirm the display, not the values."),
            "evidence": format_notes.evidence(),
        })
    name_signed = [s for s in sections if sign_sources[s] == "name"]
    if name_signed:
        issues.append({
            "code": "SECTION_SIGN_FROM_NAME",
            "severity": "review",
            "route": "forecast-section-sign-review",
            "message": (f"{len(name_signed)} section(s) are not used by any derived row, so their sign "
                        "role (adds to or subtracts from the bottom line) came from the section name."),
            "evidence": [f"{s}: {'+' if signs[s] > 0 else '-'}" for s in name_signed],
        })

    plan = {
        "name": name or (wf.properties.title or "").strip() or src.stem,
        "intent": intent,
        "connectionId": connection,
        "folderId": folder,
        "currencyFormat": currency_format,
        "pctFormat": pct_format,
        "sourcePath": str(src),
        "sourceSha256": source_sha256(src),
        "parityTargets": targets,
    }
    if grain == "month":
        plan["actualMonths"] = [month_meta(_header_period(ws.cell(header_row, c).value)[1], i + 1)
                                for i, c in enumerate(act_cols)]
        plan["forecastMonths"] = [month_meta(_header_period(ws.cell(header_row, c).value)[1], n_act + i + 1)
                                  for i, c in enumerate(fc_cols)]
    plan.update({
        "periods": periods,
        "sections": section_meta,
        "derivedRows": derived,
        "lines": lines,
        "assumptions": assumptions,
        "actuals": actuals,
        "sourceCharts": charts,
        "presentation": presentation,
        "draftReviewIssues": issues,
    })
    if fy_labels and len(fy_labels) != len(periods):
        plan["fyPeriods"] = fy_labels
        plan["fyLabel"] = _text(ws.cell(header_row, fy_col).value)
    summary = {
        "reportSheet": ws.title, "glSheet": gl_name, "grain": grain,
        "kinds": dict(Counter(line["kind"] for line in lines)),
        "lines": len(lines), "glRows": len(actuals), "drivers": len(assumptions),
        "sections": [f"{s} ({'+' if signs[s] > 0 else '-'}, {sign_sources[s]})" for s in sections],
        "derivedRows": [d["label"] for d in derived],
        "actualRange": (actual_periods[0]["label"], actual_periods[-1]["label"]),
        "forecastRange": (forecast_periods[0]["label"], forecast_periods[-1]["label"]),
        "fyColumn": get_column_letter(fy_col) if fy_col else None,
        "fyPeriods": fy_labels if fy_labels and len(fy_labels) != len(periods) else None,
        "parityTargets": targets,
        "skippedTargets": skipped,
        "ignoredColumns": extra_headers,
        "missingTargets": [m for m in TARGET_LABELS.values() if m not in targets],
        "reviewIssues": [issue["code"] for issue in issues],
        "charts": [f"{c['title']} ({c['grain']}, {len(c['series'])} series)" for c in charts],
        "unmappedCharts": unmapped_charts,
        "verifiedCells": verified_count(lines, totals, derived, period_cols),
        "presentation": {
            "title": presentation["title"],
            "money": ((presentation["formats"]["money"].get("format") or {}).get("formatString")
                      or f"{currency_format} (--currency-format: the source cells are General)"),
            "percent": next((e["format"]["formatString"] for e in presentation["formats"]["derived"].values()
                             if fpres.ef.is_percent_format(e.get("format"))), None),
            "notes": presentation["notes"],
        },
    }
    return plan, summary


def verified_count(lines, totals, derived, period_cols):
    return (len(lines) + len(totals) + len(derived)) * len(period_cols)


def _number(value, ws, col, row, label):
    if not _is_number(value):
        raise PlanDraftError(
            f"{_ref(ws, col, row)} ({label!r}): no numeric cached value ({value!r}); open the "
            "workbook in Excel, recalculate and save so cached values exist")
    return float(value)


GL_ROLES = {
    "date": lambda h, t: "date" in t,
    "month": lambda h, t: h in ("month", "period", "month label", "quarter", "fiscal period",
                                "fiscal quarter", "period label", "qtr"),
    "line": lambda h, t: h in ("line", "line item", "account", "gl account", "line item name",
                               "account name"),
    "dept": lambda h, t: bool(t & {"dept", "department", "store", "location"}),
    "vendor": lambda h, t: "vendor" in t,
    "amt": lambda h, t: bool(t & {"amount", "amt"}),
}
GL_REQUIRED = ("month", "line", "amt")
GL_ROLE_NAMES = {"month": "period", "line": "line item", "amt": "amount", "date": "date",
                 "dept": "department", "vendor": "vendor"}


def gl_header(gl):
    """``(header row, {role: column})`` of the GL sheet: the first row that maps period, line item and amount."""
    header_row, cols, seen = None, {}, []
    for row in range(1, min(gl.max_row, 40) + 1):
        found = {}
        for col in range(1, gl.max_column + 1):
            value = gl.cell(row, col).value
            if not _text(value):
                continue
            norm = re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()
            tokens = set(norm.split())
            for role, match in GL_ROLES.items():
                if match(norm, tokens):
                    found.setdefault(role, []).append(col)
        seen.append((len([r for r in GL_REQUIRED if r in found]), row, found))
        if all(role in found for role in GL_REQUIRED):
            header_row, cols = row, found
            break
    if header_row is None:
        best = max(seen, key=lambda item: item[0])
        missing = [GL_ROLE_NAMES[r] for r in GL_REQUIRED if r not in best[2]]
        raise PlanDraftError(
            f"{gl.title!r}: no header row maps all of period, line item, amount; "
            f"best candidate row {best[1]} is missing {', '.join(missing)}")
    ambiguous = [GL_ROLE_NAMES[r] for r, found in cols.items() if len(found) > 1]
    if ambiguous:
        raise PlanDraftError(f"{gl.title!r} row {header_row}: more than one column matches "
                             f"{', '.join(ambiguous)}")
    return header_row, {role: found[0] for role, found in cols.items()}


def read_actuals(gl, lines, actual_periods, grain="month"):
    """Map the GL header by column name and return plan ``actuals`` rows.

    Period, line item and amount are required; date, department/store and vendor are
    read when the sheet has them (blank otherwise).
    """
    by_line = {line["line"]: line for line in lines}
    labels = {p["label"].casefold(): p["label"] for p in actual_periods}
    word = "month" if grain == "month" else "period"
    header_row, col = gl_header(gl)

    out = []
    for row in range(header_row + 1, gl.max_row + 1):
        values = {role: gl.cell(row, c).value for role, c in col.items()}
        if all(v is None for v in values.values()):
            continue
        where = f"{gl.title}!row {row}"
        _at(where)
        label = _text(values["line"])
        if label not in by_line:
            raise PlanDraftError(f"{where}: line item {values['line']!r} is not a P&L line")
        if grain == "month":
            parsed = parse_month(values["month"])
            period = (f"{MONTH_ABBR[parsed[1] - 1].capitalize()}-{parsed[0] % 100:02d}"
                      if parsed else None)
            period = labels.get(period.casefold()) if period else None
        else:
            raw = values["month"]
            if isinstance(raw, datetime):
                raw = raw.date()
            text = raw.isoformat() if isinstance(raw, date) else (str(raw).strip() if raw is not None else None)
            period = labels.get(text.casefold()) if text is not None else None
        if period is None:
            raise PlanDraftError(f"{where}: {word} {values['month']!r} is not an actual {word} "
                                 f"({', '.join(sorted(p['label'] for p in actual_periods))})")
        if not _is_number(values["amt"]):
            raise PlanDraftError(f"{where}: amount {values['amt']!r} is not numeric")
        stamp = None
        if "date" in values:
            stamp = values["date"]
            if isinstance(stamp, datetime):
                stamp = stamp.date()
            if isinstance(stamp, date):
                stamp = stamp.isoformat()
            elif not (isinstance(stamp, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", stamp.strip())):
                raise PlanDraftError(f"{where}: date {values['date']!r} is not a date")
            stamp = stamp.strip()
        line = by_line[label]
        out.append({"date": stamp, "month": period,
                    "line": label, "section": line["section"], "secOrder": line["secOrder"],
                    "lineOrder": line["lineOrder"], "dept": _text(values.get("dept")) or "",
                    "vendor": _text(values.get("vendor")) or "", "amt": float(values["amt"])})
    if not out:
        raise PlanDraftError(f"{gl.title!r}: no GL rows below the header")
    return out


def verify(ws, vs, ev, lines, totals, derived, period_cols, tolerance):
    """Recompute each cell the way the builder does; list disagreements with the sheet."""
    problems = []

    def compare(coord, label, period, got, want):
        if not _is_number(want):
            problems.append(f"{coord} ({label!r}, {period}): sheet has no numeric cached value ({want!r})")
        elif not calc.close(got, want, tolerance):
            digits = 2 if f"{got:,.2f}" != f"{want:,.2f}" else 9  # show a sub-cent difference
            problems.append(f"{coord} ({label!r}, {period}): plan {got:,.{digits}f} "
                            f"vs sheet {want:,.{digits}f}")

    for ln in lines:
        row = ln["sourceRow"]
        for i, col in enumerate(period_cols):
            compare(f"{ws.title}!{get_column_letter(col)}{row}", ln["line"], ev.labels[i],
                    ev.matrix[ln["line"]][i], vs.cell(row, col).value)
    for sec, (row, label) in totals.items():
        for i, col in enumerate(period_cols):
            compare(f"{ws.title}!{get_column_letter(col)}{row}", label, ev.labels[i],
                    ev.totals[sec][i], vs.cell(row, col).value)
    for d in derived:
        row = d["sourceRow"]
        for i, col in enumerate(period_cols):
            cached = vs.cell(row, col).value
            if isinstance(cached, bool):
                cached = float(cached)  # Excel stores IF(cond, x) = FALSE as a boolean: 0
            got = ev.derived[d["label"]][i]
            if isinstance(cached, str) and d["kind"] == "ratio":
                continue  # an IFERROR text result is not a number to compare
            compare(f"{ws.title}!{get_column_letter(col)}{row}", d["label"], ev.labels[i], got, cached)
    return problems


def _fy_coverage(ws, vs, line_rows, period_cols, fy_col, tolerance):
    """Period columns the FY cells sum (same for every line row), else None."""
    covered = None
    for row, _, _ in line_rows:
        formula = ws.cell(row, fy_col).value
        cols = period_cols
        if _formula(formula):
            found = re.fullmatch(r"=SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)",
                                 UNQUOTED_SPACE_RE.sub("", formula), re.I)
            if found and int(found.group(2)) == row == int(found.group(4)):
                lo, hi = column_index_from_string(found.group(1)), column_index_from_string(found.group(3))
                cols = [c for c in period_cols if lo <= c <= hi]
                if not cols:
                    return None
        cells = [vs.cell(row, col).value for col in cols]
        total = vs.cell(row, fy_col).value
        if not (_is_number(total) and all(_is_number(c) for c in cells)
                and calc.close(sum(cells), total, tolerance)):
            return None
        if covered is None:
            covered = cols
        elif covered != cols:
            return None
    return covered


def parity_targets(ws, vs, ev, totals, derived, period_cols, fy_col, fy_idx, sections, signs,
                   tolerance):
    """Cached FY (else summed period) value of every total / derived row, keyed by label.

    The four classic P&L rows keep their canonical keys (Revenue, Gross Profit, EBITDA,
    Net Income); when every section is one of the four classic ones only those keys are
    emitted, so classic plans keep their historical targets. A ratio row is a target
    only when its FY cell equals the formula evaluated over the FY totals.
    Returns ``(targets, skipped)``.
    """
    classic = set(sections) <= set(SECTIONS) and all(signs[s] == LEGACY_SIGNS[s] for s in sections)
    rows = [(row, label, None) for sec, (row, label) in totals.items()]
    rows += [(d["sourceRow"], d["label"], d) for d in derived]
    rows.sort(key=lambda item: item[0])
    targets, skipped = {}, []
    for row, label, d in rows:
        metric = TARGET_LABELS.get(label.lower())
        if classic and not metric and not (d is not None and d["kind"] == "amount"):
            continue  # classic plans keep the four canonical keys plus any extra amount row
        key = metric or label
        if key in targets:
            continue
        ratio = d is not None and d["kind"] == "ratio"
        if fy_col:
            value = vs.cell(row, fy_col).value
            if ratio:
                try:
                    got = ev.aggregate(label, fy_idx)
                except calc.ExprError as error:
                    skipped.append(f"{label}: cannot be evaluated over the FY totals ({error})")
                    continue
                if not (_is_number(value) and calc.close(got, value, tolerance)):
                    skipped.append(f"{label}: FY cell {value!r} is not the formula over FY totals")
                    continue
            elif d is not None:
                got = sum(ev.derived[label][i] for i in fy_idx)
                if _is_number(value) and not calc.close(got, value, tolerance):
                    raise PlanDraftError(
                        f"{_ref(ws, fy_col, row)} ({label!r}): FY cell {value:,.2f} vs plan "
                        f"{got:,.2f} over the periods it covers")
        else:
            if ratio:
                skipped.append(f"{label}: no FY column to verify a ratio against")
                continue
            cells = [vs.cell(row, col).value for col in period_cols]
            value = sum(cells) if all(_is_number(c) for c in cells) else None
        if _is_number(value):
            targets[key] = float(value)
    return targets, skipped


# ---------------------------------------------------------------- source charts
QUARTER_LABEL_RE = re.compile(r"^\s*Q([1-4])\b", re.I)


def _range_values(wv, formula):
    """Cached values of a chart reference such as ``Quarterly!$B$4:$E$4`` (row or column)."""
    found = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)",
                         formula or "")
    if not found:
        return None
    sheet = (found.group(1) or found.group(2)).replace("''", "'")
    if sheet not in wv.sheetnames:
        return None
    min_col, min_row, max_col, max_row = range_boundaries(found.group(3).replace("$", ""))
    ws = wv[sheet]
    return [ws.cell(r, c).value for r in range(min_row, max_row + 1)
            for c in range(min_col, max_col + 1)]


def _metric_rows(ev):
    """``{lowercase label: (output name, per-period values)}`` for chartable amount rows."""
    rows = {}
    for sec, label in ev.total_labels.items():
        if label:
            rows[label.lower()] = (TARGET_LABELS.get(label.lower(), label), ev.totals[sec])
    for label, values in ev.derived.items():
        row = next(d for d in ev.plan["derivedRows"] if d["label"] == label)
        if row["kind"] == "amount":
            rows[label.lower()] = (TARGET_LABELS.get(label.lower(), label), values)
    return rows


def source_charts(wf, wv, ev, tolerance):
    """Map workbook line and bar charts whose series are P&L rows to plan ``sourceCharts``.

    A chart maps when every series title is a section-total or amount derived row and
    its categories are the plan's periods (or calendar quarters / months of a month
    plan). Mapped charts must reproduce their cached series values from the plan; a
    mismatch raises. Returns ``(charts, unmapped)`` where ``unmapped`` names charts left
    to review.
    """
    periods = ev.periods
    grain = periods[0]["grain"]
    rows = _metric_rows(ev)
    quarters = []
    if grain == "month":
        for i, p in enumerate(periods):
            key = (int(p["end"][:4]), (int(p["end"][5:7]) - 1) // 3 + 1)
            if not quarters or quarters[-1][0] != key:
                quarters.append((key, []))
            quarters[-1][1].append(i)

    charts, unmapped = [], []
    for ws in wf.worksheets:
        for chart in getattr(ws, "_charts", []):
            title = _chart_title(chart) or f"{type(chart).__name__} on {ws.title}"
            where = f"{title!r} on {ws.title!r}"
            kind = {"LineChart": "line", "BarChart": "bar"}.get(type(chart).__name__)
            if kind is None or not chart.series:
                unmapped.append(f"{where}: only line and bar charts are mapped")
                continue
            if kind == "bar" and getattr(chart, "grouping", "clustered") not in ("clustered", "standard", None):
                unmapped.append(f"{where}: only clustered bar charts are mapped")
                continue
            metrics, categories = [], None
            for s in chart.series:
                ref = s.tx.strRef.f if s.tx is not None and s.tx.strRef is not None else None
                label = (_range_values(wv, ref) or [None])[0] if ref else None
                metric = rows.get(str(label).strip().lower()) if label else None
                cat_source = None
                if s.cat is not None:
                    cat_source = s.cat.strRef if s.cat.strRef is not None else s.cat.numRef
                cat_ref = cat_source.f if cat_source is not None else None
                cats = _range_values(wv, cat_ref) if cat_ref else None
                vals = _range_values(wv, s.val.numRef.f) if s.val is not None and s.val.numRef else None
                if not metric or not cats or not vals:
                    metrics = None
                    break
                metrics.append((metric[0], metric[1], vals))
                categories = cats
            if not metrics:
                unmapped.append(f"{where}: a series is not a section total or amount row of the P&L")
                continue
            if grain != "month":
                labels = [str(c).strip().casefold() if c is not None else "" for c in categories]
                plan_labels = [p["label"].casefold() for p in periods]
                if labels != plan_labels:
                    unmapped.append(f"{where}: categories {categories} are not all {len(periods)} plan periods; "
                                    "a chart over a subset is left for review")
                    continue
                chart_grain = "period"
                buckets = [((p["label"],), [i]) for i, p in enumerate(periods)]
            elif all(QUARTER_LABEL_RE.match(str(c or "")) for c in categories):
                chart_grain = "quarter"
                labels = [int(QUARTER_LABEL_RE.match(str(c)).group(1)) for c in categories]
                if [key[1] for key, _ in quarters[:len(labels)]] != labels:
                    unmapped.append(f"{where}: quarter categories {categories} do not follow the plan months")
                    continue
                buckets = quarters[:len(labels)]
            else:
                labels = [parse_month(c) for c in categories]
                if not all(labels):
                    unmapped.append(f"{where}: categories are neither quarters nor months")
                    continue
                expected = [(int(p["end"][:4]), int(p["end"][5:7])) for p in periods[:len(labels)]]
                if labels != expected:
                    unmapped.append(f"{where}: month categories {categories} do not follow the plan months")
                    continue
                chart_grain = "month"
                buckets = [((p["label"],), [i]) for i, p in enumerate(periods)][:len(labels)]
            for name, series, vals in metrics:
                for (key, members), want in zip(buckets, vals):
                    got = sum(series[i] for i in members)
                    if not _is_number(want) or not calc.close(got, want, tolerance):
                        raise PlanDraftError(
                            f"chart {where}: series {name!r} at {key} is {want!r} in the workbook "
                            f"but {got:,.2f} from the plan")
            entry = {"title": title, "kind": kind, "grain": chart_grain,
                     "series": [name for name, _, _ in metrics],
                     "categories": [str(c) for c in categories]}
            if chart_grain != "period":
                # The source chart's own period range, so the rebuilt chart shows the same subset.
                covered = [periods[i] for _, members in buckets for i in members]
                entry["periodStart"] = covered[0]["end"][:8] + "01"
                entry["periodEnd"] = covered[-1]["end"]
            colors = fpres.series_colors(chart)
            if colors:
                entry["colors"] = colors
            entry.update(sourceSheet=ws.title, verifiedPoints=len(metrics) * len(buckets))
            charts.append(entry)
    return charts, unmapped


def _chart_title(chart):
    try:
        return "".join(run.t for p in chart.title.tx.rich.p for run in (p.r or [])).strip() or None
    except AttributeError:
        return None
