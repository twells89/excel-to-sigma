"""Table-report model: expression IR, Python evaluator, verification, Sigma spec builders.

A *table-report* plan describes a workbook whose outputs are formulas over formal
Excel Tables (inventory projection, sales pipeline, headcount plan). The drafter
(``table_report.py``) lowers every recognised formula to the small IR below; this
module evaluates the IR in Python (the local parity oracle) and renders the same IR
as Sigma formulas, so what is verified offline is exactly what is emitted.

Plan (JSON, ``planVersion`` 1)
------------------------------
``params``  [{name, value, cell}]               scalar cells the formulas read (editable)
``tables``  [{name, sheet, ref, editable, columns[{name,type,format?,expr?}], rows}]
            ``expr`` marks a calculated column; ``rows`` hold the literal columns only
``frames``  expansion | keyed | summary (see the builders below)
``checks``  every recognised source cell with Excel's cached value
``parityTargets``  the headline summary values the builder re-checks

IR (nested lists)
-----------------
literals   ["n", 1.5] ["s", "x"] ["b", true] ["d", "2026-07-01"] ["null"]
row/frame  ["c", col] this row / this summary column   ["key"] the grouping key
           ["period"] the period key   ["prev"] previous period's value
           ["p", name] a parameter
ops        ["op", sym, a, b]  ["neg", a]  ["f", NAME, *args]
lookup     ["lk", table, keyExpr, keyCol, valueCol, defaultExpr|null]
aggregate  ["agg", sum|count|avg|min|max, valueCol|null, filterExpr|null]  (summary only)
running    ["cumsum", expr, partitionCol, orderCol]                        (keyed / ledger)
           ["cummin", expr, partitionCol, orderCol]                        (ledger: running MIN)
params     ["p", name] number   ["pt", name] text   ["pd", name] date     (``Parameters`` table)

Ledger frames (``kind: "ledger"``) are a period ledger (a cash-flow style report: line items down, weeks
across in Excel) lowered to ONE table with a row per period: the spine table's period key plus manual
columns, and one calculated column per Excel row, in the sheet's row order. ``partitionCol`` may be null
there (a single running series over the whole table, ordered by the period).
"""
from __future__ import annotations

import csv
import json
import math
import re
from collections import OrderedDict
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from . import excel_formats as ef
    from . import presentation_style as ps
except ImportError:  # direct import when scripts/lib itself is on sys.path
    import excel_formats as ef
    import presentation_style as ps

PLAN_VERSION = 1
ARCHETYPE = "table-report"
PARAM_TABLE = "Parameters"
COLUMN_TYPES = ("number", "text", "datetime", "boolean")
FUNCTIONS = {"IF", "AND", "OR", "NOT", "MIN", "MAX", "ABS", "ROUND", "CEILING", "EOMONTH0",
             "MONTH", "YEAR", "IFERROR", "SUM", "MOD", "ADDDAYS", "CONCAT"}
COMPARE = {"=", "<>", "<", ">", "<=", ">="}
ARITH = {"+", "-", "*", "/", "^"}
AGGS = {"sum", "count", "avg", "min", "max"}
# Sigma functions the emitted formulas depend on (listed for dry-run validation).
SIGMA_FUNCTIONS = ("Lookup", "Coalesce", "SumIf", "CountIf", "AvgIf", "MinIf", "MaxIf", "Sum", "Count",
                   "Avg", "Min", "Max", "CumulativeSum", "CumulativeMin", "Ceiling", "DateAdd", "DateTrunc",
                   "DatePart", "Date", "IsNull", "If", "Least", "Greatest", "Abs", "Round", "Not", "Mod",
                   "Power", "Concat")


class PlanError(ValueError):
    """The plan (or an IR node in it) is malformed or unsupported."""


class EvalError(Exception):
    """A spreadsheet-style evaluation error (``#DIV/0!``, ``#N/A``...)."""

    def __init__(self, code, detail=""):
        super().__init__(f"{code} {detail}".strip())
        self.code = code


class Unmodelled(EvalError):
    """Excel semantics this drafter does not reproduce in Sigma; never caught by ``IFERROR``."""


# Errors ``IFERROR`` catches. These are the two the emitter can test for in Sigma (a zero/blank
# denominator, an exact lookup that finds nothing); any other error fails the draft loudly.
CAUGHT_ERRORS = ("#DIV/0!", "#N/A")


# ----------------------------------------------------------------- values
def parse_date(value):
    """ISO string / date / datetime -> date (time-of-day is refused, never dropped)."""
    if value is None or isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        if value.time() != datetime.min.time():
            raise PlanError(f"time-of-day is not supported in {value.isoformat()}")
        return value.date()
    if isinstance(value, str):
        try:
            return date.fromisoformat(value[:10]) if len(value) == 10 else datetime.fromisoformat(value).date()
        except ValueError as error:
            raise PlanError(f"not an ISO date: {value!r}") from error
    raise PlanError(f"not a date: {value!r}")


def iso(value):
    return value.isoformat() if isinstance(value, date) else value


def _num(x):
    if x is None:
        return 0
    if isinstance(x, bool):
        return 1 if x else 0
    if isinstance(x, (int, float)):
        return x
    raise EvalError("#VALUE!", repr(x))


def _blank_for(other):
    if isinstance(other, str):
        return ""
    if isinstance(other, bool):
        return False
    if isinstance(other, date):
        return date.min
    return 0


def _compare(sym, a, b):
    if a is None:
        a = _blank_for(b)
    if b is None:
        b = _blank_for(a)
    text_a, text_b = isinstance(a, str), isinstance(b, str)
    if text_a != text_b:  # Excel orders text above numbers; they are never equal
        if sym == "=":
            return False
        if sym == "<>":
            return True
        a_rank, b_rank = (1 if text_a else 0), (1 if text_b else 0)
        return {"<": a_rank < b_rank, ">": a_rank > b_rank, "<=": a_rank <= b_rank,
                ">=": a_rank >= b_rank}[sym]
    if isinstance(a, date) != isinstance(b, date):
        raise EvalError("#VALUE!", f"cannot compare {a!r} with {b!r}")
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[sym]


def _compare_crit(sym, a, b):
    """Excel *criteria* semantics (SUMIFS/COUNTIFS/AVERAGEIFS...), which differ from ``=A1<5``:
    a blank cell satisfies only ``<>`` (and ``=""`` / ``<>""``), and an ordering/equality criterion
    never matches a cell of the other type (text vs number)."""
    if a is None or a == "":
        if b == "" or b is None:
            return sym == "="
        return sym == "<>"
    if isinstance(b, str) != isinstance(a, str) or isinstance(b, date) != isinstance(a, date):
        return sym == "<>"
    return _compare(sym, a, b)


def _truthy(x):
    if x is None:
        return False
    if isinstance(x, str):
        raise EvalError("#VALUE!", repr(x))
    return bool(x)


def eomonth0(day):
    if not isinstance(day, date):
        raise EvalError("#VALUE!", repr(day))
    first_next = date(day.year + (day.month == 12), day.month % 12 + 1, 1)
    return first_next - timedelta(days=1)


# ----------------------------------------------------------------- IR utils
def walk(node):
    """Yield every IR node (pre-order)."""
    if isinstance(node, list) and node and isinstance(node[0], str):
        yield node
        for child in node[1:]:
            if isinstance(child, list):
                if child and isinstance(child[0], str):
                    yield from walk(child)
                else:
                    for item in child:
                        yield from walk(item)


def validate_ir(node, where="expression"):
    if node is None:
        return
    if not (isinstance(node, list) and node and isinstance(node[0], str)):
        raise PlanError(f"{where}: IR node must be a non-empty list, got {node!r}")
    tag = node[0]
    arity = {"n": 1, "s": 1, "b": 1, "d": 1, "null": 0, "c": 1, "key": 0, "period": 0, "prev": 0,
             "p": 1, "pt": 1, "pd": 1, "op": 3, "neg": 1, "lk": 5, "agg": 3, "cumsum": 3, "cummin": 3}
    if tag == "f":
        if len(node) < 3 or node[1] not in FUNCTIONS:
            raise PlanError(f"{where}: unsupported function node {node!r}")
        kids = node[2:]
    elif tag in arity:
        if len(node) - 1 != arity[tag]:
            raise PlanError(f"{where}: {tag} node needs {arity[tag]} operand(s): {node!r}")
        kids = []
        if tag == "op":
            kids = node[2:]
        elif tag == "neg":
            kids = node[1:]
        elif tag == "lk":
            kids = [node[2], node[5]]
        elif tag == "agg":
            kids = [node[3]]
        elif tag in ("cumsum", "cummin"):
            kids = [node[1]]
    else:
        raise PlanError(f"{where}: unknown IR node {tag!r}")
    if tag == "op" and node[1] not in ARITH | COMPARE:
        raise PlanError(f"{where}: unsupported operator {node[1]!r}")
    if tag == "agg" and node[1] not in AGGS:
        raise PlanError(f"{where}: unsupported aggregate {node[1]!r}")
    for kid in kids:
        validate_ir(kid, where)


class Env:
    __slots__ = ("row", "cols", "key", "period", "prev", "group", "cum", "idx", "crit")

    def __init__(self, row=None, cols=None, key=None, period=None, prev=None, group=None,
                 cum=None, idx=None, crit=False):
        self.row, self.cols, self.key, self.period = row, cols, key, period
        self.prev, self.group, self.cum, self.idx = prev, group, cum, idx
        self.crit = crit      # evaluating an aggregate criterion (Excel criteria semantics)


class Model:
    """Evaluates a plan's tables and frames (lazily, memoised)."""

    def __init__(self, plan):
        self.plan = plan
        self.tables = {t["name"]: t for t in plan.get("tables", [])}
        self.frames = {f["id"]: f for f in plan.get("frames", [])}
        self.params = {p["name"]: parse_date(p["value"]) if p.get("kind") == "date" else p["value"]
                       for p in plan.get("params", [])}
        self._table_rows, self._frame_rows, self._lookup_index = {}, {}, {}
        self._partial = {}
        self._busy = set()

    # ---- tables
    def table_rows(self, name):
        if name in self._table_rows:
            return self._table_rows[name]
        if name not in self.tables:
            raise PlanError(f"unknown table {name!r}")
        if ("table", name) in self._busy:
            raise PlanError(f"circular calculated columns involving table {name!r}")
        self._busy.add(("table", name))
        try:
            table = self.tables[name]
            literal = [c for c in table["columns"] if "expr" not in c]
            rows = []
            for values in table.get("rows", []):
                if len(values) != len(literal):
                    raise PlanError(f"table {name!r}: row has {len(values)} values for {len(literal)} columns")
                row = {}
                for column, value in zip(literal, values):
                    row[column["name"]] = parse_date(value) if column["type"] == "datetime" else value
                rows.append(row)
            self._partial[name] = rows   # lookups into this table may read its literal columns meanwhile
            for column in table["columns"]:
                if "expr" in column:
                    for number, row in enumerate(rows):
                        try:
                            value = self.ev(column["expr"], Env(row=row, idx=number))
                            row[column["name"]] = 0 if value is None else value   # a formula cell is never blank
                        except EvalError as error:
                            raise EvalError(error.code, f"table {name!r} column {column['name']!r} "
                                            f"row {number + 1}: {error}") from error
            self._table_rows[name] = rows
            return rows
        finally:
            self._partial.pop(name, None)
            self._busy.discard(("table", name))

    # ---- evaluation
    def lookup(self, table, key, key_col, value_col, default):
        index = self._lookup_index.get((table, key_col))
        if index is None:
            index = {}
            if table not in self.tables:     # a hidden grouped helper frame, addressed by its title
                frame = next((f for f in self.frames.values() if f["title"] == table), None)
                if frame is None:
                    raise PlanError(f"unknown table {table!r}")
                rows = self.frame_rows(frame["id"])
            else:
                rows = self._table_rows.get(table) or self._partial.get(table) or self.table_rows(table)
            for row in rows:
                if key_col not in row:
                    raise EvalError("#REF!", f"{table}[{key_col}] is not computed yet (circular lookup)")
                index.setdefault(row[key_col], row)
            self._lookup_index[(table, key_col)] = index
        row = index.get(key) if key is not None else None
        if row is None:
            if default is None:
                raise EvalError("#N/A", f"{table}[{key_col}] has no {key!r}")
            return default
        return row[value_col]

    def ev(self, node, env):
        tag = node[0]
        if tag == "n" or tag == "s" or tag == "b":
            return node[1]
        if tag == "d":
            return date.fromisoformat(node[1])
        if tag == "null":
            return None
        if tag == "c":
            name = node[1]
            if env.row is not None and name in env.row:
                return env.row[name]
            if env.cols is not None and name in env.cols:
                return env.cols[name]
            raise EvalError("#NAME?", f"column {name!r}")
        if tag == "key":
            return env.key
        if tag == "period":
            return env.period
        if tag == "prev":
            return env.prev
        if tag in ("p", "pt", "pd"):
            if node[1] not in self.params:
                raise EvalError("#NAME?", f"parameter {node[1]!r}")
            return self.params[node[1]]
        if tag == "neg":
            return -_num(self.ev(node[1], env))
        if tag == "op":
            return self._op(node, env)
        if tag == "f":
            return self._fn(node, env)
        if tag == "lk":
            default = None if node[5] is None else self.ev(node[5], env)
            return self.lookup(node[1], self.ev(node[2], env), node[3], node[4], default)
        if tag == "agg":
            return self._agg(node, env)
        if tag in ("cumsum", "cummin"):
            if env.cum is None or id(node) not in env.cum:
                raise EvalError("#VALUE!", f"{tag} outside a keyed or ledger frame column")
            return env.cum[id(node)][env.idx]
        raise PlanError(f"unknown IR node {tag!r}")

    def _op(self, node, env):
        sym = node[1]
        a, b = self.ev(node[2], env), self.ev(node[3], env)
        if sym in COMPARE:
            return (_compare_crit if env.crit else _compare)(sym, a, b)
        a, b = _num(a), _num(b)
        if sym == "+":
            return a + b
        if sym == "-":
            return a - b
        if sym == "*":
            return a * b
        if sym == "/":
            if b == 0:
                raise EvalError("#DIV/0!")
            return a / b
        try:
            return a ** b
        except (OverflowError, ZeroDivisionError) as error:
            raise EvalError("#NUM!") from error

    def _fn(self, node, env):
        name, args = node[1], node[2:]
        if name == "IF":
            if len(args) != 3:
                raise PlanError("IF needs three arguments")
            return self.ev(args[1] if _truthy(self.ev(args[0], env)) else args[2], env)
        if name == "IFERROR":
            try:
                return self.ev(args[0], env)
            except Unmodelled:
                raise
            except EvalError as error:
                if error.code not in CAUGHT_ERRORS:
                    raise
                return self.ev(args[1], env)
        values = [self.ev(a, env) for a in args]
        if name in ("AND", "OR") and any(v is None for v in values):
            # Excel's AND/OR skip references to blank cells (and are #VALUE! when nothing is left);
            # Sigma's three-valued logic differs again, so refuse rather than guess.
            raise Unmodelled("#VALUE!", f"{name} over a blank cell is not supported")
        if name == "AND":
            return all(_truthy(v) for v in values)
        if name == "OR":
            return any(_truthy(v) for v in values)
        if name == "NOT":
            return not _truthy(values[0])
        if name == "SUM":
            return sum(_num(v) for v in values)
        if name in ("MIN", "MAX"):
            if any(v is None for v in values):
                # Excel's MIN/MAX ignore a reference to a blank cell (it is not 0); Sigma's
                # Least/Greatest follow SQL NULL rules. Neither equals this evaluator's 0.
                raise Unmodelled("#VALUE!", f"{name} over a blank cell is not supported")
            numbers = [_num(v) for v in values]
            return min(numbers) if name == "MIN" else max(numbers)
        if name == "ABS":
            return abs(_num(values[0]))
        if name == "MOD":
            divisor = _num(values[1])
            if divisor == 0:
                raise EvalError("#DIV/0!")
            return _num(values[0]) % divisor      # sign of the divisor, like Excel
        if name == "ADDDAYS":
            if not isinstance(values[0], date):
                raise EvalError("#VALUE!", repr(values[0]))
            return values[0] + timedelta(days=int(_num(values[1])))
        if name == "CONCAT":
            if not all(isinstance(v, str) for v in values):
                raise EvalError("#VALUE!", "CONCAT joins text only")
            return "".join(values)
        if name == "ROUND":
            digits = int(_num(values[1]))
            factor = 10 ** digits
            scaled = abs(_num(values[0])) * factor
            return math.copysign(math.floor(scaled + 0.5) / factor, _num(values[0]))
        if name == "CEILING":
            number, multiple = _num(values[0]), _num(values[1])
            if multiple == 0:
                return 0
            if number > 0 > multiple:
                raise EvalError("#NUM!", "CEILING of a positive number to a negative multiple")
            return math.ceil(number / multiple) * multiple
        if name == "EOMONTH0":
            return eomonth0(values[0])
        if name == "MONTH":
            if not isinstance(values[0], date):
                raise EvalError("#VALUE!", repr(values[0]))
            return values[0].month
        if name == "YEAR":
            if not isinstance(values[0], date):
                raise EvalError("#VALUE!", repr(values[0]))
            return values[0].year
        raise PlanError(f"unsupported function {name}")

    def _agg(self, node, env):
        _, fn, column, flt = node
        if env.group is None:
            raise EvalError("#VALUE!", "aggregate outside a summary frame")
        picked = []
        for row in env.group:
            if flt is not None and not _truthy(self.ev(flt, Env(row=row, crit=True))):
                continue
            picked.append(row)
        if fn == "count" and column is None:
            return len(picked)
        values = [row[column] for row in picked]
        numbers = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
        if fn == "count":
            return len(numbers)
        if fn == "sum":
            return sum(numbers)
        if fn == "avg":
            if not numbers:
                raise EvalError("#DIV/0!")
            return sum(numbers) / len(numbers)
        if not numbers:
            return 0
        return min(numbers) if fn == "min" else max(numbers)

    # ---- frames
    def frame_rows(self, frame_id):
        if frame_id in self._frame_rows:
            return self._frame_rows[frame_id]
        frame = self.frames.get(frame_id)
        if frame is None:
            raise PlanError(f"unknown frame {frame_id!r}")
        if ("frame", frame_id) in self._busy:
            raise PlanError(f"circular frames involving {frame_id!r}")
        self._busy.add(("frame", frame_id))
        try:
            kind = frame["kind"]
            builder = {"expansion": self._expansion, "keyed": self._keyed, "summary": self._summary,
                       "ledger": self._ledger}.get(kind)
            if builder is None:
                raise PlanError(f"frame {frame_id!r}: unknown kind {kind!r}")
            rows = builder(frame)
            self._frame_rows[frame_id] = rows
            return rows
        finally:
            self._busy.discard(("frame", frame_id))

    def source_rows(self, source):
        if source["kind"] == "table":
            return self.table_rows(source["name"])
        return self.frame_rows(source["name"])

    def _eval_columns(self, frame, rows, env_for):
        """Evaluate the frame's columns across ``rows`` (column-major so running sums see all rows)."""
        for column in frame["columns"]:
            expr = column["expr"]
            cum = {}
            for node in walk(expr):
                if node[0] in ("cumsum", "cummin"):
                    series = {}
                    for number, row in enumerate(rows):
                        env = env_for(row, number)
                        env.cum, env.idx = cum, number
                        part = row[node[2]] if node[2] is not None else None
                        series.setdefault(part, []).append(
                            (row[node[3]], number, _num(self.ev(node[1], env))))
                    values = [None] * len(rows)
                    for entries in series.values():
                        running = None
                        for _, number, delta in sorted(entries, key=lambda item: (item[0], item[1])):
                            if node[0] == "cumsum":
                                running = delta if running is None else running + delta
                            else:
                                running = delta if running is None else min(running, delta)
                            values[number] = running
                    cum[id(node)] = values
            for number, row in enumerate(rows):
                env = env_for(row, number)
                env.cum, env.idx = cum, number
                try:
                    value = self.ev(expr, env)
                    row[column["name"]] = 0 if value is None else value
                except EvalError as error:
                    raise EvalError(error.code, f"frame {frame['title']!r} column {column['name']!r} "
                                    f"row {number + 1}: {error}") from error

    def _expansion(self, frame):
        axis = frame["axis"]
        axis_values = [parse_date(v) if axis["type"] == "datetime" else v for v in axis["values"]]
        rows = []
        for base_index, base in enumerate(self.table_rows(frame["base"])):
            for value in axis_values:
                row = dict(base)
                row[axis["name"]] = value
                row["__base"] = base_index
                rows.append(row)
        self._eval_columns(frame, rows, lambda row, number: Env(row=row, period=row[axis["name"]]))
        return rows

    def _keyed(self, frame):
        rows = [dict(r, __base=i) for i, r in enumerate(self.table_rows(frame["table"]))]
        by_key = {}
        for number, row in enumerate(rows):
            by_key.setdefault(row[frame["rowKey"]], []).append(number)
        prev_of = {}
        for numbers in by_key.values():
            ordered = sorted(numbers, key=lambda n: (rows[n][frame["periodKey"]], n))
            for earlier, later in zip(ordered, ordered[1:]):
                prev_of[later] = earlier
        frame_name = frame["measure"]

        def env_for(row, number):
            earlier = prev_of.get(number)
            return Env(row=row, key=row[frame["rowKey"]], period=row[frame["periodKey"]],
                       prev=rows[earlier].get(frame_name) if earlier is not None else None)
        self._eval_columns(frame, rows, env_for)
        return rows

    def _ledger(self, frame):
        """One row per period of the spine table; columns are evaluated in dependency order."""
        axis = frame["axis"]
        rows = [dict(r, __base=i) for i, r in enumerate(self.table_rows(frame["spine"]))]
        rows.sort(key=lambda r: r[axis])
        names = {c["name"] for c in frame["columns"]}
        deps = {c["name"]: {n[1] for n in walk(c["expr"]) if n[0] == "c" and n[1] in names}
                for c in frame["columns"]}
        ordered, done, remaining = [], set(), list(frame["columns"])
        while remaining:
            ready = [c for c in remaining if deps[c["name"]] <= done]
            if not ready:
                raise PlanError(f"ledger {frame['title']!r}: calculated columns reference each other in a "
                                f"cycle ({', '.join(c['name'] for c in remaining)})")
            for column in ready:
                ordered.append(column)
                done.add(column["name"])
                remaining.remove(column)
        self._eval_columns({"title": frame["title"], "columns": ordered}, rows,
                           lambda row, number: Env(row=row, period=row[axis], idx=number))
        return rows

    def _summary(self, frame):
        source_rows = self.source_rows(frame["source"])
        by = frame["by"]
        groups = OrderedDict()
        for row in source_rows:
            groups.setdefault(row[by], []).append(row)
        ordered = [k for k in frame["labels"] if k in groups] + sorted(
            (k for k in groups if k not in frame["labels"]), key=lambda k: (str(type(k)), str(k)))
        rows = []
        for key in ordered:
            cols = {by: key}
            for column in frame["columns"]:
                try:
                    value = self.ev(column["expr"], Env(cols=cols, key=key, group=groups[key]))
                    cols[column["name"]] = 0 if value is None else value
                except EvalError as error:
                    raise EvalError(error.code, f"frame {frame['title']!r} {key!r} column "
                                    f"{column['name']!r}: {error}") from error
            rows.append(cols)
        return rows

    # ---- addressing (checks)
    def cell(self, frame, column, kind, sel):
        """Value of one addressed check cell; raises KeyError when the row does not exist."""
        if kind == "table":
            return self.table_rows(frame)[sel[0]][column]
        if kind == "summary-total":
            return sum(_num(row[column]) for row in self.frame_rows(frame))
        if kind == "summary":
            for row in self.frame_rows(frame):
                if row[self.frames[frame]["by"]] == sel[0]:
                    return row[column]
            raise KeyError(f"summary row {sel[0]!r} does not exist (a grouped table would not show it)")
        rows = self.frame_rows(frame)
        spec = self.frames[frame]
        if kind == "ledger":
            for row in rows:
                if row[spec["axis"]] == sel[0]:
                    return row[column]
            raise KeyError(f"ledger period {sel[0]!r} does not exist in {spec['title']!r}")
        if kind == "ledger-total":
            return sum(_num(row[column]) for row in rows)
        if kind == "ledger-caption":
            return self.ev(spec["captions"][sel[0]]["expr"], Env())
        if kind == "keyed":
            return rows[sel[0]][column]
        if kind in ("expansion", "rowtotal"):
            axis = spec["axis"]
            mine = [r for r in rows if r["__base"] == sel[0]]
            if kind == "rowtotal":
                return sum(_num(r[column]) for r in mine)
            wanted = parse_date(sel[1]) if axis["type"] == "datetime" else sel[1]
            for row in mine:
                if row[axis["name"]] == wanted:
                    return row[column]
            raise KeyError(f"expansion row {sel!r} does not exist")
        raise PlanError(f"unknown check kind {kind!r}")


def compute(plan):
    """Return a ready ``Model`` after evaluating every table and frame (raises on any error)."""
    model = Model(plan)
    for name in model.tables:
        model.table_rows(name)
    for frame_id in model.frames:
        model.frame_rows(frame_id)
    return model


# The oracle recomputes the same arithmetic Excel did, so only float noise (summation order, the
# 15-17 significant digits Excel caches) may remain. ``tolerance=None`` selects this tight default; a
# number passed explicitly (``--tolerance``) is an absolute limit and overrides it.
ABS_FLOOR, REL_LIMIT = 1e-6, 1e-12


def tolerance_for(want, tolerance=None):
    if tolerance is not None:
        return tolerance
    return max(ABS_FLOOR, REL_LIMIT * abs(want))


def values_match(got, want, tolerance=None):
    """Numbers within the tolerance (tight relative/absolute default); text/bool/date exact; blank == empty string."""
    if isinstance(want, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", want) and isinstance(got, date):
        return got.isoformat() == want
    if got is None:
        got = ""
    if want is None:
        want = ""
    if isinstance(got, bool) or isinstance(want, bool):
        return got == want
    if isinstance(got, (int, float)) and isinstance(want, (int, float)):
        return abs(got - want) <= tolerance_for(want, tolerance)
    if isinstance(got, date) and isinstance(want, str):
        return got.isoformat() == want
    return got == want


def check_cells(plan, model, tolerance=None):
    """Compare every recorded source cell with the plan's recomputation.

    Returns ``(checked, problems)``; ``problems`` names sheet!cell, the expected cached
    value and the plan's value.
    """
    checked, problems = 0, []
    for group in plan.get("checks", []):
        for address, sel, want in group["cells"]:
            checked += 1
            try:
                got = model.cell(group["frame"], group["column"], group["kind"], sel)
            except (KeyError, IndexError) as error:
                problems.append(f"{address} ({group['frame']}.{group['column']}): {error}")
                continue
            if not values_match(got, want, tolerance):
                shown = iso(got)
                hint = ""
                if want in (None, "") and got not in (None, ""):
                    hint = " (the workbook has no cached value here; open it in Excel, recalculate and save)"
                problems.append(f"{address} ({group['frame']}.{group['column']}): plan {shown!r} "
                                f"vs workbook {want!r}{hint}")
    return checked, problems


def parity_lines(plan, model, tolerance=None):
    """``(ok, lines)``: headline parity of the plan's summary targets.

    One line per summary frame (cells matched / total), one per numeric column (plan total vs
    source total), and a FAIL line for every target that differs.
    """
    targets = plan.get("parityTargets") or []
    if not targets:
        return True, ["LOCAL PARITY: not checked - plan has no parityTargets"]
    ok, lines, frames = True, [], OrderedDict()
    for target in targets:
        group = frames.setdefault(target["frame"], {"label": target["label"].split(" / ")[0], "total": 0,
                                                    "passed": 0, "columns": OrderedDict()})
        group["total"] += 1
        try:
            got = model.cell(target["frame"], target["column"], target.get("kind", "summary"), target["sel"])
        except (KeyError, IndexError) as error:
            ok = False
            lines.append(f"PARITY FAIL {target['label']}: {error}")
            continue
        passed = values_match(got, target["value"], tolerance)
        group["passed"] += passed
        if not passed:
            ok = False
            lines.append(f"PARITY FAIL {target['label']}: plan={iso(got)!r} source={target['value']!r}")
        frame = (model.frames.get(target["frame"]) or {})
        fmt = next((c.get("format") for c in frame.get("columns", []) if c["name"] == target["column"]), None)
        if isinstance(got, (int, float)) and not isinstance(got, bool) and not ef.is_percent_format(fmt) \
                and target.get("kind") != "ledger" and isinstance(target["value"], (int, float)):
            column = group["columns"].setdefault(target["column"], [0.0, 0.0])
            column[0] += got
            column[1] += target["value"]
    heads = []
    for group in frames.values():
        state = "PASS" if group["passed"] == group["total"] else "FAIL"
        heads.append(f"PARITY {state} {group['label']}: {group['passed']}/{group['total']} summary cells")
    lines = heads + lines
    for frame_id, group in frames.items():
        for column, (got, want) in group["columns"].items():
            state = "PASS" if abs(got - want) <= tolerance_for(want, tolerance) * max(1, group["total"]) else "FAIL"
            if state == "FAIL":
                ok = False
            lines.append(f"PARITY {state} {group['label']} / {column} total: plan={got:,.2f} "
                         f"source={want:,.2f} diff={got - want:,.4f}")
    return ok, lines


def validate_plan(plan):
    """Structural validation of a (possibly hand-edited) plan; raises ``PlanError``."""
    if plan.get("archetype") != ARCHETYPE:
        raise PlanError(f"not a table-report plan (archetype {plan.get('archetype')!r})")
    if plan.get("planVersion") != PLAN_VERSION:
        raise PlanError(f"unsupported planVersion {plan.get('planVersion')!r}")
    names = [t["name"] for t in plan.get("tables", [])]
    if len(set(names)) != len(names):
        raise PlanError("duplicate table names")
    if PARAM_TABLE in names:
        raise PlanError(f"{PARAM_TABLE!r} is reserved for plan.params")
    titles = names + [f["title"] for f in plan.get("frames", [])] + [PARAM_TABLE]
    if len(set(titles)) != len(titles):
        raise PlanError("table and frame titles must be unique")
    for table in plan.get("tables", []):
        cols = [c["name"] for c in table["columns"]]
        if len(set(cols)) != len(cols):
            raise PlanError(f"table {table['name']!r}: duplicate column names")
        for column in table["columns"]:
            if column.get("type") not in COLUMN_TYPES:
                raise PlanError(f"table {table['name']!r} column {column['name']!r}: bad type")
            if re.search(r"[\[\]/]", column["name"]):
                raise PlanError(f"table {table['name']!r}: column name {column['name']!r} contains [ ] or /")
            if "expr" in column:
                validate_ir(column["expr"], f"{table['name']}.{column['name']}")
    for frame in plan.get("frames", []):
        if frame["kind"] not in ("expansion", "keyed", "summary", "ledger"):
            raise PlanError(f"frame {frame.get('id')!r}: unknown kind")
        for column in frame["columns"]:
            validate_ir(column["expr"], f"{frame['title']}.{column['name']}")
        for caption in frame.get("captions", []):
            validate_ir(caption["expr"], f"{frame['title']} caption {caption.get('cell')}")
    _validate_presentation(plan)
    return True


def _validate_presentation(plan):
    """The optional ``presentation`` block must point at real frames and carry whole series."""
    pres = plan.get("presentation")
    if not pres:
        return
    frames = {f["id"]: f for f in plan.get("frames", [])}
    for kpi in pres.get("kpis", []):
        if kpi.get("frame") not in frames or kpi.get("kind") not in ("total", "ratio"):
            raise PlanError(f"presentation KPI {kpi.get('label')!r} does not name a frame and a kind")
        names = {c["name"] for c in frames[kpi["frame"]]["columns"]}
        if not {kpi.get("column"), kpi.get("num"), kpi.get("den")} - {None} <= names:
            raise PlanError(f"presentation KPI {kpi.get('label')!r} reads a column its frame does not have")
    for chart in pres.get("charts", []):
        if chart.get("frame") not in frames or chart.get("kind") not in ("bar", "line", "donut", "pie"):
            raise PlanError(f"presentation chart {chart.get('title')!r} does not name a frame and a kind")
        for series in chart["series"]:
            if not (len(series["cells"]) == len(series["sels"]) == len(series["values"])
                    == len(chart["categories"])):
                raise PlanError(f"presentation chart {chart['title']!r}: series {series['title']!r} does not "
                                "have one verified value per category")


# ===================================================================== Sigma
SIGMA_STRING = json.dumps


def slug(text, sep="_"):
    return re.sub(r"[^a-z0-9]+", sep, str(text).lower()).strip(sep) or "x"


def alias(name):
    return re.sub(r"[^A-Za-z0-9]+", "_", str(name)).strip("_").upper() or "COL"


def q(text):
    return "'" + str(text).replace("'", "''") + "'"


def sigma_number(value):
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


BLANK_DEFAULT = {"number": "0", "text": '""', "boolean": "False", "datetime": 'Date("1900-01-01")'}
NUMERIC_FUNCTIONS = {"SUM", "ABS", "ROUND", "CEILING", "MOD", "MONTH", "YEAR"}


class Emitter:
    """Renders IR as a Sigma formula. ``col``/``key``/``period``/``prev`` are element-specific.

    Excel and Sigma disagree about blanks: Excel reads a blank cell as 0 / "" / FALSE depending on
    the other operand, Sigma reads it as NULL (which poisons arithmetic and makes ``<>`` unknown).
    The emitter therefore wraps every operand that *can* be blank (``blankable`` columns, lookups
    into blank-capable columns, IF branches returning them) in ``Coalesce(.., 0 | "" | False)``
    exactly where Excel would have coerced it, so the emitted formula agrees with the Python oracle
    (``Model``). Columns without a blank in the plan data are emitted unchanged.
    """

    def __init__(self, col, key=None, period=None, prev=None, coltype=None, count_ref=None,
                 agg_source=None, agg_col=None, blankable=None, tables=None, wrap_lookup=False,
                 utf8_strings=False):
        self.col, self.key_ref, self.period_ref, self.prev_ref = col, key, period, prev
        self.wrap_lookup = wrap_lookup      # a lookup of the group key inside a chart series: Max(Lookup(..))
        self.utf8_strings = utf8_strings    # chart / KPI feeds: labels as typed, not \\uXXXX escapes
        self.coltype, self.count_ref = coltype, count_ref
        self.agg_source = agg_source  # Emitter used inside aggregate filters (source columns)
        self.agg_col = agg_col        # name -> source column ref for aggregate values
        self._blankable = blankable or (lambda name: False)
        self.own_blank = {}           # calculated columns emitted so far -> may be blank
        self.tables = tables or {"types": {}, "blank": {}}   # per plan table: column types / blank-capable

    # ---- shape knowledge
    def blankable(self, name):
        return self.own_blank[name] if name in self.own_blank else bool(self._blankable(name))

    def column_formula(self, name, expr):
        """Formula of a calculated column. A formula cell never holds a blank in Excel (a result that is
        a reference to a blank cell shows 0), so a possibly-blank result is ``Coalesce(.., 0)``."""
        text = self.emit(expr)
        if self.may_null(expr):
            text = f"Coalesce({text}, 0)"
        self.own_blank[name] = False
        return text

    def may_null(self, node):
        """True when the emitted Sigma value can be NULL where Excel would show a blank (0 / "")."""
        tag = node[0]
        if tag == "null" or tag == "prev":
            return True
        if tag == "c":
            return self.blankable(node[1])
        if tag == "lk":
            return bool(self.tables["blank"].get(node[1], {}).get(node[4]))
        if tag == "f":
            name, args = node[1], node[2:]
            if name == "IF":
                return self.may_null(args[1]) or self.may_null(args[2])
            if name == "IFERROR":
                return self.may_null(args[0]) or self.may_null(args[1])
            if name in ("EOMONTH0", "ADDDAYS"):
                return self.may_null(args[0])
        return False

    def type_of(self, node):
        tag = node[0]
        if tag == "s":
            return "text"
        if tag in ("n",):
            return "number"
        if tag == "d":
            return "datetime"
        if tag == "b":
            return "boolean"
        if tag == "c" and self.coltype:
            return self.coltype(node[1])
        if tag == "lk":
            found = self.tables["types"].get(node[1], {}).get(node[4])
            if found:
                return found
            return "text" if node[5] is not None and node[5][0] == "s" else None
        if tag == "pt":
            return "text"
        if tag == "p":
            return "number"
        if tag == "pd":
            return "datetime"
        if tag == "neg" or tag == "agg" or tag in ("cumsum", "cummin"):
            return "number"
        if tag == "op":
            return "number" if node[1] in ARITH else "boolean"
        if tag == "f":
            name, args = node[1], node[2:]
            if name in ("AND", "OR", "NOT"):
                return "boolean"
            if name in ("IF",):
                return self.type_of(args[1]) or self.type_of(args[2])
            if name == "IFERROR":
                return self.type_of(args[0]) or self.type_of(args[1])
            if name in ("MIN", "MAX") or name in NUMERIC_FUNCTIONS:
                return "number"
            if name in ("EOMONTH0", "ADDDAYS"):
                return "datetime"
            if name == "CONCAT":
                return "text"
        return None

    def num(self, node):
        """Emit ``node`` as an arithmetic operand (a blank is 0)."""
        text = self.emit(node)
        return f"Coalesce({text}, 0)" if self.may_null(node) else text

    def operand(self, node, other):
        """Emit ``node`` as one side of a comparison with ``other`` (a blank takes the other's type)."""
        text = self.emit(node)
        if not self.may_null(node):
            return text
        kind = self.type_of(other) or self.type_of(node)
        if kind not in BLANK_DEFAULT:
            raise PlanError(f"cannot tell whether a blank in {text} should read as 0 or \"\" in this "
                            "comparison; compare it with a literal or a typed column")
        return f"Coalesce({text}, {BLANK_DEFAULT[kind]})"

    def emit(self, node):
        tag = node[0]
        if tag == "n":
            return sigma_number(node[1])
        if tag == "s":
            return json.dumps(node[1], ensure_ascii=False) if self.utf8_strings else SIGMA_STRING(node[1])
        if tag == "b":
            return "True" if node[1] else "False"
        if tag == "d":
            return f'Date("{node[1]}")'
        if tag == "null":
            return "Null()"
        if tag == "c":
            return self.col(node[1])
        if tag == "key":
            if self.key_ref is None:
                raise PlanError("key reference outside a keyed/summary frame")
            return self.key_ref
        if tag == "period":
            if self.period_ref is None:
                raise PlanError("period reference outside a keyed frame")
            return self.period_ref
        if tag == "prev":
            raise PlanError("a previous-period reference must be rewritten into a running sum")
        if tag in ("p", "pt", "pd"):
            value_col = {"p": "Value", "pt": "Text Value", "pd": "Date Value"}[tag]
            return (f'Lookup([{PARAM_TABLE}/{value_col}], {SIGMA_STRING(node[1])}, [{PARAM_TABLE}/Parameter])')
        if tag == "neg":
            return f"(-{self.num(node[1])})"
        if tag == "op":
            return self._op(node)
        if tag == "f":
            return self._fn(node)
        if tag == "lk":
            _, table, key, key_col, value_col, default = node
            key_text = self.emit(key)
            lookup = f"Lookup([{table}/{value_col}], {key_text}, [{table}/{key_col}])"
            if self.wrap_lookup:
                lookup = f"Max({lookup})"
            if default is None:
                return lookup
            if self.may_null(node):
                # the value column has blanks: a key that IS found with a blank value is Excel's 0/"",
                # not the default, so "found" is tested on the key column itself
                return f"If({self._not_found(node)}, {self.emit(default)}, {lookup})"
            return f"Coalesce({lookup}, {self.emit(default)})"
        if tag == "agg":
            return self._agg(node)
        if tag == "cumsum":
            return f"CumulativeSum({self.emit(node[1])})"
        if tag == "cummin":
            return f"CumulativeMin({self.emit(node[1])})"
        raise PlanError(f"cannot emit IR node {tag!r}")

    def _not_found(self, node):
        _, table, key, key_col, _, _ = node
        return f"IsNull(Lookup([{table}/{key_col}], {self.emit(key)}, [{table}/{key_col}]))"

    def emit_crit(self, node):
        """Emit an aggregate criterion with Excel *criteria* semantics (a blank matches only ``<>``)."""
        if node[0] == "f" and node[1] == "AND":
            return "(" + " and ".join(self.emit_crit(a) for a in node[2:]) + ")"
        if node[0] == "op" and node[1] in COMPARE:
            return self._op(node, crit=True)
        return self.emit(node)

    def _op(self, node, crit=False):
        sym, a, b = node[1], node[2], node[3]
        if sym in ("=", "<>"):
            for left, right in ((a, b), (b, a)):
                if right == ["s", ""]:
                    kind = self.type_of(left)
                    left_text = self.emit(left)
                    if kind == "text":
                        test = f"(IsNull({left_text}) or {left_text} = \"\")"
                    else:
                        test = f"IsNull({left_text})"
                    return test if sym == "=" else f"Not({test})"
        if sym == "^":
            return f"Power({self.num(a)}, {self.num(b)})"
        symbol = {"<>": "!="}.get(sym, sym)
        if sym in COMPARE:
            if crit:
                left, right = self.emit(a), self.emit(b)
                if sym == "<>" and a[0] == "c" and self.blankable(a[1]):
                    return f"(IsNull({left}) or {left} != {right})"
                return f"({left} {symbol} {right})"
            return f"({self.operand(a, b)} {symbol} {self.operand(b, a)})"
        return f"({self.num(a)} {symbol} {self.num(b)})"

    def _fn(self, node):
        name, args = node[1], node[2:]
        if name == "IFERROR":
            guard = self._guard(args[0])
            if guard is None:
                return self.emit(args[0])
            return f"If({guard}, {self.emit(args[1])}, {self.emit(args[0])})"
        if name == "IF":
            cond = self.emit(args[0])
            if self.may_null(args[0]):
                cond = f"Coalesce({cond}, False)"
            return f"If({cond}, {self.emit(args[1])}, {self.emit(args[2])})"
        if name in ("MIN", "MAX"):
            return f"{'Least' if name == 'MIN' else 'Greatest'}({', '.join(self.emit(a) for a in args)})"
        if name == "SUM":
            return "(" + " + ".join(self.num(a) for a in args) + ")"
        if name == "ABS":
            return f"Abs({self.num(args[0])})"
        if name == "MOD":
            # Excel's MOD takes the sign of the divisor; SQL/Sigma Mod may take the dividend's.
            # Mod(Mod(a, b) + b, b) is the divisor-signed remainder under either convention.
            a, b = self.num(args[0]), self.num(args[1])
            return f"Mod(Mod({a}, {b}) + {b}, {b})"
        if name == "ROUND":
            return f"Round({self.num(args[0])}, {self.num(args[1])})"
        if name == "CEILING":
            # Excel CEILING(x, 0) is 0; guard the division instead of emitting x / 0.
            a, b = self.num(args[0]), self.num(args[1])
            return f"If({b} = 0, 0, (Ceiling({a} / {b}) * {b}))"
        out = [self.emit(a) for a in args]
        if name == "AND":
            return "(" + " and ".join(out) + ")"
        if name == "OR":
            return "(" + " or ".join(out) + ")"
        if name == "NOT":
            cond = out[0]
            if self.may_null(args[0]):
                cond = f"Coalesce({cond}, False)"
            return f"Not({cond})"
        if name == "ADDDAYS":
            return f'DateAdd("day", {self.num(args[1])}, {out[0]})'
        if name == "CONCAT":
            return f"Concat({', '.join(out)})"
        if name == "EOMONTH0":
            return f'DateAdd("day", -1, DateAdd("month", 1, DateTrunc("month", {out[0]})))'
        if name == "MONTH":
            return f'DatePart("month", {out[0]})'
        if name == "YEAR":
            return f'DatePart("year", {out[0]})'
        raise PlanError(f"cannot emit function {name}")

    def _guard(self, node):
        """Boolean Sigma text that is true exactly when evaluating ``node`` raises one of the errors
        ``IFERROR`` catches (``CAUGHT_ERRORS``: a zero/blank denominator, a missed exact lookup, an
        AVERAGEIFS over nothing); ``None`` when it cannot. Branches an ``IF`` does not take, and the
        operands of a nested ``IFERROR``, contribute only when they are actually evaluated."""
        tag = node[0]
        parts = []
        if tag == "op":
            parts += [self._guard(node[2]), self._guard(node[3])]
            if node[1] == "/" and not (node[3][0] == "n" and node[3][1] != 0):
                den = self.emit(node[3])
                parts.append(f"IsNull({den}) or {den} = 0")
        elif tag == "neg":
            parts.append(self._guard(node[1]))
        elif tag == "lk":
            parts.append(self._guard(node[2]))
            if node[5] is None:
                parts.append(self._not_found(node))
            else:
                parts.append(self._guard(node[5]))
        elif tag == "agg":
            if node[1] == "avg":
                parts.append(f"IsNull({self.emit(node)})")
        elif tag == "f":
            name, args = node[1], node[2:]
            if name == "IF":
                parts.append(self._guard(args[0]))
                then_g, else_g = self._guard(args[1]), self._guard(args[2])
                if then_g or else_g:
                    parts.append(f"If({self.emit(args[0])}, {then_g or 'False'}, {else_g or 'False'})")
            elif name == "IFERROR":
                inner, fallback = self._guard(args[0]), self._guard(args[1])
                if inner and fallback:
                    parts.append(f"If({inner}, {fallback}, False)")
            else:
                parts += [self._guard(a) for a in args]
        parts = [p for p in parts if p]
        return " or ".join(parts) if parts else None

    def _agg(self, node):
        _, fn, column, flt = node
        source = self.agg_source
        if source is None:
            raise PlanError("aggregate outside a summary frame")
        value = self.agg_col(column) if column is not None else None
        cond = source.emit_crit(flt) if flt is not None else None
        # Excel's SUMIFS/COUNTIFS/MINIFS/MAXIFS give 0 for an empty match, Sigma's conditional
        # aggregates give NULL (which then poisons any arithmetic), so they are wrapped in
        # Coalesce(.., 0). AVERAGEIFS stays unwrapped: an empty match is #DIV/0! in Excel.
        if fn == "count":
            return f"Coalesce(CountIf({cond}), 0)" if cond else f"Coalesce(Count({self.count_ref}), 0)"
        name = {"sum": "Sum", "avg": "Avg", "min": "Min", "max": "Max"}[fn]
        text = f"{name}If({value}, {cond})" if cond else f"{name}({value})"
        return text if fn == "avg" else f"Coalesce({text}, 0)"


# ---- tables as the builders see them (plan tables + the parameter table)
def param_table(plan):
    """Pseudo-table holding ``plan.params`` (name, value, source cell)."""
    if not plan.get("params"):
        return None
    kinds = {p.get("kind", "number") for p in plan["params"]}
    numbers = [p["value"] for p in plan["params"] if p.get("kind", "number") == "number"]
    columns = [{"name": "Parameter", "type": "text"},
               {"name": "Value", "type": "number", "format": ef.plain_number(ef.number_decimals(numbers))},
               {"name": "Cell", "type": "text"}]
    if "text" in kinds:      # a text parameter (a selected scenario) lives in its own typed column
        columns.append({"name": "Text Value", "type": "text"})
    if "date" in kinds:
        columns.append({"name": "Date Value", "type": "datetime"})
    rows = []
    for p in plan["params"]:
        kind = p.get("kind", "number")
        row = [p["name"], p["value"] if kind == "number" else None, p.get("cell", "")]
        if "text" in kinds:
            row.append(p["value"] if kind == "text" else None)
        if "date" in kinds:
            row.append(p["value"] if kind == "date" else None)
        rows.append(row)
    return {"name": PARAM_TABLE, "sheet": "", "editable": bool(plan.get("parametersEditable", True)),
            "columns": columns, "rows": rows}


def all_tables(plan):
    extra = param_table(plan)
    return list(plan["tables"]) + ([extra] if extra else [])


def is_editable(table):
    """Editable tables (and empty ones, which VALUES cannot hold) become input tables."""
    return bool(table.get("editable")) or not table.get("rows")


def literal_columns(table):
    return [c for c in table["columns"] if "expr" not in c]


# ---- DM (read-only data tables as VALUES custom SQL)
def sql_literal(value, kind):
    if value is None:
        return "NULL::" + {"number": "number", "text": "varchar", "datetime": "date",
                           "boolean": "boolean"}[kind]
    if kind == "text":
        return q(value)
    if kind == "datetime":
        return f"{q(iso(value) if isinstance(value, date) else value)}::date"
    if kind == "boolean":
        return "TRUE" if value else "FALSE"
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise PlanError(f"not a finite number: {value!r}")
    return sigma_number(value)


def values_sql(rows, columns):
    """``select * from (values ...) as t(ALIASES)``; ``columns`` = [(alias, type)]."""
    if not rows:
        raise PlanError("a VALUES table needs at least one row (use an input table for empty tables)")
    body = ",\n".join("(" + ",".join(sql_literal(v, c[1]) for v, c in zip(row, columns)) + ")"
                      for row in rows)
    return f"select * from (values\n{body}\n) as t({', '.join(c[0] for c in columns)})"


def dm_tables(plan):
    """Read-only tables (and every frame axis spine) hosted in the data model."""
    out = []
    for table in all_tables(plan):
        if not is_editable(table):
            out.append({"key": f"t:{table['name']}", "label": table["name"], "table": table})
    for frame in plan["frames"]:
        if frame["kind"] == "expansion":
            axis = frame["axis"]
            fake = {"name": f"{frame['title']} Axis",
                    "columns": [{"name": axis["name"], "type": axis["type"]}],
                    "rows": [[v] for v in axis["values"]]}
            out.append({"key": f"axis:{frame['id']}", "label": fake["name"], "table": fake})
    return out


def dm_signature(table):
    return tuple(c["name"] for c in literal_columns(table))


def build_dm_spec(plan):
    """Data-model spec, or ``None`` when every table is editable (nothing read-only to host)."""
    items = dm_tables(plan)
    if not items:
        return None
    sigs = {}
    for item in items:
        sig = dm_signature(item["table"])
        if sig in sigs:
            raise PlanError(f"{item['label']!r} and {sigs[sig]!r} have identical column sets; the live "
                            "data-model readback could not tell them apart - rename a column")
        sigs[sig] = item["label"]
    conn = plan["connectionId"]
    elements = []
    for number, item in enumerate(items, 1):
        table = item["table"]
        cols = literal_columns(table)
        used, aliases = set(), []
        for column in cols:
            base = alias(column["name"])
            candidate, bump = base, 2
            while candidate in used:
                candidate, bump = f"{base}_{bump}", bump + 1
            used.add(candidate)
            aliases.append(candidate)
        statement = values_sql(
            [[parse_date(v) if c["type"] == "datetime" else v for v, c in zip(row, cols)]
             for row in table["rows"]], [(a, c["type"]) for a, c in zip(aliases, cols)])
        elements.append({"id": f"dm{number}", "kind": "table",
                         "source": {"kind": "sql", "connectionId": conn, "statement": statement},
                         "columns": [{"id": f"d{number}c{i}", "name": c["name"],
                                      "formula": f"[Custom SQL/{a}]"}
                                     for i, (c, a) in enumerate(zip(cols, aliases), 1)]})
    return {"name": plan["name"] + " - Model", "schemaVersion": 1, "folderId": plan["folderId"],
            "pages": [{"id": "model", "name": "Model", "elements": elements}]}


def map_dm_elements(plan, spec):
    """Map ``dm_tables`` keys -> live element ids by column-name signature (ids are reassigned)."""
    by_sig = {}
    for element in spec["pages"][0]["elements"]:
        by_sig[tuple(c["name"] for c in element["columns"])] = element["id"]
    out = {}
    for item in dm_tables(plan):
        sig = dm_signature(item["table"])
        if sig not in by_sig:
            raise PlanError(f"data model readback has no element with columns {list(sig)}")
        out[item["key"]] = by_sig[sig]
    return out


def local_dm_elements(plan):
    """Element ids as authored (dry run: no readback)."""
    return {item["key"]: f"dm{number}" for number, item in enumerate(dm_tables(plan), 1)}


# ---- seeds
def seed_csvs(plan, out_dir):
    """One paste-ready CSV per editable table (literal columns only); returns the paths."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for table in all_tables(plan):
        if not is_editable(table):
            continue
        path = out_dir / f"seed_{slug(table['name'])}.csv"
        cols = literal_columns(table)
        with open(path, "w", newline="", encoding="utf-8") as stream:
            writer = csv.writer(stream)
            writer.writerow([c["name"] for c in cols])
            for row in table.get("rows", []):
                writer.writerow(["" if v is None else v for v in row])
        written.append(path)
    return written


# ---- workbook
def number_format(plan, kind):
    """Sigma ``format`` object for a plan column ``format``: a full format object (current plans) or one of the
    legacy kind names (``percent`` / ``currency`` / ``date``) older plans carry."""
    if isinstance(kind, dict):
        return dict(kind)
    if kind == "date":
        return {"kind": "datetime", "formatString": "%m/%d/%y"}
    if kind == "percent":
        return {"kind": "number", "formatString": plan.get("pctFormat", ".1%")}
    if kind == "currency":
        return {"kind": "number", "formatString": plan.get("currencyFormat", "$,.0f")}
    return None


FALLBACK_NUMBER = ef.plain_number(2)


def _with_format(plan, column, definition):
    """Attach the column's format. Every numeric column gets an explicit one (a plan without it falls back to a
    plain number) so Sigma never inherits a ``%`` from an operand (``Amount * Probability``)."""
    fmt = number_format(plan, column.get("format"))
    if fmt is None and column.get("type") == "number":
        fmt = dict(FALLBACK_NUMBER)
    if fmt:
        definition["format"] = fmt
    return definition


def _date_format(values):
    parsed = [parse_date(v) for v in values if v is not None]
    monthly = bool(parsed) and all(d.day == 1 for d in parsed)
    return {"kind": "datetime", "formatString": "%b-%y" if monthly else "%m/%d/%y"}


class _Names:
    """Unique element ids."""

    def __init__(self):
        self.used = set()

    def get(self, prefix, name=""):
        base = f"{prefix}_{slug(name)}" if name else prefix
        candidate, bump = base, 2
        while candidate in self.used:
            candidate, bump = f"{base}_{bump}", bump + 1
        self.used.add(candidate)
        return candidate


CARD = {"backgroundColor": "#FFFFFF", "borderColor": {"kind": "theme", "ref": "colors-border"},
        "borderWidth": 1, "borderRadius": "round"}


def referenced_columns(frame, available):
    """Source columns a frame's expressions read, attrs first, in stable order."""
    needed = list(frame.get("attrs", []))
    for column in frame["columns"]:
        for node in walk(column["expr"]):
            if node[0] == "c" and node[1] in available and node[1] not in needed:
                needed.append(node[1])
            if node[0] == "cumsum":
                for name in (node[2], node[3]):
                    if name not in needed:
                        needed.append(name)
    return needed


def _calc_columns(plan, frame, fid, own, emitter):
    out = []
    for number, column in enumerate(frame["columns"], 1):
        if column["name"] in own:
            raise PlanError(f"frame {frame['title']!r}: calculated column {column['name']!r} collides "
                            "with another column; rename it in the plan")
        own.add(column["name"])
        out.append(_with_format(plan, column, {"id": f"{fid}_k{number}", "name": column["name"],
                                               "formula": emitter.column_formula(column["name"], column["expr"])}))
    return out


ORDER_NAME = "Source Order"
MAX_ORDER_LABELS = 40      # a longer grid keeps Sigma's key order (nested If depth); see SOURCE_ORDER_OMITTED


def order_labels(frame):
    """The row labels whose source-sheet order a frame keeps (``None`` when there is nothing to order)."""
    labels = frame.get("labels") or []
    if frame["kind"] not in ("summary", "keyed") or frame.get("hidden") or not labels \
            or len(labels) > MAX_ORDER_LABELS:
        return None
    return labels


def order_formula(key_ref, labels, key_type="text"):
    """Nested ``If`` giving each label its position in the source sheet (rows not listed sort last)."""
    def literal(value):
        if key_type == "datetime" or (isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)
                                      and key_type != "text"):
            return f'Date("{value}")'
        if key_type == "text" or isinstance(value, bool) or not isinstance(value, (int, float)):
            return json.dumps(str(value), ensure_ascii=False)      # UTF-8 as typed; a numeric label on a text key
        return sigma_number(value)
    text = str(len(labels) + 1)
    for position in range(len(labels), 0, -1):
        text = f"If({key_ref} = {literal(labels[position - 1])}, {position}, {text})"
    return text


def _summary_columns(plan, frame, fid, src_types, src_label, src_blank=None, info=None):
    """Group-level calculation columns of a summary frame (``src_label`` None = inside the source)."""
    by = frame["by"]
    src_blank = src_blank or {}
    ref = (lambda n: f"[{n}]") if src_label is None else (lambda n: f"[{src_label}/{n}]")
    src_em = Emitter(col=ref, key=ref(by), coltype=lambda n: src_types.get(n),
                     blankable=lambda n: src_blank.get(n, False), tables=info)
    own_types = {c["name"]: c["type"] for c in frame["columns"]}
    own_types[by] = src_types.get(by, "text")
    em = Emitter(col=lambda n: f"[{n}]", key=f"[{by}]", coltype=lambda n: own_types.get(n),
                 count_ref=ref(by), agg_source=src_em, agg_col=ref, tables=info)
    return _calc_columns(plan, frame, fid, {by}, em)


# ---- chart / KPI feeds: the frame's values computed from the frame's SOURCE columns
# Sigma resolves a KPI or chart formula only against base-level columns of its source: a grouped table's
# grouping-level calculations and group keys cannot be referenced by name (live dry-run, even with
# ``source.groupingId``). So every KPI value and chart series is re-emitted from the same IR the frame uses,
# with the aggregates evaluated at chart / KPI level over the frame's source rows.
def flatten_expr(node, cols):
    """Inline references to other frame columns (``["c", name]``) so only source-level nodes remain."""
    tag = node[0]
    if tag == "c":
        return flatten_expr(cols[node[1]]["expr"], cols) if node[1] in cols else node
    if tag == "op":
        return ["op", node[1], flatten_expr(node[2], cols), flatten_expr(node[3], cols)]
    if tag == "neg":
        return ["neg", flatten_expr(node[1], cols)]
    if tag == "f":
        return ["f", node[1]] + [flatten_expr(a, cols) for a in node[2:]]
    if tag == "lk":
        return ["lk", node[1], node[2], node[3], node[4], flatten_expr(node[5], cols) if node[5] else None]
    return node


def member_ir(by, labels, key_type="text"):
    """IR: the row key is one of the frame's labels."""
    def literal(value):
        if key_type == "text" or isinstance(value, bool) or not isinstance(value, (int, float)):
            return ["d", value] if key_type == "datetime" else ["s", str(value)]
        return ["n", value]
    tests = [["op", "=", ["c", by], literal(v)] for v in labels]
    return tests[0] if len(tests) == 1 else ["f", "OR"] + tests


def restrict_aggs(node, member):
    """Add ``member`` to the filter of every aggregate (rows outside the frame's label set drop out)."""
    tag = node[0]
    if tag == "agg":
        flt = member if node[3] is None else ["f", "AND", node[3], member]
        return ["agg", node[1], node[2], flt]
    if tag == "op":
        return ["op", node[1], restrict_aggs(node[2], member), restrict_aggs(node[3], member)]
    if tag == "neg":
        return ["neg", restrict_aggs(node[1], member)]
    if tag == "f":
        return ["f", node[1]] + [restrict_aggs(a, member) for a in node[2:]]
    return node


def leaf_terms(node, sign=1):
    """``[(sign, leaf)]`` when a flattened expression is a signed sum of aggregates / lookups, else ``None``."""
    tag = node[0]
    if tag in ("agg", "lk"):
        return [(sign, node)]
    if tag == "neg":
        return leaf_terms(node[1], -sign)
    if tag == "op" and node[1] in ("+", "-"):
        left = leaf_terms(node[2], sign)
        right = leaf_terms(node[3], sign if node[1] == "+" else -sign)
        return left + right if left is not None and right is not None else None
    if tag == "f" and node[1] == "SUM":
        parts = [leaf_terms(a, sign) for a in node[2:]]
        return None if any(p is None for p in parts) else [t for p in parts for t in p]
    return None


def _frame_feed(plan, frame, src_id, src_label, src_types, src_blank, info, wrapper, restrict=None):
    """Chart series and KPI terms of a summary frame, emitted from its source's base-level columns.

    ``series[name]``: the frame column as a chart-level formula (the chart groups by the key column);
    ``terms[name]``: the column as signed leaf terms ``{sign, element, label, formula}`` whose sum over ALL
    source rows (restricted to the frame's labels when the data has other keys) is the column's frame total.
    A lookup term always sums only the rows of the lookup table whose key is one of the frame's labels (that
    table can list keys the source does not); ``why[name]`` says why a column has no terms.
    """
    by = frame["by"]
    ref = lambda n: f"[{src_label}/{n}]"                                       # noqa: E731
    src_em = Emitter(col=ref, key=ref(by), coltype=lambda n: src_types.get(n),
                     blankable=lambda n: (src_blank or {}).get(n, False), tables=info, utf8_strings=True)

    def no_group_column(name):
        raise PlanError(f"{name!r} is a frame column, not a source column")
    own_types = {c["name"]: c["type"] for c in frame["columns"]}
    em = Emitter(col=no_group_column, key=ref(by), coltype=lambda n: own_types.get(n), count_ref=ref(by),
                 agg_source=src_em, agg_col=ref, tables=info, wrap_lookup=True, utf8_strings=True)
    plain = Emitter(col=no_group_column, key=ref(by), coltype=lambda n: own_types.get(n), count_ref=ref(by),
                    agg_source=src_em, agg_col=ref, tables=info, utf8_strings=True)
    cols = {c["name"]: c for c in frame["columns"]}
    member = member_ir(by, restrict, src_types.get(by, "text")) if restrict else None
    feed = {"element": src_id, "label": src_label, "key": ref(by), "keyType": src_types.get(by, "text"),
            "series": {}, "terms": {}, "why": {}}
    for column in frame["columns"]:
        name = column["name"]
        feed["series"][name] = feed["terms"][name] = None
        if column["type"] != "number":
            continue
        flat = flatten_expr(column["expr"], cols)
        try:
            feed["series"][name] = em.column_formula(name, flat)
        except (PlanError, KeyError):
            feed["why"][name] = "not expressible from the frame's source columns"
        leaves = leaf_terms(flat)
        if not leaves:
            feed["terms"][name] = None
            feed["why"].setdefault(name, "not a signed sum of aggregates / lookups")
            continue
        terms = []
        for sign, leaf in leaves:
            if leaf[0] == "agg":
                node = restrict_aggs(leaf, member) if member else leaf
                terms.append({"sign": sign, "element": src_id, "label": src_label,
                              "formula": plain.column_formula(name, node)})
                continue
            term, why = _lookup_term(plan, frame, leaf, sign, wrapper, info)
            if term is None:
                terms = None
                feed["why"][name] = why
                break
            terms.append(term)
        feed["terms"][name] = terms
    return feed


def _lookup_term(plan, frame, leaf, sign, wrapper, info):
    """A per-label lookup's frame total: the lookup table's value column over the rows keyed by the frame's
    labels. ``(term, None)`` or ``(None, reason)``; duplicate keys would double count, so they refuse."""
    _, table, key, key_col, value_col, default = leaf
    if key != ["key"] or default is not None or table not in wrapper:
        return None, f"the lookup into {table} is not a plain per-label lookup"
    data = plan_table(plan, table)
    names = [c["name"] for c in literal_columns(data)]
    key_type = next((c["type"] for c in data["columns"] if c["name"] == key_col), "text")
    labels = frame["labels"]
    if key_col in names:
        index = names.index(key_col)
        keys = [row[index] for row in data.get("rows") or []]
        repeated = sorted({str(k) for k in keys if k in labels and keys.count(k) > 1})
        if repeated:
            return None, (f"{table}[{key_col}] has duplicate keys ({', '.join(repeated)}): summing the lookup "
                          "would double count")
    value = f"[{table}/{value_col}]"
    cond = Emitter(col=lambda n, t=table: f"[{t}/{n}]", tables=info, utf8_strings=True).emit_crit(
        member_ir(key_col, labels, key_type))
    return {"sign": sign, "element": wrapper[table], "label": table,
            "formula": f"Coalesce(SumIf({value}, {cond}), 0)"}, None


def _frame_types(plan, frame, types):
    own = {c["name"]: c["type"] for c in frame["columns"]}
    if frame["kind"] == "keyed":
        return dict(types[frame["table"]], **own)
    if frame["kind"] == "expansion":
        return dict(types[frame["base"]], **{frame["axis"]["name"]: frame["axis"]["type"]}, **own)
    return own


def blank_map(plan):
    """``{table: {column: has a blank cell in the plan data}}`` for every literal column."""
    out = {}
    for table in all_tables(plan):
        rows = table.get("rows", [])
        out[table["name"]] = {c["name"]: any(row[i] is None for row in rows)
                              for i, c in enumerate(literal_columns(table))}
    return out


def _table_info(plan, types):
    """Column types and blank-capable flags per table (literal columns that hold a blank)."""
    info = {"types": types, "blank": blank_map(plan)}
    for table in all_tables(plan):      # calculated columns are never blank (see Emitter.column_formula)
        for column in table["columns"]:
            if "expr" in column:
                info["blank"][table["name"]][column["name"]] = False
    return info


def plan_table(plan, name):
    """A plan table (or the Parameters pseudo-table) by name."""
    return next((t for t in all_tables(plan) if t["name"] == name), {"name": name, "columns": [], "rows": []})


def distinct_count(plan, table_name, column):
    """Number of distinct values of a literal column in a plan table (rows shown by a grouping on it)."""
    table = plan_table(plan, table_name)
    names = [c["name"] for c in literal_columns(table)]
    rows = table.get("rows") or []
    if column not in names:
        return len(rows)
    index = names.index(column)
    return len({row[index] for row in rows})


def build_parts(plan, dm_id, dm_elements):
    """Return ``{"inputs", "hidden", "visible"}`` element lists for the workbook pages.

    ``dm_elements`` maps ``dm_tables`` keys to data-model element ids ({} without a DM).
    """
    validate_plan(plan)
    conn = plan["connectionId"]
    ids = _Names()
    tables = all_tables(plan)
    types = {t["name"]: {c["name"]: c["type"] for c in t["columns"]} for t in tables}
    info = _table_info(plan, types)
    blank = info["blank"]
    frame_blank = {}
    inputs, hidden, visible = [], [], []
    visible_meta, frame_elements = [], {}
    wrapper = {}
    expansion_bases = {f["base"] for f in plan["frames"] if f["kind"] == "expansion"}

    def dm_source(key):
        return {"kind": "data-model", "dataModelId": dm_id, "elementId": dm_elements[key]}

    def show(element, frame, role, rows):
        """A Report-page element and what the composer needs to size / style it."""
        visible.append(element)
        visible_meta.append({"id": element["id"], "frame": frame["id"], "role": role, "rows": rows})

    for table in tables:
        name, element_id = table["name"], ids.get("w", table["name"])
        cols = literal_columns(table)
        if is_editable(table):
            input_id, input_name = ids.get("in", name), f"{name} Input"
            inputs.append({
                "id": input_id, "kind": "input-table", "name": input_name, "inputMode": "edit",
                "source": {"kind": "empty", "connectionId": conn},
                "columns": [{"id": "ID"}] + [
                    _with_format(plan, c, {"id": alias(c["name"]), "name": c["name"],
                                           "type": {"boolean": "checkbox"}.get(c["type"], c["type"])})
                    for c in cols],
                "order": [alias(c["name"]) for c in cols]})
            source, prefix = {"kind": "table", "elementId": input_id}, input_name
        else:
            source, prefix = dm_source(f"t:{name}"), "Custom SQL"
        columns = [_with_format(plan, c, {"id": f"{element_id}_c{i}", "name": c["name"],
                                          "formula": f"[{prefix}/{c['name']}]"})
                   for i, c in enumerate(cols, 1)]
        em = Emitter(col=lambda n: f"[{n}]", coltype=lambda n, t=types[name]: t.get(n),
                     blankable=lambda n, f=blank[name]: f.get(n, False), tables=info)
        for number, column in enumerate((c for c in table["columns"] if "expr" in c), 1):
            columns.append(_with_format(plan, column, {"id": f"{element_id}_k{number}", "name": column["name"],
                                                       "formula": em.column_formula(column["name"],
                                                                                    column["expr"])}))
        if name in expansion_bases:
            columns.append({"id": f"{element_id}_jk", "name": "Join Key", "formula": "1"})
        hidden.append({"id": element_id, "kind": "table", "name": name, "source": source, "columns": columns})
        wrapper[name] = element_id

    # keyed frames absorb the one-key summaries over them (running sums partition by the group)
    merged = {}
    for frame in plan["frames"]:
        if frame["kind"] == "summary" and frame["source"]["kind"] == "frame":
            parent = next((f for f in plan["frames"] if f["id"] == frame["source"]["name"]), None)
            if parent and parent["kind"] == "keyed" and parent["rowKey"] == frame["by"]:
                merged.setdefault(parent["id"], []).append(frame)
    absorbed = {s["id"] for group in merged.values() for s in group}
    element_of = {}

    for frame in plan["frames"]:
        kind, fid = frame["kind"], ids.get("f", frame["id"])
        if frame["id"] in absorbed:
            continue
        frame_types = _frame_types(plan, frame, types)
        if kind == "expansion":
            axis, base = frame["axis"], frame["base"]
            aid = ids.get("w", f"{frame['title']} axis")
            hidden.append({"id": aid, "kind": "table", "name": f"{frame['title']} Axis",
                           "source": dm_source(f"axis:{frame['id']}"),
                           "columns": [{"id": f"{aid}_c1", "name": axis["name"],
                                        "formula": f"[Custom SQL/{axis['name']}]"},
                                       {"id": f"{aid}_jk", "name": "Join Key", "formula": "1"}]})
            source = {"kind": "join", "name": "XJ",
                      "primarySource": {"kind": "table", "elementId": wrapper[base]},
                      "joins": [{"name": "Axis", "joinType": "inner",
                                 "left": {"kind": "table", "elementId": wrapper[base]},
                                 "right": {"kind": "table", "elementId": aid},
                                 "columns": [{"left": "[Join Key]", "right": "[Join Key]"}]}]}
            needed = referenced_columns(frame, set(types[base]) | {axis["name"]})
            columns = []
            for number, name in enumerate(needed):
                if name == axis["name"]:
                    columns.append({"id": f"{fid}_a", "name": name, "formula": f"[Axis/{name}]",
                                    "format": axis.get("format") or _date_format(axis["values"])})
                else:
                    columns.append({"id": f"{fid}_b{number}", "name": name, "formula": f"[XJ/{name}]"})
            em = Emitter(col=lambda n: f"[{n}]", period=f"[{axis['name']}]",
                         coltype=lambda n: frame_types.get(n),
                         blankable=lambda n, f=blank[base]: f.get(n, False), tables=info)
            columns += _calc_columns(plan, frame, fid, {c["name"] for c in columns}, em)
            frame_blank[frame["id"]] = dict(blank[base], **em.own_blank)
            hidden.append({"id": fid, "kind": "table", "name": frame["title"], "source": source,
                           "columns": columns})
            element_of[frame["id"]] = fid
            show(_pivot(plan, frame, fid, ids, frame.get("attrs", []), axis["name"], axis["values"]),
                 frame, "pivot", len(plan_table(plan, base).get("rows") or []))
        elif kind == "keyed":
            table = frame["table"]
            needed = referenced_columns(frame, set(types[table]))
            for name in (frame["periodKey"], frame["rowKey"]):
                if name not in needed:
                    needed.insert(0, name)
            columns = [{"id": f"{fid}_b{i}", "name": n, "formula": f"[{table}/{n}]"}
                       for i, n in enumerate(needed)]
            em = Emitter(col=lambda n: f"[{n}]", key=f"[{frame['rowKey']}]",
                         period=f"[{frame['periodKey']}]", coltype=lambda n: frame_types.get(n),
                         blankable=lambda n, f=blank[table]: f.get(n, False), tables=info)
            own = {c["name"] for c in columns}
            columns += _calc_columns(plan, frame, fid, own, em)
            frame_blank[frame["id"]] = dict(blank[table], **em.own_blank)
            by_id = next(c["id"] for c in columns if c["name"] == frame["rowKey"])
            period_id = next(c["id"] for c in columns if c["name"] == frame["periodKey"])
            order_id = None
            labels = order_labels(frame)
            if labels and ORDER_NAME not in own and ps.SURFACES["order_group_sort"]:   # hidden row-level order column
                order_id = f"{fid}_ord"
                columns.append({"id": order_id, "name": ORDER_NAME, "hidden": True,
                                "formula": order_formula(f"[{frame['rowKey']}]", labels,
                                                         frame_types.get(frame["rowKey"], "text"))})
                own.add(ORDER_NAME)
            calc_ids = []
            for summary in merged.get(frame["id"], []):
                extra = _summary_columns(plan, summary, f"{fid}_s", frame_types, None,
                                         frame_blank.get(frame["id"]), info)
                for column in extra:
                    if column["name"] in own:
                        raise PlanError(f"frame {summary['title']!r}: column {column['name']!r} collides "
                                        f"with a column of {frame['title']!r}")
                    own.add(column["name"])
                columns += extra
                calc_ids += [c["id"] for c in extra]
                element_of[summary["id"]] = fid
            grouping = {"id": f"{fid}_by_key", "groupBy": [by_id], "calculations": calc_ids}
            if order_id and calc_ids:      # group-level order: Min of the hidden row-level order column
                group_order = {"id": f"{fid}_ordg", "name": f"{ORDER_NAME} (group)", "hidden": True,
                               "formula": f"Min([{ORDER_NAME}])"}
                columns.append(group_order)
                grouping["calculations"] = [group_order["id"]] + calc_ids
                grouping["sort"] = [{"columnId": group_order["id"], "direction": "ascending"}]
            element = {"id": fid, "kind": "table", "name": frame["title"],
                       "source": {"kind": "table", "elementId": wrapper[table]}, "columns": columns,
                       "sort": [{"columnId": period_id, "direction": "ascending"}],
                       "groupings": [grouping]}
            key_rows = distinct_count(plan, table, frame["rowKey"])
            if calc_ids:
                show(element, frame, "grouped", key_rows)
            else:
                hidden.append(element)
            element_of[frame["id"]] = fid
            frame_elements[frame["id"]] = {"element": fid, "grouping": f"{fid}_by_key" if calc_ids else None,
                                           "title": frame["title"]}
            for summary in merged.get(frame["id"], []):
                restrict = (plan.get("presentation") or {}).get("restrict", {}).get(summary["id"])
                frame_elements[summary["id"]] = {
                    "element": fid, "grouping": f"{fid}_by_key", "title": frame["title"],
                    "feed": _frame_feed(plan, summary, fid, frame["title"], frame_types,
                                        frame_blank.get(frame["id"]), info, wrapper, restrict)}
            show(_pivot(plan, frame, fid, ids, [frame["rowKey"]], frame["periodKey"], frame["periods"],
                        order=order_id), frame, "pivot", key_rows)
        elif kind == "ledger":
            hidden_or_visible = _ledger_elements(plan, frame, fid, ids, wrapper, types, info)
            spine_rows = len(plan_table(plan, frame["spine"]).get("rows") or [])
            for number, element in enumerate(hidden_or_visible):
                show(element, frame, "ledger" if number == 0 else "totals", spine_rows if number == 0 else 1)
            element_of[frame["id"]] = fid
            frame_elements[frame["id"]] = {
                "element": fid, "grouping": None, "title": frame["title"],
                "feed": {"element": fid, "label": frame["title"], "key": None,
                         "series": {c["name"]: f"Sum([{frame['title']}/{c['name']}])"
                                    for c in frame["columns"] if c["type"] == "number"},
                         "terms": {}}}
        else:
            source = frame["source"]
            if source["kind"] == "table":
                src_id, src_label, src_types = wrapper[source["name"]], source["name"], types[source["name"]]
                src_blank = blank[source["name"]]
            else:
                parent = next(f for f in plan["frames"] if f["id"] == source["name"])
                src_id, src_label = element_of[parent["id"]], parent["title"]
                src_types = _frame_types(plan, parent, types)
                src_blank = frame_blank.get(parent["id"])
            key_col = {"id": f"{fid}_key", "name": frame["by"], "formula": f"[{src_label}/{frame['by']}]"}
            calcs = _summary_columns(plan, frame, fid, src_types, src_label, src_blank, info)
            grouping = {"id": f"{fid}_by", "groupBy": [key_col["id"]], "calculations": [c["id"] for c in calcs]}
            tail = []
            labels = order_labels(frame)
            if labels and ps.SURFACES["order_group_sort"] and ORDER_NAME not in {c["name"] for c in calcs} \
                    and ORDER_NAME != frame["by"]:
                # source order: a hidden grouping calculation ``Min(<row-level If over the key>)`` placed first,
                # and a grouping-level sort on it (accepted live; the group-key ``hidden`` flag is dropped)
                tail = [{"id": f"{fid}_ord", "name": ORDER_NAME, "hidden": True,
                         "formula": "Min(" + order_formula(f"[{src_label}/{frame['by']}]", labels,
                                                           src_types.get(frame["by"], "text")) + ")"}]
                grouping["calculations"].insert(0, tail[0]["id"])
                grouping["sort"] = [{"columnId": tail[0]["id"], "direction": "ascending"}]
            # hidden helper frames (ledger source aggregates) stay on the Model page
            element = {
                "id": fid, "kind": "table", "name": frame["title"],
                "source": {"kind": "table", "elementId": src_id}, "columns": [key_col] + calcs + tail,
                "groupings": [grouping]}
            if frame.get("hidden"):
                hidden.append(element)
            else:
                show(element, frame, "grouped", len(frame.get("labels") or []))
                feed_restrict = (plan.get("presentation") or {}).get("restrict", {}).get(frame["id"])
                feed = _frame_feed(plan, frame, src_id, src_label, src_types, src_blank, info, wrapper,
                                   feed_restrict)
            element_of[frame["id"]] = fid
            frame_elements[frame["id"]] = {"element": fid, "grouping": f"{fid}_by", "title": frame["title"]}
            if not frame.get("hidden"):
                frame_elements[frame["id"]]["feed"] = feed
    return {"inputs": inputs, "hidden": hidden, "visible": visible, "visibleMeta": visible_meta,
            "frames": frame_elements}


def _ledger_elements(plan, frame, fid, ids, wrapper, types, info=None):
    """Ledger frame -> the period table (one row per period) and its one-row total table."""
    spine, axis = frame["spine"], frame["axis"]
    own = {c["name"]: c for c in frame["columns"]}
    spine_defs = {c["name"]: c for c in plan_table(plan, spine)["columns"]}
    here = dict(types[spine], **{name: c["type"] for name, c in own.items()})
    spine_blank = (info or {"blank": {}})["blank"].get(spine, {})
    # a number column's blank cells are emitted as 0 below, so only other spine columns stay blank-capable
    em = Emitter(col=lambda n: f"[{n}]", period=f"[{axis}]", coltype=lambda n: here.get(n),
                 blankable=lambda n: spine_blank.get(n, False) and n != axis and here.get(n) != "number",
                 tables=info)
    columns, taken = [], set()
    for number, name in enumerate(frame["order"]):
        if name in taken:
            raise PlanError(f"ledger {frame['title']!r}: column {name!r} appears twice")
        taken.add(name)
        if name in own:
            columns.append(_with_format(plan, own[name], {"id": f"{fid}_k{number}", "name": name,
                                                          "formula": em.column_formula(name, own[name]["expr"])}))
        else:       # spine columns: the period key and the manual line items (a blank cell is 0)
            ref = f"[{spine}/{name}]"
            columns.append(_with_format(plan, spine_defs.get(name, {}), {
                "id": f"{fid}_b{number}", "name": name,
                "formula": ref if name == axis or here.get(name) != "number" else f"Coalesce({ref}, 0)"}))
    axis_id = next(c["id"] for c in columns if c["name"] == axis)
    elements = [{"id": fid, "kind": "table", "name": frame["title"],
                 "source": {"kind": "table", "elementId": wrapper[spine]}, "columns": columns,
                 "sort": [{"columnId": axis_id, "direction": "ascending"}]}]
    totals = frame.get("totals") or []
    if totals:
        tid = ids.get("f", f"{frame['title']} total")
        label = frame.get("totalLabel") or "Total"
        key = {"id": f"{tid}_key", "name": "Period", "formula": SIGMA_STRING(label)}
        calcs = [{"id": f"{tid}_t{number}", "name": name, "formula": f"Sum([{frame['title']}/{name}])"}
                 for number, name in enumerate(totals, 1)]
        for calc in calcs:
            source = own.get(calc["name"]) or spine_defs.get(calc["name"])
            if source:
                _with_format(plan, source, calc)
        elements.append({"id": tid, "kind": "table", "name": f"{frame['title']} - {label}",
                         "source": {"kind": "table", "elementId": fid}, "columns": [key] + calcs,
                         "groupings": [{"id": f"{tid}_by", "groupBy": [key["id"]],
                                        "calculations": [c["id"] for c in calcs]}]})
    return elements


def _pivot(plan, frame, fid, ids, rows, axis_col, axis_values, order=None):
    """Matrix view: rows x period, Sum(measure), with row totals. ``order`` (a hidden source-order column of
    the frame element) keeps the first row dimension in the source sheet's order."""
    pid, title = ids.get("pv", frame["title"]), frame["title"]
    columns, row_by = [], []
    for number, name in enumerate(rows):
        cid = f"{pid}_r{number}"
        columns.append({"id": cid, "name": name, "formula": f"[{title}/{name}]"})
        row_by.append({"columnId": cid, "sort": {"direction": "ascending"}})
    if order and row_by:
        columns.append({"id": f"{pid}_o", "name": ORDER_NAME, "formula": f"[{title}/{ORDER_NAME}]"})
        row_by[0]["sort"] = {"direction": "ascending", "by": f"{pid}_o", "aggregation": "min"}
    axis_id, value_id, measure = f"{pid}_x", f"{pid}_v", frame["measure"]
    axis_format = (frame.get("axis") or {}).get("format") or frame.get("periodFormat")
    columns.append({"id": axis_id, "name": axis_col, "formula": f"[{title}/{axis_col}]",
                    "format": dict(axis_format) if axis_format else _date_format(axis_values)})
    source_column = next((c for c in frame["columns"] if c["name"] == measure), {})
    columns.append(_with_format(plan, source_column, {"id": value_id, "name": measure,
                                                      "formula": f"Sum([{title}/{measure}])"}))
    return {"id": pid, "kind": "pivot-table", "name": f"{title} (by period)",
            "source": {"kind": "table", "elementId": fid}, "columns": columns, "values": [value_id],
            "rowsBy": row_by, "columnsBy": [{"columnId": axis_id, "sort": {"direction": "ascending"}}],
            "axisTotals": [{"axisId": "grand-row-total", "visibility": "hidden"}], "style": dict(CARD)}


def layout_xml(pages):
    """Dense grid layout: [(page id, [(element id, col_start, col_end, row_start, row_end[, children])])].

    A placement with ``children`` (the same tuples, rows relative to the container's own grid) is a
    ``<Container>`` holding them.
    """
    blocks = []
    for page_id, placements in pages:
        lines = [f'<Page type="grid" gridTemplateColumns="repeat(24, 1fr)" gridTemplateRows="auto" '
                 f'id="{page_id}">']
        for element_id, c0, c1, r0, r1, *rest in placements:
            if not rest:
                lines.append(f'  <Element elementId="{element_id}" gridColumn="{c0} / {c1}" '
                             f'gridRow="{r0} / {r1}"/>')
                continue
            lines.append(f'  <Container elementId="{element_id}" type="grid" gridColumn="{c0} / {c1}" '
                         f'gridRow="{r0} / {r1}" gridTemplateColumns="repeat(24, 1fr)" gridTemplateRows="auto">')
            for child, cc0, cc1, cr0, cr1 in rest[0]:
                lines.append(f'    <Element elementId="{child}" gridColumn="{cc0} / {cc1}" '
                             f'gridRow="{cr0} / {cr1}"/>')
            lines.append("  </Container>")
        lines.append("</Page>")
        blocks.append("\n".join(lines))
    return '<?xml version="1.0" encoding="utf-8"?>\n' + "\n".join(blocks)


def build_wb_spec(plan, dm_id, dm_elements):
    """Nested workbook draft (``workbook_wire.wire_workbook`` flattens it): Report, Inputs, hidden Model."""
    try:
        from . import table_report_presentation as presentation
    except ImportError:  # direct import when scripts/lib itself is on sys.path
        import table_report_presentation as presentation
    parts = build_parts(plan, dm_id, dm_elements)
    page = presentation.compose_report(plan, parts)
    pages = [{"id": "report", "name": "Report", "elements": page["elements"]}]
    layout_pages = [("report", page["placements"])]
    if parts["inputs"]:      # no input tables (read-only intent): no empty Inputs page
        pages.append({"id": "inputs", "name": "Inputs", "elements": parts["inputs"]})
        layout_pages.append(("inputs", [(e["id"], 1, 25, 1 + 14 * n, 15 + 14 * n)
                                        for n, e in enumerate(parts["inputs"])]))
    pages.append({"id": "model", "name": "Model (sources)", "visibility": "hidden", "elements": parts["hidden"]})
    layout_pages.append(("model", [(e["id"], 1, 25, 1 + 12 * n, 13 + 12 * n) for n, e in enumerate(parts["hidden"])]))
    spec = {"name": plan["name"], "schemaVersion": 1, "kind": "workbook", "folderId": plan["folderId"],
            "pages": pages, "layout": layout_xml(layout_pages)}
    if page.get("settings"):
        spec["settings"] = page["settings"]
    return spec
