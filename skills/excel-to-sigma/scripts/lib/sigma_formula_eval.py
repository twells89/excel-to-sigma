"""A small evaluator of the Sigma formulas the table-report builder EMITS, with SQL NULL logic.

The table-report drafter verifies a plan with a Python oracle (``table_report_model.Model``) and then
renders the same IR as Sigma formula strings (``Emitter``). Those are two implementations that can
disagree (a blank is 0 in Excel and the oracle but NULL in Sigma; ``Mod`` signs; ``IFERROR`` scope).
This module re-evaluates the *emitted strings* of ``build_parts`` over the plan's own rows, with

* NULL propagation through arithmetic and comparisons, Kleene ``and`` / ``or`` / ``Not``,
* ``If(NULL, a, b)`` taking the else branch, ``Coalesce``, ``IsNull``,
* conditional aggregates that ignore NULL values and give NULL for an empty set,
* a configurable ``Mod`` convention (dividend-signed like SQL, or divisor-signed like Excel),

so a test can compare every verified source cell with what the emitted Sigma would show.
Supported elements: table wrappers (calculated columns), expansion joins and grouped summaries over a
table or frame. Keyed/ledger frames (running sums) are reported as skipped, never silently passed.
"""
from __future__ import annotations

import math
import re
from datetime import date, timedelta

try:
    from . import table_report_model as trm
except ImportError:  # direct import when scripts/lib itself is on sys.path
    import table_report_model as trm


class Unsupported(Exception):
    """An emitted construct this evaluator does not model (the caller counts it as skipped)."""


# ---------------------------------------------------------------------------- parsing
TOKEN = re.compile(r"""\s*(?:
    (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<str>"(?:[^"\\]|\\.)*")
  | (?P<ref>\[[^\]]*\])
  | (?P<id>[A-Za-z_][A-Za-z_0-9]*)
  | (?P<op>!=|<=|>=|[-+*/=<>(),])
)""", re.X)


def tokenize(text):
    pos, out = 0, []
    while pos < len(text):
        match = TOKEN.match(text, pos)
        if not match or match.end() == pos:
            if text[pos:].strip() == "":
                break
            raise ValueError(f"cannot tokenize {text[pos:pos + 30]!r} in {text!r}")
        pos = match.end()
        kind = match.lastgroup
        out.append((kind, match.group(kind)))
    return out


class Parser:
    def __init__(self, text):
        self.text, self.toks, self.i = text, tokenize(text), 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, value=None):
        kind, tok = self.peek()
        if value is not None and tok != value:
            raise ValueError(f"expected {value!r} got {tok!r} in {self.text!r}")
        self.i += 1
        return kind, tok

    def parse(self):
        node = self.p_or()
        if self.i != len(self.toks):
            raise ValueError(f"trailing tokens in {self.text!r}: {self.toks[self.i:]}")
        return node

    def p_or(self):
        node = self.p_and()
        while self.peek() == ("id", "or"):
            self.take()
            node = ("or", node, self.p_and())
        return node

    def p_and(self):
        node = self.p_cmp()
        while self.peek() == ("id", "and"):
            self.take()
            node = ("and", node, self.p_cmp())
        return node

    def p_cmp(self):
        node = self.p_add()
        kind, tok = self.peek()
        if kind == "op" and tok in ("=", "!=", "<", ">", "<=", ">="):
            self.take()
            node = ("cmp", tok, node, self.p_add())
        return node

    def p_add(self):
        node = self.p_mul()
        while self.peek()[0] == "op" and self.peek()[1] in "+-":
            _, tok = self.take()
            node = ("arith", tok, node, self.p_mul())
        return node

    def p_mul(self):
        node = self.p_unary()
        while self.peek()[0] == "op" and self.peek()[1] in "*/":
            _, tok = self.take()
            node = ("arith", tok, node, self.p_unary())
        return node

    def p_unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.p_unary())
        return self.p_primary()

    def p_primary(self):
        kind, tok = self.take()
        if kind == "num":
            return ("lit", float(tok) if re.search(r"[.eE]", tok) else int(tok))
        if kind == "str":
            import json
            return ("lit", json.loads(tok))
        if kind == "ref":
            return ("ref", tok[1:-1])
        if kind == "op" and tok == "(":
            node = self.p_or()
            self.take(")")
            return node
        if kind == "id":
            if tok in ("True", "False") and self.peek() != ("op", "("):
                return ("lit", tok == "True")
            self.take("(")
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.p_or())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.p_or())
            self.take(")")
            return ("call", tok, args)
        raise ValueError(f"unexpected token {tok!r} in {self.text!r}")


_PARSED = {}


def parse(text):
    if text not in _PARSED:
        _PARSED[text] = Parser(text).parse()
    return _PARSED[text]


# ---------------------------------------------------------------------------- evaluation
def _add_months(day, months):
    index = day.year * 12 + day.month - 1 + months
    year, month = divmod(index, 12)
    last = (date(year + (month == 11), (month + 1) % 12 + 1, 1) - timedelta(days=1)).day
    return date(year, month + 1, min(day.day, last))


AGGREGATES = {"Sum", "Avg", "Min", "Max", "Count", "SumIf", "AvgIf", "MinIf", "MaxIf", "CountIf"}


class Ctx:
    """Resolution context for one evaluation: ``ref(name)`` plus (for aggregates) the group's rows."""

    def __init__(self, ref, rows=None, row_ctx=None):
        self.ref, self.rows, self.row_ctx = ref, rows, row_ctx


class Evaluator:
    def __init__(self, lookup, mod="floor"):
        self.lookup, self.mod = lookup, mod      # lookup(table, valueCol, key, keyCol)

    def ev(self, node, ctx):
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "ref":
            return ctx.ref(node[1])
        if kind == "neg":
            value = self.ev(node[1], ctx)
            return None if value is None else -value
        if kind == "arith":
            a, b = self.ev(node[2], ctx), self.ev(node[3], ctx)
            if a is None or b is None:
                return None
            sym = node[1]
            if sym == "/":
                return None if b == 0 else a / b
            return a + b if sym == "+" else a - b if sym == "-" else a * b
        if kind == "cmp":
            a, b = self.ev(node[2], ctx), self.ev(node[3], ctx)
            if a is None or b is None:
                return None
            if isinstance(a, str) != isinstance(b, str):
                raise ValueError(f"comparing {a!r} with {b!r}: Sigma would not compile this")
            return {"=": a == b, "!=": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[node[1]]
        if kind in ("and", "or"):
            a = self.ev(node[1], ctx)
            b = self.ev(node[2], ctx)
            if kind == "and":
                if a is False or b is False:
                    return False
                return None if a is None or b is None else True
            if a is True or b is True:
                return True
            return None if a is None or b is None else False
        if kind == "call":
            return self.call(node[1], node[2], ctx)
        raise ValueError(f"unknown node {kind}")

    def call(self, name, args, ctx):
        if name in AGGREGATES:
            return self.aggregate(name, args, ctx)
        if name == "If":
            cond = self.ev(args[0], ctx)
            return self.ev(args[1] if cond is True else args[2], ctx)       # NULL -> else branch
        if name == "Coalesce":
            for arg in args:
                value = self.ev(arg, ctx)
                if value is not None:
                    return value
            return None
        if name == "IsNull":
            return self.ev(args[0], ctx) is None
        if name == "Not":
            value = self.ev(args[0], ctx)
            return None if value is None else not value
        if name == "Lookup":
            key = self.ev(args[1], ctx)
            value_ref, key_ref = args[0], args[2]
            if value_ref[0] != "ref" or key_ref[0] != "ref":
                raise ValueError("Lookup needs column references")
            table, value_col = value_ref[1].split("/", 1)
            table2, key_col = key_ref[1].split("/", 1)
            assert table == table2
            return None if key is None else self.lookup(table, value_col, key, key_col)
        values = [self.ev(a, ctx) for a in args]
        if name in ("Least", "Greatest"):
            if any(v is None for v in values):
                return None
            return min(values) if name == "Least" else max(values)
        if name == "Abs":
            return None if values[0] is None else abs(values[0])
        if name == "Round":
            if values[0] is None or values[1] is None:
                return None
            factor = 10 ** int(values[1])
            return math.copysign(math.floor(abs(values[0]) * factor + 0.5) / factor, values[0])
        if name == "Power":
            return None if None in values else values[0] ** values[1]
        if name == "Ceiling":
            return None if values[0] is None else math.ceil(values[0])
        if name == "Mod":
            a, b = values
            if a is None or b is None or b == 0:
                return None
            return math.fmod(a, b) if self.mod == "trunc" else a % b
        if name == "Concat":
            return None if None in values else "".join(values)
        if name == "Date":
            return date.fromisoformat(values[0])
        if name == "DateAdd":
            unit, amount, day = values
            if amount is None or day is None:
                return None
            return day + timedelta(days=int(amount)) if unit == "day" else _add_months(day, int(amount))
        if name == "DateTrunc":
            return None if values[1] is None else values[1].replace(day=1)
        if name == "DatePart":
            return None if values[1] is None else {"month": values[1].month, "year": values[1].year}[values[0]]
        raise Unsupported(f"function {name}")

    def aggregate(self, name, args, ctx):
        if ctx.rows is None:
            raise ValueError(f"{name} outside a grouped element")
        base = name.removesuffix("If")
        conditional = name.endswith("If")
        cond = args[-1] if conditional else None
        value = args[0] if base != "Count" or not conditional else None
        picked = []
        for row in ctx.rows:
            row_ctx = ctx.row_ctx(row)
            if cond is not None and self.ev(cond, row_ctx) is not True:     # NULL condition -> excluded
                continue
            picked.append(row_ctx)
        if base == "Count" and value is None:
            return len(picked)
        values = [self.ev(value, rc) for rc in picked]
        values = [v for v in values if v is not None]
        if base == "Count":
            return len(values)
        if not values:
            return None
        return {"Sum": sum, "Avg": lambda v: sum(v) / len(v), "Min": min, "Max": max}[base](values)


# ---------------------------------------------------------------------------- elements
class EmittedModel:
    """Evaluate ``build_parts`` output over the plan's rows."""

    def __init__(self, plan, mod="floor"):
        self.plan, self.mod = plan, mod
        self.parts = trm.build_parts(plan, "dm", trm.local_dm_elements(plan))
        self.elements = {}
        for group in ("inputs", "hidden", "visible"):
            for element in self.parts[group]:
                self.elements[element["id"]] = element
        self.by_name = {e["name"]: e for g in ("hidden", "visible") for e in self.parts[g]}
        self.input_ids = {e["id"] for e in self.parts["inputs"]}
        self.plan_tables = {t["name"]: t for t in trm.all_tables(plan)}
        self.frames = {f["id"]: f for f in plan["frames"]}
        self._rows = {}
        self.eval = Evaluator(self._lookup, mod)

    # ---- sources
    def _leaf_rows(self, name):
        if name in self.plan_tables:
            table = self.plan_tables[name]
            cols = trm.literal_columns(table)
            return [{c["name"]: trm.parse_date(v) if c["type"] == "datetime" else v
                     for c, v in zip(cols, row)} for row in table.get("rows", [])]
        for frame in self.plan["frames"]:
            if frame["kind"] == "expansion" and name == f"{frame['title']} Axis":
                axis = frame["axis"]
                return [{axis["name"]: trm.parse_date(v), "Join Key": 1} for v in axis["values"]]
        raise Unsupported(f"no data for {name!r}")

    def source_rows(self, element):
        """List of ``{label: row dict}`` the element's column formulas are evaluated over."""
        source = element["source"]
        kind = source["kind"]
        if kind == "data-model":
            return [{"Custom SQL": row} for row in self._leaf_rows(element["name"])]
        if kind == "table" and source["elementId"] in self.input_ids:
            parent = self.elements[source["elementId"]]
            return [{parent["name"]: row} for row in self._leaf_rows(element["name"])]
        if kind == "table":
            parent = self.elements[source["elementId"]]
            return [{parent["name"]: row} for row in self.rows_of(parent)]
        if kind == "join":
            base = self.elements[source["primarySource"]["elementId"]]
            right = self.elements[source["joins"][0]["right"]["elementId"]]
            return [{source["name"]: dict(b, __base=i), source["joins"][0]["name"]: r}
                    for i, b in enumerate(self.rows_of(base)) for r in self.rows_of(right)]
        raise Unsupported(f"source kind {kind}")

    def _lookup(self, table, value_col, key, key_col):
        element = self.by_name.get(table)
        if element is None:
            raise Unsupported(f"lookup into {table!r}")
        for row in self.rows_of(element):
            if row.get(key_col) is not None and row[key_col] == key:
                return row[value_col]
        return None

    # ---- element rows
    def rows_of(self, element):
        eid = element["id"]
        if eid in self._rows:
            return self._rows[eid]
        if element["kind"] != "table":
            raise Unsupported(f"element kind {element['kind']}")
        if any(re.search(r"Cumulative", c["formula"]) for c in element["columns"] if "formula" in c):
            raise Unsupported("running sums")
        sources = self.source_rows(element)
        names = {c["id"]: c["name"] for c in element["columns"]}
        groupings = element.get("groupings")
        out = []
        if not groupings:
            for src in sources:
                values = {}
                ctx = Ctx(self._ref(src, values))
                for column in element["columns"]:
                    values[column["name"]] = self.eval.ev(parse(column["formula"]), ctx)
                for label, row in src.items():
                    if "__base" in row:
                        values["__base"] = row["__base"]
                out.append(values)
        else:
            if len(groupings) != 1:
                raise Unsupported("several groupings")
            grouping = groupings[0]
            key_cols = [c for c in element["columns"] if c["id"] in grouping["groupBy"]]
            calc_ids = set(grouping["calculations"])
            groups = {}
            for src in sources:
                values = {}
                ctx = Ctx(self._ref(src, values))
                key = tuple(self.eval.ev(parse(c["formula"]), ctx) for c in key_cols)
                groups.setdefault(key, []).append(src)
            for key, members in groups.items():
                values = dict(zip((c["name"] for c in key_cols), key))
                for column in element["columns"]:
                    if column["id"] not in calc_ids:
                        continue
                    ctx = Ctx(lambda ref, v=values: self._group_ref(v, ref), members,
                              lambda src, v=values: Ctx(self._ref(src, {})))
                    values[column["name"]] = self.eval.ev(parse(column["formula"]), ctx)
                out.append(values)
        self._rows[eid] = out
        return out

    @staticmethod
    def _ref(src, values):
        def ref(name):
            if "/" in name:
                label, column = name.split("/", 1)
                if label not in src:
                    raise Unsupported(f"unknown source {label!r}")
                return src[label][column]
            if name in values:
                return values[name]
            raise Unsupported(f"column {name!r} is not computed yet")
        return ref

    @staticmethod
    def _group_ref(values, name):
        if name in values:
            return values[name]
        raise Unsupported(f"group-level reference {name!r}")

    # ---- checks
    def emitted_cell(self, group, sel):
        """The emitted value for one check cell; raises ``Unsupported`` for kinds not modelled."""
        kind, frame_key, column = group["kind"], group["frame"], group["column"]
        if kind == "table":
            return self.rows_of(self.by_name[frame_key])[sel[0]][column]
        frame = self.frames[frame_key]
        if kind == "summary":
            if frame["title"] not in self.by_name:        # absorbed into a keyed frame (running sums)
                raise Unsupported(f"summary {frame['title']!r} lives inside a keyed frame")
            for row in self.rows_of(self.by_name[frame["title"]]):
                if row[frame["by"]] == sel[0]:
                    return row[column]
            raise KeyError(f"no summary row {sel[0]!r}")
        if kind in ("expansion", "rowtotal"):
            rows = [r for r in self.rows_of(self.by_name[frame["title"]]) if r["__base"] == sel[0]]
            if kind == "rowtotal":
                return sum(r[column] or 0 for r in rows)
            wanted = trm.parse_date(sel[1])
            for row in rows:
                if row[frame["axis"]["name"]] == wanted:
                    return row[column]
            raise KeyError(f"no expansion row {sel!r}")
        raise Unsupported(f"check kind {kind}")


def compare_to_workbook(plan, mod="floor", tolerance=None):
    """Compare every verified source cell with the value the EMITTED formulas give.

    Returns ``(compared, skipped, problems)``; ``problems`` lists cells where the emitted formula's
    value differs from Excel's cached value (NULL only matches a blank cached value).
    """
    model = EmittedModel(plan, mod)
    compared, skipped, problems = 0, 0, []
    for group in plan["checks"]:
        for address, sel, want in group["cells"]:
            try:
                got = model.emitted_cell(group, sel)
            except Unsupported:
                skipped += 1
                continue
            compared += 1
            if not trm.values_match(got, want, tolerance):
                problems.append(f"{address} ({group['frame']}.{group['column']}): emitted {trm.iso(got)!r} "
                                f"vs workbook {want!r}")
    return compared, skipped, problems
