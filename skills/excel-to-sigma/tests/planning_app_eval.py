"""Run the planning app's emitted spec over local rows: a small Sigma element evaluator for tests.

The planning app (``build-forecast-model.py --workflows``) moves the actual / forecast split with
``Parameters.Forecast Start Month``. Asserting formula strings cannot show that a roll forward reads the
landed month, or that a closed month whose actuals have not landed keeps its forecast. This module
evaluates the data model and workbook elements the builder EMITS (column formulas, data-model /
table / join / union / custom-SQL / input-table sources, list filters) over rows a test supplies:

* ``warehouse``: ``{TABLE: [row dict keyed by landed column name]}`` for warehouse-table sources,
* ``inputs``: ``{elementId: [row dict keyed by column id]}`` for input tables,
* ``controls``: ``{controlId: value}``.

Semantics follow Sigma where the builder relies on them: NULL propagates through arithmetic and
comparisons, ``If(NULL, ...)`` takes the else branch, an aggregate over another element
(``Max([Parameters/Forecast Start Month])``) is a scalar over that element's rows, multi-key
``Lookup(Sum([E/V]), k1, [E/K1], k2, [E/K2])`` aggregates the matching rows (NULL when none) and a
left-outer join keeps unmatched primary rows. An unsupported construct raises, never passes silently.
"""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta

_TOKEN = re.compile(r"""\s*(?:
      (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
    | (?P<str>"(?:[^"\\]|\\.)*")
    | (?P<ref>\[[^\]]+\])
    | (?P<word>[A-Za-z_][A-Za-z0-9_]*)
    | (?P<op>!=|<=|>=|[-+*/^(),=<>])
)""", re.X)


class EvalError(ValueError):
    pass


def _tokens(text):
    out, pos = [], 0
    while pos < len(text):
        m = _TOKEN.match(text, pos)
        if not m or m.end() == pos:
            if not text[pos:].strip():
                break
            raise EvalError(f"cannot read {text[pos:pos + 20]!r} in {text!r}")
        pos = m.end()
        out.append((m.lastgroup, m.group(m.lastgroup)))
    return out


class _Parser:
    def __init__(self, text):
        self.text, self.toks, self.i = text, _tokens(text), 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, value=None):
        tok = self.peek()
        if tok[0] is None or (value is not None and tok[1] != value):
            raise EvalError(f"expected {value!r} at {tok[1]!r} in {self.text!r}")
        self.i += 1
        return tok

    def parse(self):
        node = self.or_()
        if self.i != len(self.toks):
            raise EvalError(f"trailing {self.peek()[1]!r} in {self.text!r}")
        return node

    def or_(self):
        node = self.and_()
        while self.peek() == ("word", "or"):
            self.take()
            node = ("or", node, self.and_())
        return node

    def and_(self):
        node = self.cmp()
        while self.peek() == ("word", "and"):
            self.take()
            node = ("and", node, self.cmp())
        return node

    def cmp(self):
        node = self.add()
        if self.peek()[0] == "op" and self.peek()[1] in ("=", "!=", "<", ">", "<=", ">="):
            op = self.take()[1]
            node = ("cmp", op, node, self.add())
        return node

    def add(self):
        node = self.mul()
        while self.peek()[0] == "op" and self.peek()[1] in "+-":
            op = self.take()[1]
            node = ("arith", op, node, self.mul())
        return node

    def mul(self):
        node = self.unary()
        while self.peek()[0] == "op" and self.peek()[1] in "*/":
            op = self.take()[1]
            node = ("arith", op, node, self.unary())
        return node

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.unary())
        return self.power()

    def power(self):
        node = self.primary()
        if self.peek() == ("op", "^"):
            self.take()
            node = ("arith", "^", node, self.unary())
        return node

    def primary(self):
        kind, tok = self.take()
        if kind == "num":
            return ("lit", float(tok) if re.search(r"[.eE]", tok) else int(tok))
        if kind == "str":
            return ("lit", tok[1:-1].replace('\\"', '"'))
        if kind == "ref":
            return ("ref", tok[1:-1])
        if kind == "op" and tok == "(":
            node = self.or_()
            self.take(")")
            return node
        if kind == "word":
            if tok in ("True", "False", "true", "false") and self.peek() != ("op", "("):
                return ("lit", tok.lower() == "true")
            if tok in ("Null", "NULL") and self.peek() != ("op", "("):
                return ("lit", None)
            self.take("(")
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.or_())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.or_())
            self.take(")")
            return ("call", tok, args)
        raise EvalError(f"unexpected {tok!r} in {self.text!r}")


_CACHE = {}


def parse(text):
    if text not in _CACHE:
        _CACHE[text] = _Parser(text).parse()
    return _CACHE[text]


def _add_months(day, months):
    index = day.year * 12 + day.month - 1 + months
    year, month = divmod(index, 12)
    last = (date(year + (month == 11), (month + 1) % 12 + 1, 1) - timedelta(days=1)).day
    return date(year, month + 1, min(day.day, last))


def _as_date(value):
    """A datetime compares as its date (the builder's date columns are month-grain)."""
    return value.date() if isinstance(value, datetime) else value


def _to_date(value):
    """A date-function argument: a date, a datetime or an ISO date string (an input-table seed)."""
    value = _as_date(value)
    return date.fromisoformat(value[:10]) if isinstance(value, str) else value


AGGREGATES = {"Sum", "Max", "Min", "Count", "SumIf", "CountIf", "MaxIf", "MinIf"}

# ---------------------------------------------------------------------------- custom SQL VALUES
_SQL_TOKEN = re.compile(r"\s*(?:(?P<str>'(?:[^']|'')*')(?:::(?P<cast>\w+))?|(?P<null>NULL)(?:::\w+)?"
                        r"|(?P<num>-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|(?P<sep>[(),]))", re.I)


def values_rows(statement):
    """Rows of a ``select * from (values (...), ...) as t(COL, ...)`` statement (the builder's shape)."""
    m = re.search(r"\(values\s*(.*)\)\s*as\s+t\(([^)]*)\)\s*$", statement, re.S | re.I)
    if not m:
        raise EvalError(f"unsupported custom SQL: {statement[:60]!r}")
    body, cols = m.group(1), [c.strip() for c in m.group(2).split(",")]
    rows, row, pos, depth = [], None, 0, 0
    while pos < len(body):
        t = _SQL_TOKEN.match(body, pos)
        if not t or t.end() == pos:
            if not body[pos:].strip():
                break
            raise EvalError(f"cannot read SQL {body[pos:pos + 20]!r}")
        pos = t.end()
        if t.group("sep") == "(":
            depth += 1
            row = []
        elif t.group("sep") == ")":
            depth -= 1
            rows.append(dict(zip(cols, row)))
        elif t.group("sep") == ",":
            continue
        elif t.group("null"):
            row.append(None)
        elif t.group("str") is not None:
            text = t.group("str")[1:-1].replace("''", "'")
            row.append(date.fromisoformat(text) if (t.group("cast") or "").lower() == "date" else text)
        else:
            num = t.group("num")
            row.append(float(num) if re.search(r"[.eE]", num) else int(num))
    return rows


# ---------------------------------------------------------------------------- workbook
class Model:
    """Evaluate a nested (unwired) workbook spec + data-model spec over local rows."""

    def __init__(self, wb_spec, dm_spec, *, warehouse=None, inputs=None, controls=None):
        self.elements = {e["id"]: e for pg in wb_spec["pages"] for e in pg["elements"]}
        self.by_name = {}
        for e in self.elements.values():
            if e.get("name"):
                self.by_name.setdefault(e["name"], e["id"])
        self.dm = {e["id"]: e for pg in dm_spec["pages"] for e in pg["elements"]}
        self.warehouse, self.inputs, self.controls = warehouse or {}, inputs or {}, controls or {}
        self._rows = {}

    # ------------------------------------------------------------------ sources
    def _dm_rows(self, element_id):
        el = self.dm[element_id]
        src = el["source"]
        if src["kind"] == "warehouse-table":
            prefix, raw = src["path"][-1], self.warehouse.get(src["path"][-1], [])
        elif src["kind"] == "sql":
            prefix, raw = "Custom SQL", values_rows(src["statement"])
        else:
            raise EvalError(f"data-model source {src['kind']}")
        return self._columns(el["columns"], [{prefix: r} for r in raw]), el.get("name") or "Custom SQL"

    def _source(self, src):
        """``[(context rows)]``: each row is ``{prefix: {column: value}}``."""
        kind = src["kind"]
        if kind == "data-model":
            rows, prefix = self._dm_rows(src["elementId"])
            return [{prefix: r} for r in rows]
        if kind == "table":
            el = self.elements[src["elementId"]]
            return [{el["name"]: r} for r in self.rows(src["elementId"])]
        if kind == "sql":
            return [{"Custom SQL": r} for r in values_rows(src["statement"])]
        if kind == "join":
            primary = self.elements[src["primarySource"]["elementId"]]
            out = [{src["name"]: r} for r in self.rows(primary["id"])]
            for leg in src["joins"]:
                if leg["joinType"] != "left-outer":
                    raise EvalError(f"join type {leg['joinType']}")
                right = self.rows(leg["right"]["elementId"])
                blank = {c.get("name") or c["id"]: None for c in self.elements[leg["right"]["elementId"]]["columns"]}
                keys = [(c["left"][1:-1], c["right"][1:-1]) for c in leg["columns"]]
                joined = []
                for ctx in out:
                    left = ctx[src["name"]]
                    hits = [r for r in right if all(left.get(lk) is not None and left.get(lk) == r.get(rk)
                                                    for lk, rk in keys)]
                    for hit in hits or [None]:
                        joined.append({**ctx, leg["name"]: hit or blank})
                out = joined
            return out
        if kind == "union":
            name = f"Union of {len(src['sources'])} Sources"
            out = []
            for i, part in enumerate(src["sources"]):
                rows = self.rows(part["elementId"])
                for r in rows:
                    out.append({name: {m["outputColumnName"]: r.get(m["sourceColumns"][i][1:-1])
                                       for m in src["matches"]}})
            return out
        raise EvalError(f"source kind {kind}")

    def rows(self, element_id):
        if element_id in self._rows:
            return self._rows[element_id]
        el = self.elements[element_id]
        if el["kind"] == "input-table":
            names = {c["id"]: c.get("name") or c["id"] for c in el["columns"]}
            rows = [{names.get(k, k): v for k, v in r.items()} for r in self.inputs.get(element_id, [])]
            for r in rows:
                for col in el["columns"]:
                    r.setdefault(names[col["id"]], None)
        else:
            self._rows[element_id] = []        # a cycle reads empty and fails the caller's assertions
            rows = self._columns(el["columns"], self._source(el["source"]))
            for f in el.get("filters") or []:
                if f["kind"] != "list" or f["mode"] != "include":
                    raise EvalError(f"filter {f['kind']}/{f['mode']}")
                name = next(c["name"] for c in el["columns"] if c["id"] == f["columnId"])
                rows = [r for r in rows if r[name] in f["values"]]
        self._rows[element_id] = rows
        return rows

    def _columns(self, columns, contexts):
        out = []
        for ctx in contexts:
            row, busy = {}, set()

            def own(name, ctx=ctx, row=row, busy=busy):
                if name in row:
                    return row[name]
                col = next((c for c in columns if c["name"] == name), None)
                if col is None:
                    raise KeyError(name)
                if name in busy:
                    raise EvalError(f"column {name} references itself")
                busy.add(name)
                row[name] = self.eval(parse(col["formula"]), ctx, own)
                return row[name]
            for col in columns:
                own(col["name"])
            out.append(row)
        return out

    # ------------------------------------------------------------------ formulas
    def scalar(self, formula):
        """A formula outside any row (a text element's ``{{...}}`` expression)."""
        return self.eval(parse(formula), {}, None)

    def _ref(self, name, ctx, own):
        if "/" in name:
            prefix, col = name.split("/", 1)
            if prefix in ctx:
                if col not in ctx[prefix]:
                    raise EvalError(f"[{name}] is not a column of {prefix}")
                return ctx[prefix][col]
            raise EvalError(f"[{name}] outside an aggregate")
        if own is not None:
            try:
                return own(name)
            except KeyError:
                pass
        if name in self.controls:
            return self.controls[name]
        raise EvalError(f"unknown reference [{name}]")

    def _target(self, nodes, ctx):
        """The other element an aggregate ranges over (named by its references)."""
        found = set()

        def walk(node):
            if node[0] == "ref" and "/" in node[1]:
                prefix = node[1].split("/", 1)[0]
                if prefix not in ctx:
                    found.add(prefix)
            elif node[0] in ("arith", "cmp"):
                walk(node[2]), walk(node[3])
            elif node[0] in ("and", "or"):
                walk(node[1]), walk(node[2])
            elif node[0] == "neg":
                walk(node[1])
            elif node[0] == "call":
                for a in node[2]:
                    walk(a)
        for n in nodes:
            walk(n)
        if len(found) != 1:
            raise EvalError(f"aggregate over {sorted(found) or 'its own rows'}: not modelled")
        prefix = found.pop()
        if prefix not in self.by_name:
            raise EvalError(f"no element named {prefix}")
        return prefix, self.rows(self.by_name[prefix])

    def _aggregate(self, name, args, ctx, own, rows=None, prefix=None):
        if rows is None:
            prefix, rows = self._target(args, ctx)
        cond = args[-1] if name.endswith("If") else None
        base = name[:-2] if cond is not None else name
        values = []
        for r in rows:
            inner = {**ctx, prefix: r}
            if cond is not None and self.eval(cond, inner, own) is not True:
                continue
            if base == "Count" and cond is not None and len(args) == 1:
                values.append(1)
                continue
            v = self.eval(args[0], inner, own)
            if v is not None:
                values.append(v)
        if base == "Count":
            return len(values)
        if not values:
            return None
        return {"Sum": sum, "Max": max, "Min": min}[base](values)

    def eval(self, node, ctx, own):
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "ref":
            return self._ref(node[1], ctx, own)
        if kind == "neg":
            v = self.eval(node[1], ctx, own)
            return None if v is None else -v
        if kind == "arith":
            a, b = self.eval(node[2], ctx, own), self.eval(node[3], ctx, own)
            if a is None or b is None:
                return None
            op = node[1]
            if op == "/":
                return None if b == 0 else a / b
            if op == "^":
                return a ** b
            return a + b if op == "+" else a - b if op == "-" else a * b
        if kind == "cmp":
            a, b = self.eval(node[2], ctx, own), self.eval(node[3], ctx, own)
            if a is None or b is None:
                return None
            a, b = _as_date(a), _as_date(b)
            if isinstance(a, str) != isinstance(b, str):
                raise EvalError(f"comparing {a!r} with {b!r}")
            return {"=": a == b, "!=": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[node[1]]
        if kind in ("and", "or"):
            a, b = self.eval(node[1], ctx, own), self.eval(node[2], ctx, own)
            if kind == "and":
                return False if (a is False or b is False) else None if (a is None or b is None) else True
            return True if (a is True or b is True) else None if (a is None or b is None) else False
        if kind == "call":
            return self.call(node[1], node[2], ctx, own)
        raise EvalError(f"node {kind}")

    def call(self, name, args, ctx, own):
        ev = lambda n: self.eval(n, ctx, own)     # noqa: E731
        if name in AGGREGATES:
            return self._aggregate(name, args, ctx, own)
        if name == "Lookup":
            if len(args) < 3 or len(args) % 2 != 1:
                raise EvalError("Lookup(value, key, keyColumn, ...)")
            pairs = [(ev(args[i]), args[i + 1]) for i in range(1, len(args), 2)]
            prefix, rows = self._target([k for _v, k in pairs], ctx)
            if any(v is None for v, _k in pairs):
                return None
            hits = [r for r in rows if all(_as_date(r[k[1].split("/", 1)[1]]) == _as_date(v) for v, k in pairs)]
            value = args[0]
            if value[0] == "call" and value[1] in AGGREGATES:
                return self._aggregate(value[1], value[2], ctx, own, rows=hits, prefix=prefix) if hits else None
            return self.eval(value, {**ctx, prefix: hits[0]}, own) if hits else None
        if name == "If":
            for i in range(0, len(args) - 1, 2):
                if ev(args[i]) is True:
                    return ev(args[i + 1])
            return ev(args[-1]) if len(args) % 2 else None
        if name == "Switch":
            key = ev(args[0])
            for i in range(1, len(args) - 1, 2):
                if key is not None and key == ev(args[i]):
                    return ev(args[i + 1])
            return ev(args[-1]) if len(args) % 2 == 0 else None
        if name == "Coalesce":
            for a in args:
                v = ev(a)
                if v is not None:
                    return v
            return None
        if name == "IsNull":
            return ev(args[0]) is None
        values = [ev(a) for a in args]
        if name == "MakeDate":
            return None if None in values else date(int(values[0]), int(values[1]), int(values[2]))
        if name == "DateAdd":
            unit, amount, day = values
            if amount is None or day is None:
                return None
            day = _to_date(day)
            return day + timedelta(days=int(amount)) if unit == "day" else _add_months(day, int(amount))
        if name == "DateTrunc":
            return None if values[1] is None else _to_date(values[1]).replace(day=1)
        if name == "DatePart":
            if values[1] is None:
                return None
            day = _to_date(values[1])
            return {"year": day.year, "month": day.month}[values[0]]
        if name == "DateDiff":
            if values[0] != "month":
                raise EvalError(f"DateDiff unit {values[0]}")
            a, b = _to_date(values[1]), _to_date(values[2])
            return None if a is None or b is None else (b.year * 12 + b.month) - (a.year * 12 + a.month)
        if name == "DateFormat":
            return None if values[0] is None else _to_date(values[0]).strftime(values[1])
        if name == "Concat":
            return None if None in values else "".join(str(v) for v in values)
        if name == "Text":
            v = values[0]
            return None if v is None else str(int(v)) if float(v).is_integer() else str(v)
        if name == "Abs":
            return None if values[0] is None else abs(values[0])
        if name == "Round":
            return None if None in values else round(values[0], int(values[1]))
        raise EvalError(f"function {name}")
