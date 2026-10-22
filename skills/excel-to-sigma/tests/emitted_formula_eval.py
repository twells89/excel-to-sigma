"""Evaluate the Sigma formulas the forecast builder EMITS, over the plan's own rows.

The drafter's oracle (``forecast_calc.Evaluation``) answers "what does the plan compute".
The workbook spec answers "what did we ask Sigma to compute". A bug in the mapping from one
to the other (a P&L line widened to its whole section, a coefficient dropped) leaves both
the oracle and every local parity check green while Sigma shows a different number. This
module closes that gap: it parses the emitted formula text for the subset the builder
uses (``Sum`` / ``SumIf`` / ``If``, arithmetic, comparisons, ``and`` / ``or``, ``[Table/Col]``
references) and evaluates it over the rows ``pnlCombined`` would hold, so tests can compare
the emitted KPI, trend and chart formulas with the oracle for every derived row.
"""
from __future__ import annotations

import re

_TOKEN = re.compile(r"""\s*(?:
      (?P<num>\d+(?:\.\d+)?)
    | (?P<str>"(?:[^"]|"")*")
    | (?P<ref>\[[^\]]+\])
    | (?P<word>[A-Za-z_][A-Za-z0-9_]*)
    | (?P<op><>|<=|>=|[-+*/(),=<>^])
)""", re.X)


class FormulaError(ValueError):
    pass


def _tokens(text):
    out, pos = [], 0
    while pos < len(text):
        found = _TOKEN.match(text, pos)
        if not found or found.end() == pos:
            if not text[pos:].strip():
                break
            raise FormulaError(f"cannot read {text[pos:pos + 20]!r} in {text!r}")
        pos = found.end()
        out.append((found.lastgroup, found.group(found.lastgroup)))
    return out


class _Parse:
    def __init__(self, text):
        self.text, self.toks, self.i = text, _tokens(text), 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, value=None):
        tok = self.peek()
        if tok[0] is None or (value is not None and tok[1].lower() != value):
            raise FormulaError(f"expected {value!r} at {tok[1]!r} in {self.text!r}")
        self.i += 1
        return tok

    def parse(self):
        node = self.or_()
        if self.i != len(self.toks):
            raise FormulaError(f"trailing {self.peek()[1]!r} in {self.text!r}")
        return node

    def or_(self):
        node = self.and_()
        while self.peek()[0] == "word" and self.peek()[1].lower() == "or":
            self.take()
            node = ("or", node, self.and_())
        return node

    def and_(self):
        node = self.cmp()
        while self.peek()[0] == "word" and self.peek()[1].lower() == "and":
            self.take()
            node = ("and", node, self.cmp())
        return node

    def cmp(self):
        node = self.add()
        if self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            node = ("cmp", op, node, self.add())
        return node

    def add(self):
        node = self.mul()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            node = ("bin", op, node, self.mul())
        return node

    def mul(self):
        node = self.pow()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            node = ("bin", op, node, self.pow())
        return node

    def pow(self):
        node = self.unary()
        if self.peek() == ("op", "^"):
            self.take()
            node = ("bin", "^", node, self.pow())
        return node

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.unary())
        return self.primary()

    def primary(self):
        kind, value = self.peek()
        if kind == "num":
            self.take()
            return ("num", float(value))
        if kind == "str":
            self.take()
            return ("str", value[1:-1].replace('""', '"'))
        if kind == "ref":
            self.take()
            return ("ref", value[1:-1])
        if kind == "word":
            self.take()
            self.take("(")
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.or_())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.or_())
            self.take(")")
            return ("call", value.lower(), args)
        if self.peek() == ("op", "("):
            self.take()
            node = self.or_()
            self.take(")")
            return node
        raise FormulaError(f"unexpected {value!r} in {self.text!r}")


def parse(text):
    return _Parse(text).parse()


def _ref_name(ref):
    """``Table/Column`` -> ``Column`` (the rows carry columns by name)."""
    return ref.rsplit("/", 1)[-1]


def _eval(node, row, rows):
    kind = node[0]
    if kind in ("num", "str"):
        return node[1]
    if kind == "ref":
        if row is None:
            raise FormulaError(f"[{node[1]}] used outside an aggregate with no row in scope")
        name = _ref_name(node[1])
        if name not in row:
            raise FormulaError(f"unknown column {node[1]!r}")
        return row[name]
    if kind == "neg":
        return -_eval(node[1], row, rows)
    if kind == "bin":
        a, b = _eval(node[2], row, rows), _eval(node[3], row, rows)
        if a is None or b is None:
            return None  # SQL NULL propagates
        op = node[1]
        if op == "+":
            return a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if op == "^":
            return a ** b
        return None if b == 0 else a / b
    if kind == "cmp":
        a, b = _eval(node[2], row, rows), _eval(node[3], row, rows)
        if a is None or b is None:
            return None
        return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[node[1]]
    if kind in ("and", "or"):
        a, b = _eval(node[1], row, rows), _eval(node[2], row, rows)
        if kind == "and":   # three-valued logic
            if a is False or b is False:
                return False
            return None if a is None or b is None else True
        if a is True or b is True:
            return True
        return None if a is None or b is None else False
    if kind == "call":
        name, args = node[1], node[2]
        if name == "if":   # If(c1, v1, c2, v2, ..., default)
            pairs, default = args[:-1], args[-1]
            for cond, value in zip(pairs[::2], pairs[1::2]):
                if _eval(cond, row, rows) is True:
                    return _eval(value, row, rows)
            return _eval(default, row, rows)
        if rows is None:
            raise FormulaError(f"{name}() needs a row set")
        if name == "sum":
            values = [_eval(args[0], r, rows) for r in rows]
            values = [v for v in values if v is not None]
            return sum(values) if values else None
        if name == "sumif":
            values = [_eval(args[0], r, rows) for r in rows if _eval(args[1], r, rows) is True]
            values = [v for v in values if v is not None]
            return sum(values) if values else None
        raise FormulaError(f"unsupported function {name}")
    raise FormulaError(f"unsupported node {kind}")


def evaluate(formula, rows, row=None):
    """Value of ``formula`` over ``rows`` (aggregates run over them; refs read ``row``)."""
    return _eval(parse(formula), row, rows)


def combined_rows(plan, ev, wb_spec):
    """The rows ``pnlCombined`` holds: actuals then forecast, with its own column formulas applied.

    ``ev`` is the oracle's ``Evaluation`` (amounts per line and period). Per-row columns the
    spec defines (the section amount columns) are evaluated from their emitted formulas.
    """
    columns = None
    for page in wb_spec["pages"]:
        for element in page["elements"]:
            if element["id"] == "pnlCombined":
                columns = element["columns"]
    if columns is None:
        raise FormulaError("the workbook spec has no pnlCombined element")
    section_of = {ln["line"]: ln["section"] for ln in plan["lines"]}
    rows = []
    for label, series in ev.matrix.items():
        for i, amount in enumerate(series):
            rows.append({"Section": section_of[label], "Line Item": label, "Month": ev.labels[i],
                         "Month Num": ev.periods[i]["num"], "Amount": amount,
                         "Period": "Actual" if i < ev.n_act else "Forecast"})
    for column in columns:
        formula = column["formula"]
        if re.fullmatch(r"\[[^\]]+\]", formula):
            continue  # a plain pass-through of a base column
        for row in rows:
            row[column["name"]] = evaluate(formula, None, row)
    return rows
