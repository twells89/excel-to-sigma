"""Shared calculation for forecast plans: periods, same-column row expressions, recomputation.

``forecast_plan`` (the drafter) and ``build-forecast-model.py`` (the builder) both
need to answer "what does this plan evaluate to": the drafter to prove a plan
reproduces the workbook's cached values, the builder to run the local parity check.
Keeping one implementation means the two cannot drift.

Row expressions come from derived rows such as ``=B9-B14`` or
``=IF(B13=0,0,B22/B13)``. They are parsed into a small JSON-able tree whose leaves
name plan entities (``sec`` section totals, ``line`` P&L lines, ``drv`` other derived
rows), so a plan carries the formula without any cell addresses.
"""
from __future__ import annotations

import math
import re


# The four classic P&L rows keep these canonical names in parity targets and chart series.
TARGET_LABELS = {"total revenue": "Revenue", "gross profit": "Gross Profit",
                 "ebitda": "EBITDA", "net income": "Net Income"}


class ExprError(ValueError):
    """A row formula is outside the supported grammar, or cannot be evaluated."""


class RowEvalError(ExprError):
    """A derived row cannot be evaluated in one period; names the row and the period."""

    def __init__(self, label, index, period, cause):
        self.label, self.index, self.period = label, index, period
        super().__init__(f"derived row {label!r} in {period}: "
                         f"{cause if isinstance(cause, ExprError) else type(cause).__name__ + ': ' + str(cause)}")


# ---------------------------------------------------------------- tokenising / parsing
_TOKEN = re.compile(r"""\s*(?:
      (?P<num>\d+(?:\.\d+)?)
    | (?P<str>"[^"]*")
    | (?P<func>[A-Za-z_][A-Za-z0-9_.]*)\s*(?=\()
    | (?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?\$?[A-Z]{1,3}\$?\d+)
    | (?P<op><>|<=|>=|[-+*/()=<>,:])
)""", re.X)
_REF = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?\$?([A-Z]{1,3})\$?(\d+)$")
SUPPORTED_CALLS = {"SUM", "IF", "IFERROR", "MAX", "MIN", "ABS"}
# (fewest, most) arguments Excel accepts; None means no upper bound.
ARITY = {"SUM": (1, None), "MAX": (1, None), "MIN": (1, None), "ABS": (1, 1),
         "IF": (2, 3), "IFERROR": (2, 2)}


def check_arity(name, count):
    """Raise ``ExprError`` when ``name`` is called with an argument count Excel rejects."""
    low, high = ARITY[name]
    if count < low or (high is not None and count > high):
        wanted = f"{low}" if low == high else (f"at least {low}" if high is None else f"{low} or {high}")
        raise ExprError(f"{name} takes {wanted} argument(s), got {count}")


def _tokens(text):
    out, pos = [], 0
    text = text.strip()
    while pos < len(text):
        found = _TOKEN.match(text, pos)
        if not found or found.end() == pos:
            raise ExprError(f"unsupported syntax near {text[pos:pos + 12]!r}")
        pos = found.end()
        kind = found.lastgroup
        out.append((kind, found.group(kind)))
    return out


class _Parser:
    def __init__(self, text):
        if not text.startswith("="):
            raise ExprError("not a formula")
        self.toks = _tokens(text[1:])
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self, value=None):
        tok = self.peek()
        if tok[0] is None or (value is not None and tok[1] != value):
            raise ExprError(f"expected {value!r}")
        self.i += 1
        return tok

    def parse(self):
        node = self.comparison()
        if self.i != len(self.toks):
            raise ExprError(f"unexpected {self.peek()[1]!r}")
        return node

    def comparison(self):
        left = self.additive()
        if self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            return ("cmp", op, left, self.additive())
        return left

    def additive(self):
        node = self.term()
        while self.peek()[1] in ("+", "-") and self.peek()[0] == "op":
            op = self.take()[1]
            node = ("bin", op, node, self.term())
        return node

    def term(self):
        node = self.unary()
        while self.peek()[1] in ("*", "/") and self.peek()[0] == "op":
            op = self.take()[1]
            node = ("bin", op, node, self.unary())
        return node

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.unary())
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        return self.primary()

    def primary(self):
        kind, value = self.peek()
        if kind == "num":
            self.take()
            return ("num", float(value))
        if kind == "str":
            self.take()
            return ("str", value[1:-1])
        if kind == "ref":
            self.take()
            first = _ref_node(value)
            if self.peek() == ("op", ":"):
                self.take()
                second = _ref_node(self.take()[1])
                return ("range", first[1], first[2], first[3], second[2], second[3])
            return first
        if kind == "func":
            name = self.take()[1].upper()
            if name not in SUPPORTED_CALLS:
                raise ExprError(f"unsupported function {name}")
            self.take("(")
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.comparison())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.comparison())
            self.take(")")
            check_arity(name, len(args))
            return ("call", name, args)
        if self.peek() == ("op", "("):
            self.take()
            node = self.comparison()
            self.take(")")
            return node
        raise ExprError(f"unexpected {value!r}")


def _ref_node(text):
    found = _REF.match(text)
    if not found:
        raise ExprError(f"bad reference {text!r}")
    sheet = (found.group(1) or found.group(2) or None)
    if sheet:
        sheet = sheet.replace("''", "'")
    return ("ref", sheet, found.group(3), int(found.group(4)))


def parse_formula(text):
    """Parse ``=...`` into an address-level tree (``ref`` / ``range`` leaves)."""
    return _Parser(text).parse()


def relative_shape(node, col_index):
    """The tree with column letters replaced by offsets from ``col_index``.

    Two period columns hold "the same formula" exactly when their relative shapes
    are equal.
    """
    from openpyxl.utils import column_index_from_string
    kind = node[0]
    if kind == "ref":
        return ("ref", node[1], column_index_from_string(node[2]) - col_index, node[3])
    if kind == "range":
        return ("range", node[1], column_index_from_string(node[2]) - col_index, node[3],
                column_index_from_string(node[4]) - col_index, node[5])
    if kind in ("bin", "cmp"):
        return (kind, node[1], relative_shape(node[2], col_index), relative_shape(node[3], col_index))
    if kind == "neg":
        return ("neg", relative_shape(node[1], col_index))
    if kind == "call":
        return ("call", node[1], tuple(relative_shape(a, col_index) for a in node[2]))
    return node


def to_entities(node, resolve):
    """Convert an address-level tree to a plan expression.

    ``resolve(sheet, col, row)`` returns ``["sec", name]``, ``["line", label]`` or
    ``["drv", label]`` and raises ``ExprError`` for anything else. A range inside
    ``SUM`` is expanded through ``resolve_range(sheet, col, row1, row2)`` supplied as
    ``resolve.range``.
    """
    kind = node[0]
    if kind == "num":
        return ["num", node[1]]
    if kind == "str":
        return ["str", node[1]]
    if kind == "ref":
        return resolve(node[1], node[2], node[3])
    if kind == "range":
        raise ExprError("a cell range is only supported inside SUM(...)")
    if kind == "neg":
        return ["neg", to_entities(node[1], resolve)]
    if kind in ("bin", "cmp"):
        return [kind, node[1], to_entities(node[2], resolve), to_entities(node[3], resolve)]
    if kind == "call":
        args = []
        for arg in node[2]:
            if arg[0] == "range":
                if node[1] != "SUM":
                    raise ExprError("a cell range is only supported inside SUM(...)")
                if arg[2] != arg[4]:
                    raise ExprError("SUM over a multi-column range is not a same-column total")
                args.extend(resolve.range(arg[1], arg[2], arg[3], arg[5]))
            else:
                args.append(to_entities(arg, resolve))
        return ["call", node[1], args]
    raise ExprError(f"unsupported node {kind}")


# ---------------------------------------------------------------- evaluating / rendering
def eval_expr(expr, env):
    """Evaluate a plan expression; ``env`` offers ``sec(name)``, ``line(label)``, ``drv(label)``."""
    kind = expr[0]
    if kind in ("num", "str"):
        return expr[1]
    if kind == "sec":
        return env.sec(expr[1])
    if kind == "line":
        return env.line(expr[1])
    if kind == "drv":
        return env.drv(expr[1])
    if kind == "neg":
        return -eval_expr(expr[1], env)
    if kind == "bin":
        a, b = eval_expr(expr[2], env), eval_expr(expr[3], env)
        if expr[1] == "+":
            return a + b
        if expr[1] == "-":
            return a - b
        if expr[1] == "*":
            return a * b
        if b == 0:
            raise ZeroDivisionError("division by zero")
        return a / b
    if kind == "cmp":
        a, b = eval_expr(expr[2], env), eval_expr(expr[3], env)
        return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[expr[1]]
    if kind == "call":
        name, args = expr[1], expr[2]
        if name not in ARITY:
            raise ExprError(f"unsupported function {name}")
        check_arity(name, len(args))
        if name == "IFERROR":
            try:
                return eval_expr(args[0], env)
            except (ZeroDivisionError, TypeError, ExprError):
                return eval_expr(args[1], env)
        if name == "IF":
            if eval_expr(args[0], env):
                return eval_expr(args[1], env)
            # Excel: IF(cond, x) with a false condition is FALSE (0 in arithmetic).
            return eval_expr(args[2], env) if len(args) == 3 else False
        values = [eval_expr(a, env) for a in args]
        if name == "ABS":
            return abs(values[0])
        return {"SUM": sum, "MAX": max, "MIN": min}[name](values)
    raise ExprError(f"unsupported node {kind}")


_PREC = {"cmp": 0, "+": 1, "-": 1, "*": 2, "/": 2}


def render_expr(expr, names=None):
    """Readable formula text over ``[Label]`` references."""
    kind = expr[0]
    if kind == "num":
        return f"{expr[1]:g}"
    if kind == "str":
        return f'"{expr[1]}"'
    if kind in ("sec", "line", "drv"):
        return f"[{expr[1]}]"
    if kind == "neg":
        return "-" + render_expr(expr[1])
    if kind in ("bin", "cmp"):
        op = expr[1]
        mine = _PREC["cmp"] if kind == "cmp" else _PREC[op]

        def side(child, right):
            text = render_expr(child)
            if child[0] in ("bin", "cmp"):
                theirs = _PREC["cmp"] if child[0] == "cmp" else _PREC[child[1]]
                if theirs < mine or (right and theirs == mine and op in ("-", "/")):
                    return f"({text})"
            return text
        return f"{side(expr[2], False)} {op} {side(expr[3], True)}"
    if kind == "call":
        return f"{expr[1]}(" + ", ".join(render_expr(a) for a in expr[2]) + ")"
    raise ExprError(f"unsupported node {kind}")


def linear_terms(expr, derived_exprs, _stack=()):
    """Express ``expr`` as ``{("sec", name) | ("line", label): coefficient}``; ``None`` if not linear.

    A pure amount row (totals, gross profit, net income, a line-to-line difference) is
    linear with no constant term; a ratio or an ``IF`` is not. Section totals and
    individual lines stay separate terms (a line is not a section), derived rows are
    expanded, and a term that cancels to zero is dropped.
    """
    kind = expr[0]
    if kind in ("sec", "line"):
        return {(kind, expr[1]): 1.0}
    if kind == "drv":
        if expr[1] in _stack or expr[1] not in derived_exprs:
            return None
        return linear_terms(derived_exprs[expr[1]], derived_exprs, _stack + (expr[1],))
    if kind == "neg":
        inner = linear_terms(expr[1], derived_exprs, _stack)
        return None if inner is None else {k: -v for k, v in inner.items()}
    if kind == "bin":
        op = expr[1]
        if op in ("+", "-"):
            a = linear_terms(expr[2], derived_exprs, _stack)
            b = linear_terms(expr[3], derived_exprs, _stack)
            if a is None or b is None:
                return None
            out = dict(a)
            for key, value in b.items():
                out[key] = out.get(key, 0.0) + (value if op == "+" else -value)
            return {k: v for k, v in out.items() if v != 0}
        const = lambda e: e[0] == "num"
        if op == "*" and (const(expr[2]) or const(expr[3])):
            number, other = (expr[2][1], expr[3]) if const(expr[2]) else (expr[3][1], expr[2])
        elif op == "/" and const(expr[3]) and expr[3][1] != 0:
            number, other = 1.0 / expr[3][1], expr[2]
        else:
            return None
        inner = linear_terms(other, derived_exprs, _stack)
        return None if inner is None else {k: v * number for k, v in inner.items() if v * number != 0}
    if kind == "call" and expr[1] == "SUM":
        out = {}
        for arg in expr[2]:
            part = linear_terms(arg, derived_exprs, _stack)
            if part is None:
                return None
            for key, value in part.items():
                out[key] = out.get(key, 0.0) + value
        return {k: v for k, v in out.items() if v != 0}
    return None


def fold_terms(terms, section_lines):
    """Split linear ``terms`` into ``(section coefficients, line coefficients)``.

    ``section_lines`` maps each section to its line labels. A section total is the sum of
    its lines, so lines that enter with one common weight fold back into a single section
    coefficient (``=EBITDA-B19-B20`` over both lines of Below the Line is one ``-1``).
    Lines that enter with different weights stay line-level, and are never collapsed into
    their section: ``=Subscription - COGS`` is one revenue line, not all of revenue.
    Coefficients are ``{name: weight}`` in the order of ``section_lines``.
    """
    sections, lines = {}, {}
    for section, labels in section_lines.items():
        whole = terms.get(("sec", section), 0.0)
        parts = [terms.get(("line", label), 0.0) for label in labels]
        if parts and all(abs(p - parts[0]) <= 1e-12 for p in parts):
            weight = whole + parts[0]
            if weight != 0:
                sections[section] = weight
            continue
        if whole != 0:
            sections[section] = whole
        for label, part in zip(labels, parts):
            if part != 0:
                lines[label] = part
    return sections, lines


def linear_over_sections(expr, line_section, derived_exprs, _stack=()):
    """``{section: coefficient}`` when ``expr`` is linear over whole sections, else ``None``.

    Kept for callers that only handle section-level rows; an expression that needs an
    individual line (one that does not enter with its section's other lines) is ``None``,
    never silently widened to the whole section.
    """
    terms = linear_terms(expr, derived_exprs, _stack)
    if terms is None:
        return None
    section_lines = {}
    for label, section in line_section.items():
        section_lines.setdefault(section, []).append(label)
    for kind, name in terms:
        if kind == "sec":
            section_lines.setdefault(name, [])
    sections, lines = fold_terms(terms, section_lines)
    return None if lines else sections


def derived_references(expr):
    """Labels of the derived rows ``expr`` reads (direct references only)."""
    kind = expr[0]
    if kind == "drv":
        return {expr[1]}
    if kind in ("bin", "cmp"):
        return derived_references(expr[2]) | derived_references(expr[3])
    if kind == "neg":
        return derived_references(expr[1])
    if kind == "call":
        return set().union(*(derived_references(a) for a in expr[2])) if expr[2] else set()
    return set()


def close(got, want, tolerance=None):
    """Whether a recomputed value equals a workbook value.

    Both sides come from the same arithmetic (the oracle recomputes what Excel computed),
    so only float noise and Excel's 15-significant-digit storage may differ. Default:
    ``abs(got - want) <= max(1e-6, 1e-9 * abs(want))``. An explicit ``tolerance`` is an
    absolute bound the caller chose (``--tolerance`` / ``--parity-tolerance``).
    """
    if tolerance is not None:
        return abs(got - want) <= tolerance
    return abs(got - want) <= max(1e-6, 1e-9 * abs(want))


# ---------------------------------------------------------------- plan periods
def plan_periods(plan):
    """``(periods, n_actual)``; ``periods`` are dicts with label, num, end, grain, role.

    New plans carry ``periods``; older month plans carry ``actualMonths`` /
    ``forecastMonths`` and are read as month periods.
    """
    if plan.get("periods"):
        periods = [dict(label=p["label"], num=p["num"], end=p.get("end"),
                        grain=p.get("grain", "month"), role=p["role"]) for p in plan["periods"]]
    else:
        periods = [dict(label=m["month"], num=m["num"], end=m.get("end"), grain="month", role=role)
                   for role, key in (("actual", "actualMonths"), ("forecast", "forecastMonths"))
                   for m in plan[key]]
    n_act = sum(1 for p in periods if p["role"] == "actual")
    if n_act == 0 or any(p["role"] == "actual" for p in periods[n_act:]):
        raise ExprError("plan periods must be actual periods followed by forecast periods")
    return periods, n_act


def month_views(periods):
    """Periods in the legacy ``{month, num, end}`` shape the builder's tables use."""
    return [{"month": p["label"], "num": p["num"], "end": p.get("end")} for p in periods]


def yoy_terms(index, last_actual, lag):
    """Exponent and base index for a lag-``lag`` growth cell at period ``index``.

    ``=<k periods back>*(1+rate)`` chains until it lands on an actual period.
    """
    hops = max(1, math.ceil((index - last_actual) / lag))
    return hops, index - hops * lag


# ---------------------------------------------------------------- recomputation
class Evaluation:
    """Every line / section / derived-row value a plan produces, per period."""

    def __init__(self, plan):
        self.plan = plan
        self.periods, self.n_act = plan_periods(plan)
        self.labels = [p["label"] for p in self.periods]
        last = self.n_act - 1
        index = {label: i for i, label in enumerate(self.labels[:self.n_act])}
        lines = plan["lines"]
        self.line_section = {ln["line"]: ln["section"] for ln in lines}
        self.sections = [s["name"] for s in plan["sections"]]
        for ln in lines:
            if ln["section"] not in self.sections:
                raise ExprError(f"line {ln['line']!r} is in section {ln['section']!r}, "
                                "which the plan does not declare")
        self.matrix = {ln["line"]: [0.0] * len(self.periods) for ln in lines}
        for row in plan["actuals"]:
            label = row["month"]
            if row["line"] not in self.matrix or label not in index:
                raise ExprError(f"actuals row for {row['line']!r} in {label!r} does not match the plan")
            self.matrix[row["line"]][index[label]] += row["amt"]
        deferred = []
        for ln in lines:
            values, base = self.matrix[ln["line"]], ln.get("juneBase", 0)
            kind = ln["kind"]
            if kind == "pct_of_total":
                deferred.append(ln)
                continue
            for i in range(self.n_act, len(self.periods)):
                j = i - self.n_act
                if kind == "manual":
                    values[i] = ln["manual"][j]
                elif kind == "flat":
                    values[i] = base
                elif kind == "yoy":
                    hops, at = yoy_terms(i, last, ln["lag"])
                    if at < 0:
                        raise ExprError(f"line {ln['line']!r}: lag {ln['lag']} reaches before the first period")
                    values[i] = values[at] * (1 + ln["rate"]) ** hops
                else:
                    values[i] = base * (1 + ln["rate"]) ** (self.periods[i]["num"] - self.periods[last]["num"])
        self.totals = {name: [0.0] * len(self.periods) for name in self.sections}
        for ln in lines:
            if ln["kind"] != "pct_of_total":
                for i, value in enumerate(self.matrix[ln["line"]]):
                    self.totals[ln["section"]][i] += value
        for ln in deferred:
            of = ln["of"]
            if of == ln["section"] or any(o["kind"] == "pct_of_total" and o["section"] == of
                                          for o in lines):
                raise ExprError(f"line {ln['line']!r}: percentage base {of!r} must be another section "
                                "without percentage lines")
            for i in range(self.n_act, len(self.periods)):
                self.matrix[ln["line"]][i] = self.totals[of][i] * ln["rate"]
        for ln in deferred:
            for i, value in enumerate(self.matrix[ln["line"]]):
                self.totals[ln["section"]][i] += value
        self.derived_exprs = {d["label"]: d["expr"] for d in plan.get("derivedRows", [])}
        self.derived = {label: self._row(label, range(len(self.periods)), per_period=True)
                        for label in self.derived_exprs}
        self.total_labels = {s["name"]: s.get("totalLabel") for s in plan["sections"]}

    # -- evaluation environments
    def _env(self, idxs, memo):
        ev = self

        class Env:
            def sec(self, name):
                return sum(ev.totals[name][i] for i in idxs)

            def line(self, label):
                return sum(ev.matrix[label][i] for i in idxs)

            def drv(self, label):
                if label not in memo:
                    if label in memo.get("__active__", ()):
                        raise ExprError(f"derived row {label!r} refers to itself")
                    memo.setdefault("__active__", set()).add(label)
                    memo[label] = eval_expr(ev.derived_exprs[label], self)
                    memo["__active__"].discard(label)
                return memo[label]
        return Env()

    def _value(self, label, idxs, period):
        """One derived-row value: a number (``FALSE`` reads as 0), or ``RowEvalError``."""
        try:
            value = eval_expr(self.derived_exprs[label], self._env(idxs, {}))
        except RowEvalError:
            raise
        except Exception as error:  # noqa: BLE001 - any evaluation failure is named, not leaked
            raise RowEvalError(label, idxs[0] if idxs else None, period, error) from None
        return float(value) if isinstance(value, bool) else value

    def _row(self, label, idxs, per_period):
        if per_period:
            return [self._value(label, [i], self.labels[i]) for i in idxs]
        return self._value(label, list(idxs), "the summed periods")

    def aggregate(self, label, idxs):
        """A derived row evaluated over the periods ``idxs`` summed (the FY column)."""
        return self._row(label, idxs, per_period=False)

    def section_total(self, name, idxs=None):
        idxs = range(len(self.periods)) if idxs is None else idxs
        return sum(self.totals[name][i] for i in idxs)

    def row_series(self, label):
        """Per-period values for a section-total label or a derived-row label, else None."""
        if label in self.derived:
            return self.derived[label]
        for name, total_label in self.total_labels.items():
            if total_label == label:
                return self.totals[name]
        return None
