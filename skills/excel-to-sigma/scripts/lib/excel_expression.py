"""Safe parsed Excel expressions shared by migration analysis and rendering.

The parser intentionally supports a bounded scalar formula language. Unsupported
or external constructs raise :class:`UnsupportedExpression`; source text is never
passed to Python ``eval``.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
import math
import re
from typing import Callable, Iterable

from openpyxl.utils.cell import column_index_from_string, get_column_letter


class ExpressionError(ValueError):
    pass


class UnsupportedExpression(ExpressionError):
    pass


@dataclass(frozen=True)
class CellReference:
    sheet: str | None
    address: str | None
    column_absolute: bool = False
    row_absolute: bool = False
    name: str | None = None


@dataclass(frozen=True)
class Number:
    value: float | int
    source: str


@dataclass(frozen=True)
class Text:
    value: str


@dataclass(frozen=True)
class Reference:
    value: CellReference


@dataclass(frozen=True)
class Range:
    start: CellReference
    end: CellReference


@dataclass(frozen=True)
class Unary:
    operator: str
    operand: object


@dataclass(frozen=True)
class Percent:
    operand: object


@dataclass(frozen=True)
class Binary:
    operator: str
    left: object
    right: object


@dataclass(frozen=True)
class Function:
    name: str
    arguments: tuple[object, ...]


@dataclass(frozen=True)
class Group:
    expression: object


_CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$")
_CELL_TOKEN = re.compile(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*")
_NUMBER = re.compile(r"(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_IDENT = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")
_SUPPORTED_FUNCTIONS = {"SUM", "SUMIFS", "IF", "ROUND", "ABS", "MIN", "MAX", "INT"}


def _unquote_sheet(value: str) -> str:
    if value.startswith("'"):
        return value[1:-1].replace("''", "'")
    return value


class _Lexer:
    def __init__(self, source: str):
        self.source = source
        self.position = 0

    def tokens(self):
        result = []
        while self.position < len(self.source):
            ch = self.source[self.position]
            if ch.isspace():
                self.position += 1
                continue
            if ch in "{}[]#;@":
                raise UnsupportedExpression(f"unsupported formula token {ch!r}")
            if ch == '"':
                result.append(("TEXT", self._quoted_text()))
                continue
            if ch == "'":
                result.append(("SHEET", self._quoted_sheet()))
                continue
            number = _NUMBER.match(self.source, self.position)
            if number:
                value = number.group(0)
                self.position = number.end()
                result.append(("NUMBER", value))
                continue
            cell = _CELL_TOKEN.match(self.source, self.position)
            if cell:
                value = cell.group(0)
                self.position = cell.end()
                result.append(("IDENT", value))
                continue
            ident = _IDENT.match(self.source, self.position)
            if ident:
                value = ident.group(0)
                self.position = ident.end()
                if self.position < len(self.source) and self.source[self.position] == "!":
                    self.position += 1
                    result.append(("SHEET", value))
                else:
                    result.append(("IDENT", value))
                continue
            two = self.source[self.position:self.position + 2]
            if two in (">=", "<=", "<>"):
                result.append(("OP", two)); self.position += 2; continue
            if ch in "+-*/^%=><(),:":
                result.append(("OP", ch)); self.position += 1; continue
            raise UnsupportedExpression(f"unsupported formula text near {self.source[self.position:]!r}")
        result.append(("EOF", ""))
        return result

    def _quoted_text(self):
        self.position += 1
        out = []
        while self.position < len(self.source):
            ch = self.source[self.position]
            self.position += 1
            if ch == '"':
                if self.position < len(self.source) and self.source[self.position] == '"':
                    out.append('"'); self.position += 1; continue
                return "".join(out)
            out.append(ch)
        raise ExpressionError("unterminated string literal")

    def _quoted_sheet(self):
        start = self.position
        self.position += 1
        while self.position < len(self.source):
            if self.source[self.position] == "'":
                if self.position + 1 < len(self.source) and self.source[self.position + 1] == "'":
                    self.position += 2; continue
                self.position += 1
                if self.position >= len(self.source) or self.source[self.position] != "!":
                    raise UnsupportedExpression("quoted text is only supported as a sheet reference")
                value = self.source[start:self.position]
                self.position += 1
                if "[" in value or "]" in value:
                    raise UnsupportedExpression("external workbook references are unsupported")
                return value
            self.position += 1
        raise ExpressionError("unterminated quoted sheet name")


class _Parser:
    def __init__(self, source: str, sheet: str | None):
        self.tokens = _Lexer(source).tokens()
        self.index = 0
        self.sheet = sheet

    def parse(self):
        result = self._comparison()
        if self._peek()[0] != "EOF":
            raise UnsupportedExpression(f"unexpected token {self._peek()[1]!r}")
        return result

    def _peek(self):
        return self.tokens[self.index]

    def _take(self):
        value = self.tokens[self.index]
        self.index += 1
        return value

    def _accept(self, value):
        if self._peek() == ("OP", value):
            self.index += 1
            return True
        return False

    def _comparison(self):
        node = self._additive()
        while self._peek() in (("OP", "="), ("OP", "<>"), ("OP", "<"),
                               ("OP", ">"), ("OP", "<="), ("OP", ">=")):
            op = self._take()[1]
            node = Binary(op, node, self._additive())
        return node

    def _additive(self):
        node = self._multiplicative()
        while self._peek() in (("OP", "+"), ("OP", "-")):
            op = self._take()[1]
            node = Binary(op, node, self._multiplicative())
        return node

    def _multiplicative(self):
        node = self._power()
        while self._peek() in (("OP", "*"), ("OP", "/")):
            op = self._take()[1]
            node = Binary(op, node, self._power())
        return node

    def _power(self):
        node = self._unary()
        if self._accept("^"):
            node = Binary("^", node, self._power())
        return node

    def _unary(self):
        if self._peek() in (("OP", "+"), ("OP", "-")):
            return Unary(self._take()[1], self._unary())
        node = self._primary()
        while self._accept("%"):
            node = Percent(node)
        return node

    def _primary(self):
        kind, value = self._take()
        if kind == "NUMBER":
            return Number(float(value) if any(c in value for c in ".eE") else int(value), value)
        if kind == "TEXT":
            return Text(value)
        if kind == "SHEET":
            return self._reference(_unquote_sheet(value))
        if kind == "IDENT":
            if self._accept("("):
                name = value.upper()
                if name not in _SUPPORTED_FUNCTIONS:
                    raise UnsupportedExpression(f"unsupported function {name}")
                args = []
                if not self._accept(")"):
                    while True:
                        args.append(self._comparison())
                        if self._accept(")"):
                            break
                        if not self._accept(","):
                            raise ExpressionError("expected ',' or ')' in function")
                return Function(name, tuple(args))
            if _CELL.match(value):
                self.index -= 1
                return self._reference(self.sheet)
            return Reference(CellReference(None, None, name=value))
        if kind == "OP" and value == "(":
            expression = self._comparison()
            if not self._accept(")"):
                raise ExpressionError("unclosed parenthesis")
            return Group(expression)
        raise ExpressionError(f"expected expression, got {value!r}")

    def _reference(self, sheet):
        kind, value = self._take()
        if kind != "IDENT" or not _CELL.match(value):
            raise UnsupportedExpression("sheet names must be followed by a cell reference")
        start = _cell_reference(sheet, value)
        if not self._accept(":"):
            return Reference(start)
        end_sheet = sheet
        if self._peek()[0] == "SHEET":
            end_sheet = _unquote_sheet(self._take()[1])
        kind, value = self._take()
        if kind != "IDENT" or not _CELL.match(value):
            raise UnsupportedExpression("range endpoint must be a cell reference")
        return Range(start, _cell_reference(end_sheet, value))


def _cell_reference(sheet, address):
    match = _CELL.match(address)
    return CellReference(sheet, f"{match.group(2).upper()}{match.group(4)}",
                         bool(match.group(1)), bool(match.group(3)))


def parse_expression(formula: str, *, sheet: str | None = None):
    if not isinstance(formula, str) or not formula.startswith("="):
        raise ExpressionError("Excel expressions must start with '='")
    source = formula[1:].strip()
    if not source:
        raise ExpressionError("empty formula")
    if "[" in source or "]" in source:
        raise UnsupportedExpression("external or structured references are unsupported")
    return _Parser(source, sheet).parse()


def _walk(node) -> Iterable[CellReference]:
    if isinstance(node, Reference):
        yield node.value
    elif isinstance(node, Range):
        yield node.start; yield node.end
    elif isinstance(node, (Unary, Percent, Group)):
        yield from _walk(node.operand if not isinstance(node, Group) else node.expression)
    elif isinstance(node, Binary):
        yield from _walk(node.left); yield from _walk(node.right)
    elif isinstance(node, Function):
        for argument in node.arguments:
            yield from _walk(argument)


def extract_references(expression) -> tuple[CellReference, ...]:
    return tuple(_walk(expression))


def expand_range(value: Range) -> tuple[CellReference, ...]:
    if value.start.name or value.end.name or value.start.sheet != value.end.sheet:
        raise UnsupportedExpression("ranges must use cells on one sheet")
    start = _CELL.match(value.start.address)
    end = _CELL.match(value.end.address)
    c1, c2 = column_index_from_string(start.group(2)), column_index_from_string(end.group(2))
    r1, r2 = int(start.group(4)), int(end.group(4))
    refs = []
    for row in range(min(r1, r2), max(r1, r2) + 1):
        for col in range(min(c1, c2), max(c1, c2) + 1):
            refs.append(CellReference(value.start.sheet, f"{get_column_letter(col)}{row}"))
    return tuple(refs)


def evaluate(expression, resolver: Callable[[CellReference], object]):
    def run(node):
        if isinstance(node, Number): return node.value
        if isinstance(node, Text): return node.value
        if isinstance(node, Reference): return resolver(node.value)
        if isinstance(node, Range): return [resolver(ref) for ref in expand_range(node)]
        if isinstance(node, Group): return run(node.expression)
        if isinstance(node, Unary):
            value = run(node.operand)
            return value if node.operator == "+" else -value
        if isinstance(node, Percent): return run(node.operand) * 0.01
        if isinstance(node, Binary):
            left, right = run(node.left), run(node.right)
            operations = {
                "+": lambda: left + right, "-": lambda: left - right,
                "*": lambda: left * right, "/": lambda: left / right,
                "^": lambda: left ** right, "=": lambda: left == right,
                "<>": lambda: left != right, "<": lambda: left < right,
                ">": lambda: left > right, "<=": lambda: left <= right,
                ">=": lambda: left >= right,
            }
            return operations[node.operator]()
        if isinstance(node, Function):
            if node.name == "IF":
                if len(node.arguments) not in (2, 3):
                    raise ExpressionError("IF requires two or three arguments")
                condition = run(node.arguments[0])
                if condition:
                    return run(node.arguments[1])
                return run(node.arguments[2]) if len(node.arguments) == 3 else False
            args = [run(arg) for arg in node.arguments]
            if node.name == "ROUND":
                if len(args) != 2:
                    raise ExpressionError("ROUND requires two arguments")
                places = int(args[1])
                quantum = Decimal("1").scaleb(-places)
                return float(Decimal(str(args[0])).quantize(quantum, rounding=ROUND_HALF_UP))
            if node.name == "ABS": return abs(args[0])
            if node.name == "INT": return math.floor(args[0])
            if node.name == "MIN": return min(_flatten(args))
            if node.name == "MAX": return max(_flatten(args))
            if node.name == "SUM": return sum(v for v in _flatten(args) if isinstance(v, (int, float)))
            if node.name == "SUMIFS":
                if len(args) < 3 or len(args) % 2 == 0:
                    raise ExpressionError("SUMIFS requires sum range plus range/criterion pairs")
                totals = args[0]
                if not isinstance(totals, list): totals = [totals]
                conditions = []
                for index in range(1, len(args), 2):
                    values = args[index] if isinstance(args[index], list) else [args[index]]
                    if len(values) != len(totals):
                        raise ExpressionError("SUMIFS ranges must have equal sizes")
                    conditions.append((values, args[index + 1]))
                return sum(value for i, value in enumerate(totals)
                           if isinstance(value, (int, float))
                           and all(values[i] == criterion for values, criterion in conditions))
        raise UnsupportedExpression(f"cannot evaluate {type(node).__name__}")
    return run(expression)


def _flatten(values):
    for value in values:
        if isinstance(value, list):
            yield from value
        else:
            yield value


def emit_sigma(expression, reference_renderer: Callable[[CellReference], str],
               range_renderer: Callable[[Range], str] | None = None) -> str:
    function_names = {"IF": "If", "ROUND": "Round", "ABS": "Abs", "MIN": "Min",
                      "MAX": "Max", "SUM": "Sum", "INT": "Floor", "SUMIFS": "SumIfs"}

    def render(node):
        if isinstance(node, Number): return node.source
        if isinstance(node, Text): return '"' + node.value.replace('"', '\\"') + '"'
        if isinstance(node, Reference): return reference_renderer(node.value)
        if isinstance(node, Range):
            if range_renderer is not None:
                return range_renderer(node)
            return f"Range({reference_renderer(node.start)}, {reference_renderer(node.end)})"
        if isinstance(node, Group): return f"({render(node.expression)})"
        if isinstance(node, Unary): return f"{node.operator}{render(node.operand)}"
        if isinstance(node, Percent): return f"({render(node.operand)} * 0.01)"
        if isinstance(node, Binary):
            operator = "!=" if node.operator == "<>" else node.operator
            return f"{render(node.left)} {operator} {render(node.right)}"
        if isinstance(node, Function):
            return f"{function_names[node.name]}(" + ", ".join(render(arg) for arg in node.arguments) + ")"
        raise UnsupportedExpression(f"cannot emit {type(node).__name__}")
    return render(expression)
