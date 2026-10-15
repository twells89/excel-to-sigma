"""Cell-granular workbook dependency graphs for formula migration."""
from __future__ import annotations

from dataclasses import dataclass

try:
    from .excel_expression import (Range, Reference, UnsupportedExpression, evaluate,
                                   expand_range, parse_expression)
except ImportError:  # standalone tests/scripts that place lib directly on sys.path
    from excel_expression import (Range, Reference, UnsupportedExpression, evaluate,
                                  expand_range, parse_expression)


@dataclass(frozen=True, order=True)
class CellIdentity:
    sheet: str
    address: str

    @property
    def id(self):
        return f"{self.sheet}!{self.address}"


class FormulaGraph:
    def __init__(self, formulas, literals, dependencies, unresolved, unsupported, names=None):
        self.formulas = formulas
        self.literals = literals
        self.dependencies = dependencies
        self.unresolved_references = unresolved
        self.unsupported_formulas = unsupported
        self.names = names or {}

    @classmethod
    def from_workbook(cls, workbook):
        formulas, literals = {}, {}
        for ws in workbook.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    identity = CellIdentity(ws.title, cell.coordinate)
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas[identity] = cell.value
                    else:
                        literals[identity] = cell.value
        existing = set(formulas) | set(literals)
        names = _defined_names(workbook)
        dependencies, unresolved, unsupported = {}, {}, {}
        for identity, formula in formulas.items():
            try:
                expression = parse_expression(formula, sheet=identity.sheet)
                refs = tuple(_expression_cells(expression))
            except UnsupportedExpression as error:
                unsupported[identity] = str(error)
                refs = ()
            deps = set()
            for ref in refs:
                if ref.name and ref.name in names:
                    deps.add(names[ref.name])
                elif ref.sheet and ref.address:
                    deps.add(CellIdentity(ref.sheet, ref.address))
            deps = tuple(sorted(deps))
            dependencies[identity] = deps
            missing = tuple(dep for dep in deps if dep not in existing)
            if missing:
                unresolved[identity] = missing
        return cls(formulas, literals, dependencies, unresolved, unsupported, names)

    def cycles(self):
        formula_nodes = set(self.formulas)
        edges = {node: tuple(dep for dep in self.dependencies.get(node, ()) if dep in formula_nodes)
                 for node in formula_nodes}
        index = 0
        stack = []
        on_stack = set()
        indices, low = {}, {}
        result = []

        def strong_connect(node):
            nonlocal index
            indices[node] = low[node] = index
            index += 1
            stack.append(node); on_stack.add(node)
            for dep in edges[node]:
                if dep not in indices:
                    strong_connect(dep); low[node] = min(low[node], low[dep])
                elif dep in on_stack:
                    low[node] = min(low[node], indices[dep])
            if low[node] == indices[node]:
                component = []
                while True:
                    member = stack.pop(); on_stack.remove(member); component.append(member)
                    if member == node:
                        break
                self_loop = len(component) == 1 and component[0] in edges[component[0]]
                if len(component) > 1 or self_loop:
                    result.append(tuple(sorted(component)))

        for node in sorted(formula_nodes):
            if node not in indices:
                strong_connect(node)
        return sorted(result)

    def topological_order(self):
        cycle_members = {node for component in self.cycles() for node in component}
        nodes = set(self.formulas) - cycle_members
        edges = {node: {dep for dep in self.dependencies.get(node, ()) if dep in nodes}
                 for node in nodes}
        result = []
        ready = sorted(node for node, deps in edges.items() if not deps)
        while ready:
            node = ready.pop(0)
            result.append(node)
            for candidate in sorted(nodes - set(result)):
                if node in edges[candidate]:
                    edges[candidate].remove(node)
                    if not edges[candidate] and candidate not in ready:
                        ready.append(candidate); ready.sort()
        return tuple(result)

    def evaluate(self):
        """Evaluate every acyclic supported formula from workbook literal values."""
        values = dict(self.literals)
        expressions = {}
        for identity, formula in self.formulas.items():
            try:
                expressions[identity] = parse_expression(formula, sheet=identity.sheet)
            except UnsupportedExpression:
                continue
        for identity in self.topological_order():
            expression = expressions.get(identity)
            if expression is None:
                continue

            def resolve(reference):
                if reference.name:
                    target = self.names.get(reference.name)
                    if target is None:
                        raise UnsupportedExpression(f"unresolved defined name {reference.name}")
                    return values.get(target)
                return values.get(CellIdentity(reference.sheet, reference.address))
            try:
                values[identity] = evaluate(expression, resolve)
            except (ArithmeticError, TypeError, ValueError, UnsupportedExpression):
                continue
        return values


def _expression_cells(expression):
    if isinstance(expression, Reference):
        if expression.value.address or expression.value.name:
            yield expression.value
        return
    if isinstance(expression, Range):
        yield from expand_range(expression)
        return
    for name in ("expression", "operand", "left", "right"):
        child = getattr(expression, name, None)
        if child is not None:
            yield from _expression_cells(child)
    for argument in getattr(expression, "arguments", ()):
        yield from _expression_cells(argument)


def _defined_names(workbook):
    result = {}
    for definition in workbook.defined_names.values():
        try:
            destinations = list(definition.destinations)
        except (AttributeError, TypeError):
            continue
        if len(destinations) != 1:
            continue
        sheet, address = destinations[0]
        address = address.replace("$", "")
        if ":" not in address:
            result[definition.name] = CellIdentity(sheet, address)
    return result
