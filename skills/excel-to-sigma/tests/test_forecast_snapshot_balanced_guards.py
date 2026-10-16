"""Engine-depth regression; symbolic construction is not live acceptance."""
import ast
from dataclasses import dataclass
import hashlib
import json
import re
import unittest
import test_forecast_snapshot as fixture


@dataclass(eq=False)
class FormulaNode:
    operator: str
    arguments: tuple


def freeze(value):
    if isinstance(value, dict):
        return tuple((freeze(k), freeze(v)) for k, v in value.items())
    if isinstance(value, (tuple, list)):
        return tuple(freeze(v) for v in value)
    return value


def children(node):
    def walk(value):
        if isinstance(value, FormulaNode):
            yield value
        elif isinstance(value, tuple):
            for child in value:
                yield from walk(child)
    return tuple(child for value in node.arguments for child in walk(value))


def depth(node):
    return 1 + max((depth(child) for child in children(node)), default=0)


def canonical(value):
    """Flatten only equal AND/OR operators; preserve operand order/duplicates."""
    if not isinstance(value, FormulaNode):
        return tuple(canonical(v) for v in value) if isinstance(value, tuple) else value
    if value.operator in ('and_', 'or_'):
        def operands(node):
            if node.operator == value.operator:
                return tuple(leaf for child in children(node) for leaf in operands(child))
            return (canonical(node),)
        return value.operator, operands(value)
    return value.operator, tuple(canonical(v) for v in value.arguments)


def fingerprint(value):
    return hashlib.sha256(json.dumps(canonical(value), separators=(',', ':')).encode()).hexdigest()


def replay_condition(program):
    """Run only formula constructors/loops from real emitted code on inert nodes."""
    tree = ast.parse(program)
    def function(name):
        return lambda *args: FormulaNode(name, tuple(freeze(v) for v in args))
    environment = {n.func.id: function(n.func.id) for n in ast.walk(tree)
                   if isinstance(n, ast.Call) and isinstance(n.func, ast.Name)}
    environment.update(range=range, len=len, RuntimeError=RuntimeError)
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and re.fullmatch(r'c\d+_\d+', node.id):
            environment[node.id] = FormulaNode('column', (node.id,))
    environment.update(op=FormulaNode('control', ('op-control',)),
                       t2=FormulaNode('table', ('snapshots',)))
    formula_variables = {'frozen_assumptions', 'current_assumptions', 'frozen_columns',
        'snapshot_keys', 'source_cells', 'source_rows', 'line_items', 'sections',
        'periods', 'capture_time', 'confirm_guard', 'confirm_guard_parts', 'allowed',
        'allowed_parts', 'effects'}
    for node in tree.body:
        selected = (isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name)
                    and (node.targets[0].id in formula_variables
                         or bool(re.fullmatch(r'a\d+', node.targets[0].id))))
        selected |= (isinstance(node, ast.For) and isinstance(node.iter, ast.Call)
                     and isinstance(node.iter.func, ast.Name) and node.iter.func.id == 'range')
        selected |= (isinstance(node, ast.Expr) and isinstance(node.value, ast.Call)
                     and isinstance(node.value.func, ast.Attribute)
                     and isinstance(node.value.func.value, ast.Name)
                     and node.value.func.value.id == 'confirm_guard_parts'
                     and node.value.func.attr == 'append')
        if selected:
            exec(compile(ast.Module(body=[node], type_ignores=[]), '<formula-only>', 'exec'), environment)
    condition = next(n.args[0] for n in ast.walk(tree) if isinstance(n, ast.Call)
                     and isinstance(n.func, ast.Attribute) and n.func.attr == 'set_condition')
    return eval(compile(ast.Expression(condition), '<condition-only>', 'eval'), environment), environment['effects']


class BalancedSnapshotGuardTests(unittest.TestCase):
    def condition(self):
        p, definition, current, bindings = fixture.SnapshotTests().candidate_fixture()
        prepared = p.render_snapshot_candidate(definition, current, bindings, phase='prepare')
        current['contents']['elements'].append({'id':'op-control', 'kind':'control',
            'controlType':'text', 'controlId':'forecast-snapshot-operation', 'value':''})
        current['contents']['layout'] = current['contents']['layout'].replace(
            '<Element elementId="operations" />', '<Element elementId="op-control" /><Element elementId="operations" />')
        candidate = p.render_snapshot_candidate(definition, current, bindings, phase='confirm',
            operation_control_id='op-control', prepared_contract=prepared['preparedContract'])
        return replay_condition(candidate['program'])

    def test_complete_confirm_condition_avoids_deep_boolean_folds(self):
        condition, effects = self.condition()
        self.assertLessEqual(depth(condition), 30)
        self.assertEqual(len(effects), 108)

    def test_boolean_reassociation_preserves_every_ordered_guard_and_effect(self):
        condition, effects = self.condition()
        # Reviewed pre-balancing baseline. Constants, lookup keys, 96 nullable
        # equality guards, 108 duplicate/conflict guards, and the foreign-key
        # blocker are all covered. Reassociation may not alter any operand.
        self.assertEqual(fingerprint(condition), 'bf45531bb374c3a77ca0c64d52e426410881d177485ca819a97c0217da1858be')
        self.assertEqual(fingerprint(tuple(effects)), 'd7099a257a5a05daf55254ac04541f860f60c499d95761be527d7331b600b1f9')


if __name__ == '__main__':
    unittest.main()
