"""Test-only evaluator of complete emitted tables, separate from the emitter.

Scalar operations use the documented-Sigma reference interpreter. This adapter
adds row provenance and native table Lookup/SumIf/CountIf/MinIf operations. It
executes no target calls and is not Sigma runtime evidence.
"""
import ast
import copy
import datetime as dt
import re
from native_formula_reference import evaluate_native


class NativeGraphReference:
    def __init__(self, inputs):
        self.tables = {}
        self.elements = {}
        for element_id, (name, rows) in inputs.items():
            self.elements[element_id] = name
            self.tables[name] = [{'origin': (element_id, i), 'values': dict(row)} for i, row in enumerate(rows)]

    def evaluate(self, formula, table, index):
        refs = {}
        # Preserve literal text, including strings that contain bracket tokens.
        parts = []; i = 0
        while i < len(formula):
            if formula[i] == '"':
                j = i + 1
                while j < len(formula):
                    if formula[j] == '\\': j += 2; continue
                    if formula[j] == '"': j += 1; break
                    j += 1
                parts.append(formula[i:j]); i = j
            elif formula[i] == '[':
                j = formula.index(']', i); token = 'graph_ref_' + str(len(refs))
                refs[token] = formula[i+1:j]; parts.append(token); i = j + 1
            else:
                parts.append(formula[i]); i += 1
        # The scalar reference handles Sigma equality syntax. Reuse its parser
        # by translating outside literals, without interpreting any function.
        raw = ''.join(parts)
        translated=[]; quote=None; position=0
        while position<len(raw):
            char=raw[position]
            if quote:
                translated.append(char)
                if char=='\\' and position+1<len(raw):position+=1;translated.append(raw[position])
                elif char==quote:quote=None
            elif char in ('"', "'"):quote=char;translated.append(char)
            elif char=='=' and (position==0 or raw[position-1]not in '<>!=') and (position+1==len(raw)or raw[position+1]!='='):
                translated.append('==')
            else:translated.append(char)
            position+=1
        raw=''.join(translated)
        tree = ast.parse(raw, mode='eval').body

        def column(node):
            if not isinstance(node, ast.Name) or node.id not in refs:
                raise AssertionError('aggregate/lookup key must be a declared column')
            value = refs[node.id]
            return value.split('/', 1) if '/' in value else (table, value)

        def ref_value(node, current_table, row_index):
            target, field = column(node)
            if target == current_table:
                return self.tables[target][row_index]['values'].get(field)
            origin = self.tables[current_table][row_index]['origin']
            matches = [row for row in self.tables[target] if row['origin'] == origin]
            if len(matches) != 1:
                raise AssertionError('cross-table row reference has no unique source provenance')
            return matches[0]['values'].get(field)

        def run(node, current_table=table, row_index=index):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == 'Lookup':
                if len(node.args) < 3 or len(node.args) % 2 != 1:
                    raise AssertionError('Lookup requires one or more local/related key pairs')
                value_node, key_node, key_column = node.args[:3]
                target, key_field = column(key_column)
                wanted = run(key_node, current_table, row_index)
                matches = [j for j, row in enumerate(self.tables[target]) if row['values'].get(key_field) == wanted]
                for extra_index in range(3, len(node.args), 2):
                    extra_target, extra_field = column(node.args[extra_index+1])
                    if extra_target != target:
                        raise AssertionError('Lookup related keys must share a table')
                    extra_wanted = run(node.args[extra_index], current_table, row_index)
                    matches = [j for j in matches if self.tables[target][j]['values'].get(extra_field) == extra_wanted]
                if isinstance(value_node, ast.Call) and value_node.func.id in ('SumIf', 'CountIf', 'MinIf'):
                    operation = value_node.func.id; args = value_node.args
                    if operation == 'CountIf':
                        return sum(bool(run(args[0], target, j)) for j in matches)
                    selected = [run(args[0], target, j) for j in matches if run(args[1], target, j)]
                    selected = [value for value in selected if value is not None]
                    return (sum(selected) if operation == 'SumIf' else min(selected)) if selected else None
                return run(value_node, target, matches[0]) if matches else None
            if isinstance(node, ast.Name) and node.id in refs:
                return ref_value(node, current_table, row_index)
            if isinstance(node, ast.Call) and node.func.id == 'If':
                return run(node.args[1] if run(node.args[0], current_table, row_index) else node.args[2], current_table, row_index)
            # Replace nested table operations with values, then let the separate
            # scalar implementation evaluate operators/functions independently.
            def literal(value):
                if isinstance(value, (dt.date, dt.datetime)):
                    return ast.Call(func=ast.Name(id='Date', ctx=ast.Load()),
                                    args=[ast.Constant(value.isoformat())], keywords=[])
                return ast.Constant(value)
            class Resolve(ast.NodeTransformer):
                def visit_Name(self, child):
                    if child.id in refs:
                        return literal(run(child, current_table, row_index))
                    return child
                def visit_Call(self, child):
                    if child.func.id in ('Lookup', 'If'):
                        return literal(run(child, current_table, row_index))
                    return self.generic_visit(child)
            return evaluate_native(ast.unparse(Resolve().visit(copy.deepcopy(node))))
        return run(tree)

    def materialize(self, shards):
        for shard in shards:
            name = shard['name']; source = shard['source']
            if source['kind'] == 'sql':
                rows = [{'origin': (shard['id'], 0), 'values': {'C0': 1}}]
            else:
                rows = copy.deepcopy(self.tables[self.elements[source['elementId']]])
            self.tables[name] = rows; self.elements[shard['id']] = name
            for column in shard['columns']:
                for index, row in enumerate(rows):
                    row['values'][column['name']] = self.evaluate(column['formula'], name, index)
        return self

    def output(self, lowered, source_cell):
        mapping = lowered['formulas'][source_cell]['targetMapping']
        table = self.tables[self.elements[mapping['logicalElementId']]]
        if len(table) != 1: raise AssertionError('public scalar output is not singleton')
        return {lane: table[0]['values'][column] for lane, column in mapping['logicalLaneColumnIds'].items()}
