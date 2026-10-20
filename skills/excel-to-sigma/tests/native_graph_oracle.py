"""Small local interpreter for emitted formulas, independent of builder equations.

This is a bounded test oracle, not an implementation of Sigma null/error or UI
semantics. SQL source queries execute in SQLite; supported scalar expressions
are parsed by Python AST with explicit operations and lazy If/Switch branches.
"""
import ast
import datetime as dt
import math
import operator
import re
import sqlite3
from decimal import Decimal, ROUND_HALF_UP


def scalar(formula, row):
    code = re.sub(r'\[([^\]]+)\]', lambda m: 'field(' + repr(m.group(1)) + ')', formula)
    code = re.sub(r'(?<![<>=!])=(?!=)', '==', code).replace('^', '**').replace('<>', '!=')
    ops = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
           ast.Div: operator.truediv, ast.Pow: operator.pow, ast.Eq: operator.eq, ast.NotEq: operator.ne,
           ast.Lt: operator.lt, ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge}
    def run(node):
        if isinstance(node, ast.Constant): return node.value
        if isinstance(node, ast.Name) and node.id == 'Null': return None
        if isinstance(node, ast.BinOp):
            a, b = run(node.left), run(node.right)
            return None if a is None or b is None else ops[type(node.op)](a, b)
        if isinstance(node, ast.Compare):
            a, b = run(node.left), run(node.comparators[0])
            return False if a is None or b is None else ops[type(node.ops[0])](a, b)
        if isinstance(node, ast.Call):
            args = node.args
            if node.func.id == 'field': return row.get(run(args[0]))
            if node.func.id == 'If':
                for i in range(0, len(args)-1, 2):
                    if run(args[i]): return run(args[i+1])
                return run(args[-1]) if len(args) % 2 else None
            if node.func.id == 'Switch':
                key = run(args[0])
                for i in range(1, len(args)-1, 2):
                    if run(args[i]) == key: return run(args[i+1])
                return run(args[-1])
            if node.func.id == 'IsNull': return run(args[0]) is None
            if node.func.id == 'Coalesce':
                for arg in args:
                    value = run(arg)
                    if value is not None: return value
                return None
            if node.func.id in ('Min', 'Max', 'Sum'):
                values = [run(arg) for arg in args]
                numbers = [value for value in values if isinstance(value, (int, float))]
                if node.func.id == 'Sum': return sum(numbers)
                if node.func.id == 'Min': return min(numbers, default=0)
                return max(numbers, default=0)
            if node.func.id == 'Greatest':
                values=[run(arg) for arg in args]
                return None if any(value is None for value in values) else max(values)
            if node.func.id == 'Date': return dt.date.fromisoformat(run(args[0]))
            if node.func.id == 'Abs':
                value=run(args[0]);return None if value is None else abs(value)
            if node.func.id == 'Round':
                number,places=(run(arg) for arg in args)
                return None if number is None else float(Decimal(str(number)).quantize(Decimal(1).scaleb(-places),rounding=ROUND_HALF_UP))
            if node.func.id == 'RoundUp':
                number, places = (run(arg) for arg in args)
                scale = 10**places
                return math.copysign(math.ceil(abs(number)*scale)/scale, number)
            if node.func.id == 'DateDiff':
                unit, a, b = (run(arg) for arg in args)
                assert unit == 'day'
                return (b-a).days
        raise AssertionError('unsupported local oracle expression: '+ast.dump(node))
    return run(ast.parse(code, mode='eval').body)


def project(columns, row):
    local, output = dict(row), {}
    for column in columns:
        value = scalar(column['formula'], local)
        output[column['name']] = local[column['name']] = value
    return output


def run_forecast(elements, wide_rows):
    by_id = {e['id']: e for e in elements}
    conn = sqlite3.connect(':memory:')
    def sql_rows(element):
        query = conn.execute(element['source']['statement'])
        names = [description[0] for description in query.description]
        return [project(element['columns'], {'Custom SQL/'+name: value for name, value in zip(names, row)})
                for row in query]
    try:
        actuals, spine = (sql_rows(by_id[key]) for key in ['actuals', 'spineWrap'])
    finally:
        conn.close()
    inputs = {row['Line Item']: row for row in wide_rows}
    forecast = []
    for row in spine:
        joined = {'FJ/'+key: value for key, value in row.items()}
        joined.update({'Inputs/'+key: value for key, value in inputs[row['Line Item']].items()})
        forecast.append(project(by_id['forecast']['columns'], joined))
    return [project(by_id['pnlCombined']['columns'], {'Union of 2 Sources/'+key: value
                     for key, value in row.items()}) for row in actuals+forecast]
