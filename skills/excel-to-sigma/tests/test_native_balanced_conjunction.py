"""Boolean grouping must not create the deep native-condition failure."""
import ast
from pathlib import Path
import sys
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))
import native_actions


class NativeBalancedConjunctionTests(unittest.TestCase):
    def test_large_conjunction_has_logarithmic_depth(self):
        expression = ast.parse(native_actions._conjoin(['guard_'+str(i) for i in range(128)]), mode='eval').body
        def depth(node):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == 'and_':
                return 1 + max(depth(argument) for argument in node.args)
            return 0
        self.assertLessEqual(depth(expression), 8)

    def test_order_duplicates_and_nonconjunctive_subexpressions_are_preserved(self):
        values = ['eq(lookup(value, key, row_id), const_num(1))',
                  'or_(is_null(before), eq(before, actual))',
                  'eq(subtract(actual, actual), const_num(0))'] * 5
        original = list(values)
        expression = ast.parse(native_actions._conjoin(values), mode='eval').body
        def leaves(node):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == 'and_':
                return [leaf for argument in node.args for leaf in leaves(argument)]
            return [ast.dump(node)]
        self.assertEqual(leaves(expression), [ast.dump(ast.parse(v, mode='eval').body) for v in values])
        self.assertEqual(values, original)

    def test_empty_and_single_operand_keep_existing_contract(self):
        self.assertEqual(native_actions._conjoin([]), 'const_bool(True)')
        value = 'or_(is_null(before), eq(before, actual))'
        self.assertEqual(native_actions._conjoin([value]), value)


if __name__ == '__main__':
    unittest.main()
