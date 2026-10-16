"""Excel operator precedence and emitted Sigma expression preserve one AST."""
import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from excel_expression import Binary, Group, Number, emit_sigma, evaluate, parse_expression


class PowerPrecedenceTests(unittest.TestCase):
    def expression(self,formula):return parse_expression(formula,sheet='Model')
    def value(self,formula):return evaluate(self.expression(formula),lambda _ref:None)

    def test_chained_powers_follow_excel_left_to_right(self):
        self.assertEqual(self.value('=2^3^2'),64)
        self.assertEqual(self.value('=2^3^2+1'),65)
        node=self.expression('=2^3^2')
        self.assertIsInstance(node,Binary)
        self.assertEqual(node.operator,'^')
        self.assertIsInstance(node.left,Binary)
        self.assertEqual(node.left.operator,'^')
        self.assertEqual(node.right,Number(2,'2'))

    def test_unary_sign_and_percent_precedence_are_preserved_in_power_chains(self):
        for formula,expected in [('=-2^3^2',64),('=+2^3^2',64),('=2^-3^2',.015625),
                                 ('=50%^2^3',.015625),('=200%^3^2',64),('=-(2^3^2)',-64),
                                 ('=-2^2',4),('=2^3*2',16)]:
            with self.subTest(formula=formula):self.assertEqual(self.value(formula),expected)

    def test_explicit_right_grouping_stays_right_grouped(self):
        self.assertEqual(self.value('=2^(3^2)'),512)
        node=self.expression('=2^(3^2)')
        self.assertIsInstance(node.right,Group)
        self.assertIsInstance(node.right.expression,Binary)
        self.assertEqual(self.value('=(2^3)^2'),64)

    def test_sigma_power_calls_preserve_the_parsed_left_or_explicit_right_tree(self):
        for formula,expected in [('=2^3^2','Power(Power(2, 3), 2)'),
                                 ('=2^3^2+1','Power(Power(2, 3), 2) + 1'),
                                 ('=2^(3^2)','Power(2, (Power(3, 2)))'),
                                 ('=A1^B1^C1','Power(Power([A1], [B1]), [C1])'),
                                 ('=50%^2^3','Power(Power((50 * 0.01), 2), 3)'),
                                 ('=-2^3^2','Power(Power(-2, 3), 2)')]:
            with self.subTest(formula=formula):
                self.assertEqual(emit_sigma(self.expression(formula),lambda ref:'['+ref.address+']'),expected)

if __name__=='__main__':unittest.main()
