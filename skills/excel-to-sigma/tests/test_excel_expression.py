"""Excel operator precedence and emitted Sigma expression preserve one AST."""
import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from excel_expression import Binary, ExcelCalcError, Group, Number, emit_sigma, evaluate, parse_expression


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


class ConcatenationPrecedenceTests(unittest.TestCase):
    def value(self, formula):
        return evaluate(parse_expression(formula, sheet='Model', workbook_mode=True), lambda _ref: None,
                        excel_errors=True)

    def test_arithmetic_binds_before_concatenation_and_comparison_follows(self):
        # Expected results follow Excel's documented operator precedence.
        for formula, expected in [('=1+2&"x"', '3x'),
                                  ('=1&2+3', '15'),
                                  ('=1+2&3+4', '37'),
                                  ('=2*3&"x"', '6x'),
                                  ('=1+2&"x"="3x"', True),
                                  ('=1+2&"x"="12x"', False),
                                  ('=1&(2+3)', '15'),
                                  ('=(1+2)&"x"', '3x')]:
            with self.subTest(formula=formula):
                self.assertEqual(self.value(formula), expected)

    def test_parentheses_can_force_invalid_text_arithmetic(self):
        with self.assertRaisesRegex(ExcelCalcError, '#VALUE!'):
            self.value('=1+(2&"x")')


class CriteriaEqualityTests(unittest.TestCase):
    def assert_criteria_results(self, values, cases):
        for argument, count, total in cases:
            for function, expected in [('COUNTIFS', count), ('SUMIFS', total)]:
                ranges='B1:B8,A1:A8' if function=='SUMIFS' else 'A1:A8'
                formula=f'={function}({ranges},{argument})'
                with self.subTest(formula=formula):
                    node=parse_expression(formula,sheet='Model',workbook_mode=True)
                    self.assertEqual(evaluate(node,lambda ref:values.get(ref.address),excel_errors=True),expected)

    def test_blank_reference_criterion_is_zero_and_empty_text_is_blank_matching(self):
        # Microsoft COUNTIFS Remarks: an empty criterion cell is treated as 0.
        # Keep empty strings separate: only A2/A3 match the literal empty text.
        values={'A1':0,'A2':None,'A3':'','A4':'0','A5':2,'A6':3,'A7':4,'A8':5,
                **{f'B{i}':2**(i-1) for i in range(1,9)},'C1':None,'C2':''}
        self.assert_criteria_results(values,[('C1',3,11),('0',3,11),('"0"',3,11),
                                            ('""',2,6),('C2',2,6),('"="&C1',2,6),('"<>"&C1',6,249),('"<>"',6,249)])

    def test_numeric_criteria_whitespace_preserves_text_operand_whitespace(self):
        values={'A1':2,'A2':'2','A3':'002','A4':' 2 ','A5':None,'A6':'','A7':1,'A8':' abc ',
                **{f'B{i}':2**(i-1) for i in range(1,9)},'C1':'>= 2','C2':' 2 ','C3':' abc '}
        self.assert_criteria_results(values,[('">= 2"',1,1),('C1',1,1),('" 2 "',4,15),
                                            ('C2',4,15),('">="&C2',1,1),('C3',1,128),('" abc "',1,128)])

    def test_nonfinite_numeric_spellings_remain_text_criteria(self):
        values={'A1':'inf','A2':'INF','A3':'NaN','A4':'-inf','A5':'Infinity','A6':2,'A7':None,'A8':'',
                **{f'B{i}':2**(i-1) for i in range(1,9)},'C1':'inf','C2':'NaN','C3':'-inf'}
        self.assert_criteria_results(values,[('"inf"',2,3),('C1',2,3),('"NaN"',1,4),('C2',1,4),
                                            ('"-inf"',1,8),('C3',1,8),('"Infinity"',1,16),('"<>NaN"',7,251)])

    def test_numeric_text_criteria_equality_and_inequality(self):
        # Independent literal expectations follow Microsoft's numeric-text
        # criteria comparison contract, not the native compiler's output.
        values={'A1':2,'A2':'2','A3':'002','A4':'abc','A5':None,'A6':'','A7':0,'A8':'0',
                'B1':10,'B2':20,'B3':40,'B4':80,'B5':160,'B6':320,'B7':640,'B8':1280,'C1':'2'}
        for argument,count,total in [('2',3,70),('"2"',3,70),('"002"',3,70),('C1',3,70),
                                     ('"="&C1',3,70),('"<>2"',5,2480),('"<>"&C1',5,2480),
                                     ('"abc"',1,80),('""',2,480),('"=0"',3,2080),
                                     ('">1"',1,10),('">"&C1',0,0)]:
            for function,expected in [('COUNTIFS',count),('SUMIFS',total)]:
                ranges='B1:B8,A1:A8' if function=='SUMIFS' else 'A1:A8'
                formula=f'={function}({ranges},{argument})'
                with self.subTest(formula=formula):
                    node=parse_expression(formula,sheet='Model',workbook_mode=True)
                    self.assertEqual(evaluate(node,lambda ref:values.get(ref.address),excel_errors=True),expected)

if __name__=='__main__':unittest.main()
