"""Hand-calculated cases for the opt-in exact native compiler, independent of its emitter."""
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/lib'))
from native_exact import compile_expression as compile_expression_raw, Literal, evaluate_bound
from excel_expression import ExcelCalcError, UnsupportedExpression
from native_calculations import _ast
from workbook_expression import WorkbookContext
from openpyxl import Workbook
from native_formula_reference import evaluate_native as evaluate_native_raw

_native_helpers={}
def evaluate_native(formula):return evaluate_native_raw(formula,_native_helpers)
def compile_expression(node,reference):
    _native_helpers.clear()
    def emit(formula):
        token=f"[test_helper_{len(_native_helpers)}]"
        _native_helpers[token]=evaluate_native_raw(formula,_native_helpers)
        return token
    return compile_expression_raw(node,reference,emit=emit)


class ExactFunctionTests(unittest.TestCase):
    def bound(self, formula, values=()):
        book=Workbook(); ws=book.active; ws.title='Renamed'
        for cell, value in dict(values).items(): ws[cell]=value
        return _ast(WorkbookContext("memory.xlsx",book,"test").bind_formula(formula, sheet='Renamed', address='Z99'))

    def value(self, formula, values=()):
        node=self.bound(formula, values)
        return evaluate_bound(node, lambda c:dict(values).get(c.split('!')[1]))

    def test_criteria_and_numeric_range_ignore_text_and_booleans(self):
        values={'A1':'West','A2':'west','A3':'East','A4':None,'B1':4,'B2':7,'B3':8,'B4':9}
        self.assertEqual(self.value('=SUMIFS(B1:B4,A1:A4,"WEST")',values),11)
        self.assertEqual(self.value('=COUNTIFS(A1:A4,"w*")',values),2)
        self.assertEqual(self.value('=COUNTIFS(A1:A4,"")',values),1)
        self.assertEqual(self.value('=COUNT(A1:B4)',values),4)
        self.assertEqual(self.value('=COUNTA(A1:B4)',values),7)
        self.assertEqual(self.value('=SUM(A1:A4)',values),0)
        self.assertEqual(self.value('=SUM(TRUE,"2",3)'),6)

    def test_lookup_first_duplicate_case_and_missing_errors(self):
        v={'A1':'x','A2':'X','A3':'z','B1':4,'B2':8,'B3':9}
        self.assertEqual(self.value('=VLOOKUP("X",A1:B3,2,FALSE)',v),4)
        self.assertEqual(self.value('=XLOOKUP("z",A1:A3,B1:B3)',v),9)
        self.assertEqual(self.value('=INDEX(B1:B3,MATCH("x",A1:A3,0))',v),4)
        self.assertEqual(self.value('=IFERROR(XLOOKUP("missing",A1:A3,B1:B3),17)',v),17)
        with self.assertRaisesRegex(ExcelCalcError,'#N/A'): self.value('=MATCH("missing",A1:A3,0)',v)

    def test_error_catching_is_distinct_from_unsupported_binding(self):
        self.assertEqual(self.value('=IFERROR(1/0,99)'),99)
        self.assertEqual(self.value('=IF(FALSE,1/0,8)'),8)
        with self.assertRaises(UnsupportedExpression):
            compile_expression({'kind':'function','name':'IFERROR','arguments':[{'kind':'reference','cell':'Missing!A1'},{'kind':'number','source':'9'}]},lambda c: (_ for _ in ()).throw(UnsupportedExpression('unresolved')))
        result=compile_expression(self.bound('=IFERROR(1/0,99)'),lambda c:Literal.number(0))
        self.assertEqual(evaluate_native(result.number),99)
        self.assertNotIn('Coalesce(Null',result.number)
        self.assertEqual(evaluate_native(result.error),'')

    def test_logic_mod_rounding_and_dates(self):
        for formula,expected in [('=AND(TRUE,OR(FALSE,TRUE))',True),('=MOD(-3,2)',1),('=MOD(3,-2)',-1),('=ROUNDUP(-12.21,1)',-12.3),('=CEILING(7,3)',9),('=CEILING(-7,3)',-6),('=MONTH(DATE(2024,2,29))',2),('=EDATE(DATE(2024,1,31),1)',45351),('=EOMONTH(DATE(2024,2,1),0)',45351),('=DATE(2024,13,0)',45657)]:
            with self.subTest(formula=formula): self.assertEqual(self.value(formula),expected)

    def test_index_row_array_and_scalar_blank(self):
        v={'A1':1,'B1':2,'A2':3,'B2':4}
        self.assertEqual(self.value('=SUM(INDEX(A1:B2,2,0))',v),7)
        self.assertEqual(self.value('=INDEX(A1:B2,1,2)',v),2)
        self.assertEqual(self.value('=A3+4',v),4)
        self.assertEqual(self.value('=COUNTA(A3)',v),0)
        self.assertEqual(self.value('=COUNTA("")'),1)

    def test_emitted_lanes_match_hand_results(self):
        cases=[('=IFERROR(1/0,99)',{},99),('=IF(FALSE,1/0,8)',{},8),
               ('=SUM(TRUE,"2",3)',{},6),('=COUNT(A1:A3)',{'A1':4,'A2':'4','A3':True},1),
               ('=COUNTA(A1:A3)',{'A1':4,'A2':'','A3':None},2),
               ('=MOD(-3,2)',{},1),('=ROUNDUP(-12.21,1)',{},-12.3),
               ('=CEILING(-7,3)',{},-6),('=MONTH(DATE(2024,2,29))',{},2),
               ('=EDATE(DATE(2024,1,31),1)',{},45351),('=EOMONTH(DATE(2024,2,1),0)',{},45351),
               ('=DATE(2024,13,0)',{},45657),
               ('=VLOOKUP("x",A1:B2,2,FALSE)',{'A1':'X','A2':'x','B1':4,'B2':8},4),
               ('=INDEX(A1:B2,2,2)',{'A1':1,'A2':3,'B1':2,'B2':4},4),
               ('=SUM(INDEX(A1:B2,2,0))',{'A1':1,'A2':3,'B1':2,'B2':4},7),
               ('=IFERROR(MATCH("missing",A1:A2,0),17)',{'A1':'x','A2':'y'},17),
               ('=IF(TRUE,"",1)',{},''),('=COUNT(TRUE,"2","ignored")',{},2),('=COUNT(INDEX(A1:B2,2,0))',{'A1':1,'A2':3,'B1':2,'B2':4},2),('=EDATE(60,1)',{},89),('=EDATE(60,0)',{},60),('=MONTH(0)',{},1),('=MATCH("x*",A1:A2,0)',{'A1':'XYZ','A2':'x'},1),('=IFERROR(XLOOKUP("x",A1:A2,B1:B2,1/0),17)',{'A1':'x','A2':'y','B1':4,'B2':8},4)]
        for formula,values,expected in cases:
            with self.subTest(formula=formula):
                value=compile_expression(self.bound(formula,values),lambda c:Literal.value(values.get(c.split('!')[1])))
                self.assertEqual(evaluate_native(value.error),'')
                kind=evaluate_native(value.kind)
                result=evaluate_native({'n':value.number,'t':value.text,'b':value.logical,'z':'Null'}[kind])
                self.assertEqual(result,expected)

    def test_dynamic_editable_wildcard_and_operator_criteria(self):
        for criterion,expected in [('w*',11),('w?st',11),('~*',8),('~~*',9),('>5',24),('',0),('<>',32)]:
            values={'A1':'West','A2':'west','A3':'*','A4':'~abc','B1':4,'B2':7,'B3':8,'B4':9,'C1':criterion}
            value=compile_expression(self.bound('=SUMIFS(B1:B4,A1:A4,C1)',values),lambda c:Literal.value(values.get(c.split('!')[1])))
            if criterion=='>5': expected=0
            if criterion=='<>':expected=28
            with self.subTest(criterion=criterion):self.assertEqual(evaluate_native(value.number),expected)

    def test_sigma_roundup_oracle_is_ceiling_and_excel_lowering_is_away_from_zero(self):
        # Sigma official RoundUp examples are the native interpreter contract;
        # Excel ROUNDUP is a separate hand-derived expected result.
        self.assertEqual(evaluate_native_raw('RoundUp(-3.141, 2)'),-3.14)
        self.assertEqual(evaluate_native_raw('RoundUp(3.141, 2)'),3.15)
        self.assertEqual(evaluate_native_raw('RoundDown(-6.25417, 3)'),-6.255)
        self.assertEqual(evaluate_native_raw('Trunc(-1.9, 0)'),-1)
        for number,digits,expected in [(-3.141,2,-3.15),(3.141,2,3.15),(-12.21,1,-12.3),(-123.1,-1,-130),(-123.1,-1.9,-130),(123.1,-1,130),(0,2,0)]:
            with self.subTest(number=number,digits=digits):
                result=compile_expression(self.bound(f'=ROUNDUP({number},{digits})'),lambda c:Literal.number(0))
                self.assertEqual(evaluate_native(result.number),expected)
                self.assertEqual(evaluate_native(result.error),'')

    def test_boolean_criteria_are_unresolved_until_actual_excel_oracle(self):
        values={'A1':None,'A2':'x','A3':True,'A4':0,'C1':'>0'}
        node=self.bound('=COUNTIFS(A1:A4,C1)',values)
        with self.assertRaisesRegex(UnsupportedExpression,'Boolean criteria.*Excel oracle'):
            evaluate_bound(node,lambda c:values.get(c.split('!')[1]))
        with self.assertRaisesRegex(UnsupportedExpression,'Boolean criteria.*Excel oracle'):
            compile_expression_raw(node,lambda c:Literal.value(values.get(c.split('!')[1])),emit=lambda f:'[helper]')
        # A numeric baseline cannot prove a formula's future Boolean branch.
        child=compile_expression_raw(self.bound('=IF(C1=7,TRUE,0)'),lambda c:Literal.number('[editable]'),emit=lambda f:'[helper]')
        parent=self.bound('=COUNTIFS(A1:A1,">0")')
        with self.assertRaisesRegex(UnsupportedExpression,'Boolean criteria.*Excel oracle'):
            compile_expression_raw(parent,lambda c:child,emit=lambda f:'[helper]')

    def test_finite_indirect_invalid_selector_is_ref_error_and_not_blank(self):
        for selector,result,error in [('Base',4,''),('Missing',0,'#REF!')]:
            node={'kind':'indirect','selector':{'kind':'text','value':selector},'alternatives':[['Base',{'kind':'number','source':'4'}]]}
            value=compile_expression(node,lambda c:Literal.number(0))
            self.assertEqual(evaluate_native(value.error),error)
            if not error:self.assertEqual(evaluate_native(value.number),result)

    def test_compiler_has_explicit_error_channel_and_bounded_helpers(self):
        v={'A1':Literal.number(3),'A2':Literal.number(4)}
        for formula in ['=SUMIFS(A1:A2,A1:A2,">2")','=IFERROR(1/0,9)','=VLOOKUP(3,A1:A2,1,FALSE)','=INDEX(A1:A2,MATCH(4,A1:A2,0))','=DATE(2024,2,29)','=ROUNDUP(-4.1,0)','=AND(TRUE,FALSE)']:
            with self.subTest(formula=formula):
                result=compile_expression(self.bound(formula),lambda c:v[c.split('!')[1]])
                self.assertIsInstance(result.error,str)
                self.assertIsInstance(result.number,str)

if __name__=='__main__': unittest.main()
