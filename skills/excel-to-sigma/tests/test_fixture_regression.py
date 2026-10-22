import sys,unittest,json
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import fixture_regression as oracle
FIX=Path(__file__).resolve().parent/'fixtures'
class FixtureRegressionTests(unittest.TestCase):
    def path(self,name):return FIX/('fixture-'+name+'.xlsx')
    def test_sales_cycle_days_oracle_retains_source_timestamp_fraction(self):
        import datetime as dt
        for created,close,expected in [
            ('2026-01-01T08:00:00','2026-01-02T20:00:00',1.5),
            ('2026-01-02T05:00:00','2026-01-01T23:00:00',-.25),
            ('2026-01-01T08:00:00','2026-01-01T08:00:00.001',.000000011574074074074073),
            ('2026-01-01T08:00:00.001','2026-01-01T08:00:00',-.000000011574074074074073),
        ]:
            with self.subTest(created=created,close=close):
                edits={'Pipeline!B2':dt.datetime.fromisoformat(created),
                       'Pipeline!C2':dt.datetime.fromisoformat(close)}
                actual=oracle.sales(self.path('sales-pipeline'),edits)['rows'][0]['Cycle Days']
                self.assertAlmostEqual(actual,expected,delta=1e-14)
    def test_independent_primary_and_secondary_denominators_and_edits(self):
        forecast=oracle.forecast(self.path('forecast-model'));self.assertEqual(len(forecast['monthlyValues']),144)
        self.assertEqual(len(forecast['rows'])*13,234)
        sales=oracle.sales(self.path('sales-pipeline'));self.assertEqual(len(sales['rows']),48);self.assertEqual(len(sales['repTargets']),4)
        changed=oracle.sales(self.path('sales-pipeline'),{'Rep Targets!B2':2500000})
        self.assertAlmostEqual(changed['repTargets'][0]['Coverage'],sales['repTargets'][0]['Coverage']/2)
        inventory=oracle.inventory(self.path('inventory-reorder'));self.assertEqual(len(inventory['rows']),36);self.assertEqual(len(inventory['purchasePlan'])*7,252)
        expense=oracle.expense(self.path('expense-approval'));self.assertEqual(len(expense['rows']),48);self.assertEqual(len(expense['reconciliation'])*9,432)
        self.assertEqual(oracle.expense(self.path('expense-approval'),{'Expense Register!J3':200.26})['rows'][1]['Exception'],'CLEAR')
        self.assertEqual(oracle.expense(self.path('expense-approval'),{'Expense Register!J3':200.27})['rows'][1]['Exception'],'RECON')
    def test_research_hand_equations_preserve_gp_override_equity_and_growth_dependencies(self):
        path=self.path('research-model');baseline=oracle.research(path)['cells'];changed=oracle.research(path,{'Model Inputs!C2':.1})['cells']
        golden=json.loads((FIX/'fixture-research-expected.json').read_text())
        # Golden rows39/44 use rounded generator EPS; source row37 is a live formula.
        for row,values in golden['rows'].items():
            if row in {'39','44'}:continue
            for col,value in zip('BCDEFGHIJKLM',values):
                actual=baseline['FY results!'+col+row]
                if value is None:self.assertIsNone(actual)
                else:self.assertAlmostEqual(actual,value,places=5)
        source=oracle.FixtureOracle(path)
        try:
            for col in 'BCDEFGHIJKLM':
                self.assertEqual(source.workbook['FY results'][col+'44'].value,'='+col+'43/'+col+'37')
                self.assertEqual(source.workbook['FY results'][col+'39'].value,'='+col+'38/'+col+'37')
                self.assertEqual(baseline['FY results!'+col+'44'],source.literal('FY results',col+'43')/baseline['FY results!'+col+'37'])
                self.assertEqual(baseline['FY results!'+col+'39'],source.literal('FY results',col+'38')/baseline['FY results!'+col+'37'])
        finally:source.close()
        self.assertEqual(len(baseline),420);self.assertNotEqual(baseline['FY results!H8'],changed['FY results!H8'])
        self.assertEqual(baseline['FY results!H13'],changed['FY results!H13']);self.assertEqual(baseline['FY results!H31'],changed['FY results!H31'])
        equity=oracle.research(path,{'FY results!H56':999})['cells'];self.assertEqual(equity['FY results!H56'],999)
if __name__=='__main__':unittest.main()
