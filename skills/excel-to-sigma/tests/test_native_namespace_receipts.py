"""Omitted REST names need complete source-bound native metadata, not guesses."""
import copy
from pathlib import Path
import sys
import unittest
SCRIPTS=Path(__file__).resolve().parents[1]/'scripts'
sys.path[:0]=[str(SCRIPTS),str(SCRIPTS/'lib')]
from lib.native_exact_plan import digest
from lib.regional_grouped_model import _projection_ddl_columns
from lib.regional_grouped_phases import _validate_namespaces
import test_regional_inherited_columns as inherited
class NativeNamespaceReceiptTest(unittest.TestCase):
    def setUp(self):
        inherited.InheritedColumnReceiptTest.setUpClass();t=inherited.InheritedColumnReceiptTest();t.setUp();t.grouping_state()
        self.current=t.current;self.child=t.child
        ddl=t.receipt()['productDdl'];columns=_projection_ddl_columns(ddl,self.current,self.child)
        self.receipt={'documentId':self.current['workbookId'],'documentVersion':self.current['documentVersion'],
            'contentsHash':digest(self.current['contents']),'elementId':self.child['id'],'ddl':ddl,'columns':columns,
            'reviewedNameDifferences':[{'id':c['id'],'restName':r.get('name'),'nativeName':c['name']} for r,c in zip(self.child['columns'],columns) if r.get('name')!=c['name']]}
        self.binding={'elementId':self.child['id'],'name':self.child['name'],
            'columns':{c['name']:{'id':c['id'],'name':c['name']} for c in columns},'nativeColumnReceipt':self.receipt}
    def test_actual_table_source_without_grouping_id_uses_full_observed_metadata(self):
        self.assertNotIn('groupingId',self.child['source'])
        _validate_namespaces(self.current,{'Product Rows':self.binding})
        bare=copy.deepcopy(self.binding);bare.pop('nativeColumnReceipt')
        with self.assertRaisesRegex(ValueError,'BLOCKED'):_validate_namespaces(self.current,{'Product Rows':bare})
    def test_foreign_stale_incomplete_or_unreviewed_receipt_is_blocked(self):
        cases=[]
        for field,value in [('documentId','foreign'),('documentVersion',True),('contentsHash','0'*64),('elementId','foreign')]:
            r=copy.deepcopy(self.receipt);r[field]=value;cases.append(r)
        r=copy.deepcopy(self.receipt);r['columns'].reverse();cases.append(r)
        r=copy.deepcopy(self.receipt);r['columns'][0]['type']='number';cases.append(r)
        r=copy.deepcopy(self.receipt);r['columns'].pop();cases.append(r)
        r=copy.deepcopy(self.receipt);r['reviewedNameDifferences']=[];cases.append(r)
        r=copy.deepcopy(self.receipt);r['ddl']=r['ddl'].replace('Formula: [Regional Sales Rows/Group Key]','Formula: 0');cases.append(r)
        for i,r in enumerate(cases):
            b=copy.deepcopy(self.binding);b['nativeColumnReceipt']=r
            with self.subTest(i=i),self.assertRaisesRegex(ValueError,'BLOCKED'):_validate_namespaces(self.current,{'Product Rows':b})
    def test_binding_cannot_override_observed_id_label(self):
        b=copy.deepcopy(self.binding);b['columns']['Revenue Number']['name']='Revenue Text'
        with self.assertRaisesRegex(ValueError,'BLOCKED'):_validate_namespaces(self.current,{'Product Rows':b})
if __name__=='__main__':unittest.main()
