import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
class DecimalAdapterTests(unittest.TestCase):
    def module(self):
        import importlib.util
        self.assertIsNotNone(importlib.util.find_spec('decimal_procedure_adapter'),'explicit text transport adapter missing')
        import decimal_procedure_adapter
        return decimal_procedure_adapter
    def contract(self):
        p=self.module();c=p.controlled_decimal_text_contract('conn')
        c.update(signature='actual-sigma-indexed-text-signature',catalogObjectId='indexed-wrapper',lookupEvidenceId='lookup',callerEvidenceId='caller',scalarMappingEvidenceId='scalar',datetimeEvidenceId='date',provisioningEvidenceId='metadata',signatureEvidenceId='actual-ui-and-roundtrip')
        c['lookup']={k:c[k] for k in ('connectionId','path','signature','declaredTypes','actionTypes','returnType')}
        return p,c
    def test_reference_wrapper_body_is_exact_reviewed_fixed_point_delegation(self):
        import hashlib
        p=self.module();source=(Path(__file__).resolve().parent/'fixtures/task6-decimal-text-procedure.sql').read_text()
        body=source.split('$$')[1].strip()
        self.assertEqual(hashlib.sha256(body.encode()).hexdigest(),p.ADAPTER_BODY)
        self.assertEqual(body.count('CALL SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO('),1)
        self.assertNotIn('FLOAT',body);self.assertIn('TRY_TO_DECIMAL(P_AMOUNT_TEXT, 18, 2)',body)

    def test_original_numeric_contract_is_not_broadened_and_new_pending_signature_stops(self):
        p,c=self.contract()
        with self.assertRaises(ValueError):p.build_decimal_text_procedure_action(p.controlled_decimal_text_contract('conn'),['op','1234.56',7,True,'2026-10-15T14:30:00.000'])
        from custom_ports import controlled_procedure_contract,build_procedure_action
        with self.assertRaises(ValueError):build_procedure_action(c,['op','1234.56',7,True,'2026-10-15T14:30:00.000'])
        self.assertNotEqual(c['path'],controlled_procedure_contract('conn')['path'])
    def test_text_amount_preserves_large_exact_cents_and_typed_datetime_formula(self):
        p,c=self.contract();args=['op','100000000000000.10',7,True,'2026-10-15T14:30:00.000']
        result=p.build_decimal_text_procedure_action(c,args)
        self.assertIn('amount=100000000000000.10',result['invocation']['expectedScalar'])
        wire=result['effects'][0]['arguments'];self.assertEqual([a['type'] for a in wire],['text','text','integer','boolean','datetime'])
        self.assertEqual(wire[1]['value']['value']['value'],args[1]);self.assertEqual(wire[4]['value'],{'type':'formula','formula':'Datetime("2026-10-15T14:30:00.000")'})
        for amount in [1234.56,'1.001','-0.00','01','1e2',' 1','+1','1000000000000000','１２.３４']:
            with self.subTest(amount=amount),self.assertRaises(ValueError):p.build_decimal_text_procedure_action(c,[args[0],amount,*args[2:]])
        first=p.build_decimal_text_procedure_action(c,['same','1.0',7,True,args[4]])
        with self.assertRaises(ValueError):p.build_decimal_text_procedure_action(c,['same','1.00',7,True,args[4]],prior_invocation=first['invocation'])
if __name__=='__main__':unittest.main()
