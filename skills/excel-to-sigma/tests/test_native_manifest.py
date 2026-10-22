"""Opt-in native bound evidence never bypasses source or runtime gates."""
import copy
import hashlib
import tempfile
import unittest
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from openpyxl import Workbook
from native_exact_plan import build_exact_native_calculation_plan
from migration_manifest import build_manifest,validate_manifest,ManifestError

class BoundNativeManifestTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.path=Path(self.tmp.name)/'renamed.xlsx';b=Workbook();b.active['A1']='=IFERROR(1/0,9)';b.save(self.path)
        self.native=build_exact_native_calculation_plan(self.path)
        self.plan={'nativeCalculationPlan':self.native,'nativeReviewApproval':{'status':'reviewed','owner':'owner','reviewer':'reviewer','sourceSha256':self.native['sourceSha256'],'nativePlanHash':self.native['nativePlanHash']},'formulaCells':[{'id':'Sheet!A1','mappingStatus':'live','targetMapping':{'logicalElementId':'native1','logicalColumnId':'result'}}]}
    def test_default_parser_stays_conservative_and_opt_in_preserves_unverified_status(self):
        self.assertEqual(build_manifest(self.path,{'formulaCells':self.plan['formulaCells']})['formulaCells'][0]['mappingStatus'],'unresolved')
        m=build_manifest(self.path,self.plan)
        self.assertEqual(m['formulaCells'][0]['mappingStatus'],'live')
        self.assertEqual(m['formulaCells'][0]['nativeCandidate']['runtimeStatus'],'UNVERIFIED_CANDIDATE')
        self.assertIsNone(m['formulaCells'][0]['verification'])
        validate_manifest(m)
    def test_source_formula_plan_and_review_hashes_cannot_be_forged(self):
        for mutation in ('ast','review','source'):
            p=copy.deepcopy(self.plan)
            if mutation=='ast':p['nativeCalculationPlan']['formulas']['Sheet!A1']['ast']['name']='SUM'
            if mutation=='review':p['nativeReviewApproval']['nativePlanHash']='0'*64
            if mutation=='source':p['nativeCalculationPlan']['sourceSha256']='0'*64
            with self.subTest(mutation=mutation),self.assertRaises(ManifestError):build_manifest(self.path,p)

if __name__=='__main__':unittest.main()
