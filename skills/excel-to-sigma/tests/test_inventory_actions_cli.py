"""The offline CLI must emit a reviewable plan and never call live services."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ORACLE = json.loads((ROOT / 'tests/fixtures/inventory-source-oracle.json').read_text())
CLI = ROOT / 'scripts/inventory-actions.py'


class InventoryActionsCliTests(unittest.TestCase):
    def test_plan_adjust_writes_blocked_exact_id_plan(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            capture = {'origin':'sigma-query','executionId':'synthetic-before','evidenceId':'synthetic-evidence',
                'sourceSha256':ORACLE['sourceSha256'],'documentId':'synthetic-doc','version':'v1',
                'binding':{'elementId':'demand','rowIdColumn':'Input Row ID',
                    'columnIds':{'Input Row ID':'id','Source Row':'ord','SKU':'sku',
                        'Week Start':'week','Forecast Units':'units','Planned Receipts':'receipts'}},
                'rows':[{**row,'Input Row ID':f'id-{i}'} for i,row in enumerate(ORACLE['demand'],1)]}
            before=folder/'capture.json';out=folder/'plan.json'
            before.write_text(json.dumps(capture))
            result=subprocess.run([sys.executable,str(CLI),'plan-adjust','--capture',str(before),
                '--operation','intent-A','--percent','100','--out',str(out)],text=True,capture_output=True)
            self.assertEqual(result.returncode,0,result.stderr)
            plan=json.loads(out.read_text())
            self.assertEqual(len(plan['updates']),72)
            self.assertEqual(plan['productionStatus'],'BLOCKED_PENDING_LIVE_EVIDENCE')
            self.assertEqual(plan['updates'][0]['rowId'],'id-1')


if __name__ == '__main__':
    unittest.main()
