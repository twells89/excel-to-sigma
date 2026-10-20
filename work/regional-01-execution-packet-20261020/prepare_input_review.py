"""Offline Regional input-only review material; no approval or target write."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

parser=argparse.ArgumentParser(description=__doc__)
for name in ('repo','task-root','connection','folder','current','out'):
    parser.add_argument('--'+name,required=True)
args=parser.parse_args()
sys.path.insert(0,str(Path(args.repo)/'skills/excel-to-sigma/scripts/lib'))
from source_inputs import build_source_input_plan,validate_source_input_plan
from source_workbook import build_source_model
from native_exact_plan import build_exact_native_calculation_plan
from migration_manifest import build_manifest
from migration_gate import gate_plan_staging_source
root=Path(args.task_root);out=Path(args.out);out.mkdir(parents=True,exist_ok=True)
source=root/'work/broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx'
reviewed=json.loads((root/'work/complete-source-record-bindings-plans-20261020/01-regional-sales-forecast-xlsx.json').read_text())
validate_source_input_plan(reviewed,source)
if hashlib.sha256(source.read_bytes()).hexdigest()!='4b46bfa6190d1ef12ea35713e3932aa77f19fb964730b4b3bbb0851294219e2c':
    raise ValueError('Regional source identity changed')
if any(not v.strip() or 'placeholder' in v.lower() for v in (args.connection,args.folder)):
    raise ValueError('actual approved staging connection/folder are required')
inputs=build_source_input_plan(source,reviewed['decisions'],connection_id=args.connection,folder_id=args.folder)
plan={'sourceSha256':inputs['sourceSha256'],'sourcePath':str(source),'inputPlanHash':inputs['inputPlanHash'],
      'sharedInputArchitecture':inputs['sharedInputArchitecture'],'executionScope':'input-tables-only'}
manifest=build_manifest(source,plan)
current=json.loads(Path(args.current).read_text())
if not current.get('workbookId') or type(current.get('documentVersion')) is not int or not current.get('contents',{}).get('pages'):
    raise ValueError('fresh complete staging document/page readback is required')
manifest['target']={'documentId':current['workbookId'],'version':current['documentVersion']}
_,source_gate=gate_plan_staging_source(plan,source_arg=str(source))
request={'status':'REQUEST_PENDING_NAMED_ROOT_REVIEW','scope':'input-tables-only',
         'owner':None,'reviewer':None,'sourceHash':manifest['sourceHash'],'planHash':manifest['planHash'],
         'mappedFeatures':inputs['sharedInputArchitecture']['mappedFeatures'],
         'reviewItems':[r['id']for r in manifest['formulaCells']],
         'pendingFormulaIds':[r['id']for r in manifest['formulaCells']],
         'targetBindings':{f:{**manifest['target'],'pageId':None}for f in inputs['sharedInputArchitecture']['inputs']},
         'auditedIssuesRequiringExplicitStagingReview':source_gate.get('issues',[])}
for name,value in [('input-plan',inputs),('native-plan',build_exact_native_calculation_plan(source,inputs)),
                   ('source-model',build_source_model(source)),('input-review-plan',plan),
                   ('input-manifest',manifest),('source-gate',source_gate),('input-approval-request',request)]:
    (out/(name+'.json')).write_text(json.dumps(value,indent=2,ensure_ascii=False)+'\n')
print('Prepared 214 pending formula obligations and five literal input tables; no staging approval or execution.')
