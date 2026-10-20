"""Read-only actual 13-source replay. Test row IDs never become target receipts."""
from __future__ import annotations
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from native_exact_plan import build_exact_native_calculation_plan,lower_exact_native_calculations


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',type=Path,required=True)
    parser.add_argument('--input-plans',type=Path,required=True)
    parser.add_argument('--out',type=Path,required=True)
    args=parser.parse_args();args.out.mkdir(parents=True,exist_ok=True)
    coverage={};totals={'originalFive':Counter(),'broaderEight':Counter()};requests={}
    for file in sorted(args.input_plans.glob('*.json')):
        inputs=json.loads(file.read_text());name=inputs['sourceName'];group='originalFive' if name.startswith('fixture-') else 'broaderEight'
        source=args.root/'outputs'/name if group=='originalFive' else args.root/'work/broader-eight-source-replay-20261020/source-originals'/name
        if hashlib.sha256(source.read_bytes()).hexdigest()!=inputs['sourceSha256']:raise ValueError('source/input hash changed')
        plan=build_exact_native_calculation_plan(source,inputs)
        assigned={feature:{'elementId':'TEST_ONLY_'+feature,'name':spec['name'],'rowIdColumn':'Row ID',
            'rows':{str(index):f'00000000-0000-0000-0000-{index:012d}' for index in range(len(spec['rows']))}}
            for feature,spec in inputs['sharedInputArchitecture']['inputs'].items()}
        lower=lower_exact_native_calculations(plan,assigned)
        expected=sum(r['status']=='ready-to-bind'for r in plan['formulas'].values())
        if len(lower['formulas'])!=expected:raise ValueError('candidate construction denominator changed')
        if any(len(e['columns'])>64 for e in lower['shards']):raise ValueError('native column limit exceeded')
        if any('__native_'in c['formula'] for e in lower['shards'] for c in e['columns']):raise ValueError('helper token escaped construction')
        counts=Counter(r['status']for r in plan['formulas'].values());totals[group].update(counts)
        status={'sourcePath':str(source),'sourceSha256':plan['sourceSha256'],'sourceGroup':group,'formulaCount':plan['formulaCount'],
            'nativePlanHash':plan['nativePlanHash'],'inputPlanHash':inputs['inputPlanHash'],'statusCounts':dict(counts),
            'unresolvedReasons':dict(Counter(r['reason']for r in plan['formulas'].values()if r['status']=='unresolved')),
            'unresolvedFormulaCells':{c:r['reason']for c,r in plan['formulas'].items()if r['status']=='unresolved'},
            'sourceErrorClasses':{c:r['localExpectedError']for c,r in plan['formulas'].items()if r.get('sourceResultClass')=='excel-error'},
            'loweredFormulaCount':len(lower['formulas']),'shardCount':len(lower['shards']),'maxFormulaLength':lower['maxFormulaLength'],
            'maxDependencyDepth':lower['maxDependencyDepth'],'rangeAggregateCount':lower['rangeAggregateCount'],
            'ordinaryRowTableCount':lower['ordinaryRowTableCount'],'helperCount':lower['helperCount'],
            'formalTemplateObligations':lower['formalTemplateObligations'],
            'appendErrorOrder':'new rows with calculation errors require reviewed stable source ordinal and runtime qualification',
            'constructionReplayIds':'TEST ONLY fixed-length UUID-shaped row IDs; no target receipts or assigned target IDs',
            'runtimeStatus':'UNVERIFIED_CANDIDATE','planPath':str(args.out/file.name)}
        coverage[name]=status;(args.out/file.name).write_text(json.dumps(plan,indent=2)+'\n')
        if name.startswith(('02 SaaS','06 13-Week')):
            requests[name]={'status':'ROOT_STAGING_REQUEST_PENDING_REAL_RECEIPTS','sourcePath':str(source),'sourceSha256':plan['sourceSha256'],
                'inputPlanPath':str(file),'inputPlanHash':inputs['inputPlanHash'],'nativePlanPath':str(args.out/file.name),
                'nativePlanHash':plan['nativePlanHash'],'candidateFormulaCount':expected,'runtimeStatus':'UNVERIFIED_CANDIDATE',
                'requiredReceipts':['actual staging target connection/folder and source-reviewed input plan','queried assigned input element/column/system row IDs with complete typed physical rows','native --lower-candidate using those observed rows','named native review and matching manifest logical mappings generated from actual lowering','approved assigned staging receipt','fresh full document/version and persisted complete document/row backup','observed SELECT 1 AS C0 singleton seed element, queriedRows [{C0:1}], and actual existing native page ID','one prepare-scaffold MCP program + confirm-scaffold assigned-ID receipt per helper table','prepare-update parenthesized REST packet + confirm-readback per table','native value/type/error parity, original five Audit verification obligations, edit/reopen/restore evidence'],
                'cli':'scripts/native-calculations.py plan --exact; lower-candidate; prepare-scaffold; confirm-scaffold; prepare-update --native-bindings; confirm-readback'}
        print(name,dict(counts),'lowered',len(lower['formulas']),'max',lower['maxFormulaLength'])
    if sum(totals['originalFive'].values())!=1865 or sum(totals['broaderEight'].values())!=1786:raise ValueError('fixed 1865/1786 source denominators changed')
    report={'formulaCount':3651,'denominators':{'originalFive':1865,'broaderEight':1786},'totals':{k:dict(v)for k,v in totals.items()},'sources':coverage}
    (args.out/'coverage.json').write_text(json.dumps(report,indent=2)+'\n')
    (args.out/'root-construction-requests.json').write_text(json.dumps(requests,indent=2)+'\n')

if __name__=='__main__':main()
