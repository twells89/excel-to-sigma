"""Approved SDK display plans; code-rep input-column exports are not evidence."""
from __future__ import annotations

import copy
import hashlib
import json
import re

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

try:
    from .input_sync import _approved_feature
except ImportError:
    from input_sync import _approved_feature


def _hash(value):
    return hashlib.sha256(json.dumps(value,sort_keys=True,separators=(',',':')).encode()).hexdigest()


def source_format_intent(raw, column_type):
    """Translate bounded Excel semantics to documented raw d3/strftime formats."""
    if column_type in {'text','boolean'} or raw in {'General','@'}:
        return {'kind':'automatic'},'EXACT',None
    if column_type=='datetime':
        pattern={'yyyy-mm-dd':'%Y-%m-%d','yyyy-mm-dd hh:mm:ss':'%Y-%m-%d %H:%M:%S'}.get(raw)
        return (({'pattern':pattern},'EXACT',None) if pattern else
                (None,'UNRESOLVED','source date pattern has no verified exact SDK mapping'))
    sections=raw.split(';')
    positive=sections[0]
    if not re.fullmatch(r'\$?(?:#,##)?0(?:\.0+)?%?',positive):
        return None,'UNRESOLVED','source numeric pattern has no verified exact SDK mapping'
    decimals=len(positive.split('.')[1].rstrip('%')) if '.' in positive else 0
    negative='(' if len(sections)>1 and '(' in sections[1] else ''
    pattern=negative+('$' if '$' in positive else '')+(',' if ',' in positive else '')+f'.{decimals}'+('%' if positive.endswith('%') else 'f')
    if len(sections)>1:
        return {'pattern':pattern},'PARTIAL','Excel accounting sections and zero-as-dash fidelity are unresolved; only supported currency/grouping/precision/negative-parentheses semantics are configured'
    return {'pattern':pattern},'EXACT',None


def currency_pattern(raw, precision):
    """Use a real d3 currency marker; currencySymbol metadata alone does not render it."""
    intent,_,_=source_format_intent(raw,'number')
    if intent and 'pattern' in intent and '$' in intent['pattern']:
        return intent['pattern']
    return f'$,.{precision}f'


def build_input_display(plan,manifest,receipt,feature=None):
    if manifest.get('sourceHash')!=plan['source']['sha256']:
        raise ValueError('input display source hash is stale')
    architecture=manifest.get('planProjection',{}).get('sharedInputArchitecture',{})
    feature=feature or architecture.get('presentationInputFeature')
    spec,target,artifact=_approved_feature(manifest,receipt,feature,require_assigned=True)
    if manifest.get('target')!={'documentId':target['documentId'],'version':target['version']}:
        raise ValueError('input display target/version binding is stale')
    workbook=load_workbook(plan['source']['path'],data_only=False)
    try:
        records=[]
        prominent=['Scope']+[f'i_model_inputs_c{i}' for i in range(2,8)] if plan['workflow']=='research' else [c['name'] for c in spec['columns']]
        for column in spec['columns']:
            name=column['name']; source_cell=None; raw=None
            provenance=column.get('provenance','')
            if '!' in provenance and not provenance.startswith('source:'):
                source_cell=provenance
                sheet,address=source_cell.rsplit('!',1)
                raw=workbook[sheet][address].number_format
            elif column.get('sourceColumn'):
                sheet,letter=column['sourceColumn'].rsplit('!',1)
                _,first,_,last=range_boundaries(spec['sourceRange'])
                source_cell=f'{sheet}!{letter}{first+1}'
                formats={workbook[sheet][f'{letter}{row}'].number_format for row in range(first+1,last+1)}
                raw=next(iter(formats)) if len(formats)==1 else None
            if raw is None and column['type'] in {'text','boolean'}:
                raw='General'
            if raw is None:
                fmt,fidelity,reason=None,'UNRESOLVED','source column has missing or heterogeneous format bindings'
            else:
                fmt,fidelity,reason=source_format_intent(raw,column['type'])
            records.append({'id':target['columnIds'][name],'name':name,'visible':name in prominent,
                            'sourceCell':source_cell,'sourceFormat':raw,'format':fmt,
                            'fidelity':fidelity,'unresolvedReason':reason})
    finally:
        workbook.close()
    # Keep the intended prominent order independent of source-column enumeration.
    records.sort(key=lambda c:(0,prominent.index(c['name'])) if c['name'] in prominent else (1,c['name']))
    contract={'sourceHash':manifest['sourceHash'],'planHash':manifest['planHash'],
              'receiptHash':_hash(receipt),'artifactHash':artifact['artifactHash'],'feature':feature,
              'target':copy.deepcopy(target),'tableName':spec['name'],'columns':records}
    contract_hash=_hash(contract)
    header=[f'# sourceHash={manifest["sourceHash"]}',f'# planHash={manifest["planHash"]}',
            f'# displayContractHash={contract_hash}',f'# execute only in document={target["documentId"]} version={target["version"]}',
            f'host = sigma.workbook.get_page({target["pageId"]!r})',
            "if not isinstance(host, Page): raise RuntimeError('expected approved page')",
            f'table = host.get_element({target["elementId"]!r})',
            "if not isinstance(table, InputTable): raise RuntimeError('expected approved input table')",
            f"if table.name() != {spec['name']!r}: raise RuntimeError('input table name is stale')"]
    for i,c in enumerate(records):
        header += [f'c{i} = table.get_column({c["id"]!r})',
                   f"if not isinstance(c{i}, (LiteralEditableColumn, SingleSelectEditableColumn, MultiSelectEditableColumn)): raise RuntimeError('expected approved editable column')",
                   f"if c{i}.name() != {c['name']!r}: raise RuntimeError('approved column name is stale')"]
    program=list(header)
    program += [f'approved_ids = {[c["id"] for c in records]!r}',
                'for observed in table.list_columns():',
                "    if observed.id() not in approved_ids and not isinstance(observed, InputTableIDColumn): raise RuntimeError('foreign input column prevents display mutation')"]
    program.append('table.set_visible_columns(['+', '.join(f'c{i}' for i,c in enumerate(records) if c['visible'])+'])')
    for i,c in enumerate(records):
        fmt=c['format']
        if not fmt: continue
        if 'pattern' in fmt: program.append(f'c{i}.set_format_pattern({fmt["pattern"]!r})')
        elif 'decimals' in fmt: program.append(f'c{i}.set_format({fmt["kind"]!r}, decimals={fmt["decimals"]})')
        else: program.append(f'c{i}.set_format({fmt["kind"]!r})')
    program.append(f"print('display requested; fresh SDK readback required: {contract_hash}')")
    readback=list(header)+['observations = []','for column in table.list_columns():',
        '    fmt = column.format()',
        "    observations.append({'id': column.id(), 'name': column.name(), 'visible': column.is_visible(), 'format': None if fmt is None else {'kind': fmt.kind(), 'decimals': fmt.decimal_places(), 'currency': fmt.currency(), 'pattern': fmt.pattern()}})",
        f"print({{'source': 'sdk-column-readback', 'contractHash': {contract_hash!r}, 'elementId': table.id(), 'columns': observations}})"]
    return {**contract,'contractHash':contract_hash,'status':'PENDING_SDK_READBACK',
            'executionBinding':{'workbookId':target['documentId'],'documentVersion':target['version']},
            'limitations':['SDK has no documented workbook-version CAS; caller must bind the execution context to this document/version and capture a fresh read-only SDK session after execution.',
                          'REST code-rep exports omit input display metadata and are not display readback evidence.'],
            'program':'\n'.join(program)+'\n','readbackProgram':'\n'.join(readback)+'\n'}


def verify_input_display(display,readback):
    keys=('sourceHash','planHash','receiptHash','artifactHash','feature','target','tableName','columns')
    if _hash({key:display[key] for key in keys})!=display.get('contractHash'):
        raise ValueError('display contract hash is stale')
    if readback.get('source')!='sdk-column-readback' or not readback.get('readbackId'):
        raise ValueError('fresh SDK column readback evidence is required')
    if (readback.get('contractHash')!=display['contractHash']
            or readback.get('documentId')!=display['target']['documentId']
            or readback.get('elementId')!=display['target']['elementId']
            or not isinstance(readback.get('documentVersion'),int)
            or isinstance(readback.get('documentVersion'),bool)
            or readback['documentVersion']<display['target']['version']):
        raise ValueError('SDK display readback target/version/contract is stale')
    records=readback.get('columns')
    if not isinstance(records,list) or any(not isinstance(c,dict) for c in records):
        raise ValueError('SDK display columns are missing')
    observed={c.get('id'):c for c in records}
    if len(observed)!=len(records): raise ValueError('duplicate SDK display columns')
    mismatches=[]
    for c in display['columns']:
        actual=observed.get(c['id']) or {}
        if actual.get('name')!=c['name'] or actual.get('visible') is not c['visible']:
            mismatches.append({'columnId':c['id'],'field':'visibility/name'})
        intended=c['format']; fmt=actual.get('format')
        if fmt is not None and not isinstance(fmt,dict):
            raise ValueError('SDK format observation must be a structured record or null')
        if intended:
            matches=(fmt is None or fmt.get('kind')=='automatic') if intended=={'kind':'automatic'} else (
                isinstance(fmt,dict) and all(fmt.get(key)==value for key,value in intended.items()))
            if not matches: mismatches.append({'columnId':c['id'],'field':'format','expected':intended,'actual':fmt})
    expected={c['id'] for c in display['columns']}
    extra=[c for c in records if c.get('id') not in expected]
    if len(extra)!=1 or extra[0].get('name')!='Row ID':
        mismatches.append({'field':'complete system Row ID column readback required'})
    if any(c.get('visible') is not False for c in records if c.get('id') not in expected):
        mismatches.append({'field':'unexpected visible columns'})
    unresolved=[{'columnId':c['id'],'sourceFormat':c['sourceFormat'],'reason':c['unresolvedReason']}
                for c in display['columns'] if c['unresolvedReason']]
    return {'status':'FAILED' if mismatches else ('PARTIAL' if unresolved else 'PASSED'),
            'supportedDisplayStatus':'FAILED' if mismatches else 'PASSED',
            'contractHash':display['contractHash'],'readbackId':readback['readbackId'],
            'documentId':readback['documentId'],'documentVersion':readback['documentVersion'],
            'mismatches':mismatches,'unresolved':unresolved,
            'verificationScope':'SDK column visibility/format metadata only; rendered output, data invariants and UI validation require separate acceptance'}
