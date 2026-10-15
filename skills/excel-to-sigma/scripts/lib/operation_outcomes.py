"""Persist outcome only after a fresh complete native business readback.

Controller programs are explicit follow-ups, not action-chain success effects.
Input-table updates have no compare-and-swap transaction; serialize controller
reconciliation per workbook and read back the resulting row before reporting it.
"""
import datetime
try:
    from .native_actions import audit_native_action_readback
    from .promotion import _header,_rows,_digest
    from .workflow_repairs import _project
except ImportError:
    from native_actions import audit_native_action_readback
    from promotion import _header,_rows,_digest
    from workflow_repairs import _project


def reconcile_operation(expected,capture,operations,*,completed_at):
    if capture.get('origin')!='sigma-query' or not capture.get('executionId'):
        raise ValueError('actual fresh business query required')
    for key in ('sourceHash','planHash','documentId','version','operationId','context'):
        if not expected.get(key) or expected[key]!=capture.get(key):raise ValueError('operation readback binding mismatch')
    datetime.datetime.fromisoformat(completed_at.replace('Z','+00:00'))
    audit=audit_native_action_readback(expected,capture)
    if audit['status']!='PASSED_COMPONENT_READBACK':raise ValueError('business outcome not confirmed: '+', '.join(audit['failures']))
    indexed=_rows(operations);row=indexed.get(expected['operationId'])
    if not row or row.get('State') not in {'PENDING','PREPARED'}:raise ValueError('pending original operation required; no replay')
    for name,key in [('Workflow','workflow'),('Context','context'),('Source Hash','sourceHash'),('Plan Hash','planHash')]:
        if row.get(name)!=expected[key]:raise ValueError('durable operation context mismatch')
    columns=operations['binding']['sdkColumns'];variables={c['name']:f'c{i}' for i,c in enumerate(columns)}
    if not {'State','Outcome','Completed At'}<=set(variables):raise ValueError('missing operation outcome columns')
    lines=_header(operations)
    lines += [f'# sourceHash={expected["sourceHash"]}; planHash={expected["planHash"]}; evidence={capture["evidenceId"]}',
        '# Controller must serialize this follow-up and capture the updated row; no SDK return proves a write.',
        'from datetime import datetime',
        f"table.update_rows([({expected['operationId']!r}, {{{variables['State']}: 'CONFIRMED', {variables['Outcome']}: {'Business readback confirmed; evidence='+capture['evidenceId']!r}, {variables['Completed At']}: datetime.fromisoformat({completed_at.replace('Z','+00:00')!r})}})])"]
    return {'businessAudit':audit,'operationId':expected['operationId'],'operationsBeforeHash':_digest(operations),
        'program':'\n'.join(lines)+'\n','status':'PENDING_OUTCOME_ROW_READBACK','productionStatus':'BLOCKED'}


def verify_outcome_row(plan,rows):
    index=_rows(rows);row=index.get(plan['operationId'])
    if not row or row.get('State')!='CONFIRMED' or plan['businessAudit']['evidenceId'] not in (row.get('Outcome') or '') or not row.get('Completed At'):
        raise ValueError('persisted outcome row not confirmed')
    return {'status':'CONFIRMED_COMPONENT_OUTCOME','operationId':plan['operationId'],'readbackHash':_digest(rows),'productionStatus':'BLOCKED'}


def build_outcome_view(binding):
    if not binding.get('elementId') or not binding.get('name'):raise ValueError('assigned operations table required')
    name=binding['name']
    return _project('workflow-confirmed-outcomes','Workflow results',binding,
        [(n,f'[{name}/{n}]') for n in ('Workflow','Target Business Key','Context','Created At','Completed At','Outcome')]+[
        ('Result',f'If([{name}/State] = "CONFIRMED", "Confirmed by business readback", [{name}/State] = "FAILED", "Failed; inspect and recover", "Pending verification")')])
