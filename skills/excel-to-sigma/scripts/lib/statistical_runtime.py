"""Offline source-bound Sigma Python preview emitter, with no REST mutation.

The caller must supply reviewed, actual element titles and paired Snowpark column
identifier/Row.as_dict keys. Output remains blocked pending an Excel-qualified ETS
provider, actual Sigma execution and the repository's phased output guards.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path


REQUIRED = {
    "history": ("Month", "SKU", "Demand"),
    "parameters": ("Key", "Value"),
    "targets": ("Period", "Month"),
}
FIELDS = (("Status", "STRING"), ("ReadyToPublish", "BOOLEAN"),
          ("ErrorCode", "STRING"), ("InputHash", "STRING"),
          ("SourceHash", "STRING"), ("Month", "STRING"), ("SKU", "STRING"),
          ("ETS", "DOUBLE"), ("Linear", "DOUBLE"),
          ("SelectedSKUTrend", "DOUBLE"), ("Selected3MMA", "DOUBLE"),
          ("FY26ETSTotal", "DOUBLE"), ("StdDev36M", "DOUBLE"),
          ("SafetyStock", "DOUBLE"), ("ReorderPoint", "DOUBLE"),
          ("IntendedSequenceDescending", "INTEGER"))
OUTPUT_NAME = "StatisticalForecastPreview"


def _binding(binding):
    if not isinstance(binding, dict) or set(binding) != set(REQUIRED):
        raise ValueError("exact history/parameters/targets bindings required")
    titles = []
    for role, fields in REQUIRED.items():
        entry = binding[role]
        if not isinstance(entry, dict) or set(entry) != {"title", "columns"}:
            raise ValueError("reviewed title and column map required")
        title, columns = entry["title"], entry["columns"]
        if not isinstance(title, str) or not title or not isinstance(columns, dict) or set(columns) != set(fields):
            raise ValueError("exact runtime source columns required")
        titles.append(title)
        identifiers, row_keys = [], []
        for name in fields:
            pair = columns[name]
            if (not isinstance(pair, dict) or set(pair) != {"identifier", "rowKey"}
                    or any(not isinstance(v, str) or not v for v in pair.values())):
                raise ValueError("observed identifier and rowKey pair required")
            identifiers.append(pair["identifier"])
            row_keys.append(pair["rowKey"])
        if len(set(identifiers)) != len(fields) or len(set(row_keys)) != len(fields):
            raise ValueError("ambiguous runtime column map")
    if len(set(titles)) != len(titles):
        raise ValueError("ambiguous input element titles")
    return binding


def emit_preview_program(binding):
    """Return a reviewable program, never a production-ready port."""
    binding = _binding(binding)
    kernel = Path(__file__).with_name("statistical_forecast.py").read_text()
    code = kernel + "\n" + '''
import hashlib
import json

SOURCE_BINDING = ''' + repr(binding) + '''
OUTPUT_COLUMNS = ''' + repr([name for name, _ in FIELDS]) + '''
OUTPUT_TYPES = ''' + repr(dict(FIELDS)) + '''


def read_bound_rows(role):
    spec = SOURCE_BINDING[role]
    frame = sigma.get_element(spec['title'])
    observed = list(frame.columns)
    columns = spec['columns']
    if any(observed.count(pair['identifier']) != 1 for pair in columns.values()):
        raise StatisticalPortError('RUNTIME_COLUMN_MISMATCH')
    rows = []
    for record in frame.collect():
        raw = record.as_dict()
        if any(pair['rowKey'] not in raw for pair in columns.values()):
            raise StatisticalPortError('RUNTIME_ROW_KEY_MISMATCH')
        rows.append({name: raw[pair['rowKey']] for name, pair in columns.items()})
    return rows


input_hash = None
try:
    history = read_bound_rows('history')
    parameters = read_bound_rows('parameters')
    targets = read_bound_rows('targets')
    raw_input = {'history': sorted(history, key=lambda r: (str(r['Month']), str(r['SKU']))),
                 'parameters': sorted(parameters, key=lambda r: str(r['Key'])),
                 'targets': sorted(targets, key=lambda r: str(r['Period']))}
    input_hash = hashlib.sha256(json.dumps(raw_input, sort_keys=True, default=str,
                    separators=(',', ':'), allow_nan=False).encode()).hexdigest()
    if len(parameters) != 4 or set(r['Key'] for r in parameters) != {
            'Selected SKU', 'Seasonality (months)', 'Service level', 'Lead time (months)'}:
        raise StatisticalPortError('INVALID_PARAMETER_KEYS')
    p = {r['Key']: r['Value'] for r in parameters}
    if len(targets) != 12 or set(r['Period'] for r in targets) != set(range(1, 13)):
        raise StatisticalPortError('INVALID_TARGET_KEYS')
    ordered_targets = [r['Month'] for r in sorted(targets, key=lambda r: r['Period'])]
    result = calculate(history, p['Selected SKU'], p['Seasonality (months)'],
                       p['Service level'], p['Lead time (months)'], ordered_targets)
    summaries = {r['SKU']: r for r in result['Summary']}
    rows = []
    for period_index, period in enumerate(result['Periods']):
        for sku in SKUS:
            summary = summaries[sku]
            rows.append(dict(Status=result['Status'], ReadyToPublish=False,
                ErrorCode='EXCEL_ETS_PROVIDER_UNQUALIFIED', InputHash=input_hash,
                SourceHash=SOURCE_SHA256, Month=period['Month'], SKU=sku,
                ETS=None, Linear=period['Linear'][sku],
                SelectedSKUTrend=period['SelectedSKUTrend'],
                Selected3MMA=period['Selected3MMA'], FY26ETSTotal=None,
                StdDev36M=summary['StdDev36M'], SafetyStock=summary['SafetyStock'],
                ReorderPoint=None,
                IntendedSequenceDescending=result['IntendedSequenceDescending'][period_index]))
except Exception as error:
    rows = [dict(Status='FAILURE', ReadyToPublish=False,
                 ErrorCode=str(error) if isinstance(error, StatisticalPortError) else 'RUNTIME_INPUT_ERROR',
                 InputHash=input_hash, SourceHash=SOURCE_SHA256)]

try:
    values = [[typed_output_value(row.get(name), OUTPUT_TYPES[name]) for name in OUTPUT_COLUMNS]
              for row in rows]
    output_frame = session.create_dataframe(values, schema=''' + repr(", ".join('"'+name+'" '+type_ for name,type_ in FIELDS)) + ''')
except Exception:
    fallback = dict(Status='FAILURE', ReadyToPublish=False,
                    ErrorCode='RUNTIME_OUTPUT_TYPE_ERROR', SourceHash=SOURCE_SHA256)
    output_frame = session.create_dataframe([[typed_output_value(fallback.get(name), OUTPUT_TYPES[name])
                                              for name in OUTPUT_COLUMNS]],
        schema=''' + repr(", ".join('"'+name+'" '+type_ for name,type_ in FIELDS)) + ''')
sigma.output(''' + repr(OUTPUT_NAME) + ''', output_frame)
'''
    return {"code": code, "codeHash": hashlib.sha256(code.encode()).hexdigest(),
            "output": OUTPUT_NAME, "schema": list(FIELDS),
            "sourceBindingHash": hashlib.sha256(json.dumps(binding, sort_keys=True).encode()).hexdigest(),
            "runtimeStatus": "UNRESOLVED", "productionStatus": "BLOCKED",
            "reason": "EXCEL_ETS_PROVIDER_UNQUALIFIED"}
