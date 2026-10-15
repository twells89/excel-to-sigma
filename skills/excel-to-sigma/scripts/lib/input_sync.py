"""Typed, resumable input-table seeding plans and reviewable Sigma SDK programs.

This module never queries Sigma.  The caller must query the current table with the
companion MCP describe/query tools, pass those rows to :func:`plan_seed`, execute
the emitted program, and query again before recording a checkpoint.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from collections.abc import Mapping, Sequence


MAX_BATCH_ROWS = 2000


class InputSyncError(ValueError):
    pass


def _json_value(value):
    if isinstance(value, dt.datetime):
        return {"$type": "datetime", "value": value.isoformat()}
    if isinstance(value, dt.date):
        return {"$type": "date", "value": value.isoformat()}
    if isinstance(value, Mapping):
        return {str(key): _json_value(item) for key, item in sorted(value.items())}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    return value


def _digest(value):
    encoded = json.dumps(_json_value(value), sort_keys=True, separators=(",", ":"),
                         ensure_ascii=False)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _typed(value):
    if value is None:
        tag = "null"
    elif isinstance(value, bool):
        tag = "bool"
    elif isinstance(value, int):
        tag = "int"
    elif isinstance(value, float):
        tag = "float"
    elif isinstance(value, dt.datetime):
        tag, value = "datetime", value.isoformat()
    elif isinstance(value, dt.date):
        tag, value = "date", value.isoformat()
    elif isinstance(value, str):
        tag = "str"
    else:
        raise InputSyncError(f"unsupported business-key type: {type(value).__name__}")
    return {"type": tag, "value": value}


def _key(row, columns, origin):
    values = []
    for column in columns:
        if column not in row or row[column] is None:
            raise InputSyncError(f"missing business key {column!r} in {origin} row")
        values.append(_typed(row[column]))
    return values


def _key_token(key):
    return json.dumps(key, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def _index(rows, key_columns, origin):
    result = {}
    for row in rows:
        if not isinstance(row, Mapping):
            raise InputSyncError(f"{origin} rows must be objects")
        key = _key(row, key_columns, origin)
        token = _key_token(key)
        if token in result:
            raise InputSyncError(f"duplicate {origin} business key: {token}")
        result[token] = (key, dict(row))
    return result


def _same(left, right):
    return _json_value(left) == _json_value(right) and type(left) is type(right)


def _prior_index(prior_seed):
    result = {}
    if not prior_seed:
        return result
    for record in prior_seed.get("seedRecords", []):
        key = record.get("key")
        if isinstance(key, list) and isinstance(record.get("lastMigrationValue"), dict):
            result[_key_token(key)] = record
    return result


def plan_seed(source_rows, current_rows, key_columns, row_id_column, prior_seed=None,
              snapshot=None):
    """Plan inserts and safe row-ID updates without overwriting observed user edits."""
    source_rows = list(source_rows)
    current_rows = list(current_rows)
    if not key_columns or not all(isinstance(column, str) and column for column in key_columns):
        raise InputSyncError("key_columns must contain at least one column name")
    source = _index(source_rows, key_columns, "source")
    current = _index(current_rows, key_columns, "current")
    previous = _prior_index(prior_seed)
    inserts, updates, conflicts, unchanged, orphans, records = [], [], [], [], [], []

    for token in sorted(source):
        key, wanted = source[token]
        if token not in current:
            entry = {"key": key, "source": wanted}
            inserts.append(entry)
            records.append({"key": key, "lastMigrationValue": wanted})
            continue
        _current_key, observed = current[token]
        row_id = observed.get(row_id_column)
        if row_id is None:
            raise InputSyncError(f"current row for {token} has no real server row ID in {row_id_column!r}")
        comparable = {name: value for name, value in observed.items() if name != row_id_column}
        changed_columns = [name for name, value in wanted.items()
                           if name not in key_columns and not _same(comparable.get(name), value)]
        if not changed_columns:
            unchanged.append({"key": key, "rowId": row_id, "source": wanted})
            records.append({"key": key, "lastMigrationValue": wanted})
            continue

        last = (previous.get(token) or {}).get("lastMigrationValue")
        safe, conflicting = {}, []
        for name in changed_columns:
            if isinstance(last, dict) and name in last and _same(comparable.get(name), last[name]):
                safe[name] = wanted[name]
            else:
                conflicting.append(name)
        if conflicting:
            conflicts.append({"key": key, "rowId": row_id, "columns": conflicting,
                              "source": wanted, "current": comparable,
                              "lastMigrationValue": last})
            records.append({"key": key, "lastMigrationValue": last or wanted})
        else:
            updates.append({"key": key, "rowId": row_id, "changes": safe,
                            "source": wanted, "current": comparable,
                            "lastMigrationValue": last})
            records.append({"key": key, "lastMigrationValue": wanted})

    for token in sorted(set(current) - set(source)):
        key, row = current[token]
        orphans.append({"key": key, "rowId": row.get(row_id_column), "current": row})

    source_hash = _digest(list(source_rows))
    logical = {"sourceHash": source_hash, "keyColumns": list(key_columns),
               "rowIdColumn": row_id_column, "inserts": inserts, "updates": updates,
               "conflicts": conflicts, "unchanged": unchanged, "orphans": orphans}
    return {**logical, "planHash": _digest(logical), "snapshot": dict(snapshot or {}),
            "seedRecords": records,
            "limitations": [
                "No compare-and-swap is available between MCP query and write.",
                "Re-query after interruptions and before each emitted write program.",
                "Orphans are reported and are never deleted automatically.",
            ]}


def _parse_datetime(value, date_only=False):
    if isinstance(value, dt.datetime):
        if value.tzinfo is not None:
            value = value.astimezone(dt.timezone.utc).replace(tzinfo=None)
        return value
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day, 12)
    if isinstance(value, str):
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            parsed = dt.date.fromisoformat(value)
            return dt.datetime(parsed.year, parsed.month, parsed.day, 12)
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone(dt.timezone.utc).replace(tzinfo=None)
        return parsed
    return value


def _sdk_literal(value, column_type=None):
    if column_type == "datetime":
        date_only = (isinstance(value, dt.date) and not isinstance(value, dt.datetime)) or (
            isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)
        )
        value = _parse_datetime(value)
        if date_only:
            return f"datetime({value.year}, {value.month}, {value.day}, 12)"
    if isinstance(value, dt.datetime):
        return (f"datetime({value.year}, {value.month}, {value.day}, {value.hour}, "
                f"{value.minute}, {value.second})")
    if isinstance(value, dt.date):
        return f"datetime({value.year}, {value.month}, {value.day}, 12)"
    if value is None:
        return "None"
    return repr(value)


def _validate_spec(spec):
    required = ("pageId", "connectionId", "name", "columns")
    missing = [key for key in required if not spec.get(key)]
    if missing:
        raise InputSyncError("table spec missing: " + ", ".join(missing))
    names = [column.get("name") for column in spec["columns"]]
    if len(names) != len(set(names)) or any(not name for name in names):
        raise InputSyncError("editable column names must be unique and nonempty")
    if any(column.get("provenance") == "system" for column in spec["columns"]):
        raise InputSyncError("system columns cannot be seeded as editable columns")


def _row_expression(row, variables, types):
    fields = []
    for name, variable in variables.items():
        if name in row:
            fields.append(f"{variable}: {_sdk_literal(row[name], types.get(name))}")
    return "{" + ", ".join(fields) + "}"


def emit_construction_program(spec, seed_rows, *, source_hash, plan_hash):
    """Emit a restricted-dialect program that creates one independent input table."""
    _validate_spec(spec)
    if len(seed_rows) > MAX_BATCH_ROWS:
        raise InputSyncError("construction seed exceeds 2000 rows; create first, then emit checkpointed write chunks")
    columns = spec["columns"]
    types = {column["name"]: column["type"] for column in columns}
    lines = [
        f"# sourceHash={source_hash}", f"# planHash={plan_hash}",
        f"host = sigma.workbook.get_page({spec['pageId']!r})",
        "if not isinstance(host, Page):", "    raise RuntimeError('expected current page host')",
    ]
    schema = spec.get("writebackSchemaId")
    create = f"host.add_input_table({spec['connectionId']!r}"
    if schema:
        create += f", writeback_schema_id={schema!r}"
    frame = spec.get("frame") or [1, 1, 24, 12]
    if not (isinstance(frame, Sequence) and len(frame) == 4):
        raise InputSyncError("frame must contain x, y, width, height")
    lines += [f"created = {create})", f"created.set_name({spec['name']!r})",
              f"created.set_edit_mode({spec.get('editMode', 'draft')!r})",
              f"created.set_frame({frame[0]}, {frame[1]}, {frame[2]}, {frame[3]})"]
    for index, column in enumerate(columns):
        lines.append(f"created_c{index} = created.add_editable_column({column['name']!r}, {column['type']!r})")
    created_ids = ", ".join(f"created_c{index}.id()" for index in range(len(columns)))
    lines += [f"created_column_ids = [{created_ids}]",
              "system_columns = created.add_system_columns(['row_id'])",
              "created_id = created.id()",
              f"host = sigma.workbook.get_page({spec['pageId']!r})",
              "table = host.get_element(created_id)",
              "if not isinstance(table, InputTable):", "    raise RuntimeError('expected input table')"]
    variables = {}
    for index, column in enumerate(columns):
        variable = f"c{index}"; variables[column["name"]] = variable
        lines += [f"{variable} = table.get_column(created_column_ids[{index}])",
                  f"if not isinstance({variable}, LiteralEditableColumn):",
                  f"    raise RuntimeError('expected editable column: {column['name']}')"]
    if seed_rows:
        lines.append("row_ids = table.upsert_rows([")
        lines.extend(f"    {_row_expression(row, variables, types)}," for row in seed_rows)
        lines.append("])")
    lines += ["print('inputTableId=' + table.id())"]
    for name, variable in variables.items():
        lines.append(f"print({('column:' + name + '=')!r} + {variable}.id())")
    return "\n".join(lines) + "\n"


def emit_write_programs(seed_plan, target, max_batch_rows=MAX_BATCH_ROWS):
    """Emit checkpointed update/insert programs using assigned server IDs."""
    if not 1 <= max_batch_rows <= MAX_BATCH_ROWS:
        raise InputSyncError("max_batch_rows must be between 1 and 2000")
    for key in ("pageId", "elementId", "columnIds"):
        if not target.get(key):
            raise InputSyncError(f"target requires {key}")
    actions = [("update", item) for item in seed_plan.get("updates", [])]
    actions += [("insert", item) for item in seed_plan.get("inserts", [])]
    columns = dict(target["columnIds"])
    types = dict(target.get("columnTypes") or {})
    programs = []
    for offset in range(0, len(actions), max_batch_rows):
        chunk = actions[offset:offset + max_batch_rows]
        chunk_index = offset // max_batch_rows
        checkpoint = _digest({"sourceHash": seed_plan["sourceHash"],
                              "planHash": seed_plan["planHash"],
                              "target": target, "chunkIndex": chunk_index,
                              "actions": chunk})
        lines = [f"# sourceHash={seed_plan['sourceHash']}", f"# planHash={seed_plan['planHash']}",
                 f"# checkpointId={checkpoint}",
                 f"host = sigma.workbook.get_page({target['pageId']!r})",
                 "if not isinstance(host, Page):", "    raise RuntimeError('expected current page host')",
                 f"table = host.get_element({target['elementId']!r})",
                 "if not isinstance(table, InputTable):", "    raise RuntimeError('expected input table')"]
        variables = {}
        for index, (name, column_id) in enumerate(columns.items()):
            variable = f"c{index}"; variables[name] = variable
            lines += [f"{variable} = table.get_column({column_id!r})",
                      f"if not isinstance({variable}, LiteralEditableColumn):",
                      f"    raise RuntimeError('expected editable column: {name}')"]
        update_rows = [item for kind, item in chunk if kind == "update"]
        insert_rows = [item for kind, item in chunk if kind == "insert"]
        if update_rows:
            lines.append("table.update_rows([")
            for item in update_rows:
                lines.append(f"    ({item['rowId']!r}, {_row_expression(item['changes'], variables, types)}),")
            lines.append("])")
        if insert_rows:
            lines.append("row_ids = table.upsert_rows([")
            lines.extend(f"    {_row_expression(item['source'], variables, types)}," for item in insert_rows)
            lines.append("])")
        lines.append(f"print({'checkpointId=' + checkpoint!r})")
        programs.append({"chunkIndex": chunk_index, "checkpointId": checkpoint,
                         "actionCount": len(chunk), "program": "\n".join(lines) + "\n"})
    return programs


def build_shared_input_architecture(plan, *, connection_id, folder_id):
    """Return the approved logical shared-input architecture before manifest hashing."""
    if not connection_id or not folder_id:
        raise InputSyncError("connection_id and folder_id are required target scope")
    policy_rows = [
        {"Category": "Airfare", "Per-item Limit": 1200, "Receipt Threshold": 25, "Policy Owner": "Travel"},
        {"Category": "Lodging", "Per-item Limit": 350, "Receipt Threshold": 25, "Policy Owner": "Travel"},
        {"Category": "Meals", "Per-item Limit": 90, "Receipt Threshold": 25, "Policy Owner": "Finance"},
        {"Category": "Ground Transport", "Per-item Limit": 180, "Receipt Threshold": 25, "Policy Owner": "Finance"},
        {"Category": "Software", "Per-item Limit": 500, "Receipt Threshold": 0, "Policy Owner": "Finance"},
        {"Category": "Supplies", "Per-item Limit": 250, "Receipt Threshold": 25, "Policy Owner": "Finance"},
    ]
    source_inputs = list(plan.get("sourceInputs") or [])
    bindings = {item["id"]: item["sourceCellId"] for item in source_inputs}
    def input_type(value):
        if isinstance(value, bool):
            return "boolean"
        if isinstance(value, (int, float)):
            return "number"
        if isinstance(value, (dt.date, dt.datetime)):
            return "datetime"
        return "text"
    research_rows = [{"Scope": "shared", **{
        item["id"]: item.get("defaultValue") for item in source_inputs
    }}]
    mapped = ["sales-as-of-date", "sales-timing-status-quoting-repair",
              "expense-policy", "research-model-inputs", "research-literal-overrides"]
    return {
        "targetScope": {"connectionId": connection_id, "folderId": folder_id},
        "mappedFeatures": mapped,
        "inputs": {
            "salesSettings": {
                "kind": "input-table", "name": "Sales Settings", "keyColumns": ["Setting"],
                "connectionId": connection_id,
                "sourceBinding": "Assumptions!F2",
                "rows": [{"Setting": "AsOfDate", "AsOfDate": "2026-06-30"}],
                "columns": [{"name": "Setting", "type": "text", "provenance": "source"},
                            {"name": "AsOfDate", "type": "datetime", "provenance": "source"}],
            },
            "expensePolicy": {
                "kind": "input-table", "name": "Expense Policy", "keyColumns": ["Category"],
                "connectionId": connection_id,
                "sourceRange": "Expense Policy!A2:D7", "rows": policy_rows,
                "columns": [
                    {"name": "Category", "type": "text", "provenance": "source"},
                    {"name": "Per-item Limit", "type": "number", "provenance": "source"},
                    {"name": "Receipt Threshold", "type": "number", "provenance": "source"},
                    {"name": "Policy Owner", "type": "text", "provenance": "source"},
                ],
                "unknownCategoryBehavior": "explicit-null-error",
                "duplicateOrMissingKeyBehavior": "block-before-join",
            },
            "researchModelInputs": {
                "kind": "input-table", "name": "Research Model Inputs", "keyColumns": ["Scope"],
                "connectionId": connection_id,
                "rows": research_rows, "editableColumnBindings": list(bindings),
                "columns": [{"name": "Scope", "type": "text", "provenance": "migration"}] + [
                    {"name": item["id"], "type": input_type(item.get("defaultValue")),
                     "provenance": item["sourceCellId"]}
                    for item in source_inputs
                ],
                "sourceCellBindings": bindings, "joinGrain": ["Scope"],
                "excludedFromGrouping": ["Row ID", "Created At", "Created By", "Updated At", "Updated By"],
            },
        },
        "joins": {
            "sales": {"type": "left_outer", "key": "constant setting scope", "cardinality": "many-to-one"},
            "expense": {"type": "left_outer", "key": "Category", "cardinality": "many-to-one"},
            "research": {"type": "left_outer", "key": "constant Scope", "cardinality": "many-to-one"},
        },
        "calculations": {
            "TimingStatus": 'If([Stage] = "Closed Won" or [Stage] = "Closed Lost", "Closed", If(DateTrunc("day", [Close Date]) < [AsOfDate], "Slipped", "Open"))',
            "PolicyExcess": "If(IsNull([Per-item Limit]), Null, Max([Amount] - [Per-item Limit], 0))",
        },
        "provenance": {
            "referenceQuotingRepair": {
                "source": "Pipeline Timing Status formulas contain C2<\\'Assumptions\\'!$F$2",
                "replacement": "reviewed day-grain shared AsOfDate join",
                "status": "reviewed-replacement-required",
            }
        },
        "acceptance": {
            "sales": "edit 2026-06-30 to 2026-07-01, verify boundary rows, restore",
            "expense": "edit Meals 90 to 400, verify all matching rows/counts, restore",
            "research": "edit growth and literal-year overrides, verify downstream years, restore",
            "persistence": "reopen/refresh and query unique key counts after each restore",
        },
    }
