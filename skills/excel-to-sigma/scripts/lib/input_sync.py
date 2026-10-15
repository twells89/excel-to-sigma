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
        if isinstance(key, list) and any(isinstance(record.get(field), dict) for field in (
                "confirmedMigrationValue", "pendingMigrationValue", "lastMigrationValue")):
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
            prior = previous.get(token) or {}
            records.append({"key": key,
                            "confirmedMigrationValue": prior.get("confirmedMigrationValue"),
                            "pendingMigrationValue": wanted})
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
            records.append({"key": key, "confirmedMigrationValue": wanted,
                            "pendingMigrationValue": None})
            continue

        prior = previous.get(token) or {}
        confirmed = prior.get("confirmedMigrationValue")
        pending = prior.get("pendingMigrationValue")
        legacy = prior.get("lastMigrationValue")
        if confirmed is None and isinstance(legacy, dict):
            confirmed = legacy
        if isinstance(pending, dict) and all(
                name in comparable and _same(comparable[name], value)
                for name, value in pending.items()):
            confirmed = pending
        safe, conflicting = {}, []
        for name in changed_columns:
            if (isinstance(confirmed, dict) and name in confirmed
                    and _same(comparable.get(name), confirmed[name])):
                safe[name] = wanted[name]
            else:
                conflicting.append(name)
        if conflicting:
            conflicts.append({"key": key, "rowId": row_id, "columns": conflicting,
                              "source": wanted, "current": comparable,
                              "confirmedMigrationValue": confirmed})
            records.append({"key": key, "confirmedMigrationValue": confirmed,
                            "pendingMigrationValue": pending})
        else:
            updates.append({"key": key, "rowId": row_id, "changes": safe,
                            "source": wanted, "current": comparable,
                            "confirmedMigrationValue": confirmed})
            records.append({"key": key, "confirmedMigrationValue": confirmed,
                            "pendingMigrationValue": wanted})

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


def sigma_input_type(value):
    """Map a typed source literal to the ordinary input-table type vocabulary."""
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (dt.date, dt.datetime)):
        return "datetime"
    if isinstance(value, (int, float)):
        return "number"
    return "text"


def evaluate_timing_status(stage, close_date, as_of_date):
    """Reference behavior for the emitted day-grain sales calculation."""
    if stage in {"Closed Won", "Closed Lost"}:
        return "Closed"
    close_day = _parse_datetime(close_date).date()
    as_of_day = _parse_datetime(as_of_date).date()
    return "Slipped" if close_day < as_of_day else "Open"


def evaluate_expense_exception(row, policy_limit):
    """Reference behavior for the joined policy exception calculation."""
    if policy_limit is None:
        return {"policyExcess": None, "exception": None}
    excess = max(row["Amount"] - policy_limit, 0)
    if abs(row["Card Amount"] - row["Ledger Amount"]) > 0.01:
        status = "RECON"
    elif excess > 0:
        status = "OVER LIMIT"
    elif row["Receipt Status"] == "Missing":
        status = "MISSING RECEIPT"
    elif row["Approval Status"] != "Approved":
        status = "PENDING"
    else:
        status = "CLEAR"
    return {"policyExcess": excess, "exception": status}


def _manifest_architecture(manifest):
    projection = manifest.get("planProjection") or {}
    architecture = projection.get("sharedInputArchitecture")
    if not isinstance(architecture, dict) or not isinstance(architecture.get("inputs"), dict):
        raise InputSyncError("manifest does not contain hash-bound sharedInputArchitecture")
    return architecture


def _construction_spec(input_spec, binding):
    return {
        "pageId": binding.get("pageId"),
        "connectionId": input_spec.get("connectionId"),
        "name": input_spec.get("name"),
        "columns": input_spec.get("columns"),
        "writebackSchemaId": input_spec.get("writebackSchemaId"),
        "editMode": input_spec.get("editMode", "draft"),
        "frame": input_spec.get("frame", [1, 1, 24, 12]),
    }


def build_artifact_records(manifest, architecture, target_bindings):
    """Hash every approved feature's exact schema, seed rows, and target binding."""
    if architecture != _manifest_architecture(manifest):
        raise InputSyncError("architecture is not the immutable manifest architecture")
    inputs = architecture["inputs"]
    if set(target_bindings or {}) != set(inputs):
        raise InputSyncError("target bindings must exactly cover approved input tables")
    artifacts = {}
    wiring = build_operational_input_elements(architecture)
    wiring_hash = _digest(wiring)
    for feature in sorted(inputs):
        binding = target_bindings[feature]
        for field in ("documentId", "version", "pageId"):
            if binding.get(field) is None or binding.get(field) == "":
                raise InputSyncError(f"target binding for {feature} requires {field}")
        spec = _construction_spec(inputs[feature], binding)
        _validate_spec(spec)
        material = {
            "sourceHash": manifest.get("sourceHash"),
            "planHash": manifest.get("planHash"),
            "feature": feature,
            "constructionSpec": spec,
            "seedRows": inputs[feature].get("rows") or [],
            "targetBinding": binding,
            "workbookWiringHash": wiring_hash,
        }
        artifacts[feature] = {
            "feature": feature,
            "specHash": _digest(spec),
            "seedHash": _digest(material["seedRows"]),
            "targetBindingHash": _digest(binding),
            "workbookWiringHash": wiring_hash,
            "artifactHash": _digest(material),
        }
    return artifacts


def _approved_feature(manifest, receipt, feature, *, require_assigned=False):
    if receipt.get("status") != "APPROVED_STAGING":
        raise InputSyncError("input artifact requires an APPROVED_STAGING receipt")
    if (receipt.get("sourceHash") != manifest.get("sourceHash")
            or receipt.get("planHash") != manifest.get("planHash")):
        raise InputSyncError("approved staging receipt is stale")
    architecture = _manifest_architecture(manifest)
    if feature not in architecture["inputs"]:
        raise InputSyncError(f"feature is not approved: {feature}")
    bindings = receipt.get("targetBindings") or {}
    expected = build_artifact_records(manifest, architecture, bindings)
    if receipt.get("artifacts") != expected:
        raise InputSyncError("approved artifact binding does not match manifest architecture")
    binding = bindings[feature]
    if require_assigned:
        for field in ("elementId", "columnIds"):
            if not binding.get(field):
                raise InputSyncError(f"approved target binding for {feature} requires {field}")
        expected_columns = {column["name"] for column in architecture["inputs"][feature]["columns"]}
        if set(binding["columnIds"]) != expected_columns:
            raise InputSyncError("approved column IDs do not exactly cover the input schema")
    return architecture["inputs"][feature], binding, expected[feature]


def emit_workbook_wiring(manifest, staging_receipt):
    """Return approved REST-native join fragments bound to the artifact receipt."""
    architecture = _manifest_architecture(manifest)
    for feature in architecture["inputs"]:
        _approved_feature(manifest, staging_receipt, feature)
    return {
        "sourceHash": manifest["sourceHash"], "planHash": manifest["planHash"],
        "artifacts": staging_receipt["artifacts"],
        "wiring": build_operational_input_elements(architecture),
    }


def _row_expression(row, variables, types):
    fields = []
    for name, variable in variables.items():
        if name in row:
            fields.append(f"{variable}: {_sdk_literal(row[name], types.get(name))}")
    return "{" + ", ".join(fields) + "}"


def emit_construction_program(manifest, staging_receipt, feature):
    """Emit construction only from an immutable manifest and approved artifact receipt."""
    input_spec, binding, artifact = _approved_feature(manifest, staging_receipt, feature)
    spec = _construction_spec(input_spec, binding)
    seed_rows = input_spec.get("rows") or []
    _validate_spec(spec)
    if len(seed_rows) > MAX_BATCH_ROWS:
        raise InputSyncError("construction seed exceeds 2000 rows; create first, then emit checkpointed write chunks")
    columns = spec["columns"]
    types = {column["name"]: column["type"] for column in columns}
    lines = [
        f"# sourceHash={manifest['sourceHash']}", f"# planHash={manifest['planHash']}",
        f"# artifactHash={artifact['artifactHash']}", f"# feature={feature}",
        f"# targetDocumentId={binding['documentId']}", f"# targetVersion={binding['version']}",
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


def emit_write_programs(seed_plan, manifest, staging_receipt, feature,
                        max_batch_rows=MAX_BATCH_ROWS):
    """Emit writes bound to the manifest, approved artifact, and assigned target IDs."""
    if not 1 <= max_batch_rows <= MAX_BATCH_ROWS:
        raise InputSyncError("max_batch_rows must be between 1 and 2000")
    _input_spec, target, artifact = _approved_feature(
        manifest, staging_receipt, feature, require_assigned=True)
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
                              "approvedSourceHash": manifest["sourceHash"],
                              "approvedPlanHash": manifest["planHash"],
                              "artifactHash": artifact["artifactHash"],
                              "target": target, "chunkIndex": chunk_index,
                              "actions": chunk})
        lines = [f"# sourceHash={manifest['sourceHash']}", f"# planHash={manifest['planHash']}",
                 f"# reconciliationSourceHash={seed_plan['sourceHash']}",
                 f"# reconciliationPlanHash={seed_plan['planHash']}",
                 f"# artifactHash={artifact['artifactHash']}", f"# feature={feature}",
                 f"# targetDocumentId={target['documentId']}",
                 f"# targetVersion={target['version']}",
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
                         "artifactHash": artifact["artifactHash"], "feature": feature,
                         "targetBinding": target, "actionCount": len(chunk),
                         "program": "\n".join(lines) + "\n"})
    return programs


def build_shared_input_architecture(plan, *, connection_id, folder_id, workflows):
    """Build only the selected source workbook's hash-bound input architecture."""
    if not connection_id or not folder_id:
        raise InputSyncError("connection_id and folder_id are required target scope")
    if isinstance(workflows, str):
        workflows = [workflows]
    selected = list(dict.fromkeys(workflows or []))
    unknown = set(selected) - {"sales", "expense", "research"}
    if not selected or unknown:
        raise InputSyncError("workflows must select sales, expense, and/or research")
    shared_sources = plan.get("sharedInputSources") or {}
    architecture = {
        "workflows": selected,
        "targetScope": {"connectionId": connection_id, "folderId": folder_id},
        "mappedFeatures": [], "featureToInput": {}, "inputs": {}, "joins": {},
        "workbookWiring": {}, "calculations": {}, "provenance": {},
        "acceptance": {"persistence": "reopen/refresh and query unique key counts after each restore"},
        "acceptanceChecks": [],
    }
    if "sales" in selected:
        source = shared_sources.get("salesSettings") or {}
        if (source.get("sourceBinding") != "Assumptions!F2"
                or not isinstance(source.get("rows"), list) or len(source["rows"]) != 1):
            raise InputSyncError("sales plan requires validated Assumptions!F2 sharedInputSources")
        if (source["rows"][0].get("Setting") != "AsOfDate"
                or source["rows"][0].get("AsOfDate") is None):
            raise InputSyncError("sales source row must bind AsOfDate to Assumptions!F2")
        plan_seed(source["rows"], [], ["Setting"], "Row ID")
        architecture["inputs"]["salesSettings"] = {
            "kind": "input-table", "name": "Sales Settings", "keyColumns": ["Setting"],
            "connectionId": connection_id, "sourceBinding": source["sourceBinding"],
            "rows": source["rows"],
            "columns": [{"name": "Setting", "type": "text", "provenance": "source"},
                        {"name": "AsOfDate", "type": "datetime", "provenance": "source"}],
        }
        architecture["mappedFeatures"] += ["sales-as-of-date", "sales-timing-status-day-grain"]
        architecture["featureToInput"].update({
            "sales-as-of-date": "salesSettings",
            "sales-timing-status-day-grain": "salesSettings",
        })
        architecture["joins"]["sales"] = {
            "type": "left_outer", "key": "constant setting scope", "cardinality": "many-to-one"}
        architecture["workbookWiring"]["salesSource"] = {
            "elementId": source.get("sourceElementId", "salesSource"),
            "name": source.get("sourceElementName", "Sales Source"),
            "consumerSourceElementId": "salesWithSettings"}
        architecture["calculations"]["TimingStatus"] = (
            'If([Stage] = "Closed Won" or [Stage] = "Closed Lost", "Closed", '
            'If(DateTrunc("day", [Close Date]) < DateTrunc("day", [AsOfDate]), '
            '"Slipped", "Open"))')
        architecture["provenance"]["dayGrainTimingDecision"] = {
            "source": '=IF(H2="Closed Won","Closed",IF(H2="Closed Lost","Closed",IF(C2<\'Assumptions\'!$F$2,"Slipped","Open")))',
            "replacement": "compare DateTrunc day for both Close Date and shared AsOfDate",
            "status": "reviewed-equivalence"}
        architecture["acceptance"]["sales"] = (
            "edit source-derived AsOfDate from 2026-06-30 to 2026-07-01, verify boundary rows, restore")
        architecture["acceptanceChecks"].append("sales-as-of-date-boundary")
    if "expense" in selected:
        source = shared_sources.get("expensePolicy") or {}
        if (source.get("sourceRange") != "Policy!A1:D7"
                or not isinstance(source.get("rows"), list) or not source["rows"]):
            raise InputSyncError("expense plan requires validated Policy!A1:D7 sharedInputSources")
        expected_policy_columns = {
            "Category", "Per-item Limit", "Receipt Threshold", "Policy Owner"}
        if any(set(row) != expected_policy_columns for row in source["rows"]):
            raise InputSyncError("expense policy rows must preserve all Policy!A1:D7 columns")
        plan_seed(source["rows"], [], ["Category"], "Row ID")
        architecture["inputs"]["expensePolicy"] = {
            "kind": "input-table", "name": "Expense Policy", "keyColumns": ["Category"],
            "connectionId": connection_id, "sourceRange": source["sourceRange"],
            "rows": source["rows"],
            "columns": [{"name": "Category", "type": "text", "provenance": "source"},
                        {"name": "Per-item Limit", "type": "number", "provenance": "source"},
                        {"name": "Receipt Threshold", "type": "number", "provenance": "source"},
                        {"name": "Policy Owner", "type": "text", "provenance": "source"}],
            "unknownCategoryBehavior": "explicit-null-error",
            "duplicateOrMissingKeyBehavior": "block-before-join",
        }
        architecture["mappedFeatures"].append("expense-policy")
        architecture["featureToInput"]["expense-policy"] = "expensePolicy"
        architecture["joins"]["expense"] = {
            "type": "left_outer", "key": "Category", "cardinality": "many-to-one"}
        architecture["workbookWiring"]["expenseSource"] = {
            "elementId": source.get("sourceElementId", "expenseSource"),
            "name": source.get("sourceElementName", "Expense Source"),
            "consumerSourceElementId": "expensesWithPolicy"}
        architecture["calculations"]["PolicyExcess"] = (
            "If(IsNull([Per-item Limit]), Null, Max([Amount] - [Per-item Limit], 0))")
        architecture["acceptance"]["expense"] = (
            "edit source-derived Meals limit 90 to 400, verify all matching rows/counts, restore")
        architecture["acceptanceChecks"].append("expense-meals-limit-propagation")
    if "research" in selected:
        source_inputs = list(plan.get("sourceInputs") or [])
        ids = [item.get("id") for item in source_inputs]
        cells = [item.get("sourceCellId") for item in source_inputs]
        if (any(not value for value in ids + cells) or len(set(ids)) != len(ids)
                or len(set(cells)) != len(cells)):
            raise InputSyncError("research sourceInputs require unique ids and sourceCellIds")
        bindings = {item["id"]: item["sourceCellId"] for item in source_inputs}
        architecture["inputs"]["researchModelInputs"] = {
            "kind": "input-table", "name": "Research Model Inputs", "keyColumns": ["Scope"],
            "connectionId": connection_id,
            "rows": [{"Scope": "shared", **{
                item["id"]: item.get("defaultValue") for item in source_inputs}}],
            "editableColumnBindings": list(bindings),
            "columns": [{"name": "Scope", "type": "text", "provenance": "migration"}] + [
                {"name": item["id"], "type": sigma_input_type(item.get("defaultValue")),
                 "provenance": item["sourceCellId"]} for item in source_inputs],
            "sourceCellBindings": bindings, "joinGrain": ["Scope"],
            "excludedFromGrouping": ["Row ID", "Created At", "Created By", "Updated At", "Updated By"],
        }
        architecture["mappedFeatures"] += ["research-model-inputs", "research-literal-overrides"]
        architecture["featureToInput"].update({
            "research-model-inputs": "researchModelInputs",
            "research-literal-overrides": "researchModelInputs"})
        architecture["joins"]["research"] = {
            "type": "left_outer", "key": "constant Scope", "cardinality": "many-to-one"}
        architecture["acceptance"]["research"] = (
            "edit growth and literal-year overrides, verify downstream years, restore")
        architecture["acceptanceChecks"] += [
            "research-growth-propagation", "research-literal-override-propagation"]
    return architecture


def _rest_columns(spec, prefix):
    return [{"id": f"{prefix}_{index}", "name": column["name"], "type": column["type"]}
            for index, column in enumerate(spec["columns"])]


def build_operational_input_elements(architecture):
    """Build executable ordinary inputs and native join/calculation fragments."""
    inputs = architecture.get("inputs") or {}
    connection = (architecture.get("targetScope") or {}).get("connectionId")
    wiring = architecture.get("workbookWiring") or {}
    if not connection:
        raise InputSyncError("approved target connection is required")
    elements, consumers, validations = [], {}, {}
    if "salesSettings" in inputs:
        spec = inputs["salesSettings"]
        source = wiring.get("salesSource") or {}
        if not source.get("elementId") or not source.get("name"):
            raise InputSyncError("approved sales source wiring is required")
        elements.extend([
            {"id": "salesSettings", "kind": "input-table", "name": spec["name"],
             "inputMode": "edit", "source": {"kind": "empty", "connectionId": connection},
             "columns": _rest_columns(spec, "sales_setting")},
            {"id": "salesInputBase", "kind": "table", "name": "Sales Input Base",
             "source": {"kind": "table", "elementId": source["elementId"]},
             "columns": [
                 {"id": "sales_" + str(index), "name": name,
                  "formula": f"[{source['name']}/{source_name}]"}
                 for index, (name, source_name) in enumerate([
                     ("Opportunity ID", "Opportunity ID"), ("Created Date", "Created Date"),
                     ("Close Date", "Expected Close"), ("Account", "Account"),
                     ("Region", "Region"), ("Segment", "Segment"), ("Rep", "Rep"),
                     ("Stage", "Stage"), ("Probability", "Probability"), ("ACV", "ACV"),
                     ("Forecast Category", "Forecast Category"),
                     ("Weighted ACV", "Weighted ACV"), ("Cycle Days", "Cycle Days")])
             ] + [
                 {"id": "sales_setting", "name": "Setting", "formula": '"AsOfDate"'}]},
            {"id": "salesWithSettings", "kind": "table", "name": "Sales With Settings",
             "source": {"kind": "join", "name": "SalesJoin",
                        "primarySource": {"kind": "table", "elementId": "salesInputBase"},
                        "joins": [{"name": "Settings", "joinType": "left-outer",
                                   "left": {"kind": "table", "elementId": "salesInputBase"},
                                   "right": {"kind": "table", "elementId": "salesSettings"},
                                   "columns": [{"left": "[Setting]", "right": "[Setting]"}]}]},
             "columns": [
                 {"id": "sales_join_" + str(index), "name": name,
                  "formula": f"[SalesJoin/{name}]"}
                 for index, name in enumerate([
                     "Opportunity ID", "Created Date", "Close Date", "Account", "Region",
                     "Segment", "Rep", "Stage", "Probability", "ACV", "Forecast Category",
                     "Weighted ACV", "Cycle Days"])
             ] + [
                 {"id": "sales_join_asof", "name": "AsOfDate", "formula": "[Settings/AsOfDate]"},
                 {"id": "sales_timing", "name": "TimingStatus",
                  "formula": architecture["calculations"]["TimingStatus"].replace(
                      "[AsOfDate]", "[Settings/AsOfDate]")}]},
        ])
        consumers["sales"] = source["consumerSourceElementId"]
        validations["salesSettings"] = {"uniqueKeys": ["Setting"]}
    if "expensePolicy" in inputs:
        spec = inputs["expensePolicy"]
        source = wiring.get("expenseSource") or {}
        if not source.get("elementId") or not source.get("name"):
            raise InputSyncError("approved expense source wiring is required")
        elements.extend([
            {"id": "expensePolicy", "kind": "input-table", "name": spec["name"],
             "inputMode": "edit", "source": {"kind": "empty", "connectionId": connection},
             "columns": _rest_columns(spec, "expense_policy")},
            {"id": "expenseInputBase", "kind": "table", "name": "Expense Input Base",
             "source": {"kind": "table", "elementId": source["elementId"]},
             "columns": [
                 {"id": "expense_" + str(index), "name": name,
                  "formula": f"[{source['name']}/{name}]"}
                 for index, name in enumerate([
                     "Expense ID", "Date", "Employee", "Department", "Category", "Vendor",
                     "Amount", "Receipt Status", "Approval Status", "Card Amount", "Ledger Amount"])]},
            {"id": "expensesWithPolicy", "kind": "table", "name": "Expenses With Policy",
             "source": {"kind": "join", "name": "ExpenseJoin",
                        "primarySource": {"kind": "table", "elementId": "expenseInputBase"},
                        "joins": [{"name": "Policy", "joinType": "left-outer",
                                   "left": {"kind": "table", "elementId": "expenseInputBase"},
                                   "right": {"kind": "table", "elementId": "expensePolicy"},
                                   "columns": [{"left": "[Category]", "right": "[Category]"}]}]},
             "columns": [
                 {"id": "expense_join_" + str(index), "name": name,
                  "formula": f"[ExpenseJoin/{name}]"}
                 for index, name in enumerate([
                     "Expense ID", "Date", "Employee", "Department", "Category", "Vendor",
                     "Amount", "Receipt Status", "Approval Status", "Card Amount", "Ledger Amount"])
             ] + [
                 {"id": "expense_join_limit", "name": "Per-item Limit", "formula": "[Policy/Per-item Limit]"},
                 {"id": "expense_join_receipt", "name": "Receipt Threshold", "formula": "[Policy/Receipt Threshold]"},
                 {"id": "expense_join_owner", "name": "Policy Owner", "formula": "[Policy/Policy Owner]"},
                 {"id": "expense_policy_status", "name": "PolicyMatchStatus",
                  "formula": 'If(IsNull([Policy/Category]), "ERROR: missing policy", "Matched")'},
                 {"id": "expense_recon", "name": "Recon Variance",
                  "formula": "[ExpenseJoin/Card Amount] - [ExpenseJoin/Ledger Amount]"},
                 {"id": "expense_excess", "name": "PolicyExcess",
                  "formula": "If(IsNull([Policy/Category]), Null, Max([ExpenseJoin/Amount] - [Policy/Per-item Limit], 0))"},
                 {"id": "expense_exception", "name": "Exception",
                  "formula": ('If(IsNull([Policy/Category]), Null, '
                              'If(Abs([Recon Variance]) > 0.01, "RECON", '
                              'If([PolicyExcess] > 0, "OVER LIMIT", '
                              'If([ExpenseJoin/Receipt Status] = "Missing", "MISSING RECEIPT", '
                              'If([ExpenseJoin/Approval Status] <> "Approved", "PENDING", "CLEAR")))))')}]},
        ])
        consumers["expense"] = source["consumerSourceElementId"]
        validations["expensePolicy"] = {"uniqueKeys": ["Category"]}
    return {"elements": elements, "consumerSources": consumers,
            "preJoinValidation": validations}
