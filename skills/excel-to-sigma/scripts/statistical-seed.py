#!/usr/bin/env python3
"""Offline, resumable source-key reconciliation for the three statistical inputs.

This command never queries or writes Sigma. It emits one bounded SDK write chunk
from a fresh actual source and query receipt. Re-query after that chunk, then
run again with the prior plan. Never replay an emitted write program.
"""
from __future__ import annotations

import argparse
from datetime import date, datetime
import hashlib
import json
import math
import re
from pathlib import Path
import runpy
import sys

LIB = Path(__file__).resolve().parent / "lib"
sys.path.insert(0, str(LIB))
from input_sync import _row_expression, plan_seed  # noqa: E402
from statistical_forecast import SOURCE_SHA256  # noqa: E402

OPERATOR = runpy.run_path(str(Path(__file__).with_name("statistical-operator.py")))
TABLES = OPERATOR["TABLES"]


def _rows_from_receipt(receipt, role, element_id):
    if isinstance(receipt, list) and all(isinstance(row, dict) for row in receipt):
        matches = [item for item in receipt if isinstance(item.get("value"), dict)
                   and item["value"].get("role") == role
                   and item["value"].get("elementId") == element_id]
        if len(matches) == 1 and "value" in matches[0]:
            payload = matches[0]["value"]["result"]["structuredContent"]
            rows = payload.get("rows")
        else:
            rows = receipt
    elif isinstance(receipt, dict):
        rows = receipt.get("rows")
        if rows is None and isinstance(receipt.get("structuredContent"), dict):
            rows = receipt["structuredContent"].get("rows")
    else:
        rows = None
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ValueError("ACTUAL_QUERY_ROWS_REQUIRED")
    return rows


def _normalize(value, kind):
    if kind == "datetime":
        if isinstance(value, datetime):
            value = value.date()
        if isinstance(value, date):
            return value.isoformat()
        if isinstance(value, str):
            return date.fromisoformat(value[:10]).isoformat()
        raise ValueError("INVALID_DATE_VALUE")
    if kind == "number":
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
            raise ValueError("INVALID_NUMERIC_VALUE")
        return float(value)
    if kind == "text":
        if not isinstance(value, str):
            raise ValueError("INVALID_TEXT_VALUE")
        return value
    raise ValueError("UNSUPPORTED_COLUMN_TYPE")


def seed_plan(*, source_packet_path, source_original, actual_source, query_receipt,
              describe_receipt, role, row_id_key, prior_seed=None, max_rows=25):
    packet_hash = OPERATOR["source_packet"](source_packet_path, source_original)
    packet = json.loads(Path(source_packet_path).read_text())
    if role not in TABLES or not isinstance(row_id_key, str) or not row_id_key:
        raise ValueError("ACTUAL_ROLE_AND_ROW_ID_REQUIRED")
    if not 1 <= max_rows <= 2000:
        raise ValueError("INVALID_BATCH_SIZE")
    source = json.loads(Path(actual_source).read_text())
    if type(source.get("documentVersion")) is not int or not source.get("workbookId"):
        raise ValueError("ACTUAL_WORKBOOK_SOURCE_REQUIRED")
    elements = source.get("contents", {}).get("elements", [])
    title, names, types, _frame = TABLES[role]
    matches = [item for item in elements if item.get("name") == title]
    if len(matches) != 1:
        raise ValueError("ACTUAL_INPUT_TABLE_MISSING_OR_AMBIGUOUS")
    element = matches[0]
    if element.get("kind") != "input-table" or element.get("source", {}).get("kind") != "empty":
        raise ValueError("ACTUAL_INPUT_TABLE_KIND_MISMATCH")
    actual_columns = [c for c in element.get("columns", []) if c.get("name") is not None]
    if [(c.get("name"), c.get("type")) for c in actual_columns] != list(zip(names, types)):
        raise ValueError("ACTUAL_INPUT_TABLE_COLUMNS_MISMATCH")
    column_ids = {c["name"]: c["id"] for c in actual_columns}
    describes = json.loads(Path(describe_receipt).read_text())
    matches = [item["value"] for item in describes if isinstance(item.get("value"), dict)
               and item["value"].get("role") == role
               and item["value"].get("elementId") == element["id"]]
    if len(matches) != 1:
        raise ValueError("ACTUAL_DDL_DESCRIBE_REQUIRED")
    ddl = matches[0]["result"]["structuredContent"]["ddl"]
    if (f'-- Workbook ID: "{source["workbookId"]}"' not in ddl
            or f'-- Element ID: "{element["id"]}"' not in ddl
            or any(f'"{column_ids[name]}" {kind}' not in ddl for name, kind in zip(names, types))):
        raise ValueError("ACTUAL_DDL_IDENTITY_MISMATCH")
    row_ids = re.findall(r'^\s+"([^"]+)" text -- "ID" \|', ddl, re.MULTILINE)
    if row_ids != [row_id_key]:
        raise ValueError("ACTUAL_ROW_ID_ALIAS_MISMATCH")
    role_spec = packet["roles"][role]
    source_rows = [{name: _normalize(row[name], kind) for name, kind in role_spec["columns"].items()}
                   for row in role_spec["rows"]]
    receipt_bytes = Path(query_receipt).read_bytes()
    receipt_hash = hashlib.sha256(receipt_bytes).hexdigest()
    raw_rows = _rows_from_receipt(json.loads(receipt_bytes), role, element["id"])
    current_rows = []
    for raw in raw_rows:
        if row_id_key not in raw:
            raise ValueError("ACTUAL_QUERY_ROW_ID_MISSING")
        record = {row_id_key: raw[row_id_key]}
        for name, kind in role_spec["columns"].items():
            found = [key for key in (name, column_ids[name]) if key in raw]
            if len(found) != 1:
                raise ValueError("ACTUAL_QUERY_COLUMN_KEY_MISSING_OR_AMBIGUOUS:" + name)
            record[name] = _normalize(raw[found[0]], kind)
        current_rows.append(record)
    if prior_seed is not None and prior_seed.get("snapshot", {}).get("queryReceiptHash") == receipt_hash:
        raise ValueError("REQUERY_REQUIRED_AFTER_PRIOR_BATCH")
    plan = plan_seed(source_rows, current_rows, role_spec["naturalKey"], row_id_key,
                     prior_seed=prior_seed,
                     snapshot={"queryReceiptHash": receipt_hash,
                               "workbookId": source["workbookId"],
                               "workbookVersion": source["documentVersion"],
                               "elementId": element["id"], "sourcePacketHash": packet_hash},
                     column_types=role_spec["columns"])
    if prior_seed is not None and prior_seed.get("sourceHash") != plan["sourceHash"]:
        raise ValueError("PRIOR_SEED_SOURCE_CHANGED")
    actions = [("update", item) for item in plan["updates"]] + [("insert", item) for item in plan["inserts"]]
    if plan["conflicts"]:
        return plan, None
    if not actions:
        return plan, None
    chunk = actions[:max_rows]
    checkpoint = hashlib.sha256(json.dumps({"planHash": plan["planHash"],
                                             "snapshot": plan["snapshot"],
                                             "actions": chunk}, sort_keys=True).encode()).hexdigest()
    page_id = next((p["id"] for p in source["contents"].get("pages", [])
                    if p.get("name") == "Statistical Demand Forecast — candidate preview"), None)
    if not page_id:
        raise ValueError("ACTUAL_INPUT_PAGE_MISSING")
    lines = ["from datetime import datetime", f"# sourceHash={SOURCE_SHA256}",
             f"# sourcePacketHash={packet_hash}", f"# queryReceiptHash={receipt_hash}",
             f"# expectedWorkbookId={source['workbookId']}",
             f"# expectedWorkbookVersion={source['documentVersion']}",
             f"# checkpointId={checkpoint}",
             "# Repeat full source/version and query before execution; SDK has no row CAS.",
             f"page = sigma.workbook.get_page({page_id!r})",
             "if not isinstance(page, Page): raise RuntimeError('expected input page')",
             f"table = page.get_element({element['id']!r})",
             "if not isinstance(table, InputTable): raise RuntimeError('expected observed input table')"]
    variables = {}
    for index, name in enumerate(names):
        var = f"c{index}"; variables[name] = var
        lines += [f"{var} = table.get_column({column_ids[name]!r})",
                  f"if not isinstance({var}, LiteralEditableColumn): raise RuntimeError('expected editable {name}')"]
    updates = [item for kind, item in chunk if kind == "update"]
    inserts = [item for kind, item in chunk if kind == "insert"]
    if updates:
        lines.append("table.update_rows([")
        lines.extend(f"    ({item['rowId']!r}, {_row_expression(item['changes'], variables, role_spec['columns'])}),"
                     for item in updates)
        lines.append("])")
    if inserts:
        lines.append("row_ids = table.upsert_rows([")
        lines.extend(f"    {_row_expression(item['source'], variables, role_spec['columns'])},"
                     for item in inserts)
        lines.append("])")
        lines.append("print('insertedRowIds=' + repr(row_ids))")
    lines.append(f"print('checkpointId={checkpoint}')")
    return plan, {"program": "\n".join(lines) + "\n", "checkpointId": checkpoint,
                  "actionCount": len(chunk), "remainingCount": len(actions) - len(chunk),
                  "snapshot": plan["snapshot"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("source-packet", "source-original", "actual-source", "query-receipt", "describe-receipt", "out"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--role", choices=tuple(TABLES), required=True)
    parser.add_argument("--row-id-key", required=True)
    parser.add_argument("--prior-seed", type=Path)
    parser.add_argument("--max-rows", type=int, default=25)
    args = parser.parse_args()
    prior = json.loads(args.prior_seed.read_text()) if args.prior_seed else None
    plan, chunk = seed_plan(source_packet_path=args.source_packet,
                            source_original=args.source_original,
                            actual_source=args.actual_source,
                            query_receipt=args.query_receipt,
                            describe_receipt=args.describe_receipt, role=args.role,
                            row_id_key=args.row_id_key, prior_seed=prior,
                            max_rows=args.max_rows)
    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / f"seed-plan-{args.role}.json").write_text(json.dumps(plan, indent=2) + "\n")
    if chunk:
        (args.out / f"seed-next-{args.role}.py").write_text(chunk["program"])
        (args.out / f"seed-next-{args.role}-metadata.json").write_text(
            json.dumps({key: value for key, value in chunk.items() if key != "program"}, indent=2) + "\n")
    print(json.dumps({"role": args.role, "inserts": len(plan["inserts"]),
                      "updates": len(plan["updates"]), "conflicts": len(plan["conflicts"]),
                      "orphans": len(plan["orphans"]), "nextActionCount": chunk["actionCount"] if chunk else 0}))


if __name__ == "__main__":
    main()
