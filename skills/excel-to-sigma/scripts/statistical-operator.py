#!/usr/bin/env python3
"""Offline, phased operator code for the Statistical Demand Forecast preview.

Generation performs no Sigma call. Execute each phase-1 program separately in
Sigma MCP Code Mode, then capture actual readback before generating phase 2.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
from statistical_forecast import SOURCE_SHA256  # noqa: E402
from statistical_runtime import CANDIDATE_REQUIRED, emit_preview_program  # noqa: E402

COUNTS = {"history": 108, "parameters": 1, "targets": 12}
TABLES = {
    "history": ("Statistical History (editable)", ("Month", "SKU", "Demand"),
                ("datetime", "text", "number"), (1, 1, 24, 12)),
    "parameters": ("Statistical Parameters (editable)",
                   ("Scope", "SelectedSKU", "Seasonality", "ServiceLevel", "LeadTime"),
                   ("text", "text", "number", "number", "number"), (1, 15, 24, 8)),
    "targets": ("Statistical Target Months (editable)", ("Period", "Month"),
                ("number", "datetime"), (1, 25, 24, 10)),
}


def source_packet(path):
    packet = json.loads(Path(path).read_text())
    if packet.get("sourceHash") != SOURCE_SHA256 or set(packet.get("roles", {})) != set(COUNTS):
        raise ValueError("SOURCE_PACKET_MISMATCH")
    for role, expected in COUNTS.items():
        spec = packet["roles"][role]
        if (len(spec.get("rows", [])) != expected
                or list(spec.get("columns", {})) != list(TABLES[role][1])
                or list(spec.get("naturalKey", [])) != (["Month", "SKU"] if role == "history" else ["Scope"] if role == "parameters" else ["Period"])):
            raise ValueError("SOURCE_PACKET_ROLE_MISMATCH")
    return packet


def phase1_program(role, page_id, connection_id):
    if role not in TABLES or not page_id or not connection_id:
        raise ValueError("OBSERVED_PAGE_AND_WRITE_CONNECTION_REQUIRED")
    name, columns, types, frame = TABLES[role]
    lines = [f"# sourceHash={SOURCE_SHA256}", f"# role={role}",
             "# CREATE ONLY: query first and skip if this role already exists.",
             f"page = sigma.workbook.get_page({page_id!r})",
             "if not isinstance(page, Page): raise RuntimeError('expected current page')",
             f"created = page.add_input_table({connection_id!r})",
             f"created.set_name({name!r})", "created.set_edit_mode('draft')",
             f"created.set_frame({', '.join(map(str, frame))})"]
    for index, (column, kind) in enumerate(zip(columns, types)):
        lines.append(f"created_c{index} = created.add_editable_column({column!r}, {kind!r})")
    lines += ["created.add_system_columns(['row_id'])",
              "print('role=' + " + repr(role) + ")",
              "print('title=' + " + repr(name) + ")",
              "print('inputTableId=' + created.id())",
              *[f"print('column:{column}=' + created_c{index}.id())"
                for index, column in enumerate(columns)],
              "print('Read back Snowpark identifiers and Row.as_dict keys before phase 2')"]
    return "\n".join(lines) + "\n"


def phase1_continue_parameters(page_id, element_id):
    """Continue a readback-proven partial parameters table without creating one."""
    if not page_id or not element_id:
        raise ValueError("OBSERVED_PARTIAL_PARAMETERS_REQUIRED")
    name, columns, types, frame = TABLES["parameters"]
    lines = [f"# sourceHash={SOURCE_SHA256}",
             "# CONTINUATION ONLY: operator must confirm this element has zero editable columns.",
             f"page = sigma.workbook.get_page({page_id!r})",
             "if not isinstance(page, Page): raise RuntimeError('expected current page')",
             f"created = page.get_element({element_id!r})",
             "if not isinstance(created, InputTable): raise RuntimeError('expected existing parameters input table')",
             f"created.set_frame({', '.join(map(str, frame))})"]
    for index, (column, kind) in enumerate(zip(columns, types)):
        lines.append(f"created_c{index} = created.add_editable_column({column!r}, {kind!r})")
    lines += ["created.add_system_columns(['row_id'])",
              "print('role=parameters')", f"print('title={name}')",
              "print('inputTableId=' + created.id())",
              *[f"print('column:{column}=' + created_c{index}.id())"
                for index, column in enumerate(columns)]]
    return "\n".join(lines) + "\n"


def phase2_program(readback, *, candidate_ets=False):
    if readback.get("sourceHash") != SOURCE_SHA256 or not readback.get("workbookId") or not readback.get("workbookVersion"):
        raise ValueError("OBSERVED_WORKBOOK_READBACK_REQUIRED")
    roles = readback.get("roles")
    if not isinstance(roles, dict) or set(roles) != set(CANDIDATE_REQUIRED):
        raise ValueError("THREE_OBSERVED_ROLES_REQUIRED")
    binding = {}
    element_ids = []
    for role, names in CANDIDATE_REQUIRED.items():
        item = roles[role]
        if (not item.get("elementId") or item.get("rowCount") != COUNTS[role]
                or item.get("naturalKeysVerified") is not True
                or not item.get("readbackSnapshotId")):
            raise ValueError("OBSERVED_ROLE_READBACK_INCOMPLETE:" + role)
        element_ids.append(item["elementId"])
        binding[role] = {"title": item["title"], "columns": item["columns"]}
        if set(binding[role]["columns"]) != set(names):
            raise ValueError("OBSERVED_COLUMNS_MISMATCH:" + role)
    if len(set(element_ids)) != 3:
        raise ValueError("AMBIGUOUS_OBSERVED_ELEMENTS")
    if not candidate_ets:
        raise ValueError("TYPED_PARAMETER_TABLE_REQUIRES_CANDIDATE_PREVIEW")
    return emit_preview_program(binding, candidate_ets=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-packet", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    commands = parser.add_subparsers(dest="phase", required=True)
    create = commands.add_parser("phase1")
    create.add_argument("--page-id", required=True)
    create.add_argument("--write-connection-id", required=True)
    continue_params = commands.add_parser("phase1-continue-parameters")
    continue_params.add_argument("--page-id", required=True)
    continue_params.add_argument("--element-id", required=True)
    bind = commands.add_parser("phase2")
    bind.add_argument("--readback", type=Path, required=True)
    bind.add_argument("--candidate-ets", action="store_true")
    args = parser.parse_args()
    source_packet(args.source_packet)
    args.out.mkdir(parents=True, exist_ok=True)
    if args.phase == "phase1":
        for role in COUNTS:
            (args.out / f"phase1-create-{role}.py").write_text(
                phase1_program(role, args.page_id, args.write_connection_id))
        print("Generated three create-only programs; no Sigma call or seed was made.")
    elif args.phase == "phase1-continue-parameters":
        (args.out / "phase1-continue-parameters.py").write_text(
            phase1_continue_parameters(args.page_id, args.element_id))
        print("Generated existing-table continuation; no Sigma call or seed was made.")
    else:
        program = phase2_program(json.loads(args.readback.read_text()), candidate_ets=args.candidate_ets)
        (args.out / "phase2-python-preview.py").write_text(program["code"])
        (args.out / "phase2-program-metadata.json").write_text(json.dumps(
            {key: value for key, value in program.items() if key != "code"}, indent=2) + "\n")
        print("Generated bound Python preview; no Sigma call was made.")


if __name__ == "__main__":
    main()
