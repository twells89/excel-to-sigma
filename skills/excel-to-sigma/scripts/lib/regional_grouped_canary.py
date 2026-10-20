"""Source-bound, offline Regional01 B4 grouped construction candidate.

This module emits logical columns and a staged execution contract. It never
creates Sigma objects or treats authored names as server-assigned identities.
The formal 144-row Revenue template is proved from the saved source ASTs.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
import copy
import hashlib
import json
from pathlib import Path
import re

from .native_exact import compile_expression, lane_dict
from .native_exact_plan import HelperGraph, input_scalar, digest, _map_helpers
from .source_inputs import validate_source_input_plan

SOURCE_SHA = "4b46bfa6190d1ef12ea35713e3932aa77f19fb964730b4b3bbb0851294219e2c"
CELL = "Summary!B4"
SALES_ROWS = tuple(range(2, 146))
MAX_TABLES = 8
MAX_FORMULA_COLUMNS = 128


def _require(condition, message):
    if not condition:
        raise ValueError(message)


def _relative_ast(ast, row):
    """Replace only same-row source addresses; all other edges remain visible."""
    if isinstance(ast, list):
        return [_relative_ast(value, row) for value in ast]
    if not isinstance(ast, dict):
        return ast
    result = {}
    for key, value in ast.items():
        if key == "cells":
            result[key] = [re.sub(rf"(?<=!)([A-Z]+){row}$", r"\1@ROW", c)
                           for c in value]
        elif key == "activeDomain":
            result[key] = "@ROW:@ROW"
        else:
            result[key] = _relative_ast(value, row)
    return result


def build_candidate(source_path, input_plan, native_plan, oracle):
    source_path = Path(source_path)
    _require(hashlib.sha256(source_path.read_bytes()).hexdigest() == SOURCE_SHA,
             "Regional01 source bytes changed")
    validate_source_input_plan(input_plan, source_path)
    _require(input_plan["sourceSha256"] == native_plan["sourceSha256"] == SOURCE_SHA,
             "source plans do not match source bytes")
    _require(native_plan["inputPlanHash"] == input_plan["inputPlanHash"],
             "native/input plan binding mismatch")
    _require(digest({k: v for k, v in native_plan.items() if k != "nativePlanHash"}) == native_plan["nativePlanHash"],
             "native plan hash is stale")
    sales = input_plan["sharedInputArchitecture"]["inputs"]["sales-forecast"]
    _require(len(sales["rows"]) == len(SALES_ROWS) == 144, "sales row domain changed")
    _require(sales["sourceRowNumbers"] == list(SALES_ROWS), "sales source ordinal changed")
    _require(native_plan["bindings"]["Summary!A4"] == {
        "kind": "input", "featureId": "summary-regions", "rowIndex": 0,
        "column": "Region", "valueType": "text"}, "B4 region criterion lost source input")
    b4 = native_plan["formulas"][CELL]
    _require(b4["status"] == "ready-to-bind" and b4["ast"]["name"] == "SUMIFS",
             "B4 is no longer a bound SUMIFS")
    args = b4["ast"]["arguments"]
    expected_ranges = (("F", "table:tblSalesForecast:#data"),
                       ("A", "table:tblSalesForecast:#data"),
                       ("C", "table:tblSalesForecast:#data"),
                       ("C", "table:tblSalesForecast:#data"))
    for node, (col, kind) in zip((args[0], args[1], args[3], args[5]), expected_ranges):
        _require(node["kind"] == "range" and node["sourceKind"] == kind and
                 node["cells"] == [f"Forecast Input!{col}{row}" for row in SALES_ROWS],
                 "B4 formal table range changed; append semantics must be reviewed")
    _require(args[2] == {"kind": "reference", "cell": "Summary!A4"},
             "B4 region criterion is not live A4")
    _require(b4["originalFormula"] ==
             '=SUMIFS(tblSalesForecast[Revenue],tblSalesForecast[Region],$A4,tblSalesForecast[Month],">="&DATE(2026,1,1),tblSalesForecast[Month],"<"&EDATE(DATE(2026,1,1),3))',
             "B4 date criteria changed")
    formulas = native_plan["formulas"]
    revenue_ids = [f"Forecast Input!F{row}" for row in SALES_ROWS]
    first = formulas[revenue_ids[0]]
    signature = _relative_ast(first["ast"], 2)
    for row, cell in zip(SALES_ROWS, revenue_ids):
        record = formulas[cell]
        _require(record["status"] == "ready-to-bind" and
                 _relative_ast(record["ast"], row) == signature and
                 record["dependencies"] == [f"Forecast Input!D{row}", f"Forecast Input!E{row}"],
                 f"Revenue row template breaks at {cell}")
    _require(oracle["sourceSha256"] == SOURCE_SHA and oracle["sheet"] == "Summary" and
             oracle["address"] == "B4" and oracle["numericValue"] == "1544984.02" and
             oracle["excel"]["valueClass"] == "real", "independent Excel B4 oracle missing")

    # Existing typed compiler provides the five public lanes, including Excel
    # blank/text/logical/error coercion. It is applied once after proving every
    # source Revenue cell has the same row-relative expression.
    graph = HelperGraph(prefix="__regional_r")
    value = compile_expression(first["ast"],
        lambda cell: input_scalar(cell, native_plan["bindings"],
            lambda binding: f'[Sales/{binding["column"]}]'),
        emit=graph.emit, epoch=native_plan["dateEpoch"])
    columns = []
    for index, helper in enumerate(graph.nodes):
        formula = _map_helpers(helper["formula"], graph.prefix,
            lambda match: f'[h{int(match.group()[len(graph.prefix):-2])}]')
        columns.append({"id": f"h{index}", "name": f"h{index}", "formula": formula})
    for lane, formula in lane_dict(value).items():
        formula = _map_helpers(formula, graph.prefix,
            lambda match: f'[h{int(match.group()[len(graph.prefix):-2])}]')
        columns.append({"id": f"revenue_{lane}", "name": f"Revenue {lane.title()}",
                        "formula": formula})
    columns += [
        {"id": "region_key", "name": "Region Key", "formula": "Lower([Sales/Region])"},
        {"id": "year_key", "name": "Year Key", "formula": 'DatePart("year", [Sales/Month])'},
        {"id": "quarter_key", "name": "Quarter Key", "formula": 'DatePart("quarter", [Sales/Month])'},
        {"id": "group_key", "name": "Group Key",
         "formula": 'Concat([Region Key], "|", Text([Year Key]), "|", Text([Quarter Key]))'},
        {"id": "source_ordinal", "name": "Source Ordinal",
         "formula": "[Sales/Source Ordinal]"},
        {"id": "source_key", "name": "Source Key",
         "formula": "[Sales/Source Key]"},
        {"id": "first_error_ordinal", "name": "First Error Ordinal",
         "formula": 'If([Revenue Error] != "", [Source Ordinal], Null)'},
    ]
    _require(all("__regional_r" not in c["formula"] and len(c["formula"]) <= 6000
                 for c in columns), "row helper did not resolve or formula exceeds limit")
    # Grouped result is a flat source derived from the live calculated row
    # table. B4's exact date interval is represented by 2026/Q1. This date
    # equivalence is valid only for non-null civil dates and needs live TZ proof.
    elements = [
        {"id": "regional_sales_rows", "kind": "calculated-table",
         "sourceFeature": "sales-forecast", "columns": columns},
        {"id": "regional_sales_region_quarter", "kind": "grouped-source",
         "sourceElement": "regional_sales_rows",
         "keys": ["Group Key"],
         "calculations": [
             {"id": "revenue_sum", "name": "Revenue Sum", "aggregate": "Sum",
              "column": "Revenue Number"},
             {"id": "first_error_ordinal", "name": "First Error Ordinal", "aggregate": "Min",
              "column": "First Error Ordinal"}],
         "sourceMethod": "Grouping.to_source",
         "keyComponents": ["Region Key", "Year Key", "Quarter Key"]},
        {"id": "regional_summary_b4", "kind": "summary-table",
         "sourceFeature": "summary-regions",
         "columns": [
             {"id": "group_key", "name": "Group Key",
              "formula": 'Concat(Lower([Summary Regions/Region]), "|2026|1")'},
             {"id": "matched_revenue", "name": "Matched Revenue",
              "formula": 'Lookup([Grouped/Revenue Sum], [Group Key], [Grouped/Group Key])'},
             {"id": "matched_error_ordinal", "name": "Matched Error Ordinal",
              "formula": 'Lookup([Grouped/First Error Ordinal], [Group Key], [Grouped/Group Key])'},
             {"id": "b4_number", "name": "B4 Number", "formula": "Coalesce([Matched Revenue], 0)"},
             {"id": "b4_text", "name": "B4 Text", "formula": '""'},
             {"id": "b4_logical", "name": "B4 Logical", "formula": "False"},
             {"id": "b4_kind", "name": "B4 Kind", "formula": '"n"'},
             {"id": "b4_error", "name": "B4 Error",
              "formula": 'If(IsNull([Matched Error Ordinal]), "", Lookup([Rows/Revenue Error], [Matched Error Ordinal], [Rows/Source Ordinal]))'}]}]
    formula_columns = sum(len(e.get("columns", [])) + len(e.get("calculations", []))
                          for e in elements)
    _require(len(elements) <= MAX_TABLES and formula_columns <= MAX_FORMULA_COLUMNS,
             "grouped canary construction budget exceeded")
    source_mapping = {cell: {"targetElement": "regional_sales_rows",
                             "laneColumnIds": {lane: f"revenue_{lane}" for lane in
                                               ("number", "text", "logical", "kind", "error")},
                             "sourceOrdinal": row - 1,
                             "sourceKey": f"Forecast Input!row:{row}"}
                      for row, cell in zip(SALES_ROWS, revenue_ids)}
    source_mapping[CELL] = {"targetElement": "regional_summary_b4", "sourceRow": 0,
                            "laneColumnIds": {lane: f"b4_{lane}" for lane in
                                              ("number", "text", "logical", "kind", "error")}}
    return {"status": "OFFLINE_GROUPED_CANARY_PENDING_STAGING", "runtimeStatus": "UNVERIFIED_CANDIDATE",
            "sourceSha256": SOURCE_SHA, "inputPlanHash": input_plan["inputPlanHash"],
            "nativePlanHash": native_plan["nativePlanHash"], "sourceFormulaSha256": b4["formulaSha256"],
            "sourceCell": CELL, "excelOracle": {"kind": "number", "number": "1544984.02"},
            "sourceRangeAppendPolicy": "formal-table-data-body; includes appended sales rows",
            "sourceRowCount": 144, "sourceMappings": source_mapping,
            "sourceDependencies": b4["dependencies"], "sourceLayer": b4["layer"],
            "elements": elements, "constructionBudget": {"tables": len(elements),
                                                    "formulaColumns": formula_columns,
                                                    "maxTables": MAX_TABLES,
                                                    "maxFormulaColumns": MAX_FORMULA_COLUMNS},
            "qualificationsRequired": [
                "Persist Source Ordinal and Source Key with all 144 ordinary sales rows; observe actual row IDs.",
                "Prove deterministic Source Ordinal assignment and first-error order for appended rows.",
                "Prove grouped DatePart year/quarter matches Excel >=2026-01-01 and <2026-04-01 under date and timezone edits.",
                "Prove Sigma grouped Sum/Min, flat grouping source, three-key join/lookup, and zero for no match.",
                "Qualify Excel case-insensitive/wildcard/blank/error SUMIFS criteria before accepting arbitrary A4 edits.",
                "Use assigned IDs and strict full-contents REST formula patch/readback if SDK nested math serialization changes text.",
                "Query final five B4 lanes after baseline, edits, reopen and restore; source B4 must remain mapped."],
            "unsupported": ["full Regional01 migration", "other 51 SUMIFS cells",
                            "four source VLOOKUP NAME errors", "charts", "production promotion"]}


def source_rows(input_plan):
    """Independent decimal fixture for comparing an emitted candidate baseline."""
    rows = input_plan["sharedInputArchitecture"]["inputs"]["sales-forecast"]["rows"]
    return [{**row, "Source Ordinal": i + 1,
             "Source Key": f"Forecast Input!row:{i+2}"} for i, row in enumerate(rows)]


def reference_b4(rows, region):
    total = Decimal(0)
    for row in rows:
        month = date.fromisoformat(row["Month"][:10]) if row["Month"] is not None else None
        if (isinstance(row["Region"], str) and row["Region"].casefold() == region.casefold()
                and month is not None and date(2026, 1, 1) <= month < date(2026, 4, 1)):
            units = Decimal(str(row["Units"] or 0))
            price = Decimal(str(row["Avg Price"] or 0))
            total += units * price
    return total


def bind_assigned_formula(formula, *, namespaces, local_columns):
    """Bind logical bracket references only after IDs/names come from readback.

    `namespaces` maps Sales/Rows/Grouped/Summary Regions to readback element
    name and exact logical-to-actual column names. Literal quoted text, such as
    a regex character class, is never scanned as a reference.
    """
    result = []
    index = 0
    while index < len(formula):
        if formula[index] == '"':
            end = index + 1
            while end < len(formula):
                if formula[end] == '\\':
                    end += 2
                elif formula[end] == '"':
                    end += 1
                    break
                else:
                    end += 1
            _require(end <= len(formula) and formula[end - 1] == '"', "unterminated formula text")
            result.append(formula[index:end]); index = end
        elif formula[index] == '[':
            end = formula.find(']', index + 1)
            _require(end != -1, "unterminated formula reference")
            key = formula[index + 1:end]
            if '/' in key:
                namespace, column = key.split('/', 1)
                binding = namespaces.get(namespace)
                _require(binding is not None and column in binding.get("columns", {}),
                         f"unassigned native reference [{key}]")
                result.append(f'[{binding["name"]}/{binding["columns"][column]}]')
            else:
                _require(key in local_columns, f"unassigned local column [{key}]")
                result.append(f'[{local_columns[key]}]')
            index = end + 1
        else:
            result.append(formula[index]); index += 1
    bound = ''.join(result)
    _require(len(bound) <= 6000, "assigned formula exceeds 6000 characters")
    return bound


def prepare_scaffold_program(candidate, *, element_kind, page_id, source_element_id,
                             grid_row):
    """One documented SDK table scaffold; print assigned IDs for readback.

    The caller supplies observed page/source IDs. Grouping and the flat grouped
    table require their own qualification phase after the row formulas patch.
    """
    _require(element_kind in ("calculated-table", "summary-table"), "unsupported scaffold phase")
    _require(all(isinstance(x, str) and x and not x.startswith("@")
                 for x in (page_id, source_element_id)), "observed page/source IDs required")
    _require(isinstance(grid_row, int) and grid_row >= 1, "positive placement row required")
    logical = next(e for e in candidate["elements"] if e["kind"] == element_kind)
    name = "Regional Sales Rows" if element_kind == "calculated-table" else "Regional Summary B4"
    lines = [
        "# Structural placeholder only. Read back all assigned IDs before REST formulas.",
        f'p = sigma.workbook.get_page({json_quote(page_id)})',
        f't = p.add_table(sigma.data.source.element({json_quote(source_element_id)}))',
        f't.set_name({json_quote(name)})',
        f't.set_frame(1, {grid_row}, 24, 12)',
        "inherited = t.list_columns()", "assigned = []",
    ]
    for column in logical["columns"]:
        lines += [f'c = t.add_column({json_quote(column["name"])}, sigma.formula.const_number(0))',
                  'assigned.append({"logicalColumnId": ' + json_quote(column["id"])
                  + ', "actualColumnId": c.id(), "name": ' + json_quote(column["name"]) + '})']
    lines += ["for c in inherited:", "    c.delete()",
              f'print({{"logicalElementId": {json_quote(logical["id"])}, "actualElementId": t.id(), "name": {json_quote(name)}, "columns": assigned}})']
    return "\n".join(lines) + "\n"


def json_quote(value):
    return json.dumps(value, ensure_ascii=False)


def prepare_contents_patch(candidate, current, receipt, *, element_kind):
    """Prepare an exact, one-element formula patch against fresh scaffold IDs.

    The caller owns REST transport and must use the returned document version as
    an If-Match guard. Grouping creation/calculations are a separate SDK phase.
    """
    _require(candidate["status"] == "OFFLINE_GROUPED_CANARY_PENDING_STAGING", "candidate status changed")
    _require(element_kind in ("calculated-table", "summary-table"), "unsupported patch phase")
    _require(receipt.get("sourceSha256") == candidate["sourceSha256"] and
             receipt.get("documentId") == current.get("workbookId") and
             receipt.get("documentVersion") == current.get("documentVersion"),
             "stale source/document scaffold receipt")
    _require(isinstance(current.get("documentVersion"), int) and
             isinstance(current.get("contents", {}).get("elements"), list),
             "fresh full document readback required")
    logical = next(e for e in candidate["elements"] if e["kind"] == element_kind)
    binding = receipt.get("elements", {}).get(logical["id"], {})
    _require(binding.get("elementId") and binding.get("name") and
             set(binding.get("columns", {})) == {c["id"] for c in logical["columns"]},
             "exact assigned element/column receipt required")
    original = current["contents"]
    updated = copy.deepcopy(original)
    matches = [e for e in updated["elements"] if e.get("id") == binding["elementId"]]
    _require(len(matches) == 1 and matches[0].get("name") == binding["name"],
             "assigned scaffold element changed")
    target = matches[0]
    actual = {c.get("id"): c for c in target.get("columns", [])}
    _require(set(actual) == set(binding["columns"].values()),
             "assigned scaffold columns changed")
    namespaces = receipt.get("namespaces", {})
    local = {c["name"]: actual[binding["columns"][c["id"]]]["name"]
             for c in logical["columns"]}
    changes = []
    for column in logical["columns"]:
        actual_column = actual[binding["columns"][column["id"]]]
        _require(actual_column.get("name") == column["name"] and
                 actual_column.get("formula") == "0", "scaffold placeholder changed")
        bound = bind_assigned_formula(column["formula"], namespaces=namespaces,
                                      local_columns=local)
        actual_column["formula"] = bound
        changes.append({"logicalColumnId": column["id"], "actualColumnId": actual_column["id"],
                        "formula": bound})
    packet = {"status": "REST_FORMULA_PATCH_PENDING_EXECUTION", "sourceSha256": candidate["sourceSha256"],
              "documentId": current["workbookId"], "ifMatchDocumentVersion": current["documentVersion"],
              "elementId": binding["elementId"], "beforeContentsHash": digest(original),
              "afterContentsHash": digest(updated), "contents": updated, "changes": changes}
    return packet


def confirm_contents_patch(packet, readback):
    _require(packet.get("status") == "REST_FORMULA_PATCH_PENDING_EXECUTION" and
             readback.get("workbookId") == packet.get("documentId") and
             isinstance(readback.get("documentVersion"), int) and
             readback["documentVersion"] > packet["ifMatchDocumentVersion"],
             "fresh saved document readback required")
    _require(digest(readback.get("contents")) == packet["afterContentsHash"],
             "saved full contents differ from exact candidate patch")
    return {"status": "CONFIRMED_FORMULA_PATCH_STRUCTURE_ONLY",
            "sourceSha256": packet["sourceSha256"], "documentId": packet["documentId"],
            "documentVersion": readback["documentVersion"], "contentsHash": packet["afterContentsHash"],
            "runtimeStatus": "UNVERIFIED_CANDIDATE"}
