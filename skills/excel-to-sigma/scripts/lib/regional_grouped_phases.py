"""Guarded, resumable structural and formula phases for Regional outputs.

SDK programs only add documented native tables/groupings. REST packets carry
the complete contents plus an If-Match document version; callers must GET and
confirm each phase before continuing. No transport is performed here.
"""
from __future__ import annotations

import copy
import json

from .native_exact_plan import digest
from .regional_grouped_output import bind_full_output_formulas


def _require(condition, message):
    if not condition:
        raise ValueError("BLOCKED Regional construction: " + message)


def _element(current, element_id):
    _require(isinstance(current.get("documentVersion"), int) and
             isinstance(current.get("contents", {}).get("elements"), list),
             "fresh full-document readback/version required")
    found = [e for e in current["contents"]["elements"] if e.get("id") == element_id]
    _require(len(found) == 1 and isinstance(found[0].get("name"), str),
             "actual element ID absent or ambiguous")
    return found[0]


def _page(current, page_id):
    _require(any(p.get("id") == page_id for p in current["contents"].get("pages", [])),
             "actual page ID absent")


def _fresh_inputs(graph, current, assigned_rows):
    _require(assigned_rows.get("status") == "FULL_ASSIGNED_ROWS_VERIFIED" and
             assigned_rows.get("documentId") == current.get("workbookId") and
             assigned_rows.get("documentVersion") == current.get("documentVersion") and
             assigned_rows.get("contentsHash") == digest(current.get("contents")) and
             assigned_rows.get("sourceSha256") == graph.get("sourceSha256") and
             assigned_rows.get("inputPlanHash") == graph.get("inputPlanHash") and
             len(assigned_rows.get("features", {})) == 5,
             "fresh five-input query/physical-ID reconciliation required")


def prepare_output_scaffold_program(graph, current, assigned_rows, *, output_element, source_element_id,
                                    page_id, grid_row):
    """One output table scaffold using observed source/page IDs."""
    _require(graph.get("status") == "FULL_REGIONAL_FORMULAS_PENDING_ASSIGNED_BINDINGS" and
             output_element in ("region-summary", "product-summary"),
             "output graph/phase changed")
    _fresh_inputs(graph, current, assigned_rows)
    source = _element(current, source_element_id)
    _page(current, page_id)
    _require(source.get("kind") == "input-table" and isinstance(grid_row, int) and grid_row >= 1,
             "ordinary source input and placement required")
    _require(not any(e.get("name") == {"region-summary": "Regional Summary",
                                         "product-summary": "Regional Product Months"}[output_element]
                     for e in current["contents"]["elements"]),
             "output scaffold already exists; reconcile prior phase")
    output_name = {"region-summary": "Regional Summary", "product-summary": "Regional Product Months"}[output_element]
    lines = [
        "# STAGING ONLY. Inspect assigned IDs and input rows before formula patch.",
        f'p = sigma.workbook.get_page({json.dumps(page_id)})',
        f't = p.add_table(sigma.data.source.element({json.dumps(source_element_id)}))',
        f't.set_name({json.dumps(output_name)})',
        f't.set_frame(1, {grid_row}, 24, 12)',
        'inherited = t.list_columns()',
        'assigned = []',
    ]
    for name in graph["formulaColumns"][output_element]:
        lines.append(f'c = t.add_column({json.dumps(name)}, sigma.formula.const_num(0))')
        lines.append(f'assigned.append({{"logicalName": {json.dumps(name)}, "actualId": c.id(), "name": c.name()}})')
    lines.extend(['for c in inherited:', '    c.delete()',
                  f'print({{"logicalElement": {json.dumps(output_element)}, "elementId": t.id(), "columns": assigned}})'])
    return {"status": "REGIONAL_OUTPUT_SCAFFOLD_PENDING_EXECUTION", "sourceSha256": graph["sourceSha256"],
            "documentId": current["workbookId"], "documentVersion": current["documentVersion"],
            "contentsHash": digest(current["contents"]), "logicalElement": output_element,
            "sourceElementId": source_element_id, "program": "\n".join(lines) + "\n"}


def _validate_namespaces(current, namespaces):
    _require(isinstance(namespaces, dict), "assigned namespace receipt required")
    for alias, binding in namespaces.items():
        _require(isinstance(binding, dict) and isinstance(binding.get("elementId"), str) and
                 isinstance(binding.get("name"), str) and isinstance(binding.get("columns"), dict),
                 "assigned namespace incomplete: " + alias)
        actual = _element(current, binding["elementId"])
        _require(actual["name"] == binding["name"], "assigned namespace name changed: " + alias)
        by_id = {c.get("id"): c for c in actual.get("columns", [])}
        _require(len(by_id) == len(actual.get("columns", [])), "ambiguous namespace column IDs: " + alias)
        for logical, column in binding["columns"].items():
            _require(isinstance(column, dict) and isinstance(column.get("id"), str) and
                     isinstance(column.get("name"), str), "incomplete column binding: " + alias + "/" + logical)
            actual_column = by_id.get(column["id"])
            _require(actual_column is not None, "assigned column absent: " + alias + "/" + logical)
            if actual_column.get("name") is None:
                if (column["id"] == "ID" and column["name"] == "ID" and
                        actual.get("kind") == "input-table"):
                    continue
                ddl = binding.get("nativeDescribe", {}).get("structuredContent", {}).get("ddl", "")
                _require(f'"{column["id"]}"' in ddl and f'"{column["name"]}"' in ddl and
                         actual.get("source", {}).get("kind") == "table" and
                         actual.get("source", {}).get("groupingId"),
                         "unnamed grouped column needs observed native DDL: " + alias + "/" + logical)
            else:
                _require(actual_column["name"] == column["name"],
                         "assigned column name changed: " + alias + "/" + logical)


def _validate_namespace_roles(current, namespaces, assigned_rows):
    input_aliases = {"Regions": "regions", "Summary Regions": "summary-regions",
                     "Summary Months": "summary-months", "Summary Products": "summary-products"}
    for alias, feature in input_aliases.items():
        if alias in namespaces:
            binding = namespaces[alias]
            receipt = assigned_rows["features"][feature]
            _require(binding["elementId"] == receipt["elementId"] and
                     binding["name"] == receipt["name"],
                     "output namespace is not reconciled editable input: " + alias)
    if "Rows" in namespaces:
        rows = _element(current, namespaces["Rows"]["elementId"])
        _require(rows.get("kind") == "table" and
                 rows.get("source") == {"kind": "table", "elementId": assigned_rows["features"]["sales-forecast"]["elementId"]},
                 "shared Revenue row source lineage changed")
    if "Grouped" in namespaces:
        grouped = _element(current, namespaces["Grouped"]["elementId"])
        _require("Rows" in namespaces and grouped.get("kind") == "table" and
                 grouped.get("source", {}).get("elementId") == namespaces["Rows"]["elementId"] and
                 grouped.get("source", {}).get("groupingId"),
                 "quarter grouped source lineage changed")
    if "Product Grouped" in namespaces:
        product = _element(current, namespaces["Product Grouped"]["elementId"])
        parent = _element(current, product.get("source", {}).get("elementId"))
        _require("Rows" in namespaces and product.get("kind") == parent.get("kind") == "table" and
                 product.get("source", {}).get("groupingId") and
                 parent.get("source", {}).get("elementId") == namespaces["Rows"]["elementId"],
                 "product/month grouped source lineage changed")
    if "Region Summary" in namespaces:
        region = _element(current, namespaces["Region Summary"]["elementId"])
        _require(region.get("kind") == "table" and
                 region.get("source") == {"kind": "table", "elementId": assigned_rows["features"]["summary-regions"]["elementId"]},
                 "region total source lineage changed")
    if "Product Summary" in namespaces:
        product_summary = _element(current, namespaces["Product Summary"]["elementId"])
        _require(product_summary.get("kind") == "table" and
                 product_summary.get("source") == {"kind": "table", "elementId": assigned_rows["features"]["summary-months"]["elementId"]},
                 "product summary source lineage changed")
    if "Totals" in namespaces:
        totals = _element(current, namespaces["Totals"]["elementId"])
        _require("Region Summary" in namespaces and totals.get("kind") == "table" and
                 totals.get("source", {}).get("elementId") == namespaces["Region Summary"]["elementId"] and
                 totals.get("source", {}).get("groupingId"),
                 "grand totals grouped source lineage changed")


def prepare_output_contents_patch(graph, current, scaffold, assigned_rows, *, output_element,
                                  output_element_id, source_element_id, assigned_columns,
                                  namespaces):
    """Exact full-document formula patch for all columns in one output table."""
    _require(graph.get("status") == "FULL_REGIONAL_FORMULAS_PENDING_ASSIGNED_BINDINGS" and
             output_element in ("region-summary", "product-summary", "totals") and
             scaffold.get("documentId") == current.get("workbookId") and
             scaffold.get("sourceSha256") == graph["sourceSha256"] and
             scaffold.get("logicalElement") == output_element and
             scaffold.get("sourceElementId") == source_element_id,
             "source/scaffold phase identity differs")
    _fresh_inputs(graph, current, assigned_rows)
    output = _element(current, output_element_id)
    source = _element(current, source_element_id)
    _require(output.get("kind") == "table" and
             output.get("source", {}).get("elementId") == source_element_id and
             source.get("kind") in ("input-table", "table"),
             "output table/source lineage differs")
    names = graph["formulaColumns"][output_element]
    _require(set(assigned_columns) == set(names), "complete assigned output columns required")
    actual_by_id = {c.get("id"): c for c in output.get("columns", [])}
    for name, col_id in assigned_columns.items():
        _require(isinstance(col_id, str) and actual_by_id.get(col_id, {}).get("name") == name and
                 actual_by_id[col_id].get("formula") == "0", "fresh placeholder/column ID differs: " + name)
    _require(len(set(assigned_columns.values())) == len(assigned_columns),
             "assigned output IDs are not unique")
    _validate_namespaces(current, namespaces)
    _validate_namespace_roles(current, namespaces, assigned_rows)
    bound = bind_full_output_formulas(graph, namespaces, element=output_element)
    updated = copy.deepcopy(current["contents"])
    target = next(e for e in updated["elements"] if e.get("id") == output_element_id)
    for column in target["columns"]:
        if column["id"] in assigned_columns.values():
            column["formula"] = bound[column["name"]]
    return {"status": "REGIONAL_OUTPUT_FORMULAS_PENDING_REST", "sourceSha256": graph["sourceSha256"],
            "documentId": current["workbookId"], "ifMatchDocumentVersion": current["documentVersion"],
            "beforeContentsHash": digest(current["contents"]), "expectedContentsHash": digest(updated),
            "logicalElement": output_element, "outputElementId": output_element_id,
            "assignedColumns": assigned_columns, "contents": updated}


def confirm_output_contents_patch(packet, after):
    _require(packet.get("status") == "REGIONAL_OUTPUT_FORMULAS_PENDING_REST" and
             after.get("workbookId") == packet.get("documentId") and
             after.get("documentVersion") == packet.get("ifMatchDocumentVersion", -1) + 1 and
             digest(after.get("contents")) == packet.get("expectedContentsHash"),
             "saved output formula patch/readback differs")
    return {"status": "REGIONAL_OUTPUT_FORMULAS_CONFIRMED",
            "logicalElement": packet["logicalElement"], "outputElementId": packet["outputElementId"],
            "documentId": after["workbookId"], "documentVersion": after["documentVersion"],
            "contentsHash": digest(after["contents"]), "sourceSha256": packet["sourceSha256"]}


def prepare_totals_grouping_program(graph, current, region_confirmation, assigned_rows, *,
                                    region_element_id, page_id, grid_row):
    """Group four live region records into five source totals and a label row."""
    _require(graph.get("status") == "FULL_REGIONAL_FORMULAS_PENDING_ASSIGNED_BINDINGS" and
             region_confirmation.get("status") == "REGIONAL_OUTPUT_FORMULAS_CONFIRMED" and
             region_confirmation.get("logicalElement") == "region-summary" and
             region_confirmation.get("outputElementId") == region_element_id and
             region_confirmation.get("documentId") == current.get("workbookId") and
             region_confirmation.get("documentVersion") == current.get("documentVersion") and
             region_confirmation.get("contentsHash") == digest(current.get("contents")),
             "confirmed region-summary readback required")
    _fresh_inputs(graph, current, assigned_rows)
    region = _element(current, region_element_id)
    _page(current, page_id)
    _require(region.get("kind") == "table" and isinstance(grid_row, int) and grid_row >= 1,
             "actual region table/placement required")
    expected_names = set(graph["formulaColumns"]["region-summary"])
    by_name = {c.get("name"): c for c in region.get("columns", [])}
    _require(expected_names <= set(by_name) and
             all(by_name[name].get("formula") not in (None, "0") for name in expected_names),
             "all saved source region formulas required")
    expected = [{"id": c["id"], "name": c.get("name"), "formula": c.get("formula")}
                for c in region["columns"]]
    lines = [
        "# STAGING ONLY. Read back grouping/flat IDs, query five totals, then patch label/error lanes.",
        "from sigma.levelTable import LevelTable",
        f'p = sigma.workbook.get_page({json.dumps(page_id)})',
        f'region = p.get_element({json.dumps(region_element_id)})',
        'if not isinstance(region, LevelTable): raise RuntimeError("actual region table changed")',
        f'if region.id() != {json.dumps(region_element_id)} or region.name() != {json.dumps(region["name"])}: raise RuntimeError("region identity changed")',
        'actual = []',
        'for c in region.list_columns():',
        '    actual.append({"id": c.id(), "name": c.name(), "formula": c.formula()})',
        f'if actual != {json.dumps(expected)}: raise RuntimeError("region source columns/formulas changed")',
        f'key = region.get_column({json.dumps(by_name["Total Key"]["id"])})',
        'grouping = region.set_groupings([[key]])[0]',
        'calcs = []',
    ]
    for spec in graph["totalGrouping"]["calculations"].values():
        lines.append(f'source = region.get_column({json.dumps(by_name[spec["source"]]["id"])})')
        lines.append(f'calc = grouping.set_calculation({json.dumps(spec["sum"])}, sigma.formula.sum(source))')
        lines.append(f'calcs.append({{"name": {json.dumps(spec["sum"])}, "id": calc.id()}})')
        lines.append(f'error_source = region.get_column({json.dumps(by_name[spec["errorSource"]]["id"])})')
        lines.append(f'calc = grouping.set_calculation({json.dumps(spec["errorOrdinal"])}, sigma.formula.min(error_source))')
        lines.append(f'calcs.append({{"name": {json.dumps(spec["errorOrdinal"])}, "id": calc.id()}})')
    lines.extend(['flat = p.add_table(grouping.to_source())',
                  'flat.set_name("Regional Totals")',
                  f'flat.set_frame(1, {grid_row}, 24, 12)',
                  'assigned = []'])
    for name in graph["formulaColumns"]["totals"]:
        lines.append(f'c = flat.add_column({json.dumps(name)}, sigma.formula.const_num(0))')
        lines.append(f'assigned.append({{"logicalName": {json.dumps(name)}, "actualId": c.id(), "name": c.name()}})')
    lines.append('print({"regionElementId": region.id(), "totalsElementId": flat.id(), "calculations": calcs, "outputColumns": assigned})')
    return {"status": "REGIONAL_TOTALS_GROUPING_PENDING_EXECUTION", "sourceSha256": graph["sourceSha256"],
            "documentId": current["workbookId"], "documentVersion": current["documentVersion"],
            "contentsHash": digest(current["contents"]), "logicalElement": "totals",
            "sourceElementId": region_element_id, "program": "\n".join(lines) + "\n"}


def bind_full_214_source_manifest(graph, current, assigned_rows, namespaces):
    """Freeze every source formula against read-back target element/column IDs.

    This is a construction receipt, not formula or runtime qualification. The
    grouped totals row has no source input physical row and is tied to its
    source DAG and the assigned grouped element instead.
    """
    _require(graph.get("status") == "FULL_REGIONAL_FORMULAS_PENDING_ASSIGNED_BINDINGS" and
             graph.get("sourceFormulaCount") == 214, "full source formula graph required")
    _fresh_inputs(graph, current, assigned_rows)
    _validate_namespaces(current, namespaces)
    _validate_namespace_roles(current, namespaces, assigned_rows)
    roles = {"sales-rows": "Rows", "region-summary": "Region Summary",
             "product-summary": "Product Summary", "totals": "Totals"}
    for alias in roles.values():
        _require(alias in namespaces, "actual assigned output namespace missing: " + alias)
    expected_counts = {"sales-rows": 144, "region-summary": 28,
                       "product-summary": 36, "totals": 6}
    manifest = {}
    for cell, mapping in graph["sourceMappings"].items():
        role = mapping["targetElement"]
        _require(role in roles, "source target role unknown: " + cell)
        namespace = namespaces[roles[role]]
        lanes = {}
        for lane, column_name in mapping["laneColumnNames"].items():
            column = namespace["columns"].get(column_name)
            _require(isinstance(column, dict) and column.get("id"),
                     "assigned lane column missing: " + cell + "/" + lane)
            lanes[lane] = column["id"]
        _require(set(lanes) == {"number", "text", "logical", "kind", "error"},
                 "five source lanes incomplete: " + cell)
        source_feature = {"sales-rows": "sales-forecast",
                          "region-summary": "summary-regions",
                          "product-summary": "summary-months"}.get(role)
        source_row_id = None
        if source_feature:
            row_index = (mapping["sourceOrdinal"] - 1 if role == "sales-rows"
                         else mapping["targetRowIndex"])
            source_row_id = assigned_rows["features"][source_feature]["rows"].get(str(row_index))
            _require(isinstance(source_row_id, str) and source_row_id,
                     "actual source row ID missing: " + cell)
        manifest[cell] = {"sourceFormulaSha256": mapping["sourceFormulaSha256"],
                          "sourceAstHash": mapping["sourceAstHash"],
                          "targetElementId": namespace["elementId"],
                          "targetElementName": namespace["name"],
                          "targetField": mapping["targetField"],
                          "targetRowIndex": mapping.get("targetRowIndex"),
                          "sourceInputFeature": source_feature,
                          "sourceInputPhysicalRowId": source_row_id,
                          "sourceDependencies": mapping["sourceDependencies"],
                          "laneColumnIds": lanes,
                          "runtimeStatus": "PENDING_QUALIFICATION"}
    counts = {role: sum(m["targetElement"] == role for m in graph["sourceMappings"].values())
              for role in roles}
    _require(counts == expected_counts and len(manifest) == 214,
             "assigned source formula coverage differs")
    return {"status": "FULL_REGIONAL_214_ASSIGNED_PENDING_RUNTIME",
            "documentId": current["workbookId"],
            "documentVersion": current["documentVersion"],
            "contentsHash": digest(current["contents"]),
            "sourceSha256": graph["sourceSha256"],
            "inputPlanHash": graph["inputPlanHash"],
            "nativePlanHash": graph["nativePlanHash"],
            "sourceFormulaCount": 214,
            "roleCounts": counts,
            "sourceMappings": manifest,
            "requiredRuntimeChecks": ["all 214 source values and five lanes",
                "source Excel VLOOKUP capture reconciliation",
                "five editable input change propagation and restoration",
                "appended record and source ordinal behavior",
                "two chart bindings and saved/reopened roles"]}
