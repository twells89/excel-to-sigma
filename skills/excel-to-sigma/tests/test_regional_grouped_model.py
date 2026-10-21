"""The grouped planner must account for the whole source, without scalar shards."""
from collections import Counter
import copy
from decimal import Decimal
import gzip
import json
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.regional_grouped_model import (build_grouped_model, evaluate_sumifs, evaluate_lookup,
    prepare_product_key_scaffold, prepare_product_key_patch,
    confirm_product_key_patch, prepare_product_rows_scaffold_program,
    prepare_product_error_patch, confirm_product_error_patch,
    prepare_product_grouping_program)
from lib.regional_grouped_output import (build_full_output_graph, bind_full_output_formulas,
    validate_full_assigned_rows)
from lib.regional_grouped_phases import (prepare_output_scaffold_program,
    prepare_output_contents_patch, confirm_output_contents_patch,
    prepare_totals_grouping_program, bind_full_214_source_manifest)
from native_graph_reference import NativeGraphReference

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "regional-grouped"
INPUT = FIXTURES / "input-plan.json.gz"
NATIVE = FIXTURES / "native-plan.json.gz"
SAVED = FIXTURES / "saved-b4-readback.json"
CANARY = FIXTURES / "canary.json.gz"


def synthetic_assigned_rows(input_plan):
    """Generate query-shaped evidence with explicit test-only physical IDs."""
    elements = []
    observed = {}
    for feature, spec in input_plan["sharedInputArchitecture"]["inputs"].items():
        names = list(spec["rows"][0]) + ["Source Key", "Source Ordinal"]
        element_id = "test-only-element-" + feature
        elements.append({"id": element_id, "name": feature, "kind": "input-table",
                         "columns": [{"id": name, "name": name} for name in names] + [{"id": "ID"}]})
        rows = []
        for index, values in enumerate(spec["rows"]):
            source_cells = list(spec["sourceCells"][str(index)].values())
            key = (source_cells[0] if len(source_cells) == 1 and feature not in ("sales-forecast", "summary-regions") else
                   source_cells[0].rsplit("!", 1)[0] + "!row:" + str(spec["sourceRowNumbers"][index]))
            rows.append({**values, "Source Key": key, "Source Ordinal": index + 1,
                         "ID": f"test-only-{feature}-{index}"})
        observed[feature] = {"elementId": element_id, "name": feature,
                             "rowIdColumn": "ID", "queryRows": rows}
    current = {"workbookId": "test-only-full-inputs", "documentVersion": 1,
               "contents": {"elements": elements}}
    return validate_full_assigned_rows(input_plan, current, observed)


class RegionalGroupedModelTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.inputs = json.loads(gzip.decompress(INPUT.read_bytes()))
        cls.native = json.loads(gzip.decompress(NATIVE.read_bytes()))
        cls.model = build_grouped_model(cls.inputs, cls.native)

    def test_all_214_obligations_and_bounded_structure(self):
        model = self.model
        self.assertEqual(set(model["sourceMappings"]), set(self.native["formulas"]))
        self.assertEqual(Counter(x["kind"] for x in model["sourceMappings"].values()),
                         Counter({"record-revenue": 144, "region-quarter": 16,
                                  "product-month": 36, "lookup": 4,
                                  "ratio": 4, "total": 9, "text": 1}))
        self.assertEqual(len(model["elements"]), 7)
        self.assertLessEqual(model["constructionBudget"]["logicalRoleColumns"], 128)
        self.assertEqual(set(model["editableInputs"]), set(self.inputs["sharedInputArchitecture"]["inputs"]))
        self.assertEqual(model["sourceMappings"]["Summary!B4"]["period"],
                         {"start": "2026-01-01", "endExclusive": "2026-04-01"})

    def test_all_52_baselines_and_forward_edits(self):
        sales = self.inputs["sharedInputArchitecture"]["inputs"]["sales-forecast"]["rows"]
        for cell, mapping in self.model["sourceMappings"].items():
            if mapping["kind"] not in ("region-quarter", "product-month"):
                continue
            expected = self.native["formulas"][cell]["localExpected"]
            self.assertEqual(evaluate_sumifs(mapping, sales, self.inputs).quantize(Decimal(".01")),
                             Decimal(str(expected)).quantize(Decimal(".01")), cell)
        changed = [dict(row) for row in sales]
        changed[0]["Units"] += 1
        b4 = self.model["sourceMappings"]["Summary!B4"]
        self.assertEqual(evaluate_sumifs(b4, changed, self.inputs) -
                         evaluate_sumifs(b4, sales, self.inputs), Decimal(str(sales[0]["Avg Price"])))
        inputs = self.inputs["sharedInputArchitecture"]["inputs"]
        for cell, mapping in self.model["sourceMappings"].items():
            if mapping["kind"] not in ("region-quarter", "product-month"):
                continue
            selector = mapping["selector"]
            criterion = inputs[selector["featureId"]]["rows"][selector["rowIndex"]][selector["column"]]
            matches = []
            for index, row in enumerate(sales):
                value = row["Region"] if mapping["kind"] == "region-quarter" else row["Product Line"]
                if value.casefold() != criterion.casefold():
                    continue
                if mapping["kind"] == "region-quarter":
                    if mapping["period"]["start"] <= row["Month"][:10] < mapping["period"]["endExclusive"]:
                        matches.append(index)
                else:
                    month_selector = mapping["period"]["exactMonthSelector"]
                    month = inputs[month_selector["featureId"]]["rows"][month_selector["rowIndex"]][month_selector["column"]]
                    if row["Month"] == month:
                        matches.append(index)
            self.assertTrue(matches, cell)
            edited = [dict(row) for row in sales]
            edited[matches[0]]["Units"] += 1
            self.assertEqual(evaluate_sumifs(mapping, edited, self.inputs) -
                             evaluate_sumifs(mapping, sales, self.inputs),
                             Decimal(str(sales[matches[0]]["Avg Price"])), cell)
            self.assertEqual(evaluate_sumifs(mapping, sales, self.inputs,
                             errors={matches[0] + 1: "#DIV/0!"}), "#DIV/0!", cell)
            changed_inputs = copy.deepcopy(self.inputs)
            changed_inputs["sharedInputArchitecture"]["inputs"][selector["featureId"]]["rows"][selector["rowIndex"]][selector["column"]] = "No such source selector"
            self.assertEqual(evaluate_sumifs(mapping, sales, changed_inputs), Decimal(0), cell)

    def test_all_214_source_forward_baselines(self):
        sales = self.inputs["sharedInputArchitecture"]["inputs"]["sales-forecast"]["rows"]
        mappings = self.model["sourceMappings"]
        values = {}
        def value(cell):
            if cell in values:
                return values[cell]
            mapping = mappings[cell]
            kind = mapping["kind"]
            if kind == "record-revenue":
                row = sales[mapping["sourceOrdinal"] - 1]
                result = Decimal(str(row["Units"] or 0)) * Decimal(str(row["Avg Price"] or 0))
            elif kind in ("region-quarter", "product-month"):
                result = evaluate_sumifs(mapping, sales, self.inputs)
            elif kind == "lookup":
                result = evaluate_lookup(mapping, self.inputs)
            elif kind == "total":
                result = sum((value(edge) for edge in mapping["summedSourceCells"]), Decimal(0))
            elif kind == "ratio":
                numerator = value(mapping["numeratorSourceCell"])
                denominator = value(mapping["denominatorSourceCell"])
                result = numerator / Decimal(str(denominator)) if denominator else Decimal(0)
            else:
                result = mapping["sourceAst"]["value"]
            values[cell] = result
            return result
        for cell, record in self.native["formulas"].items():
            actual = value(cell)
            expected = record["localExpected"]
            if isinstance(actual, Decimal):
                places = Decimal(".00000001") if mappings[cell]["kind"] == "ratio" else Decimal(".01")
                self.assertEqual(actual.quantize(places), Decimal(str(expected)).quantize(places), cell)
            else:
                self.assertEqual(actual, expected, cell)
        self.assertEqual(len(values), 214)

    def test_criteria_boundary_and_first_included_error(self):
        sales = [dict(row) for row in self.inputs["sharedInputArchitecture"]["inputs"]["sales-forecast"]["rows"]]
        b4 = self.model["sourceMappings"]["Summary!B4"]
        sales[0]["Month"] = "2026-04-01"
        self.assertEqual(evaluate_sumifs(b4, sales, self.inputs),
                         evaluate_sumifs(b4, sales, self.inputs, errors={1: "#DIV/0!"}))
        sales[0]["Month"] = "2026-01-01"
        self.assertEqual(evaluate_sumifs(b4, sales, self.inputs, errors={1: "#DIV/0!", 2: "#VALUE!"}),
                         "#DIV/0!")

    def test_lookup_first_duplicate_and_missing(self):
        lookup = self.model["sourceMappings"]["Summary!G4"]
        regions = [dict(row) for row in self.inputs["sharedInputArchitecture"]["inputs"]["regions"]["rows"]]
        self.assertEqual(evaluate_lookup(lookup, self.inputs), self.native["formulas"]["Summary!G4"]["localExpected"])
        duplicate = {**regions[0], "Region": regions[0]["Region"], "FY26 Target": 999999}
        self.assertEqual(evaluate_lookup(lookup, self.inputs, region_rows=[duplicate, *regions]), 999999)
        self.assertEqual(evaluate_lookup(lookup, self.inputs, selector_value="missing"), "#N/A")
        reordered = json.loads(json.dumps(self.native))
        reordered["formulas"]["Summary!G4"]["ast"]["arguments"][1]["cells"].reverse()
        from lib.native_exact_plan import digest
        reordered["formulas"]["Summary!G4"]["astHash"] = digest(reordered["formulas"]["Summary!G4"]["ast"])
        reordered["nativePlanHash"] = digest({k: v for k, v in reordered.items() if k != "nativePlanHash"})
        with self.assertRaisesRegex(ValueError, "first-duplicate"):
            build_grouped_model(self.inputs, reordered)

    def test_full_output_formulas_and_source_forward_edges(self):
        assigned = synthetic_assigned_rows(self.inputs)
        graph = build_full_output_graph(self.model, json.loads(gzip.decompress(CANARY.read_bytes())), assigned)
        self.assertEqual(graph["sourceFormulaCount"], 214)
        self.assertEqual(graph["constructionBudget"], {"tables": 7, "formulaAndAggregateColumns": 145})
        self.assertEqual({k: len(v) for k, v in graph["formulaColumns"].items()},
                         {"region-summary": 56, "product-summary": 25, "totals": 30})
        self.assertTrue(all(len(m["laneColumnNames"]) == 5 for m in graph["sourceMappings"].values()))
        self.assertEqual(graph["sourceMappings"]["Summary!F4"]["targetRowIndex"], 0)
        self.assertEqual(graph["sourceMappings"]["Summary!B8"]["targetField"], "Quarter Total 1")
        self.assertEqual(graph["sourceMappings"]["Summary!F8"]["targetField"], "Grand Total")
        self.assertEqual(graph["sourceMappings"]["Summary!H4"]["targetRowIndex"], 0)
        self.assertIn("test-only-regions-0", graph["formulaColumns"]["region-summary"]["Target Matched Row ID"])
        self.assertIn("[Summary Months/Month]", graph["formulaColumns"]["product-summary"]["Product 1 Matched Revenue"])
        names = {
            "Summary Regions": ["Region", "Source Ordinal", "Source Key"],
            "Summary Months": ["Month", "Source Key"],
            "Summary Products": ["Product Line", "Row ID"],
            "Regions": ["Region", "FY26 Target", "Row ID"],
            "Grouped": ["Revenue Sum", "First Error Ordinal", "Group Key"],
            "Product Grouped": ["Revenue Sum", "First Error Ordinal", "Product Key", "Month Key"],
            "Rows": ["Revenue Error", "Source Ordinal"],
            "Region Summary": ["Quarter 1 Error", "Quarter 2 Error", "Quarter 3 Error", "Quarter 4 Error", "Annual Error", "Source Ordinal"],
            "Totals": [name for spec in graph["totalGrouping"]["calculations"].values()
                       for name in (spec["sum"], spec["errorOrdinal"])],
        }
        namespaces = {name: {"name": name, "columns": {column: {"name": column, "id": column}
                      for column in columns}} for name, columns in names.items()}
        for element in graph["formulaColumns"]:
            bound = bind_full_output_formulas(graph, namespaces, element=element)
            self.assertEqual(set(bound), set(graph["formulaColumns"][element]))
            self.assertFalse(any("__native" in formula for formula in bound.values()))
        namespaces["Totals"]["columns"]["Quarter Total 1 Sum"]["name"] = "Observed Quarter Sum"
        rebound = bind_full_output_formulas(graph, namespaces, element="totals")
        self.assertEqual(rebound["Quarter Total 1 Number"], "[Observed Quarter Sum]")

    def test_totals_error_lanes_use_grouped_ordinals_without_collapsed_lookup(self):
        assigned = synthetic_assigned_rows(self.inputs)
        graph = build_full_output_graph(self.model, json.loads(gzip.decompress(CANARY.read_bytes())), assigned)
        region = graph["formulaColumns"]["region-summary"]
        totals = graph["formulaColumns"]["totals"]
        for quarter in range(1, 5):
            lane = f"Quarter Total {quarter}"
            self.assertEqual(region[f"Quarter {quarter} Error Ordinal"],
                             f'If([Quarter {quarter} Error] != "", [Source Ordinal], Null)')
            self.assertEqual(graph["totalGrouping"]["calculations"][lane]["errorSource"],
                             f"Quarter {quarter} Error Ordinal")
            self.assertEqual(totals[lane + " Error"],
                             f'If(IsNull([{lane} First Error Ordinal]), "", "#VALUE!")')
        self.assertEqual(region["Annual Error Ordinal"],
                         'If([Annual Error] != "", [Source Ordinal], Null)')
        self.assertEqual(totals["Grand Total Error"],
                         'If(IsNull([Grand Total First Error Ordinal]), "", "#VALUE!")')
        self.assertTrue(all("Lookup([Region Summary/" not in totals[lane + " Error"]
                            for lane in ("Quarter Total 1", "Quarter Total 2",
                                         "Quarter Total 3", "Quarter Total 4", "Grand Total")))

    def test_emitted_grouped_graph_all_70_nonrecord_baselines(self):
        assigned = synthetic_assigned_rows(self.inputs)
        graph = build_full_output_graph(self.model, json.loads(gzip.decompress(CANARY.read_bytes())), assigned)
        inputs = self.inputs["sharedInputArchitecture"]["inputs"]
        sales = inputs["sales-forecast"]["rows"]
        region_rows = []
        product_rows = []
        for ordinal, source in enumerate(sales, 1):
            revenue = float(Decimal(str(source["Units"] or 0)) * Decimal(str(source["Avg Price"] or 0)))
            year = source["Month"][:4]
            quarter = (int(source["Month"][5:7]) - 1) // 3 + 1
            region_rows.append({"Revenue Number": revenue, "Revenue Error": "", "Source Ordinal": ordinal,
                                "Group Key": source["Region"].lower() + "|" + year + "|" + str(quarter)})
            product_rows.append({"Revenue Number": revenue, "Revenue Error": "", "Source Ordinal": ordinal,
                                 "Product Key": source["Product Line"].lower(), "Month Key": source["Month"]})
        def grouped(rows, keys):
            result = {}
            for row in rows:
                key = tuple(row[k] for k in keys)
                group = result.setdefault(key, {k: row[k] for k in keys} | {"Revenue Sum": 0.0,
                                                                           "First Error Ordinal": None})
                group["Revenue Sum"] += row["Revenue Number"]
            return list(result.values())
        region_grouped = grouped(region_rows, ("Group Key",))
        product_grouped = grouped(product_rows, ("Product Key", "Month Key"))
        def table_rows(feature):
            return [{**row, "ID": f"test-only-{feature}-{i}", "Row ID": f"test-only-{feature}-{i}",
                     "Source Ordinal": i + 1,
                     "Source Key": next(iter(inputs[feature]["sourceCells"][str(i)].values()))}
                    for i, row in enumerate(inputs[feature]["rows"])]
        reference = NativeGraphReference({
            "regions": ("Regions", table_rows("regions")),
            "region-input": ("Summary Regions", table_rows("summary-regions")),
            "months": ("Summary Months", table_rows("summary-months")),
            "products": ("Summary Products", table_rows("summary-products")),
            "rows": ("Rows", region_rows),
            "region-grouped": ("Grouped", region_grouped),
            "product-grouped": ("Product Grouped", product_grouped),
        })
        reference.tables["Region Summary"] = copy.deepcopy(reference.tables["Summary Regions"])
        for column, formula in graph["formulaColumns"]["region-summary"].items():
            for index, row in enumerate(reference.tables["Region Summary"]):
                row["values"][column] = reference.evaluate(formula, "Region Summary", index)
        reference.tables["Product Summary"] = copy.deepcopy(reference.tables["Summary Months"])
        for column, formula in graph["formulaColumns"]["product-summary"].items():
            for index, row in enumerate(reference.tables["Product Summary"]):
                row["values"][column] = reference.evaluate(formula, "Product Summary", index)
        total_values = {}
        for spec in graph["totalGrouping"]["calculations"].values():
            total_values[spec["sum"]] = sum(row["values"][spec["source"]]
                                             for row in reference.tables["Region Summary"])
            ordinals = [row["values"][spec["errorSource"]]
                        for row in reference.tables["Region Summary"]
                        if row["values"][spec["errorSource"]] is not None]
            total_values[spec["errorOrdinal"]] = min(ordinals) if ordinals else None
        reference.tables["Totals"] = [{"origin": ("totals", 0), "values": total_values}]
        for column, formula in graph["formulaColumns"]["totals"].items():
            reference.tables["Totals"][0]["values"][column] = reference.evaluate(formula, "Totals", 0)
        tables = {"region-summary": "Region Summary", "product-summary": "Product Summary", "totals": "Totals"}
        for cell, mapping in graph["sourceMappings"].items():
            if mapping["kind"] == "record-revenue":
                continue
            row = reference.tables[tables[mapping["targetElement"]]][mapping["targetRowIndex"]]["values"]
            number = row[mapping["laneColumnNames"]["number"]]
            error = row[mapping["laneColumnNames"]["error"]]
            self.assertEqual(error, "", cell)
            expected = self.native["formulas"][cell]["localExpected"]
            if mapping["kind"] == "text":
                self.assertEqual(row[mapping["laneColumnNames"]["text"]], expected, cell)
            else:
                places = Decimal(".00000001") if mapping["kind"] == "ratio" else Decimal(".01")
                self.assertEqual(Decimal(str(number)).quantize(places),
                                 Decimal(str(expected)).quantize(places), cell)

    def test_full_output_scaffold_patch_and_totals_grouping_contract(self):
        import ast
        from lib.native_exact_plan import digest
        assigned = synthetic_assigned_rows(self.inputs)
        graph = build_full_output_graph(self.model, json.loads(gzip.decompress(CANARY.read_bytes())), assigned)
        namespaces = {}
        elements = []
        required = {
            "Summary Regions": ("input-table", ["Region", "Source Ordinal", "Source Key"]),
            "Regions": ("input-table", ["Region", "FY26 Target", "Row ID"]),
            "Grouped": ("table", ["Revenue Sum", "First Error Ordinal", "Group Key"]),
            "Rows": ("table", ["Revenue Error", "Source Ordinal"]),
        }
        for alias, (kind, names) in required.items():
            element_id = "id-" + alias
            columns = [{"id": alias + "-" + name, "name": name, "formula": "1"} for name in names]
            source = ({"kind": "table", "elementId": "id-Sales"} if alias == "Rows" else
                      {"kind": "table", "elementId": "id-Rows", "groupingId": "test-group"}
                      if alias == "Grouped" else None)
            element = {"id": element_id, "name": alias, "kind": kind, "columns": columns}
            if source:
                element["source"] = source
            elements.append(element)
            namespaces[alias] = {"elementId": element_id, "name": alias,
                                 "columns": {name: {"id": alias + "-" + name, "name": name} for name in names}}
        current = {"workbookId": "test-doc", "documentVersion": 10,
                   "contents": {"pages": [{"id": "test-page"}], "elements": elements,
                       "layout": '<Page type="grid" id="test-page" gridTemplateColumns="repeat(24, 1fr)"/>'}}
        elements.append({"id": "id-Sales", "name": "Sales", "kind": "input-table",
                         "columns": [{"id": "ID"}]})
        assigned["features"]["sales-forecast"].update({"elementId": "id-Sales", "name": "Sales"})
        assigned["features"]["summary-regions"].update({"elementId": "id-Summary Regions", "name": "Summary Regions"})
        assigned["features"]["regions"].update({"elementId": "id-Regions", "name": "Regions"})
        assigned.update({"documentId": current["workbookId"], "documentVersion": current["documentVersion"],
                         "contentsHash": digest(current["contents"])})
        scaffold = prepare_output_scaffold_program(graph, current, assigned, output_element="region-summary",
            source_element_id="id-Summary Regions", page_id="test-page", grid_row=30)
        with self.assertRaisesRegex(ValueError, "exact reconciled editable role"):
            prepare_output_scaffold_program(graph, current, assigned, output_element="region-summary",
                source_element_id="id-Regions", page_id="test-page", grid_row=30)
        ast.parse(scaffold["program"])
        observed = copy.deepcopy(current)
        observed["documentVersion"] += 1
        output_columns = {name: "out-" + name for name in graph["formulaColumns"]["region-summary"]}
        observed["contents"]["elements"].append({"id": "region-output", "name": "Regional Summary",
            "kind": "table", "source": {"kind": "table", "elementId": "id-Summary Regions"},
            "columns": [{"id": col_id, "name": name, "formula": "0"}
                        for name, col_id in output_columns.items()]})
        assigned.update({"documentVersion": observed["documentVersion"],
                         "contentsHash": digest(observed["contents"])})
        patch = prepare_output_contents_patch(graph, observed, scaffold, assigned,
            output_element="region-summary", output_element_id="region-output",
            source_element_id="id-Summary Regions", assigned_columns=output_columns,
            namespaces=namespaces)
        wrong = copy.deepcopy(observed)
        wrong["contents"]["elements"][-1]["source"]["elementId"] = "id-Regions"
        assigned["contentsHash"] = digest(wrong["contents"])
        with self.assertRaisesRegex(ValueError, "exact reconciled editable role"):
            prepare_output_contents_patch(graph, wrong, scaffold, assigned,
                output_element="region-summary", output_element_id="region-output",
                source_element_id="id-Summary Regions", assigned_columns=output_columns,
                namespaces=namespaces)
        assigned["contentsHash"] = digest(observed["contents"])
        self.assertEqual(patch["ifMatchDocumentVersion"], observed["documentVersion"])
        self.assertNotEqual(patch["beforeContentsHash"], patch["expectedContentsHash"])
        saved = copy.deepcopy(observed)
        saved["documentVersion"] += 1
        saved["contents"] = patch["contents"]
        confirmation = confirm_output_contents_patch(patch, saved)
        assigned.update({"documentVersion": saved["documentVersion"],
                         "contentsHash": digest(saved["contents"])})
        totals = prepare_totals_grouping_program(graph, saved, confirmation, assigned,
            region_element_id="region-output", page_id="test-page", grid_row=60)
        ast.parse(totals["program"])
        self.assertEqual(totals["program"].count("grouping.set_calculation("), 10)
        self.assertIn("grouping.to_source()", totals["program"])
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            confirm_output_contents_patch(patch, observed)

    def test_final_214_manifest_uses_installed_recipes_and_confirmed_ids(self):
        from lib.native_exact_plan import digest
        from lib.regional_grouped_canary import bind_assigned_formula
        canary = json.loads(gzip.decompress(CANARY.read_bytes()))
        assigned = synthetic_assigned_rows(self.inputs)
        graph = build_full_output_graph(self.model, canary, assigned)
        elements = []
        namespaces = {}
        def add(alias, element_id, element_name, kind, source, names, grouping=None):
            columns = [{"id": alias + ":" + name, "name": name, "formula": "0"} for name in names]
            element = {"id": element_id, "name": element_name, "kind": kind, "columns": columns}
            if source is not None:
                element["source"] = {"kind": "table", "elementId": source}
                if grouping:
                    element["source"]["groupingId"] = grouping
            elements.append(element)
            namespaces[alias] = {"elementId": element_id, "name": element_name,
                "columns": {name: {"id": alias + ":" + name, "name": name} for name in names}}
            return element
        inputs = (("Sales", "sales-forecast"), ("Regions", "regions"),
                  ("Summary Regions", "summary-regions"), ("Summary Months", "summary-months"),
                  ("Summary Products", "summary-products"))
        for alias, feature in inputs:
            receipt = assigned["features"][feature]
            names = [c["name"] for c in self.inputs["sharedInputArchitecture"]["inputs"][feature]["columns"]]
            names += ["Source Key", "Source Ordinal", "Row ID"]
            element = add(alias, receipt["elementId"], receipt["name"], "input-table", None, names)
            element["columns"].append({"id": "ID"})
        record = canary["elements"][0]["columns"]
        row_names = [c["name"] for c in record] + ["Product Key", "Month Key", "Revenue Sum"]
        rows = add("Rows", "Rows", "Rows", "table", assigned["features"]["sales-forecast"]["elementId"], row_names)
        rows["columns"].append({"id": "rows-aggregate-error", "name": "First Error Ordinal",
                                "formula": "Min([First Error Ordinal (1)])"})
        grouped = add("Grouped", "Grouped", "Grouped", "table", "Rows",
                      ["Group Key", "Revenue Sum", "First Error Ordinal"], "g1")
        product_rows = add("Product Rows", "Product Rows", "Product Rows", "table", "Rows",
                           ["Product Key", "Month Key", "Revenue Number", "Revenue Error",
                            "Source Ordinal", "Product Error Ordinal", "Revenue Sum", "First Error Ordinal"])
        product_grouped = add("Product Grouped", "Product Grouped", "Product Grouped", "table",
                              "Product Rows", ["Product Key", "Month Key", "Revenue Sum",
                                               "First Error Ordinal"], "g2")
        region_names = list(graph["formulaColumns"]["region-summary"])
        for spec in graph["totalGrouping"]["calculations"].values():
            region_names += [spec["sum"], spec["errorOrdinal"]]
        region = add("Region Summary", "Region Summary", "Region Summary", "table",
                     assigned["features"]["summary-regions"]["elementId"], region_names)
        product = add("Product Summary", "Product Summary", "Product Summary", "table",
                      assigned["features"]["summary-months"]["elementId"],
                      list(graph["formulaColumns"]["product-summary"]))
        totals_names = list(graph["formulaColumns"]["totals"])
        for spec in graph["totalGrouping"]["calculations"].values():
            totals_names += [spec["sum"], spec["errorOrdinal"]]
        totals = add("Totals", "Totals", "Totals", "table", "Region Summary", totals_names, "g3")
        def column(element, name):
            return next(c for c in element["columns"] if c.get("name") == name)
        local = {c["name"]: c["name"] for c in record}
        for recipe in record:
            column(rows, recipe["name"])["formula"] = bind_assigned_formula(
                recipe["formula"], namespaces=namespaces, local_columns=local)
            if recipe["name"] in ("h0", "h1", "h2"):
                column(rows, recipe["name"])["formula"] = column(rows, recipe["name"])["formula"].replace(
                    "Not(", "Not (")
        column(rows, "Product Key")["formula"] = f'Lower([{namespaces["Sales"]["name"]}/Product Line])'
        column(rows, "Month Key")["formula"] = f'[{namespaces["Sales"]["name"]}/Month]'
        column(rows, "Revenue Sum")["formula"] = "Sum([Revenue Number])"
        for name in ("Group Key", "Revenue Sum", "First Error Ordinal"):
            column(grouped, name)["formula"] = f'[Rows/{name}]'
        for name in ("Product Key", "Month Key", "Revenue Number", "Revenue Error", "Source Ordinal"):
            column(product_rows, name)["formula"] = f'[Rows/{name}]'
        column(product_rows, "Product Error Ordinal")["formula"] = 'If([Revenue Error] != "", [Source Ordinal], Null)'
        column(product_rows, "Revenue Sum")["formula"] = "Sum([Revenue Number])"
        column(product_rows, "First Error Ordinal")["formula"] = "Min([Product Error Ordinal])"
        for name in ("Product Key", "Month Key", "Revenue Sum", "First Error Ordinal"):
            column(product_grouped, name)["formula"] = f'[Product Rows/{name}]'
        for role, element in (("region-summary", region), ("product-summary", product), ("totals", totals)):
            for name, formula in bind_full_output_formulas(graph, namespaces, element=role).items():
                column(element, name)["formula"] = formula
        for spec in graph["totalGrouping"]["calculations"].values():
            column(region, spec["sum"])["formula"] = f'Sum([{spec["source"]}])'
            column(region, spec["errorOrdinal"])["formula"] = f'Min([{spec["errorSource"]}])'
            for name in (spec["sum"], spec["errorOrdinal"]):
                column(totals, name)["formula"] = f'[Region Summary/{name}]'
        rows["groupings"] = [{"id": "g1", "groupBy": [column(rows, "Group Key")["id"]],
                              "calculations": [column(rows, "Revenue Sum")["id"],
                                               "rows-aggregate-error"]}]
        product_rows["groupings"] = [{"id": "g2", "groupBy": [column(product_rows, "Product Key")["id"],
                                                              column(product_rows, "Month Key")["id"]],
                                      "calculations": [column(product_rows, "Revenue Sum")["id"],
                                                       column(product_rows, "First Error Ordinal")["id"]]}]
        total_calc_ids = []
        for spec in graph["totalGrouping"]["calculations"].values():
            total_calc_ids += [column(region, spec["sum"])["id"],
                               column(region, spec["errorOrdinal"])["id"]]
        region["groupings"] = [{"id": "g3", "groupBy": [column(region, "Total Key")["id"]],
                                "calculations": total_calc_ids}]
        current = {"workbookId": "test-full-assigned", "documentVersion": 8,
                   "contents": {"elements": elements}}
        def fresh(doc):
            receipt = copy.deepcopy(assigned)
            receipt.update({"documentId": doc["workbookId"], "documentVersion": doc["documentVersion"],
                            "contentsHash": digest(doc["contents"])})
            return receipt
        phases = {}
        for key, status in (("product-key", "PRODUCT_KEY_FORMULAS_CONFIRMED"),
                            ("product-error", "PRODUCT_ERROR_FORMULA_CONFIRMED"),
                            ("region-summary", "REGIONAL_OUTPUT_FORMULAS_CONFIRMED"),
                            ("product-summary", "REGIONAL_OUTPUT_FORMULAS_CONFIRMED"),
                            ("totals", "REGIONAL_OUTPUT_FORMULAS_CONFIRMED")):
            phases[key] = {"status": status, "documentId": current["workbookId"],
                           "documentVersion": 7, "contentsHash": "a" * 64,
                           "sourceSha256": graph["sourceSha256"]}
        phases["product-key"]["rowElementId"] = "Rows"
        phases["product-error"].update({"rowElementId": "Rows", "productElementId": "Product Rows",
                                        "errorColumnId": namespaces["Product Rows"]["columns"]["Product Error Ordinal"]["id"]})
        for key, alias in (("region-summary", "Region Summary"),
                           ("product-summary", "Product Summary"), ("totals", "Totals")):
            phases[key].update({"logicalElement": key, "outputElementId": alias})
        def bind(doc=current, graph_value=graph, phase_value=phases, aliases=namespaces):
            return bind_full_214_source_manifest(graph_value, doc, fresh(doc), aliases,
                source_input_plan=self.inputs, source_native_plan=self.native,
                reviewed_canary=canary, phase_confirmations=phase_value)
        manifest = bind()
        self.assertEqual(manifest["sourceFormulaCount"], 214)
        self.assertEqual(manifest["roleCounts"], {"sales-rows": 144, "region-summary": 28,
                                                  "product-summary": 36, "totals": 6})
        self.assertEqual(manifest["sourceMappings"]["Summary!G4"]["sourceInputPhysicalRowId"],
                         assigned["features"]["summary-regions"]["rows"]["0"])
        self.assertIsNone(manifest["sourceMappings"]["Summary!B8"]["sourceInputPhysicalRowId"])
        for alias, name in (("Region Summary", "Quarter 1 Number"),
                            ("Rows", "h0"), ("Product Rows", "Product Error Ordinal"),
                            ("Product Grouped", "Revenue Sum"), ("Totals", "Quarter Total 1 Sum")):
            changed = copy.deepcopy(current)
            target = next(e for e in changed["contents"]["elements"] if e["id"] == alias)
            column(target, name)["formula"] = "0"
            with self.assertRaisesRegex(ValueError, "installed formula|grouping recipe"):
                bind(changed)
        changed = copy.deepcopy(current)
        target = next(e for e in changed["contents"]["elements"] if e["id"] == "Rows")
        column(target, "h0")["formula"] = column(target, "h0")["formula"].replace("Not (", "Not  (", 1)
        with self.assertRaisesRegex(ValueError, "installed formula"):
            bind(changed)
        changed = copy.deepcopy(current)
        target = next(e for e in changed["contents"]["elements"] if e["id"] == "Product Rows")
        target["groupings"][0]["groupBy"].reverse()
        with self.assertRaisesRegex(ValueError, "installed grouping keys"):
            bind(changed)
        changed_graph = copy.deepcopy(graph)
        changed_graph["formulaColumns"]["region-summary"]["Quarter 1 Number"] = "0"
        with self.assertRaisesRegex(ValueError, "source-bound full graph differs"):
            bind(graph_value=changed_graph)
        changed_graph = copy.deepcopy(graph)
        changed_graph["sourceMappings"]["Summary!B4"]["sourceDependencies"] = []
        with self.assertRaisesRegex(ValueError, "source-bound full graph differs"):
            bind(graph_value=changed_graph)
        incomplete = copy.deepcopy(phases)
        del incomplete["product-summary"]
        with self.assertRaisesRegex(ValueError, "all completed construction phase"):
            bind(phase_value=incomplete)
        missing = copy.deepcopy(namespaces)
        del missing["Region Summary"]["columns"]["Ratio Number"]
        with self.assertRaisesRegex(ValueError, "assigned lane column missing|installed recipe column"):
            bind(aliases=missing)

    def test_renamed_source_file_and_unsupported_ast_fail_closed(self):
        changed = json.loads(json.dumps(self.native))
        changed["sourcePath"] = "/renamed/forecast.xlsx"
        # The source filename is diagnostic; the pinned AST and source hash are authoritative.
        changed.pop("nativePlanHash")
        from lib.native_exact_plan import digest
        changed["nativePlanHash"] = digest(changed)
        self.assertEqual(len(build_grouped_model(self.inputs, changed)["sourceMappings"]), 214)
        changed["formulas"]["Summary!B4"]["ast"]["arguments"][3]["sourceKind"] = "cell_range"
        changed["nativePlanHash"] = digest({k: v for k, v in changed.items() if k != "nativePlanHash"})
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            build_grouped_model(self.inputs, changed)

    def test_generated_renamed_sheet_profile_is_not_address_hardcoded(self):
        from lib.native_exact_plan import digest
        generated = json.loads(json.dumps(self.native).replace("Forecast Input!", "Demand Records!")
                               .replace("Summary!", "Overview!").replace("Regions!", "Territories!"))
        generated["sourcePath"] = "/generated/renamed-workbook.xlsx"
        for record in generated["formulas"].values():
            record["astHash"] = digest(record["ast"])
        generated["nativePlanHash"] = digest({k: v for k, v in generated.items() if k != "nativePlanHash"})
        rebuilt = build_grouped_model(self.inputs, generated)
        self.assertEqual(rebuilt["recordProfile"]["sheet"], "Demand Records")
        self.assertIn("Overview!B4", rebuilt["sourceMappings"])
        self.assertEqual(len(rebuilt["sourceMappings"]), 214)

    def test_five_input_physical_row_reconciliation_rejects_drift(self):
        inputs = self.inputs["sharedInputArchitecture"]["inputs"]
        elements = []
        observed = {}
        for feature, spec in inputs.items():
            columns = list(spec["rows"][0]) + ["Source Key", "Source Ordinal"]
            elements.append({"id": feature, "name": feature, "kind": "input-table",
                             "columns": [{"id": name, "name": name} for name in columns] + [{"id": "ID"}]})
            rows = []
            for index, values in enumerate(spec["rows"]):
                cells = list(spec["sourceCells"][str(index)].values())
                key = (cells[0] if len(cells) == 1 and feature not in ("sales-forecast", "summary-regions") else
                       cells[0].rsplit("!", 1)[0] + "!row:" + str(spec["sourceRowNumbers"][index]))
                rows.append({**values, "Source Key": key, "Source Ordinal": index + 1,
                             "ID": f"physical-{feature}-{index}"})
            observed[feature] = {"elementId": feature, "name": feature,
                                 "rowIdColumn": "ID", "queryRows": rows}
        current = {"workbookId": "test-only", "documentVersion": 1,
                   "contents": {"elements": elements}}
        self.assertEqual(set(validate_full_assigned_rows(self.inputs, current, observed)["features"]), set(inputs))
        self.assertEqual(observed["summary-regions"]["queryRows"][0]["Source Key"], "Summary!row:4")
        self.assertEqual(observed["summary-months"]["queryRows"][0]["Source Key"], "Summary!A12")
        self.assertEqual(observed["summary-products"]["queryRows"][0]["Source Key"], "Summary!B11")
        observed["sales-forecast"]["queryRows"][0]["Month"] = "2026-01-01T12:00:00.000Z"
        self.assertEqual(set(validate_full_assigned_rows(self.inputs, current, observed)["features"]), set(inputs))
        for drift in ("2026-01-01T12:34:56.000Z", "2026-01-01T12:00:00.000+01:00"):
            observed["sales-forecast"]["queryRows"][0]["Month"] = drift
            with self.assertRaisesRegex(ValueError, "queried business input differs"):
                validate_full_assigned_rows(self.inputs, current, observed)
        observed["sales-forecast"]["queryRows"][0]["Month"] = "2026-01-01"
        observed["regions"]["queryRows"][0]["FY26 Target"] += 1
        with self.assertRaisesRegex(ValueError, "queried business input differs"):
            validate_full_assigned_rows(self.inputs, current, observed)

    def test_guarded_product_grouping_phases(self):
        current = json.loads(SAVED.read_text())
        rows = next(e for e in current["contents"]["elements"] if e.get("name") == "Regional Sales Rows")
        sales_id = rows["source"]["elementId"]
        page_id = current["contents"]["pages"][0]["id"]
        scaffold = prepare_product_key_scaffold(self.model, current, row_element_id=rows["id"],
                                                 sales_element_id=sales_id, page_id=page_id)
        self.assertIn("const_num(0)", scaffold["program"])
        self.assertIn("isinstance(rows, LevelTable)", scaffold["program"])
        observed = copy.deepcopy(current)
        observed["documentVersion"] += 1
        target = next(e for e in observed["contents"]["elements"] if e.get("id") == rows["id"])
        target["columns"].extend([{"id": "test-product-id", "name": "Product Key", "formula": "0"},
                                   {"id": "test-month-id", "name": "Month Key", "formula": "0"}])
        patch = prepare_product_key_patch(self.model, observed, scaffold, row_element_id=rows["id"],
                                          sales_element_id=sales_id, product_column_id="test-product-id",
                                          month_column_id="test-month-id")
        self.assertEqual(patch["ifMatchDocumentVersion"], observed["documentVersion"])
        self.assertEqual(next(c for c in patch["contents"]["elements"] if c.get("id") == rows["id"])
                         ["columns"][-1]["formula"], "[Regional B4 Sales Inputs/Month]")
        saved = copy.deepcopy(observed)
        saved["contents"] = patch["contents"]
        saved["documentVersion"] += 1
        confirmed = confirm_product_key_patch(patch, saved)
        projection = prepare_product_rows_scaffold_program(self.model, saved, confirmed,
            row_element_id=rows["id"], sales_element_id=sales_id, page_id=page_id, grid_row=40)
        self.assertIn("Product Error Ordinal", projection["program"])
        projected = copy.deepcopy(saved)
        projected["documentVersion"] += 1
        inherited = []
        for name in ("Revenue Error", "Source Ordinal", "Product Key", "Month Key", "Revenue Number"):
            inherited.append({"id": "projected-" + name, "name": name,
                              "formula": f'[{rows["name"]}/{name}]'})
        projected["contents"]["elements"].append({
            "id": "test-product-rows", "name": "Regional Product Rows", "kind": "table",
            "source": {"kind": "table", "elementId": rows["id"]},
            "columns": inherited + [{"id": "test-product-error", "name": "Product Error Ordinal",
                                     "formula": "0"}]})
        error_patch = prepare_product_error_patch(self.model, projected, projection,
            row_element_id=rows["id"], product_element_id="test-product-rows",
            error_column_id="test-product-error")
        self.assertEqual(error_patch["ifMatchDocumentVersion"], projected["documentVersion"])
        error_saved = copy.deepcopy(projected)
        error_saved["contents"] = error_patch["contents"]
        error_saved["documentVersion"] += 1
        error_confirmed = confirm_product_error_patch(error_patch, error_saved)
        grouping = prepare_product_grouping_program(self.model, error_saved, error_confirmed,
            product_element_id="test-product-rows", page_id=page_id, grid_row=40)
        self.assertIn("grouping.to_source()", grouping["program"])
        self.assertIn("sigma.formula.min(error_order)", grouping["program"])
        stale = copy.deepcopy(error_saved)
        stale["documentVersion"] += 1
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            prepare_product_grouping_program(self.model, stale, error_confirmed,
                product_element_id="test-product-rows", page_id=page_id, grid_row=40)


if __name__ == "__main__":
    unittest.main()
