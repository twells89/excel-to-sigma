"""Focused provenance and hostile binding checks for the isolated B4 fixture."""
import copy
import json
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.regional_canary_inputs import (build_canary_input_plan, prepare_input_create_program,
    prepare_input_seed_program, reconcile_input_rows)
from lib.regional_grouped_canary import (prepare_contents_patch, confirm_contents_patch,
    prepare_grouping_program, observe_b4_query)

ROOT = Path(__file__).resolve().parents[5]
WORK = ROOT / "work"
SOURCE = WORK / "broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx"
BASE = WORK / "complete-source-record-bindings-plans-20261020/01-regional-sales-forecast-xlsx.json"
CANDIDATE = WORK / "regional-grouped-canary-20261020/candidate.json"


@unittest.skipUnless(SOURCE.exists() and BASE.exists() and CANDIDATE.exists(),
                     "source-bound Regional canary packet is not installed")
class RegionalCanaryInputsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = json.loads(BASE.read_text())
        cls.candidate = json.loads(CANDIDATE.read_text())
        cls.plan = build_canary_input_plan(SOURCE, cls.base, cls.candidate)

    def setup_document(self):
        elements = []
        bindings = {}
        queries = {}
        for feature, spec in self.plan["features"].items():
            prefix = "test-only-sales" if feature == "sales-forecast" else "test-only-summary"
            columns = [{"id": f"{prefix}-c{i}", "name": c["name"], "type": c["type"]}
                       for i, c in enumerate(spec["columns"])]
            columns.append({"id": f"{prefix}-id", "name": "ID", "type": "text"})
            element = {"id": prefix, "kind": "input-table", "name": spec["name"],
                       "source": {"kind": "empty", "connectionId": "test-only-connection"},
                       "columns": columns}
            elements.append(element)
            bindings[feature] = {"documentId": "test-only-document",
                "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
                "elementId": prefix, "connectionId": "test-only-connection",
                "pageId": "test-only-page", "columnIds": {c["name"]: c["id"] for c in columns[:-1]},
                "rowIdColumn": "ID", "rowIdColumnId": columns[-1]["id"]}
            queries[feature] = {"elementId": prefix, "documentVersion": 7,
                "snapshotId": f"test-only-query-{feature}",
                "rows": [{**row, "ID": f"test-only-physical-{feature}-{i}"}
                         for i, row in enumerate(spec["rows"])]}
        current = {"workbookId": "test-only-document", "documentVersion": 7,
                   "contents": {"kind": "workbook", "pages": [{"id": "test-only-page"}],
                                "elements": elements, "layout": ""}}
        return current, bindings, queries

    def observed(self, current, bindings, queries):
        return {feature: {"binding": bindings[feature], "query": queries[feature],
                          "receipt": reconcile_input_rows(self.plan, current, feature=feature,
                              binding=bindings[feature], query=queries[feature])}
                for feature in bindings}

    def namespace(self, element, *, system_alias=None):
        columns = {c["name"]: {"id": c["id"], "name": c["name"]}
                   for c in element["columns"]}
        if system_alias:
            columns["Row ID"] = columns["ID"]
        return {"elementId": element["id"], "name": element["name"], "columns": columns}

    def test_projection_exact_two_inputs_and_pending_denominator(self):
        self.assertEqual(self.plan["status"], "CANARY_QUALIFICATION_ONLY")
        self.assertEqual({k: len(v["rows"]) for k, v in self.plan["features"].items()},
                         {"sales-forecast": 144, "summary-regions": 4})
        self.assertEqual(len(self.plan["pendingFormulaCellIds"]), 214)
        self.assertEqual(self.plan["omittedSourceInputFeatures"],
                         ["regions", "summary-months", "summary-products"])
        sales = self.plan["features"]["sales-forecast"]
        self.assertEqual(sales["sourceCells"]["0"]["Units"], "Forecast Input!D2")
        self.assertEqual(sales["rows"][0]["Source Key"], "Forecast Input!row:2")
        self.assertEqual(sales["rows"][0]["Source Ordinal"], 1)
        self.assertNotIn("Source Key", sales["sourceCells"]["0"])
        self.assertEqual(self.plan["features"]["summary-regions"]["sourceCells"]["0"],
                         {"Region": "Summary!A4"})

    def test_create_seed_and_reconcile_are_bounded_and_observed(self):
        current, bindings, queries = self.setup_document()
        empty = copy.deepcopy(current)
        empty["contents"]["elements"] = []
        for feature in bindings:
            create = prepare_input_create_program(self.plan, empty, feature=feature,
                page_id="test-only-page", connection_id="test-only-connection", grid_row=1)
            compile(create["program"], "<input-create>", "exec")
            self.assertNotIn("upsert_rows", create["program"])
            empty_query = {"elementId": bindings[feature]["elementId"],
                           "documentVersion": 7, "snapshotId": "test-only-empty", "rows": []}
            seed = prepare_input_seed_program(self.plan, current, feature=feature,
                binding=bindings[feature], empty_query=empty_query)
            compile(seed["program"], "<input-seed>", "exec")
            self.assertIn("row_ids = table.upsert_rows([", seed["program"])
            self.assertEqual(seed["expectedRows"], len(self.plan["features"][feature]["rows"]))
            result = reconcile_input_rows(self.plan, current, feature=feature,
                                          binding=bindings[feature], query=queries[feature])
            self.assertEqual(len(result["rows"]), seed["expectedRows"])
        wrong = copy.deepcopy(queries["summary-regions"])
        wrong["rows"][0]["Region"] = "Wrong"
        with self.assertRaisesRegex(ValueError, "value differs"):
            reconcile_input_rows(self.plan, current, feature="summary-regions",
                                 binding=bindings["summary-regions"], query=wrong)

    def test_patch_checks_input_source_namespaces_and_b4_row_id(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        row_logical = self.candidate["elements"][0]
        row_element = {"id": "test-only-row-table", "kind": "table", "name": "Rows",
                       "source": {"kind": "table", "elementId": bindings["sales-forecast"]["elementId"]},
                       "columns": [{"id": f"test-only-row-c{i}", "name": c["name"], "formula": "0"}
                                   for i, c in enumerate(row_logical["columns"])]}
        current["contents"]["elements"].append(row_element)
        namespaces = {"Sales": self.namespace(current["contents"]["elements"][0], system_alias=True),
                      "Summary Regions": self.namespace(current["contents"]["elements"][1], system_alias=True),
                      "Rows": self.namespace(row_element)}
        receipt = {"sourceSha256": self.plan["sourceSha256"],
                   "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
                   "documentId": current["workbookId"], "documentVersion": 7,
                   "elements": {row_logical["id"]: {"elementId": row_element["id"],
                       "name": row_element["name"],
                       "columns": {logical["id"]: actual["id"] for logical, actual in
                                   zip(row_logical["columns"], row_element["columns"])}}},
                   "namespaces": namespaces}
        kwargs = dict(element_kind="calculated-table", source_path=SOURCE,
                      base_input_plan=self.base, canary_input_plan=self.plan,
                      observed_inputs=observed)
        packet = prepare_contents_patch(self.candidate, current, receipt, **kwargs)
        self.assertEqual(len(packet["changes"]), len(row_logical["columns"]))
        self.assertEqual(confirm_contents_patch(packet, {"workbookId": current["workbookId"],
            "documentVersion": 8, "contents": packet["contents"]})["runtimeStatus"],
                         "UNVERIFIED_CANDIDATE")
        wrong_source = copy.deepcopy(current)
        wrong_source["contents"]["elements"][-1]["source"]["elementId"] = "wrong-sales"
        with self.assertRaisesRegex(ValueError, "sourced|source"):
            prepare_contents_patch(self.candidate, wrong_source, receipt, **kwargs)
        wrong_namespace = copy.deepcopy(receipt)
        wrong_namespace["namespaces"]["Sales"]["name"] = "Wrong Sales"
        with self.assertRaisesRegex(ValueError, "namespace"):
            prepare_contents_patch(self.candidate, current, wrong_namespace, **kwargs)
        wrong_column = copy.deepcopy(receipt)
        wrong_column["namespaces"]["Sales"]["columns"]["Units"]["id"] = "foreign"
        with self.assertRaisesRegex(ValueError, "lineage"):
            prepare_contents_patch(self.candidate, current, wrong_column, **kwargs)

    def test_grouping_and_summary_b4_physical_row_readback(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        row_logical, _, summary_logical = self.candidate["elements"]
        row_element = {"id": "test-only-row-table", "kind": "table", "name": "Regional Sales Rows",
            "source": {"kind": "table", "elementId": bindings["sales-forecast"]["elementId"]},
            "columns": [{"id": f"test-only-row-c{i}", "name": c["name"], "formula": c["formula"]}
                        for i, c in enumerate(row_logical["columns"])]}
        current["contents"]["elements"].append(row_element)
        from lib.native_exact_plan import digest
        confirmed = {"status": "CONFIRMED_FORMULA_PATCH_STRUCTURE_ONLY",
            "sourceSha256": self.candidate["sourceSha256"], "documentId": current["workbookId"],
            "documentVersion": 7, "contentsHash": digest(current["contents"])}
        grouped = prepare_grouping_program(self.candidate, current, confirmed,
            row_element_id=row_element["id"], page_id="test-only-page", grid_row=15)
        compile(grouped["program"], "<grouping>", "exec")
        self.assertIn("sigma.formula.min(error_ordinal)", grouped["program"])
        self.assertIn("grouping.to_source()", grouped["program"])
        grouped_element = {"id": "test-only-grouped", "kind": "table",
            "name": "Regional Sales Region Quarter",
            "source": {"kind": "grouping", "elementId": row_element["id"]},
            "columns": [{"id": f"test-only-group-c{i}", "name": name}
                        for i, name in enumerate(("Group Key", "Revenue Sum", "First Error Ordinal"))]}
        summary_element = {"id": "test-only-summary-table", "kind": "table", "name": "Regional Summary B4",
            "source": {"kind": "table", "elementId": bindings["summary-regions"]["elementId"]},
            "columns": [{"id": f"test-only-summary-c{i}", "name": c["name"], "formula": "0"}
                        for i, c in enumerate(summary_logical["columns"])]}
        current["contents"]["elements"] += [grouped_element, summary_element]
        namespaces = {"Sales": self.namespace(current["contents"]["elements"][0], system_alias=True),
            "Summary Regions": self.namespace(current["contents"]["elements"][1], system_alias=True),
            "Rows": self.namespace(row_element), "Grouped": self.namespace(grouped_element)}
        receipt = {"sourceSha256": self.plan["sourceSha256"],
            "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
            "documentId": current["workbookId"], "documentVersion": 7,
            "elements": {summary_logical["id"]: {"elementId": summary_element["id"],
                "name": summary_element["name"],
                "columns": {logical["id"]: actual["id"] for logical, actual in
                            zip(summary_logical["columns"], summary_element["columns"])}}},
            "namespaces": namespaces}
        kwargs = dict(element_kind="summary-table", source_path=SOURCE,
            base_input_plan=self.base, canary_input_plan=self.plan, observed_inputs=observed)
        patch = prepare_contents_patch(self.candidate, current, receipt, **kwargs)
        self.assertEqual(patch["b4Source"]["sourceRowId"],
                         observed["summary-regions"]["receipt"]["rows"]["0"]["rowId"])
        readback = {"workbookId": current["workbookId"], "documentVersion": 8,
                    "contents": patch["contents"]}
        query = {"elementId": summary_element["id"], "documentVersion": 8,
                 "snapshotId": "test-only-summary-query",
                 "rows": [{"Source Key": "Summary!row:4",
                           "Source Row ID": patch["b4Source"]["sourceRowId"],
                           "B4 Number": 1544984.02, "B4 Text": "", "B4 Logical": False,
                           "B4 Kind": "n", "B4 Error": ""}]}
        result = observe_b4_query(patch, readback, query)
        self.assertEqual(result["status"], "B4_QUERY_OBSERVED_UNVERIFIED")
        wrong = copy.deepcopy(query)
        wrong["rows"][0]["Source Row ID"] = "test-only-other-row"
        with self.assertRaisesRegex(ValueError, "physical row"):
            observe_b4_query(patch, readback, wrong)


if __name__ == "__main__":
    unittest.main()
