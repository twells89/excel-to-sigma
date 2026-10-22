"""Full row verification must not promote foreign or stale query evidence."""
import copy
import gzip
import json
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.regional_grouped_output import validate_full_assigned_rows
from lib.regional_grouped_phases import prepare_output_contents_patch
from lib.native_exact_plan import digest


def capture_query_evidence(receipt):
    """Test-only capture of a selected projection and its returned value matrix."""
    columns = list(receipt["columnIds"].items()) + [(receipt["rowIdColumn"], receipt["rowIdQueryColumnId"])]
    select = ", ".join(f'"{cid}" AS "{name}"' for name, cid in columns)
    receipt["queryRequest"] = {"query": {"type": "workbook", "workbookId": receipt["documentId"],
        "sql": f'SELECT {select} FROM "workbook"."{receipt["elementId"]}"'}}
    receipt["resultHash"] = digest({"columns": [name for name, _ in columns],
        "rows": [[row[name] for name, _ in columns] for row in receipt["queryRows"]]})


class RegionalFullReconciliationTest(unittest.TestCase):
    def setUp(self):
        fixture = Path(__file__).parent / "fixtures/regional-grouped/input-plan.json.gz"
        self.plan = json.loads(gzip.decompress(fixture.read_bytes()))
        self.current = {"workbookId": "regional-current", "documentVersion": 123,
                        "contents": {"elements": []}}
        self.observed = {}
        for feature, spec in self.plan["sharedInputArchitecture"]["inputs"].items():
            names = list(spec["rows"][0]) + ["Source Key", "Source Ordinal"]
            column_ids = {name: f"{feature}-column-{i}" for i, name in enumerate(names)}
            self.current["contents"]["elements"].append({
                "id": feature, "name": feature, "kind": "input-table",
                "columns": [{"id": cid, "name": name} for name, cid in column_ids.items()]
                           + [{"id": "ID"}]})
            rows = []
            for index, values in enumerate(spec["rows"]):
                cells = list(spec["sourceCells"][str(index)].values())
                key = (cells[0] if len(cells) == 1 and feature not in ("sales-forecast", "summary-regions")
                       else cells[0].rsplit("!", 1)[0] + "!row:" + str(spec["sourceRowNumbers"][index]))
                rows.append({**values, "Source Key": key, "Source Ordinal": index + 1,
                             "ID": f"physical-{feature}-{index}"})
            self.observed[feature] = {
                "documentId": "regional-current", "documentVersion": 123,
                "snapshotId": "query-123-" + feature,
                "elementId": feature, "name": feature, "columnIds": column_ids,
                "rowIdQueryColumnId": "ID", "rowIdColumn": "ID", "queryRows": rows[::-1]}
            capture_query_evidence(self.observed[feature])

    def test_current_bound_queries_reconcile_all_167_physical_rows(self):
        # Query order and the observed civil-noon datetime serialization are supported.
        for row in self.observed["sales-forecast"]["queryRows"]:
            row["Month"] += "T12:00:00.000Z"
        capture_query_evidence(self.observed["sales-forecast"])
        result = validate_full_assigned_rows(self.plan, self.current, self.observed)
        self.assertEqual(result["status"], "FULL_ASSIGNED_ROWS_VERIFIED")
        self.assertEqual(result["documentVersion"], 123)
        self.assertEqual(sum(len(v["rows"]) for v in result["features"].values()), 167)
        self.assertEqual(result["features"]["summary-products"]["rows"], {
            "0": "physical-summary-products-0", "1": "physical-summary-products-1",
            "2": "physical-summary-products-2"})
        self.assertEqual(result["features"]["summary-products"]["rowIdQueryColumnId"], "ID")
        self.assertEqual(result["features"]["summary-products"]["rowIdColumnName"], "ID")

    def test_each_input_requires_current_document_and_snapshot(self):
        for feature in self.observed:
            for field, value in (("documentId", "foreign-workbook"), ("documentId", None),
                                 ("documentVersion", -999), ("documentVersion", 124),
                                 ("documentVersion", "123"), ("documentVersion", None),
                                 ("snapshotId", None), ("snapshotId", ""),
                                 ("snapshotId", "   ")):
                with self.subTest(feature=feature, field=field, value=value):
                    changed = copy.deepcopy(self.observed)
                    changed[feature][field] = value
                    with self.assertRaises(ValueError):
                        validate_full_assigned_rows(self.plan, self.current, changed)

    def test_bare_row_projection_explains_missing_query_provenance(self):
        bare = {feature: {key: receipt[key] for key in ("elementId", "name", "rowIdColumn", "queryRows")}
                for feature, receipt in self.observed.items()}
        with self.assertRaisesRegex(ValueError, "fresh full input query required.*documentId.*documentVersion.*snapshotId"):
            validate_full_assigned_rows(self.plan, self.current, bare)

    def test_observed_named_row_id_column_remains_supported(self):
        self.current["contents"]["elements"][0]["columns"].append({
            "id": "physical-row-column", "name": "Row ID"})
        feature = self.current["contents"]["elements"][0]["id"]
        receipt = self.observed[feature]
        receipt["rowIdColumn"] = receipt["rowIdQueryColumnId"] = "physical-row-column"
        for row in receipt["queryRows"]:
            row["physical-row-column"] = row.pop("ID")
        capture_query_evidence(receipt)
        result = validate_full_assigned_rows(self.plan, self.current, self.observed)
        self.assertEqual(result["features"][feature]["rows"]["0"], f"physical-{feature}-0")
        self.assertEqual(result["features"][feature]["rowIdQueryColumnId"], "physical-row-column")
        self.assertEqual(result["features"][feature]["rowIdColumnName"], "Row ID")

    def test_system_id_with_unreviewed_semantic_name_fails_closed(self):
        self.current["contents"]["elements"][0]["columns"][-1]["name"] = "Unreviewed Identifier"
        with self.assertRaisesRegex(ValueError, "row-ID column identity"):
            validate_full_assigned_rows(self.plan, self.current, self.observed)

    def _product_row_id_patch(self, row_id_binding, *, selected_id="ID", selected_name=None,
                              omit_verified_field=None):
        """Exercise the public patch gate after real full-row reconciliation."""
        products = next(e for e in self.current["contents"]["elements"] if e["id"] == "summary-products")
        if selected_id != "ID":
            products["columns"].append({"id": selected_id, "name": selected_name})
        receipt = self.observed["summary-products"]
        receipt["rowIdQueryColumnId"] = selected_id
        capture_query_evidence(receipt)
        self.current["contents"]["elements"].append({"id": "product-output", "name": "Product Output",
            "kind": "table", "source": {"kind": "table", "elementId": "summary-months"},
            "columns": [{"id": "product-key", "name": "Product Key", "formula": "0"}]})
        assigned = validate_full_assigned_rows(self.plan, self.current, self.observed)
        if omit_verified_field:
            del assigned["features"]["summary-products"][omit_verified_field]
        graph = {"status": "FULL_REGIONAL_FORMULAS_PENDING_ASSIGNED_BINDINGS",
            "sourceSha256": self.plan["sourceSha256"], "inputPlanHash": self.plan["inputPlanHash"],
            "formulaColumns": {"product-summary": {"Product Key":
                'Lookup([Summary Products/Product Line], "physical-summary-products-0", [Summary Products/Row ID])'}}}
        scaffold = {"documentId": self.current["workbookId"], "sourceSha256": self.plan["sourceSha256"],
            "logicalElement": "product-summary", "sourceElementId": "summary-months"}
        namespaces = {"Summary Products": {"elementId": "summary-products", "name": "summary-products",
            "columns": {"Product Line": {"id": receipt["columnIds"]["Product Line"], "name": "Product Line"},
                        "Row ID": row_id_binding}}}
        return prepare_output_contents_patch(graph, self.current, scaffold, assigned,
            output_element="product-summary", output_element_id="product-output",
            source_element_id="summary-months", assigned_columns={"Product Key": "product-key"},
            namespaces=namespaces)

    def test_sql_id_alias_cannot_rebind_custom_row_values_to_system_id(self):
        with self.assertRaisesRegex(ValueError, "reconciled input column.*Row ID"):
            self._product_row_id_patch({"id": "ID", "name": "ID"},
                selected_id="custom-row-id", selected_name="Row ID")

    def test_system_row_id_namespace_binds_the_queried_physical_column(self):
        patch = self._product_row_id_patch({"id": "ID", "name": "ID"})
        self.assertIn("[summary-products/ID]", patch["contents"]["elements"][-1]["columns"][0]["formula"])

    def test_custom_row_id_namespace_binds_the_queried_physical_column(self):
        patch = self._product_row_id_patch({"id": "custom-row-id", "name": "Row ID"},
            selected_id="custom-row-id", selected_name="Row ID")
        self.assertIn("[summary-products/Row ID]", patch["contents"]["elements"][-1]["columns"][0]["formula"])

    def test_named_row_id_namespace_cannot_substitute_for_queried_system_id(self):
        products = next(e for e in self.current["contents"]["elements"] if e["id"] == "summary-products")
        products["columns"].append({"id": "custom-row-id", "name": "Row ID"})
        with self.assertRaisesRegex(ValueError, "reconciled input column.*Row ID"):
            self._product_row_id_patch({"id": "custom-row-id", "name": "Row ID"})

    def test_row_id_namespace_requires_physical_provenance_in_verified_receipt(self):
        current = copy.deepcopy(self.current)
        for field in ("rowIdQueryColumnId", "rowIdColumnName"):
            with self.subTest(field=field):
                self.current = copy.deepcopy(current)
                with self.assertRaisesRegex(ValueError, "reconciled input column.*Row ID"):
                    self._product_row_id_patch({"id": "ID", "name": "ID"}, omit_verified_field=field)

    def test_query_request_and_returned_rows_must_match_capture(self):
        for feature in self.observed:
            for change in ("foreign-workbook", "foreign-element", "foreign-column", "changed-row", "missing-request", "missing-hash"):
                with self.subTest(feature=feature, change=change):
                    observed = copy.deepcopy(self.observed)
                    receipt = observed[feature]
                    if change == "foreign-workbook":
                        receipt["queryRequest"]["query"]["workbookId"] = "other-workbook"
                    elif change == "foreign-element":
                        receipt["queryRequest"]["query"]["sql"] = receipt["queryRequest"]["query"]["sql"].replace(
                            f'FROM "workbook"."{feature}"', 'FROM "workbook"."foreign"')
                    elif change == "foreign-column":
                        name, cid = next(iter(receipt["columnIds"].items()))
                        receipt["queryRequest"]["query"]["sql"] = receipt["queryRequest"]["query"]["sql"].replace(
                            f'"{cid}" AS "{name}"', f'"foreign" AS "{name}"')
                    elif change == "changed-row":
                        receipt["queryRows"][0]["ID"] = "different-physical-row"
                    else:
                        del receipt["queryRequest" if change == "missing-request" else "resultHash"]
                    with self.assertRaises(ValueError):
                        validate_full_assigned_rows(self.plan, self.current, observed)

    def test_matching_values_cannot_hide_query_column_reassignment(self):
        for feature, receipt in self.observed.items():
            for name in receipt["columnIds"]:
                with self.subTest(feature=feature, column=name):
                    changed = copy.deepcopy(self.observed)
                    changed[feature]["columnIds"][name] = "foreign-column"
                    with self.assertRaises(ValueError):
                        validate_full_assigned_rows(self.plan, self.current, changed)
            for field in ("columnIds", "rowIdQueryColumnId"):
                with self.subTest(feature=feature, missing=field):
                    changed = copy.deepcopy(self.observed)
                    del changed[feature][field]
                    with self.assertRaises(ValueError):
                        validate_full_assigned_rows(self.plan, self.current, changed)
            with self.subTest(feature=feature, column="ID"):
                changed = copy.deepcopy(self.observed)
                changed[feature]["rowIdQueryColumnId"] = "foreign-row-id"
                with self.assertRaises(ValueError):
                    validate_full_assigned_rows(self.plan, self.current, changed)


if __name__ == "__main__":
    unittest.main()
