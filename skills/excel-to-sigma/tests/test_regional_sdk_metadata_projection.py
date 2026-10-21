"""Source-bound SDK display evidence for Regional Rows read guards."""
import copy
import gzip
import json
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.native_exact_plan import digest
from lib.regional_grouped_model import (
    build_grouped_model, prepare_product_key_scaffold, prepare_product_key_patch,
    confirm_product_key_patch, prepare_product_rows_scaffold_program)

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "regional-grouped"


class SDKMetadataProjectionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        inputs = json.loads(gzip.decompress((FIXTURES / "input-plan.json.gz").read_bytes()))
        native = json.loads(gzip.decompress((FIXTURES / "native-plan.json.gz").read_bytes()))
        cls.model = build_grouped_model(inputs, native)

    def setUp(self):
        self.current = json.loads((FIXTURES / "saved-b4-readback.json").read_text())
        self.rows = next(e for e in self.current["contents"]["elements"]
                         if e.get("name") == "Regional Sales Rows")
        self.sales_id = self.rows["source"]["elementId"]
        self.page_id = self.current["contents"]["pages"][0]["id"]

    def receipt(self, source=None):
        current = source or self.current
        rows = next(e for e in current["contents"]["elements"]
                    if e.get("id") == self.rows["id"])
        columns = [{"id": c["id"], "name": c.get("name"),
                    "formula": c.get("formula")} for c in rows["columns"]]
        leaf = next(c for c in columns if c["id"] == "19bGkf4AMh")
        self.assertEqual(leaf["name"], "First Error Ordinal")
        leaf["name"] = "First Error Ordinal (1)"
        return {"documentId": current["workbookId"],
                "documentVersion": current["documentVersion"],
                "contentsHash": digest(current["contents"]),
                "elementId": rows["id"], "columns": columns,
                "reviewedNameDifferences": [{"id": leaf["id"],
                    "restName": "First Error Ordinal", "sdkName": leaf["name"]}]}

    def key(self, receipt=None):
        return prepare_product_key_scaffold(self.model, self.current,
            row_element_id=self.rows["id"], sales_element_id=self.sales_id,
            page_id=self.page_id, sdk_metadata_receipt=receipt)

    def test_key_default_strict_and_explicit_alias_preserves_physical_ids(self):
        original = self.key()
        projected = self.key(self.receipt())
        self.assertIn('"id": "19bGkf4AMh", "name": "First Error Ordinal",', original["program"])
        self.assertIn('"id": "19bGkf4AMh", "name": "First Error Ordinal (1)",', projected["program"])
        self.assertIn('"id": "1zF2g5Jr80", "name": "First Error Ordinal",', projected["program"])
        self.assertEqual(original["rowElementId"], projected["rowElementId"])
        self.assertEqual(original["contentsHash"], projected["contentsHash"])
        self.assertEqual(original["program"].count('add_column('), 2)
        self.assertEqual(projected["program"].count('add_column('), 2)

    def test_receipt_rejects_foreign_stale_missing_reordered_and_drift(self):
        base = self.receipt()
        bad = []
        for key, value in (("documentId", "foreign-workbook"),
                           ("documentVersion", base["documentVersion"] + 1),
                           ("contentsHash", "0" * 64), ("elementId", "foreign-rows")):
            receipt = copy.deepcopy(base)
            receipt[key] = value
            bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"].pop()
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"][0], receipt["columns"][1] = receipt["columns"][1], receipt["columns"][0]
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"][0]["id"] = receipt["columns"][1]["id"]
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"][0]["formula"] = "0"
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"][0]["name"] = "Unreviewed unrelated name"
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["reviewedNameDifferences"] = []
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["reviewedNameDifferences"][0]["id"] = "1zF2g5Jr80"
        bad.append(receipt)
        receipt = copy.deepcopy(base)
        receipt["columns"][0]["id"] = []
        bad.append(receipt)
        for index, receipt in enumerate(bad):
            with self.subTest(index=index), self.assertRaisesRegex(ValueError, "BLOCKED"):
                self.key(receipt)
        stale = copy.deepcopy(self.current)
        stale["documentVersion"] += 1
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            prepare_product_key_scaffold(self.model, stale,
                row_element_id=self.rows["id"], sales_element_id=self.sales_id,
                page_id=self.page_id, sdk_metadata_receipt=base)

    def test_product_rows_requires_fresh_new_receipt_after_key_patch(self):
        scaffold = self.key(self.receipt())
        current = copy.deepcopy(self.current)
        current["documentVersion"] += 1
        rows = next(e for e in current["contents"]["elements"] if e.get("id") == self.rows["id"])
        rows["columns"].extend([
            {"id": "assigned-product-key", "name": "Product Key", "formula": "0"},
            {"id": "assigned-month-key", "name": "Month Key", "formula": "0"}])
        patch = prepare_product_key_patch(self.model, current, scaffold,
            row_element_id=self.rows["id"], sales_element_id=self.sales_id,
            product_column_id="assigned-product-key", month_column_id="assigned-month-key")
        saved = copy.deepcopy(current)
        saved["documentVersion"] += 1
        saved["contents"] = patch["contents"]
        confirmed = confirm_product_key_patch(patch, saved)
        program = prepare_product_rows_scaffold_program(self.model, saved, confirmed,
            row_element_id=self.rows["id"], sales_element_id=self.sales_id,
            page_id=self.page_id, grid_row=130,
            sdk_metadata_receipt=self.receipt(saved))["program"]
        self.assertIn('"id": "19bGkf4AMh", "name": "First Error Ordinal (1)",', program)
        self.assertIn('"id": "assigned-product-key", "name": "Product Key",', program)
        self.assertIn('"id": "assigned-month-key", "name": "Month Key",', program)
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            prepare_product_rows_scaffold_program(self.model, saved, confirmed,
                row_element_id=self.rows["id"], sales_element_id=self.sales_id,
                page_id=self.page_id, grid_row=130,
                sdk_metadata_receipt=self.receipt())
        wrong = self.receipt(saved)
        wrong["columns"][-1]["formula"] = "0"
        with self.assertRaisesRegex(ValueError, "BLOCKED"):
            prepare_product_rows_scaffold_program(self.model, saved, confirmed,
                row_element_id=self.rows["id"], sales_element_id=self.sales_id,
                page_id=self.page_id, grid_row=130, sdk_metadata_receipt=wrong)


if __name__ == "__main__":
    unittest.main()
