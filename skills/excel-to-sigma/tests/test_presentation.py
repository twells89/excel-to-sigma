#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
WORKSPACE = ROOT.parents[3]
sys.path.insert(0, str(SCRIPTS / "lib"))

import presentation  # noqa: E402


FIXTURES = WORKSPACE / "outputs"


class TestPresentationExtraction(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.plans = {
            path.name: presentation.extract_presentation(path)
            for path in sorted(FIXTURES.glob("fixture-*.xlsx"))
        }

    def test_extracts_all_native_charts_with_source_relationships_and_series_order(self):
        charts = [chart for plan in self.plans.values() for chart in plan["charts"]]
        self.assertEqual(len(charts), 10)
        self.assertTrue(all(chart["sourceLocation"]["worksheet"] for chart in charts))
        self.assertTrue(all(chart["sourceLocation"]["drawing"] for chart in charts))
        self.assertEqual(
            [series["name"] for series in self.plans["fixture-forecast-model.xlsx"]["charts"][0]["series"]],
            ["Revenue", "EBITDA", "Net Income"],
        )
        self.assertEqual(
            [series["name"] for series in self.plans["fixture-research-model.xlsx"]["charts"][1]["series"]],
            ["EPS", "Net debt / EBITDA"],
        )
        self.assertEqual(
            self.plans["fixture-sales-pipeline.xlsx"]["charts"][1]["categoryRef"],
            "'Overview'!$D$12:$D$21",
        )
        self.assertEqual(
            self.plans["fixture-inventory-reorder.xlsx"]["charts"][1]["unit"], "currency")

    def test_extracts_validation_contracts_and_exact_whole_number_list(self):
        inventory = self.plans["fixture-inventory-reorder.xlsx"]
        lead_days = next(rule for rule in inventory["validations"] if rule["range"] == "E2:E37")
        self.assertEqual(lead_days["sourceType"], "whole")
        self.assertEqual(lead_days["emission"], {"values": list(range(1, 91))})
        for invalid in (1.5, 0, 91):
            self.assertFalse(presentation.validation_accepts(lead_days, invalid))
        self.assertTrue(presentation.validation_accepts(lead_days, 1))
        self.assertTrue(presentation.validation_accepts(lead_days, 90))

        sales = self.plans["fixture-sales-pipeline.xlsx"]
        probability = next(rule for rule in sales["validations"] if rule["range"] == "I2:I49")
        self.assertEqual(probability["emission"], {"range": {"min": 0.0, "max": 1.0}})
        self.assertFalse(probability["allowBlank"])

    def test_preserves_conditional_rule_priority_relative_references_and_scales(self):
        rules = [rule for plan in self.plans.values() for rule in plan["conditionalFormats"]]
        priorities = {(rule["sheet"], rule["range"], rule["priority"]) for rule in rules}
        self.assertIn(("Expense Register", "O2:O49", 1), priorities)
        self.assertIn(("Expense Register", "O2:O49", 3), priorities)
        self.assertTrue(all("relativeReference" in rule for rule in rules))
        scales = [rule for rule in rules if rule["sourceType"] == "colorScale"]
        self.assertGreaterEqual(len(scales), 3)
        self.assertTrue(all(rule["status"] == "supported" for rule in scales))
        self.assertTrue(all(rule["emission"]["type"] == "backgroundScale" for rule in scales))
        self.assertTrue(all(rule["emission"]["domain"]["mid"].startswith("Percentile(") for rule in scales))

    def test_extracts_source_precision_dates_and_all_formula_provenance(self):
        expected = {
            "fixture-sales-pipeline.xlsx": 170,
            "fixture-expense-approval.xlsx": 632,
            "fixture-inventory-reorder.xlsx": 442,
            "fixture-forecast-model.xlsx": 278,
            "fixture-research-model.xlsx": 343,
        }
        self.assertEqual(
            {name: len(plan["formulaProvenance"]) for name, plan in self.plans.items()}, expected)
        self.assertEqual(
            self.plans["fixture-sales-pipeline.xlsx"]["semanticFormats"]["Pipeline!B2"]["dateGrain"],
            "day",
        )
        self.assertEqual(
            self.plans["fixture-sales-pipeline.xlsx"]["semanticFormats"]["Pipeline!I2"]["precision"],
            0,
        )
        sales_timing = next(
            item for item in self.plans["fixture-sales-pipeline.xlsx"]["formulaProvenance"]
            if item["cell"] == "Pipeline!N2"
        )
        self.assertNotIn("\\\\", sales_timing["formula"])
        chart_cells = [cell for plan in self.plans.values() for chart in plan["charts"]
                       for series in chart["series"] for cell in series["sourceFormulaCells"]]
        self.assertTrue(chart_cells)


class TestPresentationEmission(unittest.TestCase):
    def test_business_input_programs_seed_only_literal_columns(self):
        expectations = {
            "fixture-forecast-model.xlsx": ("Line Item", None),
            "fixture-sales-pipeline.xlsx": ("Probability", "Weighted ACV"),
            "fixture-inventory-reorder.xlsx": ("Lead Days", "Reorder Qty"),
            "fixture-expense-approval.xlsx": ("Approval Status", "Exception"),
        }
        for filename, (editable, derived) in expectations.items():
            plan = presentation.extract_presentation(FIXTURES / filename)
            spec = presentation.extract_business_input(plan)
            program = presentation.emit_business_input_program(spec, presentation.fixture_bindings(plan)["businessInput"])
            column_names = [column["name"] for column in spec["columns"]]
            self.assertIn(editable, column_names)
            if derived is not None:
                self.assertNotIn(derived, column_names)
            self.assertIn(f"add_editable_column({editable!r}", program["program"])
            self.assertIn("set_edit_mode('published-explore')", program["program"])
            if derived is not None:
                self.assertNotIn(f"add_editable_column({derived!r}", program["program"])
                self.assertTrue(any(column["name"] == derived for column in spec["derivedColumns"]))
                if filename == "fixture-sales-pipeline.xlsx":
                    self.assertIn("add_formula_column('Weighted ACV'", program["program"])
                    self.assertIn("add_formula_column('Cycle Days'", program["program"])
                if filename == "fixture-inventory-reorder.xlsx":
                    self.assertIn("sigma.formula.round_up", program["program"])
                    self.assertEqual(program["liveFormulaColumns"],
                                     ["Available", "Reorder Point", "Days Cover", "Reorder Qty", "Status"])
            retry_binding = {**presentation.fixture_bindings(plan)["businessInput"], "elementId": "existing-input"}
            retry = presentation.emit_business_input_program(spec, retry_binding)
            self.assertEqual(retry["mode"], "reconcile-existing")
            self.assertNotIn("program", retry)

    def test_research_reuses_existing_shared_input_identity(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-research-model.xlsx")
        spec = presentation.extract_business_input(plan)
        program = presentation.emit_business_input_program(spec, {
            "pageId": "input-page", "elementId": "0CI0ooHYzw", "rowId": "row-1",
            "scopeColumnId": "scope-id", "planHash": "plan", "stagingReceiptHash": "receipt",
        })
        self.assertEqual(program["mode"], "extend-existing")
        self.assertEqual(program["elementId"], "0CI0ooHYzw")
        self.assertEqual(program["expectedExistingEditableColumns"], 134)
        self.assertEqual(program["expectedEditableColumns"], 146)
        self.assertEqual(len(program["additionalColumns"]), 12)
        self.assertEqual(program["additionalColumns"][0]["sourceCellId"], "FY results!B56")
        self.assertEqual(program["additionalColumns"][-1]["sourceCellId"], "FY results!M56")
        self.assertEqual(program["additionalColumns"][0]["name"], "i_fy_results_b56")
        self.assertIn("table.update_rows([('row-1'", program["program"])

    def test_sales_and_expense_consume_reviewed_operational_wiring_builder(self):
        sales_plan = presentation.extract_presentation(FIXTURES / "fixture-sales-pipeline.xlsx")
        sales_binding = presentation.fixture_bindings(sales_plan)["businessInput"]
        sales_binding["elementId"] = "sales-input"
        sales = presentation.build_business_operational_wiring(sales_plan, sales_binding)
        sales_by_id = {item["id"]: item for item in sales["fragments"]["elements"]}
        timing = next(column for column in sales_by_id["salesWithSettings"]["columns"]
                      if column["name"] == "TimingStatus")
        self.assertIn('DateTrunc("day", [Settings/AsOfDate])', timing["formula"])

        expense_plan = presentation.extract_presentation(FIXTURES / "fixture-expense-approval.xlsx")
        expense_binding = presentation.fixture_bindings(expense_plan)["businessInput"]
        expense_binding["elementId"] = "expense-input"
        expense = presentation.build_business_operational_wiring(expense_plan, expense_binding)
        expense_by_id = {item["id"]: item for item in expense["fragments"]["elements"]}
        expense_calc = expense_by_id["expensesWithPolicy"]
        self.assertEqual(expense_calc["source"]["kind"], "join")
        exception = next(column for column in expense_calc["columns"] if column["name"] == "Exception")
        self.assertIn("[Policy/Per-item Limit]", json.dumps(expense_calc))
        self.assertIn("MISSING RECEIPT", exception["formula"])

    def test_emits_live_charts_formats_rules_and_three_usable_pages(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-sales-pipeline.xlsx")
        bindings = presentation.fixture_bindings(plan)
        artifact = presentation.emit_presentation(plan, bindings)

        charts = [element for element in artifact["elements"] if element["kind"].endswith("-chart")]
        self.assertEqual([chart["kind"] for chart in charts], ["bar-chart", "line-chart"])
        self.assertEqual(charts[0]["yAxis"]["columnIds"], ["chart_1_series_1"])
        self.assertEqual(charts[0]["columns"][1]["formula"], "Sum([Pipeline/Weighted ACV])")
        self.assertEqual(charts[0]["xAxis"]["sort"]["by"], "chart_1_order")
        self.assertNotIn("cached", json.dumps(charts).lower())

        self.assertEqual([page["name"] for page in artifact["pages"]], ["Overview", "Edit", "Detail"])
        self.assertEqual(artifact["pageRoles"], {"overview": "sales-overview", "edit": "sales-edit", "detail": "sales-detail"})
        self.assertIn("Probability", artifact["editableColumns"])
        self.assertIn("ACV", artifact["editableColumns"])
        self.assertIn("Weighted ACV", artifact["principalCalculatedColumns"])
        self.assertIn("Timing Status", artifact["principalCalculatedColumns"])
        self.assertLessEqual(len(artifact["mainGridColumns"]), 12)
        self.assertGreater(len(artifact["detailColumns"]), len(artifact["mainGridColumns"]))
        self.assertEqual(len(artifact["formulaAccounting"]), len(plan["formulaProvenance"]))
        self.assertIn("mapped-live-chart-check",
                      {item["targetStatus"] for item in artifact["formulaAccounting"]})

        probability_patch = artifact["elementPatches"]["sales-input"]["columnUpdates"]["probability"]
        self.assertEqual(probability_patch["range"], {"min": 0.0, "max": 1.0})

        page_spans = {}
        for page_id, body in re.findall(r'<Page\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</Page>', artifact["layout"], re.S):
            spans = []
            for element_id, c1, c2, r1, r2 in re.findall(
                r'<Element elementId="([^"]+)" gridColumn="(\d+) / (\d+)" gridRow="(\d+) / (\d+)"/>', body
            ):
                span = tuple(map(int, (c1, c2, r1, r2)))
                for other_id, other in spans:
                    overlap = span[0] < other[1] and other[0] < span[1] and span[2] < other[3] and other[2] < span[3]
                    self.assertFalse(overlap, f"{page_id}: {element_id} overlaps {other_id}")
                spans.append((element_id, span))
            page_spans[page_id] = spans
        self.assertEqual(set(page_spans), {"sales-overview", "sales-edit", "sales-detail"})

    def test_research_edit_view_keeps_shared_drivers_prominent_and_133_fields_in_detail(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-research-model.xlsx")
        bindings = presentation.fixture_bindings(plan)
        bindings["sharedInputColumns"] = [f"Source Cell {i}" for i in range(1, 134)]
        bindings["prominentEditableColumns"] = [f"Source Cell {i}" for i in range(1, 7)]
        artifact = presentation.emit_presentation(plan, bindings)

        self.assertEqual(len(artifact["editColumns"]), 6)
        self.assertEqual(len(artifact["detailColumns"]), 133)
        self.assertEqual(artifact["sharedInputScope"], "shared")
        main = {column["name"]: column for element in artifact["elements"]
                if element["id"] == "research-main-view" for column in element["columns"]}
        detail_view = next(element for element in artifact["elements"]
                           if element["id"] == "research-detail-view")
        self.assertEqual(len(detail_view["columns"]), 36)
        self.assertNotIn("Line Item", {column["name"] for column in detail_view["columns"]})
        self.assertEqual(main["Revenue"]["format"]["currencySymbol"], "$")
        self.assertEqual(main["EPS"]["format"]["formatString"], ",.2f")
        self.assertEqual(main["Net debt / EBITDA (x)"]["format"]["suffix"], "x")
        self.assertFalse(any(item["kind"] == "mixed-unit-row-format"
                             for item in artifact["unresolved"]))
        updates = artifact["elementPatches"]["research-input-existing"]["columnUpdates"]
        self.assertEqual(
            {name for name, update in updates.items() if update.get("range") == {"min": 0.0, "max": 0.25}},
            {f"i_model_inputs_c{row}" for row in range(2, 8)},
        )

    def test_forecast_manual_validation_covers_each_month_column(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-forecast-model.xlsx")
        artifact = presentation.emit_presentation(plan, presentation.fixture_bindings(plan))
        updates = artifact["elementPatches"]["forecast-input"]["columnUpdates"]
        self.assertEqual(
            {name for name, update in updates.items()
             if update.get("range") == {"min": 0.0, "max": 1000000.0}},
            {"jul_26", "aug_26", "sep_26", "oct_26", "nov_26", "dec_26"},
        )

    def test_unbound_source_chart_remains_explicitly_unresolved(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-expense-approval.xlsx")
        bindings = presentation.fixture_bindings(plan)
        bindings["charts"].pop(plan["charts"][0]["id"])
        artifact = presentation.emit_presentation(plan, bindings)

        self.assertEqual(len([e for e in artifact["elements"] if e["kind"].endswith("-chart")]), 1)
        unresolved = next(item for item in artifact["unresolved"] if item["sourceId"] == plan["charts"][0]["id"])
        self.assertEqual(unresolved["reason"], "missing live target binding")

    def test_full_contents_integration_applies_validation_and_rule_patches(self):
        plan = presentation.extract_presentation(FIXTURES / "fixture-inventory-reorder.xlsx")
        bindings = presentation.fixture_bindings(plan)
        artifact = presentation.emit_presentation(plan, bindings)
        generated_ids = {element["id"] for element in artifact["elements"]}
        element_ids = set(artifact["referencedElementIds"]) | (set(artifact["elementPatches"]) - generated_ids)
        patch_columns = {
            element_id: set(patch["columnUpdates"])
            | {column_id for rule in patch["conditionalFormats"] for column_id in rule["columnIds"]}
            for element_id, patch in artifact["elementPatches"].items()
        }
        contents = {
            "schemaVersion": 1, "kind": "workbook",
            "pages": [{"id": "existing", "name": "Existing"}],
            "elements": [
                {"id": element_id, "kind": "input-table" if element_id in artifact["elementPatches"] else "table",
                 "columns": [{"id": column_id, "type": "number"} for column_id in sorted(patch_columns.get(element_id, set()))],
                 "source": {"kind": "table", "elementId": "source"}, "inputMode": "edit"}
                for element_id in sorted(element_ids)
            ],
            "layout": '<?xml version="1.0" encoding="utf-8"?>\n<Page id="existing"></Page>',
        }
        merged = presentation.integrate_presentation(contents, artifact)
        lead_days = next(
            column for element in merged["elements"] if element["id"] == "inventory-input"
            for column in element["columns"] if column["id"] == "lead_days"
        )
        self.assertEqual(lead_days["values"], list(range(1, 91)))
        self.assertEqual(len(merged["pages"]), 4)
        self.assertEqual(len(re.findall(r"<Page\b", merged["layout"])), 4)

        repeated = presentation.integrate_presentation(merged, artifact)
        self.assertEqual(len(repeated["pages"]), 4)
        self.assertEqual(len(re.findall(r"<Page\b", repeated["layout"])), 4)
        self.assertEqual(len({element["id"] for element in repeated["elements"]}), len(repeated["elements"]))

        foreign = json.loads(json.dumps(contents))
        foreign["elements"].append({"id": artifact["elements"][0]["id"], "kind": "text", "body": "foreign"})
        with self.assertRaisesRegex(ValueError, "foreign contents"):
            presentation.integrate_presentation(foreign, artifact)


if __name__ == "__main__":
    unittest.main()
