import importlib.util
import copy
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]


def load_inference():
    spec = importlib.util.spec_from_file_location("formula_provenance_inference",
                                                  SCRIPTS / "infer-canonical-formulas.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


INFERENCE = load_inference()
_builder_spec = importlib.util.spec_from_file_location("research_builder_for_dag",
                                                       SCRIPTS / "build-research-model.py")
BUILDER = importlib.util.module_from_spec(_builder_spec)
_builder_spec.loader.exec_module(BUILDER)
from migration_manifest import validate_manifest  # noqa: E402


class TestResearchInference(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fixture = ROOT / "tests" / "fixtures" / "fixture-research-model.xlsx"
        cls.plan = INFERENCE.infer(cls.fixture, sheet="FY results")

    def test_all_source_formulas_are_accounted_without_cached_fallbacks(self):
        accounting = self.plan["formulaCells"]
        self.assertEqual(len(accounting), 343)
        self.assertEqual(len({item["id"] for item in accounting}), 343)
        self.assertEqual(
            {status: sum(item["mappingStatus"] == status for item in accounting)
             for status in ("live", "verification", "cached", "unresolved")},
            {"live": 327, "verification": 16, "cached": 0, "unresolved": 0},
        )
        self.assertEqual(self.plan["parity"], {"live": 327, "frozen": 0, "audit": 16})
        self.assertEqual(len(self.plan["migrationManifest"]["formulaCells"]), 343)
        self.assertTrue(validate_manifest(self.plan["migrationManifest"], self.fixture))
        mappings = [item["targetMapping"] for item in accounting if item["targetMapping"]]
        self.assertTrue(all(not ({"elementId", "columnId", "lineId", "checkId"} & set(mapping))
                            for mapping in mappings))
        self.assertTrue(any(mapping.get("logicalElementId") == "calc" for mapping in mappings))
        self.assertTrue(any(mapping.get("logicalCheckId") for mapping in mappings))

    def test_false_row_cycle_preserves_each_cells_actual_formula_or_literal(self):
        lines = {line["row"]: line for line in self.plan["lines"]}
        self.assertIn("2025", lines[13]["literalInputs"])
        self.assertIn("2025", lines[11]["periodFormulas"])
        self.assertNotEqual(lines[11]["kind"], "input")
        self.assertNotIn("plug-cycle", " ".join(lines[11].get("reasons", [])))

    def test_cross_sheet_drivers_are_source_bound_inputs_and_year_formulas_stay_live(self):
        inputs = self.plan["sourceInputs"]
        drivers = [item for item in inputs if item["role"] == "cross_sheet_driver"]
        self.assertEqual([item["sourceCellId"] for item in drivers],
                         [f"Model Inputs!C{row}" for row in range(2, 8)])
        revenue = next(line for line in self.plan["lines"] if line["row"] == 8)
        self.assertEqual(set(revenue["periodFormulas"]),
                         {str(year) for year in range(2025, 2031)})
        self.assertIn("[i_model_inputs_c2]", revenue["periodFormulas"]["2025"])

    def test_period_offsets_sum_percent_and_precision_render_as_live_formulas(self):
        lines = {line["row"]: line for line in self.plan["lines"]}
        self.assertEqual(lines[9]["periodFormulas"]["2020"],
                         "[i_fy_results_c8] / [i_fy_results_b8] - 1")
        self.assertNotIn("[c", lines[16]["periodFormulas"]["2025"])
        self.assertNotIn("[c", lines[63]["periodFormulas"]["2019"])
        self.assertNotIn("[c", lines[64]["periodFormulas"]["2019"])
        self.assertTrue(INFERENCE._safe_sigma("Lag([c8], 1, [Year])"))

    def test_emitted_period_formulas_are_a_static_dag_without_self_or_sibling_refs(self):
        for line in self.plan["lines"]:
            for formula in (line.get("periodFormulas") or {}).values():
                self.assertNotRegex(formula, r"\[c\d+\]")
                self.assertNotIn("Lag(", formula)
                self.assertNotIn("Lead(", formula)

    def test_generated_target_calculation_columns_form_a_dag(self):
        plan = copy.deepcopy(self.plan)
        plan["targetScope"] = {"connectionId": "connection", "folderId": "folder"}
        _dm, data_lines = BUILDER.build_dm(plan, "connection", "folder", "Research")
        workbook = BUILDER.build_wb(
            plan, "dm", {"data": "data", "linedim": "linedim"}, data_lines,
            "folder", "Research")
        calc = next(element for page in workbook["pages"] for element in page.get("elements", [])
                    if element.get("id") == "calc")
        calc_ids = {column["id"] for column in calc["columns"]}
        edges = {column["id"]: set(__import__("re").findall(r"\[([^]/]+)\]", column["formula"])) & calc_ids
                 for column in calc["columns"]}
        self.assertTrue(all(node not in deps for node, deps in edges.items()))
        self.assertFalse(edges["c11"] & {"c13"})
        self.assertFalse(edges["c13"] & {"c11"})
        visiting, visited = set(), set()
        def visit(node):
            self.assertNotIn(node, visiting, f"generated target cycle at {node}")
            if node in visited:
                return
            visiting.add(node)
            for dependency in edges[node]:
                visit(dependency)
            visiting.remove(node); visited.add(node)
        for node in edges:
            visit(node)

    def test_audit_formulas_are_computed_verification_checks(self):
        checks = self.plan["auditChecks"]
        self.assertEqual(len(checks), 16)
        self.assertEqual([item["computedValue"] for item in checks if item["address"].startswith("F")],
                         ["PASS", "PASS", "PASS", "PASS"])
        self.assertTrue(all(item["mappingStatus"] == "verification" for item in checks))


if __name__ == "__main__":
    unittest.main()
