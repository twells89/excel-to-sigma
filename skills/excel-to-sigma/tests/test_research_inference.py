import importlib.util
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

    def test_false_row_cycle_preserves_each_cells_actual_formula_or_literal(self):
        lines = {line["row"]: line for line in self.plan["lines"]}
        self.assertIn("2025", lines[13]["literalInputs"])
        self.assertIn("2025", lines[11]["periodFormulas"])
        self.assertNotEqual(lines[11]["kind"], "input")
        self.assertNotIn("plug-cycle", " ".join(lines[11].get("reasons", [])))

    def test_cross_sheet_drivers_are_source_bound_inputs_and_year_formulas_stay_live(self):
        inputs = self.plan["sourceInputs"]
        self.assertEqual([item["sourceCellId"] for item in inputs],
                         [f"Model Inputs!C{row}" for row in range(2, 8)])
        revenue = next(line for line in self.plan["lines"] if line["row"] == 8)
        self.assertEqual(set(revenue["periodFormulas"]),
                         {str(year) for year in range(2025, 2031)})
        self.assertIn("[i_model_inputs_c2]", revenue["periodFormulas"]["2025"])

    def test_period_offsets_sum_percent_and_precision_render_as_live_formulas(self):
        lines = {line["row"]: line for line in self.plan["lines"]}
        self.assertEqual(lines[9]["periodFormulas"]["2020"], "[c8] / Lag([c8], 1, [Year]) - 1")
        self.assertEqual(lines[16]["periodFormulas"]["2025"],
                         "Lag([c16], 1, [Year]) * (104.5 * 0.01)")
        self.assertEqual(lines[63]["periodFormulas"]["2019"], "Sum(([c61] + [c62]))")
        self.assertEqual(lines[64]["periodFormulas"]["2019"], "[c63] / [c31]")
        self.assertTrue(INFERENCE._safe_sigma("Lag([c8], 1, [Year])"))

    def test_audit_formulas_are_computed_verification_checks(self):
        checks = self.plan["auditChecks"]
        self.assertEqual(len(checks), 16)
        self.assertEqual([item["computedValue"] for item in checks if item["address"].startswith("F")],
                         ["PASS", "PASS", "PASS", "PASS"])
        self.assertTrue(all(item["mappingStatus"] == "verification" for item in checks))


if __name__ == "__main__":
    unittest.main()
