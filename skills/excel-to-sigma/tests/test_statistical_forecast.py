"""Generated-data checks; external Excel replay lives outside normal CI."""
from datetime import date
from pathlib import Path
import sys
import unittest


LIB = Path(__file__).resolve().parents[1] / "scripts" / "lib"
sys.path.insert(0, str(LIB))
from statistical_forecast import (StatisticalPortError, calculate, excel_serial,
                                  forecast_linear, sample_stdev, trailing_three_average)
from statistical_runtime import emit_preview_program


def month(year, index):
    return date(year + (index - 1) // 12, (index - 1) % 12 + 1, 1)


def fixture():
    rows = [{"Month": month(2023, i), "SKU": sku, "Demand": 100 + 2*i + j*40}
            for i in range(1, 37)
            for j, sku in enumerate(("SKU-A100", "SKU-B200", "SKU-C300"))]
    targets = [month(2026, i) for i in range(1, 13)]
    return rows, targets


class StatisticalForecastTests(unittest.TestCase):
    def test_excel_serial_regression_uses_real_month_lengths(self):
        dates = [date(2024, 1, 1), date(2024, 2, 1), date(2024, 3, 1)]
        self.assertEqual([excel_serial(d) for d in dates], [45292, 45323, 45352])
        self.assertAlmostEqual(forecast_linear(date(2024, 4, 1), dates, [10, 20, 30]), 40.214735283, places=8)

    def test_source_fields_follow_editable_history_and_selection(self):
        rows, targets = fixture()
        result = calculate(rows, "SKU-B200", 12, .95, 2, targets)
        self.assertEqual(result["Status"], "ETS_UNAVAILABLE")
        self.assertFalse(result["ReadyToPublish"])
        self.assertIsNone(result["Summary"][0]["FY26ETSTotal"])
        self.assertIsNone(result["Summary"][0]["ReorderPoint"])
        self.assertEqual(result["Periods"][0]["Selected3MMA"], 210)
        self.assertEqual(result["IntendedSequenceDescending"], list(range(12, 0, -1)))
        self.assertEqual(result["SourceK4"], "#NAME?")
        before = result["Periods"][0]["Linear"]["SKU-A100"]
        rows[-3]["Demand"] += 100  # Last A100 observation in interleaved long rows.
        changed = calculate(rows, "SKU-A100", 12, .95, 2, targets)
        self.assertNotEqual(before, changed["Periods"][0]["Linear"]["SKU-A100"])
        self.assertEqual(changed["Periods"][0]["Selected3MMA"], (168+170+272)/3)

    def test_sample_stdev_and_trailing_average_skip_blanks(self):
        self.assertAlmostEqual(sample_stdev([1, None, 3]), 2**.5)
        self.assertEqual(trailing_three_average([1, 2, None, 4]), 3)

    def test_guards_and_candidate_remains_unqualified(self):
        rows, targets = fixture()
        with self.assertRaisesRegex(StatisticalPortError, "MISSING_OR_INVALID_SELECTED_SKU"):
            calculate(rows, "SKU-Z", 12, .95, 2, targets)
        with self.assertRaisesRegex(StatisticalPortError, "INVALID_INVENTORY_PARAMETERS"):
            calculate(rows, "SKU-A100", 12, 1, 2, targets)
        with self.assertRaisesRegex(StatisticalPortError, "NONCONSECUTIVE_TARGET_MONTHS"):
            calculate(rows, "SKU-A100", 12, .95, 2, targets[:-1] + [date(2027, 2, 1)])
        result = calculate(rows, "SKU-A100", 12, .95, 2, targets,
                           ets_forecaster=lambda values, s, n: [values[-1]] * n)
        self.assertEqual(result["Status"], "REVIEW_REQUIRED")
        self.assertFalse(result["ReadyToPublish"])

    def test_generated_sigma_preview_reads_current_rows_and_stays_blocked(self):
        binding = {role: {"title": role, "columns": {name: {"identifier": name, "rowKey": name}
                    for name in names}} for role, names in {
                        "history": ("Month", "SKU", "Demand"),
                        "parameters": ("Key", "Value"), "targets": ("Period", "Month")}.items()}
        program = emit_preview_program(binding)
        compile(program["code"], "<statistical Sigma preview>", "exec")
        rows, targets = fixture()

        class Row:
            def __init__(self, data): self.data = data
            def as_dict(self): return self.data

        class Frame:
            def __init__(self, data):
                self.data = data
                self.columns = list(data[0])
            def collect(self): return [Row(row) for row in self.data]

        class Sigma:
            def __init__(self, frames): self.frames, self.outputs = frames, []
            def get_element(self, name): return Frame(self.frames[name])
            def output(self, name, frame): self.outputs.append((name, frame))

        class Session:
            def create_dataframe(self, values, schema): return (values, schema)

        frames = {"history": rows,
                  "parameters": [{"Key": k, "Value": v} for k, v in {
                      "Selected SKU": "SKU-A100", "Seasonality (months)": 12,
                      "Service level": .95, "Lead time (months)": 2}.items()],
                  "targets": [{"Period": i, "Month": target} for i, target in enumerate(targets, 1)]}
        sigma = Sigma(frames)
        exec(program["code"], {"sigma": sigma, "session": Session()})
        values, schema = sigma.outputs[0][1]
        self.assertEqual(len(values), 36)
        self.assertEqual(values[0][0:3], ["ETS_UNAVAILABLE", False, "EXCEL_ETS_PROVIDER_UNQUALIFIED"])
        self.assertIsNone(values[0][7])  # ETS remains null.
        self.assertIn('"ReadyToPublish" BOOLEAN', schema)
        frames["parameters"][0]["Value"] = "SKU-Z"
        sigma = Sigma(frames)
        exec(program["code"], {"sigma": sigma, "session": Session()})
        self.assertEqual(sigma.outputs[0][1][0][0][0], "FAILURE")


if __name__ == "__main__":
    unittest.main()
