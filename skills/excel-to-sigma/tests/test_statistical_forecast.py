"""Generated-data checks; external Excel replay lives outside normal CI."""
from datetime import date
from pathlib import Path
import sys
import json
import runpy
import types
import unittest
from unittest.mock import patch
import warnings


LIB = Path(__file__).resolve().parents[1] / "scripts" / "lib"
sys.path.insert(0, str(LIB))
from statistical_forecast import (StatisticalPortError, calculate, excel_serial,
                                  forecast_linear, sample_stdev, trailing_three_average,
                                  typed_output_value)
from statistical_runtime import CANDIDATE_REQUIRED, emit_preview_program
from statistical_ets_candidate import CandidateFitError, forecast_with_diagnostics


def month(year, index):
    return date(year + (index - 1) // 12, (index - 1) % 12 + 1, 1)


def fixture():
    rows = [{"Month": month(2023, i), "SKU": sku, "Demand": 100 + 2*i + j*40}
            for i in range(1, 37)
            for j, sku in enumerate(("SKU-A100", "SKU-B200", "SKU-C300"))]
    targets = [month(2026, i) for i in range(1, 13)]
    return rows, targets


class StatisticalForecastTests(unittest.TestCase):
    def test_operator_frames_fit_sigma_grid_and_partial_continuation_reuses_table(self):
        operator = runpy.run_path(str(LIB.parent / "statistical-operator.py"))
        for role, (_, _, _, frame) in operator["TABLES"].items():
            with self.subTest(role=role):
                x, y, width, height = frame
                self.assertGreaterEqual(x, 1)
                self.assertGreaterEqual(y, 1)
                self.assertGreater(width, 0)
                self.assertGreater(height, 0)
                self.assertLessEqual(x + width - 1, 24)
        continuation = operator["phase1_continue_parameters"]("observed-page", "observed-table")
        self.assertIn("page.get_element('observed-table')", continuation)
        self.assertNotIn("add_input_table", continuation)
        self.assertIn("set_frame(1, 15, 24, 8)", continuation)
        compile(continuation, "<partial continuation>", "exec")

    def test_candidate_sigma_preview_is_typed_current_and_review_only(self):
        binding = {role: {"title": role, "columns": {name: {"identifier": name, "rowKey": name}
                    for name in names}} for role, names in CANDIDATE_REQUIRED.items()}
        program = emit_preview_program(binding, candidate_ets=True)
        compile(program["code"], "<candidate preview>", "exec")
        rows, targets = fixture()
        frames = {"history": rows,
                  "parameters": [{"Scope": "Parameters", "SelectedSKU": "SKU-A100",
                                  "Seasonality": 12, "ServiceLevel": .95, "LeadTime": 2}],
                  "targets": [{"Period": i, "Month": target} for i, target in enumerate(targets, 1)]}

        class Row(dict):
            def as_dict(self): return self
        class Frame:
            def __init__(self, data): self.data, self.columns = data, list(data[0])
            def collect(self): return [Row(row) for row in self.data]
        class Sigma:
            def __init__(self): self.outputs = []
            def get_element(self, title): return Frame(frames[title])
            def output(self, name, frame): self.outputs.append((name, frame))
        class Session:
            def create_dataframe(self, values, schema): return (values, schema)
        class Fit:
            def __init__(self, values, converged=True, invalid=False):
                self.values = values
                self.mle_retvals = {"converged": converged, "warnflag": 0 if converged else 2,
                                    "iterations": 5, "fcalls": 10}
                self.invalid = invalid
            def forecast(self, horizon):
                return [float("nan") if self.invalid else self.values[-1]] * horizon
        class ETSModel:
            fail = False
            invalid = False
            def __init__(self, values, **kwargs): self.values = values
            def fit(self, **kwargs): return Fit(self.values, not self.fail, self.invalid)
        fake = types.ModuleType("statsmodels.tsa.exponential_smoothing.ets")
        fake.ETSModel = ETSModel
        def execute():
            sigma = Sigma()
            with patch.dict(sys.modules, {"statsmodels": types.ModuleType("statsmodels"),
                                          "statsmodels.tsa": types.ModuleType("statsmodels.tsa"),
                                          "statsmodels.tsa.exponential_smoothing": types.ModuleType("statsmodels.tsa.exponential_smoothing"),
                                          "statsmodels.tsa.exponential_smoothing.ets": fake}), \
                 patch("importlib.metadata.version", side_effect=lambda name: {
                     "statsmodels": "0.15.0", "scipy": "1.18.1"}[name]):
                exec(program["code"], {"sigma": sigma, "session": Session()})
            return sigma.outputs[0][1]

        values, schema = execute()
        columns = [name for name, _ in program["schema"]]
        first = dict(zip(columns, values[0]))
        self.assertEqual(len(values), 36, values[0])
        self.assertEqual((first["Status"], first["ReadyToPublish"], first["ErrorCode"]),
                         ("REVIEW_REQUIRED", False, "CANDIDATE_REVIEW_REQUIRED"))
        self.assertEqual(first["ETS"], rows[-3]["Demand"])
        self.assertIsNotNone(first["FY26ETSTotal"])
        self.assertIsNotNone(first["ReorderPoint"])
        self.assertEqual(first["ModelIdentity"], program["modelIdentity"])
        self.assertEqual(json.loads(first["RuntimePackages"]),
                         {"statsmodels": "0.15.0", "scipy": "1.18.1"})
        self.assertEqual(set(json.loads(first["FitDiagnostics"])), set(("SKU-A100", "SKU-B200", "SKU-C300")))
        self.assertIn('"FitDiagnostics" STRING', schema)
        prior_hash = first["InputHash"]
        rows[-3]["Demand"] += 100
        changed, _ = execute()
        changed_first = dict(zip(columns, changed[0]))
        self.assertNotEqual(changed_first["InputHash"], prior_hash)
        self.assertNotEqual(changed_first["ETS"], first["ETS"])
        self.assertNotEqual(changed_first["PreviewRunId"], first["PreviewRunId"])
        ETSModel.fail = True
        failed, _ = execute()
        self.assertEqual(len(failed), 1)
        failure = dict(zip(columns, failed[0]))
        self.assertEqual((failure["Status"], failure["ErrorCode"]),
                         ("FAILURE", "CANDIDATE_NONCONVERGED_FIT"))
        self.assertIsNone(failure["ETS"])
        self.assertFalse(failure["ReadyToPublish"])
        ETSModel.fail, ETSModel.invalid = False, True
        invalid, _ = execute()
        self.assertEqual(dict(zip(columns, invalid[0]))["ErrorCode"],
                         "CANDIDATE_INVALID_FORECAST_OUTPUT")
        ETSModel.invalid = False
        original_fit = ETSModel.fit
        def editing_fit(self, **kwargs):
            rows[-3]["Demand"] += 1
            return original_fit(self, **kwargs)
        with patch.object(ETSModel, "fit", editing_fit):
            stale, _ = execute()
        self.assertEqual(len(stale), 1)
        self.assertEqual(dict(zip(columns, stale[0]))["ErrorCode"],
                         "INPUT_CHANGED_DURING_RUN")

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

    def test_finite_history_that_overflows_derived_values_fails_closed(self):
        rows, targets = fixture()
        for row in rows:
            row["Demand"] = 1e308
        with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_LINEAR_Y_MEAN"):
            calculate(rows, "SKU-A100", 12, .95, 2, targets)
        with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_STDEV_VARIANCE"):
            sample_stdev([1e308, -1e308])
        with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_TRAILING_AVERAGE"):
            trailing_three_average([1e308, 1e308, 1e308])

    def test_derived_safety_and_provider_annual_sum_fail_without_preview(self):
        rows, targets = fixture()
        for row in rows:
            if row["SKU"] == "SKU-A100" and row["Month"] == date(2023, 1, 1):
                row["Demand"] = 8e153
            if row["SKU"] == "SKU-A100" and row["Month"] == date(2023, 2, 1):
                row["Demand"] = -8e153
        with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_SAFETY_STOCK"):
            calculate(rows, "SKU-A100", 12, .9999999999999999, 1.7e308, targets)
        rows, targets = fixture()
        with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_ETS_TOTAL"):
            calculate(rows, "SKU-A100", 12, .95, 2, targets,
                      ets_forecaster=lambda values, season, horizon: [1e308] * horizon)

    def test_typed_output_and_generated_runtime_emit_only_failure_on_overflow(self):
        for number in (float("nan"), float("inf"), -float("inf")):
            with self.assertRaisesRegex(StatisticalPortError, "NONFINITE_OUTPUT"):
                typed_output_value(number, "DOUBLE")
        binding = {role: {"title": role, "columns": {name: {"identifier": name, "rowKey": name}
                    for name in names}} for role, names in {
                        "history": ("Month", "SKU", "Demand"),
                        "parameters": ("Key", "Value"), "targets": ("Period", "Month")}.items()}
        program = emit_preview_program(binding)
        rows, targets = fixture()
        for row in rows:
            row["Demand"] = 1e308
        frames = {"history": rows,
                  "parameters": [{"Key": k, "Value": v} for k, v in {
                      "Selected SKU": "SKU-A100", "Seasonality (months)": 12,
                      "Service level": .95, "Lead time (months)": 2}.items()],
                  "targets": [{"Period": i, "Month": target} for i, target in enumerate(targets, 1)]}

        class Frame:
            def __init__(self, data): self.data, self.columns = data, list(data[0])
            def collect(self): return self
            def __iter__(self): return iter(self.data)
        class Sigma:
            def __init__(self): self.outputs = []
            def get_element(self, title): return Frame(frames[title])
            def output(self, name, frame): self.outputs.append((name, frame))
        class Session:
            def create_dataframe(self, values, schema): return (values, schema)
        # A Snowpark Row supplies as_dict; use dict subclass to model that API.
        class Row(dict):
            def as_dict(self): return self
        Frame.collect = lambda self: [Row(row) for row in self.data]
        sigma = Sigma()
        exec(program["code"], {"sigma": sigma, "session": Session()})
        output_rows, _ = sigma.outputs[0][1]
        self.assertEqual(len(output_rows), 1)
        self.assertEqual(output_rows[0][0:3], ["FAILURE", False, "NONFINITE_LINEAR_Y_MEAN"])
        self.assertTrue(all(value is None for value in output_rows[0][5:]))

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

    def test_candidate_rejects_nonconverged_fit_before_reading_finite_forecast(self):
        class Fit:
            mle_retvals = {"converged": False, "warnflag": 2, "iterations": 0}
            forecast_called = False
            def forecast(self, horizon):
                self.forecast_called = True
                return [100.0] * horizon
        fit = Fit()
        with patch("statistical_ets_candidate._fit_model", return_value=fit):
            with self.assertRaises(CandidateFitError) as caught:
                forecast_with_diagnostics([100.0] * 36, 12, 12)
        self.assertEqual(caught.exception.code, "NONCONVERGED_FIT")
        self.assertEqual(caught.exception.diagnostics["warnflag"], 2)
        self.assertFalse(fit.forecast_called)

    def test_candidate_requires_explicit_optimizer_status_and_complete_finite_output(self):
        class Fit:
            def __init__(self, status, output):
                self.mle_retvals, self.output = status, output
            def forecast(self, horizon): return self.output
        for status in (None, {}, {"converged": True}, {"warnflag": 0}):
            with self.subTest(status=status), patch("statistical_ets_candidate._fit_model",
                                                     return_value=Fit(status, [1.0] * 12)):
                with self.assertRaises(CandidateFitError) as caught:
                    forecast_with_diagnostics([2.0] * 36, 12, 12)
                self.assertEqual(caught.exception.code, "MISSING_OPTIMIZER_STATUS")
        for output in ([1.0] * 11, [1.0] * 11 + [float("nan")]):
            with self.subTest(output=output), patch("statistical_ets_candidate._fit_model",
                                                     return_value=Fit({"converged": True, "warnflag": 0}, output)):
                with self.assertRaises(CandidateFitError) as caught:
                    forecast_with_diagnostics([2.0] * 36, 12, 12)
                self.assertEqual(caught.exception.code, "INVALID_FORECAST_OUTPUT")
        with patch("statistical_ets_candidate._fit_model",
                   return_value=Fit({"converged": True, "warnflag": 0, "iterations": 3}, [1.0] * 12)):
            result = forecast_with_diagnostics([2.0] * 36, 12, 12)
        self.assertEqual(result["forecasts"], [1.0] * 12)
        self.assertEqual(result["diagnostics"]["iterations"], 3)

    def test_candidate_retains_warnings_without_treating_them_as_success(self):
        class Fit:
            mle_retvals = {"converged": False, "warnflag": 2}
            def forecast(self, horizon): return [0.0] * horizon
        def fit_with_warning(values, seasonality):
            warnings.warn("optimizer did not converge", RuntimeWarning)
            return Fit()
        with patch("statistical_ets_candidate._fit_model", side_effect=fit_with_warning):
            with self.assertRaises(CandidateFitError) as caught:
                forecast_with_diagnostics([0.0] * 36, 12, 12)
        self.assertEqual(caught.exception.code, "NONCONVERGED_FIT")
        self.assertEqual(caught.exception.diagnostics["warningCount"], 1)
        self.assertEqual(caught.exception.diagnostics["warnings"][0]["message"], "optimizer did not converge")


if __name__ == "__main__":
    unittest.main()
