"""Generated-data checks; external Excel replay lives outside normal CI."""
from datetime import date
import ast
from pathlib import Path
import sys
import json
import copy
import runpy
import shutil
import tempfile
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


def assert_literal_sigma_references(test, code, expected_titles):
    """Sigma discovers Python dependencies/output names from string literals."""
    calls = [node for node in ast.walk(ast.parse(code)) if isinstance(node, ast.Call)
             and isinstance(node.func, ast.Attribute) and isinstance(node.func.value, ast.Name)
             and node.func.value.id == "sigma"]
    for node in calls:
        if node.func.attr.startswith("get_") or node.func.attr == "output":
            test.assertTrue(node.args and isinstance(node.args[0], ast.Constant)
                            and isinstance(node.args[0].value, str), ast.unparse(node))
    dependency_calls = [node for node in calls if node.func.attr == "get_element"]
    test.assertEqual([node.args[0].value for node in dependency_calls], expected_titles)


class StatisticalForecastTests(unittest.TestCase):
    def test_sigma_literal_reference_ast_rejects_dynamic_names(self):
        for code in ("sigma.get_element(title)", "sigma.output(output_name, frame)",
                     "sigma.get_control('prefix' + variable)"):
            with self.subTest(code=code), self.assertRaises(AssertionError):
                assert_literal_sigma_references(self, code, [])

    def test_source_packet_rejects_finite_row_or_csv_tampering(self):
        workspace = LIB.parents[4]
        original = workspace / "broader-eight-source-replay-20261020/source-originals/05 Statistical Demand Forecast.xlsx"
        packet_dir = workspace / "statistical-sigma-preview-20261020-CANDIDATE"
        if not original.exists() or not (packet_dir / "source-inputs.json").exists():
            self.skipTest("external source packet unavailable")
        operator = runpy.run_path(str(LIB.parent / "statistical-operator.py"))
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder)
            for name in ("source-inputs.json", "history.csv", "parameters.csv", "targets.csv"):
                shutil.copyfile(packet_dir / name, target / name)
            operator["source_packet"](target / "source-inputs.json", original)
            payload = json.loads((target / "source-inputs.json").read_text())
            payload["roles"]["history"]["rows"][0]["Demand"] += 12345
            (target / "source-inputs.json").write_text(json.dumps(payload))
            with self.assertRaisesRegex(ValueError, "SOURCE_PACKET_CONTENT_MISMATCH"):
                operator["source_packet"](target / "source-inputs.json", original)
            shutil.copyfile(packet_dir / "source-inputs.json", target / "source-inputs.json")
            with (target / "history.csv").open("a") as stream:
                stream.write("tamper")
            with self.assertRaisesRegex(ValueError, "SOURCE_CSV_CONTENT_MISMATCH"):
                operator["source_packet"](target / "source-inputs.json", original)

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

    def test_phase1_source_snapshot_rejects_replay_and_changed_partial_schema(self):
        workspace = LIB.parents[4]
        packet = workspace / "statistical-sigma-preview-20261020-CANDIDATE"
        before = packet / "root-phase1-partial-source2.json"
        after = packet / "root-phase1-after-source3.json"
        if not before.exists() or not after.exists():
            self.skipTest("actual source snapshots unavailable")
        operator = runpy.run_path(str(LIB.parent / "statistical-operator.py"))
        source2, source3 = json.loads(before.read_text()), json.loads(after.read_text())
        args = {"page_id": "dsRS3ZlYIn",
                "connection_id": "cb2f5180-641f-47bd-8efa-da9d590d855a"}
        operator["phase1_snapshot"](source2, role="parameters", existing_id="H5nEnfBk1F", **args)
        operator["phase1_snapshot"](source2, role="targets", **args)
        with self.assertRaisesRegex(ValueError, "ROLE_ALREADY_EXISTS"):
            operator["phase1_snapshot"](source3, role="targets", **args)
        with self.assertRaisesRegex(ValueError, "PARTIAL_PARAMETER_STATE_CHANGED"):
            operator["phase1_snapshot"](source3, role="parameters", existing_id="H5nEnfBk1F", **args)

    def test_phase2_rejects_source_identity_drift_before_emission(self):
        workspace = LIB.parents[4]
        source_path = workspace / "statistical-sigma-preview-20261020-CANDIDATE/root-phase1-after-source3.json"
        if not source_path.exists():
            self.skipTest("actual source snapshot unavailable")
        source = json.loads(source_path.read_text())
        operator = runpy.run_path(str(LIB.parent / "statistical-operator.py"))
        roles = {}
        for role, names in CANDIDATE_REQUIRED.items():
            title = operator["TABLES"][role][0]
            element = next(e for e in source["contents"]["elements"] if e.get("name") == title)
            ids = {c["name"]: c["id"] for c in element["columns"] if c.get("name")}
            roles[role] = {"elementId": element["id"], "title": title, "columnIds": ids,
                           "rowCount": {"history": 108, "parameters": 1, "targets": 12}[role],
                           "naturalKeysVerified": True, "readbackSnapshotId": "observed-" + role,
                           "columns": {name: {"identifier": ids[name], "rowKey": ids[name]} for name in names}}
        readback = {"sourceHash": source["contents"].get("sourceHash",
                    "02f90f32fa077ee1fac7e2d7a8cdc48b61bf09f6fc69b27a855886ed9cd9d310"),
                    "workbookId": source["workbookId"],
                    "workbookVersion": source["documentVersion"],
                    "pageId": "dsRS3ZlYIn", "sourceSnapshotId": "observed-source",
                    "roles": roles}
        program = operator["phase2_program"](readback, actual_source=source,
                    source_packet_hash="0" * 64, candidate_ets=True)
        assert_literal_sigma_references(self, program["code"], [roles[role]["title"] for role in CANDIDATE_REQUIRED])
        self.assertEqual(program["sourceIdentity"]["roles"]["history"]["elementId"],
                         roles["history"]["elementId"])
        changed = {**readback, "workbookVersion": readback["workbookVersion"] + 1}
        with self.assertRaisesRegex(ValueError, "ACTUAL_SOURCE_VERSION_OR_PAGE_MISMATCH"):
            operator["phase2_program"](changed, actual_source=source,
                        source_packet_hash="0" * 64, candidate_ets=True)
        changed_roles = {**roles, "history": {**roles["history"], "elementId": roles["targets"]["elementId"]}}
        with self.assertRaisesRegex(ValueError, "ACTUAL_SOURCE_ELEMENT_MISMATCH"):
            operator["phase2_program"]({**readback, "roles": changed_roles}, actual_source=source,
                        source_packet_hash="0" * 64, candidate_ets=True)

    def test_resumable_seed_requeries_and_preserves_edits(self):
        workspace = LIB.parents[4]
        packet = workspace / "statistical-sigma-preview-20261020-CANDIDATE"
        original = workspace / "broader-eight-source-replay-20261020/source-originals/05 Statistical Demand Forecast.xlsx"
        source = packet / "root-seed-source-before5.json"
        if not source.exists() or not original.exists():
            self.skipTest("actual source snapshots unavailable")
        seed_module = runpy.run_path(str(LIB.parent / "statistical-seed.py"))
        seed = seed_module["seed_plan"]
        payload = json.loads((packet / "source-inputs.json").read_text())
        first = payload["roles"]["history"]["rows"][0]
        source_json = json.loads(source.read_text())
        element = next(e for e in source_json["contents"]["elements"] if e.get("id") == "qPoaWFS7V6")
        column_ids = {c["name"]: c["id"] for c in element["columns"] if c.get("name")}
        requests = seed_module["query_requests"](source_json["workbookId"], element["id"],
                    column_ids, ["Month", "SKU"], "EHGAUETP2V")
        describes = packet / "root-seed-describes4.json"
        def receipt(rows):
            names = ["Month", "SKU", "Demand"]
            def result(columns, matrix):
                return {"structuredContent": {"columns": columns, "rows": matrix}}
            return [{"status": "fulfilled", "value": {"role": "history", "elementId": element["id"],
                    "selectedColumnIds": [*column_ids.values(), "EHGAUETP2V"],
                    "rowIdAlias": "EHGAUETP2V", "request": {"query": requests["rows"]},
                    "startedAt": "2026-10-20T22:04:19Z", "completedAt": "2026-10-20T22:04:20Z",
                    "result": result([*column_ids.values(), "EHGAUETP2V"] if rows else [],
                                    [[*(r[n] for n in names), r["EHGAUETP2V"]] for r in rows]),
                    "count": {"request": {"query": requests["count"]},
                              "startedAt": "2026-10-20T22:04:20Z", "completedAt": "2026-10-20T22:04:21Z",
                              "result": result(["__COUNT__"], [[len(rows)]])}}}]
        with tempfile.TemporaryDirectory() as folder:
            current = Path(folder) / "query.json"
            current.write_text(json.dumps(receipt([])))
            opts = {"source_packet_path": packet / "source-inputs.json",
                    "source_original": original, "actual_source": source,
                    "actual_source_after": packet / "root-seed-source-after5.json",
                    "query_receipt": current, "role": "history",
                    "describe_receipt": describes,
                    "row_id_key": "EHGAUETP2V"}
            current.write_bytes((packet / "root-seed-executed-queries5.json").read_bytes())
            actual_plan, actual_chunk = seed(**opts)
            self.assertEqual((len(actual_plan["inserts"]), actual_chunk["actionCount"]), (108, 25))
            calls = [node.func for node in ast.walk(ast.parse(actual_chunk["program"]))
                     if isinstance(node, ast.Call)]
            # The actual SDK rejected repr; preserve the typed upsert receipt.
            self.assertNotIn("repr", [call.id for call in calls if isinstance(call, ast.Name)])
            self.assertIn("print(row_ids)\n", actual_chunk["program"])
            self.assertEqual({call.id for call in calls if isinstance(call, ast.Name)},
                             {"datetime", "isinstance", "RuntimeError", "print"})
            self.assertEqual({call.attr for call in calls if isinstance(call, ast.Attribute)},
                             {"get_page", "get_element", "get_column", "upsert_rows"})
            changed_after = Path(folder) / "changed-source.json"
            changed_after.write_text(source.read_text().replace('"documentVersion": 3', '"documentVersion": 4'))
            with self.assertRaisesRegex(ValueError, "SOURCE_CHANGED_DURING_QUERY"):
                seed(**{**opts, "actual_source_after": changed_after})
            current.write_text(json.dumps(receipt([])))
            plan, chunk = seed(**opts)
            self.assertEqual((len(plan["inserts"]), chunk["actionCount"], chunk["remainingCount"]),
                             (108, 25, 83))
            with self.assertRaisesRegex(ValueError, "REQUERY_REQUIRED_AFTER_PRIOR_BATCH"):
                seed(**opts, prior_seed=plan)
            observed = {"Month": first["Month"], "SKU": first["SKU"],
                        "Demand": first["Demand"], "EHGAUETP2V": "observed-row-id"}
            current.write_text(json.dumps(receipt([observed])))
            next_plan, next_chunk = seed(**opts, prior_seed=plan)
            self.assertEqual((len(next_plan["inserts"]), len(next_plan["unchanged"])), (107, 1))
            self.assertEqual(next_chunk["actionCount"], 25)
            edited = {**observed, "Demand": observed["Demand"] + 1}
            current.write_text(json.dumps(receipt([edited])))
            conflict, program = seed(**opts, prior_seed=plan)
            self.assertEqual(len(conflict["conflicts"]), 1)
            self.assertIsNone(program)
            forged = {**plan, "seedRecords": [{"confirmedMigrationValue": edited["Demand"]}]}
            forged_conflict, forged_program = seed(**opts, prior_seed=forged)
            self.assertEqual(len(forged_conflict["conflicts"]), 1)
            self.assertIsNone(forged_program)
            current.write_text(json.dumps(receipt([observed, {**observed, "EHGAUETP2V": "other-id"}])))
            with self.assertRaisesRegex(ValueError, "duplicate current business key"):
                seed(**opts, prior_seed=plan)
            current.write_text(json.dumps(receipt([{**observed, "SKU": "edited-key"}])))
            orphan, program = seed(**opts, prior_seed=plan)
            self.assertEqual(len(orphan["orphans"]), 1)
            self.assertIsNone(program)
            wrong_target = copy.deepcopy(receipt([]))
            wrong_target[0]["value"]["request"]["query"]["workbookId"] = "wrong"
            for unbound in ([], {"rows": []}, wrong_target):
                current.write_text(json.dumps(unbound))
                with self.assertRaisesRegex(ValueError, "BOUND_QUERY_RECEIPT_REQUIRED|EXECUTED_QUERY_TARGET_OR_STATUS_MISMATCH"):
                    seed(**opts)
            mismatched = receipt([])
            mismatched[0]["value"]["count"]["result"]["structuredContent"]["rows"] = [[1]]
            current.write_text(json.dumps(mismatched))
            with self.assertRaisesRegex(ValueError, "QUERY_COUNT_OR_SCHEMA_MISMATCH"):
                seed(**opts)
            missing_execution_time = receipt([])
            del missing_execution_time[0]["value"]["startedAt"]
            current.write_text(json.dumps(missing_execution_time))
            with self.assertRaisesRegex(ValueError, "EXECUTED_QUERY_TIMING_REQUIRED"):
                seed(**opts)

    def test_seed_date_normalization_rejects_ambiguous_values(self):
        normalize = runpy.run_path(str(LIB.parent / "statistical-seed.py"))["_normalize"]
        self.assertEqual(normalize("2025-01-01T00:00:00Z", "datetime"), "2025-01-01")
        for value in ("2025-01-01garbage", "2025-01-01T23:59:59-10:00",
                      "2025-01-01T00:00:00+01:00", "2025-01-01T13:00:00"):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "AMBIGUOUS_DATE_VALUE"):
                normalize(value, "datetime")

    def test_candidate_sigma_preview_is_typed_current_and_review_only(self):
        binding = {role: {"title": role, "columns": {name: {"identifier": name, "rowKey": name}
                    for name in names}} for role, names in CANDIDATE_REQUIRED.items()}
        identity = {"workbookId": "observed-workbook", "workbookVersion": 3,
                    "pageId": "observed-page", "sourceSnapshotId": "observed-snapshot",
                    "sourceSnapshotHash": "1" * 64,
                    "sourcePacketHash": "0" * 64,
                    "roles": {role: {"elementId": "observed-" + role,
                                     "title": role, "columns": binding[role]["columns"]}
                              for role in CANDIDATE_REQUIRED}}
        program = emit_preview_program(binding, candidate_ets=True, source_identity=identity)
        changed_identity = {**identity, "workbookVersion": 4}
        changed_program = emit_preview_program(binding, candidate_ets=True,
                                               source_identity=changed_identity)
        self.assertNotEqual(program["codeHash"], changed_program["codeHash"])
        self.assertNotEqual(program["sourceBindingHash"], changed_program["sourceBindingHash"])
        changed_role_identity = {**identity, "roles": {**identity["roles"],
                                 "history": {**identity["roles"]["history"],
                                             "elementId": "same-title-different-element"}}}
        changed_role_program = emit_preview_program(binding, candidate_ets=True,
                                                     source_identity=changed_role_identity)
        self.assertNotEqual(program["codeHash"], changed_role_program["codeHash"])
        self.assertNotEqual(program["sourceBindingHash"], changed_role_program["sourceBindingHash"])
        with self.assertRaisesRegex(ValueError, "OBSERVED_SOURCE_IDENTITY_REQUIRED"):
            emit_preview_program(binding, candidate_ets=True)
        compile(program["code"], "<candidate preview>", "exec")
        assert_literal_sigma_references(self, program["code"], [binding[role]["title"] for role in CANDIDATE_REQUIRED])
        rows, targets = fixture()
        frames = {"history": rows,
                  "parameters": [{"Scope": "Parameters", "SelectedSKU": "SKU-A100",
                                  "Seasonality": 12, "ServiceLevel": .95, "LeadTime": 2}],
                  "targets": [{"Period": i, "Month": target} for i, target in enumerate(targets, 1)]}

        class Row(dict):
            def as_dict(self): return self
        class Frame:
            def __init__(self, data): self.data, self.columns, self.session = data, list(data[0]), Session()
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
        versions = {"statsmodels": "0.15.0", "scipy": "1.18.1"}
        def execute():
            sigma = Sigma()
            with patch.dict(sys.modules, {"statsmodels": types.ModuleType("statsmodels"),
                                          "statsmodels.tsa": types.ModuleType("statsmodels.tsa"),
                                          "statsmodels.tsa.exponential_smoothing": types.ModuleType("statsmodels.tsa.exponential_smoothing"),
                                          "statsmodels.tsa.exponential_smoothing.ets": fake}), \
                 patch("importlib.metadata.version", side_effect=lambda name: versions[name]):
                exec(program["code"], {"sigma": sigma})
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
        self.assertEqual(first["SourceBindingHash"], program["sourceBindingHash"])
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
        versions["scipy"] = "1.17.0"
        package_failure, _ = execute()
        self.assertEqual(dict(zip(columns, package_failure[0]))["ErrorCode"],
                         "CANDIDATE_SCIPY_VERSION_MISMATCH")
        versions["scipy"] = "1.18.1"
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
            def __init__(self, data): self.data, self.columns, self.session = data, list(data[0]), Session()
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
        exec(program["code"], {"sigma": sigma})
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
        assert_literal_sigma_references(self, program["code"], [binding[role]["title"] for role in binding])
        rows, targets = fixture()

        class Row:
            def __init__(self, data): self.data = data
            def as_dict(self): return self.data

        class Frame:
            def __init__(self, data):
                self.data = data
                self.columns = list(data[0])
                self.session = Session()
            def collect(self): return [Row(row) for row in self.data]

        class Sigma:
            def __init__(self, frames): self.frames, self.outputs, self.reads = frames, [], []
            def get_element(self, name):
                self.reads.append(name)
                return Frame(self.frames[name])
            def output(self, name, frame): self.outputs.append((name, frame))

        class Session:
            def create_dataframe(self, values, schema): return (values, schema)

        frames = {"history": rows,
                  "parameters": [{"Key": k, "Value": v} for k, v in {
                      "Selected SKU": "SKU-A100", "Seasonality (months)": 12,
                      "Service level": .95, "Lead time (months)": 2}.items()],
                  "targets": [{"Period": i, "Month": target} for i, target in enumerate(targets, 1)]}
        sigma = Sigma(frames)
        exec(program["code"], {"sigma": sigma})
        values, schema = sigma.outputs[0][1]
        self.assertEqual(len(values), 36)
        self.assertEqual(values[0][0:3], ["ETS_UNAVAILABLE", False, "EXCEL_ETS_PROVIDER_UNQUALIFIED"])
        self.assertIsNone(values[0][7])  # ETS remains null.
        self.assertIn('"ReadyToPublish" BOOLEAN', schema)
        self.assertEqual(sigma.reads, ["history", "parameters", "targets"] * 2)
        class ChangingSigma(Sigma):
            def get_element(self, name):
                if len(self.reads) == 3:
                    self.frames["parameters"] = [dict(row) for row in self.frames["parameters"]]
                    self.frames["parameters"][0]["Value"] = "SKU-B200"
                return super().get_element(name)
        changing = ChangingSigma(copy.deepcopy(frames))
        exec(program["code"], {"sigma": changing})
        self.assertEqual(changing.outputs[0][1][0][0][0:3],
                         ["FAILURE", False, "INPUT_CHANGED_DURING_RUN"])
        self.assertEqual(changing.reads, ["history", "parameters", "targets"] * 2)
        frames["parameters"][0]["Value"] = "SKU-Z"
        sigma = Sigma(frames)
        exec(program["code"], {"sigma": sigma})
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
