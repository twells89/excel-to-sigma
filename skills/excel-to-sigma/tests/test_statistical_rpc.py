"""Focused contract tests for the isolated statistical ETS provider."""
from datetime import date
import ast
import copy
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.statistical_ets_rpc import (RpcContractError, decode_fit_response,
                                     encode_fit_request, server_fit)
from lib.statistical_rpc_runtime import (PROCEDURE, emit_procedure_sql,
                                         emit_rpc_preview_program)
import statistical_ets_candidate as candidate


BINDING = "a" * 64
INPUT = "b" * 64
RUN = "12345678-1234-1234-1234-123456789abc"
VERSIONS = {"statsmodels": "0.15.0", "scipy": "1.18.1"}


def request(sku="SKU-A100"):
    return encode_fit_request([100.0 + i for i in range(36)], 12, 12,
        source_binding_hash=BINDING, input_hash=INPUT, preview_run_id=RUN, sku=sku)


class FakeFit:
    mle_retvals = {"converged": True, "warnflag": 0, "iterations": 4, "fcalls": 11}

    def forecast(self, horizon):
        return [float(201 + i) for i in range(horizon)]


class RpcTests(unittest.TestCase):
    def test_emitted_sigma_route_preserves_review_only_rows_and_typed_failure(self):
        meta_path = SCRIPTS.parents[3] / "statistical-observed-preview6-R1-CANDIDATE/program-metadata.json"
        if not meta_path.exists():
            self.skipTest("observed source metadata unavailable")
        identity = json.loads(meta_path.read_text())["sourceIdentity"]
        binding = {r: {"title": v["title"], "columns": v["columns"]}
                   for r, v in identity["roles"].items()}
        code = emit_rpc_preview_program(binding, identity)["code"]

        class Row:
            def __init__(self, values):
                self.values = values

            def as_dict(self):
                return self.values

        class Frame:
            def __init__(self, spec, rows, session):
                self.columns = [pair["identifier"] for pair in spec["columns"].values()]
                self.rows = [{spec["columns"][name]["rowKey"]: value
                              for name, value in row.items()} for row in rows]
                self.session = session

            def collect(self):
                return [Row(values) for values in self.rows]

        class Session:
            malformed = False

            def call(self, name, raw):
                self.assert_name = name
                if self.malformed:
                    return "{}"
                return server_fit(raw, fit_fn=lambda *_: {
                    "forecasts": [201.0 + i for i in range(12)],
                    "diagnostics": {"converged": True, "warnflag": 0}},
                    version_fn=lambda package: VERSIONS[package])

            def create_dataframe(self, rows, schema):
                return {"rows": rows, "schema": schema}

        class Sigma:
            def __init__(self, frames):
                self.frames = frames
                self.outputs = {}

            def get_element(self, title):
                return self.frames[title]

            def output(self, name, frame):
                self.outputs[name] = frame

        session = Session()
        history = []
        for i in range(36):
            month = date(2023 + i // 12, i % 12 + 1, 1)
            for j, sku in enumerate(("SKU-A100", "SKU-B200", "SKU-C300")):
                history.append({"Month": month, "SKU": sku, "Demand": 100.0 + i + 40*j})
        data = {"history": history,
                "parameters": [{"Scope": "Parameters", "SelectedSKU": "SKU-A100",
                                "Seasonality": 12, "ServiceLevel": 0.95, "LeadTime": 2}],
                "targets": [{"Period": i+1, "Month": date(2026, i+1, 1)}
                            for i in range(12)]}
        frames = {binding[r]["title"]: Frame(binding[r], rows, session)
                  for r, rows in data.items()}
        sigma = Sigma(frames)
        exec(compile(code, "<rpc-sigma-test>", "exec"), {"sigma": sigma})
        output = sigma.outputs["StatisticalForecastPreview"]
        self.assertEqual(len(output["rows"]), 36)
        self.assertTrue(all(row[1] is False for row in output["rows"]))
        self.assertTrue(all(row[2] == "CANDIDATE_REVIEW_REQUIRED" for row in output["rows"]))
        self.assertEqual(session.assert_name, PROCEDURE)
        self.assertIn('"ETS" DOUBLE', output["schema"])
        session.malformed = True
        sigma = Sigma(frames)
        exec(compile(code, "<rpc-sigma-test>", "exec"), {"sigma": sigma})
        failure = sigma.outputs["StatisticalForecastPreview"]["rows"]
        self.assertEqual(len(failure), 1)
        self.assertEqual(failure[0][0:3], ["FAILURE", False, "RPC_RESPONSE_INVALID"])

    def test_server_preserves_unchanged_local_provider_values_and_diagnostics(self):
        raw = request()
        with patch.object(candidate, "_fit_model", return_value=FakeFit()):
            local = candidate.forecast_with_diagnostics([100.0 + i for i in range(36)], 12, 12)
            response = decode_fit_response(server_fit(raw,
                fit_fn=candidate.forecast_with_diagnostics,
                version_fn=lambda name: VERSIONS[name]), raw)
        self.assertEqual(response["forecasts"], local["forecasts"])
        self.assertEqual(response["diagnostics"], local["diagnostics"])
        self.assertEqual(response["runtimePackages"], VERSIONS)
        self.assertEqual(response["identity"], json.loads(raw)["identity"])

    def test_invalid_requests_and_versions_fail_closed(self):
        for bad, code in (("{", "RPC_REQUEST_INVALID"),
                          ("x" * 8193, "RPC_REQUEST_INVALID"),
                          ('{"identity":{},"identity":{}}', "RPC_REQUEST_INVALID"),
                          (request().replace('"SKU-A100"', '"OTHER"'), "RPC_IDENTITY_INVALID")):
            with self.subTest(code=code):
                response = json.loads(server_fit(bad))
                self.assertEqual(response["status"], "FAILURE")
                self.assertEqual(response["errorCode"], code)
                self.assertIsNone(response["forecasts"])
        raw = request()
        response = decode_fit_response(server_fit(raw,
            fit_fn=lambda *_: self.fail("fit must not run on wrong version"),
            version_fn=lambda name: "0" if name == "statsmodels" else "1.18.1"), raw)
        self.assertEqual(response["errorCode"], "CANDIDATE_VERSION_MISMATCH")

    def test_response_identity_size_and_numeric_guards(self):
        raw = request()
        with self.assertRaisesRegex(RpcContractError, "RPC_RESPONSE_INVALID"):
            decode_fit_response("x" * 16385, raw)
        valid = json.loads(server_fit(raw, fit_fn=lambda *_: {
            "forecasts": [1.0] * 12, "diagnostics": {}},
            version_fn=lambda name: VERSIONS[name]))
        tampered = copy.deepcopy(valid)
        tampered["identity"]["inputHash"] = "c" * 64
        with self.assertRaisesRegex(RpcContractError, "RPC_RESPONSE_IDENTITY_MISMATCH"):
            decode_fit_response(json.dumps(tampered), raw)
        tampered = copy.deepcopy(valid)
        tampered["forecasts"][0] = float("nan")
        with self.assertRaisesRegex(RpcContractError, "RPC_RESPONSE_INVALID"):
            decode_fit_response(json.dumps(tampered), raw)
        with self.assertRaisesRegex(RpcContractError, "RPC_VALUES_INVALID"):
            encode_fit_request([float("nan")] * 36, 12, 12,
                source_binding_hash=BINDING, input_hash=INPUT,
                preview_run_id=RUN, sku="SKU-A100")

    def test_generated_sql_is_pure_and_caller_uses_literal_sources(self):
        sql = emit_procedure_sql()
        self.assertTrue(sql.startswith("CREATE PROCEDURE " + PROCEDURE))
        self.assertNotIn("CREATE OR REPLACE", sql)
        self.assertIn("EXECUTE AS CALLER", sql)
        self.assertIn("'statsmodels==0.15.0'", sql)
        self.assertIn("'scipy==1.18.1'", sql)
        self.assertIn("def forecast_with_diagnostics(", sql)
        self.assertIn("def calculate(", sql)
        handler = sql.split("$$", 2)[1]
        tree = ast.parse(handler)
        server = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "server_fit")
        forbidden = {"sql", "table", "write", "collect", "requests", "urlopen", "open"}
        attrs = {n.attr for n in ast.walk(server) if isinstance(n, ast.Attribute)}
        self.assertFalse(attrs & forbidden)
        self.assertIn("return server_fit(request_json)", handler)

        meta_path = SCRIPTS.parents[3] / "statistical-observed-preview6-R1-CANDIDATE/program-metadata.json"
        if not meta_path.exists():
            self.skipTest("observed source metadata unavailable")
        identity = json.loads(meta_path.read_text())["sourceIdentity"]
        binding = {r: {"title": v["title"], "columns": v["columns"]}
                   for r, v in identity["roles"].items()}
        program = emit_rpc_preview_program(binding, identity)
        code = program["code"]
        self.assertEqual(code.count("session.call("), 1)
        self.assertNotIn("from statsmodels", code)
        self.assertIn("ReadyToPublish=False", code)
        self.assertIn("INPUT_CHANGED_DURING_RUN", code)
        self.assertIn("sigma.output('StatisticalForecastPreview'", code)
        calls = [n for n in ast.walk(ast.parse(code)) if isinstance(n, ast.Call)
                 and isinstance(n.func, ast.Attribute) and n.func.attr == "get_element"]
        self.assertEqual([n.args[0].value for n in calls],
                         [binding[r]["title"] for r in ("history", "parameters", "targets")])


if __name__ == "__main__":
    unittest.main()
