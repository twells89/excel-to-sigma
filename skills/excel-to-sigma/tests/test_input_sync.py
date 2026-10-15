#!/usr/bin/env python3
from __future__ import annotations

import copy
import datetime as dt
import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))

from input_sync import (  # noqa: E402
    InputSyncError,
    build_shared_input_architecture,
    emit_construction_program,
    emit_write_programs,
    plan_seed,
)
from input_orchestration import gate_input_completion, gate_input_staging  # noqa: E402


def load_script(filename, module_name):
    spec = importlib.util.spec_from_file_location(module_name, SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class TestSeedPlanner(unittest.TestCase):
    def test_rerun_inserts_nothing_and_partial_retry_reconciles_server_rows(self):
        source = [{"Business ID": "A", "Amount": 10}, {"Business ID": "B", "Amount": 0}]
        first = plan_seed(source, [{"Business ID": "A", "Amount": 10, "Row ID": "r-a"}],
                          ["Business ID"], "Row ID", snapshot={"id": "q1", "ageSeconds": 2})
        self.assertEqual([row["source"]["Business ID"] for row in first["inserts"]], ["B"])
        self.assertEqual(first["updates"], [])

        current = [{"Business ID": "A", "Amount": 10, "Row ID": "r-a"},
                   {"Business ID": "B", "Amount": 0, "Row ID": "r-b"}]
        rerun = plan_seed(source, current, ["Business ID"], "Row ID", prior_seed=first,
                          snapshot={"id": "q2", "ageSeconds": 0})
        self.assertEqual(rerun["inserts"], [])
        self.assertEqual(rerun["updates"], [])
        self.assertEqual(len(rerun["unchanged"]), 2)
        self.assertEqual(rerun["snapshot"]["id"], "q2")

    def test_source_change_updates_by_real_row_id_but_preserves_user_edit_as_conflict(self):
        baseline = plan_seed([{"Key": "A", "Value": 1}], [], ["Key"], "Row ID")
        safe = plan_seed([{"Key": "A", "Value": 2}],
                         [{"Key": "A", "Value": 1, "Row ID": "server-42"}],
                         ["Key"], "Row ID", prior_seed=baseline)
        self.assertEqual(safe["updates"][0]["rowId"], "server-42")
        self.assertEqual(safe["updates"][0]["changes"], {"Value": 2})

        edited = plan_seed([{"Key": "A", "Value": 2}],
                           [{"Key": "A", "Value": 9, "Row ID": "server-42"}],
                           ["Key"], "Row ID", prior_seed=baseline)
        self.assertEqual(edited["updates"], [])
        self.assertEqual(edited["conflicts"][0]["columns"], ["Value"])
        self.assertEqual(edited["conflicts"][0]["current"]["Value"], 9)

    def test_typed_keys_reject_duplicates_and_missing_but_keep_zero_distinct(self):
        with self.assertRaisesRegex(InputSyncError, "missing business key"):
            plan_seed([{"Key": None}], [], ["Key"], "Row ID")
        with self.assertRaisesRegex(InputSyncError, "duplicate source business key"):
            plan_seed([{"Key": 1}, {"Key": 1}], [], ["Key"], "Row ID")
        with self.assertRaisesRegex(InputSyncError, "duplicate current business key"):
            plan_seed([{"Key": "A"}], [{"Key": "A", "Row ID": "1"},
                                          {"Key": "A", "Row ID": "2"}], ["Key"], "Row ID")
        plan = plan_seed([{"Key": 0}, {"Key": "0"}], [], ["Key"], "Row ID")
        self.assertEqual([entry["key"][0]["type"] for entry in plan["inserts"]],
                         ["int", "str"])

    def test_orphans_are_reported_and_never_deleted(self):
        plan = plan_seed([], [{"Key": "old", "Row ID": "r-old"}], ["Key"], "Row ID")
        self.assertEqual(plan["orphans"][0]["rowId"], "r-old")
        self.assertNotIn("deletes", plan)


class TestMcpPrograms(unittest.TestCase):
    def setUp(self):
        self.spec = {
            "pageId": "page-1", "connectionId": "conn-1", "name": "Settings",
            "columns": [
                {"name": "Setting", "type": "text", "provenance": "source"},
                {"name": "As Of Date", "type": "datetime", "provenance": "source"},
            ],
        }

    def test_construction_creates_columns_then_refetches_and_seeds_once(self):
        program = emit_construction_program(
            self.spec,
            [{"Setting": "AsOfDate", "As Of Date": dt.date(2026, 6, 30)}],
            source_hash="a" * 64, plan_hash="b" * 64,
        )
        self.assertLess(program.index("add_editable_column('As Of Date'"),
                        program.index("upsert_rows(["))
        self.assertLess(program.index("add_system_columns(['row_id'])"),
                        program.index("upsert_rows(["))
        self.assertIn("sigma.workbook.get_page('page-1')", program)
        self.assertIn("get_element(created_id)", program)
        self.assertIn("get_column(created_column_ids[1])", program)
        self.assertIn("isinstance(table, InputTable)", program)
        self.assertIn("isinstance(c1, LiteralEditableColumn)", program)
        self.assertIn("datetime(2026, 6, 30, 12)", program)
        self.assertEqual(program.count("upsert_rows(["), 1)
        self.assertNotIn(".query(", program)

    def test_write_programs_use_actual_ids_chunk_at_2000_and_preserve_timestamp(self):
        rows = [{"Key": str(index), "At": dt.datetime(2026, 6, 30, 8, 45, 12)}
                for index in range(2001)]
        seed = plan_seed(rows, [], ["Key"], "Row ID")
        programs = emit_write_programs(
            seed,
            {"pageId": "p", "elementId": "e", "columnIds": {"Key": "c-key", "At": "c-at"}},
        )
        self.assertEqual(len(programs), 2)
        self.assertTrue(all("get_element('e')" in item["program"] for item in programs))
        self.assertIn("get_column('c-at')", programs[0]["program"])
        self.assertIn("datetime(2026, 6, 30, 8, 45, 12)", programs[0]["program"])
        self.assertEqual([item["chunkIndex"] for item in programs], [0, 1])
        self.assertTrue(all(item["checkpointId"] for item in programs))


class TestSharedArchitecture(unittest.TestCase):
    def test_exact_sales_expense_and_research_bindings_are_plan_hash_material(self):
        source_inputs = [
            {"id": "i_model_inputs_c2", "sourceCellId": "Model Inputs!C2",
             "sheet": "Model Inputs", "address": "C2", "defaultValue": 0.05,
             "values": {"2025": 0.05}, "role": "cross_sheet_driver"},
            {"id": "i_fy_results_b8", "sourceCellId": "FY results!B8",
             "sheet": "FY results", "address": "B8", "defaultValue": 100,
             "values": {"2024": 100}, "role": "source_literal"},
        ]
        architecture = build_shared_input_architecture(
            {"sourceInputs": source_inputs}, connection_id="conn", folder_id="folder")
        self.assertEqual(architecture["targetScope"],
                         {"connectionId": "conn", "folderId": "folder"})
        sales = architecture["inputs"]["salesSettings"]
        self.assertEqual(sales["rows"][0]["AsOfDate"], "2026-06-30")
        self.assertEqual(sales["sourceBinding"], "Assumptions!F2")
        self.assertEqual(architecture["calculations"]["TimingStatus"],
                         'If([Stage] = "Closed Won" or [Stage] = "Closed Lost", "Closed", If(DateTrunc("day", [Close Date]) < [AsOfDate], "Slipped", "Open"))')
        self.assertIn("referenceQuotingRepair", architecture["provenance"])

        policy = architecture["inputs"]["expensePolicy"]
        self.assertEqual([row["Category"] for row in policy["rows"]],
                         ["Airfare", "Lodging", "Meals", "Ground Transport", "Software", "Supplies"])
        self.assertEqual(next(row for row in policy["rows"] if row["Category"] == "Meals")["Per-item Limit"], 90)
        self.assertEqual(policy["sourceRange"], "Expense Policy!A2:D7")
        self.assertEqual(policy["unknownCategoryBehavior"], "explicit-null-error")

        research = architecture["inputs"]["researchModelInputs"]
        self.assertEqual(research["keyColumns"], ["Scope"])
        self.assertEqual(len(research["rows"]), 1)
        self.assertEqual(research["rows"][0]["i_model_inputs_c2"], 0.05)
        self.assertEqual(set(research["editableColumnBindings"]),
                         {"i_model_inputs_c2", "i_fy_results_b8"})
        self.assertEqual(research["sourceCellBindings"]["i_fy_results_b8"], "FY results!B8")
        self.assertEqual([column["type"] for column in research["columns"]],
                         ["text", "number", "number"])
        self.assertEqual(research["joinGrain"], ["Scope"])
        self.assertNotIn("Row ID", research["joinGrain"])


class TestInputOrchestration(unittest.TestCase):
    @mock.patch("input_orchestration.gate_staging_manifest", return_value=True)
    def test_staging_calls_manifest_gate_and_requires_exact_feature_scope(self, gate):
        manifest = {"sourceHash": "a" * 64, "planHash": "b" * 64}
        architecture = {"mappedFeatures": ["sales", "expense"]}
        approval = {"sourceHash": "a" * 64, "planHash": "b" * 64,
                    "reviewer": "Named Reviewer", "mappedFeatures": ["expense", "sales"]}
        result = gate_input_staging(manifest, approval, architecture)
        gate.assert_called_once_with(manifest, approval)
        self.assertEqual(result["status"], "APPROVED_STAGING")
        for field in ("sourceHash", "planHash", "reviewer", "mappedFeatures"):
            bad = copy.deepcopy(approval); bad.pop(field)
            with self.assertRaisesRegex(ValueError, field):
                gate_input_staging(manifest, bad, architecture)

    @mock.patch("input_orchestration.gate_completion_manifest", return_value=True)
    def test_completion_calls_manifest_gate_and_consumes_recorded_runtime_evidence(self, gate):
        evidence = {"sourceHash": "a" * 64, "planHash": "b" * 64,
                    "querySnapshotId": "query-1", "readbackId": "read-1",
                    "checkpointIds": ["chunk-0"], "status": "passed",
                    "uniqueKeyCounts": {"researchModelInputs": {"rows": 1, "distinct": 1}},
                    "persistenceReadbackId": "reopen-1",
                    "acceptanceChecks": [{"name": "research edit/restore", "status": "passed",
                                           "restored": True}]}
        result = gate_input_completion(
            {"sourceHash": "a" * 64, "planHash": "b" * 64}, evidence)
        gate.assert_called_once()
        self.assertEqual(result["status"], "VERIFIED_COMPLETE")


class TestTypedWorkbookExtraction(unittest.TestCase):
    def test_business_id_is_retained_and_system_id_requires_explicit_provenance(self):
        module = load_script("xlsx-to-input-csv.py", "typed_input_csv")
        headers = ["ID", "Amount", "Updated At"]
        self.assertEqual(module.select_columns(headers, system_columns={"Updated At"}),
                         [(0, "ID"), (1, "AMOUNT")])

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "fixture.xlsx"
            wb = Workbook(); ws = wb.active
            ws.append(["ID", "Amount"]); ws.append([0, None]); ws.append([1, 0])
            tab = Table(displayName="Entries", ref="A1:B3")
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
            ws.add_table(tab); wb.save(path)
            loaded = load_workbook(path, data_only=True)
            _, found = module.find_table(loaded, "Entries")
            rows = module.extract_typed_rows(loaded.active, found)
            self.assertEqual(rows, [{"ID": 0, "AMOUNT": None}, {"ID": 1, "AMOUNT": 0}])


class TestInputSyncCli(unittest.TestCase):
    def test_plan_command_writes_reviewable_plan_and_program_bundle(self):
        module = load_script("input-sync.py", "input_sync_cli")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "source.json").write_text(json.dumps([{"Key": "A", "Value": 1}]))
            (root / "current.json").write_text("[]")
            (root / "target.json").write_text(json.dumps({
                "pageId": "p", "elementId": "e",
                "columnIds": {"Key": "k", "Value": "v"},
            }))
            code = module.main([
                "plan", "--source-rows", str(root / "source.json"),
                "--current-rows", str(root / "current.json"),
                "--keys", "Key", "--row-id", "Row ID",
                "--snapshot-id", "query-1", "--snapshot-age-seconds", "4",
                "--target", str(root / "target.json"),
                "--out", str(root / "bundle.json"),
            ])
            self.assertEqual(code, 0)
            bundle = json.loads((root / "bundle.json").read_text())
            self.assertEqual(bundle["plan"]["snapshot"], {"id": "query-1", "ageSeconds": 4})
            self.assertEqual(bundle["programs"][0]["actionCount"], 1)
            self.assertIn("checkpointId=", bundle["programs"][0]["program"])


if __name__ == "__main__":
    unittest.main()
