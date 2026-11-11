#!/usr/bin/env python3
"""Approve / Send back: approver-only update-rows decisions on the Approvals table."""
from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEMO_KIT = ROOT.parents[1] / "docs" / "demo-kit"
sys.path.insert(0, str(ROOT / "scripts" / "lib"))
sys.path.insert(0, str(ROOT / "scripts"))

import demo_workflows as dw  # noqa: E402

APPROVER = "tj+1@sigmacomputing.com"


@unittest.skipUnless(DEMO_KIT.is_dir(), "demo kit not present (standalone skill copy)")
class ApproveTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = dw.build_workflow_manifest(DEMO_KIT / "demo-forecast-macros.bas")

    def elements(self, approvers=APPROVER):
        return dw.workflow_elements(dw.with_approvers(self.base, approvers), "conn-1", "FY27 forecast")

    def decision(self, wf, button):
        return next(b for b in wf["decisions"] if b["id"] == button)["actions"][0]["effects"]

    def test_approve_updates_only_submitted_rows_for_approvers(self):
        (effect,) = self.decision(self.elements(), "btnApprove")
        self.assertEqual((effect["effect"], effect["tableElementId"]), ("update-rows", "approvals"))
        formula = effect["whichRows"]["formula"]
        self.assertEqual(effect["whichRows"]["type"], "formula")
        self.assertIn('[Status] = "Submitted"', formula)
        self.assertIn(APPROVER, formula)
        self.assertIn("CurrentUserEmail()", formula)
        self.assertIn("Lower(CurrentUserEmail())", formula)
        cid = dw.COLUMN_IDS
        self.assertEqual(effect["values"][cid[("approvals", "Status")]]["value"]["value"], "Approved")
        self.assertEqual(effect["values"][cid[("approvals", "Decided By")]]["formula"], "CurrentUserEmail()")
        self.assertEqual(effect["values"][cid[("approvals", "Decided At")]]["formula"], "Now()")
        self.assertNotIn(cid[("approvals", "Log Order")], effect["values"])   # history order untouched

    def test_approver_match_is_exact_and_case_insensitive(self):
        formula = dw.approver_filter(["TJ+1@SigmaComputing.com", "b@x.com"])
        self.assertIn(",tj+1@sigmacomputing.com,b@x.com,", formula)      # lower-cased list
        self.assertIn('Concat(",", Lower(CurrentUserEmail()), ",")', formula)  # whole-address match only

    def test_send_back_sets_returned(self):
        (effect,) = self.decision(self.elements(), "btnReturn")
        self.assertEqual(effect["values"][dw.COLUMN_IDS[("approvals", "Status")]]["value"]["value"], "Returned")
        self.assertIn('[Status] = "Submitted"', effect["whichRows"]["formula"])

    def test_buttons_are_decisions_not_workflows(self):
        wf = self.elements()
        self.assertEqual([b["id"] for b in wf["decisions"]], ["btnApprove", "btnReturn"])
        self.assertEqual([b["text"] for b in wf["decisions"]], ["Approve", "Send back"])
        self.assertNotIn("btnApprove", [b["id"] for b in wf["buttons"]])
        manifest = dw.with_approvers(self.base, APPROVER)
        self.assertEqual(manifest["workflows"], self.base["workflows"])   # the agent never gets an approve tool

    def test_no_approvers_no_approve_buttons(self):
        for none in (None, "", " , "):
            manifest = dw.with_approvers(self.base, none)
            wf = dw.workflow_elements(manifest, "conn-1", "FY27 forecast")
            self.assertEqual(wf["decisions"], [])
            note = next(d for d in manifest["dispositions"] if d["disposition"] == "approve-omitted-no-approvers")
            self.assertIn("--approvers", note["note"])
        self.assertEqual([b["id"] for b in wf["buttons"]], ["btnRoll", "btnSnapshot", "btnSubmit"])

    def test_with_approvers_is_idempotent_and_rejects_bad_emails(self):
        once = dw.with_approvers(self.base, APPROVER)
        self.assertEqual(dw.with_approvers(once, APPROVER), once)
        self.assertEqual(len(dw.with_approvers(dw.with_approvers(self.base, None), None)["dispositions"]),
                         len(dw.with_approvers(self.base, None)["dispositions"]))
        with self.assertRaises(dw.WorkflowError):
            dw.with_approvers(self.base, "not-an-email")

    def test_no_submit_workflow_no_decisions(self):
        only = {**self.base, "workflows": [w for w in self.base["workflows"] if w["actionId"] == "roll-forward"]}
        manifest = dw.with_approvers(only, APPROVER)
        self.assertNotIn("decisions", manifest)
        self.assertEqual(dw.workflow_elements(manifest, "conn-1", "x")["decisions"], [])

    def test_approvals_has_decided_columns(self):
        cols = dict((cid, (name, typ)) for cid, name, typ in dw.TABLES["approvals"][1])
        self.assertEqual(cols["ap_dby"], ("Decided By", "text"))
        self.assertEqual(cols["ap_dat"], ("Decided At", "datetime"))
        table = next(t for t in self.elements()["tables"] if t["id"] == "approvals")
        self.assertIn("ap_dby", [c["id"] for c in table["columns"]])


SUBMIT_ONLY = {"version": 1, "dispositions": [], "workflows": [
    {"macro": "SubmitForApproval_Click", "class": "EXTERNAL_IO", "actionId": "submit-approval",
     "label": "Submit for approval", "effect": "insert-rows", "writes": True,
     "target": {"elementId": "approvals",
                "columns": ["Request", "Submitted By", "Approver", "Status", "Submitted At"]}}]}


class NoParametersStatusTests(unittest.TestCase):
    """Submit with no scenario switch or roll forward: no Parameters table, so the latest-status
    table sits on a one-row VALUES query (the shape proven live on 2026-11-09)."""

    def test_status_reads_a_one_row_values_query(self):
        wf = dw.workflow_elements(SUBMIT_ONLY, "conn-1", "FY27 forecast", no_status="Draft")
        self.assertEqual([t["id"] for t in wf["tables"]], [dw.APPROVALS])
        status = wf["status"]
        self.assertEqual(status["id"], dw.STATUS_KPI)
        self.assertEqual(status["source"], {"kind": "sql", "connectionId": "conn-1",
                                            "statement": "select * from (values\n(1)\n) as t(STATUS_ROW)"})
        (column,) = status["columns"]
        self.assertEqual(column["formula"],
                         'Coalesce(Lookup([Approvals/Status], Max([Approvals/Log Order]), '
                         '[Approvals/Log Order]), "Draft")')

    def test_status_sits_on_parameters_when_the_workbook_has_it(self):
        manifest = {**SUBMIT_ONLY, "workflows": SUBMIT_ONLY["workflows"] + [
            {"macro": "RollForwardMonth_Click", "class": "ACTION_WRITEBACK", "actionId": "roll-forward",
             "label": "Roll forward one month", "effect": "update-rows", "writes": True,
             "target": {"elementId": "parameters", "columns": ["Forecast Start Month"]}}]}
        wf = dw.workflow_elements(manifest, "conn-1", "FY27 forecast")
        self.assertEqual(wf["status"]["source"], {"kind": "table", "elementId": dw.PARAMETERS})

class SetupProgramHeaderTests(unittest.TestCase):
    def test_program_says_run_it_unchanged(self):
        wf = dw.workflow_elements(SUBMIT_ONLY, "conn-1", "FY27 forecast")
        import workbook_wire
        spec = workbook_wire.wire_workbook({"name": "wb", "pages": [{"id": "p", "name": "P",
                                                                    "elements": wf["tables"]}]})
        program = dw.emit_manual_commit_program(spec)
        head = program.split("\n\n", 1)[0]
        self.assertIn("Pass this whole file unchanged as one build_workbook program", head)
        self.assertIn("If it fails, report the error and stop.", head)


FIG = {"revenue": "Sum([PnL Combined/Amount])", "grossMargin": "Sum([PnL Combined/GM])",
       "ebitda": "Sum([PnL Combined/EBITDA])"}
GROWTH = ["North America Revenue", "EMEA Revenue", "APAC Revenue", "LATAM Revenue", "North America Cost of Revenue",
          "EMEA Cost of Revenue", "APAC Cost of Revenue", "LATAM Cost of Revenue", "North America Opex", "EMEA Opex"]
PLAN = {"lines": [*({"line": n, "kind": "growth", "rate": 0.01 + i / 1000} for i, n in enumerate(GROWTH)),
                  {"line": "APAC Opex", "kind": "flat"},
                  {"line": "LATAM Opex", "kind": "manual",
                   "manual": [610000.0, 615000.0, 622000.0, 630000.0, 640000.0, 655000.0]}],
        "periods": [{"label": "Jun-27", "num": 24, "end": "2027-06-30", "role": "actual"},
                    *({"label": m, "num": 25 + i, "end": None, "role": "forecast"}
                      for i, m in enumerate(["Jul-27", "Aug-27", "Sep-27", "Oct-27", "Nov-27", "Dec-27"]))]}


def snapshot_values(wf):
    button = next(b for b in wf["buttons"] if b["id"] == dw.BUTTON_IDS["snapshot-version"])
    (effect,) = button["actions"][0]["effects"]
    return effect["values"]


@unittest.skipUnless(DEMO_KIT.is_dir(), "demo kit not present (standalone skill copy)")
class VersionDriverTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = dw.build_workflow_manifest(DEMO_KIT / "demo-forecast-macros.bas")

    def elements(self, **extra):
        return dw.workflow_elements(self.manifest, "conn-1", "FY27 forecast", FIG, scenario_control="Scenario",
                                    no_status="Draft", **extra)

    def test_version_drivers_cover_every_rate_line_and_manual_cell(self):
        drivers = dw.version_drivers(PLAN)
        self.assertEqual([d.kind for d in drivers].count("rate"), 10)
        self.assertEqual([d.kind for d in drivers].count("manual"), 6)
        self.assertEqual(len({d.cid for d in drivers}), 16)
        rate = next(d for d in drivers if d.line == "EMEA Revenue")
        self.assertEqual((rate.name, rate.month_num, rate.default), ("EMEA Revenue rate", None, 0.011))
        manual = next(d for d in drivers if d.kind == "manual" and d.month_num == 26)
        self.assertEqual((manual.line, manual.name, manual.default), ("LATAM Opex", "LATAM Opex Aug-27", 615000.0))

    def test_driver_names_are_safe_column_paths(self):
        plan = json.loads(json.dumps(PLAN))
        plan["lines"][0]["line"] = "Salaries/Wages [US]"
        rate = dw.version_drivers(plan)[0]
        self.assertEqual(rate.line, "Salaries/Wages [US]")
        self.assertEqual(rate.name, "Salaries Wages US rate")
        self.assertIn('= "Salaries/Wages [US]"', dw.driver_formula(rate))
        plan["lines"][0]["line"] = 'Say "hi"'
        with self.assertRaisesRegex(dw.WorkflowError, "quote"):
            dw.version_drivers(plan)

    def test_snapshot_writes_every_driver_and_inputs_frozen(self):
        drivers = dw.version_drivers(PLAN)
        wf = self.elements(drivers=drivers)
        values = snapshot_values(wf)
        for d in drivers:
            formula = values[d.cid]["formula"]
            self.assertTrue(formula.startswith("Coalesce(Max(If("), formula)
            self.assertTrue(formula.endswith(f", {d.default!r})"), formula)   # unedited line: its default
            self.assertIn(f'= "{d.line}"', formula)
        self.assertIn("[Manual forecast/Month Num] = 26", values[next(d.cid for d in drivers if d.month_num == 26)]["formula"])
        self.assertEqual(values[dw.INPUTS_FROZEN[0]], {"type": "constant", "value": {"type": "text", "value": "Yes"}})
        versions = next(t for t in wf["tables"] if t["id"] == dw.VERSIONS)
        declared = {c["id"]: c for c in versions["columns"]}
        for d in drivers:
            self.assertEqual(declared[d.cid]["type"], "number")
        self.assertEqual(declared[dw.INPUTS_FROZEN[0]]["type"], "text")

    def test_migrated_history_never_sets_inputs_frozen(self):
        src = load_demo_source()
        seeds = {el: cols for el, cols, _rows in dw.source_history(src)}
        self.assertNotIn(dw.INPUTS_FROZEN[0], seeds[dw.VERSIONS])

    def test_no_drivers_keeps_todays_versions_table(self):
        before = json.loads((ROOT / "tests" / "fixtures" / "versions" / "workflow-elements-before.json").read_text())
        self.assertEqual(json.loads(json.dumps(self.elements(), sort_keys=True)), before)

    def test_snapshot_disposition_records_the_fidelity_difference(self):
        notes = " ".join(w.get("note", "") for w in self.manifest["workflows"] if w["actionId"] == "snapshot-version")
        self.assertIn("freezes the inputs and rebuilds the version", notes)


def load_demo_source():
    import importlib.util
    import tempfile
    spec = importlib.util.spec_from_file_location("make_demo_forecast_v", DEMO_KIT / "make-demo-forecast.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.build_demo_workbook(Path(tempfile.mkdtemp()))


if __name__ == "__main__":
    unittest.main()
