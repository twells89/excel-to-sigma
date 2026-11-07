#!/usr/bin/env python3
"""Approve / Send back: approver-only update-rows decisions on the Approvals table."""
from __future__ import annotations

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


if __name__ == "__main__":
    unittest.main()
