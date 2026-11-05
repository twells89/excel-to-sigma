"""Phase 6 Enhance: agent recommendation and document fragment."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))
import agent_enhance as ae  # noqa: E402

FX = ROOT / "tests" / "fixtures" / "agent-enhance"


def fx(name):
    return json.loads((FX / name).read_text(encoding="utf-8"))


def demo():
    return fx("demo-macros.json"), fx("demo-workflows.json")["workflows"]


class AgentEnhanceTests(unittest.TestCase):
    def test_act_profile_when_write_workflows(self):
        macros, wfs = demo()
        rec = ae.recommend_agent(macros, wfs)
        self.assertTrue(rec["recommend"])
        self.assertEqual(rec["profile"], "act")
        tools = {t["actionId"]: t for t in rec["tools"]}
        writes = ("roll-forward", "snapshot-version", "submit-approval")
        self.assertEqual(set(tools), {*writes, "set-scenario"})
        for a in writes:
            self.assertIs(tools[a]["requiresApproval"], True)
        self.assertIs(tools["set-scenario"]["requiresApproval"], False)
        for phrase in ("roll forward", "snapshot", "submit for approval", "switch scenario", "approval first"):
            self.assertIn(phrase, rec["pitch"])
        self.assertNotIn("email", rec["pitch"].replace("never sends email", ""))

    def test_analyze_only_when_no_actions(self):
        macros = [{"name": "RecalcAll", "class": "RECALC"}, {"name": "FormatReport", "class": "FORMAT_COSMETIC"}]
        rec = ae.recommend_agent(macros, [])
        self.assertIn(rec["profile"], ("analyze", None))
        self.assertEqual(rec["tools"], [])
        # even if a stray workflow is passed, a cosmetic/recalc-only inventory never gets write tools
        stray = [{"macro": "X", "class": "RECALC", "actionId": "roll-forward", "label": "Roll", "writes": True}]
        self.assertEqual(ae.recommend_agent(macros, stray)["tools"], [])
        with self.assertRaises(ValueError):
            ae.agent_document_fragment(rec, "n", ["forecast"])

    def test_no_macros_no_recommendation(self):
        for macros in ([], {}, {"macros": []}):
            rec = ae.recommend_agent(macros, [])
            self.assertFalse(rec["recommend"])
            self.assertIsNone(rec["profile"])
            self.assertEqual(rec["tools"], [])

    def test_fragment_shape_matches_spike(self):
        spike = fx("spike-agent.json")
        macros, wfs = demo()
        rec = ae.recommend_agent(macros, wfs)
        frag = ae.agent_document_fragment(rec, "FP&A forecast agent",
                                          ["forecast", "parameters", "versions", "approvals"],
                                          figures={"revenue": "Sum([PnL Combined/Amount])"})
        self.assertEqual(frag["chatElement"], spike["chat"])
        agent = frag["agents"][0]
        self.assertEqual(set(agent), set(spike["agent"]))
        self.assertEqual(agent["dataSources"][0], spike["agent"]["dataSources"][0])
        spike_tool = next(t for t in spike["agent"]["tools"] if t["toolId"] == "roll-forward")
        for tool in agent["tools"]:
            self.assertLessEqual({"toolId", "kind", "name", "description", "steps"}, set(tool))
            self.assertEqual(tool["kind"], "action")
            for step in tool["steps"]:
                self.assertEqual(step["kind"], "effect")
        tools = {t["toolId"]: t for t in agent["tools"]}
        got = tools["roll-forward"]["steps"][0]
        want = spike_tool["steps"][0]
        for key in ("kind", "effect", "tableElementId"):
            self.assertEqual(got[key], want[key])
        self.assertEqual(list(got["values"].values()), list(want["values"].values()))  # column ids differ
        self.assertEqual(set(got), set(want))
        self.assertEqual(got["whichRows"]["formula"].lower(), want["whichRows"]["formula"].lower())
        self.assertIs(tools["roll-forward"]["requiresApproval"], True)
        self.assertNotIn("requiresApproval", tools["set-scenario"])
        self.assertEqual(tools["set-scenario"]["steps"][0]["values"]["pa_scn"],
                         {"type": "agent-input", "inputName": "scenario"})
        for el in ("parameters", "versions", "approvals"):
            self.assertIn(f'@dataSource("{el}")', agent["instructions"])
        for tid in tools:
            self.assertIn(f'@tool("{tid}")', agent["instructions"])
        # never name system columns bare
        for col in ("CREATED_BY", "CREATED_AT"):
            self.assertNotIn(col, json.dumps(frag))

    def test_merge_adds_agent_and_chat_to_layout(self):
        macros, wfs = demo()
        rec = ae.recommend_agent(macros, wfs)
        wb = {"name": "wb", "contents": {
            "schemaVersion": 1, "kind": "workbook", "pages": [{"id": "pnl", "name": "P&L"}],
            "elements": [{"id": "forecast", "kind": "table"}, {"id": "parameters", "kind": "input-table"}],
            "layout": '<?xml version="1.0" encoding="utf-8"?>\n<Page type="grid" id="pnl">\n'
                      '  <Element elementId="forecast" gridColumn="1 / 25" gridRow="1 / 13"/>\n'
                      '  <Element elementId="parameters" gridColumn="1 / 25" gridRow="13 / 25"/>\n</Page>'}}
        frag = ae.agent_document_fragment(rec, "A", ["forecast", "parameters"], figures={"revenue": "1"})
        merged = ae.merge_fragment(wb, frag)
        self.assertEqual(merged["contents"]["agents"][0]["id"], "fpa-agent")
        self.assertIn('elementId="chat" gridColumn="1 / 25" gridRow="25 / 49"', merged["contents"]["layout"])
        self.assertNotIn("agents", wb["contents"])  # input untouched
        with self.assertRaises(ValueError):
            ae.merge_fragment(merged, frag)

    def test_cli_prints_pitch(self):
        out = subprocess.run([sys.executable, str(ROOT / "scripts" / "recommend-agent.py"),
                              "--macros", str(FX / "demo-macros.json"),
                              "--workflows", str(FX / "demo-workflows.json")],
                             capture_output=True, text=True, check=True).stdout
        self.assertIn("4 of your 7 macros", out)


if __name__ == "__main__":
    unittest.main()
