"""Phase 6 Enhance: agent recommendation and document fragment."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))
import agent_enhance as ae  # noqa: E402
import code_rep  # noqa: E402

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
        writes = ("roll-forward", "snapshot-version", "submit-approval", "set-scenario", "set-driver")
        self.assertEqual(set(tools), set(writes))
        for a in writes:  # the spec: approval on every write
            self.assertIs(tools[a]["requiresApproval"], True)
        for phrase in ("roll forward", "snapshot", "submit for approval", "switch scenario",
                       "growth driver", "approval first"):
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
                                          ["forecast", "parameters", "versions", "approvals", "rateInput"],
                                          figures={"revenue": "Sum([PnL Combined/Amount])"})
        self.assertEqual(frag["chatElement"], spike["chat"])
        agent = frag["agents"][0]
        self.assertEqual(set(agent) - {"greeting"}, set(spike["agent"]))   # greeting: static, dry-run verified
        self.assertEqual(agent["greeting"]["mode"], "static")
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
        self.assertIs(tools["set-scenario"]["requiresApproval"], True)
        self.assertEqual(tools["set-scenario"]["steps"][0]["values"]["pa_scn"],
                         {"type": "agent-input", "inputName": "scenario"})
        for el in ("parameters", "versions", "approvals"):
            self.assertIn(f'@dataSource("{el}")', agent["instructions"])
        for tid in tools:
            self.assertIn(f'@tool("{tid}")', agent["instructions"])
        drv = tools["set-driver"]
        self.assertIs(drv["requiresApproval"], True)
        # An input table's only primary key is its system row ID (live: single-row on LINE_ITEM is refused,
        # "Primary key: ID not in update row definition") and a row formula cannot read an agent input, so the
        # tool puts the line item into a hidden text control and updates the rows whose line item equals it.
        pick, step = drv["steps"]
        self.assertEqual((pick["effect"], pick["control"]), ("set-control-value", ae.DRIVER_CONTROL))
        self.assertEqual(pick["value"], {"type": "agent-input", "inputName": "line_item"})
        self.assertEqual((step["effect"], step["tableElementId"]), ("update-rows", "rateInput"))
        self.assertEqual(step["whichRows"], {"type": "formula",
                                             "formula": f"[Line item] = [{ae.DRIVER_CONTROL}]"})
        self.assertEqual(step["values"], {"RATE": {"type": "agent-input", "inputName": "rate"}})
        self.assertEqual(frag["controlElements"], [ae.driver_control()])
        self.assertEqual(ae.driver_control()["controlType"], "text")
        self.assertIn('@tool("set-driver")', agent["instructions"])
        self.assertIn('@dataSource("rateInput")', agent["instructions"])
        # without a growth-rate table in the workbook there is no driver tool
        bare = ae.agent_document_fragment(rec, "n", ["parameters", "versions", "approvals"],
                                          figures={"revenue": "1"})
        self.assertNotIn("set-driver", [t["toolId"] for t in bare["agents"][0]["tools"]])
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

    def _input_table_workbook(self):
        """A planning-app shaped workbook: rateInput is a real input table (system ID first), the model page hidden."""
        wb = self._workbook()
        doc = wb["contents"]
        rate = next(e for e in doc["elements"] if e["id"] == "rateInput")
        rate.update({"kind": "input-table", "name": "Growth rates", "columns": [
            {"id": "ID", "name": "Row ID", "hidden": True}, {"id": "LINE_ITEM", "type": "text", "name": "Line item"},
            {"id": "RATE", "type": "number", "name": "Monthly growth rate"}]})
        doc["pages"].append({"id": "model", "name": "Model (sources)", "visibility": "hidden"})
        doc["elements"].append({"id": "helper", "kind": "table"})
        doc["layout"] += ('\n<Page type="grid" id="model">\n'
                          '  <Element elementId="helper" gridColumn="1 / 25" gridRow="1 / 13"/>\n</Page>')
        return wb

    def test_agent_row_selectors_key_only_on_the_input_table_primary_key(self):
        macros, wfs = demo()
        wb = self._input_table_workbook()
        rec, frag = ae.recommend_for_workbook(macros, wfs, wb)
        merged = ae.merge_fragment(wb, frag)
        doc = merged["contents"]
        for tool in doc["agents"][0]["tools"]:
            for step in tool["steps"]:
                rows = step.get("whichRows") or {}
                if rows.get("type") == "single-row":   # only the system row ID is a primary key
                    self.assertLessEqual(set(rows["primaryKeys"]), {"ID"}, tool["toolId"])
                if step["effect"] == "set-control-value":
                    self.assertIn(step["control"], {e.get("controlId") for e in doc["elements"]})
        # the driver control is plumbing: on the hidden page, never on a visible one
        self.assertEqual(code_rep.workbook_page_by_element(merged)["ctlDriverLineItem"]["id"], "model")
        self.assertIn('<Element elementId="ctlDriverLineItem" gridColumn="1 / 25" gridRow="13 / 15"/>', doc["layout"])
        # the selector names the column the workbook actually shows
        drv = next(t for t in doc["agents"][0]["tools"] if t["toolId"] == "set-driver")
        self.assertEqual(drv["steps"][1]["whichRows"]["formula"], "[Line item] = [DriverLineItem]")
        plain = copy.deepcopy(wb)          # the plain forecast builder titles the column "Line Item"
        next(e for e in plain["contents"]["elements"] if e["id"] == "rateInput")["columns"][1]["name"] = "Line Item"
        _, frag2 = ae.recommend_for_workbook(macros, wfs, plain)
        drv2 = next(t for t in frag2["agents"][0]["tools"] if t["toolId"] == "set-driver")
        self.assertEqual(drv2["steps"][1]["whichRows"]["formula"], "[Line Item] = [DriverLineItem]")
        # a tool keyed on a business column is refused before anything is posted
        bad = copy.deepcopy(frag)
        drv = next(t for t in bad["agents"][0]["tools"] if t["toolId"] == "set-driver")
        drv["steps"] = [{"kind": "effect", "effect": "update-rows", "tableElementId": "rateInput",
                         "whichRows": {"type": "single-row", "primaryKeys": {
                             "LINE_ITEM": {"type": "agent-input", "inputName": "line_item"}}},
                         "values": {"RATE": {"type": "agent-input", "inputName": "rate"}}}]
        with self.assertRaisesRegex(ValueError, "LINE_ITEM"):
            ae.merge_fragment(wb, bad)
        # ... and so is a control step whose control the workbook does not have
        bad = copy.deepcopy(frag)
        bad["controlElements"] = []
        with self.assertRaisesRegex(ValueError, "DriverLineItem"):
            ae.merge_fragment(wb, bad)

    def _workbook(self):
        import demo_workflows as dw
        buttons = [{"id": bid, "kind": "button", "text": bid, "actions": [{"id": "a-" + aid, "trigger": "on-click",
                    "effects": dw.action_effects(aid, "FY27 forecast", {"revenue": "Sum([PnL Combined/Amount])"})}]}
                   for aid, bid in dw.BUTTON_IDS.items()]
        ids = ["forecast", "parameters", "versions", "approvals", "rateInput"]
        els = [{"id": i, "kind": "table"} for i in ids] + buttons
        rows = "".join(f'  <Element elementId="{e["id"]}" gridColumn="1 / 25" gridRow="{1 + 4 * n} / {5 + 4 * n}"/>\n'
                       for n, e in enumerate(els))
        return {"name": "wb", "contents": {"schemaVersion": 1, "kind": "workbook",
                "pages": [{"id": "pnl", "name": "P&L"}], "elements": els,
                "layout": '<?xml version="1.0" encoding="utf-8"?>\n<Page type="grid" id="pnl">\n' + rows + "</Page>"}}

    def test_button_effects_reused_by_agent_tools(self):
        import demo_workflows as dw
        macros, wfs = demo()
        wb = self._workbook()
        eff = ae.button_effects(wb)
        self.assertEqual(set(eff), set(dw.BUTTON_IDS))
        rec = ae.recommend_agent(macros, wfs)
        frag = ae.agent_document_fragment(rec, "A", ["forecast", "parameters", "versions", "approvals", "rateInput"],
                                          effects_by_action=eff)
        tools = {t["toolId"]: t for t in frag["agents"][0]["tools"]}
        for aid in dw.BUTTON_IDS:
            self.assertEqual([{k: v for k, v in s.items() if k != "kind"} for s in tools[aid]["steps"]], eff[aid])

    def test_cli_merge_into_and_emit(self):
        import demo_workflows as dw
        wb = self._workbook()
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = tmp / "wb.json"
            src.write_text(json.dumps(wb))
            before = src.read_text()
            base = [sys.executable, str(ROOT / "scripts" / "recommend-agent.py"),
                    "--macros", str(FX / "demo-macros.json"), "--workflows", str(FX / "demo-workflows.json")]
            subprocess.run(base + ["--merge-into", str(src), "--out", str(tmp / "merged.json")],
                           capture_output=True, text=True, check=True)
            self.assertEqual(src.read_text(), before)  # input untouched
            merged = json.loads((tmp / "merged.json").read_text())["contents"]
            agent = merged["agents"][0]
            self.assertEqual(len(agent["tools"]), 5)
            self.assertIn({"id": "chat", "kind": "chat", "agentId": agent["id"]}, merged["elements"])
            self.assertIn('elementId="chat"', merged["layout"])
            eff = ae.button_effects(wb)
            roll = next(t for t in agent["tools"] if t["toolId"] == "roll-forward")
            self.assertEqual(roll["steps"][0]["values"], eff["roll-forward"][0]["values"])
            subprocess.run(base + ["--emit", str(tmp / "frag.json"), "--workbook", str(src)],
                           capture_output=True, text=True, check=True)
            frag = json.loads((tmp / "frag.json").read_text())
            self.assertEqual(frag["chatElement"]["kind"], "chat")
            self.assertEqual(len(frag["agents"][0]["tools"]), 5)
            r = subprocess.run(base + ["--merge-into", str(src)], capture_output=True, text=True)
            self.assertNotEqual(r.returncode, 0)  # --out required

    def test_cli_prints_pitch(self):
        out = subprocess.run([sys.executable, str(ROOT / "scripts" / "recommend-agent.py"),
                              "--macros", str(FX / "demo-macros.json"),
                              "--workflows", str(FX / "demo-workflows.json")],
                             capture_output=True, text=True, check=True).stdout
        self.assertIn("4 of your 7 macros", out)


if __name__ == "__main__":
    unittest.main()
