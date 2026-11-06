#!/usr/bin/env python3
"""The planning-app layout a forecast gets with --workflows (lib/forecast_app.py): page set, plumbing
hidden, scenario switch wired into the forecast, landed plan table in the model, KPI / summary formulas,
setup program extras, staging approval and the agent panel."""
from __future__ import annotations

import copy
import csv
import importlib.util
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(SCRIPTS))

import agent_enhance as ae  # noqa: E402
import code_rep  # noqa: E402
import demo_workflows as dw  # noqa: E402
import forecast_app  # noqa: E402
from forecast_plan import draft_plan  # noqa: E402


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SUFFIXES = (" Cost of Revenue", " Revenue", " Opex")


def region_of(line):
    return next(line[: -len(s)] for s in SUFFIXES if line.endswith(s))


class ForecastAppTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        tmp = Path(cls._tmp.name)
        cls.build = load("bfm_app", SCRIPTS / "build-forecast-model.py")
        src = load("make_demo_forecast_app", SCRIPTS / "make-demo-forecast.py").build_demo_workbook(tmp)
        plan, _ = draft_plan(src, intent="what-if", connection="conn-1", folder="fold-1",
                             name="Vantora FY27 Rolling Forecast — rehearsal r9")
        cls.plan = plan
        cls.manifest = dw.build_workflow_manifest(SCRIPTS / "demo-forecast-macros.bas")
        plan_csv = tmp / "plan.csv"
        with plan_csv.open("w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["MONTH", "LINE_ITEM", "REGION", "PLAN_AMOUNT"])
            for ln in plan["lines"]:
                for label in plan["fyPeriods"]:
                    w.writerow([label, ln["line"], region_of(ln["line"]), "100.00"])
        cls.landing = {"database": "CSA", "schema": "TJ", "tables": [
            {"excelTable": "tblActuals", "snowflakeTable": "FPA_X_ACTUALS", "columns": [
                {"header": "Posting Date", "name": "POSTING_DATE", "type": "DATE"},
                {"header": "Month", "name": "MONTH", "type": "VARCHAR"},
                {"header": "Line Item", "name": "LINE_ITEM", "type": "VARCHAR"},
                {"header": "Region", "name": "REGION", "type": "VARCHAR"},
                {"header": "Amount", "name": "AMOUNT", "type": "NUMBER(38,6)"}]},
            {"excelTable": "tblPlan", "snowflakeTable": "FPA_X_PLAN", "csv": str(plan_csv), "columns": [
                {"header": "Month", "name": "MONTH", "type": "VARCHAR"},
                {"header": "Line Item", "name": "LINE_ITEM", "type": "VARCHAR"},
                {"header": "Region", "name": "REGION", "type": "VARCHAR"},
                {"header": "Plan Amount", "name": "PLAN_AMOUNT", "type": "NUMBER(38,6)"}]}]}
        cls.landed = cls.build.landed_actuals_source(cls.landing, "CSA.TJ.FPA_X_ACTUALS", plan)
        cls.plan_src = cls.build.landed_plan_source(cls.landing, plan)
        cls.els = dict(cls.build.LOCAL_DM_IDS)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def spec(self, workflows="default", plan_src="default"):
        wf = self.manifest if workflows == "default" else workflows
        ps = self.plan_src if plan_src == "default" else plan_src
        return self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, landed=self.landed,
                                        workflows=wf, plan_src=ps)

    @staticmethod
    def by_id(spec):
        return {e["id"]: e for pg in spec["pages"] for e in pg["elements"]}

    @staticmethod
    def owners(spec):
        return {e["id"]: pg["id"] for pg in spec["pages"] for e in pg["elements"]}

    # ---------------------------------------------------------------- pages
    def test_four_visible_pages_then_hidden_plumbing(self):
        spec = self.spec()
        self.assertEqual([(p["id"], p["name"]) for p in spec["pages"]],
                         [*forecast_app.PAGES, ("model", "Model (sources)")])
        self.assertEqual(spec["pages"][-1].get("visibility"), "hidden")
        self.assertTrue(all(p.get("visibility") != "hidden" for p in spec["pages"][:-1]))
        owners = self.owners(spec)
        for plumbing in ("parameters", "versions", "approvals", "approvalStatus", "pnlCombined", "reportFacts",
                         "summaryFacts", "metricWrap", "planLines", "forecast", "actuals", "spineWrap"):
            self.assertEqual(owners[plumbing], "model", plumbing)
        forecast_kinds = {e["kind"] for e in spec["pages"][0]["elements"]}
        self.assertNotIn("input-table", forecast_kinds)   # no raw input tables on the main page
        self.assertNotIn("kpi-chart", forecast_kinds)
        # every element of every page is placed exactly once in the layout
        wired = self.build.workbook_wire.wire_workbook(spec)
        placed = re.findall(r'elementId="([^"]+)"', wired["contents"]["layout"])
        self.assertEqual(sorted(placed), sorted(owners))

    def test_canvas_title_drops_the_run_tag(self):
        self.assertEqual(self.build._app_titles(self.plan), ("Vantora", "FY27 rolling forecast"))
        canvas = json.dumps([pg["elements"] for pg in self.spec()["pages"][:-1]])
        self.assertNotIn("rehearsal", canvas.lower())
        self.assertNotIn("Department", canvas)
        self.assertNotIn("Vendor", canvas)

    def test_header_buttons_one_primary(self):
        b = self.by_id(self.spec())
        self.assertEqual([b[i]["text"] for i in ("btnRoll", "btnSnapshot", "btnSubmit")],
                         ["Roll forward", "Snapshot", "Submit for approval"])
        self.assertEqual([b[i]["appearance"] for i in ("btnRoll", "btnSnapshot", "btnSubmit")],
                         ["outline", "outline", "filled"])
        self.assertEqual({b[i]["size"] for i in ("btnRoll", "btnSnapshot", "btnSubmit")}, {"medium"})
        status = b["hdrStatus"]["body"]
        for name in ("Draft", "Submitted", "Approved"):
            self.assertIn(f'= "{name}"', status)
        self.assertIn('"Draft")', b["approvalStatus"]["columns"][0]["formula"])

    # ---------------------------------------------------------------- scenario
    def test_scenario_switch_drives_the_forecast_rate(self):
        b = self.by_id(self.spec())
        ctl = b["ctlScenario"]
        self.assertEqual((ctl["controlType"], ctl["controlId"], ctl["value"]), ("segmented", "Scenario", "Base"))
        self.assertEqual(ctl["source"]["values"], ["Downside", "Base", "Upside"])
        eff = next(c for c in b["forecast"]["columns"] if c["name"] == "Eff Rate")["formula"]
        self.assertIn('Switch([Scenario], "Downside", [FJ/Downside Spread], "Upside", [FJ/Upside Spread], 0)', eff)
        self.assertNotIn("Selected Scenario", json.dumps(b["parameters"]))
        snap = b["btnSnapshot"]["actions"][0]["effects"][0]["values"]
        self.assertEqual(snap["ve_scn"], {"type": "control", "control": "Scenario"})

    def test_spine_carries_the_scenario_spreads(self):
        dm = self.build.build_dm_spec(self.plan, self.landed, self.plan_src, app=True)
        spine = next(e for e in dm["pages"][0]["elements"] if e["id"] == "spine")
        names = [c["name"] for c in spine["columns"]]
        self.assertEqual(names[-2:], ["Downside Spread", "Upside Spread"])
        rows = re.findall(r"\(([^()]*)\)", spine["source"]["statement"].split("values", 1)[1])
        na = next(r for r in rows if r.startswith("'North America Revenue'"))
        self.assertEqual(na.split(",")[-2:], ["-0.004000", "0.004000"])
        flat = next(r for r in rows if r.startswith("'APAC Opex'"))
        self.assertEqual(flat.split(",")[-2:], ["0.000000", "0.000000"])
        plain = self.build.build_dm_spec(self.plan, self.landed)   # no app: spine unchanged
        self.assertNotIn("Spread", json.dumps(plain))

    # ---------------------------------------------------------------- plan
    def test_landed_plan_table_joins_the_model(self):
        dm = self.build.build_dm_spec(self.plan, self.landed, self.plan_src, app=True)
        el = {e["id"]: e for e in dm["pages"][0]["elements"]}
        self.assertEqual(el["planDetail"]["source"]["path"], ["CSA", "TJ", "FPA_X_PLAN"])
        cols = {c["name"]: c["formula"] for c in el["planDetail"]["columns"]}
        self.assertEqual(cols["Plan Amount"], "[FPA_X_PLAN/PLAN_AMOUNT]")
        self.assertIn('"Jan-27", 13', cols["Month Num"])
        self.assertIn('"Jan-27", MakeDate(2027, 1, 31)', cols["Month End"])
        self.assertEqual(self.build.map_dm_elements(dm)["planDetail"], "planDetail")
        self.assertEqual(self.build.map_dm_elements(dm)["metricMap"], "metricMap")
        self.assertEqual(self.plan_src["regions"]["LATAM Opex"], "LATAM")
        b = self.by_id(self.spec())
        self.assertEqual(b["reportFacts"]["source"]["kind"], "union")
        self.assertEqual([s["elementId"] for s in b["reportFacts"]["source"]["sources"]], ["pnlCombined", "planLines"])
        self.assertIn('"LATAM Opex", "LATAM"', next(c for c in b["reportFacts"]["columns"] if c["name"] == "Region")["formula"])

    def test_plan_month_outside_the_forecast_fails_loudly(self):
        landing = copy.deepcopy(self.landing)
        bad = Path(self._tmp.name) / "bad.csv"
        bad.write_text("MONTH,LINE_ITEM,REGION,PLAN_AMOUNT\nJan-30,LATAM Opex,LATAM,1\n")
        landing["tables"][1]["csv"] = str(bad)
        with self.assertRaises(self.build.PlanError):
            self.build.landed_plan_source(landing, self.plan)

    def test_kpis_compare_to_plan_and_prior_year(self):
        b = self.by_id(self.spec())
        rev = b["kpiRevenueText"]["body"]
        self.assertIn("FY27 revenue", rev)
        self.assertIn("vs plan", rev)
        self.assertIn("vs FY26 actual", rev)
        self.assertIn('[Forecast Facts/Period] = "Plan"', rev)
        self.assertIn("[Forecast Facts/Month Num] >= 13 and [Forecast Facts/Month Num] <= 24", rev)
        self.assertIn("[Forecast Facts/Month Num] >= 1 and [Forecast Facts/Month Num] <= 12", rev)  # FY26
        self.assertIn("pts", b["kpiMarginText"]["body"])
        months = b["kpiMonthsText"]["body"]
        self.assertIn("of 12", months)
        self.assertIn("MakeDate(2027, 12, 1)", months)
        for el in b.values():
            if el.get("kind") == "text":
                self.assertEqual(forecast_app_disallowed(el["body"]), [], el["id"])

    def test_without_plan_table_no_plan_comparisons(self):
        b = self.by_id(self.spec(plan_src=None))
        self.assertNotIn("planLines", b)
        self.assertNotIn("regionVariance", b)
        self.assertNotIn("vs plan", b["kpiRevenueText"]["body"])
        self.assertEqual([c["name"] for c in b["plSummary"]["columns"] if not c.get("hidden")],
                         ["Line", "FY26 actual", "FY27 forecast"])

    def test_pl_summary_rows_and_columns(self):
        self.assertEqual([m for m, _ in self.build.summary_metrics(self.plan)],
                         ["Revenue", "Gross Margin", "Operating Expenses", "EBITDA"])
        mm = self.build.metric_map_element(self.plan)
        self.assertIn("('Gross margin',2,'Cost of Revenue',-1.0)", mm["source"]["statement"])
        b = self.by_id(self.spec())
        cols = {c["name"]: c for c in b["plSummary"]["columns"]}
        self.assertEqual([c["name"] for c in b["plSummary"]["columns"] if not c.get("hidden")],
                         ["Line", "FY26 actual", "FY27 forecast", "Plan", "Vs plan"])
        self.assertTrue(cols["Vs plan"]["formula"].startswith("(SumIf("))   # ratio of the two sums, not of terms
        self.assertEqual(cols["FY27 forecast"]["format"]["suffix"], "M")

    def test_trend_bars_by_period_with_dashed_plan_line(self):
        t = self.by_id(self.spec())["revTrend"]
        self.assertEqual(t["kind"], "combo-chart")
        self.assertEqual(t["yAxis"]["columnIds"], [{"columnId": "tr_rev", "type": "bar"},
                                                   {"columnId": "tr_plan", "type": "line"}])
        self.assertEqual(t["color"]["column"], "tr_cat")
        self.assertIn("<= 18", next(c for c in t["columns"] if c["id"] == "tr_cat")["formula"])
        self.assertEqual(t["seriesLineAreaStyle"][0]["style"]["line"]["style"], "dashed")

    def test_detail_page_shows_the_fy_in_thousands(self):
        b = self.by_id(self.spec())
        pivot = b["pnlPivot"]
        self.assertEqual(pivot["filters"][0]["startDate"], "2027-01-01")
        self.assertTrue(next(c for c in pivot["columns"] if c["id"] == "p_amt")["formula"].endswith("/ 1000"))
        self.assertNotIn("Txn Date", [c["name"] for c in b["glTable"]["columns"]])

    def test_plain_forecast_path_unchanged(self):
        plain = self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, landed=self.landed)
        self.assertEqual([p["id"] for p in plain["pages"]], ["pnl", "detail", "inputs", "model"])
        self.assertIn("kpiRev", {e["id"] for pg in plain["pages"] for e in pg["elements"]})

    # ---------------------------------------------------------------- setup program
    def test_setup_program_reapplies_formats_and_series_colour(self):
        wired = self.build.workbook_wire.wire_workbook(self.spec())
        program = dw.emit_manual_commit_program(wired, {"start": "2027-07-01"},
                                                series_colors=forecast_app.SERIES_COLORS)
        self.assertIn("c = t.get_column('RATE')\nif not isinstance(c, LiteralEditableColumn):", program)
        self.assertIn("c.set_format_pattern('.1%')", program)
        self.assertIn("c = t.get_column('MONTH_NUM')", program)
        self.assertIn("c.set_visible(False)", program)
        self.assertIn("t.upsert_rows([{c_start: '2027-07-01'}])", program)
        self.assertNotIn("c_scn", program)
        self.assertIn("ch.set_series_color(s, '#475569')", program)
        self.assertLess(program.index("set_format_pattern"), program.index("set_edit_mode('draft')"))
        compile(program, "program", "exec")

    # ---------------------------------------------------------------- agent
    def test_agent_sets_the_switch_and_fills_the_panel(self):
        wired = self.build.workbook_wire.wire_workbook(self.spec())
        macros = [{"name": w["macro"], "class": w["class"]} for w in self.manifest["workflows"]]
        rec, frag = ae.recommend_for_workbook(macros, self.manifest, wired)
        tools = {t["toolId"]: t for t in frag["agents"][0]["tools"]}
        self.assertEqual(tools["set-scenario"]["steps"][0]["effect"], "set-control-value")
        self.assertEqual(tools["set-scenario"]["steps"][0]["control"], "Scenario")
        self.assertIs(tools["set-scenario"]["requiresApproval"], True)
        self.assertEqual(frag["agents"][0]["greeting"]["mode"], "static")
        merged = ae.merge_fragment(wired, frag)
        panel = re.search(r'<Container elementId="agentPanel".*?</Container>', merged["contents"]["layout"], re.S).group(0)
        self.assertIn('elementId="chat"', panel)
        self.assertEqual(code_rep.workbook_page_by_element(merged)["chat"]["id"], "pgForecast")

    # ---------------------------------------------------------------- staging
    def test_staging_approval_must_cover_exact_issues_folder_and_names(self):
        issue = {"code": "VBA_PROJECT", "severity": "review", "message": "m"}
        report = {"sourceSha256": "a" * 64, "issues": [issue]}
        P = {"folderId": "fold-1"}
        good = {"owner": "TJ Wells", "reviewer": "Claude (controller)", "folderId": "fold-1", "sourceSha256": "a" * 64,
                "sourceReviewApproval": {"status": "approved_for_staging", "issues": [issue]}}
        self.build.check_staging_approval(P, report, good)
        for change in ({"owner": ""}, {"folderId": "other"}, {"sourceSha256": "b" * 64},
                       {"sourceReviewApproval": {"status": "pending", "issues": [issue]}},
                       {"sourceReviewApproval": {"status": "approved_for_staging", "issues": []}},
                       {"sourceReviewApproval": {"status": "approved_for_staging",
                                                 "issues": [{**issue, "message": "edited"}]}}):
            with self.assertRaises(self.build.StagingError, msg=str(change)):
                self.build.check_staging_approval(P, report, {**good, **change})
        template = self.build.staging_template(P, report)
        self.assertEqual((template["owner"], template["sourceReviewApproval"]["status"]), ("", "pending"))


def forecast_app_disallowed(body):
    import presentation_style
    return presentation_style.disallowed_span_styles(body)


if __name__ == "__main__":
    unittest.main()
