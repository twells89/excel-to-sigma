#!/usr/bin/env python3
"""The planning-app layout a forecast gets with --workflows (lib/forecast_app.py): page set, plumbing
hidden, scenario switch wired into the forecast, landed plan table in the model, KPI / summary formulas,
setup program extras, staging approval and the agent panel."""
from __future__ import annotations

import copy
import csv
import importlib.util
import io
import json
import re
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
# The demo kit (docs/demo-kit: the demo source generator and its VBA module) lives outside the skill;
# a standalone copy of the skill has no demo kit, so the tests that need it skip.
DEMO_KIT = ROOT.parents[1] / "docs" / "demo-kit"
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "tests"))

import agent_enhance as ae  # noqa: E402
import code_rep  # noqa: E402
import demo_workflows as dw  # noqa: E402
import forecast_app  # noqa: E402
from forecast_plan import draft_plan  # noqa: E402
import planning_app_eval as pe  # noqa: E402


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SUFFIXES = (" Cost of Revenue", " Revenue", " Opex")


def region_of(line):
    return next(line[: -len(s)] for s in SUFFIXES if line.endswith(s))


@unittest.skipUnless(DEMO_KIT.is_dir(), "demo kit not present (standalone skill copy)")
class ForecastAppTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        tmp = Path(cls._tmp.name)
        cls.build = load("bfm_app", SCRIPTS / "build-forecast-model.py")
        src = load("make_demo_forecast_app", DEMO_KIT / "make-demo-forecast.py").build_demo_workbook(tmp)
        plan, _ = draft_plan(src, intent="what-if", connection="conn-1", folder="fold-1",
                             name="Vantora FY27 Rolling Forecast — rehearsal r9")
        cls.plan = plan
        cls.manifest = dw.build_workflow_manifest(DEMO_KIT / "demo-forecast-macros.bas")
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

    def spec(self, workflows="default", plan_src="default", history=None):
        wf = self.manifest if workflows == "default" else workflows
        ps = self.plan_src if plan_src == "default" else plan_src
        return self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, landed=self.landed,
                                        workflows=wf, plan_src=ps, history=history)

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
        for plumbing in ("boundary", "parameters", "versions", "approvals", "approvalStatus", "pnlCombined", "reportFacts",
                         "summaryFacts", "metricWrap", "planLines", "forecast", "actuals", "spineWrap"):
            self.assertEqual(owners[plumbing], "model", plumbing)
        forecast_kinds = {e["kind"] for e in spec["pages"][0]["elements"]}
        self.assertNotIn("input-table", forecast_kinds)   # no raw input tables on the main page
        self.assertNotIn("kpi-chart", forecast_kinds)
        # every element of every page is placed exactly once in the layout
        wired = self.build.workbook_wire.wire_workbook(spec)
        placed = re.findall(r'elementId="([^"]+)"', wired["contents"]["layout"])
        self.assertEqual(sorted(placed), sorted(owners))

    def test_planning_app_with_a_partial_manifest_has_only_its_tables(self):
        def only(*ids):
            return {**self.manifest, "workflows": [w for w in self.manifest["workflows"] if w["actionId"] in ids]}
        for ids, present, absent in (
                (("snapshot-version",), {"versions", "versionsView", "btnSnapshot"},
                 {"parameters", "approvals", "approvalStatus", "approvalsView", "hdrStatus", "btnSubmit", "btnRoll"}),
                (("submit-approval",), {"approvals", "approvalStatus", "approvalsView", "hdrStatus", "btnSubmit"},
                 {"parameters", "versions", "versionsView", "btnSnapshot"}),
                (("set-scenario",), {"parameters"}, {"versions", "approvals", "pgVersions", "hdrStatus"})):
            with self.subTest(ids=ids):
                spec = self.spec(only(*ids))
                owners = self.owners(spec)
                self.assertTrue(present <= set(owners) | {p["id"] for p in spec["pages"]}, present - set(owners))
                self.assertFalse(absent & (set(owners) | {p["id"] for p in spec["pages"]}))
                text = json.dumps(spec)
                if "parameters" not in owners:   # the boundary falls back to the source's first forecast month
                    self.assertNotIn("[Parameters/", text)
                if "approvals" not in owners:
                    self.assertNotIn("[Approvals/", text)
                    self.assertNotIn("Latest approval status", text)
                wired = self.build.workbook_wire.wire_workbook(spec)
                placed = re.findall(r'elementId="([^"]+)"', wired["contents"]["layout"])
                self.assertEqual(sorted(placed), sorted(owners))

    def test_scenario_note_names_the_plans_own_scenarios(self):
        rename = {"Downside": "Bear", "Upside": "Bull"}
        P = copy.deepcopy(self.plan)
        P["scenarios"] = json.loads(json.dumps(P["scenarios"]).replace('"Downside"', '"Bear"')
                                    .replace('"Upside"', '"Bull"'))
        self.assertEqual(P["scenarios"]["names"], [rename.get(n, n) for n in self.plan["scenarios"]["names"]])
        spec = self.build.build_wb_spec(P, "dm-1", self.els, landed=self.landed, workflows=self.manifest,
                                        plan_src=self.plan_src)
        head = self.by_id(spec)["asmHead"]["text"] if "text" in self.by_id(spec)["asmHead"] else json.dumps(
            self.by_id(spec)["asmHead"])
        self.assertIn("Bear and Bull move every growth rate", head)
        self.assertNotIn("Downside", head)

    def test_header_says_live_on_snowflake_only_when_landed(self):
        title = json.dumps(self.by_id(self.spec())["hdrTitle"])
        self.assertIn("live on Snowflake", title)
        plain = self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, workflows=self.manifest)
        self.assertNotIn("Snowflake", json.dumps(self.by_id(plain)["hdrTitle"]))

    def test_set_driver_example_is_the_rate_tables_first_line(self):
        first = self.build.seed_rows(self.plan)[0][0][0]
        macros = [{"name": w["macro"], "class": w["class"]} for w in self.manifest["workflows"]]
        wired = self.build.workbook_wire.wire_workbook(self.spec())
        merged = self.build.merge_agent(wired, macros, self.manifest, quiet=True,
                                        driver_example=self.build.driver_example(self.plan))
        tool = next(t for a in merged["contents"]["agents"] for t in a["tools"] if t["toolId"] == "set-driver")
        self.assertIn(f"the exact Line Item, for example {first};", tool["description"])

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
        self.assertEqual(next(c for c in t["columns"] if c["id"] == "tr_cat")["formula"], "[Forecast Facts/Phase]")
        self.assertEqual(t["seriesLineAreaStyle"][0]["style"]["line"]["style"], "dashed")

    # ---------------------------------------------------------------- roll forward = month close
    START_KEY = "Max([Forecast Boundary/Start Key])"

    def test_actual_forecast_split_follows_the_forecast_start_month(self):
        b = self.by_id(self.spec())
        act = {c["name"]: c for c in b["actuals"]["columns"]}
        self.assertEqual(act["Month Key"]["formula"],
                         'DatePart("year", [GL Detail/Month End]) * 12 + DatePart("month", [GL Detail/Month End])')
        self.assertEqual(act["Window"]["formula"], f'If([Month Key] < {self.START_KEY}, "Actual", "Forecast")')
        self.assertEqual(b["actuals"]["filters"], [{"id": "a_wf", "columnId": act["Window"]["id"], "kind": "list",
                                                    "mode": "include", "values": ["Actual"]}])
        fc = {c["name"]: c for c in b["forecast"]["columns"]}
        self.assertEqual(fc["Start Key"]["formula"], self.START_KEY)
        self.assertEqual(fc["Landed"]["formula"], "Lookup(Sum([Actuals/Amount]), [Line Item], [Actuals/Line Item], "
                                                  "[Month Key], [Actuals/Month Key])")
        self.assertEqual(fc["Window"]["formula"], 'If([Month Key] >= [Start Key], "Forecast", IsNull([Landed]), '
                                                  '"Awaiting actuals", "Actual")')
        self.assertEqual(b["forecast"]["filters"], [{"id": "f_wf", "columnId": fc["Window"]["id"], "kind": "list",
                                                     "mode": "include", "values": ["Forecast", "Awaiting actuals"]}])
        # a closed month whose actuals have not landed keeps its forecast, flagged in the Period column
        self.assertEqual(fc["Period"]["formula"], 'If([Window] = "Awaiting actuals", "Awaiting actuals", "Forecast")')
        self.assertEqual(act["Period"]["formula"], '"Actual"')

    def test_forecast_reseeds_from_the_last_closed_month(self):
        fc = {c["name"]: c["formula"] for c in self.by_id(self.spec())["forecast"]["columns"]}
        self.assertEqual(fc["Base Key"], "[Start Key] - 1")
        self.assertEqual(fc["Last Actual"], "Lookup(Sum([Actuals/Amount]), [Line Item], [Actuals/Line Item], "
                                            "[Base Key], [Actuals/Month Key])")
        last_key = 2027 * 12 + 6   # Jun-27, the source's last actual month
        # a closed month whose actuals have not landed seeds from its own frozen forecast value
        self.assertEqual(fc["Seed"], f'Coalesce([Last Actual], If([Kind] = "flat", [June Base], '
                                     f'[June Base] * (1 + [Eff Rate]) ^ ([Base Key] - {last_key})))')
        self.assertEqual(fc["N"], 'If([Kind] = "yoy", [FJ/N], [Month Key] - [Base Key])')
        self.assertEqual(fc["Amount"], 'If([Kind] = "manual", [Eff Manual], [Kind] = "flat", [Seed], '
                                       '[Kind] = "yoy", [June Base] * (1 + [Eff Rate]) ^ [N], '
                                       '[Seed] * (1 + [Eff Rate]) ^ [N])')

    def test_every_surface_reads_the_one_boundary(self):
        spec = self.spec()
        b = self.by_id(spec)
        # defined once: a one-row hidden element; nothing else reads Parameters.Forecast Start Month
        bd = b["boundary"]
        self.assertEqual((bd["name"], bd["source"]["kind"]), ("Forecast Boundary", "sql"))
        cols = {c["name"]: c["formula"] for c in bd["columns"]}
        self.assertEqual(cols["Start Month"], "Coalesce(Max([Parameters/Forecast Start Month]), MakeDate(2027, 7, 1))")
        self.assertEqual(cols["Start Key"], 'DatePart("year", [Start Month]) * 12 + DatePart("month", [Start Month])')
        text = json.dumps(spec)
        self.assertEqual(text.count("[Parameters/Forecast Start Month]"), 1)
        self.assertEqual(text.count("MakeDate(2027, 7, 1)"), 1)
        self.assertNotIn("<= 18", text)
        phase = next(c for c in b["reportFacts"]["columns"] if c["name"] == "Phase")["formula"]
        self.assertIn(f">= {self.START_KEY}", phase)
        for el in ("hdrTitle", "kpiMonthsText"):
            self.assertIn("Max([Forecast Boundary/", b[el]["body"], el)
        roll = b["btnRoll"]["actions"][0]["effects"][0]["values"]["pa_start"]["formula"]
        self.assertEqual(roll, 'DateAdd("month", 1, Max([Forecast Boundary/Start Month]))')
        snap = b["btnSnapshot"]["actions"][0]["effects"][0]["values"]["ve_start"]["formula"]
        self.assertEqual(snap, "Max([Forecast Boundary/Start Month])")
        # the P&L detail groups by the union's Period, which the boundary now decides
        self.assertEqual(b["pnlPivot"]["columnsBy"][0]["columnId"], "p_per")

    def test_landed_actuals_map_every_period_so_a_closed_month_keeps_its_number(self):
        dm = self.build.build_dm_spec(self.plan, self.landed, self.plan_src, app=True)
        gl = next(e for e in dm["pages"][0]["elements"] if e["id"] == "glDetail")
        num = next(c for c in gl["columns"] if c["name"] == "Month Num")["formula"]
        self.assertIn('"Jun-27", 18', num)
        self.assertIn('"Jul-27", 19', num)
        self.assertIn('"Dec-27", 24', num)
        plain = self.build.build_dm_spec(self.plan, self.landed)
        gl = next(e for e in plain["pages"][0]["elements"] if e["id"] == "glDetail")
        self.assertNotIn('"Jul-27"', next(c for c in gl["columns"] if c["name"] == "Month Num")["formula"])

    def test_plain_forecast_keeps_the_static_split(self):
        plain = self.by_id(self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, landed=self.landed))
        self.assertNotIn("filters", plain["forecast"])
        self.assertNotIn("filters", plain["actuals"])
        self.assertNotIn("Parameters", json.dumps(plain["forecast"]))
        amount = next(c for c in plain["forecast"]["columns"] if c["name"] == "Amount")["formula"]
        self.assertIn("[June Base] * (1 + [Eff Rate]) ^ [N]", amount)

    def test_read_only_planning_app_also_follows_the_boundary(self):
        plan = copy.deepcopy(self.plan)
        plan["intent"] = "read-only"
        b = self.by_id(self.build.build_wb_spec(plan, "dm-1", self.els, landed=self.landed,
                                                workflows=self.manifest, plan_src=self.plan_src))
        fc = {c["name"]: c["formula"] for c in b["forecast"]["columns"]}
        self.assertEqual(fc["Start Key"], self.START_KEY)
        self.assertIn('Switch([Scenario], "Downside", [Forecast Spine/Downside Spread], '
                      '"Upside", [Forecast Spine/Upside Spread], 0)', fc["Eff Rate"])
        self.assertEqual(fc["Active Scenario"], 'Coalesce([Scenario], "Base")')
        self.assertEqual((b["ctlScenario"]["controlId"], b["ctlScenario"]["value"]), ("Scenario", "Base"))
        self.assertEqual(fc["N"], 'If([Kind] = "yoy", [Forecast Spine/N], [Month Key] - [Base Key])')
        self.assertIn("[Seed] * (1 + [Eff Rate]) ^ [N]", fc["Amount"])
        self.assertEqual(b["forecast"]["filters"][0]["values"], ["Forecast", "Awaiting actuals"])
        program, _seeded = self.build.setup_program(
            plan, self.build.workbook_wire.wire_workbook(
                self.build.build_wb_spec(plan, "dm-1", self.els, landed=self.landed,
                                         workflows=self.manifest, plan_src=self.plan_src)),
            self.manifest)
        self.assertNotIn("c_scn", program)

    # ---------------------------------------------------------------- the boundary, evaluated
    # The emitted data model + workbook run over local rows (tests/planning_app_eval.py): the landed
    # actuals (plan rows, plus a simulated Jul-27 month-end load), the plan table, the seeded input tables.
    JULY_UPLIFT = 1.0137

    def july_landed(self):
        """A simulated Jul-27 load: every line near its June actual (deterministic)."""
        june = {}
        for r in self.plan["actuals"]:
            if r["month"] == "Jun-27":
                june[r["line"]] = june.get(r["line"], 0.0) + float(r["amt"])
        return {line: round(amount * self.JULY_UPLIFT, 2) for line, amount in june.items()}

    def model(self, start, july=False, history=None):
        raw = [{"POSTING_DATE": date.fromisoformat(r["date"]), "MONTH": r["month"], "LINE_ITEM": r["line"],
                "REGION": "", "AMOUNT": float(r["amt"])} for r in self.plan["actuals"]]
        if july:
            raw += [{"POSTING_DATE": date(2027, 7, 1), "MONTH": "Jul-27", "LINE_ITEM": line, "REGION": "",
                     "AMOUNT": amount} for line, amount in self.july_landed().items()]
        with open(self.landing["tables"][1]["csv"], newline="") as f:
            plan_rows = [{**r, "PLAN_AMOUNT": float(r["PLAN_AMOUNT"])} for r in csv.DictReader(f)]
        rate, manual = self.build.seed_rows(self.plan)
        inputs = {"rateInput": [dict(zip(self.build.RATE_SEED_COLUMNS, r)) for r in rate],
                  "manualInput": [dict(zip(self.build.MANUAL_SEED_COLUMNS, r)) for r in manual],
                  "parameters": [] if start is None else [{"pa_start": start}]}
        for element, cols, rows in history or []:
            inputs[element] = [dict(zip(cols, r)) for r in rows]
        dm = self.build.build_dm_spec(self.plan, self.landed, self.plan_src, app=True)
        return pe.Model(self.spec(history=history), dm,
                        warehouse={"FPA_X_ACTUALS": raw, "FPA_X_PLAN": plan_rows},
                        inputs=inputs, controls={"Scenario": "Base"})

    @staticmethod
    def first_money(model, body):
        """The first ``$x.xM`` figure of a KPI card body, evaluated (the card's own emitted formula)."""
        expr = re.search(r"\{\{\((.*?)\) / 1000000 \|", body).group(1)
        return model.scalar(expr)

    def kpis(self, model):
        b = self.by_id(self.spec())
        return {"Revenue": self.first_money(model, b["kpiRevenueText"]["body"]),
                "Gross Margin": self.first_money(model, b["kpiMarginText"]["body"]),
                "EBITDA": self.first_money(model, b["kpiProfitText"]["body"])}

    @staticmethod
    def by_line_month(model, month_num):
        return {r["Line Item"]: r for r in model.rows("pnlCombined") if r["Month Num"] == month_num}

    @staticmethod
    def render(model, body):
        """A text element's body with each ``{{expr}}`` evaluated and the span markup dropped."""
        def value(m):
            v = model.scalar(m.group(1).split(" | ")[0])
            return "" if v is None else str(v)
        return re.sub(r"<[^>]+>", "", re.sub(r"\{\{(.*?)\}\}", value, body))

    def assert_ties_to_the_source(self, model):
        for name, value in self.kpis(model).items():
            self.assertAlmostEqual(value, self.plan["parityTargets"][name], places=2, msg=name)

    def test_before_roll_forward_fy_ties_to_the_source_even_with_july_landed(self):
        m = self.model(date(2027, 7, 1), july=True)
        self.assert_ties_to_the_source(m)
        july = self.by_line_month(m, 19)
        self.assertEqual({r["Period"] for r in july.values()}, {"Forecast"})    # the landed rows are ignored
        b = self.by_id(self.spec())
        self.assertIn("actuals through Jun-27 · live", self.render(m, b["hdrTitle"]["body"]))
        self.assertIn("6 of 12", self.render(m, b["kpiMonthsText"]["body"]))
        # Roll forward advances the one boundary by a month
        roll = b["btnRoll"]["actions"][0]["effects"][0]["values"]["pa_start"]["formula"]
        self.assertEqual(m.scalar(roll), date(2027, 8, 1))

    def test_gl_actuals_follow_the_boundary(self):
        # A landed month that is still a forecast month (Jul-27 before roll forward) is not shown as an
        # actual on P&L detail; after roll forward it is.
        gl = self.by_id(self.spec())["glTable"]
        window = next(c for c in gl["columns"] if c["name"] == "Window")
        self.assertIs(window.get("hidden"), True)
        self.assertIn("Max([Forecast Boundary/Start Key])", window["formula"])
        self.assertEqual(gl["filters"], [{"id": "g_wf", "columnId": window["id"], "kind": "list",
                                          "mode": "include", "values": ["Actual"]}])
        before = self.model(date(2027, 7, 1), july=True).rows("glTable")
        after = self.model(date(2027, 8, 1), july=True).rows("glTable")
        self.assertNotIn("Jul-27", {r["Month"] for r in before})
        self.assertEqual(len(after) - len(before), len(self.july_landed()))
        self.assertIn("Jul-27", {r["Month"] for r in after})

    def test_unseeded_parameters_fall_back_to_the_source_boundary(self):
        self.assert_ties_to_the_source(self.model(None))

    def test_after_roll_forward_july_reads_the_landed_actuals(self):
        pre = self.model(date(2027, 7, 1), july=True)
        m = self.model(date(2027, 8, 1), july=True)
        landed = self.july_landed()
        july = self.by_line_month(m, 19)
        self.assertEqual(set(july), set(landed))
        for line, row in july.items():
            self.assertEqual(row["Period"], "Actual", line)
            self.assertAlmostEqual(row["Amount"], landed[line], places=2, msg=line)
        aug, aug_pre = self.by_line_month(m, 20), self.by_line_month(pre, 20)
        lines = {ln["line"]: ln for ln in self.plan["lines"]}
        kinds = set()
        for line, row in aug.items():
            ln = lines[line]
            kinds.add(ln["kind"])
            if ln["kind"] == "growth":     # re-seeded off the July actual
                self.assertAlmostEqual(row["Amount"], landed[line] * (1 + ln["rate"]), places=2, msg=line)
            elif ln["kind"] == "flat":
                self.assertAlmostEqual(row["Amount"], landed[line], places=2, msg=line)
            elif ln["kind"] == "manual":   # typed values stay
                self.assertAlmostEqual(row["Amount"], aug_pre[line]["Amount"], places=2, msg=line)
        self.assertEqual(kinds, {"growth", "flat", "manual"})
        phase = {r["Month Num"]: r["Phase"] for r in m.rows("reportFacts")}
        self.assertEqual((phase[18], phase[19], phase[20]), ("Actual", "Actual", "Forecast"))
        b = self.by_id(self.spec())
        self.assertIn("actuals through Jul-27 · live", self.render(m, b["hdrTitle"]["body"]))
        self.assertIn("5 of 12", self.render(m, b["kpiMonthsText"]["body"]))
        # the July load (1.37% over June, not the forecast's growth) moves the FY off the source's figure
        self.assertGreater(abs(self.kpis(m)["Revenue"] - self.plan["parityTargets"]["Revenue"]), 1000)

    def test_rolled_month_without_landed_actuals_keeps_its_forecast_flagged(self):
        pre = self.model(date(2027, 7, 1))
        m = self.model(date(2027, 8, 1))           # rolled, but no Jul-27 load yet
        july, july_pre = self.by_line_month(m, 19), self.by_line_month(pre, 19)
        self.assertEqual(len(july), len(self.plan["lines"]))
        self.assertEqual(set(july), set(july_pre))   # never dropped
        for line, row in july.items():
            self.assertEqual(row["Period"], "Awaiting actuals", line)
            self.assertAlmostEqual(row["Amount"], july_pre[line]["Amount"], places=2, msg=line)
        for month in range(20, 25):                 # the rest re-forecast from the frozen July value: unchanged
            later, later_pre = self.by_line_month(m, month), self.by_line_month(pre, month)
            for line, row in later.items():
                self.assertAlmostEqual(row["Amount"], later_pre[line]["Amount"], places=2, msg=(month, line))
        self.assert_ties_to_the_source(m)
        phase = {r["Month Num"]: r["Phase"] for r in m.rows("reportFacts")}
        self.assertEqual((phase[18], phase[19]), ("Actual", "Forecast"))
        title = self.render(m, self.by_id(self.spec())["hdrTitle"]["body"])
        self.assertIn("actuals through Jun-27 · Jul-27 awaiting actuals · live on Snowflake", title)

    # ---------------------------------------------------------------- migrated history
    def test_source_logs_seed_versions_and_approvals(self):
        history = dw.source_history(self.plan["sourcePath"])
        seeds = {el: (cols, rows) for el, cols, rows in history}
        self.assertEqual(set(seeds), {"versions", "approvals"})
        vcols, vrows = seeds["versions"]
        self.assertEqual(vcols, ["ve_ver", "ve_scn", "ve_by", "ve_at", "ve_note", "ve_seq"])
        self.assertEqual(len(vrows), 7)
        self.assertEqual([r[0] for r in vrows], [f"v{i}" for i in range(1, 8)])
        self.assertEqual([r[-1] for r in vrows], list(range(1, 8)))
        self.assertRegex(vrows[0][3], r"^\d{4}-\d{2}-\d{2}$")
        acols, arows = seeds["approvals"]
        self.assertEqual(acols, ["ap_req", "ap_by", "ap_to", "ap_status", "ap_at", "ap_note", "ap_seq"])
        self.assertEqual(len(arows), 3)
        # the views and the status pill follow the entry order, not the source's (future) dates
        spec = self.spec(history=history)
        b = self.by_id(spec)
        for view in ("versionsView", "approvalsView"):
            order = next(c for c in b[view]["columns"] if c["name"] == "Log Order")
            self.assertIs(order.get("hidden"), True)
            self.assertEqual(b[view]["sort"], [{"columnId": order["id"], "direction": "descending"}])
            self.assertIn("Comment", [c["name"] for c in b[view]["columns"]])
        self.assertIn("migrated from the source workbook", b["logHead"]["body"])
        self.assertNotIn("start empty", b["logHead"]["body"])
        m = self.model(date(2027, 7, 1), history=history)
        self.assertEqual(m.rows("approvalStatus")[0]["Latest approval status"], arows[-1][3])
        # the next snapshot / request continue the numbering
        snap = b["btnSnapshot"]["actions"][0]["effects"][0]["values"]
        self.assertEqual(m.scalar(snap["ve_ver"]["formula"]), "v8")
        self.assertEqual(m.scalar(snap["ve_seq"]["formula"]), 8)
        sub = b["btnSubmit"]["actions"][0]["effects"][0]["values"]
        self.assertEqual(m.scalar(sub["ap_seq"]["formula"]), 4)
        # the setup program seeds them, last
        wired = self.build.workbook_wire.wire_workbook(spec)
        program = dw.emit_manual_commit_program(wired, {"start": "2027-07-01"}, history)
        self.assertIn("# Seed versions: 7 row(s)", program)
        self.assertIn("# Seed approvals: 3 row(s)", program)
        self.assertIn("c_ve_seq: 7}", program)
        self.assertIn("c_ve_seq = t.get_column('ve_seq')", program)
        self.assertNotIn("c_ve_rev", program)              # the source never froze the figures
        compile(program, "program", "exec")

    def test_cli_seeds_the_source_history_in_the_setup_program(self):
        import contextlib
        import io
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, landing_path, wf_path = self._cli_files(tmp)
            program_path = Path(tmp) / "setup.py"
            argv = [str(plan_path), "--actuals-table", "CSA.TJ.FPA_X_ACTUALS", "--landing", str(landing_path),
                    "--workflows", str(wf_path), "--emit-manual-commit-program", str(program_path),
                    "--out-dir", str(Path(tmp) / "out")]
            out = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main(argv), 0)
            program = program_path.read_text()
            wb = json.loads((Path(tmp) / "out" / "wb_spec.json").read_text())
        self.assertIn("HISTORY: versions 7 rows, approvals 3 rows", out.getvalue())
        self.assertIn("versions 7 rows, approvals 3 rows", out.getvalue().split("SETUP PROGRAM")[1])
        self.assertIn("get_element('versions')", program[program.index("# Seed versions"):])
        self.assertIn("migrated from the source workbook", json.dumps(wb))

    def test_detail_page_shows_the_fy_in_thousands(self):
        b = self.by_id(self.spec())
        pivot = b["pnlPivot"]
        self.assertEqual(pivot["filters"][0]["startDate"], "2027-01-01")
        self.assertTrue(next(c for c in pivot["columns"] if c["id"] == "p_amt")["formula"].endswith("/ 1000"))
        self.assertNotIn("Txn Date", [c["name"] for c in b["glTable"]["columns"]])
        # the cells carry the K, so the title does not repeat the unit; the GL title carries no row count
        # (a month close lands more rows than the source had, so a frozen count goes stale)
        self.assertEqual(pivot["name"], "Monthly P&L — FY27")
        self.assertEqual(b["glTable"]["name"], "GL actuals")

    def test_log_views_show_blanks_not_null(self):
        history = dw.source_history(self.plan["sourcePath"])
        b = self.by_id(self.spec(history=history))
        versions = {c["name"]: c for c in b["versionsView"]["columns"]}
        # figures the source never froze read as a dash, not the word "null"
        for name in ("FY27 revenue", "Gross margin", "EBITDA"):
            self.assertEqual(versions[name]["format"].get("displayNullAs"), "—", name)
            self.assertEqual(versions[name]["format"]["kind"], "number", name)
        start = versions["Forecast start"]
        self.assertEqual(start["formula"], 'Coalesce(DateFormat([Versions/Forecast Start Month], "%b-%y"), "—")')
        self.assertNotIn("format", start)
        for view in ("versionsView", "approvalsView"):
            cols = {c["name"]: c for c in b[view]["columns"]}
            self.assertTrue(cols["Comment"]["formula"].startswith("Coalesce("), view)
            self.assertTrue(cols["Comment"]["formula"].endswith(', "")'), view)
            stamp = cols["Saved at" if view == "versionsView" else "Submitted at"]["format"]["formatString"]
            self.assertNotIn("%I", stamp)                  # migrated rows carry a date, not a midnight time
            self.assertNotIn("%H", stamp)
        # every migrated row and the next few snapshots fit without an inner scroll
        layout = self.build.workbook_wire.wire_workbook(self.spec(history=history))["contents"]["layout"]
        def rows(element_id):
            m = re.search(r'elementId="%s"[^>]*gridRow="(\d+) / (\d+)"' % element_id, layout)
            return int(m.group(2)) - int(m.group(1))
        self.assertGreaterEqual(rows("versionsView"), 24)
        self.assertGreaterEqual(rows("approvalsView"), 13)

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

    def test_agent_reads_fiscal_year_facts_that_tie_to_the_kpis(self):
        # Live: grounded only in the forecast window, the agent reported its total as "FY27 revenue" and said
        # the workbook had no plan. It reads one clean monthly source: Actual + Forecast + Plan by fiscal year.
        spec = self.spec()
        self.assertEqual(self.owners(spec)["agentFacts"], "model")
        facts = self.by_id(spec)["agentFacts"]
        self.assertEqual([c["name"] for c in facts["columns"]],
                         ["Fiscal year", "Month", "Period", "Section", "Line item", "Region", "Amount",
                          "Active scenario"])
        fy = re.sub(r"\s+total$", "", self.plan["fyLabel"], flags=re.I)
        revenue = next(sec for sec in self.plan["sections"] if sec["sign"] > 0)["name"]
        def fy_sum(rows, period):
            return sum(r["Amount"] for r in rows if r["Fiscal year"] == fy and r["Section"] == revenue
                       and (r["Period"] != "Plan" if period == "fc" else r["Period"] == "Plan"))
        for start in (date(2027, 7, 1), date(2027, 8, 1)):     # before and after a roll forward
            m = self.model(start, july=True)
            rows = m.rows("agentFacts")
            self.assertEqual({r["Period"] for r in rows}, {"Actual", "Forecast", "Plan"})
            self.assertEqual({r["Fiscal year"] for r in rows if r["Period"] == "Plan"}, {fy})
            self.assertAlmostEqual(fy_sum(rows, "fc"), self.kpis(m)["Revenue"], places=2, msg=str(start))
            plan_rev = forecast_app.Facts(*[sorted({r["Month Num"] for r in m.rows("reportFacts")
                                                    if r["Period"] == "Plan"})] * 2).row({revenue: 1}, "plan")
            self.assertAlmostEqual(fy_sum(rows, "Plan"), m.scalar(plan_rev), places=2)
            self.assertEqual({r["Active scenario"] for r in rows}, {"Base"})
        pre = self.model(date(2027, 7, 1), july=True).rows("agentFacts")
        self.assertAlmostEqual(fy_sum(pre, "fc"), self.plan["parityTargets"]["Revenue"], places=2)
        self.assertGreater(fy_sum(pre, "fc"), sum(r["Amount"] for r in self.model(date(2027, 7, 1)).rows("forecast")
                                                  if r["Section"] == revenue))   # more than the forecast window
        # the agent is grounded in it, told how a fiscal year adds up, and offers the plan question
        wired = self.build.workbook_wire.wire_workbook(spec)
        macros = [{"name": w["macro"], "class": w["class"]} for w in self.manifest["workflows"]]
        _, frag = ae.recommend_for_workbook(macros, self.manifest, wired)
        agent = frag["agents"][0]
        self.assertIn({"kind": "table", "elementId": "agentFacts"}, agent["dataSources"])
        self.assertIn('@dataSource("agentFacts")', agent["instructions"])
        self.assertIn("sum of its Actual and Forecast rows", agent["instructions"])
        self.assertIn("plan is the sum of its Plan rows", agent["instructions"])
        self.assertIn("never report its total as a fiscal-year figure", agent["instructions"])
        self.assertIn("how does revenue compare with plan?", agent["greeting"]["message"])
        # without a landed plan the agent is not invited to answer a plan question it cannot
        bare = self.build.workbook_wire.wire_workbook(self.spec(plan_src=None))
        _, frag = ae.recommend_for_workbook(macros, self.manifest, bare)
        agent = frag["agents"][0]
        self.assertNotIn("plan", agent["greeting"]["message"])
        self.assertNotIn("Plan rows", agent["instructions"])
        self.assertIn({"kind": "table", "elementId": "agentFacts"}, agent["dataSources"])

    # ---------------------------------------------------------------- design round 2
    def test_input_tables_hide_row_id_and_fit_columns_to_content(self):
        # A number column stretched across a wide input table renders its edit pencil far left and its
        # right-aligned name far right, which reads as an empty unlabelled column. Autofit and the
        # presentation preset (live, d2) still stretch the columns, and the taller presentation rows clip
        # the tables: keep the dense grid, turn autofit off and size each column to its content.
        b = self.by_id(self.spec())
        for tid in ("rateInput", "manualInput"):
            cols = {c["id"]: c for c in b[tid]["columns"]}
            self.assertIs(cols["ID"].get("hidden"), True, tid)
            self.assertNotIn("ID", b[tid]["order"])
            style = b[tid]["tableStyle"]
            self.assertEqual((style["preset"], style["gridLines"], style["autofitColumns"]),
                             ("spreadsheet", "horizontal", "hidden"), tid)
        wired = self.build.workbook_wire.wire_workbook(self.spec())
        rate_rows = [["North America Cost of Revenue", 0.008], ["EMEA", 0.01]]   # seed_rows shape
        widths = dw.input_column_widths(wired, [("rateInput", ["LINE_ITEM", "RATE"], rate_rows)])
        # every visible typed column gets a width; hidden ones and the system ID do not
        self.assertEqual(set(widths["rateInput"]), {"LINE_ITEM", "RATE"})
        self.assertEqual(set(widths["manualInput"]), {"LINE_ITEM", "MONTH", "AMOUNT"})
        # the longest value (or the header plus its edit pencil) sets the width, within bounds
        self.assertGreater(widths["rateInput"]["LINE_ITEM"], widths["manualInput"]["LINE_ITEM"])
        for table in widths.values():
            for px in table.values():
                self.assertTrue(dw.MIN_COLUMN_PX <= px <= dw.MAX_COLUMN_PX, px)
        # the setup program re-applies the style (preset first: switching preset resets the rest), then the widths
        program = dw.emit_manual_commit_program(wired, {"start": "2027-07-01"}, column_widths=widths)
        block = program[program.index("get_element('rateInput')"):program.index("get_element('manualInput')")]
        self.assertIn("s = t.table_style()", block)
        self.assertLess(block.index("s.set_preset('spreadsheet')"), block.index("s.set_grid_lines('horizontal')"))
        self.assertIn("s.set_autofit_columns(False)", block)
        self.assertIn(f"t.set_column_width(t.get_column('RATE'), {widths['rateInput']['RATE']})", block)
        self.assertLess(block.index("set_autofit_columns"), block.index("set_column_width"))
        compile(program, "program", "exec")
        # Assumptions: the narrow tables sit side by side with the impact card, no separate manual heading
        layout = wired["contents"]["layout"]
        def cols(element_id):
            m = re.search(r'elementId="%s"[^>]*gridColumn="(\d+) / (\d+)"[^>]*gridRow="(\d+) / (\d+)"'
                          % element_id, layout)
            return tuple(int(g) for g in m.groups())
        rate, manual, impact = cols("rateInput"), cols("manualInput"), cols("asmImpactCard")
        self.assertEqual((rate[1], manual[1]), (manual[0], impact[0]))
        self.assertEqual(impact[1], 25)
        self.assertEqual({rate[2:], manual[2:], impact[2:]}, {rate[2:]})
        self.assertNotIn('elementId="manualHead"', layout)
        plain = self.build.build_wb_spec(copy.deepcopy(self.plan), "dm-1", self.els, landed=self.landed)
        self.assertNotIn("tableStyle", next(e for pg in plain["pages"] for e in pg["elements"]
                                            if e["id"] == "rateInput"))

    def test_header_actions_are_one_compact_right_aligned_group(self):
        spec = self.build.workbook_wire.wire_workbook(self.spec())
        layout = spec["contents"]["layout"]
        def cols(element_id):
            m = re.search(r'elementId="%s"[^>]*gridColumn="(\d+) / (\d+)"' % element_id, layout)
            return int(m.group(1)), int(m.group(2))
        band = re.search(r'<Container elementId="hdrBand"[^>]*>(.*?)</Container>', layout, re.S).group(1)
        self.assertNotIn("<Container", band)   # no nested container: the server's layout check misplaces it
        self.assertEqual(re.findall(r'elementId="(btn\w+)"', band), ["btnRoll", "btnSnapshot", "btnSubmit"])
        spans = [cols(b) for b in ("btnRoll", "btnSnapshot", "btnSubmit")]
        self.assertEqual([s[0] for s in spans[1:]], [s[1] for s in spans[:-1]])   # adjacent, no empty cells
        self.assertEqual(spans[-1][1], 25)                                         # primary flush right
        self.assertLessEqual(spans[-1][1] - spans[0][0], 8)                        # compact: a third of the band
        self.assertGreater(spans[-1][1] - spans[-1][0], spans[0][1] - spans[0][0])  # widest label gets most room
        self.assertEqual(cols("hdrStatus")[0], cols("ctlScenario")[1])            # status next to the switch
        self.assertLessEqual(cols("hdrStatus")[1], spans[0][0])
        b = self.by_id(self.spec())
        self.assertEqual({b[i]["align"] for i in ("btnRoll", "btnSnapshot", "btnSubmit")}, {"stretch"})

    def test_pl_summary_variance_coloured_by_favourability(self):
        s = self.by_id(self.spec())["plSummary"]
        rules = {r["style"]["color"]: r for r in s["conditionalFormats"]}
        self.assertEqual(set(rules), {forecast_app.GOOD, forecast_app.BAD})
        for rule in rules.values():
            self.assertEqual((rule["type"], rule["condition"], rule["columnIds"]), ("single", "formula", ["ps_vs"]))
        bad = rules[forecast_app.BAD]["formula"]
        # a cost line under plan is favourable: its variance sign is flipped before the test
        self.assertEqual(bad, 'If([Line] = "Operating expenses", -1, 1) * [Vs plan] < 0')
        self.assertEqual(rules[forecast_app.GOOD]["formula"], bad[:-3] + "> 0")
        self.assertEqual(self.build.cost_lines(self.plan), ["Operating expenses"])

    def test_plain_dm_leaves_a_landed_plan_table_out(self):
        dm = self.build.build_dm_spec(self.plan, self.landed, self.plan_src)   # plan resolved, but no app
        self.assertNotIn("planDetail", {e["id"] for e in dm["pages"][0]["elements"]})

    def _cli_files(self, tmp, bad_plan=False):
        tmp = Path(tmp)
        plan_path = tmp / "plan.json"
        plan_path.write_text(json.dumps(self.plan), encoding="utf-8")
        landing = copy.deepcopy(self.landing)
        if bad_plan:   # a plan month the forecast does not have: an app run must refuse it
            bad = tmp / "bad-plan.csv"
            bad.write_text("MONTH,LINE_ITEM,REGION,PLAN_AMOUNT\nJan-30,LATAM Opex,LATAM,1\n")
            landing["tables"][1]["csv"] = str(bad)
        landing_path = tmp / "landing.json"
        landing_path.write_text(json.dumps(landing), encoding="utf-8")
        wf_path = tmp / "workflows.json"
        wf_path.write_text(json.dumps(self.manifest), encoding="utf-8")
        return plan_path, landing_path, wf_path

    def test_plain_cli_run_never_reads_the_landed_plan_table(self):
        import contextlib
        import io
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, landing_path, wf_path = self._cli_files(tmp, bad_plan=True)
            base = [str(plan_path), "--actuals-table", "CSA.TJ.FPA_X_ACTUALS", "--landing", str(landing_path),
                    "--out-dir", str(Path(tmp) / "out")]
            out = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main(base), 0)
            self.assertNotIn("PLAN SOURCE", out.getvalue())
            dm = json.loads((Path(tmp) / "out" / "dm_spec.json").read_text())
            self.assertNotIn("planDetail", {e["id"] for e in dm["pages"][0]["elements"]})
            err = io.StringIO()
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
                self.assertEqual(self.build.main(base + ["--workflows", str(wf_path)]), 1)
            self.assertIn("plan table month 'Jan-30'", err.getvalue())

    def test_cli_setup_program_sizes_the_input_table_columns_from_the_seeds(self):
        import contextlib
        import io
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, landing_path, wf_path = self._cli_files(tmp)
            program_path = Path(tmp) / "setup.py"
            argv = [str(plan_path), "--actuals-table", "CSA.TJ.FPA_X_ACTUALS", "--landing", str(landing_path),
                    "--workflows", str(wf_path), "--emit-manual-commit-program", str(program_path),
                    "--out-dir", str(Path(tmp) / "out")]
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main(argv), 0)
            program = program_path.read_text()
        for tid, col in (("rateInput", "LINE_ITEM"), ("rateInput", "RATE"), ("manualInput", "AMOUNT")):
            block = program[program.index(f"get_element('{tid}')"):]
            self.assertRegex(block[:block.index("set_edit_mode('draft')")],
                             r"t\.set_column_width\(t\.get_column\('%s'\), \d+\)" % col)

    def test_cli_setup_program_binds_against_the_recorded_sdk_and_seeds_last(self):
        import ast
        import contextlib
        import io
        import code_mode_check
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, landing_path, wf_path = self._cli_files(tmp)
            program_path = Path(tmp) / "setup.py"
            argv = [str(plan_path), "--actuals-table", "CSA.TJ.FPA_X_ACTUALS", "--landing", str(landing_path),
                    "--workflows", str(wf_path), "--emit-manual-commit-program", str(program_path),
                    "--out-dir", str(Path(tmp) / "out")]
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main(argv), 0)
            program = program_path.read_text()
        errors, calls = code_mode_check.check_program(program)
        self.assertEqual(errors, [])
        self.assertGreater(calls, 50)
        methods = [n.func.attr for n in ast.walk(ast.parse(program))
                   if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)]
        for used in ("table_style", "set_preset", "set_grid_lines", "set_autofit_columns", "set_column_width",
                     "set_series_color", "upsert_rows"):
            self.assertIn(used, methods)
        # the seeds are the only writes that cannot be repeated: every style/format/mode call precedes them
        lines = program.splitlines()
        first_write = next(i for i, ln in enumerate(lines) if ".upsert_rows(" in ln)
        for marker in ("set_series_color", "set_edit_mode", "set_column_width", "set_format_pattern", "set_preset"):
            self.assertLess(max(i for i, ln in enumerate(lines) if marker in ln), first_write, marker)

    def test_setup_program_with_a_wrong_sdk_call_is_refused_at_emit(self):
        wired = self.build.workbook_wire.wire_workbook(self.spec())
        el = next(e for e in code_rep.workbook_elements(code_rep.document(wired)) if e["id"] == "rateInput")
        el["tableStyle"]["gridLines"] = "rows"      # not a TableGridLines value
        with self.assertRaisesRegex(dw.WorkflowError, "recorded Sigma SDK signatures.*'rows' is not one of"):
            dw.emit_manual_commit_program(wired, {"start": "2027-07-01"})

    def test_post_preflights_the_agent_merge_before_creating_anything(self):
        import contextlib
        import io
        from unittest import mock
        from migration_gate import gate_plan_staging_source
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, landing_path, wf_path = self._cli_files(tmp)
            _src, report = gate_plan_staging_source(self.plan, plan_path=str(plan_path))
            approval = Path(tmp) / "approval.json"
            approval.write_text(json.dumps({
                "owner": "Maria Lopez", "reviewer": "Jane Reviewer", "folderId": self.plan["folderId"],
                "sourceSha256": report["sourceSha256"],
                "sourceReviewApproval": {"status": "approved_for_staging", "issues": report["issues"],
                                         "confirmation": self.build.TERMINAL_CONFIRMATION}}))
            macros = Path(tmp) / "macros.json"
            macros.write_text(json.dumps([{"name": w["macro"], "class": w["class"]}
                                          for w in self.manifest["workflows"]]))
            argv = [str(plan_path), "--post", "--staging-approval", str(approval), "--actuals-table",
                    "CSA.TJ.FPA_X_ACTUALS", "--landing", str(landing_path), "--workflows", str(wf_path),
                    "--agent-macros", str(macros), "--out-dir", str(Path(tmp) / "out")]
            err = io.StringIO()
            with mock.patch.object(self.build, "api") as api, \
                    mock.patch.object(self.build, "env", return_value={"SIGMA_BASE_URL": "https://x",
                                                                       "SIGMA_API_TOKEN": "t",
                                                                       "SIGMA_AUTH_METHOD": "test"}), \
                    mock.patch.object(ae, "merge_fragment",
                                      side_effect=ValueError("workbook already has an element chat")), \
                    contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
                self.assertEqual(self.build.main(argv), 1)
            api.assert_not_called()
            self.assertIn("AGENT ERROR", err.getvalue())

    # ---------------------------------------------------------------- staging
    def test_staging_approval_must_cover_exact_issues_folder_and_names(self):
        issue = {"code": "VBA_PROJECT", "severity": "review", "message": "m"}
        report = {"sourceSha256": "a" * 64, "issues": [issue]}
        P = {"folderId": "fold-1"}
        confirmed = self.build.TERMINAL_CONFIRMATION
        good = {"owner": "Maria Lopez", "reviewer": "Jane Reviewer", "folderId": "fold-1", "sourceSha256": "a" * 64,
                "sourceReviewApproval": {"status": "approved_for_staging", "issues": [issue],
                                         "confirmation": confirmed}}
        self.build.check_staging_approval(P, report, good)
        # a hand-written "approved" JSON without the reviewer's terminal confirmation is refused
        for review in ({"status": "approved_for_staging", "issues": [issue]},
                       {"status": "approved_for_staging", "issues": [issue], "confirmation": "approved"}):
            with self.assertRaisesRegex(self.build.StagingError, "confirmation"):
                self.build.check_staging_approval(P, report, {**good, "sourceReviewApproval": review})
        for change in ({"owner": ""}, {"folderId": "other"}, {"sourceSha256": "b" * 64},
                       {"sourceReviewApproval": {"status": "pending", "issues": [issue]}},
                       {"sourceReviewApproval": {"status": "approved_for_staging", "issues": []}},
                       {"sourceReviewApproval": {"status": "approved_for_staging",
                                                 "issues": [{**issue, "message": "edited"}],
                                                 "confirmation": confirmed}}):
            with self.assertRaises(self.build.StagingError, msg=str(change)):
                self.build.check_staging_approval(P, report, {**good, **change})
        template = self.build.staging_template(P, report)
        self.assertEqual((template["owner"], template["sourceReviewApproval"]["status"]), ("", "pending"))

    def test_staging_template_records_the_named_approval_from_flags(self):
        # The owner/reviewer approval is given on the command line by the reviewer at their own terminal
        # (they type APPROVE), never by hand-editing JSON; the exact audited issues are kept and nothing
        # is marked verified.
        import contextlib
        import io
        from unittest import mock
        from migration_gate import gate_plan_staging_source
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, _landing, _wf = self._cli_files(tmp)
            _src, report = gate_plan_staging_source(self.plan, plan_path=str(plan_path))
            out_json = Path(tmp) / "staging-approval.json"
            base = [str(plan_path), "--staging-approval-template", str(out_json)]
            names = ["--approve-for-staging", "--owner", "TJ Wells", "--reviewer", "Jane Reviewer"]
            err, out = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                self.assertEqual(self.build.main(base + ["--approve-for-staging", "--owner", "TJ Wells"]), 1)
            self.assertIn("--owner and --reviewer", err.getvalue())
            self.assertFalse(out_json.exists())
            # the reviewer types APPROVE at an interactive terminal: approved, issues listed to them first
            prompt = io.StringIO()
            with mock.patch.object(self.build.sys, "stdin", TtyInput("APPROVE\n")), \
                    contextlib.redirect_stdout(out), contextlib.redirect_stderr(prompt):
                self.assertEqual(self.build.main(base + names), 0)
            for issue in report["issues"]:
                self.assertIn(issue["code"], prompt.getvalue())
            self.assertIn("type APPROVE", prompt.getvalue())
            approval = json.loads(out_json.read_text())
            self.assertEqual((approval["owner"], approval["reviewer"]), ("TJ Wells", "Jane Reviewer"))
            self.assertEqual(approval["sourceReviewApproval"]["status"], "approved_for_staging")
            self.assertEqual(approval["sourceReviewApproval"]["issues"], report["issues"])
            self.assertEqual(approval["sourceReviewApproval"]["confirmation"], self.build.TERMINAL_CONFIRMATION)
            self.assertNotIn('"verified": true', out_json.read_text())
            self.build.check_staging_approval(self.plan, report, approval)
            self.assertIn("APPROVED FOR STAGING", out.getvalue())
            # without --approve-for-staging the names are recorded but the status stays pending
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main(base + ["--owner", "TJ Wells"]), 0)
            pending = json.loads(out_json.read_text())
            self.assertEqual((pending["owner"], pending["sourceReviewApproval"]["status"]), ("TJ Wells", "pending"))
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.build.main([str(plan_path), "--approve-for-staging", "--owner", "a",
                                                  "--reviewer", "b"]), 1)   # only with the template

    def test_agent_cannot_self_approve_staging(self):
        # An agent's shell (or a pipe / CI) has no terminal: the copy-pasted approve command refuses and
        # writes nothing; at a terminal anything but APPROVE also refuses.
        import contextlib
        import io
        from unittest import mock
        with tempfile.TemporaryDirectory() as tmp:
            plan_path, _landing, _wf = self._cli_files(tmp)
            out_json = Path(tmp) / "staging-approval.json"
            argv = [str(plan_path), "--staging-approval-template", str(out_json), "--approve-for-staging",
                    "--owner", "Maria Lopez", "--reviewer", "Jane Reviewer"]
            for stdin, why in ((PipeInput("APPROVE\n"), "interactive terminal"),
                               (TtyInput("yes\n"), "not approved"),
                               (TtyInput(""), "not approved")):
                err = io.StringIO()
                with mock.patch.object(self.build.sys, "stdin", stdin), \
                        contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
                    self.assertEqual(self.build.main(argv), 1)
                self.assertIn(why, err.getvalue())
                self.assertFalse(out_json.exists())


class PipeInput(io.StringIO):
    """stdin as an agent's shell or a pipe sees it: readable, not a terminal."""
    def isatty(self):
        return False


class TtyInput(io.StringIO):
    """stdin at a person's interactive terminal."""
    def isatty(self):
        return True

def forecast_app_disallowed(body):
    import presentation_style
    return presentation_style.disallowed_span_styles(body)


if __name__ == "__main__":
    unittest.main()
