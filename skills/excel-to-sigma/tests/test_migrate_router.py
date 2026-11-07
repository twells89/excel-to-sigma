#!/usr/bin/env python3
"""scripts/migrate.py plan: archetype routing, BLOCKED exit code, review template."""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import forecast_pnl_fixture as fixture  # noqa: E402
import table_report_fixture as table_fixture  # noqa: E402
from migration_gate import source_sha256  # noqa: E402

spec = importlib.util.spec_from_file_location("migrate_router", SCRIPTS / "migrate.py")
MIGRATE = importlib.util.module_from_spec(spec)
spec.loader.exec_module(MIGRATE)


class TestRouter(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name).resolve()
        self.out = self.dir / "work"

    def plan(self, path, *extra):
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(io.StringIO()):
            code = MIGRATE.main(["plan", str(path), "--out-dir", str(self.out), *extra])
        return code, (json.loads(stdout.getvalue()) if stdout.getvalue() else None)

    def table_workbook(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Entries"
        ws.append(["Account", "Amount"])
        for index in range(1, 21):
            ws.append([f"A{index}", index * 10])
        ws.add_table(Table(displayName="tblEntries", ref="A1:B21"))
        path = self.dir / "table.xlsx"
        wb.save(path)
        return path

    def macro_enabled_forecast(self, xlsx=None):
        """A forecast source as a real .xlsm: the 08 fixture's vbaProject.bin (one recorded
        macro) injected with the macro-enabled workbook content type and its relationship.
        Default source: the P&L fixture (no Excel Tables)."""
        if xlsx is None:
            xlsx = self.dir / "pnl.xlsx"
            fixture.build_pnl(xlsx)
        with zipfile.ZipFile(ROOT / "tests" / "fixtures" / "broader-eight"
                             / "08 Inventory Replenishment Plan.xlsm") as donor:
            vba = donor.read("xl/vbaProject.bin")
        xlsm = xlsx.with_suffix(".xlsm")
        with zipfile.ZipFile(xlsx) as src, zipfile.ZipFile(xlsm, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    text = data.decode("utf-8").replace(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
                    assert "macroEnabled.main" in text
                    data = text.replace("</Types>", '<Default Extension="bin" '
                                        'ContentType="application/vnd.ms-office.vbaProject"/></Types>').encode()
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    data = data.decode("utf-8").replace("</Relationships>", '<Relationship Id="rIdVBA" '
                        'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                        'Target="vbaProject.bin"/></Relationships>').encode()
                dst.writestr(item, data)
            dst.writestr("xl/vbaProject.bin", vba)
        return xlsx, xlsm

    def demo_source(self):
        """The demo forecast (formal Tables tblActuals, tblPlan, ...) as .xlsx and .xlsm."""
        spec = importlib.util.spec_from_file_location("make_demo_router", SCRIPTS / "make-demo-forecast.py")
        maker = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(maker)
        return self.macro_enabled_forecast(Path(maker.build_demo_workbook(self.dir)))

    @staticmethod
    def enhance_absent(test, result):
        text = "\n".join(result["nextCommands"])
        test.assertNotIn("recommend-agent.py", text)
        test.assertNotIn("demo-manifest", text)
        test.assertNotIn("--workflows", text)
        test.assertNotIn("refs/agent-enhance.md", result["readRefs"])
        test.assertFalse(any("--workflows" in h for h in result["userHandoffs"]))

    def test_forecast_route_offers_landing_and_enhance(self):
        try:
            import oletools  # noqa: F401
        except ImportError:
            self.skipTest("oletools is not installed")
        xlsx, xlsm = self.macro_enabled_forecast()
        # No Excel Tables: nothing to land, so no landing step and no landed-actuals flags.
        _, plain = self.plan(xlsx)
        text = "\n".join(plain["nextCommands"])
        self.assertNotIn("land-in-snowflake.py", text)
        self.assertNotIn("--actuals-table", text)
        self.enhance_absent(self, plain)
        # Real macro-enabled forecast source, real macro-classify (no mock): Enhance (Phase 6)
        # is offered after the build.
        code, enabled = self.plan(xlsm)
        self.assertEqual(code, 0)
        self.assertEqual(enabled["archetype"], "forecast")
        self.assertIsInstance(enabled["macros"], list)
        self.assertEqual([(m["name"], m["class"], m["status"]) for m in enabled["macros"]],
                         [("Macro1", "NAVIGATION", "AUTO")])
        self.assertEqual(json.loads((self.out / "macros.json").read_text()), enabled["macros"])
        cmds = enabled["nextCommands"]
        self.assertIn("macro-classify.py", cmds[0])
        manifest = next(c for c in cmds if "demo-manifest" in c)
        self.assertTrue(manifest.split("--out ")[1].split()[0].endswith("workflows.json"))
        self.assertLess(next(i for i, c in enumerate(cmds) if "demo-manifest" in c),
                        next(i for i, c in enumerate(cmds) if "recommend-agent.py" in c))
        self.assertIn("refs/agent-enhance.md", enabled["readRefs"])

    def test_forecast_with_tables_and_macros_lists_the_run_in_order(self):
        # The demo shape: Excel Tables + macros. nextCommands is the order the run takes; every build
        # carries the landed actuals, the workflows and one setup-program name; the agent rides the post.
        from unittest import mock
        real = MIGRATE.run_json
        demo_macros = json.loads((ROOT / "tests" / "fixtures" / "agent-enhance" / "demo-macros.json").read_text())

        def fake(*args):
            return demo_macros if Path(args[0]).name == "macro-classify.py" else real(*args)
        _, xlsm = self.demo_source()
        with mock.patch.object(MIGRATE, "run_json", fake):
            code, result = self.plan(xlsm)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "forecast")
        cmds = result["nextCommands"]

        def kind(c):
            for key, label in (("macro-classify.py", "classify"), ("land-in-snowflake.py --", None),
                               ("draft-forecast-plan.py", "draft"), ("demo-manifest", "manifest"),
                               ("recommend-agent.py", "pitch")):
                if key in c and label:
                    return label
            if "land-in-snowflake.py" in c:
                return "land-execute" if "--execute" in c else "land-dry"
            if "build-forecast-model.py" in c:
                return "post" if "--post" in c else "build-dry"
            return c
        self.assertEqual([kind(c) for c in cmds], ["classify", "land-dry", "land-execute", "draft", "manifest",
                                                   "build-dry", "pitch", "post"])
        program = str(self.out / "setup-program.py")
        everything = json.dumps(result)
        self.assertNotIn("manual-commit.py", everything)
        self.assertNotIn("--merge-into", "\n".join(cmds))
        self.assertIn("tblActuals, tblPlan", next(c for c in cmds if kind(c) == "land-dry"))
        builds = [c for c in cmds if "build-forecast-model.py" in c] + result["stagingAlternative"]
        self.assertEqual(len(builds), 5)
        for c in builds:
            for flag in ("--actuals-table", "--landing", "--workflows " + str(self.out / "workflows.json"),
                         "--emit-manual-commit-program " + program):
                self.assertIn(flag, c)
        post = cmds[-1]
        self.assertIn("--agent-macros " + str(self.out / "macros.json"), post)
        self.assertIn("--post --review-dispositions", post)
        self.assertNotIn("--agent-macros", cmds[cmds.index(post) - 2])   # the dry run precedes the yes
        template, approve, staged = result["stagingAlternative"]
        # the agent's template step carries no approval; approving is the reviewer's own terminal step
        self.assertIn("--staging-approval-template", template)
        self.assertNotIn("--approve-for-staging", template)
        self.assertNotIn("--owner", template)
        self.assertIn("--approve-for-staging", approve)
        self.assertIn("--owner", approve)
        self.assertIn("USER HANDOFF", approve)
        self.assertIn("own terminal", approve)
        self.assertIn("--post --staging-approval", staged)
        self.assertIn("--agent-macros", staged)
        setup = [h for h in result["userHandoffs"] if "setup program" in h]
        self.assertEqual(len(setup), 1)
        self.assertIn(program, setup[0])
        self.assertTrue(any("Smoke pass" in h for h in result["userHandoffs"]))
        gate = json.loads((self.out / "macro-gate.json").read_text())
        self.assertEqual(result["macroGate"], str(self.out / "macro-gate.json"))
        self.assertEqual((gate["procedures"], gate["dropped"]), (7, []))
        self.assertEqual(len(gate["nonAuto"]), 6)
        self.assertFalse(any(r["verified"] for r in gate["records"]))
        self.assertTrue(any("resolve-targets" in d for d in result["askUser"]))

    def test_macro_enabled_non_forecast_source_gets_no_enhance(self):
        try:
            import oletools  # noqa: F401
        except ImportError:
            self.skipTest("oletools is not installed")
        # 08 Inventory Replenishment Plan.xlsm: real macros, table-report archetype.
        _, result = self.plan(ROOT / "tests" / "fixtures" / "broader-eight"
                              / "08 Inventory Replenishment Plan.xlsm")
        self.assertEqual(result["archetype"], "table-report")
        self.assertEqual([m["status"] for m in result["macros"]], ["AUTO"])
        self.enhance_absent(self, result)
        self.assertNotIn("land-in-snowflake.py", "\n".join(result["nextCommands"]))
        self.assertNotIn("--emit-manual-commit-program", "\n".join(result["nextCommands"]))
        self.assertFalse((self.out / "macros.json").exists())
        # Non-forecast table workbooks never get the forecast-only landing/Enhance commands.
        _, table = self.plan(self.table_workbook())
        self.enhance_absent(self, table)
        self.assertNotIn("--emit-manual-commit-program", "\n".join(table["nextCommands"]))

    def test_macro_classify_without_a_procedure_list_gets_no_enhance(self):
        from unittest import mock
        _, xlsm = self.macro_enabled_forecast()
        real = MIGRATE.run_json

        def raises(*args):
            raise RuntimeError("olevba exploded")

        for label, outcome in (("empty list", []), ("error dict", {"error": "no VBA"}),
                               ("raised", raises)):
            def fake(*args, outcome=outcome):
                if Path(args[0]).name == "macro-classify.py":
                    return outcome(*args) if callable(outcome) else outcome
                return real(*args)
            with self.subTest(label), mock.patch.object(MIGRATE, "run_json", fake):
                code, result = self.plan(xlsm)
                self.assertEqual(code, 0)
                self.assertEqual(result["archetype"], "forecast")
                self.enhance_absent(self, result)
                self.assertIn("build-forecast-model.py", "\n".join(result["nextCommands"]))
                self.assertFalse((self.out / "macros.json").exists())
        self.assertIn("macro-classify failed (olevba exploded)", " ".join(result["notes"]))

    def test_forecast_route_seeds_through_the_setup_program_not_a_paste(self):
        path = self.dir / "pnl.xlsx"
        fixture.build_pnl(path)
        _, result = self.plan(path)
        handoffs = result["userHandoffs"]
        self.assertFalse(any("paste" in h.lower() for h in handoffs), handoffs)
        program = str(self.out / "setup-program.py")
        post = next(c for c in result["nextCommands"] if "--post" in c)
        self.assertIn("--emit-manual-commit-program " + program, post)
        setup = [h for h in handoffs if "setup program" in h]
        self.assertEqual(len(setup), 1, handoffs)
        self.assertIn("Sigma MCP Code Mode build_workbook", setup[0])
        self.assertIn(program, setup[0])
        self.assertIn("once", setup[0])
        self.assertIn("rate", setup[0])
        self.assertNotIn("Parameters", setup[0])

    def test_enhanced_forecast_has_one_setup_handoff_that_also_seeds_parameters(self):
        try:
            import oletools  # noqa: F401
        except ImportError:
            self.skipTest("oletools is not installed")
        _, xlsm = self.macro_enabled_forecast()
        _, result = self.plan(xlsm)
        handoffs = result["userHandoffs"]
        self.assertFalse(any("paste" in h.lower() for h in handoffs), handoffs)
        setup = [h for h in handoffs if "setup program" in h]
        self.assertEqual(len(setup), 1, handoffs)
        self.assertIn("--workflows", setup[0])
        self.assertIn("Parameters", setup[0])
        self.assertIn("manual commit", setup[0])

    def test_phase6_docs_name_the_demo_manifest_step_and_no_plan_tasks(self):
        import re
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        phase6 = skill[skill.index("## Phase 6"):skill.index("## What this skill does NOT do")]
        self.assertRegex(phase6, r"macro-actions\.py demo-manifest \S+ --out \S*workflows\.json")
        self.assertNotIn("Paste seed", skill)
        for name in ("agent-enhance.md", "input-tables.md", "forecast-recipes.md", "sigma-build-gotchas.md"):
            text = (ROOT / "refs" / name).read_text(encoding="utf-8")
            self.assertIsNone(re.search(r"\(Task \d+\)", text), name)
        self.assertIsNone(re.search(r"\(Task \d+\)", phase6))

    def test_forecast_archetype_routes_to_drafter_then_builder(self):
        path = self.dir / "pnl.xlsx"
        fixture.build_pnl(path)
        code, result = self.plan(path)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "forecast")
        self.assertEqual(result["auditStatus"], "AUTO")
        self.assertIn("refs/forecast-recipes.md", result["readRefs"])
        self.assertEqual(result["beforePost"], ["refs/coverage-contract.md"])
        self.assertNotIn("refs/coverage-contract.md", result["readRefs"])
        commands = result["nextCommands"]
        self.assertIn("draft-forecast-plan.py", commands[0])
        self.assertIn(str(self.out / "plan.json"), commands[0])
        self.assertIn("build-forecast-model.py", commands[1])
        self.assertNotIn("--post", commands[1])
        self.assertIn("--post --review-dispositions", commands[2])
        self.assertEqual(result["modelMap"]["P&L"]["ACTUAL"], 18)
        self.assertTrue((self.out / "migration-audit.json").is_file())
        self.assertTrue((self.out / "discovery.json").is_file())
        self.assertIsNone(result["reviewTemplate"])  # AUTO -> nothing to disposition
        self.assertFalse((self.out / "review.template.json").exists())

    def test_sumifs_workbook_the_drafter_rejects_routes_to_cell_model(self):
        # SUMIFS actuals but a forecast formula the drafter cannot read: the route must not
        # name the drafter.
        path = self.dir / "pnl.xlsx"

        def odd_forecast(pnl, gl, drivers):
            pnl["E5"] = "=D5*1.05"
        fixture.build_pnl(path, mutate=odd_forecast)
        code, result = self.plan(path)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "cell-model")
        self.assertIn("forecast drafter rejects the shape", result["routeReason"])
        self.assertIn("unrecognised forecast cell '=D5*1.05'", result["routeReason"])
        self.assertFalse(any("draft-forecast-plan.py" in c for c in result["nextCommands"]))
        self.assertIn("source-workbook.py", result["nextCommands"][0])
        self.assertIn("refs/source-workbook-model.md", result["readRefs"])

    def test_quarterly_budget_with_discovered_sections_routes_to_the_drafter(self):
        # 03 Retail Quarterly Budget: Net Sales / COGS / SG&A / Other sections, quarter
        # headers, year-over-year and percent-of-sales forecasts, an external-link cell.
        source = (Path(__file__).resolve().parent / "fixtures" / "broader-eight"
                  / "03 Retail Quarterly Budget.xlsx")
        code, result = self.plan(source)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "forecast")
        self.assertIn("forecast drafter accepts this workbook: 10 lines", result["routeReason"])
        self.assertIn("draft-forecast-plan.py", result["nextCommands"][0])

    def test_accepted_forecast_reports_the_drafter_shape(self):
        path = self.dir / "pnl.xlsx"
        fixture.build_pnl(path)
        _, result = self.plan(path)
        self.assertIn("forecast drafter accepts this workbook: 6 lines", result["routeReason"])

    def test_input_table_archetype(self):
        code, result = self.plan(self.table_workbook())
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "input-table")
        self.assertEqual(result["tables"], [{"name": "tblEntries", "sheet": "Entries", "rows": 20}])
        self.assertIn("refs/input-tables.md", result["readRefs"])
        self.assertIn("xlsx-to-input-csv.py", result["nextCommands"][0])

    def test_formal_tables_with_formula_grids_route_to_the_table_report_drafter(self):
        path = self.dir / "staff.xlsx"
        table_fixture.build_staff(path)
        code, result = self.plan(path)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "table-report")
        self.assertIn("table-report drafter accepts this workbook: 1 table(s), 2 frames, 28 source cells verified",
                      result["routeReason"])
        self.assertIn("refs/table-report-recipes.md", result["readRefs"])
        self.assertIn("refs/input-tables.md", result["readRefs"])
        commands = result["nextCommands"]
        self.assertIn("draft-table-report.py", commands[0])
        self.assertIn(str(self.out / "plan.json"), commands[0])
        self.assertIn("build-table-report.py", commands[1])
        self.assertNotIn("--post", commands[1])
        self.assertIn("--post --review-dispositions", commands[2])
        self.assertFalse(any("draft-forecast-plan.py" in c for c in commands))
        self.assertIn("Rebuild the source charts (CHARTS_PRESENT)", result["userHandoffs"])
        self.assertTrue(any("RUNNING_BALANCE_PARTITION_REVIEW" in n for n in result["notes"]))

    def test_the_broader_eight_table_workbooks_route_to_table_report(self):
        fixtures = ROOT / "tests" / "fixtures" / "broader-eight"
        for name, tables in (("04 Sales Pipeline Forecast.xlsx", 3), ("06 13-Week Cash Flow.xlsx", 4),
                             ("07 Headcount and Opex Plan.xlsx", 2), ("08 Inventory Replenishment Plan.xlsm", 3)):
            code, result = self.plan(fixtures / name)
            self.assertEqual(result["archetype"], "table-report", name)
            self.assertIn(f"{tables} table(s)", result["routeReason"])
            # the forecast drafter's rejection is kept as context, not as the route
            self.assertIn("forecast drafter rejects the shape", result["routeReason"])
        _, inventory = self.plan(fixtures / "08 Inventory Replenishment Plan.xlsm")
        self.assertEqual([m["status"] for m in inventory["macros"]], ["AUTO"])  # macro-classify no longer crashes

    def test_the_13_week_cash_flow_routes_to_table_report_instead_of_cell_model(self):
        code, result = self.plan(ROOT / "tests" / "fixtures" / "broader-eight" / "06 13-Week Cash Flow.xlsx")
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "table-report")
        self.assertIn("table-report drafter accepts this workbook: 4 table(s), 3 frames, 166 source cells verified",
                      result["routeReason"])
        self.assertIn("forecast drafter rejects the shape", result["routeReason"])   # kept as context only
        self.assertIn("draft-table-report.py", result["nextCommands"][0])
        self.assertFalse(any("source-workbook.py" in c for c in result["nextCommands"]))
        self.assertIn("Set the selected scenario and other parameters in the Parameters input table",
                      " ".join(result["userHandoffs"]))

    def test_a_forecast_workbook_is_still_routed_to_the_forecast_drafter(self):
        code, result = self.plan(ROOT / "tests" / "fixtures" / "broader-eight" / "02 SaaS Driver Forecast FY26.xlsx")
        self.assertEqual(result["archetype"], "forecast")

    def test_table_workbook_with_an_unrecognised_formula_names_the_cell_and_stays_cell_model(self):
        good = self.dir / "staff.xlsx"
        table_fixture.build_staff(good)
        bad = self.dir / "staff-bad.xlsx"
        table_fixture.patch_cell(good, bad, "Dept", "B4", formula="INDIRECT(\"Cost!G4\")")
        code, result = self.plan(bad)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "cell-model")
        self.assertIn("table-report drafter rejects the shape", result["routeReason"])
        self.assertIn("Dept!B4", result["routeReason"])
        self.assertFalse(any("draft-table-report.py" in c for c in result["nextCommands"]))

    def test_unknown_archetype_asks_the_user(self):
        wb = Workbook()
        wb.active["A1"] = "notes"
        path = self.dir / "unknown.xlsx"
        wb.save(path)
        code, result = self.plan(path)
        self.assertEqual(code, 0)
        self.assertEqual(result["archetype"], "unknown")
        self.assertTrue(result["askUser"])

    def test_research_archetype_uses_draft_inference(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "FY results"
        ws.append(["PROFIT AND LOSS", 2022, 2023, 2024])
        ws.append(["Revenue", 10, 11, 12])
        path = self.dir / "research.xlsx"
        wb.save(path)
        code, result = self.plan(path)
        self.assertEqual(result["archetype"], "research")
        self.assertIn("--draft", result["nextCommands"][0])
        self.assertIn("refs/research-recipes.md", result["readRefs"])

    def test_blocked_workbook_exits_two_and_still_prints_the_record(self):
        wb = Workbook()
        wb.active["A1"] = "x"
        path = self.dir / "blocked.xlsx"
        wb.save(path)
        with zipfile.ZipFile(path, "a") as archive:
            archive.writestr("xl/activeX/activeX1.xml", "<activeXControl/>")
        code, result = self.plan(path)
        self.assertEqual(code, 2)
        self.assertEqual(result["auditStatus"], "BLOCKED")
        self.assertEqual(result["archetype"], "blocked")
        self.assertEqual(result["nextCommands"], [])
        self.assertIn("BLOCKED stops everything.", result["notes"])

    def test_review_template_carries_source_sha_and_one_entry_per_review_issue(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Category", "Value"])
        ws.append(["A", 1])
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=2), titles_from_data=True)
        ws.add_chart(chart, "D2")
        path = self.dir / "chart.xlsx"
        wb.save(path)
        code, result = self.plan(path)
        self.assertEqual(code, 0)
        self.assertEqual(result["auditStatus"], "REVIEW_REQUIRED")
        self.assertIn("refs/presentation-migration.md", result["readRefs"])
        template = json.loads((self.out / "review.template.json").read_text())
        self.assertEqual(result["reviewTemplate"], str(self.out / "review.template.json"))
        self.assertEqual(template["sourceSha256"], source_sha256(path))
        self.assertEqual(result["sourceSha256"], source_sha256(path))
        self.assertEqual(template["issues"], [{"code": "CHARTS_PRESENT", "owner": "", "disposition": "",
                                               "verified": False}])
        # The template is not itself a valid disposition (verified must be true).
        from migration_gate import MigrationGateError, gate_source
        with self.assertRaises(MigrationGateError):
            gate_source(path, str(self.out / "review.template.json"))

    def test_missing_source_is_an_error_not_a_traceback(self):
        code, result = self.plan(self.dir / "nope.xlsx")
        self.assertEqual(code, 1)
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
