#!/usr/bin/env python3
"""table-report period ledgers (the 13-week cash flow shape): the recognisers (table rows, cadence,
recursion, running balance, scenario INDIRECT family, keyless tables), the oracle, the Sigma emission,
the loud failures, and the 06 13-Week Cash Flow fixture end to end."""
from __future__ import annotations

import copy
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import table_report as tr  # noqa: E402
import table_report_fixture as fx  # noqa: E402
import table_report_model as trm  # noqa: E402
from test_table_report import BUILD_CLI, DRAFT_CLI, check_spec_refs, run_cli  # noqa: E402

CASH_FLOW = ROOT / "tests" / "fixtures" / "broader-eight" / "06 13-Week Cash Flow.xlsx"


class Workdir(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name).resolve()

    def cash(self, name="cash", **options):
        path = self.dir / f"{name}.xlsx"
        info = fx.build_cash(path, **options)
        plan, summary = tr.draft_table_report(path, intent="what-if", connection="conn", folder="folder")
        return path, plan, summary, info

    def rejected(self, path):
        with self.assertRaises(tr.PlanDraftError) as caught:
            tr.draft_table_report(path, intent="what-if", connection="conn", folder="folder")
        return str(caught.exception)

    def patched(self, sheet, coord, **change):
        good = self.dir / "good.xlsx"
        fx.build_cash(good)
        bad = self.dir / f"bad-{coord}.xlsx"
        fx.patch_cell(good, bad, sheet, coord, **change)
        return bad

    def patched_row(self, row, template, cols="BCDEF"):
        """Rewrite one ledger row's formula in every period column (``{c}`` = the column letter)."""
        source = self.dir / "row-0.xlsx"
        fx.build_cash(source)
        for number, col in enumerate(cols, 1):
            target = self.dir / f"row-{number}.xlsx"
            fx.patch_cell(source, target, "Cash", f"{col}{row}", formula=template.format(c=col))
            source = target
        return source


def ledger(plan):
    return next(f for f in plan["frames"] if f["kind"] == "ledger")


def column(plan, name):
    return next(c for c in ledger(plan)["columns"] if c["name"] == name)


def rows_of(plan):
    return trm.compute(plan).frame_rows(ledger(plan)["id"])


# --------------------------------------------------------------------------- recognisers
class TestLedgerRecognisers(Workdir):
    def setUp(self):
        super().setUp()
        self.path, self.plan, self.info, self.book = self.cash()

    def test_every_formula_cell_is_verified_and_the_frames_are_named(self):
        self.assertEqual((self.info["verifiedCells"], self.info["formulaCells"]), (69, 69))
        self.assertEqual([k for k, _ in self.info["frames"]], ["summary", "summary", "ledger"])
        frame = ledger(self.plan)
        self.assertEqual((frame["title"], frame["axis"], frame["spine"]), ("Cash", "Wk", "Cash Periods"))
        self.assertEqual(frame["order"][:4], ["Wk", "Week Ending", "Beginning Cash", "Collections"])
        self.assertEqual(frame["totals"][:3], ["Collections", "New Sales", "Total Receipts"])

    def test_every_period_of_every_row_matches_an_independent_recomputation(self):
        expected = self.book["expected"]
        rows = rows_of(self.plan)
        for name, values in expected.items():
            for row, want in zip(rows, values):
                got = row[name]
                if isinstance(want, date):
                    self.assertEqual(got, want, name)
                else:
                    self.assertAlmostEqual(got, want, places=6, msg=name)

    def test_table_row_reads_a_hidden_grouped_helper_and_defaults_to_zero(self):
        expr = column(self.plan, "Payments")["expr"]
        self.assertEqual(expr, ["lk", "tblOut by Pay Week", ["period"], "Pay Week", "Amount", ["n", 0]])
        helper = next(f for f in self.plan["frames"] if f["title"] == "tblOut by Pay Week")
        self.assertEqual((helper["kind"], helper["helper"], helper["hidden"], helper["labels"]),
                         ("summary", True, True, [1, 4]))
        self.assertEqual(helper["columns"][0]["expr"], ["agg", "sum", "Amount", None])
        weeks_without_bills = [r["Payments"] for r in rows_of(self.plan) if r["Wk"] in (2, 3, 5)]
        self.assertEqual(weeks_without_bills, [0, 0, 0])      # Lookup misses are 0, like SUMIFS

    def test_cadence_is_if_mod_over_the_period(self):
        expr = column(self.plan, "Payroll")["expr"]
        self.assertEqual(expr[:2], ["f", "IF"])
        self.assertEqual(expr[2], ["op", "=", ["f", "MOD", ["period"], ["n", 2]], ["n", 0]])
        self.assertEqual([r["Payroll"] for r in rows_of(self.plan)], [0, 60, 0, 60, 0])

    def test_growth_recursion_is_lowered_to_its_closed_form(self):
        expr = column(self.plan, "New Sales")["expr"]
        self.assertEqual(expr, ["op", "*", ["p", "Seed receipts"],
                                ["op", "^", ["op", "+", ["n", 1], ["p", "Weekly growth"]],
                                 ["op", "-", ["period"], ["n", 1]]]])
        self.assertAlmostEqual(rows_of(self.plan)[4]["New Sales"], 80 * 1.1 ** 4)
        self.assertIn("RECURSION_CLOSED_FORM", self.info["reviewIssues"])

    def test_date_row_steps_from_a_date_parameter(self):
        self.assertEqual(column(self.plan, "Week Ending")["expr"],
                         ["f", "ADDDAYS", ["pd", "Week 1 ends"], ["op", "*", ["n", 7], ["op", "-", ["period"], ["n", 1]]]])
        self.assertEqual(column(self.plan, "Week Ending")["type"], "datetime")
        param = next(p for p in self.plan["params"] if p["name"] == "Week 1 ends")
        self.assertEqual((param["kind"], param["value"]), ("date", "2026-01-02"))
        self.assertEqual(rows_of(self.plan)[4]["Week Ending"], date(2026, 1, 30))

    def test_running_balance_is_opening_plus_cumulative_net_and_beginning_is_ending_less_net(self):
        self.assertEqual(column(self.plan, "Ending Cash")["expr"],
                         ["op", "+", ["p", "Opening cash"], ["cumsum", ["c", "Net"], None, "Wk"]])
        self.assertEqual(column(self.plan, "Beginning Cash")["expr"],
                         ["op", "-", ["c", "Ending Cash"], ["c", "Net"]])
        rows = rows_of(self.plan)
        for earlier, later in zip(rows, rows[1:]):      # Excel's carry: this week begins where last week ended
            self.assertAlmostEqual(later["Beginning Cash"], earlier["Ending Cash"])
        self.assertEqual(rows[0]["Beginning Cash"], fx.CASH_OPENING)
        self.assertIn("LEDGER_RUNNING_BALANCE_REVIEW", self.info["reviewIssues"])

    def test_running_min_over_the_period_columns_is_a_cumulative_min(self):
        self.assertEqual(column(self.plan, "Lowest Ending")["expr"], ["cummin", ["c", "Ending Cash"], None, "Wk"])
        ends = [r["Ending Cash"] for r in rows_of(self.plan)]
        self.assertEqual([r["Lowest Ending"] for r in rows_of(self.plan)],
                         [min(ends[:i + 1]) for i in range(len(ends))])

    def test_sum_over_a_column_span_of_rows_and_derived_rows(self):
        self.assertEqual(column(self.plan, "Total Out")["expr"],
                         ["f", "SUM", ["c", "Payroll"], ["c", "Rent"], ["c", "Payments"]])
        self.assertEqual(column(self.plan, "Cushion")["expr"],
                         ["op", "-", ["c", "Ending Cash"], ["p", "Minimum cash"]])    # a named cell is a parameter

    def test_typed_rows_become_editable_spine_columns(self):
        spine = next(t for t in self.plan["tables"] if t["name"] == "Cash Periods")
        self.assertEqual([c["name"] for c in spine["columns"]], ["Wk", "Rent"])
        self.assertEqual(spine["rows"], [[i + 1, v] for i, v in enumerate(fx.CASH_RENT)])
        self.assertTrue(spine["editable"] and spine["synthetic"])
        self.assertIn("LEDGER_MANUAL_LINES", self.info["reviewIssues"])

    def test_total_column_is_verified_and_emitted_as_a_one_row_table(self):
        frame = ledger(self.plan)
        self.assertEqual(frame["totalLabel"], "Total")
        groups = {g["column"]: g for g in self.plan["checks"] if g["kind"] == "ledger-total"}
        self.assertEqual(set(groups), set(frame["totals"]))
        self.assertAlmostEqual(trm.compute(self.plan).cell(frame["id"], "Net", "ledger-total", []),
                               sum(self.book["expected"]["Net"]))

    def test_caption_is_verified_but_not_emitted(self):
        frame = ledger(self.plan)
        self.assertEqual(frame["captions"], [{"cell": "Cash!A1", "expr": ["f", "CONCAT", ["s", "Case: "],
                                                                        ["pt", "Scenario"]]}])
        self.assertIn("CAPTION_FORMULA_VERIFIED_ONLY", self.info["reviewIssues"])
        parts = trm.build_parts(self.plan, "<dm>", {})
        self.assertFalse(any("Case:" in (c.get("formula") or "") for e in parts["visible"]
                             for c in e.get("columns", [])))

    def test_section_headings_are_not_rows(self):
        self.assertNotIn("Receipts", ledger(self.plan)["order"])
        self.assertNotIn("Disbursements", ledger(self.plan)["order"])


class TestScenarioFamily(Workdir):
    def test_indirect_over_a_named_range_family_becomes_a_keyed_lookup(self):
        _, plan, info, book = self.cash()
        scen = next(t for t in plan["tables"] if t["name"] == "Collection rate by Scenario")
        self.assertEqual([c["name"] for c in scen["columns"]], ["Scenario", "Collection rate"])
        self.assertEqual(scen["rows"], [["Base", 0.9], ["Low", 0.5], ["High", 1.0]])
        self.assertEqual(column(plan, "Collections")["expr"],
                         ["op", "*", ["lk", "tblIn by Week", ["period"], "Week", "Amount", ["n", 0]],
                          ["lk", "Collection rate by Scenario", ["pt", "Scenario"], "Scenario",
                           "Collection rate", None]])
        param = next(p for p in plan["params"] if p["name"] == "Scenario")
        self.assertEqual((param["kind"], param["value"]), ("text", "Base"))
        self.assertIn("INDIRECT_SCENARIO_LOOKUP", info["reviewIssues"])
        issue = next(i for i in plan["draftReviewIssues"] if i["code"] == "INDIRECT_SCENARIO_LOOKUP")
        self.assertIn("Lookup", issue["message"])
        self.assertEqual(issue["evidence"][0], "Cash!B7")

    def test_the_other_scenarios_are_recomputed_as_information_not_parity(self):
        _, plan, _, _ = self.cash()
        result = plan["scenarioResults"][ledger(plan)["id"]]
        self.assertEqual((result["parameter"], result["selected"]), ("Scenario", "Base"))
        self.assertEqual(set(result["results"]), {"Base", "Low", "High"})
        for name in ("Base", "Low", "High"):
            want = fx.cash_expected(name)
            got = result["results"][name]
            self.assertAlmostEqual(got["Cash / Ending Cash / Wk 5"], want["Ending Cash"][-1], places=6)
            self.assertAlmostEqual(got["Cash / Net / Total"], sum(want["Net"]), places=6)
            self.assertAlmostEqual(got["Cash / Lowest Ending / Wk 5"], min(want["Ending Cash"]), places=6)
        base_total = result["results"]["Base"]["Cash / Total Receipts / Total"]
        self.assertGreater(result["results"]["High"]["Cash / Total Receipts / Total"], base_total)
        self.assertLess(result["results"]["Low"]["Cash / Total Receipts / Total"], base_total)
        self.assertIn("not parity", result["informational"])
        self.assertTrue(all(t["frame"] != "scenario" for t in plan["parityTargets"]))

    def test_a_workbook_saved_on_another_scenario_verifies_against_that_scenario(self):
        _, plan, _, book = self.cash(scenario="Low")
        self.assertAlmostEqual(rows_of(plan)[0]["Collections"], 300 * 0.5)
        self.assertEqual(plan["scenarioResults"][ledger(plan)["id"]]["selected"], "Low")

    def test_the_selected_scenario_must_be_one_of_the_family(self):
        bad = self.dir / "unknown.xlsx"
        fx.build_cash(bad, selected="Nope")
        message = self.rejected(bad)
        self.assertIn("not one of the named ranges Rate_*", message)
        self.assertIn("Cash!B7", message)

    def test_indirect_of_anything_else_is_refused(self):
        bad = self.patched("Cash", "B7", formula='SUMIFS(tblIn[Amount],tblIn[Week],B$3)*INDIRECT("Inputs!B11")')
        message = self.rejected(bad)
        self.assertIn("Cash!B7", message)
        self.assertIn("INDIRECT", message)

    def test_indirect_outside_a_ledger_stays_refused(self):
        # the lowering is recognised inside a ledger only; elsewhere INDIRECT is a loud failure
        message = self.rejected(self.patched("Cash", "B7", formula='INDIRECT("Rate_Base")'))
        self.assertIn("INDIRECT", message)


class TestKeylessTables(Workdir):
    def test_a_table_without_a_unique_key_drafts_and_is_flagged_not_keyed(self):
        _, plan, info, _ = self.cash()
        issue = next(i for i in plan["draftReviewIssues"] if i["code"] == "KEYLESS_TABLE_ROWS")
        self.assertEqual(issue["evidence"], ["tblOut"])          # tblIn has a unique Invoice column
        out = next(t for t in plan["tables"] if t["name"] == "tblOut")
        self.assertEqual([r[0] for r in out["rows"]], ["Acme", "Acme", "Bolt"])
        self.assertNotIn("Row", [c["name"] for c in out["columns"]])   # no synthetic key was needed
        self.assertEqual(rows_of(plan)[0]["Payments"], 50.0)            # both Acme rows summed in week 1


# --------------------------------------------------------------------------- loud failures
class TestLedgerFailures(Workdir):
    def test_a_tampered_cached_value_names_the_cell(self):
        message = self.rejected(self.patched("Cash", "D15", value=999))
        self.assertIn("Cash!D15", message)
        self.assertIn("999", message)

    def test_an_unsupported_function_names_the_cell(self):
        message = self.rejected(self.patched_row(17, "ROUNDUP({c}16,0)"))
        self.assertIn("Cash!B17", message)
        self.assertIn("unsupported function ROUNDUP", message)

    def test_a_row_that_is_not_one_formula_across_names_the_deviating_cell(self):
        message = self.rejected(self.patched("Cash", "E15", formula="E9-E14+1"))
        self.assertIn("Cash!E15", message)
        self.assertIn("copied across", message)

    def test_a_recursion_that_is_not_growth_or_a_day_step_is_refused(self):
        message = self.rejected(self.patched_row(8, "{c}8*2+1", cols="CDEF"))
        self.assertIn("Cash!C8", message)
        self.assertIn("not a recognised recursion", message)

    def test_a_total_cell_that_is_not_a_sum_across_the_periods_is_refused(self):
        message = self.rejected(self.patched("Cash", "G11", formula="SUM(B11:E11)"))
        self.assertIn("Cash!G11", message)
        self.assertIn("SUM across the period columns", message)

    def test_a_formula_that_reads_another_period_is_refused(self):
        message = self.rejected(self.patched_row(9, "{c}7+{c}8+$B$7"))
        self.assertIn("Cash!B9", message)
        self.assertIn("fixed-column reference", message)

    def test_a_relative_reference_outside_the_ledger_is_refused(self):
        message = self.rejected(self.patched_row(8, "{c}7*Inputs!{c}6"))
        self.assertIn("Cash!B8", message)
        self.assertIn("relative reference to a cell outside the ledger", message)

    def test_a_formula_outside_the_ledger_is_never_skipped(self):
        stray = self.dir / "stray.xlsx"
        fx.build_cash(stray, stray=True)
        message = self.rejected(stray)
        self.assertIn("Cash!B30", message)

    def test_a_carry_forward_row_must_be_beginning_plus_net(self):
        message = self.rejected(self.patched_row(16, "{c}5*{c}15"))
        self.assertIn("Cash!B16", message)


# --------------------------------------------------------------------------- model pieces
class TestLedgerModelPieces(unittest.TestCase):
    def evaluate(self, node, **env):
        model = trm.Model({"tables": [], "frames": [], "params": [{"name": "d", "value": "2026-03-01", "kind": "date"},
                                                                    {"name": "t", "value": "Base", "kind": "text"}]})
        return model.ev(node, trm.Env(**env))

    def test_mod_adddays_and_concat_follow_excel(self):
        self.assertEqual(self.evaluate(["f", "MOD", ["n", 7], ["n", 3]]), 1)
        self.assertEqual(self.evaluate(["f", "MOD", ["n", -7], ["n", 3]]), 2)          # sign of the divisor
        self.assertEqual(self.evaluate(["f", "ADDDAYS", ["pd", "d"], ["n", 7]]), date(2026, 3, 8))
        self.assertEqual(self.evaluate(["f", "CONCAT", ["s", "Case: "], ["pt", "t"]]), "Case: Base")
        with self.assertRaises(trm.EvalError):
            self.evaluate(["f", "MOD", ["n", 1], ["n", 0]])
        with self.assertRaises(trm.EvalError):
            self.evaluate(["f", "CONCAT", ["s", "a"], ["n", 1]])

    def test_running_sum_and_min_without_a_partition_follow_the_period_order(self):
        frame = {"id": "f", "title": "f", "kind": "ledger", "spine": "S", "axis": "Wk", "order": [],
                 "columns": [{"name": "Run", "type": "number", "expr": ["cumsum", ["c", "V"], None, "Wk"]},
                             {"name": "Low", "type": "number", "expr": ["cummin", ["c", "Run"], None, "Wk"]}]}
        plan = {"archetype": "table-report", "planVersion": 1, "params": [], "checks": [], "frames": [frame],
                "tables": [{"name": "S", "columns": [{"name": "Wk", "type": "number"},
                                                      {"name": "V", "type": "number"}],
                            "rows": [[3, -5], [1, 4], [2, -6]]}]}
        rows = trm.compute(plan).frame_rows("f")
        self.assertEqual([(r["Wk"], r["Run"], r["Low"]) for r in rows], [(1, 4, 4), (2, -2, -2), (3, -7, -7)])

    def test_a_ledger_column_cycle_is_a_plan_error(self):
        frame = {"id": "f", "title": "f", "kind": "ledger", "spine": "S", "axis": "Wk", "order": [],
                 "columns": [{"name": "A", "type": "number", "expr": ["c", "B"]},
                             {"name": "B", "type": "number", "expr": ["c", "A"]}]}
        plan = {"archetype": "table-report", "planVersion": 1, "params": [], "checks": [], "frames": [frame],
                "tables": [{"name": "S", "columns": [{"name": "Wk", "type": "number"}], "rows": [[1]]}]}
        with self.assertRaises(trm.PlanError):
            trm.compute(plan)

    def test_new_ir_nodes_are_validated(self):
        for bad in (["cummin", ["c", "x"]], ["pt"], ["f", "MOD"], ["pd", "a", "b"]):
            with self.assertRaises(trm.PlanError):
                trm.validate_ir(bad)
        trm.validate_ir(["cummin", ["c", "x"], None, "Wk"])

    def test_emitter_renders_the_new_nodes(self):
        em = trm.Emitter(col=lambda n: f"[{n}]", period="[Wk]")
        self.assertEqual(em.emit(["f", "MOD", ["period"], ["n", 2]]), "Mod([Wk], 2)")
        self.assertEqual(em.emit(["op", "^", ["n", 2], ["period"]]), "Power(2, [Wk])")
        self.assertEqual(em.emit(["cummin", ["c", "E"], None, "Wk"]), "CumulativeMin([E])")
        self.assertEqual(em.emit(["f", "ADDDAYS", ["pd", "d"], ["n", 7]]),
                         'DateAdd("day", 7, Lookup([Parameters/Date Value], "d", [Parameters/Parameter]))')
        self.assertEqual(em.emit(["pt", "t"]), 'Lookup([Parameters/Text Value], "t", [Parameters/Parameter])')

    def test_parameter_table_gains_typed_columns_only_when_needed(self):
        plain = trm.param_table({"params": [{"name": "a", "value": 1, "cell": "S!B1"}]})
        self.assertEqual([c["name"] for c in plain["columns"]], ["Parameter", "Value", "Cell"])
        typed = trm.param_table({"params": [{"name": "a", "value": 1, "cell": "S!B1"},
                                            {"name": "s", "value": "Base", "cell": "S!B2", "kind": "text"},
                                            {"name": "d", "value": "2026-01-02", "cell": "S!B3", "kind": "date"}]})
        self.assertEqual([c["name"] for c in typed["columns"]],
                         ["Parameter", "Value", "Cell", "Text Value", "Date Value"])
        self.assertEqual(typed["rows"], [["a", 1, "S!B1", None, None], ["s", None, "S!B2", "Base", None],
                                         ["d", None, "S!B3", None, "2026-01-02"]])

    def test_other_grids_still_refuse_non_numeric_parameters(self):
        reg = tr.Registry.__new__(tr.Registry)      # only param bookkeeping is under test
        reg.wf = None
        reg.params, reg.param_names, reg.covered = {}, {}, {}
        reg.params[("Inputs", "B3")] = {"name": "Scenario", "value": "Base", "cell": "Inputs!B3", "kind": "text"}
        with self.assertRaises(tr.PlanDraftError):
            reg.param("Inputs", "B3", "X!A1")      # numbers only unless the ledger context widens ``allow``


# --------------------------------------------------------------------------- emission and builder
class TestLedgerEmission(Workdir):
    def setUp(self):
        super().setUp()
        self.path, self.plan, self.info, self.book = self.cash()
        self.parts = trm.build_parts(self.plan, "<dm>", {})

    def test_the_ledger_is_one_table_sorted_by_period_with_formulas_in_sheet_order(self):
        element = next(e for e in self.parts["visible"] if e["name"] == "Cash")
        self.assertEqual([c["name"] for c in element["columns"]][:5],
                         ["Wk", "Week Ending", "Beginning Cash", "Collections", "New Sales"])
        formulas = {c["name"]: c["formula"] for c in element["columns"]}
        self.assertEqual(formulas["Wk"], "[Cash Periods/Wk]")
        self.assertEqual(formulas["Rent"], "Coalesce([Cash Periods/Rent], 0)")
        self.assertEqual(formulas["Payroll"],
                         'If((Mod([Wk], 2) = 0), Lookup([Parameters/Value], "Payroll", [Parameters/Parameter]), 0)')
        self.assertEqual(formulas["Payments"],
                         "Coalesce(Lookup([tblOut by Pay Week/Amount], [Wk], [tblOut by Pay Week/Pay Week]), 0)")
        self.assertEqual(formulas["Ending Cash"],
                         '(Lookup([Parameters/Value], "Opening cash", [Parameters/Parameter]) + '
                         'CumulativeSum([Net]))')
        self.assertEqual(formulas["Beginning Cash"], "([Ending Cash] - [Net])")
        self.assertEqual(formulas["Lowest Ending"], "CumulativeMin([Ending Cash])")
        self.assertEqual(formulas["New Sales"],
                         '(Lookup([Parameters/Value], "Seed receipts", [Parameters/Parameter]) * '
                         'Power((1 + Lookup([Parameters/Value], "Weekly growth", [Parameters/Parameter])), '
                         '([Wk] - 1)))')
        self.assertIn("Lookup([Parameters/Text Value], \"Scenario\", [Parameters/Parameter])", formulas["Collections"])
        self.assertEqual(formulas["Week Ending"].split(",")[0], 'DateAdd("day"')
        self.assertEqual(element["sort"][0]["direction"], "ascending")
        self.assertNotIn("groupings", element)         # one running series over the whole table

    def test_helpers_are_hidden_grouped_tables_and_the_total_is_one_grouped_row(self):
        hidden = {e["name"]: e for e in self.parts["hidden"]}
        helper = hidden["tblOut by Pay Week"]
        self.assertEqual([c["formula"] for c in helper["columns"]],
                         ["[tblOut/Pay Week]", "Coalesce(Sum([tblOut/Amount]), 0)"])
        self.assertEqual(len(helper["groupings"]), 1)
        visible = {e["name"]: e for e in self.parts["visible"]}
        self.assertNotIn("tblOut by Pay Week", visible)
        total = visible["Cash - Total"]
        self.assertEqual(total["columns"][0]["formula"], '"Total"')
        self.assertEqual(total["groupings"][0]["groupBy"], [total["columns"][0]["id"]])
        self.assertEqual(total["columns"][-1]["formula"], "Sum([Cash/Net])")

    def test_spec_references_resolve(self):
        import workbook_wire
        spec = workbook_wire.wire_workbook(trm.build_wb_spec(self.plan, "<dm>", {}))
        check_spec_refs(self, spec)

    def test_read_only_intent_hosts_the_tables_in_the_data_model_including_typed_parameters(self):
        plan = tr.draft_table_report(self.path, intent="read-only", connection="c", folder="f")[0]
        dm = trm.build_dm_spec(plan)
        statements = [e["source"]["statement"] for e in dm["pages"][0]["elements"]]
        self.assertTrue(any("NULL::varchar" in s and "'Base'" in s for s in statements))      # Parameters
        self.assertTrue(any("'2026-01-02'::date" in s for s in statements))
        import workbook_wire
        check_spec_refs(self, workbook_wire.wire_workbook(
            trm.build_wb_spec(plan, "<dm>", trm.map_dm_elements(plan, dm))))

    def test_builder_dry_run_reports_parity_pass_and_writes_the_specs(self):
        plan_path = self.dir / "plan.json"
        code, out, err = run_cli(DRAFT_CLI, [str(self.path), "--connection", "c", "--folder", "f",
                                             "--out", str(plan_path)])
        self.assertEqual(code, 0, err)
        self.assertIn("frame ledger: Cash", out)
        code, out, err = run_cli(BUILD_CLI, [str(plan_path), "--out-dir", str(self.dir / "build")])
        self.assertEqual(code, 0, err)
        self.assertIn("LOCAL PARITY: 69/69 source cells reproduced", out)
        self.assertIn("PARITY PASS Cash: 14/14 summary cells", out)
        self.assertNotIn("FAIL", out)
        self.assertTrue((self.dir / "build" / "wb_spec.json").exists())
        self.assertTrue((self.dir / "build" / "seed_cash_periods.csv").exists())

    def test_the_oracle_notices_an_edited_input_table(self):
        plan = copy.deepcopy(self.plan)
        next(t for t in plan["tables"] if t["name"] == "tblIn")["rows"][0][1] += 500
        model = trm.compute(plan)
        _, problems = trm.check_cells(plan, model)
        self.assertTrue(any("Cash!" in p for p in problems))


# --------------------------------------------------------------------------- 06 13-Week Cash Flow
class TestCashFlowFixture(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.plan, cls.info = tr.draft_table_report(CASH_FLOW, intent="what-if", connection="conn",
                                                   folder="folder")
        cls.model = trm.compute(cls.plan)

    def test_every_formula_cell_is_verified(self):
        self.assertEqual((self.info["verifiedCells"], self.info["formulaCells"]), (166, 166))
        self.assertEqual([k for k, _ in self.info["frames"]], ["summary", "summary", "ledger"])
        self.assertEqual(self.info["tables"], ["tblAR", "tblAP", "Collection Rate by Scenario", "13WCF Periods"])
        self.assertEqual(self.plan["archetype"], "table-report")

    def test_headline_values_match_excel(self):
        frame = ledger(self.plan)
        rows = self.model.frame_rows(frame["id"])
        self.assertEqual(len(rows), 13)
        self.assertAlmostEqual(rows[-1]["Ending Cash"], 1619042.88368296, places=6)
        self.assertAlmostEqual(rows[-1]["Lowest Ending Cash to Date"], 899581.997838681, places=6)
        cell = lambda name: self.model.cell(frame["id"], name, "ledger-total", [])  # noqa: E731
        self.assertAlmostEqual(cell("Total Receipts"), 3183042.88368296, places=6)
        self.assertEqual(cell("Total Disbursements"), 2814000)
        self.assertAlmostEqual(cell("Net Cash Flow"), 369042.88368296, places=6)
        self.assertAlmostEqual(rows[-1]["Ending Cash"] - 1250000, cell("Net Cash Flow"), places=6)

    def test_parity_targets_cover_ending_cash_totals_and_the_lowest_balance(self):
        labels = [t["label"] for t in self.plan["parityTargets"]]
        self.assertEqual(sum(1 for label in labels if "/ Ending Cash /" in label), 13)
        for name in ("Total Receipts", "Total Disbursements", "Net Cash Flow"):
            self.assertIn(f"13WCF / {name} / 13-Wk Total", labels)
        self.assertIn("13WCF / Lowest Ending Cash to Date / Week 13", labels)
        self.assertEqual(len(labels), 23)
        ok, lines = trm.parity_lines(self.plan, self.model)
        self.assertTrue(ok)
        self.assertIn("PARITY PASS 13WCF: 23/23 summary cells", lines)

    def test_the_three_scenarios_are_recomputed_and_the_saved_one_matches_the_cache(self):
        result = self.plan["scenarioResults"][ledger(self.plan)["id"]]
        self.assertEqual((result["parameter"], result["selected"]), ("Scenario", "Base"))
        results = result["results"]
        self.assertEqual(set(results), {"Base", "Downside", "Upside"})
        base = results["Base"]
        self.assertAlmostEqual(base["13WCF / Ending Cash / Week 13"], 1619042.88368296, places=6)
        self.assertAlmostEqual(base["13WCF / Net Cash Flow / 13-Wk Total"], 369042.88368296, places=6)
        # independent expectation: collections scale linearly with the scenario rate, nothing else moves
        collections = 2009250 / 0.95
        for name, rate in (("Downside", 0.82), ("Upside", 1.0)):
            shift = collections * (rate - 0.95)
            self.assertAlmostEqual(results[name]["13WCF / Net Cash Flow / 13-Wk Total"],
                                   369042.88368296 + shift, places=4)
            self.assertAlmostEqual(results[name]["13WCF / Ending Cash / Week 13"],
                                   1619042.88368296 + shift, places=4)
            self.assertAlmostEqual(results[name]["13WCF / Total Disbursements / 13-Wk Total"], 2814000)
        self.assertLess(results["Downside"]["13WCF / Lowest Ending Cash to Date / Week 13"],
                        base["13WCF / Lowest Ending Cash to Date / Week 13"])

    def test_review_issues_and_parameters(self):
        self.assertEqual(self.info["reviewIssues"][:5],
                         ["INDIRECT_SCENARIO_LOOKUP", "LEDGER_TRANSPOSED", "RECURSION_CLOSED_FORM",
                          "LEDGER_RUNNING_BALANCE_REVIEW", "CAPTION_FORMULA_VERIFIED_ONLY"])
        self.assertIn("KEYLESS_TABLE_ROWS", self.info["reviewIssues"])
        kinds = {p["name"]: p.get("kind", "number") for p in self.plan["params"]}
        self.assertEqual(kinds["Scenario"], "text")
        self.assertEqual(kinds["Week 1 Ending"], "date")
        scenario = next(p for p in self.plan["params"] if p["name"] == "Scenario")
        self.assertEqual(scenario["options"], ["Base", "Downside", "Upside"])
        keyless = next(i for i in self.plan["draftReviewIssues"] if i["code"] == "KEYLESS_TABLE_ROWS")
        self.assertEqual(keyless["evidence"], ["tblAP"])

    def test_the_source_chart_is_listed_against_the_verified_ending_cash_column(self):
        self.assertEqual(len(self.info["unmappedCharts"]), 1)
        self.assertIn("series read verified g_13wcf.Ending Cash", self.info["unmappedCharts"][0])

    def test_builder_runs_with_parity_pass_and_specs_resolve(self):
        import json
        import workbook_wire
        with tempfile.TemporaryDirectory() as tmp:
            plan = Path(tmp) / "plan.json"
            code, out, err = run_cli(DRAFT_CLI, [str(CASH_FLOW), "--connection", "c", "--folder", "f",
                                                 "--out", str(plan)])
            self.assertEqual(code, 0, err)
            code, out, err = run_cli(BUILD_CLI, [str(plan), "--out-dir", str(Path(tmp) / "build")])
            self.assertEqual(code, 0, err)
            self.assertIn("LOCAL PARITY: 166/166 source cells reproduced", out)
            self.assertIn("PARITY PASS 13WCF: 23/23 summary cells", out)
            self.assertNotIn("FAIL", out)
            spec = json.loads((Path(tmp) / "build" / "wb_spec.json").read_text())
            check_spec_refs(self, spec)
            self.assertTrue((Path(tmp) / "build" / "seed_13wcf_periods.csv").exists())
        plan = copy.deepcopy(self.plan)
        spec = workbook_wire.wire_workbook(trm.build_wb_spec(plan, "<dm>", {}))
        names = {e["name"] for e in spec["contents"]["elements"] if e.get("name")}
        self.assertTrue({"13WCF", "13WCF - 13-Wk Total", "tblAR by Expected Week", "tblAP by Pay Week",
                         "Parameters", "Collection Rate by Scenario"} <= names)


if __name__ == "__main__":
    unittest.main()
