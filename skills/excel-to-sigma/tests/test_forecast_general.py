#!/usr/bin/env python3
"""Forecast drafter + builder beyond the fixed four sections and monthly columns.

Covers discovered sections and their sign roles, quarter periods, year-over-year /
percent-of-total / alternative-growth / external-link forecast cells, generic derived rows
and parity, the second-pass percentage elements in the workbook, and regressions that keep
the classic four-section plans (and the 02 SaaS fixture) unchanged.
"""
from __future__ import annotations

import contextlib
import copy
import importlib.util
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import forecast_calc as calc  # noqa: E402
import forecast_pnl_fixture as fixture  # noqa: E402
from forecast_plan import PlanDraftError, draft_plan, parse_quarter  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
RETAIL = FIXTURES / "broader-eight" / "03 Retail Quarterly Budget.xlsx"
SAAS = FIXTURES / "broader-eight" / "02 SaaS Driver Forecast FY26.xlsx"


def load_script(filename, name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


DRAFT = load_script("draft-forecast-plan.py", "general_draft_cli")
BUILD = load_script("build-forecast-model.py", "general_builder")
ELS = {"glDetail": "glDetail", "spine": "spine", "assumptions": "assumptions"}


def run_main(module, argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = module.main(argv)
    return code, out.getvalue(), err.getvalue()


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        self.path = self.dir / "budget.xlsx"

    def quarterly(self, **kwargs):
        built = fixture.build_quarterly(self.path, **kwargs)
        plan, info = draft_plan(self.path, intent="what-if", connection="conn", folder="fold")
        return plan, info, built

    def quarterly_fails(self, pattern, **kwargs):
        fixture.build_quarterly(self.path, **kwargs)
        with self.assertRaisesRegex(PlanDraftError, pattern):
            draft_plan(self.path, intent="what-if", connection="conn", folder="fold")

    def write_plan(self, plan, name="plan.json"):
        path = self.dir / name
        path.write_text(json.dumps(plan), encoding="utf-8")
        return path


# ---------------------------------------------------------------- calc module
class TestRowExpressions(unittest.TestCase):
    def tree(self, text, resolve=None):
        return calc.to_entities(calc.parse_formula(text), resolve or Resolver())

    def test_parse_evaluate_and_render_a_ratio_row(self):
        expr = self.tree("=IF(B13=0,0,B22/B13)")
        self.assertEqual(calc.render_expr(expr), "IF([Sales] = 0, 0, [Profit] / [Sales])")

        class Env:
            def sec(self, name): return 200.0
            def drv(self, name): return 50.0
        self.assertAlmostEqual(calc.eval_expr(expr, Env()), 0.25)

        class Zero(Env):
            def sec(self, name): return 0.0
        self.assertEqual(calc.eval_expr(expr, Zero()), 0)

    def test_iferror_catches_division_by_zero_only_inside_it(self):
        expr = self.tree("=IFERROR(B22/B13,0)")

        class Env:
            def sec(self, name): return 0.0
            def drv(self, name): return 5.0
        self.assertEqual(calc.eval_expr(expr, Env()), 0)
        bare = self.tree("=B22/B13")
        with self.assertRaises(ZeroDivisionError):
            calc.eval_expr(bare, Env())

    def test_unsupported_syntax_is_rejected(self):
        for text in ("=B1^2", "=B1&B2", "=INDEX(B1:B3,1)", "=B1%"):
            with self.subTest(text=text):
                with self.assertRaises(calc.ExprError):
                    calc.parse_formula(text)

    def test_relative_shape_matches_across_columns(self):
        a = calc.relative_shape(calc.parse_formula("=B9-B14"), 2)
        b = calc.relative_shape(calc.parse_formula("=C9-C14"), 3)
        c = calc.relative_shape(calc.parse_formula("=C9-C15"), 3)
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)

    def test_linear_expansion_gives_section_coefficients(self):
        derived = {"GM": ["bin", "-", ["sec", "Sales"], ["sec", "COGS"]]}
        expr = ["bin", "-", ["drv", "GM"], ["sec", "SGA"]]
        self.assertEqual(calc.linear_over_sections(expr, {}, derived),
                         {"Sales": 1.0, "COGS": -1.0, "SGA": -1.0})
        ratio = ["bin", "/", ["drv", "GM"], ["sec", "Sales"]]
        self.assertIsNone(calc.linear_over_sections(ratio, {}, derived))

    def test_yoy_terms_chain_until_an_actual_period(self):
        # four actual quarters (indexes 0-3), lag 4: forecast 4 and 5 read 0 and 1
        self.assertEqual(calc.yoy_terms(4, 3, 4), (1, 0))
        self.assertEqual(calc.yoy_terms(5, 3, 4), (1, 1))
        # a lag shorter than the forecast chains through forecast periods
        self.assertEqual(calc.yoy_terms(6, 3, 2), (2, 2))


class Resolver:
    names = {13: ["sec", "Sales"], 22: ["drv", "Profit"]}

    def __call__(self, sheet, col, row):
        return list(self.names[row])

    def range(self, sheet, col, a, b):
        return [list(self.names[r]) for r in range(a, b + 1) if r in self.names]


# ---------------------------------------------------------------- sections
class TestDiscoveredSections(Base):
    def test_sections_totals_and_derived_rows_are_discovered(self):
        plan, info, built = self.quarterly()
        self.assertEqual([(s["name"], s["sign"], s["signSource"]) for s in plan["sections"]],
                         [("Net Sales", 1, "derived-rows"), ("COGS", -1, "derived-rows"),
                          ("SG&A", -1, "derived-rows"), ("Other Income/Expense", -1, "derived-rows")])
        self.assertEqual([s["totalLabel"] for s in plan["sections"]],
                         ["Total Net Sales", "Total COGS", "Total SG&A", "Total Other Income/Expense"])
        self.assertEqual([(d["label"], d["kind"]) for d in plan["derivedRows"]],
                         [("Gross Margin", "amount"), ("Gross Margin %", "ratio"),
                          ("Operating Profit", "amount"), ("Pre-tax Income", "amount")])
        pretax = plan["derivedRows"][-1]
        self.assertEqual(pretax["coefficients"],
                         {"Net Sales": 1, "COGS": -1, "SG&A": -1, "Other Income/Expense": -1})
        self.assertEqual(pretax["formula"], "[Operating Profit] - [Other Income/Expense]")
        self.assertEqual([l["secOrder"] for l in plan["lines"]], [1, 1, 2, 2, 3, 3, 3, 4, 4])

    def test_ratio_row_keeps_its_formula_and_cached_values(self):
        plan, _, built = self.quarterly()
        ratio = next(d for d in plan["derivedRows"] if d["label"] == "Gross Margin %")
        self.assertEqual(ratio["formula"], "IF([Net Sales] = 0, 0, [Gross Margin] / [Net Sales])")
        self.assertEqual(list(ratio["values"]), fixture.QUARTERS)
        for label, value in zip(fixture.QUARTERS, built["derived"]["Gross Margin %"]):
            self.assertAlmostEqual(ratio["values"][label], value, places=9)

    def test_parity_targets_cover_every_total_and_derived_row(self):
        plan, info, built = self.quarterly()
        self.assertEqual(set(plan["parityTargets"]), set(built["targets"]))
        for label, value in built["targets"].items():
            self.assertAlmostEqual(plan["parityTargets"][label], value, places=6)
        # the FY column covers the last four quarters only
        self.assertEqual(plan["fyPeriods"], fixture.QUARTERS[2:])
        self.assertEqual(info["ignoredColumns"], ["YoY %"])
        self.assertTrue(any("Gross Margin %" in note for note in info["skippedTargets"]))

    def test_classic_four_section_plan_keeps_classic_targets_and_months(self):
        built = fixture.build_pnl(self.path)
        plan, _ = draft_plan(self.path, intent="what-if", connection="c", folder="f")
        self.assertEqual([(s["name"], s["sign"]) for s in plan["sections"]],
                         [("Revenue", 1), ("Cost of Revenue", -1), ("Operating Expenses", -1),
                          ("Below the Line", -1)])
        self.assertEqual(list(plan["parityTargets"]), ["Revenue", "Gross Profit", "EBITDA", "Net Income"])
        self.assertEqual(plan["actualMonths"][0], {"month": "Jan-26", "num": 1, "end": "2026-01-31"})
        self.assertEqual([p["label"] for p in plan["periods"]], fixture.MONTHS)
        self.assertNotIn("fyPeriods", plan)
        self.assertEqual([d["label"] for d in plan["derivedRows"]], ["Gross Profit", "EBITDA", "Net Income"])

    def test_a_renamed_section_is_accepted_with_its_sign_from_the_derived_row(self):
        def rename(pnl, gl, drivers):
            for row in range(1, pnl.max_row + 1):
                if pnl.cell(row, 1).value == "Below the Line":
                    pnl.cell(row, 1, "Capex")
        fixture.build_pnl(self.path, mutate=rename)
        plan, _ = draft_plan(self.path, intent="what-if", connection="c", folder="f")
        capex = next(s for s in plan["sections"] if s["name"] == "Capex")
        self.assertEqual((capex["sign"], capex["signSource"]), (-1, "derived-rows"))
        self.assertEqual(plan["draftReviewIssues"][-1]["code"], "FLAT_CARRY_POLICY")

    def test_sign_from_the_name_table_is_flagged_for_review(self):
        sections = [s for s in fixture.QUARTERLY_SECTIONS if s[0] != "Other Income/Expense"]
        derived = [("Gross Margin", "{Net Sales}-{COGS}")]
        plan, _, _ = self.quarterly(sections=sections, derived=derived)
        sga = next(s for s in plan["sections"] if s["name"] == "SG&A")
        self.assertEqual((sga["sign"], sga["signSource"]), (-1, "name"))
        issue = next(i for i in plan["draftReviewIssues"] if i["code"] == "SECTION_SIGN_FROM_NAME")
        self.assertEqual(issue["evidence"], ["SG&A: -"])

    def test_ambiguous_section_with_no_deciding_derived_row_fails_naming_it(self):
        self.quarterly_fails(r"section 'Other Income/Expense': no derived row uses it.*both revenue and cost",
                             derived=[("Gross Margin", "{Net Sales}-{COGS}")])

    def test_derived_rows_that_disagree_on_a_sign_fail_naming_the_section(self):
        self.quarterly_fails(r"section 'COGS' has ambiguous sign.*'Gross Margin' \(-\).*'Odd' \(\+\)",
                             derived=[("Gross Margin", "{Net Sales}-{COGS}"), ("Odd", "{Net Sales}+{COGS}")])

    def test_a_section_weighted_other_than_one_fails(self):
        self.quarterly_fails(r"section 'COGS': derived row 'Half' weights it by -0\.5",
                             derived=[("Half", "{Net Sales}-0.5*{COGS}")])

    def test_total_and_derived_cells_are_verified_against_the_workbook(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        cogs, profit = built["rows"]["Total COGS"], built["rows"]["Operating Profit"]
        self.quarterly_fails(rf"Budget!C{cogs} \('Total COGS', Q2 FY25\)",
                             cached_overrides={("Budget", f"C{cogs}"): 1.0})
        self.quarterly_fails(rf"Budget!G{profit} \('Operating Profit', Q2 FY26 Bud\)",
                             cached_overrides={("Budget", f"G{profit}"): 1.0})

    def test_derived_row_formula_must_be_the_same_in_every_period(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        row, cogs = built["rows"]["Gross Margin"], built["rows"]["Total COGS"]

        def drift(report, gl, drivers):
            report[f"F{row}"] = f"=F{built['rows']['Total Net Sales']}-F{cogs - 1}"
        self.quarterly_fails(rf"Budget!F{row} \('Gross Margin'\).*differs from", mutate=drift)

    def test_derived_row_reading_outside_its_column_fails(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        row = built["rows"]["Gross Margin"]

        def wander(report, gl, drivers):
            for col in "BCDEFG":
                report[f"{col}{row}"] = f"=A{built['rows']['Total COGS']}"
        self.quarterly_fails(r"Gross Margin.*neither a SUMIFS actuals line nor a recognised same-column",
                             mutate=wander)


# ---------------------------------------------------------------- periods
class TestPeriods(Base):
    def test_quarter_periods_roles_and_missing_end_dates(self):
        plan, info, _ = self.quarterly()
        self.assertEqual([(p["label"], p["role"], p["grain"], p["end"]) for p in plan["periods"]],
                         [(q, "actual" if i < 4 else "forecast", "quarter", None)
                          for i, q in enumerate(fixture.QUARTERS)])
        self.assertNotIn("actualMonths", plan)
        self.assertNotIn("forecastMonths", plan)
        self.assertEqual(info["actualRange"], ("Q1 FY25", "Q4 FY25"))
        self.assertEqual(info["forecastRange"], ("Q1 FY26 Bud", "Q2 FY26 Bud"))

    def test_calendar_quarters_carry_end_dates_and_gl_text_matches(self):
        labels = ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025", "Q1 2026", "Q2 2026"]
        plan, _, _ = self.quarterly(labels=labels)
        self.assertEqual([p["end"] for p in plan["periods"]],
                         ["2025-03-31", "2025-06-30", "2025-09-30", "2025-12-31", "2026-03-31", "2026-06-30"])
        self.assertEqual({a["month"] for a in plan["actuals"]}, set(labels[:4]))

    def test_quarter_headers_must_be_consecutive(self):
        labels = ["Q1 FY25", "Q2 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26"]
        self.quarterly_fails(r"quarter headers are not consecutive quarters \(2025-Q2 then 2025-Q4\)",
                             labels=labels)

    def test_gl_period_outside_the_actual_periods_fails(self):
        def stray(report, gl, drivers):
            gl["A2"] = "Q9 FY99"
        self.quarterly_fails(r"GL Detail!row 2: period 'Q9 FY99' is not an actual period", mutate=stray)

    def test_week_ordinal_headers_make_week_periods(self):
        labels = [f"Wk {i}" for i in range(1, 7)]
        plan, info, built = self.quarterly(labels=labels)
        self.assertEqual({p["grain"] for p in plan["periods"]}, {"week"})
        self.assertEqual([p["label"] for p in plan["periods"]], labels)
        self.assertEqual([p["end"] for p in plan["periods"]], [None] * 6)
        self.assertEqual(plan["fyPeriods"], labels[2:])
        out = self.dir / "build"
        code, stdout, _ = run_main(BUILD, [str(self.write_plan(plan)), "--out-dir", str(out)])
        self.assertEqual(code, 0)
        self.assertIn("PARITY PASS Pre-tax Income:", stdout)

    def test_week_start_dates_are_week_periods_with_end_dates(self):
        from datetime import datetime, timedelta
        start = datetime(2026, 1, 5)
        stamps = {q: start + timedelta(days=7 * i) for i, q in enumerate(fixture.QUARTERS)}

        def dates(report, gl, drivers):
            for i, q in enumerate(fixture.QUARTERS):
                report.cell(4, 2 + i, stamps[q])
            for row in range(2, gl.max_row + 1):
                gl.cell(row, 1, stamps[gl.cell(row, 1).value])
        plan, _, built = self.quarterly(mutate=dates)
        self.assertEqual([p["label"] for p in plan["periods"]],
                         [stamps[q].date().isoformat() for q in fixture.QUARTERS])
        self.assertEqual(plan["periods"][0]["end"], "2026-01-11")
        self.assertEqual({p["grain"] for p in plan["periods"]}, {"week"})
        self.assertEqual({a["month"] for a in plan["actuals"]}, {p["label"] for p in plan["periods"][:4]})

    def test_week_ordinals_must_be_consecutive(self):
        labels = ["Wk 1", "Wk 2", "Wk 4", "Wk 5", "Wk 6", "Wk 7"]
        self.quarterly_fails(r"week headers are not consecutive weeks \(week 2 then week 4\)", labels=labels)

    def test_parse_quarter_variants(self):
        self.assertEqual(parse_quarter("Q3 FY25"), (2025, 3, True))
        self.assertEqual(parse_quarter("Q3 FY26 Bud"), (2026, 3, True))
        self.assertEqual(parse_quarter("Q2 2026"), (2026, 2, False))
        self.assertEqual(parse_quarter("2026 Q1"), (2026, 1, False))
        self.assertEqual(parse_quarter("FY26 Q4"), (2026, 4, True))
        for text in ("Q5 2026", "Jan-26", "Q1", "FY26 Total", None, 5):
            self.assertIsNone(parse_quarter(text))


# ---------------------------------------------------------------- forecast cell kinds
class TestForecastKinds(Base):
    def line(self, plan, name):
        return next(l for l in plan["lines"] if l["line"] == name)

    def test_year_over_year_growth_reads_the_same_period_last_year(self):
        plan, _, _ = self.quarterly()
        line = self.line(plan, "In-Store")
        self.assertEqual((line["kind"], line["lag"], line["rate"]), ("yoy", 4, 0.04))
        self.assertNotIn("juneBase", line)
        self.assertEqual(line["driverCell"], "Drivers!B4")

    def test_alternative_growth_syntax_is_growth_from_the_last_actual(self):
        plan, _, _ = self.quarterly()
        line = self.line(plan, "Online")
        self.assertEqual((line["kind"], line["juneBase"], line["rate"]), ("growth", 480.0, 0.10))

    def test_percent_of_total_names_the_base_section(self):
        plan, _, _ = self.quarterly()
        line = self.line(plan, "Merchandise")
        self.assertEqual((line["kind"], line["of"], line["rate"]), ("pct_of_total", "Net Sales", 0.6))
        self.assertEqual(line["driverCell"], "Drivers!B6")

    def test_external_link_is_a_manual_line_with_a_review_issue(self):
        plan, _, built = self.quarterly()
        line = self.line(plan, "Allocation")
        self.assertEqual((line["kind"], line["manual"]), ("manual", [61.0, 64.0]))
        issue = next(i for i in plan["draftReviewIssues"] if i["code"] == "EXTERNAL_LINK_VALUE")
        row = built["rows"]["Allocation"]
        (evidence,) = issue["evidence"]
        self.assertIn(f"Budget!F{row}", evidence)
        self.assertIn(f"Budget!G{row}", evidence)
        self.assertIn("[Alloc.xlsx]Alloc", evidence)

    def test_zero_actual_line_has_no_gl_rows_and_is_accepted(self):
        plan, _, _ = self.quarterly()
        self.assertFalse([a for a in plan["actuals"] if a["line"] == "Allocation"])

    def test_nonzero_literal_actual_fails(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        row = built["rows"]["Payroll"]

        def literal(report, gl, drivers):
            report[f"C{row}"] = 5
        self.quarterly_fails(rf"Budget!C{row} \('Payroll'\): row is neither a SUMIFS actuals line", mutate=literal)

    def test_year_over_year_lag_must_be_uniform(self):
        probe = self.dir / "probe.xlsx"
        row = fixture.build_quarterly(probe)["rows"]["In-Store"]

        def drift(report, gl, drivers):
            report[f"G{row}"] = f"=D{row}*(1+Drivers!$B$4)"
        self.quarterly_fails(rf"Budget!G{row} \('In-Store'\): forecast cells are not uniform", mutate=drift)

    def test_percent_of_its_own_section_total_fails(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        row, own = built["rows"]["Merchandise"], built["rows"]["Total COGS"]

        def own_total(report, gl, drivers):
            for col in "FG":
                report[f"{col}{row}"] = f"={col}{own}*Drivers!$B$6"
        self.quarterly_fails(rf"Budget!F{row} \('Merchandise'\): percentage base row {own} is not the total",
                             mutate=own_total)

    def test_percent_of_a_line_row_fails(self):
        probe = self.dir / "probe.xlsx"
        built = fixture.build_quarterly(probe)
        row, other = built["rows"]["Merchandise"], built["rows"]["In-Store"]

        def line_base(report, gl, drivers):
            for col in "FG":
                report[f"{col}{row}"] = f"={col}{other}*Drivers!$B$6"
        self.quarterly_fails(r"percentage base row \d+ is not the total", mutate=line_base)

    def test_a_percentage_base_section_that_holds_percentage_lines_is_circular(self):
        plan, _, _ = self.quarterly()
        plan["lines"][0] = {**plan["lines"][0], "kind": "pct_of_total", "of": "COGS", "rate": 0.1}
        with self.assertRaisesRegex(calc.ExprError, "percentage base 'COGS' must be another section"):
            calc.Evaluation(plan)

    def test_rates_may_come_from_more_than_one_driver_column(self):
        probe = self.dir / "probe.xlsx"
        row = fixture.build_quarterly(probe)["rows"]["Payroll"]

        def second_column(report, gl, drivers):
            for col in "FG":
                prev = chr(ord(col) - 1)
                report[f"{col}{row}"] = f"={prev}{row}+{prev}{row}*Drivers!$D$8"
            drivers["D8"] = 0.02
        plan, _, _ = self.quarterly(mutate=second_column)
        self.assertEqual(self.line(plan, "Payroll")["driverCell"], "Drivers!D8")
        self.assertEqual(len(plan["assumptions"]), 6)

    def test_recomputed_plan_reproduces_the_fy_targets(self):
        plan, _, built = self.quarterly()
        ev = calc.Evaluation(plan)
        fy = [i for i, label in enumerate(ev.labels) if label in plan["fyPeriods"]]
        for label in ("Gross Margin", "Operating Profit", "Pre-tax Income"):
            self.assertAlmostEqual(ev.aggregate(label, fy), built["targets"][label], places=6)
        self.assertAlmostEqual(ev.section_total("COGS", fy), built["targets"]["Total COGS"], places=6)
        # percent-of-total lines are a share of the forecast Net Sales total, not the actuals
        merch = ev.matrix["Merchandise"]
        self.assertAlmostEqual(merch[4], ev.totals["Net Sales"][4] * 0.6, places=9)

    def test_draft_summary_names_sections_derived_rows_and_verified_cells(self):
        _, info, _ = self.quarterly()
        self.assertEqual(info["grain"], "quarter")
        self.assertIn("Net Sales (+, derived-rows)", info["sections"])
        self.assertEqual(info["verifiedCells"], (9 + 4 + 4) * 6)

    def test_cli_prints_the_general_shape(self):
        fixture.build_quarterly(self.path)
        code, stdout, _ = run_main(DRAFT, [str(self.path), "--connection", "c", "--folder", "f",
                                           "--out", str(self.dir / "plan.json")])
        self.assertEqual(code, 0)
        self.assertIn("9 lines (flat 1, growth 2, manual 3, pct_of_total 2, yoy 1)", stdout)
        self.assertIn("periods: Q1 FY25-Q4 FY25 actual", stdout)
        self.assertIn("sections: Net Sales (+, derived-rows); COGS (-, derived-rows)", stdout)
        self.assertIn("derived rows: Gross Margin, Gross Margin %, Operating Profit, Pre-tax Income", stdout)
        self.assertIn("FY column H over 4 of the periods", stdout)
        self.assertIn("columns not carried (outside the periods and FY total): YoY %", stdout)


# ---------------------------------------------------------------- builder
class TestBuilderGeneral(Base):
    def build(self, plan, *argv):
        plan_path = self.write_plan(plan)
        out = self.dir / "build"
        code, stdout, stderr = run_main(BUILD, [str(plan_path), "--out-dir", str(out), *argv])
        return code, stdout, stderr, out

    def test_dry_run_checks_every_total_and_derived_row(self):
        plan, _, built = self.quarterly()
        code, stdout, _, out = self.build(plan)
        self.assertEqual(code, 0)
        for label, value in built["targets"].items():
            self.assertIn(f"PARITY PASS {label}: plan={value:,.2f} source={value:,.2f}", stdout)
        self.assertTrue((out / "wb_spec.json").is_file())

    def test_parity_failure_on_a_generic_row_stops_before_writing(self):
        plan, _, _ = self.quarterly()
        plan["parityTargets"]["Operating Profit"] += 5
        code, stdout, stderr, out = self.build(plan)
        self.assertEqual(code, 1)
        self.assertIn("PARITY FAIL Operating Profit:", stdout)
        self.assertIn("LOCAL PARITY FAILED", stderr)
        self.assertFalse(out.exists())

    def test_a_target_with_no_matching_row_fails(self):
        plan, _, _ = self.quarterly()
        plan["parityTargets"]["Mystery"] = 1.0
        code, stdout, _, _ = self.build(plan)
        self.assertEqual(code, 1)
        self.assertIn("PARITY FAIL Mystery: plan has no row to recompute it", stdout)

    def test_parity_sums_only_the_fy_periods(self):
        plan, _, built = self.quarterly()
        computed = BUILD.parity(plan)
        self.assertAlmostEqual(computed["Total Net Sales"], built["targets"]["Total Net Sales"], places=6)
        without = {k: v for k, v in plan.items() if k != "fyPeriods"}
        self.assertGreater(BUILD.parity(without)["Total Net Sales"], computed["Total Net Sales"])

    def test_dm_spine_carries_yoy_bases_percent_rows_and_null_dates(self):
        plan, _, _ = self.quarterly()
        dm = BUILD.build_dm_spec(plan)
        sql = {e["id"]: e["source"]["statement"] for e in dm["pages"][0]["elements"]}
        self.assertIn("OF_SECTION", sql["spine"])
        # In-Store is year-over-year: Q1 FY26 Bud reads Q1 FY25 (1000.00) with exponent 1
        self.assertIn("('In-Store','Net Sales',1,1,'Q1 FY26 Bud',5,NULL::date,1,'growth',0.04,NULL,1000.00,NULL)",
                      sql["spine"])
        self.assertIn("('In-Store','Net Sales',1,1,'Q2 FY26 Bud',6,NULL::date,1,'growth',0.04,NULL,1100.00,NULL)",
                      sql["spine"])
        self.assertIn("'Merchandise','COGS',2,3,'Q1 FY26 Bud',5,NULL::date,1,'pct',0.6,NULL,NULL,'Net Sales'",
                      sql["spine"])
        self.assertIn("(NULL::date,'Q1 FY25',1,NULL::date,'In-Store','Net Sales'", sql["glDetail"])

    def test_percent_lines_get_a_second_pass_over_forecast_section_totals(self):
        plan, _, _ = self.quarterly()
        spec = BUILD.build_wb_spec(plan, "<dm>", ELS)
        elements = {e["id"]: e for p in spec["pages"] for e in p["elements"]}
        self.assertIn("forecastCalc", elements)
        totals = elements["forecastTotals"]
        self.assertEqual(totals["groupings"], [{"id": "bySectionMonth", "groupBy": ["t_mn", "t_sec"],
                                                "calculations": ["t_amt"]}])
        final = elements["forecast"]
        self.assertEqual(final["source"]["primarySource"], {"kind": "table", "elementId": "forecastCalc"})
        self.assertEqual(final["source"]["joins"][0]["columns"],
                         [{"left": "[Month Num]", "right": "[Month Num]"},
                          {"left": "[Of Section]", "right": "[Section]"}])
        amount = next(c for c in final["columns"] if c["name"] == "Amount")["formula"]
        self.assertIn('If([PJ/Kind] = "pct", [PJ/Eff Rate] * Coalesce([Totals/Section Amount], 0)', amount)
        base_amount = next(c for c in elements["forecastCalc"]["columns"] if c["name"] == "Amount")["formula"]
        self.assertIn('[Kind] = "pct", 0', base_amount)
        self.assertIn('elementId="forecastCalc"', spec["layout"])
        self.assertIn('elementId="forecastTotals"', spec["layout"])

    def test_plans_without_percent_lines_keep_one_forecast_table(self):
        fixture.build_pnl(self.path)
        plan, _ = draft_plan(self.path, intent="what-if", connection="c", folder="f")
        spec = BUILD.build_wb_spec(plan, "<dm>", ELS)
        ids = {e["id"] for p in spec["pages"] for e in p["elements"]}
        self.assertNotIn("forecastCalc", ids)
        self.assertIn("forecast", ids)

    def test_sections_other_than_the_classic_four_get_amount_columns_and_formulas(self):
        plan, _, _ = self.quarterly()
        spec = BUILD.build_wb_spec(plan, "<dm>", ELS)
        elements = {e["id"]: e for p in spec["pages"] for e in p["elements"]}
        names = [c["name"] for c in elements["pnlCombined"]["columns"]]
        self.assertEqual(names[-4:], ["Net Sales Amt", "COGS Amt", "SG&A Amt", "Other Income Expense Amt"])
        sga = next(c for c in elements["pnlCombined"]["columns"] if c["name"] == "SG&A Amt")
        self.assertIn('[Union of 2 Sources/Section] = "SG&A"', sga["formula"])
        self.assertEqual([elements[k]["name"] for k in ("kpiRev", "kpiEbitda", "kpiNI", "kpiGM")],
                         ["Total Net Sales", "Operating Profit", "Pre-tax Income", "Gross Margin %"])
        self.assertEqual(elements["kpiEbitda"]["columns"][0]["formula"],
                         "Sum([PnL Combined/Net Sales Amt]) - Sum([PnL Combined/COGS Amt]) - "
                         "Sum([PnL Combined/SG&A Amt])")
        self.assertEqual(elements["monthlyTrend"]["name"], "Quarterly Net Sales, Operating Profit & Pre-tax Income")

    def test_quarter_labels_order_the_trend_and_pivot_by_period_number(self):
        plan, _, _ = self.quarterly()
        spec = BUILD.build_wb_spec(plan, "<dm>", ELS)
        elements = {e["id"]: e for p in spec["pages"] for e in p["elements"]}
        trend = elements["monthlyTrend"]
        self.assertEqual(trend["xAxis"]["sort"], {"direction": "ascending", "by": "t_ord", "aggregation": "min"})
        pivot = elements["pnlPivot"]
        self.assertEqual(pivot["columnsBy"][1]["sort"],
                         {"direction": "ascending", "by": "p_mn", "aggregation": "min"})
        self.assertEqual(len({c["name"] for c in pivot["columns"]}), len(pivot["columns"]))

    def test_seed_csvs_carry_rates_for_every_driver_kind_and_blank_dates(self):
        plan, _, _ = self.quarterly()
        out = self.dir / "seed"
        out.mkdir()
        rate, manual = BUILD.seed_csvs(plan, str(out))
        rates = Path(rate).read_text().splitlines()
        self.assertEqual(rates[0], "LINE_ITEM,RATE")
        self.assertEqual({r.split(",")[0] for r in rates[1:]},
                         {"In-Store", "Online", "Merchandise", "Shrink", "Payroll"})
        self.assertIn("Marketing,Q1 FY26 Bud,5,,90.00", Path(manual).read_text())

    def test_post_needs_dispositions_for_the_external_link(self):
        plan, _, _ = self.quarterly()
        plan_path = self.write_plan(plan)
        with mock.patch.object(BUILD, "env", side_effect=AssertionError("auth reached")):
            code, _, stderr = run_main(BUILD, [str(plan_path), "--post"])
        self.assertEqual(code, 1)
        self.assertIn("EXTERNAL_LINK_VALUE", stderr)

    def test_classic_drafted_plan_builds_exactly_like_its_legacy_form(self):
        fixture.build_pnl(self.path)
        plan, _ = draft_plan(self.path, intent="what-if", connection="c", folder="f")
        legacy = {k: v for k, v in plan.items() if k not in ("periods", "sections", "derivedRows")}
        self.assertEqual(BUILD.build_dm_spec(plan), BUILD.build_dm_spec(legacy))
        self.assertEqual(BUILD.build_wb_spec(plan, "<dm>", ELS), BUILD.build_wb_spec(legacy, "<dm>", ELS))
        self.assertEqual(BUILD.parity(plan)["Net Income"], BUILD.parity(legacy)["Net Income"])


# ---------------------------------------------------------------- 03 Retail Quarterly Budget
class TestRetailQuarterlyBudget(Base):
    def draft(self):
        plan_path = self.dir / "retail.json"
        code, stdout, stderr = run_main(DRAFT, [str(RETAIL), "--connection", "conn", "--folder", "fold",
                                                "--out", str(plan_path)])
        self.assertEqual(code, 0, stderr)
        return json.loads(plan_path.read_text()), stdout, plan_path

    def test_drafts_and_verifies_every_cell(self):
        plan, stdout, _ = self.draft()
        self.assertIn("10 lines (flat 2, growth 1, manual 3, pct_of_total 2, yoy 2)", stdout)
        self.assertIn("124 GL rows (GL Detail)", stdout)
        self.assertIn("periods: Q3 FY25-Q2 FY26 actual, Q3 FY26 Bud-Q4 FY26 Bud forecast", stdout)
        self.assertIn("verified: 102 line / total / derived-row cells", stdout)
        self.assertIn("source chart mapped + values verified: Net Sales by Quarter (period, 1 series)", stdout)
        self.assertEqual([s["name"] for s in plan["sections"]],
                         ["Net Sales", "COGS", "SG&A", "Other Income/Expense"])
        self.assertEqual([s["sign"] for s in plan["sections"]], [1, -1, -1, -1])
        self.assertEqual([d["label"] for d in plan["derivedRows"]],
                         ["Gross Margin", "Operating Profit", "Pre-tax Income"])
        self.assertEqual(plan["fyPeriods"], ["Q1 FY26", "Q2 FY26", "Q3 FY26 Bud", "Q4 FY26 Bud"])
        kinds = {l["line"]: l["kind"] for l in plan["lines"]}
        self.assertEqual(kinds, {
            "In-Store Sales": "yoy", "E-Commerce Sales": "yoy", "Merchandise Cost": "pct_of_total",
            "Shrink": "pct_of_total", "Store Payroll": "growth", "Marketing": "manual",
            "Occupancy": "flat", "Corporate Overhead Allocation": "manual", "Interest Income": "manual",
            "Depreciation": "flat"})
        self.assertEqual({i["code"] for i in plan["draftReviewIssues"]},
                         {"FLAT_CARRY_POLICY", "EXTERNAL_LINK_VALUE"})
        (chart,) = plan["sourceCharts"]
        self.assertEqual((chart["kind"], chart["grain"], chart["series"], chart["verifiedPoints"]),
                         ("bar", "period", ["Total Net Sales"], 6))

    def test_parity_targets_match_the_workbook_fy_column(self):
        plan, _, _ = self.draft()
        targets = plan["parityTargets"]
        self.assertAlmostEqual(targets["Total Net Sales"], 14402043.63, places=2)
        self.assertAlmostEqual(targets["Total COGS"], 9055995.86, places=2)
        self.assertAlmostEqual(targets["Operating Profit"], 1377182.36, places=2)
        self.assertAlmostEqual(targets["Pre-tax Income"], 1218182.36, places=2)
        self.assertEqual(set(targets), {"Total Net Sales", "Total COGS", "Gross Margin", "Total SG&A",
                                        "Total Other Income/Expense", "Operating Profit", "Pre-tax Income"})

    def test_local_build_passes_parity_on_every_derived_row(self):
        plan, _, plan_path = self.draft()
        out = self.dir / "build"
        code, stdout, _ = run_main(BUILD, [str(plan_path), "--out-dir", str(out)])
        self.assertEqual(code, 0)
        for fragment in ("PARITY PASS Total Net Sales: plan=14,402,043.63",
                         "PARITY PASS Total COGS: plan=9,055,995.86",
                         "PARITY PASS Gross Margin: plan=5,346,047.77",
                         "PARITY PASS Total SG&A: plan=3,968,865.41",
                         "PARITY PASS Total Other Income/Expense: plan=159,000.00",
                         "PARITY PASS Operating Profit: plan=1,377,182.36",
                         "PARITY PASS Pre-tax Income: plan=1,218,182.36"):
            self.assertIn(fragment, stdout)
        self.assertIn("EXTERNAL_LINK_VALUE", stdout)
        wb = json.loads((out / "wb_spec.json").read_text())
        ids = {e["id"] for e in wb["contents"]["elements"]}
        self.assertTrue({"forecastCalc", "forecastTotals", "forecast", "sourceChart1"} <= ids)
        bar = next(e for e in wb["contents"]["elements"] if e["id"] == "sourceChart1")
        self.assertEqual(bar["kind"], "bar-chart")
        self.assertEqual(bar["columns"][0]["formula"], "[PnL Combined/Month]")

    def test_a_cell_that_disagrees_with_the_plan_fails_the_draft_with_its_address(self):
        import shutil
        import zipfile
        bad = self.dir / "bad.xlsx"
        shutil.copy(RETAIL, bad)
        # Tamper with the cached Operating Profit for Q3 FY26 Bud (F29).
        with zipfile.ZipFile(bad) as source:
            names = {n: source.read(n) for n in source.namelist()}
        sheet = next(n for n in names if n.startswith("xl/worksheets/sheet1."))
        text = names[sheet].decode("utf-8")
        import re
        pattern = re.compile(r'(<c r="F29"[^>]*><f>[^<]*</f><v>)[^<]*(</v>)')
        self.assertRegex(text, pattern)
        names[sheet] = pattern.sub(r"\g<1>1\g<2>", text).encode("utf-8")
        with zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as target:
            for name, data in names.items():
                target.writestr(name, data)
        with self.assertRaisesRegex(PlanDraftError, r"Budget!F29 \('Operating Profit', Q3 FY26 Bud\)"):
            draft_plan(bad, intent="what-if", connection="c", folder="f")


# ---------------------------------------------------------------- 02 regression
class TestSaasRegression(Base):
    def test_plan_fields_match_the_pre_generalisation_golden(self):
        plan, _ = draft_plan(SAAS, intent="what-if", connection="c", folder="f")
        golden = json.loads((FIXTURES / "forecast-regression" / "02-saas-plan-fields.json").read_text())
        for field, expected in golden.items():
            self.assertEqual(json.dumps(plan[field], sort_keys=True), json.dumps(expected, sort_keys=True),
                             field)
        self.assertNotIn("fyPeriods", plan)

    def test_builds_with_the_same_four_metrics(self):
        plan, _ = draft_plan(SAAS, intent="what-if", connection="c", folder="f")
        plan_path = self.write_plan(plan)
        code, stdout, _ = run_main(BUILD, [str(plan_path), "--out-dir", str(self.dir / "b")])
        self.assertEqual(code, 0)
        for metric in ("Revenue", "Gross Profit", "EBITDA", "Net Income"):
            self.assertIn(f"PARITY PASS {metric}:", stdout)


if __name__ == "__main__":
    unittest.main()
