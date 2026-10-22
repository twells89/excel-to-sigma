#!/usr/bin/env python3
"""table-report: the oracle and the EMITTED Sigma formulas must agree with Excel.

Review findings these tests pin down (each was a silent wrong conversion): blank cells (Excel 0 / "",
Sigma NULL), expansion pivots that merge rows, ``Mod`` signs, ``IFERROR`` scope, ``XLOOKUP`` defaults for
blank values, case-insensitive text / wildcard labels, and a loose 0.01 tolerance. The core evidence is
``sigma_formula_eval``: it evaluates the emitted formula strings (SQL three-valued NULL logic) over the
plan's rows and compares every verified cell with Excel's cached value, not just with the oracle.
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "scripts" / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import sigma_formula_eval as se  # noqa: E402
import table_report as tr  # noqa: E402
import table_report_fixture as fx  # noqa: E402
import table_report_model as trm  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "broader-eight"


def draft(path):
    return tr.draft_table_report(path, intent="read-only", connection="c", folder="f")


def emitted_problems(test, plan):
    """Every verified cell, evaluated through the emitted formulas under BOTH Mod conventions."""
    results = {}
    for mod in ("floor", "trunc"):
        compared, skipped, problems = se.compare_to_workbook(plan, mod)
        test.assertEqual(problems, [], f"emitted formulas disagree with the workbook (Mod={mod})")
        results[mod] = (compared, skipped)
    return results["floor"]


def formulas(plan, element_name):
    parts = trm.build_parts(plan, "dm", trm.local_dm_elements(plan))
    element = next(e for g in ("hidden", "visible") for e in parts[g] if e["name"] == element_name)
    return {c["name"]: c.get("formula") for c in element["columns"]}


class WorkbookCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def path(self, name="book.xlsx"):
        return self.dir / name


# ---------------------------------------------------------------------------- 1. blanks
ROWS = [("East", "Open", 100, 5), ("East", None, 40, None), ("East", "Closed", 70, 1),
        ("West", "Open", 30, 2), ("West", None, 80, 3)]


def build_blanks(path):
    """Blank Status / Incoming cells read through arithmetic, IF, COUNTIFS "<>x" and SUMIFS."""
    book = fx._Book()
    ws = book.sheet("Data")
    book.table(ws, "tblD", ["Region", "Status", "Amt", "Incoming", "Avail", "Flag"],
               [list(r) + [None, None] for r in ROWS])
    this = lambda c: f"tblD[[#This Row],[{c}]]"  # noqa: E731
    for i, (_, status, amt, incoming) in enumerate(ROWS, 2):
        book.put(ws, f"E{i}", f"={this('Amt')}+{this('Incoming')}", amt + (incoming or 0))
        book.put(ws, f"F{i}", f'=IF({this("Status")}<>"Closed","Live","Done")',
                 "Live" if status != "Closed" else "Done")
    summary = book.sheet("Summary")
    summary["A1"] = "By region"
    for c, head in enumerate(["Region", "Not closed", "Amt not closed", "Avail", "Open"], 1):
        summary.cell(3, c, head)
    for r, region in enumerate(["East", "West"], 4):
        summary.cell(r, 1, region)
        mine = [x for x in ROWS if x[0] == region]
        book.put(summary, f"B{r}", f'=COUNTIFS(tblD[Region],$A{r},tblD[Status],"<>Closed")',
                 sum(1 for x in mine if x[1] != "Closed"))
        book.put(summary, f"C{r}", f'=SUMIFS(tblD[Amt],tblD[Region],$A{r},tblD[Status],"<>Closed")',
                 sum(x[2] for x in mine if x[1] != "Closed"))
        book.put(summary, f"D{r}", f"=SUMIFS(tblD[Avail],tblD[Region],$A{r})",
                 sum(x[2] + (x[3] or 0) for x in mine))
        book.put(summary, f"E{r}", f'=COUNTIFS(tblD[Region],$A{r},tblD[Status],"Open")',
                 sum(1 for x in mine if x[1] == "Open"))
    book.save(path)


class TestBlankSemantics(WorkbookCase):
    def test_blanks_follow_excel_in_the_oracle_and_the_emitted_formulas(self):
        path = self.path()
        build_blanks(path)
        plan, summary = draft(path)
        compared, skipped = emitted_problems(self, plan)
        self.assertEqual((compared, skipped), (summary["verifiedCells"], 0))

    def test_blank_capable_operands_are_coalesced_where_excel_coerces_them(self):
        path = self.path()
        build_blanks(path)
        plan, _ = draft(path)
        calc = formulas(plan, "tblD")
        self.assertEqual(calc["Avail"], "([Amt] + Coalesce([Incoming], 0))")
        self.assertEqual(calc["Flag"], 'If((Coalesce([Status], "") != "Closed"), "Live", "Done")')
        grouped = formulas(plan, "By region")
        self.assertEqual(grouped["Not closed"],
                         'Coalesce(CountIf((IsNull([tblD/Status]) or [tblD/Status] != "Closed")), 0)')
        # a column with no blank keeps its plain formula (no needless wrappers)
        self.assertNotIn("Coalesce([tblD/Amt]", grouped["Amt not closed"])
        self.assertIn("SumIf([tblD/Amt],", grouped["Amt not closed"])

    def test_the_evaluator_catches_the_unfixed_emission(self):
        """Guard against a vacuous evaluator: the pre-fix formulas give the wrong answers."""
        ev = se.Evaluator(lambda *a: None)
        row = {"Amt": 40, "Incoming": None, "Status": None}
        ctx = se.Ctx(lambda name: row[name])
        self.assertIsNone(ev.ev(se.parse("([Amt] + [Incoming])"), ctx))                       # was emitted
        self.assertEqual(ev.ev(se.parse("([Amt] + Coalesce([Incoming], 0))"), ctx), 40)       # Excel
        self.assertEqual(ev.ev(se.parse('If(([Status] != "Closed"), "Live", "Done")'), ctx), "Done")
        self.assertEqual(ev.ev(se.parse('If((Coalesce([Status], "") != "Closed"), "Live", "Done")'), ctx), "Live")

    def test_min_and_max_over_a_blank_cell_are_refused(self):
        """Excel's MIN skips a blank reference; Sigma's Least does not: not modelled, so loud."""
        book = fx._Book()
        ws = book.sheet("Data")
        book.table(ws, "tblM", ["Key", "A", "B", "Lo"], [["x", 5, None, None], ["y", 3, 4, None]])
        book.put(ws, "D2", "=MIN(tblM[[#This Row],[A]],tblM[[#This Row],[B]])", 5)
        book.put(ws, "D3", "=MIN(tblM[[#This Row],[A]],tblM[[#This Row],[B]])", 3)
        summary = book.sheet("S")
        summary["A1"] = "Lows"
        summary.cell(3, 1, "Key")
        summary.cell(3, 2, "N")
        for r, key in enumerate(["x", "y"], 4):
            summary.cell(r, 1, key)
            book.put(summary, f"B{r}", f"=COUNTIFS(tblM[Key],$A{r})", 1)
        book.save(self.path())
        with self.assertRaisesRegex(tr.PlanDraftError, r"MIN over a blank cell is not supported"):
            draft(self.path())

    def test_a_formula_returning_a_blank_cell_is_zero(self):
        """=XLOOKUP(..) that lands on a blank cell shows 0 in Excel; the column must be 0, not NULL."""
        book = fx._Book()
        rates = book.sheet("Rates")
        book.table(rates, "tblRates", ["Product", "Rate"], [["Widget", 0.1], ["Gadget", None]])
        orders = book.sheet("Orders")
        book.table(orders, "tblO", ["Product", "Qty", "Rate", "Net"],
                   [["Widget", 10, None, None], ["Gadget", 10, None, None], ["Gizmo", 10, None, None]])
        this = lambda c: f"tblO[[#This Row],[{c}]]"  # noqa: E731
        for i, (product, rate) in enumerate([("Widget", 0.1), ("Gadget", 0), ("Gizmo", -1)], 2):
            book.put(orders, f"C{i}", f"=_xlfn.XLOOKUP({this('Product')},tblRates[Product],tblRates[Rate],-1)", rate)
            book.put(orders, f"D{i}", f"={this('Qty')}*(1+{this('Rate')})", 10 * (1 + rate))
        summary = book.sheet("S")
        summary["A1"] = "Net"
        summary.cell(3, 1, "Product")
        summary.cell(3, 2, "Net")
        for r, product in enumerate(["Widget", "Gadget", "Gizmo"], 4):
            summary.cell(r, 1, product)
            book.put(summary, f"B{r}", f"=SUMIFS(tblO[Net],tblO[Product],$A{r})",
                     {"Widget": 11.0, "Gadget": 10.0, "Gizmo": 0.0}[product])
        book.save(self.path())
        plan, _ = draft(self.path())
        emitted_problems(self, plan)
        # the default must not hide a key that IS found with a blank value: found-ness is tested on the key
        rate = formulas(plan, "tblO")["Rate"]
        self.assertIn("If(IsNull(Lookup([tblRates/Product], [Product], [tblRates/Product])), (-1), "
                      "Lookup([tblRates/Rate], [Product], [tblRates/Product]))", rate)
        self.assertTrue(rate.startswith("Coalesce("), rate)

    def test_numeric_criteria_never_match_a_blank_cell(self):
        """SUMIFS(.., Qty, "<5") skips blanks (but ``=A1<5`` would not): the oracle uses criteria semantics."""
        book = fx._Book()
        ws = book.sheet("Data")
        data = [("a", 10.0, 3), ("a", 20.0, None), ("a", 40.0, 9), ("b", 1.0, None), ("b", 2.0, 1)]
        book.table(ws, "tblQ", ["Grp", "Amt", "Qty"], [list(r) for r in data])
        summary = book.sheet("S")
        summary["A1"] = "Sizes"
        for c, head in enumerate(["Grp", "Small", "NotNine", "AtLeastOne"], 1):
            summary.cell(3, c, head)
        for r, grp in enumerate(["a", "b"], 4):
            summary.cell(r, 1, grp)
            mine = [x for x in data if x[0] == grp]
            book.put(summary, f"B{r}", f'=SUMIFS(tblQ[Amt],tblQ[Grp],$A{r},tblQ[Qty],"<5")',
                     sum(x[1] for x in mine if x[2] is not None and x[2] < 5))
            book.put(summary, f"C{r}", f'=SUMIFS(tblQ[Amt],tblQ[Grp],$A{r},tblQ[Qty],"<>9")',
                     sum(x[1] for x in mine if x[2] != 9))          # <> matches the blank cells
            book.put(summary, f"D{r}", f'=COUNTIFS(tblQ[Grp],$A{r},tblQ[Qty],">=1")',
                     sum(1 for x in mine if x[2] is not None and x[2] >= 1))
        book.save(self.path())
        plan, _ = draft(self.path())
        emitted_problems(self, plan)
        grouped = formulas(plan, "Sizes")
        self.assertIn("(IsNull([tblQ/Qty]) or [tblQ/Qty] != 9)", grouped["NotNine"])
        self.assertIn("([tblQ/Qty] < 5)", grouped["Small"])        # NULL < 5 is unknown -> excluded, as Excel


# ---------------------------------------------------------------------------- 2. pivot
def build_roster(path, unique_key=True, show_name=False):
    book = fx._Book()
    people = book.sheet("People")
    rows = [list(p) for p in fx.PEOPLE]
    header = ["Name", "Dept", "Salary", "Start"]
    if not unique_key:                       # nothing in the table identifies a row
        rows = [[p[1], p[1], 120000.0, date(2025, 6, 1)] for p in fx.PEOPLE]
    book.table(people, "tblPeople", header, rows)
    cost = book.sheet("Cost")
    cost["A1"] = "Cost by person"
    heads = ["#"] + (["Name"] if show_name else []) + ["Dept"]
    for c, head in enumerate(heads, 1):
        cost.cell(3, c, head)
    first_month = len(heads) + 1
    for j, month in enumerate(fx.MONTHS, first_month):
        cost.cell(3, j, month)
        cost.cell(3, j).number_format = fx.DATE_FORMAT
    for i, (name, dept, salary, start) in enumerate(fx.PEOPLE, 1):
        r = 3 + i
        cost.cell(r, 1, i)
        col = 2
        if show_name:
            book.put(cost, f"B{r}", f"=INDEX(tblPeople[Name],$A{r})", name if unique_key else dept)
            col += 1
        book.put(cost, f"{fx.get_column_letter(col)}{r}", f"=INDEX(tblPeople[Dept],$A{r})", dept)
        for j, month in enumerate(fx.MONTHS, first_month):
            letter = fx.get_column_letter(j)
            book.put(cost, f"{letter}{r}",
                     f"=IF(INDEX(tblPeople[Start],$A{r})<={letter}$3,INDEX(tblPeople[Salary],$A{r})/12,0)",
                     (salary if unique_key else 120000.0) / 12 if (start if unique_key else date(2025, 6, 1)) <= month else 0)
    book.save(path)


class TestExpansionPivotRows(WorkbookCase):
    def pivot_rows(self, plan):
        parts = trm.build_parts(plan, "dm", trm.local_dm_elements(plan))
        pivot = next(e for e in parts["visible"] if e["kind"] == "pivot-table")
        by = [r["columnId"] for r in pivot["rowsBy"]]
        return [c["name"] for c in pivot["columns"] if c["id"] in by]

    def distinct_pivot_rows(self, plan):
        """How many pivot rows the emitted expansion produces (distinct values of the row dimensions)."""
        model = se.EmittedModel(plan)
        frame = next(f for f in plan["frames"] if f["kind"] == "expansion")
        rows = model.rows_of(model.by_name[frame["title"]])
        return len({tuple(r[n] for n in self.pivot_rows(plan)) for r in rows})

    def test_a_shared_attribute_no_longer_merges_roster_rows(self):
        build_roster(self.path())
        plan, summary = draft(self.path())
        self.assertEqual(self.pivot_rows(plan), ["Name", "Dept"])        # unique key leads
        self.assertEqual(self.distinct_pivot_rows(plan), len(fx.PEOPLE))    # one pivot row per person
        self.assertIn("EXPANSION_ROW_KEY_ADDED", summary["reviewIssues"])
        issue = next(i for i in plan["draftReviewIssues"] if i["code"] == "EXPANSION_ROW_KEY_ADDED")
        self.assertIn("tblPeople[Name]", issue["message"])

    def test_attributes_that_already_identify_rows_are_unchanged(self):
        build_roster(self.path(), show_name=True)
        plan, summary = draft(self.path())
        self.assertEqual(self.pivot_rows(plan), ["Name", "Dept"])
        self.assertNotIn("EXPANSION_ROW_KEY_ADDED", summary["reviewIssues"])
        self.assertEqual(self.distinct_pivot_rows(plan), len(fx.PEOPLE))

    def test_no_attribute_columns_is_one_row_per_base_row_too(self):
        parts_plan, _ = draft(_staff(self.dir))
        self.assertEqual(self.distinct_pivot_rows(parts_plan), len(fx.PEOPLE))

    def test_a_table_with_no_unique_column_is_rejected_naming_the_grid(self):
        build_roster(self.path(), unique_key=False)
        with self.assertRaisesRegex(tr.PlanDraftError,
                                    r"Cost!B4: the columns this grid shows from tblPeople .* do not identify one row"):
            draft(self.path())


def _staff(directory):
    path = directory / "staff.xlsx"
    fx.build_staff(path)
    return path


# ---------------------------------------------------------------------------- 3. MOD
class TestModSign(WorkbookCase):
    PAIRS = [(-7, 3), (7, -3), (-7, -3), (7, 3), (6, 3), (-6, 3)]

    def build(self):
        book = fx._Book()
        ws = book.sheet("Data")
        book.table(ws, "tblN", ["Id", "A", "B", "Rem"], [[f"p{i}", a, b, None] for i, (a, b) in enumerate(self.PAIRS, 1)])
        for i, (a, b) in enumerate(self.PAIRS, 2):
            book.put(ws, f"D{i}", "=MOD(tblN[[#This Row],[A]],tblN[[#This Row],[B]])", a % b)   # Excel: sign of b
        summary = book.sheet("S")
        summary["A1"] = "Remainders"
        summary.cell(3, 1, "Id")
        summary.cell(3, 2, "Rem")
        for r, i in enumerate(range(1, len(self.PAIRS) + 1), 4):
            summary.cell(r, 1, f"p{i}")
            book.put(summary, f"B{r}", f"=SUMIFS(tblN[Rem],tblN[Id],$A{r})", self.PAIRS[i - 1][0] % self.PAIRS[i - 1][1])
        book.save(self.path())

    def test_negative_operands_agree_under_either_sql_mod_convention(self):
        self.build()
        plan, _ = draft(self.path())
        self.assertEqual(formulas(plan, "tblN")["Rem"], "Mod(Mod([A], [B]) + [B], [B])")
        compared, _ = emitted_problems(self, plan)       # floor (Excel-like) and trunc (SQL dividend sign)
        self.assertGreater(compared, len(self.PAIRS))

    def test_the_plain_form_would_have_failed_the_dividend_signed_convention(self):
        ev = se.Evaluator(lambda *a: None, mod="trunc")
        ctx = se.Ctx(lambda name: {"A": -7, "B": 3}[name])
        self.assertEqual(ev.ev(se.parse("Mod([A], [B])"), ctx), -1)                   # Excel says 2
        self.assertEqual(ev.ev(se.parse("Mod(Mod([A], [B]) + [B], [B])"), ctx), 2)


# ---------------------------------------------------------------------------- 4. IFERROR
class TestIfErrorScope(WorkbookCase):
    def build(self, text_error=False):
        book = fx._Book()
        q = book.sheet("Q")
        rows = [["a", 10, 2], ["b", 5, 0], ["zz", 4, 2]]
        book.table(q, "tblQ", ["Item", "Qty", "Div", "Scaled", "Guarded", "Ratio"], [r + [None] * 3 for r in rows])
        r_sheet = book.sheet("R")
        book.table(r_sheet, "tblR", ["Key", "Val"], [["a", 8], ["b", 6]])
        this = lambda c: f"tblQ[[#This Row],[{c}]]"  # noqa: E731
        scaled = {"a": 4, "b": -1, "zz": -1}      # 8/2; 6/0 -> fallback; missed lookup/2 -> fallback
        guarded = {"a": 5, "b": 0, "zz": 2}       # the untaken IF branch must not trigger the fallback
        for i, (item, qty, div) in enumerate(rows, 2):
            book.put(q, f"D{i}", f"=IFERROR(_xlfn.XLOOKUP({this('Item')},tblR[Key],tblR[Val])/{this('Div')},-1)",
                     scaled[item])
            book.put(q, f"E{i}", f"=IFERROR(IF({this('Div')}=0,0,{this('Qty')}/{this('Div')}),-1)", guarded[item])
            if text_error:
                book.put(q, f"F{i}", f"=IFERROR({this('Item')}+1,0)", 0)
            else:
                book.put(q, f"F{i}", f"=IFERROR({this('Qty')}/{this('Div')},0)", qty / div if div else 0)
        summary = book.sheet("S")
        summary["A1"] = "Totals"
        for c, head in enumerate(["Item", "Scaled", "Guarded"], 1):
            summary.cell(3, c, head)
        for r, (item, _, _) in enumerate(rows, 4):
            summary.cell(r, 1, item)
            book.put(summary, f"B{r}", f"=SUMIFS(tblQ[Scaled],tblQ[Item],$A{r})", scaled[item])
            book.put(summary, f"C{r}", f"=SUMIFS(tblQ[Guarded],tblQ[Item],$A{r})", guarded[item])
        book.save(self.path())

    def test_a_missed_lookup_numerator_and_untaken_branches_match_excel(self):
        self.build()
        plan, _ = draft(self.path())
        emitted_problems(self, plan)
        calc = formulas(plan, "tblQ")
        # the guard covers the zero denominator AND the missed lookup, tested on the key column itself
        self.assertEqual(
            calc["Scaled"],
            "If(IsNull(Lookup([tblR/Key], [Item], [tblR/Key])) or IsNull([Div]) or [Div] = 0, (-1), "
            "(Lookup([tblR/Val], [Item], [tblR/Key]) / [Div]))")
        # IF(Div=0, 0, Qty/Div): the division only counts when the IF takes that branch
        self.assertEqual(calc["Guarded"],
                         "If(If(([Div] = 0), False, IsNull([Div]) or [Div] = 0), (-1), "
                         "If(([Div] = 0), 0, ([Qty] / [Div])))")
        self.assertEqual(calc["Ratio"], "If(IsNull([Div]) or [Div] = 0, 0, ([Qty] / [Div]))")

    def test_an_error_the_emitter_cannot_test_for_is_refused_not_swallowed(self):
        """IFERROR(text+1, 0) is a #VALUE! Excel catches; Sigma would not compile it. Never coerced."""
        self.build(text_error=True)
        with self.assertRaisesRegex(tr.PlanDraftError, r"#VALUE!"):
            draft(self.path())


# ---------------------------------------------------------------------------- 5. text matching
class TestTextMatching(WorkbookCase):
    def build(self, labels=("East", "West"), data=None, status_literal="Open"):
        data = data or [("East", "Open", 5), ("East", "Closed", 7), ("West", "Open", 2)]
        book = fx._Book()
        ws = book.sheet("Data")
        book.table(ws, "tblT", ["Region", "Status", "Amt", "Flag"], [list(r) + [None] for r in data])
        for i, (_, status, _) in enumerate(data, 2):
            book.put(ws, f"D{i}", f'=IF(tblT[[#This Row],[Status]]="{status_literal}","Y","N")',
                     "Y" if status.lower() == status_literal.lower() else "N")
        summary = book.sheet("S")
        summary["A1"] = "By region"
        summary.cell(3, 1, "Region")
        summary.cell(3, 2, "Amt")
        for r, label in enumerate(labels, 4):
            summary.cell(r, 1, label)
            book.put(summary, f"B{r}", f"=SUMIFS(tblT[Amt],tblT[Region],$A{r})",
                     sum(x[2] for x in data if x[0].lower() == label.lower()))
        book.save(self.path())

    def test_a_clean_workbook_is_accepted(self):
        self.build()
        plan, _ = draft(self.path())
        emitted_problems(self, plan)

    def test_wildcard_characters_in_a_row_label_are_refused_naming_the_cell(self):
        self.build(labels=("East*", "West"), data=[("East*", "Open", 5), ("East", "Closed", 7), ("West", "Open", 2)])
        with self.assertRaisesRegex(tr.PlanDraftError, r"S!A4: row label 'East\*' contains \* \? or ~"):
            draft(self.path())

    def test_a_question_mark_label_is_refused_too(self):
        self.build(labels=("Where?", "West"), data=[("Where?", "Open", 5), ("West", "Open", 2)])
        with self.assertRaisesRegex(tr.PlanDraftError, r"S!A4: row label 'Where\?'"):
            draft(self.path())

    def test_case_variants_of_a_group_label_are_refused(self):
        """Excel's SUMIFS merges East/east into one row; a Sigma grouped table keeps two."""
        self.build(data=[("East", "Open", 5), ("east", "Open", 7), ("West", "Open", 2)])
        with self.assertRaisesRegex(tr.PlanDraftError, r"Data!A3: tblT\[Region\] holds 'east' and 'East'"):
            draft(self.path())

    def test_a_text_literal_the_data_holds_in_another_case_is_refused(self):
        """IF([@Status]="OPEN") is TRUE for 'Open' in Excel, FALSE in Sigma."""
        self.build(status_literal="OPEN")
        with self.assertRaisesRegex(tr.PlanDraftError, r"holds 'Open' but tblT\[Flag\] compares it with 'OPEN'"):
            draft(self.path())

    def test_wildcard_criteria_text_is_still_refused(self):
        book = fx._Book()
        ws = book.sheet("Data")
        book.table(ws, "tblT", ["Region", "Status", "Amt"], [["East", "Open", 5], ["West", "Open", 2]])
        summary = book.sheet("S")
        summary["A1"] = "By region"
        summary.cell(3, 1, "Region")
        summary.cell(3, 2, "Amt")
        for r, label in enumerate(["East", "West"], 4):
            summary.cell(r, 1, label)
            book.put(summary, f"B{r}", f'=SUMIFS(tblT[Amt],tblT[Region],$A{r},tblT[Status],"Op*")',
                     {"East": 5, "West": 2}[label])
        book.save(self.path())
        with self.assertRaisesRegex(tr.PlanDraftError, r"wildcard criterion 'Op\*'"):
            draft(self.path())

    def test_index_match_keys_with_wildcards_or_case_duplicates_are_refused(self):
        def build(keys):
            book = fx._Book()
            rates = book.sheet("Rates")
            book.table(rates, "tblRates", ["Product", "Rate"], [[k, 0.1] for k in keys])
            ws = book.sheet("Data")
            book.table(ws, "tblT", ["Product", "Qty", "Rate"], [[keys[0], 5, None]])
            book.put(ws, "C2", "=INDEX(tblRates[Rate],MATCH(tblT[[#This Row],[Product]],tblRates[Product],0))", 0.1)
            summary = book.sheet("S")
            summary["A1"] = "Qty"
            summary.cell(3, 1, "Product")
            summary.cell(3, 2, "Qty")
            summary.cell(4, 1, keys[0])
            book.put(summary, "B4", "=SUMIFS(tblT[Qty],tblT[Product],$A4)", 5)
            book.save(self.path())
        build(["Widget", "WIDGET"])
        with self.assertRaisesRegex(tr.PlanDraftError, r"duplicate values \(text keys are compared ignoring case\)"):
            draft(self.path())


# ---------------------------------------------------------------------------- 6. tolerance
class TestTolerance(WorkbookCase):
    def test_a_cached_value_off_by_half_a_cent_is_a_mismatch(self):
        source = self.dir / "orders.xlsx"
        fx.build_orders(source)
        plan, _ = draft(source)
        want = next(cell[2] for g in plan["checks"] for cell in g["cells"]
                    if isinstance(cell[2], float) and abs(cell[2]) > 1)
        self.assertTrue(trm.values_match(want + 1e-10, want))
        self.assertFalse(trm.values_match(want + 0.004, want))      # passed under the old flat 0.01
        self.assertFalse(trm.values_match(0.5004, 0.5))             # a percent drifting 0.04 points
        self.assertTrue(trm.values_match(want + 0.004, want, 0.01))  # an explicit absolute override still works

    def test_the_default_limit_is_float_noise_only(self):
        self.assertAlmostEqual(trm.tolerance_for(0.0), 1e-6)
        self.assertAlmostEqual(trm.tolerance_for(14_402_043.63), 1.4402043630e-5)
        self.assertEqual(trm.tolerance_for(123.0, 0.01), 0.01)

    def test_the_broader_eight_fixtures_verify_at_the_tight_default(self):
        for name in ("04 Sales Pipeline Forecast.xlsx", "06 13-Week Cash Flow.xlsx",
                     "07 Headcount and Opex Plan.xlsx", "08 Inventory Replenishment Plan.xlsm"):
            with self.subTest(name):
                plan, summary = tr.draft_table_report(FIXTURES / name, intent="editable", connection="c",
                                                      folder="f")
                model = trm.compute(plan)
                _, problems = trm.check_cells(plan, model)
                self.assertEqual(problems, [])
                ok, lines = trm.parity_lines(plan, model)
                self.assertTrue(ok, lines)


# ---------------------------------------------------------------------------- the fixtures, through the emitted formulas
class TestEmittedFormulasOnTheFixtures(unittest.TestCase):
    def test_every_modelled_cell_of_04_and_07_agrees_with_the_workbook(self):
        for name, floor in (("04 Sales Pipeline Forecast.xlsx", 200), ("07 Headcount and Opex Plan.xlsx", 600)):
            with self.subTest(name):
                plan, summary = tr.draft_table_report(FIXTURES / name, intent="editable", connection="c", folder="f")
                for mod in ("floor", "trunc"):
                    compared, skipped, problems = se.compare_to_workbook(plan, mod)
                    self.assertEqual(problems, [])
                    self.assertGreaterEqual(compared, floor)
                    self.assertEqual(skipped, 0)


class TestEvaluator(unittest.TestCase):
    def ev(self, text, **row):
        return se.Evaluator(lambda *a: None).ev(se.parse(text), se.Ctx(lambda name: row[name]))

    def test_three_valued_logic(self):
        self.assertIsNone(self.ev("([A] = 1)", A=None))
        self.assertIs(self.ev("(IsNull([A]) or ([A] = 1))", A=None), True)
        self.assertIs(self.ev("(([A] = 1) and False)", A=None), False)
        self.assertIsNone(self.ev("(([A] = 1) and True)", A=None))
        self.assertEqual(self.ev("If(([A] = 1), 1, 2)", A=None), 2)     # NULL condition -> else
        self.assertEqual(self.ev("Coalesce([A], [B], 9)", A=None, B=None), 9)
        self.assertIsNone(self.ev("Least([A], 1)", A=None))
        self.assertEqual(self.ev("(-[A])", A=3), -3)


if __name__ == "__main__":
    unittest.main()
