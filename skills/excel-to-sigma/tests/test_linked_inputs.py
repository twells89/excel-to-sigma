#!/usr/bin/env python3
"""Linked input table Code Mode program emitter (validated live 2026-10-22)."""
from __future__ import annotations

import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS / "lib"))

from linked_inputs import (  # noqa: E402
    LinkedInputError,
    emit,
    lint_program,
    normalize_spec,
    parse_printed_ids,
    readback_queries,
)

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "linked-input"
SPEC = json.loads((FIXTURES / "spec.json").read_text(encoding="utf-8"))
CHECK_PY = Path(
    "/Users/tjwells/Library/Application Support/Claude/local-agent-mode-sessions/"
    "187fa98f-13d8-4528-bc58-05c343d73291/e3ed1a55-e316-42e3-8180-4e940664ee99/rpm/"
    "plugin_016bCC8zHHjeBE2MJRMHVmMS/skills/sigma-workbook-quickstart/scripts/check.py")


def spec(**changes):
    out = copy.deepcopy(SPEC)
    out.update(changes)
    return out


class GoldenProgramTests(unittest.TestCase):
    def test_programs_match_golden_text(self):
        result = emit(SPEC)
        for name in ("call1", "call2", "call3"):
            self.assertEqual(result["programs"][name], (FIXTURES / "golden" / (name + ".py")).read_text(encoding="utf-8"), name)
        self.assertEqual(result["readback"], json.loads((FIXTURES / "golden" / "readback.json").read_text(encoding="utf-8")))
        self.assertEqual(result["manifest"], json.loads((FIXTURES / "golden" / "manifest.json").read_text(encoding="utf-8")))

    def test_proven_shape_calls(self):
        programs = emit(SPEC)["programs"]
        self.assertIn("data_page.set_visibility(False)", programs["call1"])
        self.assertIn("sigma.formula.max(plan_base.get_column_by_name('Default Rate'))", programs["call1"])
        self.assertIn("any([c.name() == 'Default Rate' for c in t.columns()])", programs["call1"])
        self.assertIn("lookup_columns={'Line Item': 'Line Item'}", programs["call2"])
        self.assertIn("source_columns={'Baseline Rate': 'Baseline Rate'}", programs["call2"])
        self.assertIn("sigma.formula.coalesce(e0, i0)", programs["call2"])
        self.assertIn("lt.update_rows([", programs["call3"])
        self.assertIn("isinstance(v0, LiteralEditableColumn)", programs["call3"])
        self.assertIn("isinstance(lt, LinkedInputTable)", programs["call3"])

    def test_programs_obey_dialect_lint(self):
        for text in emit(SPEC)["programs"].values():
            self.assertEqual(lint_program(text), [])

    def test_lint_rejects_forbidden_constructs(self):
        for bad in ("def f():\n    pass\n", "x = 1 % 2\n", "x = [1, 2][0:1]\n", "x = f'a'\n",
                    "x = 'a' 'b'\n", "x = [a for a in [1] for b in [2]]\n", "x = 'a'.format()\n",
                    "while True:\n    pass\n", "x = sum([1])\n"):
            self.assertTrue(lint_program(bad), bad)

    @unittest.skipUnless(CHECK_PY.exists(), "Sigma quickstart check.py not installed")
    def test_quickstart_checker_is_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            for name, text in emit(SPEC)["programs"].items():
                path = Path(tmp) / (name + ".py")
                path.write_text(text, encoding="utf-8")
                done = subprocess.run([sys.executable, str(CHECK_PY), str(path)], capture_output=True, text=True)
                self.assertEqual(done.returncode, 0, name + done.stdout + done.stderr)

    def test_ids_fill_placeholders(self):
        ids = {"dataPageId": "dp1", "pageId": "pg1", "planBaseId": "pb1", "linkedTableId": "lt1", "workbookId": "wb1"}
        result = emit(SPEC, ids)
        self.assertIn("get_page('dp1')", result["programs"]["call2"])
        self.assertIn("get_element('pb1')", result["programs"]["call2"])
        self.assertIn("get_element('lt1')", result["programs"]["call3"])
        for call in result["manifest"]["calls"]:
            self.assertEqual(call["unresolvedPlaceholders"], [])
        self.assertEqual(result["manifest"]["calls"][1]["target"], {"type": "existing", "workbookId": "wb1"})
        self.assertEqual(result["manifest"]["calls"][0]["target"]["type"], "new")

    def test_existing_workbook_and_page(self):
        s = spec(workbook={"id": "wb-9"}, page={"id": "pg-9"})
        result = emit(s)
        self.assertNotIn("old_pages", result["programs"]["call1"])
        self.assertNotIn("add_page('Rate overrides')", result["programs"]["call1"])
        self.assertIn("print('pageId=' + 'pg-9')", result["programs"]["call1"])
        self.assertIn("get_page('pg-9')", result["programs"]["call2"])
        self.assertEqual(result["manifest"]["calls"][0]["target"], {"type": "existing", "workbookId": "wb-9"})

    def test_element_source_without_inherited_columns(self):
        s = spec(source={"kind": "element", "elementId": "el-1"}, inherited=[],
                 formulas=[], rows=[{"key": {"Line Item": "A"}, "values": {"Override Rate": 1}}])
        programs = emit(s)["programs"]
        self.assertIn("sigma.data.source.element('el-1')", programs["call1"])
        self.assertIn("plan_base.set_groupings([plan_base.get_column_by_name('Line Item')])\n", programs["call1"])
        self.assertNotIn("source_columns", programs["call2"])
        for text in programs.values():
            self.assertEqual(lint_program(text), [])

    def test_types_arithmetic_and_multi_key(self):
        s = spec(key=["Region", "Line Item"],
                 inherited=[{"name": "Baseline Total", "aggregate": "sum", "of": "Amount", "as": "Baseline"}],
                 editable=[{"name": "Adj", "type": "number"}, {"name": "When", "type": "datetime"},
                           {"name": "Frozen", "type": "boolean"}],
                 formulas=[{"name": "Plan", "formula": {"op": "coalesce", "args": [
                     {"col": "Adj"}, {"op": "multiply", "args": [{"col": "Baseline"},
                                                                 {"op": "add", "args": [{"const": 1}, {"op": "divide", "args": [{"op": "number", "args": [{"col": "Adj"}]}, {"const": 100}]}]}]}]}},
                           {"name": "Delta", "formula": {"op": "subtract", "args": [{"col": "Plan"}, {"col": "Baseline"}]}}],
                 rows=[{"key": {"Region": "West", "Line Item": "A"},
                        "values": {"Adj": 5, "When": "2026-09-12", "Frozen": True}}])
        result = emit(s)
        self.assertIn("datetime(2026, 9, 12, 12)", result["programs"]["call3"])
        self.assertIn("v2: True", result["programs"]["call3"])
        self.assertIn("{k0: 'West', k1: 'A'}", result["programs"]["call3"])
        self.assertIn("sigma.formula.coalesce(e0, sigma.formula.multiply(i0, sigma.formula.add(sigma.formula.const_num(1), "
                      "sigma.formula.divide(sigma.formula.number(e0), sigma.formula.const_num(100)))))", result["programs"]["call2"])
        self.assertIn("sigma.formula.subtract(f0, i0)", result["programs"]["call2"])
        self.assertTrue(any("multi-column key" in w for w in result["warnings"]))
        for text in result["programs"].values():
            self.assertEqual(lint_program(text), [])


class ValidationTests(unittest.TestCase):
    def fails(self, fragment, **changes):
        with self.assertRaises(LinkedInputError) as caught:
            normalize_spec(spec(**changes))
        self.assertIn(fragment, str(caught.exception))

    def test_key_must_be_non_empty_and_unique(self):
        self.fails("key must be a non-empty list", key=[])
        self.fails("key must be a non-empty list", key=None)
        self.fails("key columns must be unique", key=["A", "a"])
        self.fails("key[0]", key=[""])

    def test_seed_keys_unique_and_complete(self):
        rows = [{"key": {"Line Item": "A"}, "values": {"Override Rate": 1}},
                {"key": {"Line Item": "A"}, "values": {"Override Rate": 2}}]
        self.fails("duplicate seed key", rows=rows)
        self.fails("exactly the key columns", rows=[{"key": {"Other": "A"}, "values": {"Override Rate": 1}}])
        self.fails("is empty", rows=[{"key": {"Line Item": " "}, "values": {"Override Rate": 1}}])
        self.fails("non-empty string or number", rows=[{"key": {"Line Item": None}, "values": {"Override Rate": 1}}])
        self.fails("not an editable column", rows=[{"key": {"Line Item": "A"}, "values": {"Baseline Rate": 1}}])

    def test_value_types(self):
        for column, value in (("Override Rate", "x"), ("Override Rate", True), ("Rationale", 3)):
            self.fails("expected", rows=[{"key": {"Line Item": "A"}, "values": {column: value}}])
        self.fails("datetime must be", editable=[{"name": "When", "type": "datetime"}], formulas=[],
                   rows=[{"key": {"Line Item": "A"}, "values": {"When": "yesterday"}}])
        self.fails(".type must be one of", editable=[{"name": "X", "type": "date"}])

    def test_name_collisions(self):
        self.fails("collides with a source column", inherited=[{"name": "Default Rate", "aggregate": "max", "of": "Default Rate"}])
        self.fails("collides with a source column", inherited=[{"name": "Line Item", "aggregate": "max", "of": "Default Rate"}])
        self.fails("collides with a source column", sourceColumns=["Line Item", "Default Rate", "Baseline Rate"])
        self.fails("collides with another column", editable=[{"name": "baseline rate", "type": "number"}])
        self.fails("collides with another column", formulas=[{"name": "Override Rate", "formula": {"const": 1}}])

    def test_source_columns_are_checked_when_given(self):
        self.fails("is not in sourceColumns", sourceColumns=["Line Item", "Other"])
        self.fails("source.memberColumn", sourceColumns=["Line Item"], key=["Line Item"], inherited=[])

    def test_formula_descriptors(self):
        self.fails("unknown or later column", formulas=[{"name": "F", "formula": {"col": "Nope"}}])
        self.fails("unsupported op", formulas=[{"name": "F", "formula": {"op": "pow", "args": [{"const": 1}, {"const": 2}]}}])
        self.fails("takes exactly 2 args", formulas=[{"name": "F", "formula": {"op": "add", "args": [{"const": 1}]}}])
        self.fails("const must be a finite number", formulas=[{"name": "F", "formula": {"const": "1"}}])
        self.fails("exactly one of", formulas=[{"name": "F", "formula": {"col": "Override Rate", "const": 1}}])
        self.fails("unknown or later column", formulas=[
            {"name": "A", "formula": {"col": "B"}}, {"name": "B", "formula": {"const": 1}}])

    def test_workbook_page_source_and_modes(self):
        self.fails("workbook must be", workbook=None)
        self.fails("workbook.description", workbook={"name": "x"})
        self.fails("exactly one of {name}", page={})
        self.fails("page.id needs an existing workbook", page={"id": "p"})
        self.fails("page.name must differ", page={"name": "data"})
        self.fails("source.kind", source={"kind": "table"})
        self.fails("source.memberColumn", source={"kind": "dataModel", "inodeId": "x"})
        self.fails("table.editMode", table={"editMode": "everyone"})
        self.fails("frame must be", table={"frame": [1, 2, 3]})
        self.fails("at least one editable", editable=[])
        self.fails("more than 2000", rows=[{"key": {"Line Item": str(i)}, "values": {"Override Rate": 1}} for i in range(2001)])

    def test_cli_reports_validation_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "bad.json"
            bad.write_text(json.dumps(spec(key=[])), encoding="utf-8")
            done = subprocess.run([sys.executable, str(SCRIPTS / "linked-input-program.py"), "emit",
                                   "--spec", str(bad), "--out-dir", str(Path(tmp) / "o")], capture_output=True, text=True)
            self.assertEqual(done.returncode, 1)
            self.assertIn("key must be a non-empty list", done.stderr)
            self.assertFalse((Path(tmp) / "o").exists())


class ReadbackTests(unittest.TestCase):
    def test_placeholders_without_describe_ids(self):
        queries = {q["name"]: q["sql"] for q in readback_queries(normalize_spec(SPEC))["queries"]}
        self.assertIn('"workbook"."<ELEMENT ID>"', queries["rows"])
        self.assertIn('"<COLUMN ID: Line Item>"', queries["distinct_keys"])

    def test_sql_with_describe_ids(self):
        ids = {"describe": {"elementId": "EL1", "columns": {
            "Line Item": "c1", "Baseline Rate": "c2", "Override Rate": "c3", "Rationale": "c4", "Effective Rate": "c5"}}}
        out = readback_queries(normalize_spec(SPEC), ids)
        q = {item["name"]: item["sql"] for item in out["queries"]}
        self.assertEqual(q["distinct_keys"],
                         'SELECT COUNT(*) AS "row_count", COUNT(DISTINCT "c1") AS "distinct_keys" FROM "workbook"."EL1"')
        self.assertEqual(q["non_null_editable"],
                         'SELECT COUNT("c3") AS "Override Rate non-null", COUNT("c4") AS "Rationale non-null" FROM "workbook"."EL1"')
        self.assertEqual(q["sums"],
                         'SELECT SUM("c3") AS "Override Rate sum", SUM("c2") AS "Baseline Rate sum", '
                         'SUM("c5") AS "Effective Rate sum" FROM "workbook"."EL1"')
        self.assertIn("""WHERE "c1" IN ('Subscription - Enterprise', 'Hosting')""", q["seeded_rows"])
        self.assertEqual(out["expected"]["nonNullEditable"], {"Override Rate": 2, "Rationale": 2})
        self.assertEqual(out["expected"]["numericEditableSums"], {"Override Rate": 0.07})

    def test_quotes_are_escaped(self):
        s = spec(rows=[{"key": {"Line Item": "O'Brien \"x\""}, "values": {"Override Rate": 1}}])
        q = {item["name"]: item["sql"] for item in readback_queries(normalize_spec(s))["queries"]}
        self.assertIn("'O''Brien \"x\"'", q["seeded_rows"])
        program = emit(s)["programs"]["call3"]
        self.assertEqual(lint_program(program), [])
        compile(program, "call3", "exec")

    def test_multi_key_distinct_uses_concat(self):
        s = spec(key=["Region", "Line Item"], rows=[])
        q = {item["name"]: item["sql"] for item in readback_queries(normalize_spec(s))["queries"]}
        self.assertIn("COUNT(DISTINCT CONCAT(CAST(", q["distinct_keys"])
        self.assertNotIn("seeded_rows", q)


class IdsTests(unittest.TestCase):
    def test_parse_printed_ids_and_cli_roundtrip(self):
        text = ("dataPageId=dp\npageId=pg\nplanBaseId=pb\nprimaryKeyColumns=1\nlinkedTableId=lt\n"
                "column:Line Item=c-1\ncolumn:Override Rate=c-3\nnoise\n")
        ids = parse_printed_ids(text)
        self.assertEqual(ids["linkedTableId"], "lt")
        self.assertEqual(ids["printedColumnIds"], {"Line Item": "c-1", "Override Rate": "c-3"})
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            (tmp / "out.txt").write_text(text, encoding="utf-8")
            script = str(SCRIPTS / "linked-input-program.py")
            subprocess.run([sys.executable, script, "ids", "--from-output", str(tmp / "out.txt"),
                            "--out", str(tmp / "ids.json"), "--workbook-id", "wb"], check=True)
            self.assertEqual(json.loads((tmp / "ids.json").read_text())["workbookId"], "wb")
            (tmp / "spec.json").write_text(json.dumps(SPEC), encoding="utf-8")
            subprocess.run([sys.executable, script, "emit", "--spec", str(tmp / "spec.json"),
                            "--ids", str(tmp / "ids.json"), "--out-dir", str(tmp / "o")], check=True, capture_output=True)
            self.assertIn("get_element('lt')", (tmp / "o" / "call3.py").read_text())


if __name__ == "__main__":
    unittest.main()
