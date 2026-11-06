"""Phase 1b: landing workbook Tables in Snowflake (identifiers, types, CSV, commands, CLI)."""
from __future__ import annotations

import csv
import datetime as dt
import json
import os
import stat
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

import snowflake_landing as sl  # noqa: E402

CLI = ROOT / "scripts" / "land-in-snowflake.py"


def make_source(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Actuals"
    ws.append(["Posting Date", "Line Item", "Net Revenue ($)", "2026 Plan"])
    ws.append([dt.datetime(2026, 1, 5), "Rev", 100.5, "a"])
    ws.append([46023, "COGS", None, "b"])          # serial with a date format; blank amount
    ws["A3"].number_format = "yyyy-mm-dd"
    ws["A2"].number_format = "yyyy-mm-dd"
    ws.add_table(Table(displayName="tblActuals", ref="A1:D3"))
    wb.save(path)
    return path


FAKE_SNOW = r'''#!/usr/bin/env python3
import json, os, sys
log = os.environ["FAKE_SNOW_LOG"]
sql = sys.argv[sys.argv.index("-q") + 1] if "-q" in sys.argv else ""
with open(log, "a") as f:
    f.write(json.dumps(sys.argv[1:]) + "\n")
if sql.startswith("SHOW TABLES"):
    print(json.dumps([{"name": "FPA_T_ACTUALS"}] if os.environ.get("FAKE_EXISTS") else []))
elif "COUNT(*)" in sql:
    print(json.dumps([{"N": int(os.environ.get("FAKE_COUNT", "2"))}]))
else:
    print("[]")
'''

FAKE_SIGMA = r'''#!/usr/bin/env python3
import json, os, sys
with open(os.environ["FAKE_SIGMA_LOG"], "a") as f:
    f.write(json.dumps(sys.argv[1:]) + "\n")
print("{}")
'''


def install(bindir: Path, name: str, body: str):
    p = bindir / name
    p.write_text(body)
    p.chmod(p.stat().st_mode | stat.S_IEXEC)


class IdentifierTests(unittest.TestCase):
    def test_identifier_normalizes(self):
        self.assertEqual(sl.snowflake_identifier("Net Revenue ($)"), "NET_REVENUE")
        self.assertEqual(sl.snowflake_identifier("2026 Plan"), "C_2026_PLAN")
        with self.assertRaises(ValueError) as ctx:
            sl.snowflake_identifiers(["Net Revenue", "net-revenue"])
        self.assertIn("Net Revenue", str(ctx.exception))
        self.assertIn("net-revenue", str(ctx.exception))


class TypeAndCsvTests(unittest.TestCase):
    def test_month_column_lands_as_date(self):
        rows = [{"Month": dt.datetime(2026, 1, 1)}, {"Month": dt.date(2026, 2, 1)}]
        self.assertEqual(sl.infer_column_types(rows, ["Month"]), {"Month": "DATE"})
        serial = [{"Month": 46023}, {"Month": 46054}]
        self.assertEqual(sl.infer_column_types(serial, ["Month"]), {"Month": "DATE"})
        self.assertEqual(sl.infer_column_types([{"Amount": 46023}], ["Amount"]), {"Amount": "NUMBER(38,6)"})
        with tempfile.TemporaryDirectory() as d:
            src = make_source(Path(d) / "s.xlsx")
            plan = sl.landing_plan(src, ["tblActuals"], "DB", "SC", "FPA_T_")
            cols = {c["header"]: c["type"] for c in plan["tables"][0]["columns"]}
            self.assertEqual(cols["Posting Date"], "DATE")
            self.assertEqual(plan["tables"][0]["snowflakeTable"], "FPA_T_ACTUALS")

    def test_blank_actuals_become_null(self):
        with tempfile.TemporaryDirectory() as d:
            src = make_source(Path(d) / "s.xlsx")
            plan = sl.landing_plan(src, ["tblActuals"], "DB", "SC", "FPA_T_")
            t = plan["tables"][0]
            types = {c["name"]: c["type"] for c in t["columns"]}
            self.assertEqual(types["NET_REVENUE"], "NUMBER(38,6)")
            self.assertEqual([c["name"] for c in t["columns"]],
                             ["POSTING_DATE", "LINE_ITEM", "NET_REVENUE", "C_2026_PLAN"])
            with open(t["csv"], newline="") as f:
                rows = list(csv.reader(f))
            self.assertEqual(rows[2], ["2026-01-01", "COGS", "", "b"])
            self.assertEqual(rows[1][0], "2026-01-05")
            self.assertEqual(t["rowCount"], 2)


class CommandTests(unittest.TestCase):
    def plan(self, d):
        src = make_source(Path(d) / "s.xlsx")
        return sl.landing_plan(src, ["tblActuals"], "DB", "SC", "FPA_T_")

    def test_commands_refuse_existing_without_replace(self):
        with tempfile.TemporaryDirectory() as d:
            plan = self.plan(d)
            joined = "\n".join(" ".join(c) for c in sl.snow_commands(plan, "tj", False, None))
            self.assertNotIn("CREATE OR REPLACE", joined)
            self.assertNotIn("CREATE SCHEMA", joined)
            self.assertIn("CREATE TABLE", joined)
            replaced = "\n".join(" ".join(c) for c in sl.snow_commands(plan, "tj", True, None))
            self.assertIn("CREATE OR REPLACE TABLE", replaced)

    def test_grant_emitted_only_when_requested(self):
        with tempfile.TemporaryDirectory() as d:
            plan = self.plan(d)
            none = "\n".join(" ".join(c) for c in sl.snow_commands(plan, "tj", False, None))
            self.assertNotIn("GRANT", none)
            some = [c[c.index("-q") + 1] for c in sl.snow_commands(plan, "tj", False, "SIGMA_ROLE")]
            grants = [s for s in some if s.startswith("GRANT")]
            self.assertEqual(grants, ['GRANT SELECT ON TABLE DB.SC.FPA_T_ACTUALS TO ROLE SIGMA_ROLE'])
            self.assertTrue(some[-1].startswith("SELECT COUNT(*)"))


class CliTests(unittest.TestCase):
    def run_cli(self, d, extra=(), env=None):
        d = Path(d)
        bindir = d / "bin"
        bindir.mkdir(exist_ok=True)
        install(bindir, "snow", FAKE_SNOW)
        install(bindir, "sigma", FAKE_SIGMA)
        src = make_source(d / "s.xlsx")
        e = dict(os.environ, PATH=f"{bindir}{os.pathsep}{os.environ['PATH']}",
                 FAKE_SNOW_LOG=str(d / "snow.log"), FAKE_SIGMA_LOG=str(d / "sigma.log"),
                 LANG="en_US.UTF-8", LC_ALL="en_US.UTF-8", **(env or {}))
        return subprocess.run(
            [sys.executable, str(CLI), str(src), "--tables", "tblActuals", "--database", "DB",
             "--schema", "SC", "--table-prefix", "FPA_T_", "--snow-connection", "tj",
             "--out", str(d / "landing.json"), *extra],
            text=True, capture_output=True, env=e)

    def test_dry_run_prints_plan_and_runs_nothing(self):
        with tempfile.TemporaryDirectory() as d:
            r = self.run_cli(d)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertIn("CREATE TABLE", r.stdout)
            self.assertFalse((Path(d) / "snow.log").exists())
            landing = json.loads((Path(d) / "landing.json").read_text())
            self.assertEqual(landing["database"], "DB")
            self.assertEqual(landing["tables"][0]["snowflakeTable"], "FPA_T_ACTUALS")

    def test_refuses_existing_without_replace(self):
        with tempfile.TemporaryDirectory() as d:
            r = self.run_cli(d, ["--execute"], {"FAKE_EXISTS": "1"})
            self.assertEqual(r.returncode, 2)
            self.assertIn("table DB.SC.FPA_T_ACTUALS exists; pass --replace to overwrite", r.stderr)
            self.assertNotIn("CREATE", (Path(d) / "snow.log").read_text())

    def test_execute_succeeds_and_syncs_when_asked(self):
        with tempfile.TemporaryDirectory() as d:
            r = self.run_cli(d, ["--execute", "--grant-select-to", "SIGMA_ROLE",
                                 "--sync-sigma-connection", "conn-1", "--sigma-profile", "prof"])
            self.assertEqual(r.returncode, 0, r.stderr)
            log = (Path(d) / "snow.log").read_text()
            self.assertIn("GRANT SELECT ON TABLE DB.SC.FPA_T_ACTUALS TO ROLE SIGMA_ROLE", log)
            sig = [json.loads(x) for x in (Path(d) / "sigma.log").read_text().splitlines()]
            self.assertEqual(len(sig), 1)
            self.assertIn("sync", sig[0])
            self.assertIn('{"path": ["DB", "SC", "FPA_T_ACTUALS"]}', sig[0])

    def test_row_count_mismatch_fails(self):
        with tempfile.TemporaryDirectory() as d:
            r = self.run_cli(d, ["--execute"], {"FAKE_COUNT": "7"})
            self.assertNotEqual(r.returncode, 0)
            self.assertIn("row count mismatch", r.stderr)


class DocConsistencyTests(unittest.TestCase):
    def test_docs_name_the_table_the_landing_actually_creates(self):
        # snowflake_table_name strips a leading "tbl": --tables tblActuals --table-prefix FPA_
        # lands FPA_ACTUALS, so every --actuals-table example must name that table.
        import re
        landed = sl.snowflake_table_name("tblActuals", "FPA_")
        self.assertEqual(landed, "FPA_ACTUALS")
        docs = [ROOT / "SKILL.md", *sorted((ROOT / "refs").glob("*.md"))]
        examples = [(p.name, m) for p in docs
                    for m in re.findall(r"--actuals-table\s+\S*?\.(FPA_\w+)", p.read_text(encoding="utf-8"))]
        self.assertTrue(examples)
        for name, table in examples:
            self.assertEqual(table, landed, name)
        for p in docs:
            self.assertNotIn("FPA_TBL", p.read_text(encoding="utf-8").upper(), p.name)


if __name__ == "__main__":
    unittest.main()
