#!/usr/bin/env python3
"""Phase 1b — land workbook Tables in Snowflake with the ``snow`` CLI.

Default is a dry run (prints the plan and the SQL). ``--execute`` runs it, verifies each row
count equals the source's, and exits non-zero on a mismatch. An existing table without
``--replace`` exits 2. ``--sync-sigma-connection <id>`` then makes Sigma see the tables.
"""
import argparse, json, subprocess, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
import snowflake_landing as sl  # noqa: E402


def run(argv):
    return subprocess.run(argv, text=True, capture_output=True)


def first_value(stdout):
    data = json.loads(stdout or "[]")
    if not data:
        return None
    return list(data[0].values())[0]


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("file")
    p.add_argument("--tables", required=True, help="comma-separated Excel Table names")
    p.add_argument("--database", required=True)
    p.add_argument("--schema", required=True)
    p.add_argument("--table-prefix", default="")
    p.add_argument("--snow-connection", required=True)
    p.add_argument("--grant-select-to", help="role granted SELECT on each landed table (that grant only)")
    p.add_argument("--replace", action="store_true")
    p.add_argument("--execute", action="store_true")
    p.add_argument("--sync-sigma-connection", help="Sigma connection id; sync each landed table path after --execute")
    p.add_argument("--sigma-profile", help="Sigma CLI profile for the sync")
    p.add_argument("--out", required=True, help="landing.json path (CSVs are written beside it)")
    a = p.parse_args(argv)

    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        plan = sl.landing_plan(Path(a.file), [t for t in a.tables.split(",") if t], a.database,
                               a.schema, a.table_prefix, out_dir=out.parent)
    except ValueError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    plan["snowConnection"] = a.snow_connection
    out.write_text(json.dumps(plan, indent=2))
    commands = sl.snow_commands(plan, a.snow_connection, a.replace, a.grant_select_to)
    for table in plan["tables"]:
        print(f"{sl.fqn(plan, table)}: {table['rowCount']} rows from {table['excelTable']}")
    print(f"plan written to {out}")
    if not a.execute:
        print("DRY RUN — commands:")
        for c in commands:
            print("  " + " ".join(f"'{x}'" if " " in x else x for x in c))
        return 0

    if not a.replace:
        for table in plan["tables"]:
            r = run(sl.exists_command(plan, table, a.snow_connection))
            if r.returncode != 0:
                print(r.stderr, file=sys.stderr)
                return 1
            if json.loads(r.stdout or "[]"):
                print(f"table {sl.fqn(plan, table)} exists; pass --replace to overwrite", file=sys.stderr)
                return 2
    per_table = len(commands) // len(plan["tables"])
    for i, table in enumerate(plan["tables"]):
        for c in commands[i * per_table:(i + 1) * per_table]:
            r = run(c)
            if r.returncode != 0:
                print(f"FAILED: {c[c.index('-q') + 1][:120]}\n{r.stderr}", file=sys.stderr)
                return 1
        count = first_value(r.stdout)   # the last command of each table is the COUNT(*)
        if count is None or int(count) != table["rowCount"]:
            print(f"row count mismatch for {sl.fqn(plan, table)}: warehouse={count} source={table['rowCount']}",
                  file=sys.stderr)
            return 1
        print(f"landed {sl.fqn(plan, table)}: {count} rows (matches source)")
    if a.sync_sigma_connection:
        for table in plan["tables"]:
            r = run(sl.sigma_sync_command(plan, table, a.sync_sigma_connection, a.sigma_profile))
            if r.returncode != 0:
                print(f"Sigma sync failed for {table['snowflakeTable']}: {r.stderr}", file=sys.stderr)
                return 1
            print(f"synced {sl.fqn(plan, table)} to Sigma connection {a.sync_sigma_connection}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
