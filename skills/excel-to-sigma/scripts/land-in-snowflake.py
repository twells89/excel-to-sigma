#!/usr/bin/env python3
"""Phase 1b — land workbook Tables in Snowflake with the ``snow`` CLI.

Default is a dry run (prints the plan and the SQL, and runs one read-only ``SHOW TABLES`` per target
to say whether ``--execute`` will need ``--replace``). ``--execute`` runs it, verifies each row
count equals the source's, and exits non-zero on a mismatch. An existing table without
``--replace`` exits 2. ``--sync-sigma-connection <id>`` then makes Sigma see the tables.
"""
import argparse, json, os, subprocess, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))
import snowflake_landing as sl  # noqa: E402


def run(argv, timeout=None, stdin=None):
    return subprocess.run(argv, text=True, capture_output=True, timeout=timeout, stdin=stdin)


# The dry run's existence probe is advisory: it must never hang on a ``snow`` connection that waits
# for an SSO / externalbrowser sign-in, and never crash on output that is not JSON.
PROBE_TIMEOUT_SECONDS = 60
NOT_CHECKED = "; --execute checks again and stops if one exists"


def first_value(stdout):
    data = json.loads(stdout or "[]")
    if not data:
        return None
    return list(data[0].values())[0]


def existing_tables_note(plan, connection, replace):
    """Read-only SHOW TABLES per target, so the dry run says up front whether --execute needs --replace."""
    found = []
    for table in plan["tables"]:
        try:
            r = run(sl.exists_command(plan, table, connection), timeout=PROBE_TIMEOUT_SECONDS,
                    stdin=subprocess.DEVNULL)   # a sign-in prompt reads EOF instead of waiting
        except subprocess.TimeoutExpired:
            return (f"EXISTING TABLES: not checked (snow did not answer within {PROBE_TIMEOUT_SECONDS}s; "
                    f"is the connection waiting for a browser sign-in?){NOT_CHECKED}")
        except OSError as error:
            return f"EXISTING TABLES: not checked ({error}){NOT_CHECKED}"
        if r.returncode != 0:
            reason = (r.stderr or r.stdout).strip().splitlines()
            return f"EXISTING TABLES: not checked ({reason[-1] if reason else f'exit {r.returncode}'}){NOT_CHECKED}"
        try:
            rows = json.loads(r.stdout or "[]")
        except json.JSONDecodeError:
            first = (r.stdout.strip().splitlines() or ["?"])[0][:80]
            return f"EXISTING TABLES: not checked (snow printed non-JSON output: {first!r}){NOT_CHECKED}"
        if rows:
            found.append(sl.fqn(plan, table))
    if not found:
        return "EXISTING TABLES: none of " + ", ".join(sl.fqn(plan, t) for t in plan["tables"]) + " exist yet"
    names = ", ".join(found)
    verb = "exists" if len(found) == 1 else "exist"
    if replace:
        return f"EXISTING TABLES: {names} already {verb}; --replace overwrites {'it' if len(found) == 1 else 'them'}"
    return f"EXISTING TABLES: {names} already {verb}; --execute stops (exit 2) unless you pass --replace"


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
    p.add_argument("--sigma-profile", default=os.environ.get("SIGMA_CLI_PROFILE") or None,
                   help="Sigma CLI profile for the sync (default: $SIGMA_CLI_PROFILE)")
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
        print(existing_tables_note(plan, a.snow_connection, a.replace))
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
