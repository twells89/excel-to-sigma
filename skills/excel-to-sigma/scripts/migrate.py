#!/usr/bin/env python3
"""Route an Excel workbook: audit + discover + classify, then print the next steps as JSON.

usage: migrate.py plan <file.xlsx|.xlsm> [--out-dir DIR]
       migrate.py resolve-targets [--folder-name NAME] [--connection-name NAME] [--sigma-profile NAME]

Writes migration-audit.json, discovery.json, macros.json (forecast with macros), macro-gate.json
(macro-enabled sources: the Phase 0b disposition gate record) and (when the audit has review issues)
review.template.json into DIR (default: the workbook's folder). Prints a JSON routing
record: audit status and issues, archetype, the refs worth reading, the decisions to
ask the user, the next commands in the order they run, the staging alternative to the
production post, and the steps only the user can do in the Sigma UI.
Exit 2 only when the audit is BLOCKED.

resolve-targets looks up the target folder and the write-enabled connection by name
(read-only REST GETs) and prints their ids for the drafter's --folder / --connection.
"""
from __future__ import annotations

import argparse
import json
import shlex
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "lib"))

import sigma_auth  # noqa: E402
import sigma_targets  # noqa: E402
from forecast_plan import PlanDraftError, draft_plan  # noqa: E402
from migration_gate import source_sha256  # noqa: E402
from table_report import NoTableReportShape, PlanDraftError as TableReportDraftError  # noqa: E402
from table_report import draft_table_report  # noqa: E402
from xlsx_audit import audit_workbook  # noqa: E402

MACRO_SUFFIXES = {".xlsm", ".xlsb", ".xls"}
RESEARCH_HINTS = ("FY results", "PER SHARE", "PROFIT AND LOSS", "__FDSCACHE__")
PRESENTATION_CODES = {"CHARTS_PRESENT", "DATA_VALIDATION_PRESENT", "CONDITIONAL_FORMATTING_PRESENT"}


def run_json(*args):
    out = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True)
    if out.returncode:
        lines = (out.stderr or out.stdout).strip().splitlines()
        raise RuntimeError(lines[-1] if lines else f"exit {out.returncode}")
    return json.loads(out.stdout)


def archetype(discovery):
    if any(hint in name for name in discovery.get("sheets", []) for hint in RESEARCH_HINTS):
        return "research"
    report = {s["sheet"]: s["classes"] for s in discovery.get("reportSheets", [])}
    if any(c.get("ACTUAL") or c.get("DRIVER_GROWTH") for c in report.values()):
        return "forecast"
    if discovery.get("tables"):
        return "input-table"
    return "unknown"


# The setup program every forecast build writes (--emit-manual-commit-program) and the post prints
# as NEXT; one name everywhere.
SETUP_PROGRAM = "<workdir>/setup-program.py"
# The forecast route's one post-create MCP step (replaces pasting seed CSVs in the UI).
FORECAST_SETUP_HANDOFF = (
    f"Run the setup program the post wrote ({SETUP_PROGRAM}) once via Sigma MCP Code Mode build_workbook "
    "(target existing) on the new workbook: it enables manual commit on every input table and seeds the "
    "rate and manual input tables with the seed CSV rows (fresh workbook only; the writes do not query first)")
RESOLVE_TARGETS = ("resolve both ids by name: python scripts/migrate.py resolve-targets --folder-name "
                   "'<folder name>' --connection-name '<connection name>'   # read-only; refuses a "
                   "connection without write access")

ROUTES = {
    "forecast": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/model-taxonomy.md", "refs/forecast-recipes.md"],
        "decisions": ["intent: read-only | editable | what-if", "write-enabled connectionId", "target folderId",
                      RESOLVE_TARGETS],
        "next": [],      # built by forecast_commands(): the order depends on Tables and macros
        "handoffs": [FORECAST_SETUP_HANDOFF, "Publish the workbook"],
    },
    "research": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/research-recipes.md", "refs/research-template-coa.md"],
        "decisions": ["sheet name and first/last line-item rows", "write-enabled connectionId", "target folderId",
                      RESOLVE_TARGETS],
        "next": [
            "python scripts/infer-canonical-formulas.py <file> --sheet=<sheet> --first-row=<n> "
            "--last-row=<n> --out <workdir>/plan.json --draft",
            "python scripts/build-research-model.py <workdir>/plan.json --conn <conn> --folder <folder>   # dry run",
            "python scripts/build-research-model.py <workdir>/plan.json --conn <conn> --folder <folder> "
            "--post --review-dispositions <workdir>/review.json",
        ],
        "handoffs": ["Publish the workbook"],
    },
    "input-table": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/input-tables.md", "refs/excel-translation.md"],
        "decisions": ["fact Table and canonical grain", "write-enabled connectionId"],
        "next": [
            "python scripts/xlsx-to-input-csv.py <file> --table <table> --out <workdir>/seed.csv",
            "python scripts/build-input-table-wb.py --name <name> --connection <conn> --columns <COL:type,...>",
            "python scripts/build-dm-on-view.py ...   # after the warehouse view exists",
        ],
        "handoffs": ["Paste seed.csv into the input table", "Publish", "Create warehouse view on the input table"],
    },
    # Formal Tables + formula grids (SUMIFS/COUNTIFS summaries, ordinal roster x period grids,
    # running balances, period ledgers such as a 13-week cash flow) that draft-table-report.py
    # reproduces and verifies cell by cell.
    "table-report": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/model-taxonomy.md", "refs/input-tables.md",
                 "refs/table-report-recipes.md"],
        "decisions": ["intent: read-only | editable | what-if (which tables people edit)",
                      "write-enabled connectionId", "target folderId", RESOLVE_TARGETS],
        "next": [
            "python scripts/draft-table-report.py <file> --intent <intent> --connection <conn> "
            "--folder <folder> --out <workdir>/plan.json   # fails loudly on unknown shapes; verifies "
            "every formula cell against Excel's cached values",
            "python scripts/build-table-report.py <workdir>/plan.json --out-dir <workdir>/build   "
            "# dry run: draft mode + local parity + seed CSVs",
            "python scripts/build-table-report.py <workdir>/plan.json --post "
            "--review-dispositions <workdir>/review.json   # production: strict dispositions",
        ],
        "handoffs": ["Seed the empty input tables from seed_*.csv through Sigma MCP Code Mode "
                     "(InputTable.upsert_rows; query first - it does not enforce keys)",
                     "Query a few summary cells and the running balance against the plan's parity targets "
                     "(CumulativeSum partition and grouped-table behaviour are live-validated only then)",
                     "Set the selected scenario and other parameters in the Parameters input table "
                     "(a workbook control is a filter, not a formula parameter)",
                     "Rebuild the source charts (CHARTS_PRESENT)", "Publish the workbook"],
    },
    # SUMIFS / driver signals, but not the shape the forecast drafter and builder support.
    "cell-model": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/model-taxonomy.md", "refs/source-workbook-model.md",
                 "refs/workbook-expression-context.md", "refs/native-calculations.md"],
        "decisions": ["edit intent per sheet (read-only report / editable plan / live what-if)",
                      "which sheets and ranges carry the model", "write-enabled connectionId", "target folderId",
                      RESOLVE_TARGETS],
        "next": [
            "python scripts/source-workbook.py <file> --out <workdir>/source-model.json   "
            "# source-bound cell/formula inventory",
            "Build source-specific with Sigma MCP Code Mode per refs/model-taxonomy.md; no bundled "
            "builder fits this shape (see routeReason)",
        ],
        "handoffs": [],
    },
    "unknown": {
        "refs": ["refs/sigma-build-gotchas.md", "refs/model-taxonomy.md"],
        "decisions": ["what the workbook is used for (entry vs report)"],
        "next": ["python scripts/xlsx-discover.py <file>   # read the human report and ask the user"],
        "handoffs": [],
    },
    "blocked": {
        "refs": [],
        "decisions": ["remove or replace the blocker (see issues), then re-run plan"],
        "next": [],
        "handoffs": [],
    },
}


# Replaces FORECAST_SETUP_HANDOFF when the build had --workflows: that build writes the program with
# the workflow tables added, and the program is run exactly once per workbook.
ENHANCE_SETUP_HANDOFF = (
    f"Run the setup program the --workflows post wrote ({SETUP_PROGRAM}) once via Sigma MCP Code Mode "
    "build_workbook (target existing) on the new workbook: it enables manual commit on every input table "
    "(backend actions and agent tools can only write input tables with manual commit enabled) and seeds the "
    "Parameters row plus the rate and manual input tables (fresh workbook only)")
# The live proof of the workflows and the agent, before the run is called clean (off camera).
SMOKE_PASS_HANDOFF = (
    "Smoke pass once the workbook is seeded, before calling the run clean: switch scenario (forecast changes), "
    "roll forward (start month +1, the closed month reads actuals), snapshot (one Versions row, label vN), "
    "submit for approval (one Approvals row, status pill Submitted) through the buttons or call_agent; read "
    "each back with Sigma MCP query; then restore the start month, any edited rate and delete the test rows, "
    "and re-check parity. See SKILL.md Phase 6 'Smoke pass'")
MACRO_GATE = "<workdir>/macro-gate.json"


def forecast_commands(tables, enhanced):
    """The forecast route in the order it runs. With Tables, Phase 1b lands them and every build reads
    the landed actuals; with macros (enhanced), the workflows manifest, the Phase 6 pitch and the
    agent ride the same builds. Returns ``(next_commands, staging_alternative)``."""
    landed = (" --actuals-table <db>.<schema>.<prefix>ACTUALS --landing <workdir>/landing.json"
              if tables else "")
    workflows = " --workflows <workdir>/workflows.json" if enhanced else ""
    build = (f"python scripts/build-forecast-model.py <workdir>/plan.json{landed}{workflows} "
             f"--emit-manual-commit-program {SETUP_PROGRAM} --out-dir <workdir>/build")
    agent = " --agent-macros <workdir>/macros.json" if enhanced else ""
    commands = []
    if tables:
        land = ("python scripts/land-in-snowflake.py <file> --tables <tables> --database <db> --schema <schema> "
                "--table-prefix <prefix> --snow-connection <snow-conn> --out <workdir>/landing.json")
        commands += [
            f"{land}   # Phase 1b dry run: SQL plus a read-only check for existing tables (source Tables: "
            + ", ".join(tables) + "; land the actuals and plan Tables)",
            f"{land} --execute --grant-select-to <role> --sync-sigma-connection <conn>   "
            "# after the user's OK; add --replace only when the dry run says the tables exist",
        ]
    commands.append("python scripts/draft-forecast-plan.py <file> --intent <intent> --connection <conn> "
                    "--folder <folder> --out <workdir>/plan.json   # fails loudly on unknown shapes; verifies "
                    "against cached values")
    if enhanced:
        commands.append("python scripts/macro-actions.py demo-manifest <file> --out <workdir>/workflows.json   "
                        "# classify the macros into workflows (never executes VBA)")
    commands.append(f"{build}   # dry run: draft mode + local parity; writes specs and the setup program")
    if enhanced:
        commands.append("python scripts/recommend-agent.py --macros <workdir>/macros.json --workflows "
                        "<workdir>/workflows.json   # Phase 6 pitch only (read it aloud); with a yes the post "
                        "below carries the agent via --agent-macros, in the same POST")
    commands.append(f"{build}{agent} --post --review-dispositions <workdir>/review.json   # production: strict "
                    f"verified dispositions; creates the workbook and rewrites {SETUP_PROGRAM}")
    staging = [
        f"{build} --staging-approval-template <workdir>/staging-approval.json   # writes the exact audited "
        "issues for the plan's folder, status pending; hand them to the named reviewer",
        f"{build} --staging-approval-template <workdir>/staging-approval.json --owner '<owner>' "
        "--reviewer '<reviewer>' --approve-for-staging   # USER HANDOFF: the named reviewer runs this in "
        "their own terminal; it lists every issue and waits for them to type APPROVE (it refuses without "
        "an interactive terminal, so the agent cannot approve); nothing is marked verified",
        f"{build}{agent} --post --staging-approval <workdir>/staging-approval.json   # staging post into that "
        f"folder; creates the workbook and rewrites {SETUP_PROGRAM}",
    ]
    return commands, staging


def macro_gate(macros):
    """Phase 0b gate record: every procedure surfaced, none dropped, nothing verified by the router."""
    records = [{"name": m.get("name"), "class": m.get("class"), "status": m.get("status"),
                "sigmaTarget": m.get("sigmaTarget"),
                "disposition": ("drop after formula parity" if m.get("status") == "AUTO"
                                else "implement, then verify live (pending)"),
                "verified": False} for m in macros]
    return {"procedures": len(records), "surfaced": [r["name"] for r in records], "dropped": [],
            "nonAuto": [r["name"] for r in records if r["status"] != "AUTO"], "records": records}


def plan(args):
    source = Path(args.workbook).expanduser().resolve()
    if not source.is_file():
        print(f"source workbook not found: {source}", file=sys.stderr)
        return 1
    workdir = Path(args.out_dir or source.parent).expanduser().resolve()
    workdir.mkdir(parents=True, exist_ok=True)

    audit = audit_workbook(source)
    (workdir / "migration-audit.json").write_text(json.dumps(audit, indent=2) + "\n", encoding="utf-8")
    blocked = audit["status"] == "BLOCKED"

    discovery, macros, notes = {}, None, []
    if blocked:
        kind = "blocked"
    else:
        try:
            discovery = run_json(HERE / "xlsx-discover.py", source, "--json")
        except RuntimeError as error:
            print(f"discovery failed: {error}", file=sys.stderr)
            return 1
        (workdir / "discovery.json").write_text(json.dumps(discovery, indent=2) + "\n", encoding="utf-8")
        kind = archetype(discovery)
    route_reason = None
    forecast_rejection = None
    if kind == "forecast":
        # Same shape checks the drafter applies, so the route never names a command that
        # fails on this workbook. Placeholder connection/folder: nothing is written.
        try:
            _, shape = draft_plan(source, intent="what-if", connection="<connection>", folder="<folder>")
            route_reason = (f"forecast drafter accepts this workbook: {shape['lines']} lines "
                            f"{shape['kinds']} on {shape['reportSheet']!r}")
        except PlanDraftError as error:
            kind = "cell-model"
            forecast_rejection = f"SUMIFS/driver signals, but the forecast drafter rejects the shape: {error}"
            route_reason = forecast_rejection
        except Exception as error:  # a drafter defect must never break routing; say so instead
            kind = "cell-model"
            forecast_rejection = (f"SUMIFS/driver signals, but the forecast drafter could not run "
                                  f"({type(error).__name__}: {error})")
            route_reason = forecast_rejection
    if kind in ("cell-model", "input-table") and discovery.get("tables"):
        # Formal Tables with formula-driven summaries: named only when the table-report drafter
        # accepts the workbook and reproduces every cached cell.
        try:
            _, shape = draft_table_report(source, intent="what-if", connection="<connection>",
                                          folder="<folder>")
            kind = "table-report"
            route_reason = (f"table-report drafter accepts this workbook: {len(shape['tables'])} table(s), "
                            f"{len(shape['frames'])} frames, {shape['verifiedCells']} source cells verified"
                            + (f" (forecast drafter: {forecast_rejection})" if forecast_rejection else ""))
        except NoTableReportShape:
            pass  # not a near miss: keep the input-table / cell-model reasoning as is
        except TableReportDraftError as error:
            note = f"table-report drafter rejects the shape: {error}"
            route_reason = f"{route_reason}; {note}" if route_reason else note
        except Exception as error:  # a drafter defect must never break routing; say so instead
            note = f"table-report drafter could not run ({type(error).__name__}: {error})"
            route_reason = f"{route_reason}; {note}" if route_reason else note
    route = ROUTES[kind]
    refs = list(route["refs"])
    next_commands = list(route["next"])
    handoffs = list(route["handoffs"])
    staging_alternative = []
    gate = None

    enhanced = False
    if not blocked and source.suffix.lower() in MACRO_SUFFIXES:
        try:
            macros = run_json(HERE / "macro-classify.py", "--json", source)
        except RuntimeError as error:
            macros = {"error": str(error)}
            notes.append(f"macro-classify failed ({error}); resolve it before building.")
        refs.append("refs/macro-handling.md")
        if isinstance(macros, list) and macros:
            gate = macro_gate(macros)
            (workdir / "macro-gate.json").write_text(json.dumps(gate, indent=2) + "\n", encoding="utf-8")
        enhanced = kind == "forecast" and isinstance(macros, list) and bool(macros)
        if enhanced:
            refs.append("refs/agent-enhance.md")
            (workdir / "macros.json").write_text(json.dumps(macros, indent=2) + "\n", encoding="utf-8")
            handoffs = [ENHANCE_SETUP_HANDOFF if h == FORECAST_SETUP_HANDOFF else h for h in handoffs]
            handoffs.insert(1, SMOKE_PASS_HANDOFF)
    if kind == "forecast":
        next_commands, staging_alternative = forecast_commands(
            [t["name"] for t in discovery.get("tables", [])], enhanced)
    if gate:   # Phase 0b runs first, whatever the route: read every non-AUTO record to the user
        next_commands.insert(0, f"python scripts/macro-classify.py <file>   # Phase 0b: surface all "
                                f"{gate['procedures']} procedures ({len(gate['nonAuto'])} non-AUTO) before "
                                f"building; the gate record is {MACRO_GATE}")
    if any(i.get("code") in PRESENTATION_CODES for i in audit.get("issues", [])):
        refs.append("refs/presentation-migration.md")

    digest = source_sha256(source)
    review_issues = [i for i in audit.get("issues", []) if i.get("severity") != "blocker"]
    review_template = {
        "sourceSha256": digest,
        "issues": [{"code": i["code"], "owner": "", "disposition": "", "verified": False}
                   for i in review_issues],
    }
    if review_template["issues"]:
        (workdir / "review.template.json").write_text(json.dumps(review_template, indent=2) + "\n",
                                                      encoding="utf-8")

    def fill(command):
        return command.replace("<file>", shlex.quote(str(source))).replace("<workdir>", shlex.quote(str(workdir)))

    result = {
        "source": str(source),
        "sourceSha256": digest,
        "auditStatus": audit["status"],
        "issues": [{"code": i["code"], "severity": i.get("severity"), "route": i.get("route")}
                   for i in audit.get("issues", [])],
        "archetype": kind,
        "routeReason": route_reason,
        "modelMap": {s["sheet"]: s["classes"] for s in discovery.get("reportSheets", [])},
        "tables": [{"name": t["name"], "sheet": t["sheet"], "rows": t["rows"]} for t in discovery.get("tables", [])],
        "macros": macros,
        "readRefs": list(dict.fromkeys(refs)),
        "beforePost": ["refs/coverage-contract.md"],
        "askUser": route["decisions"],
        "nextCommands": [fill(c) for c in next_commands],
        "stagingAlternative": [fill(c) for c in staging_alternative],
        "macroGate": (fill(MACRO_GATE) if gate else None),
        "userHandoffs": [fill(h) for h in handoffs],
        "reviewTemplate": str(workdir / "review.template.json") if review_template["issues"] else None,
        "notes": notes + [
            "Local builds (no --post) run in draft mode; pending review issues do not block them.",
            "Fill review.template.json (owner, disposition, verified) only after the feature is built "
            "and checked, before --post.",
            "BLOCKED stops everything." if blocked else "Not blocked.",
        ] + (["The drafter may add plan review issues (e.g. FLAT_CARRY_POLICY); add their codes to "
              "review.json too before --post.",
              "stagingAlternative replaces the production post (the last nextCommand) when review issues are "
              "honestly pending and the target is an isolated staging folder. Its approval step is the "
              "named reviewer's, typed in their own terminal; the agent runs the template and the post only.",
              "Export SIGMA_CLI_PROFILE=<profile> once (off camera) instead of passing --sigma-profile on "
              "each command."] if kind == "forecast" else []
             ) + (["The drafter adds plan review issues (e.g. GROUPED_ROWS_FROM_DATA, "
                   "RUNNING_BALANCE_PARTITION_REVIEW, INDIRECT_SCENARIO_LOOKUP, LEDGER_TRANSPOSED); add their "
                   "codes to review.json too before --post."]
                  if kind == "table-report" else []),
    }
    print(json.dumps(result, indent=2))
    return 2 if blocked else 0


def resolve_targets(args):
    if not (args.folder_name or args.connection_name):
        print("resolve-targets: pass --folder-name and/or --connection-name", file=sys.stderr)
        return 1
    try:
        sigma_auth.apply_cli_args(args)
        found = sigma_targets.resolve_targets(sigma_targets.sigma_api, args.folder_name, args.connection_name)
    except (sigma_targets.TargetError, sigma_auth.SigmaAuthError) as error:
        print(f"resolve-targets: {error}", file=sys.stderr)
        return 1
    folder, connection = found.get("folder"), found.get("connection")
    flags = " ".join(([f"--connection {connection['id']}"] if connection else [])
                     + ([f"--folder {folder['id']}"] if folder else []))
    print(json.dumps({"folderId": folder and folder["id"], "folder": folder,
                      "connectionId": connection and connection["id"], "connection": connection,
                      "draftFlags": flags}, indent=2))
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("plan", help="audit, discover and print next steps")
    p.add_argument("workbook")
    p.add_argument("--out-dir", help="where to write audit/discovery/review-template JSON")
    r = sub.add_parser("resolve-targets", help="look up the folder and write-enabled connection ids by name")
    r.add_argument("--folder-name", help="exact folder name (case-insensitive)")
    r.add_argument("--connection-name", help="exact connection name (case-insensitive); must be write-enabled")
    sigma_auth.add_cli_args(r)
    args = parser.parse_args(argv)
    return plan(args) if args.command == "plan" else resolve_targets(args)


if __name__ == "__main__":
    raise SystemExit(main())
