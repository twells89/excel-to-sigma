#!/usr/bin/env python3
"""
build-table-report.py - builds the Sigma data model + workbook for a table-report plan
(drafted by draft-table-report.py): formal Tables become input tables (editable) or
data-model VALUES elements (read-only); summary grids become grouped tables, ordinal
roster x period grids become a cross-joined expansion + pivot, labelled period grids with
running balances become a keyed table (CumulativeSum partitioned by label).

Default = dry run: re-verifies the plan against its recorded source cells, runs the local
parity check, writes dm_spec.json / wb_spec.json / seed CSVs to --out-dir (default
/tmp/table-report-build), no API calls; pending review issues are reported as DRAFT MODE
and do not block. A parity or cell mismatch exits 1 before anything is written.
`--post` builds live (data model POST + readback, workbook dry-run validate, create) and keeps
the strict source-bound `--review-dispositions` contract.

Editable tables are created as EMPTY input tables (structure only); seed them afterwards
with the seed_*.csv files through Sigma MCP Code Mode (see refs/input-tables.md).
"""
import argparse
import importlib.util
import json
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_LIB = _HERE / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
import workbook_wire  # noqa: E402
from lib import table_report_model as trm  # noqa: E402
from lib.migration_gate import MigrationGateError, gate_plan_source  # noqa: E402


def _forecast_helpers():
    """env/token/api from build-forecast-model.py (same auth + HTTP contract)."""
    spec = importlib.util.spec_from_file_location("build_forecast_model_helpers", _HERE / "build-forecast-model.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("plan", help="table-report plan JSON")
    parser.add_argument("--source", help="source workbook used to produce the plan")
    parser.add_argument("--review-dispositions", help="source-bound review dispositions JSON")
    parser.add_argument("--post", action="store_true", help="create the data model and workbook")
    parser.add_argument("--out-dir", default="/tmp/table-report-build",
                        help="where specs and seed CSVs are written (default: /tmp/table-report-build)")
    parser.add_argument("--parity-tolerance", type=float, default=0.01,
                        help="absolute tolerance for plan cells and parityTargets (default: 0.01)")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    with open(args.plan, encoding="utf-8") as stream:
        plan = json.load(stream)
    try:
        source_path, gate_report = gate_plan_source(
            plan, source_arg=args.source, dispositions_path=args.review_dispositions,
            plan_path=args.plan, draft=not args.post)
    except MigrationGateError as error:
        print(f"MIGRATION GATE: {error}", file=sys.stderr)
        return error.exit_code
    print(f"MIGRATION GATE: {gate_report['status']} "
          f"source={source_path.name} sha256={gate_report['sourceSha256']}")
    if gate_report.get("pendingReview"):
        print("DRAFT MODE: review pending for " + ", ".join(gate_report["pendingReview"])
              + " - fine for a local build; resolve before --post.")

    # parity: every recorded source cell, then the headline targets - a mismatch stops the run
    try:
        trm.validate_plan(plan)
        model = trm.compute(plan)
    except (trm.PlanError, trm.EvalError) as error:
        print(f"PLAN ERROR: {error}", file=sys.stderr)
        return 1
    checked, problems = trm.check_cells(plan, model, args.parity_tolerance)
    if problems:
        print(f"LOCAL PARITY FAILED - {len(problems)} of {checked} source cells differ; no specs written:",
              file=sys.stderr)
        for problem in problems[:10]:
            print("  " + problem, file=sys.stderr)
        return 1
    print(f"LOCAL PARITY: {checked}/{checked} source cells reproduced from the plan's table data")
    ok, lines = trm.parity_lines(plan, model, args.parity_tolerance)
    for line in lines:
        print(line)
    if not ok:
        print("LOCAL PARITY FAILED - no specs written; fix the plan (or the source) before building.",
              file=sys.stderr)
        return 1

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        seeds = trm.seed_csvs(plan, out_dir)
        dm_spec = trm.build_dm_spec(plan)
        if dm_spec is not None:
            (out_dir / "dm_spec.json").write_text(json.dumps(dm_spec, indent=1), encoding="utf-8")
        elif (out_dir / "dm_spec.json").exists():
            (out_dir / "dm_spec.json").unlink()
        if not args.post:
            wb_spec = workbook_wire.wire_workbook(trm.build_wb_spec(plan, "<dm-id>", trm.local_dm_elements(plan)))
            (out_dir / "wb_spec.json").write_text(json.dumps(wb_spec, indent=1), encoding="utf-8")
    except trm.PlanError as error:
        print(f"PLAN ERROR: {error}", file=sys.stderr)
        return 1
    print("seed CSVs (paste/seed into the empty input tables): " + (", ".join(p.name for p in seeds) or "none"))
    if dm_spec is None:
        print("no read-only tables: no data model is needed (every table is an editable input table)")
    if not args.post:
        written = "dm_spec.json, wb_spec.json" if dm_spec is not None else "wb_spec.json"
        print(f"DRY RUN - specs written to {out_dir}/ ({written}). Re-run with --post to build live.")
        return 0

    helpers = _forecast_helpers()
    env = helpers.env()
    base, tok = env["SIGMA_BASE_URL"], helpers.token(env)
    dm_id, elements = None, {}
    if dm_spec is not None:
        dm = helpers.api(base, tok, "POST", "/v2/dataModels/spec", dm_spec)
        dm_id = dm["dataModelId"]
        print("dataModelId:", dm_id)
        full = helpers.api(base, tok, "GET", f"/v2/dataModels/{dm_id}/spec")
        try:
            elements = trm.map_dm_elements(plan, full)
        except trm.PlanError as error:
            print(f"PLAN ERROR: {error}", file=sys.stderr)
            return 1
    wb_spec = workbook_wire.wire_workbook(trm.build_wb_spec(plan, dm_id, elements))
    (out_dir / "wb_spec.json").write_text(json.dumps(wb_spec, indent=1), encoding="utf-8")
    checked_wb = helpers.api(base, tok, "POST", "/v2/workbooks", workbook_wire.dry_run_request(wb_spec))
    workbook_wire.assert_dry_run_success(checked_wb)
    wb = helpers.api(base, tok, "POST", "/v2/workbooks", wb_spec)
    print("workbookId:", wb["workbookId"], "| url:", wb.get("url"))
    print("NEXT: seed the input tables from the seed CSVs (Sigma MCP Code Mode), query a few summary "
          "cells against the parity targets, then Publish.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
