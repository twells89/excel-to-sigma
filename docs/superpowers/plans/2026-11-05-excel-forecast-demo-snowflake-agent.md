# Excel Forecast Demo — Snowflake Landing + Agent Enhance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A believable macro-heavy FP&A forecast `.xlsm`, two excel-to-sigma enhancements (Snowflake landing, Sigma Agent Enhance phase), a polished migrated Sigma workbook TJ approves, then a recorded, lower-thirded LinkedIn cut of a real run.

**Architecture:** The source file is shaped so the skill's proven **forecast** route accepts it (`migrate.py plan` → `archetype: forecast`, `draft-forecast-plan.py` reproduces every cached value). The new landing step loads the actuals/plan/dimension tabs into Snowflake with `snow`, and `build-forecast-model.py` gains an option to source GL actuals from that table instead of inline `VALUES`. Editable assumptions stay input tables read directly by the workbook (no warehouse view). Macros become native actions; a new trailing **Phase 6 — Enhance** recommends and authors a workbook agent over those actions.

**Tech Stack:** Python 3 (`unittest`, openpyxl, oletools), Snowflake CLI `snow` 3.16.0, Sigma CLI profile `tj-wells-1989 (1)`, Sigma MCP Code Mode (`sigma_aws_us_west_2_production`), Node + Playwright video pipeline from `~/sigma-git-e2e-video`.

**Spec:** `docs/superpowers/specs/2026-11-05-excel-forecast-demo-snowflake-agent-design.md` (read it first; it carries the why, anti-goals and Claude's calls).

## Global Constraints

- Repo/branch: `twells89/excel-to-sigma`, branch `feat/snowflake-landing-agent-enhance`, worktree `~/excel-to-sigma-agent-demo`. All skill paths below are relative to `skills/excel-to-sigma/`; run scripts from that directory.
- Sigma: `SIGMA_CLI_PROFILE='tj-wells-1989 (1)'`, org `https://app.sigmacomputing.com/tj-wells-1989`, API host `aws-api.sigmacomputing.com`.
- Warehouse connection: **`ymb68310`**, `cb2f5180-641f-47bd-8efa-da9d590d855a` (Snowflake account `ymb68310`). Never `OAuth Snowflake` (`bc0319f8…`).
- **No warehouse views, no Snowflake views over Sigma writeback tables, no stored procedures, no visible workarounds.** If a step needs one, stop and bring it to TJ.
- No Snowflake grants or Sigma permission changes without TJ's OK. No Excel Trust Center changes.
- Never execute VBA. Never send real email.
- Non-`@sigmacomputing` packages: no version released within the last 3 days. No new npm packages for the video (reuse `~/sigma-git-e2e-video/node_modules`).
- Never renumber the skill's phases (0, 0b, 0c, 0d, 1–5). New steps: **Phase 1b — Land source tabs in Snowflake** and trailing **Phase 6 — Enhance**.
- Tests: `python -m unittest discover -s tests -v` from the skill dir (CI runs this). Export `LANG=en_US.UTF-8 LC_ALL=en_US.UTF-8`.
- Subagent models: implementers **Sonnet**; mechanical copy/format **Haiku**; final whole-branch review **Opus**.
- The competing product/post is never named or shown anywhere (code, docs, video).
- Snowflake landing (TJ's decision 2026-11-05): schema **`CSA.TJ`**, `snow` connection **`tj`**. Tables are prefixed: rehearsal `FPA_REH_<TABLE>`, on-camera `FPA_<TABLE>`. After creating each table, grant `SELECT` on that table to role **`SIGMA_ROLE`** (the `cb2f5180` connection's role; same per-table pattern every existing `CSA.TJ` table uses). No other grants, no schema creation, no future grants.
- Sigma folders: rehearsal `Excel Demo — Rehearsal`, on-camera `FP&A Forecast`.

## Review Focus

1. **Re-running landing into a schema that already has the tables** → expect a clear refusal unless `--replace` is passed, never a silent duplicate append. (Test in Task 4.)
2. **Excel headers with spaces/punctuation/mixed case** (`Net Revenue ($)`) → land as stable quoted-safe UPPER_SNAKE identifiers, and the builder's column map uses the same names. (Test in Task 4.)
3. **Excel date cells and serial numbers** in the month column → land as `DATE`, not `NUMBER`/`VARCHAR`. (Test in Task 4.)
4. **A macro inventory with no ACTION/WHATIF classes** (only RECALC/cosmetic) → Enhance recommends an analyze-only agent or none, never invents write tools. (Test in Task 6.)
5. **Blank actual cells** → landed as `NULL` and parity still matches Excel's cached totals (blank-as-zero per `refs/source-blank-validation.md`). (Test in Task 4.)

---

### Task 1: Environment proof

**Files:** none committed. Record findings in `work/demo/env.md` (gitignored `work/`).

- [ ] **Step 1:** Repair `snow`: `brew reinstall snowflake-cli`. Verify: `snow --version` prints `3.16.0` without a traceback.
- [ ] **Step 2:** `sigma -p "tj-wells-1989 (1)" api connections get --params connectionId=cb2f5180-641f-47bd-8efa-da9d590d855a` → record user/role/default database, warehouse, write schema.
- [ ] **Step 3:** Find the `snow` connection that is the same Snowflake user/role (`snow connection list`, then `snow sql -c <name> -q "select current_user(), current_role(), current_database()"`). Expected: one of `tj`/`default` matches Step 2. If none matches or the role can't `CREATE SCHEMA` in that database, **stop and ask TJ**.
- [ ] **Step 4:** (done) Probe Sigma visibility: `sigma -p "tj-wells-1989 (1)" api connections` path lookup (or MCP `search` with `entityTypes:["table"]`, `filters.connectionId=cb2f5180…`). Drop the probe table.
- [ ] **Step 5:** MCP: call `begin_session`; confirm the org URL is `tj-wells-1989`. If not, all Code Mode work uses the CLI/REST path and note it.
- [ ] **Step 6:** `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt`; run the full test suite once. Expected: green baseline (record count).

### Task 2: Feasibility spike (throwaway, gate)

**Files:** `work/demo/spike/` only. Nothing merged.

Prove in a scratch workbook in `Excel Demo — Rehearsal`, on `cb2f5180`:

- [ ] **Step 1:** A DM on a small Snowflake table (`region, month, revenue`) + a workbook input table `Assumptions(scenario, region, growth_pct)` + a `Parameters` input table (`selected_scenario`, `forecast_start_month`). A workbook formula column computes `revenue * (1 + growth_pct)` by looking up the input table on `[region]` and `selected_scenario`. Verify by MCP `query`: editing `growth_pct` changes the result after publish.
- [ ] **Step 2:** A button action `update-rows` on `Parameters.forecast_start_month` (+1 month) and an `insert-rows` action into `Approvals(status, submitted_by, submitted_at)`. Verify both rows via query after firing.
- [ ] **Step 3:** `document.agents` with one Act agent whose tools reference those actions (`requiresApproval: true`) plus a `chat` element; publish; run it via `sigma api workbook-agents` run / MCP `call_agent` with "roll the forecast forward one month". Expected: approval requested, then the Parameters row changes. Clone shapes from the in-org **Planning copilot** workbook `85e92ad3-346a-41e5-89e5-89b5849f4953` (`sigma api workbooks spec get`).
- [ ] **Step 4:** Write `work/demo/spike/FINDINGS.md`: the exact working spec fragments (input-table lookup formula, action shapes, agent shape). **Gate:** if any of Steps 1–3 can't be done natively, stop and report to TJ with the redesign options. Do not proceed.

### Task 3: Demo source workbook

**Files:**
- Create: `scripts/make-demo-forecast.py`
- Create: `scripts/demo-forecast-macros.bas`
- Create: `tests/test_demo_forecast_source.py`
- Output (not committed): `work/demo/source/<Company>_FY27_Rolling_Forecast_v7_FINAL.xlsx` and `.xlsm`

**Interfaces:**
- Produces: `build_demo_workbook(out: Path, seed: int = 27) -> Path` (writes `.xlsx`); CLI `make-demo-forecast.py --out <dir>`. Tab names exactly: `Instructions`, `Actuals`, `Plan`, `Dimensions`, `Assumptions`, `Forecast`, `Variance`, `Versions`, `Approvals`. Formal Tables: `tblActuals`, `tblPlan`, `tblRegions`, `tblProducts`, `tblDrivers`, `tblVersions`, `tblApprovals`.
- VBA procedures (exact names): `RefreshActuals_Click`, `SetScenario`, `RollForwardMonth_Click`, `SnapshotVersion_Click`, `SubmitForApproval_Click`, `RecalcAll`, `FormatReport`.

The `Forecast` tab must be the P&L-in-cells shape `lib/forecast_plan.py` accepts: section rows (`Revenue`, `Cost of Revenue`, `Operating Expenses`), line rows per region (4 regions), month columns `Jan-26 … Dec-27`, actual months = `SUMIFS` into `Actuals`, forecast months = `=<prior>*(1+Assumptions!<cell>)`, derived rows (`Gross Margin`, `EBITDA`). Model it on `tests/forecast_pnl_fixture.py` but bigger and believable (real-looking numbers, seasonality, an `Instructions` tab with "Click Refresh Actuals first…", buttons drawn as shapes labeled with the macro names). Actuals ≈ 18 months × 4 regions × 3 product lines.

- [ ] **Step 1: Write failing tests** in `tests/test_demo_forecast_source.py`:
  - `test_routes_as_forecast`: run `migrate.py plan` on the generated `.xlsx` → `archetype == "forecast"`, `auditStatus != "BLOCKED"`.
  - `test_drafter_reproduces_cached_values`: `draft_plan(src, intent="what-if", connection="<c>", folder="<f>")` raises nothing; `shape["lines"] >= 12`.
  - `test_macro_classes`: `macro-classify.py --json scripts/demo-forecast-macros.bas` classifies: `RefreshActuals_Click`→EXTERNAL_DATA_REFRESH, `SetScenario`→WHATIF_SCENARIO, `SnapshotVersion_Click`→ACTION_WRITEBACK, `SubmitForApproval_Click`→ACTION_WRITEBACK (email noted), `RollForwardMonth_Click`→ACTION_WRITEBACK or TRANSFORM_ROLLUP, `RecalcAll`→RECALC, `FormatReport`→FORMAT_COSMETIC.
- [ ] **Step 2:** Run → FAIL (module missing).
- [ ] **Step 3:** Implement the generator and `.bas`. Iterate on the source shape (not on the drafter) until the drafter accepts it; if it truly needs a drafter change, make it a separate, tested commit.
- [ ] **Step 4:** Run the tests → PASS; full suite still green.
- [ ] **Step 5: Produce the `.xlsm`.** It must carry *these* macros, so a fixture's VBA project can't be reused. Use Excel via AppleScript only if it needs no Trust Center change; otherwise **hand off to TJ**: open the `.xlsx` in Excel, Developer → Visual Basic → File → Import `demo-forecast-macros.bas`, assign buttons, save as `.xlsm` into `work/demo/source/`. Then re-run `migrate.py plan` on the `.xlsm` → macros section lists all 7.
- [ ] **Step 6:** Commit generator, `.bas`, tests.

### Task 4: Phase 1b — Land source tabs in Snowflake

**Files:**
- Create: `scripts/lib/snowflake_landing.py`, `scripts/land-in-snowflake.py`
- Modify: `scripts/build-forecast-model.py` (actuals source option)
- Test: `tests/test_snowflake_landing.py`, add a case to `tests/test_builders.py`

**Interfaces:**
- Produces:
  - `snowflake_identifier(header: str) -> str` (UPPER_SNAKE, `[A-Z0-9_]`, leading digit prefixed `C_`; collisions → error naming both headers).
  - `infer_column_types(rows: list[dict], headers: list[str]) -> dict[str, str]` → `DATE` | `NUMBER(38,6)` | `VARCHAR`; Excel `datetime`/`date` and month-header serials → `DATE`.
  - `landing_plan(source: Path, tables: list[str], database: str, schema: str, prefix: str) -> dict` (Snowflake table = prefix + identifier of the Excel table name minus `tbl`, e.g. `FPA_REH_ACTUALS`) → `{"tables":[{"excelTable","snowflakeTable","columns":[{"header","name","type"}],"rowCount","csv"}]}` (writes one CSV per table next to the plan).
  - `snow_commands(plan: dict, connection: str, replace: bool, grant_role: str | None) -> list[list[str]]` — argv lists per table: `CREATE TABLE` (or `CREATE OR REPLACE TABLE` only when `replace`), `PUT file://… @%TABLE`, `COPY INTO`, `GRANT SELECT ON TABLE <FQN> TO ROLE <grant_role>` (only when given), then a `SELECT COUNT(*)` check. Never `CREATE SCHEMA`.
  - CLI: `land-in-snowflake.py <file> --tables tblActuals,tblPlan,tblRegions,tblProducts --database <DB> --schema <SCHEMA> --table-prefix <PREFIX> --snow-connection <name> [--grant-select-to <ROLE>] [--replace] [--execute] --out <workdir>/landing.json`. Default prints the plan + commands (dry run); `--execute` runs them, then verifies each row count equals `rowCount` and exits non-zero on mismatch. Existing table without `--replace` → exit 2 with message `table <FQN> exists; pass --replace to overwrite`.
  - `build-forecast-model.py --actuals-table <DB.SCHEMA.TABLE> --landing <workdir>/landing.json`: GL actuals element becomes a warehouse-table source on the plan's connection (path from landing.json) instead of inline `VALUES`; column names come from landing.json. Without the flag, behavior is unchanged.
- [ ] **Step 1: Failing tests** (`tests/test_snowflake_landing.py`): `test_identifier_normalizes` (`"Net Revenue ($)"`→`NET_REVENUE`, `"2026 Plan"`→`C_2026_PLAN`, collision raises); `test_month_column_lands_as_date` (datetime cells and an Excel serial `46023` with date format → `DATE`); `test_blank_actuals_become_null` (blank cell → empty CSV field, column stays `NUMBER`); `test_commands_refuse_existing_without_replace` (plan + `replace=False` → no `CREATE OR REPLACE`; CLI with a fake `snow` that reports the table exists → exit 2 and the message above); `test_row_count_mismatch_fails` (fake `snow` returns a wrong count → non-zero exit). Builder: `test_forecast_actuals_from_landed_table` — dry-run spec has no `VALUES` for actuals, references the landed table path, and local parity totals equal the inline-VALUES build.
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3:** Implement. Fake `snow` in tests = a tiny script on `PATH` in a temp dir (pattern: see `tests/test_inventory_actions_cli.py` for CLI tests).
- [ ] **Step 4:** Run → PASS; full suite green.
- [ ] **Step 5: Live check** into `CSA.TJ` with `--table-prefix FPA_REH_ --grant-select-to SIGMA_ROLE --execute`; row counts match; after `sigma api connections sync` on each path, Sigma sees the four tables on `cb2f5180`. Add `test_grant_emitted_only_when_requested` to Step 1.
- [ ] **Step 6:** Commit.

### Task 5: Macros → native workflows

**Files:**
- Modify (only as needed): `scripts/lib/native_actions.py`, `scripts/macro-actions.py`
- Create: `scripts/lib/demo_workflows.py` if the four demo workflows don't fit existing builders
- Test: `tests/test_native_actions.py` (new cases)

**Interfaces:**
- Consumes: Task 2 `FINDINGS.md` shapes; Task 3 macro names; existing `forecast_snapshot` (CommitForecast replacement) and `forecast_scenario` components.
- Produces: an action manifest the workbook build wires to buttons and that Task 6 reads: `[{"macro": str, "class": str, "actionId": str, "label": str, "effect": "update-rows"|"insert-rows"|"control", "target": {"elementId": str, "columns": [str]}, "writes": bool}]`, written to `<workdir>/workflows.json`.

Mapping (from the spec's calls): `SetScenario` → `Parameters.selected_scenario` (a one-row Parameters input table rather than a workbook control, because the skill documents that a control filters but cannot feed a formula parameter; the header shows it as a picker); `RollForwardMonth_Click` → `update-rows` `Parameters.forecast_start_month` +1 month; `SnapshotVersion_Click` → `insert-rows` into `Versions` (reuse `forecast_snapshot` freezing); `SubmitForApproval_Click` → `insert-rows` into `Approvals` (`status='Submitted'`, user, `now()`) and the header shows the latest status; `RefreshActuals_Click`/`RecalcAll` → dispositioned as replaced by live Snowflake/recalc; `FormatReport` → styling.

- [ ] **Step 1: Failing tests:** `test_demo_manifest_routes` (the classified demo `.bas` → manifest has exactly 4 entries with `writes` true for Roll/Snapshot/Submit and false for SetScenario; RefreshActuals/RecalcAll/FormatReport are dispositioned, not dropped silently — they appear under `dispositions`); `test_submit_never_emails` (no `call-api`/email effect in the manifest).
- [ ] **Step 2:** Run → FAIL. **Step 3:** Implement. **Step 4:** PASS + suite green.
- [ ] **Step 5:** Commit.

### Task 6: Phase 6 — Enhance (Sigma Agent recommendation)

**Files:**
- Create: `scripts/lib/agent_enhance.py`, `scripts/recommend-agent.py`, `refs/agent-enhance.md`
- Test: `tests/test_agent_enhance.py`

**Interfaces:**
- Consumes: macro-classify JSON; `<workdir>/workflows.json` (Task 5); workbook element ids from the built workbook.
- Produces:
  - `recommend_agent(macros: dict, workflows: list[dict]) -> dict` → `{"recommend": bool, "profile": "act"|"analyze"|None, "reasons": [str], "tools": [{"actionId","label","requiresApproval": bool}], "pitch": str}`.
  - `agent_document_fragment(rec: dict, name: str, grounding_element_ids: list[str]) -> dict` → `{"agents":[…], "chatElement": {…}}` in the exact shapes proven in Task 2.
  - CLI: `recommend-agent.py --macros macros.json --workflows workflows.json [--emit --workbook <id>]`: prints the recommendation (the on-camera moment: plain language, e.g. *"3 of your macros were workflows people trigger by hand. A Sigma Agent can run them on request, with approval before each write."*); `--emit` writes the fragment for the build to merge.
- [ ] **Step 1: Failing tests:** `test_act_profile_when_write_workflows` (demo inputs → `profile=="act"`, 3 write tools all `requiresApproval: True`, SetScenario tool `requiresApproval: False`); `test_analyze_only_when_no_actions` (macros only RECALC/FORMAT → `profile in ("analyze", None)`, `tools == []`); `test_no_macros_no_recommendation` (`.xlsx` with no VBA → `recommend False`); `test_fragment_shape_matches_spike` (fragment keys/required fields equal the Task 2 shapes, checked against a fixture copied from FINDINGS).
- [ ] **Step 2:** FAIL. **Step 3:** Implement. **Step 4:** PASS + suite green.
- [ ] **Step 5: Live:** merge the fragment into the rehearsal workbook (MCP Code Mode or `workbooks spec` PUT), publish, run the scripted prompt from the spec via `call_agent`; verify the driver change, snapshot row and approval row by query.
- [ ] **Step 6:** Commit.

### Task 7: Skill docs + router

**Files:**
- Modify: `SKILL.md` (Prerequisites `snow` section rewritten for landing; new `### Phase 1b — Land source tabs in Snowflake`; new `## Phase 6 — Enhance: recommend a Sigma Agent`; scriptability matrix rows for landing + agents; "does NOT do" updated), `refs/macro-handling.md` (link to Enhance), `scripts/migrate.py` (forecast route: landing + `--actuals-table` in `nextCommands`; `recommend-agent.py` step when macros exist; remove "Paste seed CSVs" handoff when the seed path is MCP-scripted)
- Test: `tests/test_migrate_router.py`, `tests/test_packaging.py` if it lists scripts
- [ ] **Step 1: Failing test:** `test_forecast_route_offers_landing_and_enhance` (macro-enabled forecast source → `nextCommands` contains `land-in-snowflake.py` and `recommend-agent.py`; non-macro `.xlsx` → no `recommend-agent.py`).
- [ ] **Step 2:** FAIL. **Step 3:** Implement + docs. **Step 4:** PASS + suite green.
- [ ] **Step 5:** Commit.

### Task 8: Rehearsal migration + presentation polish → TJ approval

**Files:** `work/demo/rehearsal/` (run log, plan.json, landing.json, workflows.json, screenshots). Any builder/presentation fixes are tested commits.

- [ ] **Step 1:** Run the skill exactly as on camera, from `migrate.py plan` through Phase 6, into `Excel Demo — Rehearsal` / `CSA.TJ` prefix `FPA_REH_`. Save a full terminal transcript (`script -q work/demo/rehearsal/run.log`).
- [ ] **Step 2:** Parity: the skill's local + live parity gates pass (forecast revenue, Gross Margin, EBITDA tie to Excel cached values).
- [ ] **Step 3: Look and feel** per spec pages (Forecast / Assumptions / Versions & approvals; agent chat on Forecast). Use the `sigma-workbooks` skill and `scripts/build-presentation.py`/layout-last. Screenshot every page (MCP `screenshot`) and review against: finance planning app, light theme, brand-aligned, no `Unknown column`/`Invalid Query`, no empty surfaces.
- [ ] **Step 4:** List every place the run showed a hack, a retry, an error, or a manual UI step. Each one is fixed or redesigned, then re-run Step 1. Exit criterion: a clean run with zero such items.
- [ ] **Step 5: HARD GATE — TJ approves the built workbook** (send link + screenshots). No recording before a yes.

### Task 9: PR

- [ ] **Step 1:** Full suite green; `git push -u origin feat/snowflake-landing-agent-enhance`; `gh pr create --repo twells89/excel-to-sigma` with summary of Phase 1b + Phase 6 + demo generator, test evidence, and the rehearsal parity result. Bind it with ccd_pr tools.
- [ ] **Step 2:** Final whole-branch review (Opus subagent); address findings.

### Task 10: Video

**Files:** new dir `~/sigma-excel-migration-video/` copied from `~/sigma-git-e2e-video` (`rec.mjs`, `overlay.mjs`, `shellify.py`, `shell_render.mjs`, `assemble.py`, `vo.py`, `talk_track.py`, `kit/`, `brand/`, `node_modules` symlink). New `scenes/*.mjs`, `vo_script.json`, `TALK_TRACK.md`.

Scene list (≈3:30): `01-title` → `02-excel` (the `.xlsm`: tabs, buttons, `v7_FINAL` name) → `03-inventory` (terminal: `migrate.py plan`, macro routes) → `04-gate` (disposition gate) → `05-land` (terminal: `land-in-snowflake.py --execute`, row counts) → `06-build` (terminal build + Sigma workbook appearing) → `07-workflows` (buttons: roll forward, snapshot, submit; rows appear) → `08-agent` (terminal `recommend-agent.py` pitch, then the chat doing the scripted request with approvals) → `09-close` (governed workbook, versions/approvals, end card). Lower-thirds per phase: *Inventory · Macro gate · Land in Snowflake · Build · Workflows · Agent*.

- [ ] **Step 1:** Draft `vo_script.json` + `TALK_TRACK.md` (lines carry the "bring the spreadsheet to the warehouse" point; never name the competitor). **TJ OKs the talk track.**
- [ ] **Step 2:** Excel scene capture: `screencapture -v -l <windowId>` of the Excel window if Screen Recording permission already exists; otherwise TJ records 15 s himself.
- [ ] **Step 3:** On-camera run into `FP&A Forecast` / `CSA.TJ` prefix `FPA_` (fresh), transcript captured for `shellify`; Sigma scenes recorded with `rec.mjs` against the resulting workbook. Agent take retried until clean.
- [ ] **Step 4:** `assemble.py` → `out/excel-to-sigma-preview-silent.mp4`, `out/excel-to-sigma-preview-guide.mp4` (scratch voice), and the final `TALK_TRACK.md` with times. Verify: duration 3:00–4:15, every lower-third on its scene, no frame shows the competitor, a credential, or an error.
- [ ] **Step 5:** Hand TJ the three files.
