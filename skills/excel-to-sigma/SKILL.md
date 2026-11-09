---
name: excel-to-sigma
description: >-
  Convert an Excel (.xlsx) workbook — especially a planning / budget / forecast
  data-entry model — into a Sigma data model + workbook. Use when the user has
  an Excel file (formal Tables, pivots, charts, or a cell-formula budget model)
  and wants to recreate it in Sigma. Covers formal-Table discovery, Excel-formula
  translation, and the input-table → warehouse-view → data-model pattern that
  preserves data entry. Requires Sigma client credentials and (for the data-entry
  half) a write-enabled Sigma connection.
---

# Excel → Sigma Conversion

> **STATUS: input-table data-entry path VALIDATED end-to-end (2026-06-08).** The
> full chain — Excel formal Table → Sigma input table → warehouse view → data
> model reading `FROM` that view → section rollups — was proven with **exact
> parity** on a representative 540-row forecast model. The read/output half (DM +
> workbook from a landed seed) was validated earlier (2026-06-03). The discovery
> + formula-translation layer is not a general one-shot converter; two archetypes
> are the exception, each with a router gate, a plan drafter that verifies every
> formula cell against Excel's cached values, and a builder with local parity:
> **forecast** (cell-drawn P&L, any sections, month/quarter/week periods) and
> **table-report** (formula grids over formal Tables). Both were validated live in
> Sigma on 2026-10-22 with exact parity. Shapes outside them fail loudly.
> Complex-workbook support means **inventory + explicit routing + parity**, not a
> promise that arbitrary Excel programs convert automatically.

**Start here — do not read every ref.** Run the router first; it picks the archetype
and names the refs, decisions and commands for this workbook:

```bash
python scripts/migrate.py plan <file.xlsx|.xlsm> --out-dir <workdir>
```

It runs the audit, discovery and (for macro files) macro classification, writes
`migration-audit.json`, `discovery.json` and `review.template.json` to `<workdir>`,
and prints JSON: `auditStatus`, `archetype`, `readRefs`, `beforePost`, `askUser`,
`nextCommands`, `optIns`, `userHandoffs`, `routeReason`. Exit code 2 means `BLOCKED`. Read the refs
it lists, ask the user the `askUser` questions, then run the `nextCommands` in order.
On the forecast route, landing the source Tables in Snowflake (Phase 1b) and the workflow
planning app + agent (Phase 6) are yes/no questions in `askUser`: by default `nextCommands`
is the plain build, and each offered option is listed in `optIns` with its opted-in commands
(or re-run `plan` with `--land-tables` / `--enhance`). The planning app is offered only when a
non-AUTO macro is an action write-back, external I/O or what-if scenario macro. Pass
`--intent read-only|editable|what-if` to fill the drafter's intent; a read-only build without
the planning app has no input tables, so it writes no setup program.
Routes are tried in order and each is named only when its drafter accepts the
workbook: research template → `forecast` → `table-report` → `input-table`. A
SUMIFS/driver workbook both drafters reject routes to `cell-model` (source
inventory, then a source-specific Sigma MCP build) with the drafters' reasons in
`routeReason`.

| When | Read |
|---|---|
| Always | `refs/sigma-build-gotchas.md` — SQL element formula prefix, DM POST envelope, `contents` lifecycle, id reassignment, publish gate |
| Before `--post` or declaring completion | `refs/coverage-contract.md` — the fail-loud contract and hard completion gates |
| Formal Table to input table (the data-entry path) | `refs/input-tables.md` (the source-swap goes through a warehouse view, not a direct DM binding), `refs/excel-translation.md` |
| Cell-drawn planning / forecast model (no formal Table) | `refs/model-taxonomy.md`, `refs/forecast-recipes.md` |
| Formula grids over formal Tables (pipeline by rep, headcount plan, inventory projection, 13-week cash flow) | `refs/table-report-recipes.md` |
| Broker / FactSet research model, or a fleet off one template | `refs/research-recipes.md`, `refs/research-template-coa.md` |
| `.xlsm` / macros | `refs/macro-handling.md` |
| Charts, validation, conditional formatting, formats to rebuild | `refs/presentation-migration.md`; `refs/generic-source-presentation.md` for the generic presentation inventory |
| Live Sigma API capability questions (Actions, Python, procedures) | `refs/code-rep-capabilities.md` |
| Source cell/formula inventory or formula-level native calculations | `refs/source-workbook-model.md`, `refs/workbook-expression-context.md`, `refs/source-inputs-and-reshape.md`, `refs/native-calculations.md` |

**Maintainer-only — skip for a customer migration.** These document fixture
acceptance machinery and reviewed canary packets, not the migration path:
`refs/acceptance-and-promotion.md`, `refs/forecast-scenario.md`,
`refs/forecast-snapshot.md`, `refs/custom-ports.md`, `refs/source-blank-validation.md`,
`refs/regional-grouped-canary.md`. Read one only when the user names that
workflow or you are changing it.

**Draft vs production.** Builds without `--post` are *draft mode*: local specs and
a local parity check, with audit review issues reported as pending
(`DRAFT MODE: review pending for ...`) instead of blocking. `--post` is
*production*: it keeps the full source-bound `--review-dispositions` contract
(Phase 0). `BLOCKED` stops both. Never fabricate `verified: true` to get past a
draft; draft mode exists so you do not need to.

**UI steps.** When the work reaches a step only the user can do in the Sigma UI
(CSV paste, Publish, Create warehouse view), collect every pending UI action into
one checklist, give it to the user, and stop. Do not interleave them.

The `sigma-workbooks` companion skill, if installed, plus the Sigma OpenAPI, is the
canonical workbook spec; find it through the agent's skill loader, and use
`refs/sigma-build-gotchas.md` when it is unavailable.

---

## The one big idea

Excel planning models are **data-entry tools**, not dashboards. Forecasters type
numbers into budget sheets; formulas roll them into an income statement. The
faithful Sigma translation keeps that data entry alive — which is exactly what
**input tables** are for.

The decisive structural insight (validated): in well-structured enterprise
workbooks, the value lives in **formal Excel Tables** at a tidy grain. Most of the
thousands of `INDEX/MATCH/SUMIFS` formulas are **wide→long reshaping**, not
business logic — they *evaporate* when the data lands at that grain in a Sigma DM.
What's left ("report" sheets drawn in cells) is rebuilt by hand, same as any
migration.

So the migration is two halves:

1. **Read / output half** — the formal Table(s) become DM elements; the income-
   statement SUMIFS become grouped DM metrics + workbook charts. (Build against a
   seed first; see `refs/excel-translation.md`.)
2. **Data-entry half** — the formal Table that forecasters edit becomes a Sigma
   **input table**; the DM re-points to read from it. (See `refs/input-tables.md`.)

---

## Preserve the inputs — never ship a read-only port

**A spreadsheet is an app people type into. If the conversion is read-only, you've
demoted it.** Don't default every table to a read-only DM element. Classify each
source table by *how it's used* and route accordingly:

| Source table is… | → Sigma | Editable? |
|---|---|---|
| **Entered** (users type values: budget inputs, a tracker, subscriptions) | **input table** (empty / CSV / linked) | ✅ like Excel/Sheets |
| **Derived** (formula rollups: the income statement, SUMIFS summaries) | read-only DM element + metrics | ❌ recomputes |
| **Reference dim** (lookup lists users maintain) | input table if maintained, else read-only | depends |
| **Mixed** (literal source columns beside formula columns, or literal overrides within formulas) | route literal cells and live calculations separately; review overrides and edit intent | depends |

Phase 0 records literal and formula provenance by column and cell. Route an
all-literal table as a candidate input source and an all-formula table as a
computed candidate. A mixed Table is not wholly derived or wholly editable:
retain its literal inputs and calculate its formula columns live. Ask how users
edit ambiguous cells before selecting the input surface.

### What's actually buildable via API today (validated 2026-06-10 by MCP query)

> **Hard-won correction.** A linked input table authored through the REST workbook
> code representation has **inherited columns that don't resolve** — every row shows
> **"multiple values"** (publish doesn't fix it; re-POSTing a known-good UI-built spec
> reproduces the break), so REST linked tables stay unsupported. **Sigma MCP Code Mode
> linked tables are validated live (2026-10-22):** `scripts/linked-input-program.py`
> emits the 3-call create / seed / edit programs. See `refs/input-tables.md`.

So, honestly, by usage pattern:

| Need | API today? | How |
|---|---|---|
| **Edit-in-place of the migrated rows** (the common spreadsheet case) | ✅ with Sigma MCP Code Mode | create the empty input-table **structure** with REST, then pre-query and seed it through MCP `InputTable.upsert_rows()`; `update_rows()` targets returned/system row IDs. UI CSV remains the fallback. |
| **Net-new entry at a grain** (blank forecast grid) | ✅ via MCP Code Mode | REST linked-off-spine creation still leaves context columns as "multiple values" — do not use it. Build the linked table through MCP: `scripts/linked-input-program.py` (validated live 2026-10-22, one row per key). |
| **Augment / annotate existing rows** (notes/flags beside live data) | ✅ via MCP Code Mode | same MCP linked-table recipe (editable columns beside inherited source columns); published-mode edits by non-editors and warehouse views over the table remain untested |

REST can scaffold input-table *structure* and build read-only DMs. Sigma MCP
Code Mode can seed and update rows in an existing empty input table and can create,
seed and edit **linked** input tables (validated live 2026-10-22; this skill's
`scripts/linked-input-program.py` only emits the programs, it never calls MCP).
Warehouse-view creation still requires UI/live review.

Design so the editable surfaces exist from day one, and seed them through MCP or
the UI. **Don't build a read-only model and call it a migration.**

---

## Prerequisites

### Sigma access
Sign in with OAuth first; client credentials are only the unattended fallback. Run
these from the skill directory. Every Sigma-calling script resolves auth through
`scripts/lib/sigma_auth.py`:

```bash
export SIGMA_BASE_URL='https://<your-sigma-api-host>'   # e.g. https://aws-api.sigmacomputing.com
eval "$(bash scripts/browser-login.sh)"                  # browser OAuth (PKCE); refresh token stays in the OS keychain
# or, with the Sigma CLI:  sigma auth login -p <name>   then pass --sigma-profile <name> (or SIGMA_CLI_PROFILE=<name>)
python scripts/build-forecast-model.py plan.json --post --sigma-profile <name>   # example
python scripts/get_token.py --workdir <dir>              # shell-neutral auth.json handoff (0600)
```
With a CLI profile, `export SIGMA_CLI_PROFILE='<profile>'` once (off camera) so no command has to
carry `--sigma-profile` (`land-in-snowflake.py`'s connection sync reads it too).

Look up the target folder and the write-enabled connection by name, read-only, instead of
filtering a `sigma api files list` by hand:

```bash
python scripts/migrate.py resolve-targets --folder-name '<folder name>' --connection-name '<connection name>'
# prints folderId, connectionId and draftFlags ("--connection <id> --folder <id>") for draft-forecast-plan.py;
# a missing or ambiguous name lists the candidates; a connection without writeAccess is refused
```

Resolution order: `SIGMA_API_TOKEN` (+ `SIGMA_BASE_URL`) in the environment, then
`<SIGMA_WORKDIR>/auth.json` / `./auth.json`, then the Sigma CLI profile
(`--sigma-profile` / `SIGMA_CLI_PROFILE`), then the provider (`get_token.py`: keychain
browser session, then `SIGMA_CLIENT_ID`/`SIGMA_CLIENT_SECRET` from the environment or
`~/.sigma-migration/env`). Tokens last ~1h: they refresh after 50 minutes when the mint
time is known and once on a 401. Details: `scripts/AUTH_RUNTIME.md`,
`refs/browser-oauth-login.md`.

The **data-entry half requires a write-enabled connection** (input tables write
back to a `SIGDS_` schema). Confirm via `GET /v2/connections` → `writeAccess: true`.
Verified write connection on tj-wells-1989: `cb2f5180-…` (Snowflake ymb68310).

### Python
From this skill directory, install its pinned dependencies:
`python3 -m venv .venv && .venv/bin/pip install -r requirements.txt`.
Then run every command with `.venv/bin/python`. This file and `refs/` write commands as
`python scripts/…`; `migrate.py` prints them with the interpreter that ran it, so
`.venv/bin/python scripts/migrate.py plan …` prints `.venv/bin/python scripts/…` commands.

### Optional: `snow` CLI
Used by **Phase 1b** (`scripts/land-in-snowflake.py`) to land the workbook's Tables in
Snowflake, and to look up the full path of a Sigma-created warehouse view
(`SIGMA_WRITE_DB.<schema>.<view>`) when the UI truncates it. Sign in once with
`snow connection add` / SSO browser auth, then name the connection with
`--snow-connection <name>`. The landing script runs `snow sql` with an argv list (no shell),
is a dry run by default, and issues only `CREATE TABLE`, the load, row-count verification and
(when asked) `GRANT SELECT` on each landed table. It never creates schemas, views, stored
procedures or future grants. View lookup:
`snow sql -c <conn> --format json -q "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME FROM <DB>.INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME ILIKE '<view>%'"`.

---

## Phase 0 — Audit, then discover

The capability audit is a hard gate before conversion:

```bash
python scripts/audit-xlsx.py <file.xlsx|file.xlsm> --out migration-audit.json
python scripts/audit-openapi.py --out sigma-code-rep-audit.json
```

`AUTO` can proceed. `REVIEW_REQUIRED` is reported, not fatal: `audit-xlsx.py` exits 0
for it (pass `--strict` to exit 1 in a production gate) and exits 2 for `BLOCKED`.
Local/draft builds proceed with the review issues pending; every issue must be
routed, implemented, and verified before `--post` or declaring the migration
complete. `BLOCKED` cannot proceed until its blocker is removed. Then run
`scripts/xlsx-discover.py` for the model/grain-specific report (`migrate.py plan`
runs both).

For a source-bound cell, table, and formula-dependency inventory, run
`python scripts/source-workbook.py SOURCE.xlsx --out source-model.json` and read
`refs/source-workbook-model.md`. This offline inventory preserves unresolved
formulas explicitly; it does not approve a migration or change the audit and
production-write gates.

The whole-workbook Excel migration entrypoints—standalone canonical inference,
batch conversion, the direct research-workbook builder, and the source-bound
forecast/research plan builders—re-run this audit against the source bytes.
Production (`--post`, and batch conversion) requires, for `REVIEW_REQUIRED`,
`--review-dispositions review.json`; the only accepted format is source-bound and
issue-specific (`migrate.py plan` writes a `review.template.json` to fill in).
Local forecast/research plan builds without `--post`, and `infer-canonical-formulas.py` /
`build-research-workbook.py` with the opt-in `--draft` flag, run in draft mode and
only report the pending codes:

```json
{
  "sourceSha256": "<64 lowercase hex characters>",
  "issues": [
    {
      "code": "CHARTS_PRESENT",
      "owner": "Named migration owner",
      "disposition": "Implemented as Sigma chart and render-checked",
      "verified": true
    }
  ]
}
```

Every reported issue code needs a nonempty owner and disposition plus
`verified: true`. The hash must match the workbook currently being converted;
copying a review record to a changed workbook fails. A disposition never
overrides `BLOCKED`.

Walk the `.xlsx` (OOXML = a ZIP of XML; `openpyxl` reads it). Produce an inventory:

- **Formal Tables** (`xl/tables/*.xml`) — name, sheet, range, columns. These are
  the migration's Rosetta Stone: each is a DM-element / input-table candidate.
- **Pivots, charts, slicers** — count + anchor; map to `pivot-table` / chart /
  control later.
- **Formula census** — tally functions; flag the untranslatable (`OFFSET`,
  `INDIRECT`, volatile locators) and the "report drawn in cells" sheets.
- **Degenerate-column warnings** — constant / decayed columns (broken fiscal
  defaults, stale lookup spines). **Derive these, don't port them.**

State the inventory back to the user and confirm the **canonical grain** of the
fact Table before building.

### Phase 0b — Macro inventory & disposition gate (if `.xlsm`/`.xlsb`/`.xls`)

If the file is macro-enabled, run the macro pass **before building** — a dropped
macro is a silent loss of behavior (see `refs/macro-handling.md`):

```bash
python scripts/macro-classify.py <file.xlsm>          # report
python scripts/macro-classify.py --json <file.xlsm>   # contract for build steps
```

Each procedure is classified (olevba-extracted VBA) and routed: **AUTO**,
**CONTROL**, **ACTION**, **STORED_PROC**, **PYTHON_REVIEW**, **DATA_REVIEW**, or
**FLAG**. Actions, Python, and stored procedures are now code-representable, but
still require implementation plus live side-effect/parity checks. Surface every
non-AUTO record before build. Never silently omit a macro.

`migrate.py plan` records the gate in `<workdir>/macro-gate.json` (and names it as
`macroGate`): every procedure surfaced, `dropped: []`, the `nonAuto` list, and per procedure
its class, status, Sigma route and disposition (`implement, then verify live (pending)`, or
`drop after formula parity` for AUTO) with `verified: false`. Read the non-AUTO records to the
user from the `macro-classify.py <file>` report; nothing in the gate is marked verified.

### Phase 0c — Model map & intent (for "report drawn in cells" / planning models)

For a cell-drawn report, `xlsx-discover.py` inventories **ACTUAL** aggregations,
**DRIVER_GROWTH**, **MANUAL** literals among formulas, **DEPENDENCY_LINK** direct
cell references, and other **DERIVED** formulas. A direct reference such as
`13WCF!C7 = B23` is a live link to the prior week's ending cash; preserve its
source cell in the plan and recompute it. Do not replace it with a flat constant
or a cached formula result. A flat forecast is a separate, explicit policy that
needs source evidence; the syntax `=prior` alone does not prove one.

**Forecast path:** `migrate.py plan` -> `draft-forecast-plan.py` -> dry-run
`build-forecast-model.py` -> `build-forecast-model.py --post`. The drafter derives the
plan from the workbook: sections are discovered from the sheet with a sign role taken
from the derived rows (Gross Margin = Net Sales − COGS, …); periods may be months,
quarters (`Q3 FY25`) or weeks; forecast cells may be growth, YoY, percent-of-total,
flat, manual or an external link (kept as manual values with an `EXTERNAL_LINK_VALUE`
review issue); rates may come from several driver cells. It fails loudly on anything
else, and recomputes every line, total and derived row per period against Excel's
cached values before writing the plan. It raises `FLAT_CARRY_POLICY` for `=prior`
lines and `DRIVER_FORMULA_FROZEN` for a driver cell that holds a formula (the plan keeps
only its cached rate); the dry run reports it as pending and `--post` needs a source-bound
disposition. The builder's local parity check compares the plan to `parityTargets`
(every derived row's FY or total value) and exits 1 on any mismatch; when the FY
column covers only some periods, the KPI cards cover exactly those periods. Workbook
line and bar charts whose series are the plan's totals or derived rows are recorded
as `sourceCharts` (every point checked against cached values) and the builder emits a
matching Sigma chart over the same period range; the drafter lists any other chart as
not mapped, to rebuild by hand before dispositioning `CHARTS_PRESENT`.

**Table-report path:** `migrate.py plan` -> `draft-table-report.py` -> dry-run
`build-table-report.py` -> `build-table-report.py --post`. For workbooks whose
outputs are formula grids over formal Excel Tables (SUMIFS/COUNTIFS summaries by
row label and column header or date window, XLOOKUP/INDEX-MATCH lookups, running
balances, row × period expansions) and for period ledgers such as a 13-week cash flow
(table-sourced lines, a named-scenario factor, growth and cadence lines, beginning and
ending cash, running minimum). Tables people edit become input tables (seed them
through Sigma MCP), read-only tables become data-model data, and every summary cell is
recomputed from the table data and compared with Excel's cached value before the plan
is written. See `refs/table-report-recipes.md` for the Sigma shapes (SumIf wrapped in
`Coalesce(…, 0)`, Lookup, CumulativeSum by key, constant-key cross join).
The Report page keeps the source's presentation: Excel number formats become explicit Sigma
formats on every numeric column (never an inherited `%`), summary rows keep the sheet's order, the
header fill / bold and the title cell carry over, source charts that read verified frame cells are
emitted as native charts, and a header band + KPI cards (totals and a ratio of totals) lead the page;
every Report-page table / pivot uses the presentation table preset (and is sized for its taller rows); KPI totals and chart points are re-verified against the source cells before a spec is written
(`refs/table-report-recipes.md`, Presentation).
To match an existing Sigma workbook's look (theme, card styles, KPI type, text type, table preset), run
`scripts/style-reference.py <workbookId or URL> --out look.json` (one read-only GET) and draft with
`draft-table-report.py ... --look look.json`; the arrangement stays ours and Excel's formats, row
order and titles still win; under a look the reference's theme styles the table headers (only Excel's bold is kept, not its
fill) (`refs/table-report-recipes.md`, "Matching the look of an org workbook").

Present the detected map and confirm edit intent, source freshness, and formula
dependencies before selecting a build. The live-actuals/driver-growth union in
`refs/forecast-recipes.md` applies only when those source roles and grain are
present. Archetype builders have source binding **and shape preconditions**:
unknown sheets, tables, sections, or periods require source-specific MCP
construction or explicit review. Never coerce them into the bundled fixture
profile or freeze formulas merely to reproduce baseline values. Verify live
calculations and source-bound parity before calling a migration complete.

### Phase 0d — Template-family fingerprint & routing (financial-statement models)

Before choosing a build path, **fingerprint the file** (see `refs/research-template-coa.md`):
a **broker/FactSet equity-research model** (line-items down rows, YEARS across columns, sections
like PROFIT AND LOSS / PER SHARE DATA / BALANCE SHEET; often `__FDSCACHE__` + thousands of named
ranges) routes to the **research archetype**, NOT the FP&A forecast path:

```bash
python scripts/infer-canonical-formulas.py <file.xlsx> --sheet="FY results" \
       --first-row=7 --last-row=<end> --out plan.json --draft
python scripts/build-research-model.py plan.json --conn <writeConn> --folder <id>   # dry run
python scripts/build-research-model.py plan.json --conn <writeConn> --folder <id> \
       --review-dispositions review.json --post
```

The standalone inference CLI runs the source gate before parsing formulas; `--draft`
reports pending review issues instead of failing (it is strict by default). Use the
same source-bound review file for production inference and build. `infer-canonical-formulas.py` then extracts
the hard-coded grid + one canonical formula per derived line, records the absolute source
path and SHA256, breaks plug-cycles, and identifies
formula cells that need a cached-value fallback. Cached fallbacks can preserve displayed
values, but they are not a completed live-formula migration: any nonzero `parity.frozen`
adds `CACHED_FORMULA_FALLBACK` and requires its own source-bound disposition.
`build-research-model.py` verifies the saved source identity and gate before it writes local
specs or authenticates, then posts the DM (data page) + workbook (computation + transpose + pivot).
Use `--source <file.xlsx>` when a plan moved and its saved path no longer resolves; the
file must still match the plan's `sourceSha256`.

**For a fleet of similar files off one template** (the ~850-file case), use
`scripts/batch-convert.py` — it fingerprints, maps each file's line items to a canonical
chart-of-accounts (`refs/research-template-coa.md` + `coa.json`), converts, parity-gates, and
emits a **triage manifest** (AUTO_PARITY / NEEDS_REVIEW / BLOCKED / FAILED), including
the full audit status/issues and source SHA256, with no silent truncation. `--post-review`
does not bypass the gate; reviewed files need `--review-dispositions` (a JSON file for one
source or a directory containing `<sourceSha256>.json` for a batch).
Read `refs/research-template-coa.md` first. Everything else (a cell-drawn P&L with a data tab,
SUMIFS actuals, driver growth) stays on the FP&A path below.

## Phase 1 — Pick the grain & extract the seed (`scripts/xlsx-to-input-csv.py`)

For the fact Table, emit a **paste-ready CSV** with:
- **only the data-entry columns**, headers in `UPPER_SNAKE_CASE` matching the
  Sigma input-table column names exactly (so a clipboard paste aligns 1:1),
- **no system columns** — `Row ID` / `Created at|by` / `Last updated at|by`
  auto-populate in Sigma; including them breaks the paste.

Parse Excel date serials → ISO dates here.

## Phase 1b — Land source tabs in Snowflake (`scripts/land-in-snowflake.py`)

Optional, forecast path, and only with the user's yes (`migrate.py plan` asks "land source
Tables in Snowflake?"; it needs the `snow` CLI and a Snowflake connection). Instead of
inlining GL actuals as `VALUES`, land the workbook's Excel Tables in the warehouse so the
model reads real tables. Without it the plain build reads the actuals from the plan.

```bash
.venv/bin/python scripts/land-in-snowflake.py book.xlsx --tables tblActuals,tblPlan \
    --database <DB> --schema <SCHEMA> --table-prefix FPA_ --snow-connection <snow-conn> \
    --out work/landing.json                      # dry run: prints the plan and SQL
# with the user's OK (it creates tables):
... --execute [--grant-select-to <ROLE>] [--replace] \
    [--sync-sigma-connection <connectionId> --sigma-profile "<profile>"]
.venv/bin/python scripts/build-forecast-model.py work/plan.json \
    --actuals-table <DB>.<SCHEMA>.FPA_ACTUALS --landing work/landing.json --out-dir work/build
```

- Table names drop a leading `tbl` and take the prefix: `--tables tblActuals --table-prefix FPA_`
  lands `FPA_ACTUALS` (`landing.json` lists each `snowflakeTable`).

- Headers become stable UPPER_SNAKE identifiers (`Net Revenue ($)` -> `NET_REVENUE`); the
  builder's column map reads the same names from `landing.json`.
- Excel date cells land as `DATE`; integer serials land as `DATE` only in columns whose header contains date/month/period/day, otherwise `NUMBER`; blank cells land as `NULL`
  (blank-as-zero parity per `refs/source-blank-validation.md`).
- The dry run runs one read-only `SHOW TABLES LIKE` per target and prints `EXISTING TABLES:`,
  so you know before `--execute` whether `--replace` is needed (if `snow` cannot answer it says
  "not checked" and `--execute` checks again).
- Re-running into a schema that already holds a table exits 2 unless `--replace` is passed;
  it never appends a duplicate. `--execute` verifies each row count against the source and
  exits non-zero on a mismatch.
- `--grant-select-to <ROLE>` (optional) grants `SELECT` on each landed table and nothing else;
  pass it only when Sigma's connection role cannot read the schema already.
  `--sync-sigma-connection` makes Sigma see the new table paths.
- `--replace` recreates with `CREATE OR REPLACE TABLE ... COPY GRANTS`, so the replaced table's
  existing grants are kept. Proven live on ymb68310 (2026-11-09): a table landed with
  `--grant-select-to SIGMA_ROLE`, then re-landed with `--replace` and new rows and no grant
  flag, still lists `SELECT` to `SIGMA_ROLE`, and `SIGMA_ROLE` reads the new row count. Do not
  pass `--grant-select-to` again on a replace; the grant carries over.
- `--database`, `--schema`, `--table-prefix` and `--grant-select-to` must each be one plain
  Snowflake identifier (`^[A-Za-z_][A-Za-z0-9_$]*$`; the prefix may be empty); anything else
  (dots, quotes, `;`, a leading `-`) is refused before any SQL runs, and `--sigma-profile` is
  checked like every other Sigma CLI profile name.
- Needs no warehouse views and no stored procedures. Ask the user before `--execute`,
  any grant, or a Sigma connection sync.

## Phase 2 — Build the read/output DM (against a seed) + reuse-check

Before POSTing a new DM, check whether an existing Sigma data model already
covers the same grain (same formal-Table columns / warehouse view). Prefer
reusing or extending that DM over spawning another copy. When no match, land
the seed (inline `VALUES` custom-SQL for ≤~1k rows; PUT-to-stage + COPY INTO for
larger) and POST the DM (fact + dim elements + relationships + a flattened
`*_REPORT` join element for grouping).

**Post-DM readback (hard gate):** after `POST /v2/dataModels/spec`, GET the
live spec and map server-assigned element/column ids before any workbook build
(`build-forecast-model.py` / `build-research-model.py` already do this via
`map_dm*`). Never trust authored ids past the POST. Full build rules in
`refs/sigma-build-gotchas.md`.

## Phase 3 — Stand up the input table (`scripts/build-input-table-wb.py`)

Build the **empty input-table structure via REST**; seed rows through Sigma MCP
Code Mode when connected, or load them through the UI. See
`refs/input-tables.md`.

**Empty input table (API for structure).** POST an **empty input-table element** at
the fact grain (`source: { kind: empty, connectionId: <write-conn> }`,
`inputMode: explore`, data columns + system columns
`ID/CREATED_AT/CREATED_BY/UPDATED_AT/UPDATED_BY`). Then either pre-query the live
element and seed missing natural keys with MCP `InputTable.upsert_rows()`, or use
the UI to **paste / upload** the Phase-1 CSV and **Publish**. Finally use
**Warehouse views → Create new** in the UI.

> **Do NOT auto-build linked input tables via the REST spec.** It looks like it works
> (the POST succeeds, `/elements` shows the columns) but the **inherited columns
> resolve to "multiple values"** — the REST spec omits the linked correlation
> state (validated
> 2026-06-10, publish doesn't fix it). `build-input-table-wb.py --linked` only
> scaffolds the PK/grain + entry columns. If a usage pattern needs live context
> columns beside editable cells, use the Sigma MCP Code Mode path (validated live
> 2026-10-22): `scripts/linked-input-program.py` emits the 3-call create / seed /
> edit programs over a one-row-per-key source. See `refs/input-tables.md`.

## Phase 4 — Source-swap the DM onto the view (`scripts/build-dm-on-view.py`)

Re-point (or build) the DM's fact element to `SELECT … FROM <warehouse view>` on
the write connection. Display names stay identical, so the workbook + metrics are
unchanged. This is **API**. Verify the section rollups hit the same parity targets
as Phase 2.

## Phase 5 — Build / repoint the workbook, layout last, parity

Build the dashboard pages (Forecast Summary pivot + KPIs, trend charts) per
the `sigma-workbooks` companion skill when available, sourced from the DM. Builders author a nested
`pages[].elements` draft; **`scripts/lib/workbook_wire.py` assembles layout XML
as the last write**, then wraps via `code_rep` into
`{ name, folderId, contents: { schemaVersion, kind, pages, elements, layout } }`.
Live builders first validate with `POST /v2/workbooks` + `dryRun: true`, then
create through the same endpoint. Data-model specs stay unwrapped.

Confirm every element compiles and the totals tie out (parity hard gate —
Phase 0c/0d source-bound live evaluation, Phase 2/4 section rollups, research
`batch-convert.py` AUTO_PARITY bucket). Never skip the parity comparison.

### Security: RLS / CLS

Excel workbooks have no native row-/column-level security construct to port —
access control is file-share / workbook-protection only. Detect workbook
password protection / sheet protection during Phase 0 and note it for the
customer; do not invent Sigma RLS from spreadsheet structure. If the landed
warehouse tables already carry RLS policies outside Excel, leave those to the
warehouse / Sigma admin — this skill does not apply RLS automatically.

---

## Scriptability matrix (OpenAPI re-audited 2026-10)

| Step | Path |
|---|---|
| workbook create/validate/update | ✅ `POST /v2/workbooks` (`dryRun` then create), `PUT .../contents` |
| input-table **structure** (columns/types; title, style, sort) | ✅ code representation |
| input validation (literal/range/dynamic-list/file constraints) | ✅ code representation |
| **linked** input table (inherited columns resolve) | ❌ REST spec POST yields "multiple values"; ✅ MCP Code Mode `add_linked_input_table()` validated live 2026-10-22 (`scripts/linked-input-program.py`); UI works |
| seed **records inside REST workbook code representation** | ❌ records are excluded from the spec |
| seed/update rows in an existing input table | ✅ Sigma MCP Code Mode `build_workbook` with `InputTable.upsert_rows()` / `update_rows()`; pre-query natural keys because inserts do not deduplicate them automatically |
| **Publish** for UI edits | ⚠️ UI; MCP row writes persisted and were immediately queryable in the 2026-10-15 regression, while published-mode UI permissions remain untested |
| **warehouse view** on the input table | ⚠️ UI only |
| land workbook Tables in Snowflake (Phase 1b) | ✅ `scripts/land-in-snowflake.py` (`snow` CLI; dry run by default; `--execute`, row-count verified) + `--sync-sigma-connection` |
| enable **manual commit** on input tables (needed for backend actions / agent tools to write) | ❌ REST spec POST never sets it; ✅ post-create MCP Code Mode `set_edit_mode('draft')` then `set_edit_mode('published-all-users')` per input table (`build-forecast-model.py --emit-manual-commit-program`, which also seeds the forecast's rate / manual / Parameters tables) |
| macros -> workflow buttons + approvals (`--workflows`) | ✅ `macro-actions.py demo-manifest` -> `workflows.json` -> `build-forecast-model.py --workflows` |
| workbook **agent** (Phase 6 Enhance) | ✅ `scripts/recommend-agent.py` (pitch) + `build-forecast-model.py --agent-macros` on the post (`--merge-into` only for specs the forecast builder did not build); `requiresApproval` is enforced in the chat UI, not over the REST/MCP run API |
| DM **source-swap** to `FROM <view>` | ✅ API (DM spec PUT/POST) |
| read/output DM + workbook build | ✅ API |
| API/writeback/navigation/stored-procedure actions | ✅ code representation; runtime-test every effect |
| Python computation | ✅ `code` + `code-output`; package/runtime + parity gate |
| scheduled/webhook automated action | ✅ limited document surface; ownership/failure review |
| column protection / broader admin policy | ⚠️ UI/admin policy |

For a REST-created empty input table, query it with the MCP companion tools, run
`scripts/input-sync.py architecture --workflow ...` for the one source workflow,
then `stage` with the output plan, exact source, named
approval, and target bindings. Emit construction/native joins only from the
approved manifest/receipt. After querying assigned IDs, re-stage those bindings
and run `plan` to reconcile typed natural keys. Review its Code Mode chunks,
execute through MCP, and query/read back again. The CLI never queries or mutates
MCP itself. `complete` requires all workflow-specific target, checkpoint,
edit/restore, persistence, and table evidence. Each table readback must report
the exact approved row count and typed `naturalKeys`; missing, extra, or
same-count substituted keys block completion. Otherwise hand off
CSV/clipboard import and **Publish**. In either path,
**Create warehouse view** remains an explicit UI hand-off. Linked tables use the MCP
recipe in `refs/input-tables.md` (`scripts/linked-input-program.py`), never the REST spec.

`stage` audits the exact source bytes and always refuses `BLOCKED` or blocker
severity. Pending review issues may be built in the isolated test workspace
with explicit staging approval; this does not assert verified dispositions.
Approval includes the existing source/plan hashes, named `owner` and `reviewer`,
all formula `reviewItems`, exact architecture `mappedFeatures` and
`targetBindings`, plus `sourceReviewApproval`:

```json
{"status": "approved_for_staging", "issues": ["replace with the exact full audited issue objects"]}
```

Obtain the issue objects from `migration_gate.gate_plan_staging_source(plan,
source_arg=source, plan_path=plan_path)[1]["issues"]` and review their complete
scope. Omitted or changed issues fail approval. Finalize connection, folder,
input `editMode`, and all logical architecture before hashing the manifest;
rebuild and reapprove if the plan changes. The receipt reports
`APPROVED_STAGING`, `PENDING_STAGING`, `pendingSourceIssues`, and
`sourceReviewRequirements` (each full issue plus its `issueHash`). Cached or
unresolved formula mappings still cannot stage. Production `--post` source
gates retain their strict verified-disposition contract.

Record Task 2 component readback/edit/restore checks as evidence JSON while
source-feature work is pending; defer `complete` until the later workflow
acceptance gate. Global completion re-audits the source and requires exact
`manifest.sourceReviewEvidence` coverage, independent of the receipt. Each
record needs `code`, `issueHash`, `status: "passed"`, a real `evidenceId`,
`sourceHash`, `planHash`, `targetDocumentId`, `targetVersion`, `targetId`, and
`targetBinding` with actual document/version/element and optional column/check
IDs. `targetId` must match that binding. These are recorded runtime checks,
never synthesized from staging approval; chart rendering, validation behavior,
and formatting parity remain pending until actually tested.

---

## Phase 6 — Enhance: recommend a Sigma Agent (`scripts/recommend-agent.py`)

Last step, after parity, and only with the user's yes to "build the workflow planning app +
agent?" (`migrate.py plan` asks it when a non-AUTO macro is an action write-back, external I/O
or what-if scenario macro; navigation, recalc or cosmetic macros alone get no offer). `migrate.py
plan` writes the macro inventory to `<workdir>/macros.json`; classify the macros into workflows with

```bash
.venv/bin/python scripts/macro-actions.py demo-manifest book.xlsm --out work/workflows.json
```

(the positional is the `.xlsm` or a `.bas` macro module; VBA is only read, never run; add
`--landed` when Phase 1b landed the source Tables, so the refresh disposition says so).
The build emits only the input tables its workflows target: Parameters for set-scenario /
roll-forward, Versions for snapshot, Approvals and the status pill for submit.

**Approve / Send back (`--approvers`).** With a Submit workflow, pass
`build-forecast-model.py --approvers a@x.com,b@y.org` to add approver-only decision buttons beside the
Approvals table on the Versions & approvals page (Approve filled, Send back outline, and a caption
"Only approvers can approve: ..."). Each is an `update-rows` on Approvals whose row filter is
`[Status] = "Submitted"` and the signed-in user's email (case-insensitive, whole address) being in the
build-time list, so for anyone else it matches no row; only Submitted rows match, so migrated history is
never rewritten, and two outstanding requests both get the decision. It writes Status (Approved /
Returned), Decided By and Decided At; the header status badge follows the latest entry. The decisions are
kept out of the workbook agent's tools. Without `--approvers` the buttons are omitted and
`workflows.json` gets an `approve-omitted-no-approvers` disposition saying so.
Read the macro inventory and `workflows.json`, say in plain
language which macros were workflows people triggered by hand, and offer a Sigma
workbook agent that runs them on request, with approval before every write:

```bash
.venv/bin/python scripts/recommend-agent.py --macros work/macros.json --workflows work/workflows.json
```

That prints the pitch only. With a yes, the agent reaches the workbook one way: add
`--agent-macros work/macros.json` to the post (see **One run, one POST** below). Do not run
`recommend-agent.py --merge-into` for a forecast build: its merged spec is never posted.
`--merge-into <workbook-spec> --out <merged>` / `--emit` are only for a workbook spec this
skill's forecast builder did not build. A macro inventory with no action or scenario classes
(only recalc / cosmetic) gets no agent and no write tools. See `refs/agent-enhance.md`.

The forecast run with both opt-ins, in order (`migrate.py plan ... --land-tables --enhance`
prints it as `nextCommands`, with the staging post as `stagingAlternative`; without them it is
in `optIns`): `macro-classify.py` (Phase 0b) → `land-in-snowflake.py` dry run,
then `--execute` (Phase 1b) → `draft-forecast-plan.py` → `macro-actions.py demo-manifest` →
`build-forecast-model.py` dry run (`--actuals-table --landing --workflows
--emit-manual-commit-program work/setup-program.py`) → `recommend-agent.py` (pitch) → the post
with the same flags plus `--agent-macros` → the setup program once → the smoke pass.
The post builds and checks the setup program from the local spec before it creates the data
model, so a program that cannot be built stops the run before anything is posted.

Prerequisites the agent depends on:
- Build with `build-forecast-model.py --workflows work/workflows.json
  --emit-manual-commit-program work/setup-program.py` so the Parameters / Versions /
  Approvals input tables and workflow buttons exist. Pass the same
  `--emit-manual-commit-program` path on the post too: the post rewrites the program against
  the created workbook and prints `NEXT: run <path> ...`; run that file.
- **Manual commit.** Backend actions and agent tools can only write to input tables
  that have manual commit enabled, and a REST spec POST never enables it. After the
  workbook is created, run the setup program from `--emit-manual-commit-program` once via
  Sigma MCP Code Mode `build_workbook` (`set_edit_mode('draft')`, then
  `set_edit_mode('published-all-users')` per input table). `published-all-users` is data
  entry for everyone with workbook access. After the mode switches the program seeds, with
  one `InputTable.upsert_rows` batch per table: the single Parameters row (Base scenario,
  first forecast month; emitted only when the build had `--workflows`), the Growth Rates
  and Manual Forecast input tables with the seed CSV rows (editable / what-if plans), and —
  with `--workflows`, when the source has them — the Versions / Approvals logs with the source
  workbook's own `tblVersions` / `tblApprovals` rows (migrated history: version, scenario, who,
  when, comment; the frozen FY figures the source never recorded stay blank). Run it
  once, on a fresh workbook: the writes do not query first and keys are not enforced. No
  seed CSV is pasted in the UI. The seeds come last: every call above them (formats, table
  style, column widths, series colours, edit modes) is safe to re-run. Before writing the
  program the builder binds every call against the recorded SDK signatures
  (`scripts/lib/sdk_signatures.py`, transcribed from the Sigma MCP server's
  `sdk://sdk/sigma/*.py`; checked by `scripts/lib/code_mode_check.py`), so a wrong method,
  arity or value (a grid-lines or edit-mode name) stops the build instead of the one-time run.
  Refresh that file from `read_resource` when the SDK changes.
- Agent SQL that touches input-table system columns (`CREATED_BY` ...) fails with
  "Ambiguous column reference"; ground tools in the user columns. See
  `refs/sigma-build-gotchas.md`.

**One run, one POST.** `build-forecast-model.py ... --workflows work/workflows.json
--agent-macros work/macros.json --post` first runs the agent merge against the local spec (a merge
that cannot apply stops the run with `AGENT ERROR` before anything is created), then posts the data
model, merges the agent into the workbook spec that carries the real data-model id, and dry-runs
and creates the workbook; `build/wb_spec.posted.json` and `build/posted-ids.json` record what was
posted. Only pass `--agent-macros` after the user said yes to the agent.

**Staging build while review is pending.** Production `--post` keeps the strict verified
review-disposition gate. To build in an isolated staging folder while issues such as
`VBA_PROJECT` / `FLAT_CARRY_POLICY` are honestly pending, the agent writes the approval template
(it lists the exact audited issues, status `pending`) and hands it to the named reviewer:

```bash
.venv/bin/python scripts/build-forecast-model.py work/plan.json ... \
    --staging-approval-template work/staging-approval.json
```

**User handoff — the reviewer approves, not the agent.** Policy: the staging approval must be
given by the named human reviewer at their own terminal; the agent must not approve, and must not
pipe, script or fake the answer. The reviewer checks every issue, then runs the approve command
**in their own terminal** (no hand edit of the JSON). It prints every audited issue and waits for
them to type `APPROVE`; without an interactive terminal (an agent's shell, a pipe, CI) it refuses
and writes nothing. The post refuses an approval file without the terminal confirmation that
command records, so a hand-written `approved_for_staging` JSON is not accepted. (These checks
support the policy; they are not an access control: anyone with a terminal can type `APPROVE`.)

```bash
.venv/bin/python scripts/build-forecast-model.py work/plan.json ... \
    --staging-approval-template work/staging-approval.json \
    --owner '<owner name>' --reviewer '<reviewer name>' --approve-for-staging
```

Once the file says `approved_for_staging`, the agent posts:

```bash
.venv/bin/python scripts/build-forecast-model.py work/plan.json ... --agent-macros work/macros.json \
    --post --staging-approval work/staging-approval.json
```

`--approve-for-staging` needs both names and the typed `APPROVE`; it sets
`sourceReviewApproval.status` to `approved_for_staging` and records the terminal confirmation
(`sourceReviewApproval.confirmation`), which `--staging-approval` requires.
Without it the names are recorded and the status stays `pending`. The approval is bound to the
source bytes, the exact issue objects and the plan's `folderId`; nothing is marked verified.

**Smoke pass (before the run is called clean).** The ACTION / CONTROL dispositions stay "live
side-effect check pending" until this runs once on the seeded workbook, off camera, so the
first execution of each path is never on camera. Log each step:
1. Switch scenario (agent `set-scenario`, or the switch): the FY figures change.
2. Roll forward (button or agent): Parameters `Forecast Start Month` +1 month; the closed month
   reads landed actuals on the chart, P&L detail and header; months left drop by one.
3. Snapshot: one new Versions row labelled `v<N>` (an integer) with the frozen FY figures.
4. Submit for approval: one new Approvals row, status `Submitted`; the header pill follows.
5. Agent `set-driver` on one line item: exactly one Growth Rates row changes.
6. Read each back with Sigma MCP `query`; then restore — start month, the edited rate, delete
   the test Versions / Approvals rows — and re-check parity and the header.
`requiresApproval` is enforced only in the chat UI, so check the approval prompt there by hand.
Elements on the hidden model page cannot be queried through MCP `query` (`Table workbook.<id>
not found`); compile-check them through the visible elements that read them and the page
screenshots.

### The planning app (`--workflows`)

With `--workflows` the forecast is laid out as a planning app (`scripts/lib/forecast_app.py`),
not the plain dashboard: **Forecast** (header band with title, Downside / Base / Upside
switch, approval-status pill and Roll forward / Snapshot / Submit for approval; four KPI
cards — FY revenue vs plan and vs prior FY, gross margin %, EBITDA, forecast months left; a
revenue chart with actual and forecast bars and a dashed plan line; the agent panel; forecast
vs plan by region and a P&L summary), **Assumptions** (editable growth rates and manual lines
with a live impact card), **P&L detail** (the FY's monthly P&L in $000s, actual | forecast, and
the GL actuals) and **Versions & approvals** (newest-first read-only views of the logs). Every
helper — Parameters / Versions / Approvals input tables, the status table, the union and
summary joins, the one `Forecast Boundary` — sits on the hidden `Model (sources)` page.
The logs keep a hidden `Log Order` (entry order): migrated rows carry the source's own dates,
which can lie after today, so "newest first" and the latest-approval pill follow entry order.
Without a Parameters table (Submit but no scenario switch or roll forward), the status table
sits on a one-row `VALUES` query instead. Proven live on tj-wells-1989 (2026-11-09): an empty
log reads the no-status text, then follows the row with the highest `Log Order` (Submitted,
then Approved).

- **Scenario switch.** When the driver table has scenario columns beside the live rate
  (`Downside | Base | Upside`), the drafter writes `plan.scenarios`; the data-model spine
  carries each non-default scenario's spread (scenario rate minus the default scenario's rate)
  and the forecast's `Eff Rate` adds `Switch([Scenario], ...)` — a control's `controlId` is a
  formula handle, so the switch drives the forecast directly. Edited growth rates stay the base
  the spread applies to. Snapshot records the switch's value; the agent's set-scenario tool
  sets the control.
- **Roll forward = a month just closed.** The actual / forecast split follows
  `Parameters.Forecast Start Month`, defined ONCE as the hidden one-row `Forecast Boundary`
  element (Start Month = the Parameters value, or the source's first forecast month until the
  setup program seeds the row; Start Key; Last Closed Month). Every other element reads it as
  `Max([Forecast Boundary/...])` — the Actuals wrap (landed rows before it), the forecast
  (spine rows from it on), the trend colouring (`Forecast Facts/Phase`), header subtitle,
  months-left card, P&L detail groups, and the Roll forward / Snapshot actions — so nothing
  repeats `Max([Parameters/...])`. Each growth / flat line re-seeds off the line's landed
  actual in the month before the boundary (`Lookup(Sum([Actuals/Amount]), [Line Item], ...,
  [Base Key], [Actuals/Month Key])`, N counted from there). Manual lines keep their typed values;
  year-over-year lines keep the source seeding. Before roll forward the FY figures tie to the
  source exactly (actual rows landed for a month the source still forecasts are ignored).
  **A closed month whose actuals have not landed keeps its forecast values** (`Landed` is NULL
  → the row stays, Period `Awaiting actuals`, a third P&L detail group; the header shows
  "Jul-27 awaiting actuals"; the chart keeps it in the forecast colour) — never blank or
  dropped; the months after it re-forecast from that frozen value, so FY figures do not move
  until the load arrives. The landed `Month Num` maps every period, so a late load keeps its
  number. `tests/test_forecast_app.py` evaluates the emitted model over local rows
  (`tests/planning_app_eval.py`): pre-roll parity, post-roll landed July, missing-actuals.
  A plain run (no `--workflows`) keeps the static split.
- **Plan figures.** When a `--workflows` run's `--landing` includes the plan table (`tblPlan`:
  Month, Line Item, Region, Plan Amount), it joins the data model as `Plan Detail` (a warehouse table, with the
  forecast's sections and month numbers attached by `Switch`); every month and line must be a
  forecast period / line or the build stops. The region chart and the plan comparisons appear
  only when it is there. A plain run (no `--workflows`) never reads the plan table. The P&L
  summary colours `Vs plan` by favourability: a cost line (every section it adds is a cost
  section, e.g. Operating expenses) under plan is green.
- **Header actions.** Roll forward / Snapshot / Submit for approval are one right-aligned
  toolbar: stretched buttons in adjacent cells of the header band, primary last and flush right
  (no nested container — the server's layout check misplaces a container inside a container);
  the status pill sits right after the scenario switch.
- **Input tables** hide the system `ID` column, keep the dense grid (horizontal lines; the
  presentation preset's taller rows clip a 10-row driver table) and turn autofit off; the setup
  program sizes each visible column to its longest seed value or header
  (`demo_workflows.input_column_widths`, `set_column_width`). A number column stretched across a
  wide table (autofit or not) shows its edit pencil far from its right-aligned name and reads
  as an empty column. On Assumptions the two tables sit side by side with the impact card.
- **What the POST cannot carry** — input-table typed-column `format` / `hidden`, and a combo
  chart's per-series colour — the setup program re-applies through Code Mode
  (`set_format_pattern`, `set_visible(False)`, `set_series_color`; and the input tables'
  `tableStyle` through `table_style()` and their column widths through `set_column_width`) before it
  switches the edit modes.

---

## What this skill does NOT do (yet)

- One-shot `.xlsx` → finished migration. Phase 0–1 (discovery + formula
  translation) is spiked, not a hardened converter. The data-entry path (Phase
  3–4) is fully validated.
- Power Query / ODBC / SharePoint-Online source pipelines are detected and
  routed, but not automatically recreated. Land them as governed warehouse or
  connector sources and verify refresh/data parity.
- Run the agent's writes unattended: `requiresApproval` is enforced in the chat UI, not over
  the REST/MCP run API, so approval-before-write holds only in the chat experience. It also
  never executes VBA or sends email.
- `.xlsm` / VBA / macros — **detected + classified + routed** (Phase 0b,
  `scripts/macro-classify.py` + `refs/macro-handling.md`); never blindly
  executed. Representable routes include Actions, Python, and stored procedures,
  but each still needs implementation and live verification.

See `refs/excel-translation.md` for the full converter-build design and the
~60–70% pipeline reuse from `tableau-to-sigma`.
