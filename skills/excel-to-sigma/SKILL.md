---
name: excel-to-sigma
description: >-
  Convert an Excel (.xlsx) workbook — especially a planning / budget / forecast
  data-entry model — into a Sigma data model + workbook. Use when the user has
  an Excel file (formal Tables, pivots, charts, or a cell-formula budget model)
  and wants to recreate it in Sigma. Covers formal-Table discovery, Excel-formula
  translation, and the input-table → warehouse-view → data-model pattern that
  preserves data entry. Requires a Sigma SIGMA_API_TOKEN and (for the data-entry
  half) a write-enabled Sigma connection.
user-invocable: true
---

# Excel → Sigma Conversion

> **STATUS: input-table data-entry path VALIDATED end-to-end (2026-06-08).** The
> full chain — Excel formal Table → Sigma input table → warehouse view → data
> model reading `FROM` that view → section rollups — was proven with **exact
> parity** on a representative 540-row forecast model. The read/output half (DM +
> workbook from a landed seed) was validated earlier (2026-06-03). The discovery
> + formula-translation layer is designed and spiked, not yet a one-shot
> converter — see `refs/excel-translation.md`.

**Read ALL of the following before replying or taking any action:**
- `refs/input-tables.md` — the validated input-table workflow + the API-vs-UI split. **This is the crown jewel of the skill: the source-swap goes through a warehouse view, NOT a direct DM source binding.**
- `refs/excel-translation.md` — the translation surface, the formal-Table-only MVP rule, and what "formula translation" actually means here (grain selection, not business logic).
- `refs/sigma-build-gotchas.md` — the hard-won Sigma spec rules (SQL element formula prefix, DM POST envelope, element-name/ID reassignment, publish gate).
- `~/sigma-skills/sigma-workbooks/SKILL.md` + the Sigma OpenAPI — canonical workbook spec.

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

A Phase-0 heuristic: a formal Table that's *fed by* formulas/other sheets and only
*read* downstream = **derived**; a Table whose cells are typed values with no
inbound formula = **entered**. Route entered → input-table builder, derived →
read-only DM/metric builder. When unsure, ask the user how they use that sheet.

### Doing inputs *today*, before the bulk-seed API

Bulk-seeding an input table from the Excel rows is a separate in-progress API.
Until it lands, you can **still ship editable surfaces** — pick by usage:

- **Net-new entry at a known grain** (forecasts, plans): **linked input table off a
  dimension spine** — grain auto-populates, forecasters fill the measure. Fully
  API today, no seed. (`refs/input-tables.md` → "the powerful path".)
- **Augment / annotate existing rows** (flags, planned values, notes alongside
  live data): **linked input table off the converted read-only element** — users
  see all their migrated rows as live context columns *and* type into added entry
  columns. Fully API today, no seed. (Reference build:
  `~/excel-convert-tests/saas_editable.json` — 130 subs shown live + editable
  PLANNED_SEATS/RENEWAL_RISK/NOTES.)
- **Edit the original values in place** (change a budget number that's already
  there): needs an **empty/CSV input table holding the data** → that's the
  bulk-seed gap. Bridge = UI CSV paste now; seamless once the seed API ships.

Design so the editable surfaces exist from day one; the seed API just fills the
"edit-in-place" case later. **Don't build a read-only model and call it a
migration.**

---

## Prerequisites

### Sigma access
```bash
# neutral cred file written by the sigma-migration setup (see refs/input-tables.md)
bash -c 'source ~/.sigma-migration/env; TOKEN=$(curl -s -X POST "$SIGMA_BASE_URL/v2/auth/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "grant_type=client_credentials&client_id=$SIGMA_CLIENT_ID&client_secret=$SIGMA_CLIENT_SECRET" \
  | python3 -c "import sys,json;print(json.load(sys.stdin)[\"access_token\"])"); <cmd>'
```
The **data-entry half requires a write-enabled connection** (input tables write
back to a `SIGDS_` schema). Confirm via `GET /v2/connections` → `writeAccess: true`.
Verified write connection on tj-wells-1989: `cb2f5180-…` (Snowflake ymb68310).

### Python
`scripts/` use `openpyxl` (+ `python-dateutil` for the fixture generator). Install
into a venv: `python3 -m venv .venv && .venv/bin/pip install openpyxl python-dateutil`.

### Optional: `snow` CLI
Used only to look up the **full path** of a Sigma-created warehouse view
(`SIGMA_WRITE_DB.<schema>.<view>`) when the UI truncates it. SSO browser-auth;
`snow sql -c default --format json -q "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME FROM <DB>.INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME ILIKE '<view>%'"`.

---

## Phase 0 — Discover the workbook (`scripts/xlsx-discover.py`)

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

## Phase 1 — Pick the grain & extract the seed (`scripts/xlsx-to-input-csv.py`)

For the fact Table, emit a **paste-ready CSV** with:
- **only the data-entry columns**, headers in `UPPER_SNAKE_CASE` matching the
  Sigma input-table column names exactly (so a clipboard paste aligns 1:1),
- **no system columns** — `Row ID` / `Created at|by` / `Last updated at|by`
  auto-populate in Sigma; including them breaks the paste.

Parse Excel date serials → ISO dates here.

## Phase 2 — Build the read/output DM (against a seed)

Land the seed (inline `VALUES` custom-SQL for ≤~1k rows; PUT-to-stage + COPY INTO
for larger) and POST the DM (fact + dim elements + relationships + a flattened
`*_REPORT` join element for grouping). Verify rollups against the seed. Full
build rules in `refs/sigma-build-gotchas.md`.

## Phase 3 — Stand up the input table (`scripts/build-input-table-wb.py`)

Two authoring modes — **prefer linked** when the grain is a dimension product
(the usual case for planning models); use **empty + CSV** for seeding starting
values. See `refs/input-tables.md` for the full shapes.

**Mode A — Linked (preferred, fully API).** POST a **dimension-spine element**
(warehouse / data-model dimension, or a custom-SQL `CALENDAR × CATEGORY × BU`
cross-join) **plus** a linked input table off it:
`kind: input-table`, `source: { kind: linked, from: <spineElementId> }`, a primary
key column (`{ id, key: <spineColumnId> }`), any linked dimension columns
(`formula: '[Spine/Col]'`, auto-locked), and the **entry column(s)** forecasters
fill (`type: number`). The **grain rows are inherited from the spine** — no CSV
paste. Then UI: **Publish** → input-table element → **Warehouse views → Create
new** → note the `database.schema.view` path. (Two UI clicks, not three.)

**Mode B — Empty + CSV.** POST an **empty input-table element** at the fact grain
(`source: { kind: empty, connectionId: <write-conn> }`, `inputMode: explore`, data
columns + system columns `ID/CREATED_AT/CREATED_BY/UPDATED_AT/UPDATED_BY`). Then
UI: open the input table → **paste / upload** the Phase-1 CSV → **Publish** →
**Warehouse views → Create new**. Use when rows are starting values, not a clean
dimension cross-product.

## Phase 4 — Source-swap the DM onto the view (`scripts/build-dm-on-view.py`)

Re-point (or build) the DM's fact element to `SELECT … FROM <warehouse view>` on
the write connection. Display names stay identical, so the workbook + metrics are
unchanged. This is **API**. Verify the section rollups hit the same parity targets
as Phase 2.

## Phase 5 — Build / repoint the workbook & verify

Build the dashboard pages (Forecast Summary pivot + KPIs, trend charts) per
`~/sigma-skills/sigma-workbooks`, sourced from the DM. Confirm every element
compiles (`verify-workbook`) and the totals tie out.

---

## Scriptability matrix (verified vs. live OpenAPI + connections; input-table re-verified 2026-06-10)

| Step | Path |
|---|---|
| input-table **structure** (columns/types; title `name`, `tableStyle`, `sort` round-trip) | ✅ API (workbook spec) |
| **linked**-table grain — rows from a spine element (`source.kind: linked` + `from`) | ✅ API — **no CSV paste** |
| seed **data** load into an *empty/CSV* table (CSV upload / paste) | ⚠️ UI only — no REST endpoint (avoid via linked mode) |
| **Publish** (commits writeback) | ⚠️ UI |
| **warehouse view** on the input table | ⚠️ UI only |
| DM **source-swap** to `FROM <view>` | ✅ API (DM spec PUT/POST) |
| read/output DM + workbook build | ✅ API |
| data validation / column protection / data-entry permission | ⚠️ UI (linked dimension columns are auto-locked) |

With **linked mode** the only mandatory UI steps are **Publish** and **Create
warehouse view** — everything else, including the grain, is scripted. Bake those
two into the run as explicit "do this, then tell me the view path" hand-offs.

---

## What this skill does NOT do (yet)

- One-shot `.xlsx` → finished migration. Phase 0–1 (discovery + formula
  translation) is spiked, not a hardened converter. The data-entry path (Phase
  3–4) is fully validated.
- Power Query / ODBC / SharePoint-Online sources — MVP assumes inline data +
  Snowflake landing (the other source paths are additive; see
  `refs/excel-translation.md`).
- `.xlsm` / VBA / macros — flagged in Phase 0, never executed.

See `refs/excel-translation.md` for the full converter-build design and the
~60–70% pipeline reuse from `tableau-to-sigma`.
