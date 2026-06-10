# Input tables — the data-entry half (VALIDATED 2026-06-08)

This is the part that was blocked until input tables shipped. It's now proven
end-to-end with exact parity. Read this before building the data-entry path.

## The mental model that matters

An Excel planning model is a **data-entry tool**. The faithful Sigma translation
keeps forecasters typing numbers — that's what input tables are for. But input
tables are **not** a data-model source kind. They are **workbook/document
elements** that write back to the warehouse, and the data model reads that
written-back data **through a warehouse view**.

```
Excel formal Table (forecasters type ForecastAmount)
      │  migrate the grain + values
      ▼
Sigma input table  (workbook element: kind: input-table, source.kind: empty)
      │  Sigma writes back on PUBLISH → SIGDS_-prefixed table (NOT directly queryable)
      ▼
Warehouse view on the input table   (UI: "Warehouse views → Create new")
      │
      ▼
Data model element  →  source: SELECT … FROM <database>.<schema>.<view>
      │
      ▼
Workbook charts / metrics  (unchanged — display names identical)
```

> **The original "flip the DM source to the input table, one line" plan was
> WRONG.** The swap targets a *warehouse view of* the input table. That adds one
> (UI) step the naive plan didn't account for. Everything else holds.

## The API-vs-UI split (the single most important fact)

Verified against the live Sigma OpenAPI and `/v2/connections` on 2026-06-08:

| Step | Path | Notes |
|---|---|---|
| input-table **structure** | ✅ API — `POST /v2/workbooks/spec` | columns, types; titles (`name`) + `tableStyle`/`sort` round-trip. Validation/protection are UI. |
| **linked**-table grain (rows from a spine element) | ✅ **API** | `source.kind: linked` + `from: <parentElementId>` — grain **inherited from the parent**, no CSV paste. Preferred when the grain is a dimension product (see "Linked input tables" below). |
| seed **data** load into an *empty/CSV* table | ⚠️ **UI only** | CSV upload or clipboard paste; no REST endpoint exists. Only `*/materialization*` endpoints exist, which are unrelated. Avoid entirely by using a linked table. |
| **Publish** | ⚠️ UI | data commits to the warehouse only on publish |
| **warehouse view** | ⚠️ **UI only** | input-table element → Warehouse views → Create new |
| DM **source-swap** | ✅ API — `POST`/`PUT /v2/dataModels/spec` | `SELECT … FROM <view>` |
| **"Editable in published version"** data-entry permission | ⚠️ **UI only** | The input-table element spec carries ONLY `id/kind/source/inputMode/columns` — there is NO permission field (verified 2026-06-09 against WANDA's viewer-editable input tables; none expose it). Default = editable in draft only. To let viewers edit in the published workbook: input table → **"Editable in" → Editable in published version – Explore / all users** → **Publish**. Still a Sigma Beta feature. `inputMode` ('explore'/'view') is a separate, unrelated element setting. |

So a converter auto-builds the input-table **structure** and the **DM-on-view**,
and hands the user three explicit clicks in between. Don't try to script the
upload or the view — there's no endpoint. (`swapSources` exists on dataModels but
swaps between *registered sources*, not to an ad-hoc input-table view.)

## Input-table element spec shape (proven)

```yaml
- id: forecastInput
  kind: input-table
  name: Forecast Entry          # element TITLE — persists in the spec & round-trips on GET
  source:
    kind: empty
    connectionId: cb2f5180-…    # MUST be a write-enabled connection (writeAccess: true)
  inputMode: explore            # 'explore' = editable; 'view' also seen on WANDA
  tableStyle: { preset: presentation }   # optional — tableStyle + sort round-trip too
  columns:
    - id: REGION                # data-entry columns: give a `type`
      type: text
    - id: MONTH_DATE
      type: datetime
    - id: CATEGORY_CODE
      type: number
    - id: FORECAST_AMOUNT
      type: number
    - id: ID                    # system columns: NO `type` — Sigma auto-manages them
    - id: CREATED_AT
    - id: CREATED_BY
    - id: UPDATED_AT
    - id: UPDATED_BY
```

- A write-enabled connection is required (`GET /v2/connections` → `writeAccess: true`).
- Column types: `text`, `number`, `datetime`, `checkbox` (and single/multi-select
  via data validation, configured in UI).
- The system columns (`ID` = Row ID; `CREATED_*`/`UPDATED_*` = row edit history)
  auto-populate. **Exclude them from the seed CSV.**
- **Titles & customizations now round-trip**: the element `name` is the visible
  title; `tableStyle` and `sort` persist in the spec too (treat like a `table`
  element). (Data validation / column protection / data-entry permission remain
  UI — see the matrix above.)

## Linked input tables — the powerful path (VERIFIED 2026-06-10, API-authorable)

For a planning model the grain is almost always **derivable from dimensions**
(Region × Branch × Month × Category …). Instead of seeding an *empty* table with a
pasted CSV, build a **dimension-spine element** and a **linked input table** off
it. The grain rows are **inherited from the parent automatically** — forecasters
fill only the measure column. This is fully **POST-authorable** (verified on a
feature-enabled org), which eliminates the manual CSV-paste step for the grain.

```yaml
# 1) the spine: any element whose rows define the grain — a warehouse/data-model
#    dimension, or a custom-SQL cross-join (CALENDAR × CATEGORY × BU), etc.
- id: spine
  kind: table
  source: { kind: data-model, dataModelId: <dm>, elementId: <el> }   # or warehouse-table / sql
  columns:
    - { id: storeKey, formula: '[D_STORE/Store Key]' }
    - { id: storeName, formula: '[D_STORE/Store Name]' }

# 2) the linked input table — rows come FROM the spine
- id: forecastEntry
  kind: input-table
  source:
    kind: linked
    from: spine                 # the PARENT element id (rows are inherited from it)
  inputMode: explore
  columns:
    - id: pk
      key: storeKey             # PRIMARY KEY → references the PARENT's column id (static row identifier)
    - id: storeNameLinked
      formula: '[D_STORE/Store Name]'   # LINKED column — inherits live parent value, NOT editable
    - id: FORECAST_AMOUNT
      type: number              # OWN entry column — the only thing forecasters type
    - id: UPDATED_AT            # system edit-history columns (no type)
    - id: UPDATED_BY
```

Why this is the strong default for Excel planning models:

- **Grain is API-seeded, not pasted.** The rows come from the spine — no CSV paste,
  no 2,000-row paste cap, no manual step. Build the whole thing in one POST.
- **Stale-spine problem solved.** The Ledcor model's calendar spine ended *before*
  the forecast window. With a linked table you point at a **freshly generated**
  spine (e.g. a custom-SQL `CALENDAR × CATEGORY × BU` cross-join element) and the
  grain is correct by construction — exactly the "derive, don't port" rule.
- **Dimension columns are locked for free.** Linked columns are non-editable by
  definition, so the grain keys can't be edited — you get column protection on the
  dimensions without the separate UI step. Only the entry column(s) are writable.
- **`from` / `key` survive POST.** The `from: spine` and `key: storeKey` references
  use your spec ids and Sigma maps them (like layout `elementId`s) — verified.

Use **empty + CSV** (below) instead when you're seeding **starting values**
(e.g. last year's actuals as a baseline to adjust), where the rows aren't simply a
dimension cross-product. Use **linked** when the grain is a dimension product and
forecasters enter net-new measures.

## The seed CSV (paste-ready)

Emit only the data-entry columns, headers in `UPPER_SNAKE_CASE` matching the
input-table column ids exactly:

```
REGION,BRANCH,SUB_BRANCH,MONTH_DATE,CATEGORY_CODE,FORECAST_AMOUNT
West,Seattle,SEA-North,2025-09-01,4000,94592
…
```

Paste limit is 2,000 rows × 25 columns; CSV upload limit is 200 MB / UTF-8. For
the 540-row fixture, clipboard paste was used.

## The warehouse view

After publish, the input-table element menu → **Warehouse views → Create new**
lands a queryable view (e.g. `SIGMA_WRITE_DB.SIGMA_WRITE.FORECAST_ENTRY`). The
view exposes **only the data columns** (system columns dropped); a Snowflake
`datetime` lands as `TIMESTAMP_LTZ`. The UI dialog truncates the path — get the
full one from the dialog tooltip or via `snow`:

```bash
snow sql -c default --format json -q \
  "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME \
   FROM SIGMA_WRITE_DB.INFORMATION_SCHEMA.VIEWS WHERE TABLE_NAME ILIKE 'FORECAST_ENTRY%'"
```

## Gotchas that cost time (so they don't again)

1. **Data is invisible until Publish.** A pre-publish MCP/REST query of the input
   table returns **0 rows** — the query layer reads the published version.
   Publish, then re-query. (This looks like "the upload failed"; it didn't.)
2. **Don't repoint a DM seeded with different data.** If you built the read-half
   DM against a different seed (different category codes / grain), build a *fresh*
   DM for the view rather than repointing — otherwise dimension joins silently
   miss and parity breaks.
3. **DM element `name` is not honored on POST** — every element reads back as
   "Custom SQL". Identify elements by inspecting columns, not name.
4. **DM element + column IDs are REASSIGNED on POST.** Always `describe` the live
   data model (Sigma MCP) to get the real element/column ids before querying.
5. **SIGDS_ tables aren't directly queryable** and must not be modified — always
   go through the warehouse view.

## Going further: governance for the forecaster experience (UI)

Once the structure exists, configure in the workbook UI:
- **Data validation** — turn `CATEGORY_CODE` into a single-select against the
  category list so forecasters pick, not type.
- **Column protection** — lock the dimension keys + system columns so only
  `FORECAST_AMOUNT` is editable.
- **Data entry permissions** — who can write.
