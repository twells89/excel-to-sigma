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
| input-table **structure** | ✅ API — `POST /v2/workbooks/spec` | columns, types, validation, protection |
| seed **data** load | ⚠️ **UI only** | CSV upload or clipboard paste; no REST endpoint exists. Only `*/materialization*` endpoints exist, which are unrelated. |
| **Publish** | ⚠️ UI | data commits to the warehouse only on publish |
| **warehouse view** | ⚠️ **UI only** | input-table element → Warehouse views → Create new |
| DM **source-swap** | ✅ API — `POST`/`PUT /v2/dataModels/spec` | `SELECT … FROM <view>` |

So a converter auto-builds the input-table **structure** and the **DM-on-view**,
and hands the user three explicit clicks in between. Don't try to script the
upload or the view — there's no endpoint. (`swapSources` exists on dataModels but
swaps between *registered sources*, not to an ad-hoc input-table view.)

## Input-table element spec shape (proven)

```yaml
- id: forecastInput
  kind: input-table
  name: Forecast Entry          # NOTE: element name often reverts to "New Input Table" in the UI; harmless
  source:
    kind: empty
    connectionId: cb2f5180-…    # MUST be a write-enabled connection (writeAccess: true)
  inputMode: explore            # 'explore' = editable; 'view' also seen on WANDA
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
