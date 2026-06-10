# Excel → Sigma — Quickstart

End-to-end on the bundled sample fixture. ~10 min; three of the steps are clicks.

## 0. Setup
```bash
cd <this skill>/scripts
python3 -m venv .venv && .venv/bin/pip install openpyxl python-dateutil
# Sigma creds in ~/.sigma-migration/env  (SIGMA_BASE_URL / SIGMA_CLIENT_ID / SIGMA_CLIENT_SECRET)
```

## 1. Make (or bring) an Excel model
```bash
.venv/bin/python make-sample-forecast.py        # writes "Sample Forecast.xlsx" (540-row formal Table + dims + a SUMPRODUCT report)
```
Or point the next steps at a real `.xlsx`.

## 2. Discover
```bash
.venv/bin/python xlsx-discover.py "Sample Forecast.xlsx"
```
Lists formal Tables, their grain, pivots/charts, and a formula census. Confirm the
fact Table + grain with the user.

## 3. Extract the paste-ready seed
```bash
.venv/bin/python xlsx-to-input-csv.py "Sample Forecast.xlsx" --table tblForecast --out forecast_input_paste.csv
```
Entry columns only, UPPER_SNAKE_CASE headers, dates as ISO. (System columns are
excluded — Sigma auto-populates them.)

## 4. Build the input-table workbook  (API)

**Preferred — linked off a dimension spine (grain auto-populates, no CSV paste):**
```bash
# spine.sql = a SELECT that generates the grain (e.g. CALENDAR × CATEGORY × BU
# with a surrogate GRAIN_KEY); forecasters fill only FORECAST_AMOUNT
.venv/bin/python build-input-table-wb.py --name "Forecast Entry" --connection cb2f5180-… \
  --spine-sql spine.sql --spine-cols GRAIN_KEY,REGION,MONTH_DATE,CATEGORY_CODE \
  --key GRAIN_KEY --entry-cols FORECAST_AMOUNT:number
```

**Alternative — empty table seeded by CSV paste (for starting values):**
```bash
.venv/bin/python build-input-table-wb.py --name "Forecast Entry" --connection cb2f5180-… \
  --columns REGION:text,BRANCH:text,SUB_BRANCH:text,MONTH_DATE:datetime,CATEGORY_CODE:number,FORECAST_AMOUNT:number
```

## 5. Load data + view  (UI — see refs/input-tables.md)
1. *(empty mode only)* Open the workbook → the input table → **paste**
   `forecast_input_paste.csv`. *(linked mode skips this — the grain came from the spine.)*
2. **Publish** (data commits to the warehouse only on publish).
3. Input-table element → **Warehouse views → Create new** → copy the
   `database.schema.view` path.

## 6. Build the DM on the view  (API)
```bash
.venv/bin/python build-dm-on-view.py --view SIGMA_WRITE_DB.SIGMA_WRITE.FORECAST_ENTRY \
  --connection cb2f5180-… --categories-from "Sample Forecast.xlsx"
# prints the dataModelId
```

## 7. Verify parity
Query the DM's report element grouped by Statement Section (Sigma MCP) and compare
to the targets the discover step computed. The fixture's targets:
Revenue **10,902,533** · Contract Costs **−5,603,043** · Admin **−2,221,185** ·
Allocation **−415,557** · Net Contribution **2,662,748**.

> Validated run (2026-06-08): exact parity on all sections. workbook
> `cb6c53ef-…`, view `SIGMA_WRITE_DB.SIGMA_WRITE.FORECAST_ENTRY`, DM `04e462d8-…`.
