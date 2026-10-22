# Table-report recipes (formal Tables + formula grids)

For workbooks whose outputs are formulas over **formal Excel Tables** - an inventory
projection, a sales pipeline by owner, a headcount plan - and that are neither a
`forecast` P&L nor a plain data-entry table. `scripts/migrate.py plan` names this route
only when `draft-table-report.py` accepts the workbook, i.e. when it can reproduce **every
formula cell** from the table data. Pair with `refs/sigma-build-gotchas.md` and
`refs/input-tables.md`.

```
draft-table-report.py <file> --intent <i> --connection <c> --folder <f> --out plan.json
build-table-report.py plan.json --out-dir build/          # dry run: parity + specs + seed CSVs
build-table-report.py plan.json --post --review-dispositions review.json
```

## What is recognised

| Excel | Plan / Sigma |
|---|---|
| Formal Table, literal columns | `tables[]`; editable -> empty **input table** (+ `seed_<table>.csv`), read-only -> data-model VALUES element. Both are wrapped by a workbook `table` element named after the Excel table |
| Table calculated column (`tbl[[#This Row],[Col]]` or same-row A1 refs such as `=B2*C2`; arithmetic, IF/AND/OR/MIN/MAX/ABS/ROUND/CEILING/EOMONTH/MONTH/DATE) | a calculated column on the wrapper (copied-down formulas are compared in relative form) |
| `XLOOKUP(key, tblK[K], tblV[V], default)` / `INDEX(tblV[V], MATCH(key, tblK[K], 0))` | `Coalesce(Lookup([tblV/V], key, [tblK/K]), default)` (key column must be unique) |
| Scalar literal cell read by a formula (`Assumptions!$B$3`) | a row of the editable `Parameters` table: `Lookup([Parameters/Value], "Benefits load", [Parameters/Parameter])` |
| **summary** grid: label column + columns, each cell `SUMIFS/COUNTIFS/AVERAGEIFS(tbl[..], tbl[Key], $label, ...)` | grouped table (`groupings`): group by `Key`, other criteria become `SumIf`/`CountIf`/`AvgIf` conditions, each wrapped in `Coalesce(.., 0)` because Sigma returns NULL for an empty match where Excel returns 0 (AVERAGEIFS is left unwrapped: #DIV/0! in Excel); literal date windows `">="&DATE(..)` + `"<="&EOMONTH(..)` fold to date literals |
| arithmetic over the row's other cells, `SUM(B4:G4)`, `IFERROR(a/b,0)`, `IF(.., "ORDER", "")`, `CEILING(a,b)` | group-level calculations over the grouped table's columns |
| typed values in a summary column (quotas) | synthetic editable `<Grid> Inputs` table keyed by the row label, read back with `Lookup` |
| `SUM(INDEX(tblWide[[Jan]:[Dec]], MATCH(label, tblWide[Key], 0), 0))` | a synthetic calculated column on the wide table (`Total (Jan to Dec)`) + `Lookup` |
| **expansion** grid: ordinal `1..N` rows, `INDEX(tbl[Col], $A5)` attributes, date-headed period columns, optional `SUM` row total | the table **cross-joined** with a period spine (inner join on a constant `Join Key = 1`), the cell formula as a calculated column, a pivot (rows x period) |
| aggregates over an expansion grid's columns (`SUMIFS('Grid'!$P:$P, 'Grid'!$C:$C, $A4)`, `COUNTIFS(.., 'Grid'!$O:$O, ">0")`) | `Sum`/`CountIf` over the expansion's rows (a period-column criterion becomes `[Period] = <header date>`) |
| **keyed** grid: labels x date headers, cells `SUMIFS(tbl[V], tbl[RowKey], $A4, tbl[PeriodKey], B$3)` | the table itself is the long frame (one row per label x period) |
| running balance `prev -/+ SUMIFS terms`, opening term in the first column | `Opening + CumulativeSum([Net Change])` partitioned by the row key (the element is grouped by the key), ordered by the period |
| `MIN/MAX/SUM(B4:M4)` over the period columns | group-level `Min/Max/Sum` of the measure; one-key summaries over a keyed frame become that grouping's calculations |

## Verification oracle (never skipped)

The drafter lowers every formula to an IR (`lib/table_report_model.py`), evaluates the IR in
Python from the table data, and compares **every** recognised cell with Excel's cached value
(tolerance 0.01; text flags exact, `""` equals blank). It then requires every formula cell in
the workbook to be one of those verified cells. Any unrecognised formula, any mismatch and any
formula cell outside a recognised grid raises `PlanDraftError` naming `Sheet!Cell` - nothing is
coerced and cached values are never frozen into the plan as a conversion.

`build-table-report.py` re-runs the same check from the plan alone (`checks[]` hold the cached
values) and then prints headline parity per summary frame and numeric column; a mismatch exits 1
before any spec is written or posted.

## Shapes that fail loudly (not a fallback)

- `VLOOKUP`/`OFFSET`/`INDIRECT`/named ranges, `ROUNDUP`/`INT`/`MOD`/text functions, `&` text building,
  non-numeric (e.g. date) parameter cells, `DATE` with cell arguments,
  `EOMONTH(x, k)` with `k != 0`, approximate matches, wildcard criteria.
- A summary row label that has no rows in the data (a grouped table would drop the row).
- Aggregates not filtered by the row label; two grouping keys in one grid; mixed sources.
- A keyed grid that is not one-to-one with the table's (label, period) rows (sparse grids need a
  spine this builder does not have); an ordinal grid that does not cover all N table rows.
- Duplicate lookup keys; mixed literal/formula table columns; non-uniform copied-down formulas.

## Editable vs read-only

`plan.tables[].editable` defaults to `true` for every table a formula reads (so a typed value can
change a result) unless `--intent read-only`. Editable tables, the synthetic `Inputs` tables and
`Parameters` are created **empty** and seeded afterwards with the `seed_*.csv` files through Sigma MCP
Code Mode (`InputTable.upsert_rows`; query first - it enforces no key). Read-only tables are data-model
VALUES elements. Empty tables (e.g. a log written by a macro) are always input tables, with
guessed column types flagged as `TABLE_COLUMN_TYPES_GUESSED`.

## Sigma formulas used (dry-run validate before relying on them)

`Lookup`, `Coalesce`, `If`, `IsNull`, `And`/`Or`/`Not`, `Sum`/`Count`/`Avg`/`Min`/`Max` and
`SumIf`/`CountIf`/`AvgIf`/`MinIf`/`MaxIf`, `CumulativeSum`, `Ceiling`, `Least`/`Greatest`, `Abs`,
`Round`, `Date("YYYY-MM-DD")`, `DateAdd`, `DateTrunc`, `DatePart`. Least certain (no live proof in
this repo): `CumulativeSum` partitioning by a grouping inside one element, a pivot sourced from a
grouped table, an input-table element as the source of a wrapper `table`, a cross join through a
constant `Join Key` column, `AvgIf`/`MaxIf`, and `Least`/`Greatest`. The plan raises review issues for
the running balance (`RUNNING_BALANCE_PARTITION_REVIEW`) and the data-driven rows
(`GROUPED_ROWS_FROM_DATA`, `EXTRA_GROUP_KEYS`).

## Charts

Source charts are listed as not mapped (`CHARTS_PRESENT` stays a review item); rebuild them from the
verified summary frames with `refs/presentation-migration.md`.
