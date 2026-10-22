# Table-report recipes (formal Tables + formula grids)

For workbooks whose outputs are formulas over **formal Excel Tables** - an inventory
projection, a sales pipeline by owner, a headcount plan, a 13-week cash flow - and that are neither a
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
| **period ledger**: line items down, period ordinals `1..N` across a header row (`Week \| 1 \| 2 ... \| 13-Wk Total`), see the section below | ONE table with a row per period (the spine) and a calculated column per Excel row, in the sheet's row order; a one-row grouped total table |

## Period ledgers (13-week cash flow)

`find_ledgers` (`lib/table_report_ledger.py`) runs before the grid finders: a header row of >= 3 consecutive
integers after a text label, an optional text total header after them, and the rows below it (two blank rows
end the ledger). Section headings (a label alone on a row) are display only. Each row is one of:

| Excel row | Plan column (IR) | Sigma |
|---|---|---|
| typed numbers (rent, debt service) | a column of the editable spine table `<Sheet> Periods` | `Coalesce([<Sheet> Periods/Rent], 0)` |
| `SUMIFS(tbl[V], tbl[Period], B$4 [, tbl[C], "lit"])` | `lk` into a hidden grouped **helper** frame (`tbl by Period`, `V` summed, `0` default) | `Coalesce(Lookup([tbl by Period/V], [Week], [tbl by Period/Period]), 0)` |
| `... * INDIRECT("Coll_"&Scenario)` | `lk` into the scenario table built from the named-range family `Coll_*` keyed by the selected-scenario text parameter | `Lookup([<Rate> by Scenario/<Rate>], Lookup([Parameters/Text Value], "Scenario", [Parameters/Parameter]), [<Rate> by Scenario/Scenario])` |
| seed, then `=prev*(1+rate)` | closed form `seed*(1+rate)^(period-first)` | `seed * Power((1 + rate), ([Week] - 1))` |
| date seed, then `=prev+7` | `ADDDAYS(seed, 7*(period-first))` from a **date** parameter | `DateAdd("day", (7 * ([Week] - 1)), Lookup([Parameters/Date Value], ...))` |
| `IF(MOD(B$4,2)=0, amount, 0)` (cadence) | generic `IF`/`MOD` over the period | `If((Mod([Week], 2) = 0), amount, 0)` |
| `=B10+B11`, `=SUM(B15:B18)`, `=B23-MinCash` | same-period column arithmetic; a named single cell is a parameter | `([Payroll] + [Rent] + ...)` |
| `Beginning = prev Ending` (first = `OpeningCash`) with `Ending = Beginning + Net` | `Ending = Opening + CumulativeSum(Net)`, `Beginning = Ending - Net` | ungrouped table sorted by the period: `Opening + CumulativeSum([Net Cash Flow])` |
| `MIN($B$23:B23)` (running window from the first period) | `cummin` | `CumulativeMin([Ending Cash])` |
| total column `SUM(B10:N10)` | verified per row; a one-row grouped table emits the totals | `Sum([13WCF/Total Receipts])` grouped by the constant `"13-Wk Total"` |
| `="Scenario: "&Scenario` above the header | verified text caption, not emitted (`CAPTION_FORMULA_VERIFIED_ONLY`) | - |

Text and date parameters live in the `Parameters` table's `Text Value` / `Date Value` columns (added only
when a plan has such a parameter). Workbook controls are filter controls, not formula parameters: change
the scenario by editing the `Scenario` row of the Parameters input table. The plan's `scenarioResults`
recomputes the headline targets (totals, the last ending balance, the lowest balance) under every scenario
of the family; Excel cached only the selected one, so the others are informational, not parity targets.
A table with no unique text column (a bill list that repeats vendors) needs no key to be aggregated; it is
flagged `KEYLESS_TABLE_ROWS` because seeding its input table is positional.

Recognised only inside a ledger: named single cells (`OpeningCash`), `INDIRECT("prefix"&<named cell>)`,
`MOD`, text/date parameters. Everywhere else they still fail loudly. Within a ledger a formula may read the
same period only (a fixed-column reference, another period, or a relative cell outside the ledger fails
with its cell); `prev` is supported only as the three recursions above.

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

- `VLOOKUP`/`OFFSET`/`INDIRECT`/named ranges, `ROUNDUP`/`INT`/text functions, `&` text building,
  non-numeric (e.g. date) parameter cells, `DATE` with cell arguments,
  `EOMONTH(x, k)` with `k != 0`, approximate matches, wildcard criteria. (A period ledger accepts the
  `INDIRECT("prefix"&<named cell>)` scenario idiom, named single cells, `MOD` and text/date parameters;
  see above.)
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
`Round`, `Date("YYYY-MM-DD")`, `DateAdd`, `DateTrunc`, `DatePart`; ledgers add `Mod`, `Power`,
`CumulativeMin` and `Concat` (captions are not emitted). Least certain (no live proof in
this repo): `CumulativeSum` partitioning by a grouping inside one element, a pivot sourced from a
grouped table, an input-table element as the source of a wrapper `table`, a cross join through a
constant `Join Key` column, `AvgIf`/`MaxIf`, and `Least`/`Greatest`. Ledger-only and unproven:
`Lookup` into a hidden grouped table's calculation (the SUMIFS helpers) and a `Lookup` whose key is
another `Lookup` (the scenario factor), `CumulativeSum`/`CumulativeMin` over an ungrouped table ordered by
its sort, `Power`/`Mod`/`CumulativeMin` themselves, a grouped table keyed by a constant text-literal
column (the one-row total), columns that reference columns defined later in the same element, and the
`Text Value` / `Date Value` typed columns of the Parameters input table. The plan raises review issues for
the running balance (`RUNNING_BALANCE_PARTITION_REVIEW`) and the data-driven rows
(`GROUPED_ROWS_FROM_DATA`, `EXTRA_GROUP_KEYS`).

## Charts

Source charts are listed as not mapped (`CHARTS_PRESENT` stays a review item); rebuild them from the
verified summary frames with `refs/presentation-migration.md`.
