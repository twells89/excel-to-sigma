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
| `IF(MOD(B$4,2)=0, amount, 0)` (cadence) | generic `IF`/`MOD` over the period | `If((Mod(Mod([Week], 2) + 2, 2) = 0), amount, 0)` (the divisor-signed form, as Excel's MOD) |
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
(float noise only: `abs <= max(1e-6, 1e-12*|value|)`, `--tolerance` sets an absolute limit; text flags exact, `""` equals blank). It then requires every formula cell in
the workbook to be one of those verified cells. Any unrecognised formula, any mismatch and any
formula cell outside a recognised grid raises `PlanDraftError` naming `Sheet!Cell` - nothing is
coerced and cached values are never frozen into the plan as a conversion.

Blanks follow Excel, not SQL: an operand that can be blank (a blank cell in the data, or a lookup into one) is emitted as `Coalesce(x, 0)` / `Coalesce(x, "")` where Excel reads it as 0 / "", aggregate criteria use Excel's criteria rules (a blank matches only `<>`), and a formula cell that would be a blank is `0`. `scripts/lib/sigma_formula_eval.py` evaluates the EMITTED formulas with SQL NULL logic and compares every verified cell with Excel's cached value.

`build-table-report.py` re-runs the same check from the plan alone (`checks[]` hold the cached
values) and then prints headline parity per summary frame and numeric column; a mismatch exits 1
before any spec is written or posted.

## Shapes that fail loudly (not a fallback)

- `VLOOKUP`/`OFFSET`/`INDIRECT`/named ranges, `ROUNDUP`/`INT`/text functions, `&` text building,
  non-numeric (e.g. date) parameter cells, `MIN`/`MAX`/`AND`/`OR` over a blank cell, `IFERROR` around any error other than a
  zero/blank denominator or a missed exact lookup, a row label or `MATCH` key containing `* ? ~`, text that differs only in
  case (grouped labels, compared literals, lookup keys), an expansion grid whose shown columns do not identify one table row
  (a unique column is added when one exists, flagged `EXPANSION_ROW_KEY_ADDED`), `DATE` with cell arguments,
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

## Presentation (the Report page)

Presentation never changes a value; every figure below is re-derived from the verified cells before a spec is
written (`build-table-report.py` prints `KPI PASS` / `CHART PASS` lines and exits 1 on any difference).
Styling is ported from the sigma-skills `sigma-workbooks` skill (commit 95e67b7: `styling.md` Recipes 1-6,
`styling.rb`) in `lib/presentation_style.py`; the recipes never override the source - source formats, row
order and header fills win. Plan field `presentation` (`lib/table_report_presentation.py`) holds the title,
accent, `kpis[]` and `charts[]`.

| Source | Sigma |
|---|---|
| cell `number_format` (`lib/excel_formats.py`) | column `format`: `{"kind": "number", "formatString": <d3>, prefix?, suffix?}` / `{"kind": "datetime", "formatString": <strftime>}`. `#,##0` -> `,.0f`, `$#,##0.00` -> `$,.2f`, `0.0%` -> `.1%`, `€#,##0` -> `,.0f` + `prefix`, `yyyy-mm-dd` -> `%Y-%m-%d`, `mmm-yy` -> `%b-%y`. Never an Excel mask (`$#,##0` is a 400) |
| `(...)` negative section | the d3 `(` sign: `#,##0;(#,##0)` -> `(,.0f` (**unconfirmed live**; `translate(native_parens=False)` falls back to a minus sign) |
| `[Red]` negative section | `conditionalFormats: [{"type": "single", "columnIds": [...], "condition": "<", "value": 0, "style": {"color": "#EF4444"}}]` on the table / pivot |
| General / unknown code | an explicit plain number (`,.Nf`, N from the values) so Sigma never inherits `%` from an operand (`Amount * Probability`) |
| a code with no exact equivalent (fractions, `[h]`, scaling commas, conditions, accounting zero dash) | the closest plain format + review issue `NUMBER_FORMAT_APPROXIMATED` listing the codes and cells |
| summary / keyed row order | a hidden grouping calculation `Source Order` = `Min(If([Key] = "a", 1, If(...)))`, FIRST in `groupings[].calculations`, plus a grouping-level `sort: [{columnId, direction: ascending}]` (accepted live though the offline OpenAPI omits it). A hidden group KEY is silently un-hidden by Sigma, and element `sort` only reaches the base level, so neither is used. The keyed frame keeps a hidden base-level order column for its pivot (`rowsBy[0].sort = {by, aggregation: min}`) and adds `Source Order (group)` = `Min([Source Order])` as its group calculation. Chart `xAxis.sort` uses an order column of its own. Expansion pivots keep ascending key order |
| header fill / bold / font color | `tableStyle.textStyles.header` (pivots also `columnHeader`); under a look only the bold is kept (the theme styles headers) |
| (default look) | every table and pivot the Report page shows gets `tableStyle.preset: "presentation"` (`SURFACES["default_table_preset"]`; the OpenAPI `pivot-table` `tableStyle` takes `preset` too), merged with the Excel header `textStyles`; the hidden Model / Inputs pages are left alone. Those rows render ~2.4x taller, so tables are sized with 2.5x the row allowance and the cap scales with it. Under a look the reference's preset wins; none in the reference means none |
| sheet title cell | header band title (text element tinted with the accent) |
| bar / column (clustered, stacked, 100%), line, doughnut, pie whose categories are a grid's row labels (or verified axis cells) and whose series read one verified column each | `bar-chart` (`stacking: none/stacked/normalized`, `orientation: horizontal` for bars), `line-chart`, `donut-chart` / `pie-chart` sourced from the frame's SOURCE element (a grouped table's calculations and keys cannot be referenced by name, even with `groupingId`): x = the key column, each series = the frame's own formula at chart level (`Coalesce(SumIf([tblPipeline/Weighted Amount], <window>), 0)`; a lookup of the key becomes `Max(Lookup(..))`), a `list` filter on the key when the data has keys the grid does not list; title / series order / category order from the source. Colors: `settings.theme.overrides.categoricalScheme` (the only spec path for multi-series and donut colors) with the accent first; a single series takes `color: {by: single}` |
| anything else (subset of rows, scatter / area / combo, values outside a verified frame) | listed in `unmappedCharts` with the reason (`CHARTS_PRESENT` stays a review item) |
| KPI row | per additive column the source totals or that is named Total / Commit / Quota-like (never a period column, an average or a count of a filter) plus a ratio of two of them over the **totals**: the frame column is re-emitted as signed leaf terms over its source rows (`Coalesce(SumIf(..), 0)` per aggregate; a typed-input / lookup column is `Sum([Inputs/Quota])` of its table), restricted to the grid's labels when the data has other keys. One source -> a `kpi-chart`; terms from several tables (Coverage = pipeline / quota inputs) -> a card container + text with `{{ratio \| fmt}}` (`SURFACES["kpi_text_card"]`, off drops such KPIs). `IFERROR(a/b, 0)` -> `If(IsNull(den) or den = 0, 0, num / den)`; 2-4 equal peer cards: muted 14 px title (`name {text, color #64748B, fontSize 14, fontWeight normal}`), bold 28 px value, no accent. A text card's body is `<span style="color: #64748B; font-size: 14px">LABEL</span>` + a BLANK line + `<span style="font-size: 28px">**{{expr \| fmt}}**</span>` (span styles may only use color, background-color, font-size, font-family) |

Page order: header band, KPI row (`1/7 .. 19/25`, 7 rows), the hero chart full width (`1/25`, 16 rows; bar charts with 12 or fewer categories get `xAxis.format.labels.labelAngle: 0`), divider + quiet `DETAIL` label, then the tables sized to their row count (`table_height`). A supporting chart whose categories are a grid's row labels sits beside that grid's table (table `1/15`, chart `15/25`, same rows, the taller height); other supporting charts pair in the chart row. A narrow table takes `1/15`. The first summary grid leads; the expansion / keyed pivots follow. No input tables (read-only intent) means no Inputs page. Anti-patterns avoided (styling.md checklist): focal point, no automatic 50/50, no repeated KPI
band, one accent (not sprayed), no nested cards (KPIs carry their own card style), quiet labels instead of
competing headings, no wide-table squeeze or empty card space, no navy slideware band.

Gate: `check_emitted` (and `build-table-report.py`, exit 1) evaluates the EMITTED KPI and chart formulas over the plan's
base rows with `scripts/lib/sigma_formula_eval.py` and compares them with the oracle; the drafter drops a KPI / chart that
fails or cannot be emitted from source columns (lookup tables with duplicate keys, a label holding `}}` / `|` in a
text card, a ledger chart over some periods without the list filter) and records review issue
`PRESENTATION_ELEMENT_DROPPED`. Lookup KPIs always sum only the rows keyed by the frame's labels. Scaling commas
(`#,##0,"K"`) show the full number without the unit (`NUMBER_FORMAT_APPROXIMATED`); a grid longer than
`MAX_ORDER_LABELS` (40) keeps Sigma's key order (`SOURCE_ORDER_OMITTED`). Running-sum frames (ledger, keyed) are
`EMITTED UNVERIFIED`: the evaluator does not model them.

Verified live in a dry-run + render: the grouping `sort` / hidden `Min` calculation, source-level KPI and chart
formulas, the KPI name / value styling, the text-card body and `labelAngle`. Still unconfirmed: the d3 `(` sign,
`hidden` on a chart sort column, a `list` filter on a chart, `tableStyle.textStyles.header` vs `columnHeader` on
pivots. Emitted KPI / series formulas are evaluated offline over the plan's base rows
(`scripts/lib/sigma_formula_eval.py`) and must equal the oracle's frame cells.

## Matching the look of an org workbook (`--look`)

"Make it look like Golden Orders": `scripts/style-reference.py <workbookId | URL> --out look.json [--sigma-profile P]`
makes ONE read-only `GET /v2/workbooks/{id}/spec` and writes the look (`lib/reference_look.py`, `extract_look`; `--spec
saved.json` reads a saved GET response with no network). `draft-table-report.py --look look.json` records it as
`presentation.look` (with `source {workbookId, name}`); `build-table-report.py` applies it. Without `--look` the
default look is byte-identical to before.

| Copied from the reference | How it is applied |
|---|---|
| `settings.theme`, exactly (a named org theme stays a name) | `settings.theme` = that theme; no `categoricalScheme` or other override of ours on top, so charts / donuts take the theme's palette and no per-chart colour is written |
| the most common `style` per kind (container, kpi-chart, chart, table) | that kind's card style, else no `style` at all (the theme decides); the header band container takes the reference's container style |
| most common kpi `name` (`color`, `fontSize`, `fontWeight`) and `value` styling | kpi-charts and the text-card KPIs (spans with only color / font-size, bold via `**`); none -> theme default |
| first text's span `font-size` / colour (or heading level), second text's span (subline), `theme.overrides.colors.text` | header band title / subline, quiet `DETAIL` label colour |
| most common `tableStyle.preset`, and the theme itself | `tableStyle.preset` (`SURFACES["look_table_preset"]`); the reference's theme styles the table headers, so `textStyles.header` is just `{fontWeight: bold}` when the Excel header is bold (no Excel fill or font colour) |

Not copied: the arrangement (header band, KPI row, hero chart, DETAIL, chart beside its table), and anything Excel
decides: number formats, header bold, sheet row order, the title text (Excel's header FILL and font colour only
apply to the default look; under a look the theme styles headers). None of our own colours
(`#0F172A`, `#64748B`, `#F0F7FF`, `#E2E8F0`, the accent) is written when a look is applied; Excel's `[Red]` negatives stay
(`style {color, backgroundColor: none}`: Sigma paints its default green highlight when the background is omitted). A look
whose preset is `presentation`, whose theme sets `tableStyles.cellSpacing` medium / large, or that is only a named theme
sizes tables with 2.5x the row allowance (those rows render ~2.4x taller); a text-card KPI under a look with no KPI type
uses Sigma's own 18 px / 40 px bold type.

## Charts

Source charts that map (table above) are emitted as native Sigma charts and verified point by point;
everything else is listed as not mapped (`CHARTS_PRESENT` stays a review item); rebuild those from the
verified summary frames with `refs/presentation-migration.md`.
