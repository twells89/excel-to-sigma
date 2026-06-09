# Excel → Sigma

A Claude Code skill that migrates **Excel (`.xlsx`) planning / budget / forecast
models** to Sigma — preserving the part that matters most about a spreadsheet:
**data entry**.

> **Private staging repo.** This is the development home for the `excel-to-sigma`
> skill. It is laid out as a plugin (`.claude-plugin/plugin.json` +
> `skills/excel-to-sigma/`) so it drops straight into the
> [`sigma-migration-skills`](https://github.com/twells89/sigma-migration-skills)
> monorepo (`plugins/excel-to-sigma/`) when it's ready to ship.

## What it does

Excel planning models aren't dashboards — they're data-entry tools. The faithful
Sigma translation keeps forecasters typing numbers, which is what **input tables**
are for. The skill covers:

1. **Discovery** — inventory formal Excel Tables, pivots, charts, and a formula
   census (flagging the untranslatable).
2. **Read/output half** — formal Tables → Sigma data-model elements; the
   income-statement SUMIFS become grouped DM metrics + workbook charts.
3. **Data-entry half (validated)** — the formal Table forecasters edit becomes a
   Sigma **input table**; the data model re-points to read it through a
   **warehouse view**.

```
Excel formal Table → Sigma input table → (publish) → SIGDS_ writeback
   → warehouse view → data model (FROM view) → workbook charts/metrics
```

## Status

- ✅ **Input-table data-entry path: validated end-to-end** (2026-06-08) with exact
  parity on a 540-row forecast model.
- ✅ Read/output DM + workbook build: validated (2026-06-03).
- 🚧 Discovery + Excel-formula translation: spiked, not yet a one-shot converter.

See `skills/excel-to-sigma/SKILL.md` for the full workflow and `QUICKSTART.md` for
a runnable walkthrough on the bundled sample fixture.

## Layout

```
.claude-plugin/plugin.json          plugin manifest
skills/excel-to-sigma/
  SKILL.md                          phased workflow (read first)
  QUICKSTART.md                     runnable end-to-end on the sample fixture
  refs/
    input-tables.md                 the validated input-table workflow + API/UI split
    excel-translation.md            translation surface, formal-Table rule, converter design
    sigma-build-gotchas.md          DM/workbook spec rules
  scripts/
    make-sample-forecast.py         generate a synthetic .xlsx fixture (no customer data)
    xlsx-discover.py                Phase 0 — inventory tables/visuals/formulas
    xlsx-to-input-csv.py            Phase 1 — extract a Table to a paste-ready CSV
    build-input-table-wb.py         Phase 3 — POST the input-table workbook (API)
    build-dm-on-view.py             Phase 4 — POST the DM reading FROM the view (API)
```

## Requirements

- A Sigma `SIGMA_API_TOKEN` (creds in `~/.sigma-migration/env`).
- A **write-enabled** Sigma connection for the data-entry half (input tables write
  back to a `SIGDS_` schema).
- Python with `openpyxl` (+ `python-dateutil` for the fixture generator).

## License

MIT
