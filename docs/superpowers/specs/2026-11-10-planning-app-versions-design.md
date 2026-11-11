# Planning app versions: freeze the inputs, compare to today

Date: 2026-11-10. Skill: `excel-to-sigma`, forecast route, the planning app built with
`build-forecast-model.py --workflows` (`scripts/lib/forecast_app.py`, `scripts/lib/demo_workflows.py`).

## Why

A Snapshot in the planning app saves a Versions row with the scenario, the forecast start month,
three FY totals (revenue, gross margin, EBITDA) and who/when. That is less than the source
workbook keeps, and there is nothing to do with it afterwards.

The source macro (`SnapshotVersion_Click` in `docs/demo-kit/demo-forecast-macros.bas`) copies the
whole Forecast range `A4:Z30` (every line item, every month) into the Versions sheet as values. Its
comment: "freeze the forecast as values and log a row in tblVersions". No formula in the `.xlsm`
reads the Versions sheet, no macro restores a version, and Submit writes a typed label ("FY27 v7")
rather than pointing at a version. So in the source, a version is a frozen record of the full
forecast that people look at and compare against. It is never reopened or edited.

The migrated app should keep that meaning and make it useful: every saved version can be compared
to today's forecast by line item and month.

## Goals

- Snapshot freezes enough to reproduce that version's full forecast (every line, every month).
- A comparison of any Sigma-saved version against today's forecast: headline totals, change by
  line item, and the line × month grid.
- Works for any forecast the skill converts, not only Vantora: the columns come from the source's
  own line items.
- The button and the agent's snapshot tool keep writing the same row (the agent tool reuses the
  button's effects, as today).

## Non-goals

- Reopening, restoring or editing a version. Out of scope for now; a later "working versions"
  feature can follow the scenario-grain pattern in sigma-skills
  (`sigma-workbooks/reference/workflows/planning-apps.md`).
- Any change to the plain forecast build (no `--workflows`). It has no Snapshot.
- Recovering numbers for the migrated versions (v1–v7 in Vantora). The source file kept only their
  log rows.

## Anti-goals

- A version that silently differs from what was frozen. If the rebuilt totals do not match the
  stored totals, the build fails, not the viewer.
- Labels that imply more than the data holds: migrated versions must not look comparable.
- Warehouse views, stored procedures, or anything outside the workbook and its input tables.

## Decided

1. **Freeze the inputs, not the output grid.** A live test (2026-11-10, three throwaway workbooks,
   since trashed) showed a button `update-rows` writes one value to every matched row: a constant
   filled all rows of a version, but a formula reading each row's own columns (`[Base] * 2`,
   `[Live]`, or the table-qualified `[Grid live/Live]`) left them empty, and once failed the
   database write. Copying each cell's live value with a button is therefore not possible. One cell
   per effect (about 300 effects for Vantora) was rejected as too heavy.
2. **What one Snapshot row stores** (one `insert-rows` into Versions, as today):
   - Version label, Log Order, Saved By (`CurrentUserEmail()`), Saved At (`Now()`), Comment — as today.
   - Scenario (the scenario control's value) and Forecast Start Month — as today.
   - One column per growth-rate line item, holding that line's current rate from the Growth rates
     input table (`rateInput`). Vantora: 10 columns.
   - One column per manual-forecast cell (line item × month) from the Manual forecast input table
     (`manualInput`). Vantora: 6 columns.
   - FY revenue, gross margin, EBITDA — as today, now used as the rebuild check.
   Each driver value is a single value pulled from its input table by line item (and month for
   manual cells), the same kind of formula the FY totals already use in this insert.
3. **Rebuild.** A hidden element recalculates each Sigma-saved version's forecast by line item and
   month from its stored inputs, with the same math as the live `forecast` element (rate, scenario
   spread, manual values, start month boundary, actuals for closed months).
4. **Rebuild check.** For every Sigma-saved version, the rebuilt FY revenue, gross margin and EBITDA
   equal the stored ones to the cent. Checked in the local parity step and in the live smoke pass.
5. **Compare view**, on the Versions & approvals page (`pgVersions`), above the existing logs:
   - A "Compare to" control listing Sigma-saved versions, defaulting to the most recent.
   - Headline: FY revenue, gross margin, EBITDA — today, that version, and the change.
   - Where it moved: each line item's FY change, biggest absolute moves first.
   - The grid: line items × months in the source's own layout, with a switch for Today / Version /
     Change.
6. **Stated limits, on the page:**
   - Migrated versions show "Numbers not recorded in the source" and are not offered in Compare to.
   - "Rebuilt from the saved inputs with actuals as they are now": a restated closed month changes
     a rebuilt version (Excel's pasted values would not).
7. **Fidelity record.** The Snapshot workflow disposition says the source froze the output grid as
   values and the migration freezes the inputs and rebuilds, with the restated-actuals difference.

## Calls made after review (builder questions settled here)

These were settled by the author, not the user; check them while reading.

1. **The rebuilt grid covers every row the live P&L shows**, including section totals and derived
   rows (gross profit, EBITDA, `derivedRows`), in the same dependency order as the live forecast
   stages (so a `pct_of_total` line reads the version's own rebuilt lines). The source macro froze
   all of `A4:Z30`, totals included.
2. **One source of math.** The builder generates the live forecast and the version rebuild from the
   same function, parameterised only by where the scenario, start month and drivers come from (the
   control / Parameters / input tables for live; the version's columns for a rebuild). No parallel
   hand-written formula set. Only the `--workflows` path calls the rebuild, so the plain build stays
   byte-identical to `main`.
3. **A stronger check than FY totals.** Besides the FY-total check, the smoke pass snapshots once and
   confirms the rebuild of that version equals today's live grid cell by cell (line × month), before
   anything changes.
4. **Which versions are comparable:** a new boolean Versions column, `Inputs Frozen`, written `true` by
   Snapshot. Migrated rows (seeded from the source log) leave it empty, show "Numbers not recorded
   in the source", and are not offered in Compare to. Empty driver columns are never used as the
   signal. (A fresh build cannot hold a Sigma snapshot saved before this feature, so no second label
   is needed.)
5. **Stored values are the effective values** the live forecast used: for a rate,
   `Coalesce(input, default rate)`; for a manual cell, `Coalesce(input, source manual value)`. A
   version is self-contained, the rebuild applies no defaults, and a later change to defaults never
   alters an old version.
6. **Shape changes are out of scope.** Driver columns are fixed when the workbook is built. Adding or
   removing line items means re-converting, which builds a new workbook with a new Versions table.
   Roll forward does not change the manual cells' columns: they are keyed by the source's manual
   line × month, not by the forecast window.
7. **If the first probe fails** (one insert cannot pull each driver value by line item), stop and
   report back before building anything else. The FY totals in today's Snapshot already use the same
   kind of conditional pull inside `insert-rows`, so this is expected to pass.
8. **Agent:** no new tool. The snapshot tool keeps reusing the button's effects, so it writes the new
   columns automatically. The compare element is added to the agent's read-only data sources, so it
   can answer "what changed since v9?". Its instructions describe the comparison generically; they do
   not list the driver columns.

## Constraints

- Existing rules hold: no warehouse views, no views over writeback tables, no visible workarounds;
  agents never type APPROVE; the setup program runs unchanged as one program.
- New Versions columns must be declared in `demo_workflows.TABLES` (or built alongside it) so the
  setup program, the migrated-history seed and the agent's grounding all see them.
- Input tables keep the hidden `Log Order` ordering rule.
- The plain (non-workflow) forecast build stays byte-identical to `main`.

## Left to the builder

- Element ids, column ids and names for the new Versions columns and the rebuild/compare elements.
- How the rebuild is expressed (a cross join of versions × the forecast spine with per-line lookups
  into the version's columns, or another shape), as long as it uses the live forecast's math.
- Exact layout of the compare section within the page, following the planning app's existing style.
- How driver values are pulled in the insert (`SumIf`, `Lookup`, or similar) — whichever the first
  live probe proves.

## To prove first (live, before building on it)

1. One Snapshot click writes every driver value (10 rates + 6 manual cells for Vantora) into one
   Versions row, each pulled by line item from its input table.
2. The rebuild reproduces the stored FY totals to the cent for a version saved in the Base scenario
   and one saved in Upside, and still does after a Roll forward.
