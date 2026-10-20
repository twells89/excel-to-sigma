# Regional01 B4 grouped native canary — offline operator packet

This packet covers **Summary!B4 only**, plus the shared 144-row
`tblSalesForecast[Revenue]` formula template. B4 is a Region/Q1 `SUMIFS`, with
the region read from editable `Summary!A4`; `Summary!B3` is the literal text
`Q1` and is not a formula input. B4's dates are formula constants
`DATE(2026,1,1)` and `EDATE(...,3)`. The source formula, AST, source hash, and
Excel result `1,544,984.02` are verified against the original workbook and the
independent Excel oracle. This packet is **offline candidate evidence** only.

Run `scripts/regional-grouped-canary.py` with the exact original workbook,
reviewed v2 input plan, native v2 plan, and independent Excel oracle JSONL.
It emits `candidate.json`, including all 144 Revenue source-cell mappings,
the distinct B4 five-lane mapping, complete B4 dependency list, the actual
source formula SHA, 3 logical elements, 27 formula columns, and unresolved
qualifications. The hard canary cap is 8 tables / 128 formula columns. The
source AST uses `table:tblSalesForecast:#data` with rows 2:145, so **appended
sales rows remain in scope**. The candidate cannot be interpreted as a closed
144-row SUMIFS range.

The logical construction is:

1. An ordinary editable sales input with 144 literal rows and five authored
   fields (`Region`, `Product Line`, `Month`, `Units`, `Avg Price`). Before
   staging the canary, add two metadata fields to each input row: immutable
   `Source Key = Forecast Input!row:N` and `Source Ordinal = N-1`. These are
   migration metadata, not extra Excel literals or fake Sigma IDs. Reconcile
   all 144 physical row IDs after seed. New appended rows need a reviewed
   monotonic ordinal before error-order parity can be accepted.
2. A calculated table sourced from that input. The existing typed native
   compiler emits one shared Revenue row template with five public lanes
   (number, text, logical, kind, error), plus live Region/Year/Quarter/group
   keys, source key, and source ordinal. All 144 original `F2:F145` ASTs are checked
   row-relative against that template; matching formula text alone is not the
   proof. Group key uses `Lower(Region)`, `DatePart("year", Month)`, and
   `DatePart("quarter", Month)`.
3. A single grouping by that composite key. It computes `Sum(Revenue Number)`
   and `Min(First Error Ordinal)`; `Grouping.to_source()` provides a flat source.
   The B4 summary table is sourced from the four source-bound region labels and
   looks up key `Lower(A4)|2026|1`, coalescing a missing numeric group to zero.
   Its error lane follows the lowest applicable ordinal to the originating
   row's Revenue Error. No source result is copied as a Sigma formula.

## Minimal staging sequence and exact contracts

1. Take a fresh full document GET, input-row query, original-layout backup,
   and assigned input/page ID receipt. Confirm the saved source/input/native
   hashes in `candidate.json`. The input plan currently lacks `Source Key` and
   `Source Ordinal`; update and review the staging input architecture, seed
   receipt, and manifest before building. Do not substitute a test UUID or a
   guessed row ID. Query the actual row-ID column alias after each REST build;
   the prior virtual alias may normalize to `ID`.
2. Call `prepare_scaffold_program(candidate, element_kind="calculated-table",
   page_id=<observed>, source_element_id=<observed sales input>, grid_row=<free>)`.
   It emits one documented `get_page` / `add_table(source.element(...))` /
   `add_column(..., const_number(0))` SDK program. Execute only in a disposable
   staging document. Read back its actual element ID and every column ID, and
   check the full document/layout for foreign changes.
3. Build an assigned receipt with `sourceSha256`, `documentId`,
   `documentVersion`, `elements[regional_sales_rows] = {elementId, name,
   columns: {logicalColumnId: actualColumnId}}`, and
   `namespaces.Sales = {name: actual input element name, columns: {logical
   field name: actual field name}}`. Run
   `prepare_contents_patch(candidate, fresh_document, receipt,
   element_kind="calculated-table")`. Its `contents` is a full-document PUT
   body; use `ifMatchDocumentVersion`, then GET and
   `confirm_contents_patch(packet, readback)`. The preparer rejects a changed
   placeholder, unknown reference, stale version, or formula over 6,000 chars.
   The confirmer checks exact saved full contents. A structural readback is
   **not runtime parity**.
4. Obtain the row table handle in an MCP session, call
   `set_groupings([[group_key_column]])`, then
   `grouping.set_calculation("Revenue Sum", sigma.formula.sum(revenue_number))`.
   Qualify the SDK constructor for `Min(First Error Ordinal)` against the actual
   SDK before adding it; `Min` is documented as a Sigma formula, but the cached
   SDK reference only establishes `set_calculation` and the `sum` example.
   Call `grouping.to_source()` and add one flat grouped table. GET the saved
   grouping, actual grouped element/column IDs, and query grouped rows.
   Verify year/quarter date semantics, `Sum`, `Min`, and source edit propagation.
5. Use `prepare_scaffold_program(..., element_kind="summary-table",
   source_element_id=<observed summary-regions input>, ...)`, read back IDs,
   then patch from a fresh document with namespaces for `Summary Regions`,
   `Grouped`, and `Rows`. The `Grouped` namespace must resolve `Group Key`,
   `Revenue Sum`, and `First Error Ordinal`; `Rows` resolves `Revenue Error`
   and `Source Ordinal`. Patch every formula with an exact version guard and
   confirm saved full contents. Bind B4 to the first summary-regions source
   row through its observed row ID; other summary rows in the table are not
   accepted as migrated formula cells here.
6. Query B4 number/text/logical/kind/error baseline and compare with Excel.
   Edit one ordinary sales `Units`, `Avg Price`, `Month`, and `Region`; query
   after each edit and reopen. Check zero, blank, no-match A4, duplicate group
   membership, and a test-only error path if a typed error can be injected
   without altering source meaning. Restore every input and confirm B4 again.
   Keep before/after document, row, grouping, and query receipts. The offline
   graph interpreter has already checked baseline, edits, zero, null,
   no-match, and synthetic first-error ordering, but cannot stand in for this
   live sequence.

The exact source labels and calendar constants are preserved. B4's full
Excel `SUMIFS` semantics still require proof for case-insensitive and wildcard
criteria, blank/error criteria, duplicate rows, date/time zone boundaries,
and newly appended error rows. If any fails, leave this formula blocked and
retain its source mapping. The other 51 Regional `SUMIFS`, four source
`VLOOKUP` `#NAME?` errors, charts, and full-workbook promotion are outside
this canary.

Cached SDK references: `work/native-grouped-docs-20261020/grouping.md`,
`levelTable.md`, `source.md`, and `tables-skill.md`. Sigma formula references:
[DatePart](https://help.sigmacomputing.com/docs/datepart),
[Concat](https://help.sigmacomputing.com/docs/concat),
[Min](https://help.sigmacomputing.com/docs/min).
