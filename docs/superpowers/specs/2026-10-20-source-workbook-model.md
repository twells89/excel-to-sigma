# Source workbook model

This is the first bounded implementation in the already authorized migration-gap work. The existing builders recognize too few workbook shapes. A common, source-bound inventory must precede source-specific construction so a mixed table, renamed sheet, or scratch area cannot silently lose formulas.

## Scope

Add an offline, JSON-serializable model of workbook facts, with a small inventory CLI. This does not build Sigma elements, approve migrations, execute VBA, or make unsupported formulas buildable. Keep the current migration manifest and its production gates unchanged.

Public interfaces in `skills/excel-to-sigma/scripts/lib/source_workbook.py`:

```python
build_source_model(source_path, *, block_decisions=None) -> dict
validate_source_model(model, source_path=None) -> None
```

CLI: `scripts/source-workbook.py SOURCE --out MODEL.json [--block-decisions DECISIONS.json]`. Validation and source inspection happen before writing the requested output. Failure returns nonzero and does not overwrite an existing output.

## Model contract

- `modelVersion` is `1`; `sourceSha256` hashes the unchanged source bytes; `sourceName` is its basename. Do not include absolute local paths, wall-clock timestamps, credentials, or Sigma runtime IDs.
- Model collections are `sheets`, `cells`, `tables`, `blocks`, `definedNames`, and `unresolved`. Include a `counts` object with `formulas`, `charts`, and `tables`. Stable IDs must encode their component boundaries without collisions, even when sheet names contain punctuation or apostrophes.
- Each populated cell retains sheet, address, OOXML/openpyxl data type, and number format. Distinguish `literal` from `formula` using the cell's actual formula type, not a leading equals sign. Preserve typed booleans, numbers, text, errors, dates, and datetimes in JSON without conflating null, empty string, zero, or false. Dates and datetimes have explicit type tags and ISO values. Blanks referenced by supported formulas remain identifiable even when absent from the populated-cell list.
- Formula cells retain their exact source expression. Use the shared `excel_expression` parsed expression model for reference analysis. Record resolved cell dependencies, including every cell inside bounded ranges, and explicit unresolved analysis for unsupported syntax, unresolved names, and missing sheets. An unsupported formula must never appear to have a verified empty dependency set. Defined names retain source expression and local scope; distinguish successful single-cell resolution from unresolved/range/constant definitions. Resolve local versus workbook scope correctly or leave the reference explicitly unresolved. No cache becomes a literal input or completed live calculation.
- All formulas are inventoried, including formula cells in tables, input-looking rows, hidden sheets, and scratch blocks. Preserve cycles as source facts; the inventory must not cut a dependency or evaluate formulas.
- Sheets retain visibility, protection state, merged ranges, and counts of source charts, validation rules, and conditional-format rules/ranges. Tables retain source sheet, name, range, headers in order, and each column's literal/formula/blank cell addresses. Blank or mixed columns remain ambiguous facts; a formula percentage does not determine editability. Preserve header/body/totals boundaries and do not treat totals formulas as entered records.
- Table names and source ranges are inputs to the model. No fixture names, known sheet names, four-section vocabulary, or fixed period labels select its behavior.
- Formal tables become unreviewed candidate blocks. Optional decisions use `{ "sourceSha256": "...", "blocks": [{ "sheet": "...", "range": "A1:C10", "role": "input|calculation|reference|scratch" }] }`. Require the matching source hash, existing sheet, bounded valid range, supported role, and no overlapping reviewed block ranges. The decision changes only role metadata. It must not exclude any cell or formula. A referenced scratch sheet remains in the model and dependency graph.
- `validate_source_model` rejects duplicate identities, malformed model/version/counts, formulas mislabeled as literals, broken dependency identities, invalid decision bindings, and changed source bytes when a source path is supplied. Validation is a structural/source-identity check, not a claim that calculations are supported or migration is complete.
- Building twice from the same source and decisions produces identical JSON data. Source files must not be saved or modified; embedded VBA must not run.

## Acceptance evidence

Normal CI uses generated temporary workbooks, including renamed sheets, quoted names, mixed tables, protected/hidden scratch cells, table totals, scoped names, formula-looking text, dates, blank references, unsupported structured references, and cyclic formulas. The external eight-workbook replay uses the byte-preserved originals and a source-hash-pinned expected inventory: formula counts `[214, 222, 132, 228, 109, 166, 625, 90]`, chart counts `[2, 1, 1, 2, 1, 1, 1, 1]`; totals are 1,786 formulas and 10 charts. It must retain all twelve `13WCF!C7:N7` prior-week cash links. The pipeline's main table must retain six literal columns and three formula columns independently. An external replay does not make the full migration pass.

The five original fixtures, existing tests, migration-manifest version, `BLOCKED` behavior, and strict production-write gates remain unchanged.

## Subsequent work, outside this component

1. Consume reviewed source blocks in generic ordinary-input-table/native-join builders, reusing `input_sync.plan_seed` and preserving user edits.
2. Extend the shared expression compiler for structured references, lookups, SUMIFS, date conditions, arbitrary sections and month/quarter/week dependencies. Qualify a formula-writing route: the current SDK export lost operand grouping during Research readback. Preserve that as a failed qualification, not proof of incorrect native evaluation.
3. Implement reviewed wide-to-long extraction and employee-by-period calculation expansion; retain reverse cell mappings and all 625 headcount formulas, not only its 480 cost cells.
4. Bind charts, validation, formatting, units, ordering, and overview/edit/detail layouts to the common source map, including all ten charts in these eight files.
5. Implement reviewed statistical and macro ports with independent expected results, selected-record guards, operation IDs, retries, failure handling, and explicit external procedure contracts. `truth.json` labels the statistical ETS values Python approximations; authoritative Excel results or an approved replacement are needed for exact parity. Unknown ERP contracts and email delivery remain outside automatic execution.
6. Run per-workbook staging edit/reopen/restore, published-action, permission-role, and screenshot acceptance. Repair original URLs only after staging passes and row/document backups exist.
