# Source-bound native calculation candidates

`scripts/native-calculations.py` is an offline compiler and guarded artifact
preparer. It reads the source with `WorkbookContext` and the reviewed ordinary
input plan, retaining the exact source SHA256, every formula cell, bound AST,
source dependencies, range domains, and source-derived local expected values.
It never reads formula caches, writes a workbook, calls Sigma, or runs macros.

```sh
python3 scripts/native-calculations.py plan --source source.xlsx \
  --input-plan reviewed-input-plan.json --out native-plan.json
```

A source literal maps to an editable field only through the plan's exact
`sourceCells` reverse mapping. Reviewed external replacement formulas map to
their approved ordinary input field while preserving the original external
formula and binding identity. Ambiguous reshaped source headers, unmapped
business literals, Excel errors, unsupported functions/ranges and dependent
formulas remain named `unresolved` obligations. The planner reports these
without a cached or constant fallback. Formula status `ready-to-bind` means
local source-derived lowering is available; it does not claim Excel or Sigma
runtime parity.

The compiler's candidate form set is numeric/text/Boolean literals, exact
cell references, grouped unary `+`/`-`, percent, binary `+`, `-`, `*`, `/`,
`^`, comparisons and concatenation; scalar `IF`, `ROUND`, `ABS`, `INT`,
`MIN`, `MAX`, `SUM`; one-cell bound structured references; and rectangular
bound ranges expanded only inside `SUM`. Every referenced business literal
must have one unambiguous reviewed input provenance binding. Candidate forms
still require saved-source and fresh-value qualification.

Exact-lowering gaps include `SUMIFS`, `COUNTIFS`, `IFERROR`, `VLOOKUP`,
`XLOOKUP`, `INDEX`, `MOD`, `AND`, `ROUNDUP`, `CEILING`, `FORECAST.LINEAR`,
`FORECAST.ETS`, `NORM.S.INV`, `STDEV.S`, general criteria/range functions,
and text/Boolean coercion in arithmetic. Finite `INDIRECT` records its
alternatives and DAG edges, but remains unresolved until invalid-selector
and error behavior is exact. Whole-column or incomplete dynamic domains,
multi-cell arrays, source error cells, and ambiguous blanks stay explicit
obligations. Unsupported forms have no cached or constant fallback.

`lower-candidate` accepts actual physical row IDs and emits fully parenthesized
REST formula text plus a one-row scalar table and at most 64 independent formula
columns per shard. Formula dependencies move to a subsequent shard. Existing
ordinary input tables continue to expose their business fields. Scalar
`Lookup` references bind through the assigned input name, observed physical
row ID and observed row-ID column alias. The generated text uses documented
`Coalesce(..., 0)` numeric blank handling and explicit parentheses to avoid
the SDK export grouping defect. No raw formula parser method is assumed.

For a guarded full-contents update, save a complete fresh document GET and
full queried input rows, then run `backup` to persist the exact restoration
material. `prepare-update` requires that backup, the source input plan, a
manifest with matching input-plan hash and reviewed `live` formula mappings,
the `APPROVED_STAGING` assigned input receipt, the fresh document/version and
full queried rows. It checks actual input element, editable column, row-ID
alias and row identities. It emits one contents PUT body per shard while
copying all foreign contents, pages, events and schema. A prior shard must
have an exact saved readback receipt before the next can be prepared. Run
`confirm-readback` on the fresh saved GET; any changed expression structure
fails closed.

The native plan deliberately does not upgrade manifest review status or
complete any production gate. The current legacy manifest parser cannot mark
all new bound AST forms live; those mappings need separate reviewed manifest
integration. The REST artifact is a candidate until source, native and visual
runtime parity after input edit, reopen and restore are evidenced. Statistical
forecasting, dynamic locators, array spills, general criteria/range functions
and unreviewed source literal inputs remain unresolved.
