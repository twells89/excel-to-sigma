# Bounded Regional scaling assessment

**Do not execute the current 309-table Regional candidate.** The scalar fallback is useful compiler construction coverage; it is not the final ordinary-table architecture. Reducing this to a useful grouped build requires material compiler work, not a batch-size change.

Corrected offline counts (test-only fixed-width row IDs, no target receipts):

| Category | Tables | Columns | Cause |
|---|---:|---:|---|
| Criteria row helpers | 192 | 10,668 | A separate row-helper graph for every one of 52 SUMIFS; repeated 144-row original-error ordering and criteria/type/error computations |
| Scalar helpers | 45 | 1,721 | 1,676 deduplicated helper nodes, packed by dependency depth, plus one key/table; original Revenue cells are still compiled separately |
| Formal Revenue row template | 2 | 12 | One proven shared row expression and five typed result lanes already exist, but scalar source-cell outputs are not rewired to this common row layer |
| Aggregate outputs | 52 | 260 | One row predicate/value/error/order table per SUMIFS, rather than grouped region/period/product records |
| Public scalar outputs | 18 | 1,088 | 214 × five lanes plus table keys, at 12 source cells/table |
| **Total** | **309** | **13,749** | Maximum formula length 2,406; dependency depth 28 |

157 tables are already 64 columns wide; 17 public tables have 61 columns. Raising an arbitrary 16-column limit will not solve this. Formula text totals about 2.9 million characters. The source AST contains 144 multiply formulas, 52 SUMIFS (16 region-quarter and 36 month-product), four lookups, four ratios, nine totals and one text expression. Nested DATE/EDATE calls create 32/16 function nodes. `scaling.json` provides full AST node and profile counts for all 13 workbooks. The source denominator remains 3,651 (3,515 ready, two reviewed replacements, 134 unresolved), with every source/native plan hash equal to the approved R1 replay.

The present implementation correctly proves formal Revenue's symbolic row template across all 144 source rows and provides its appendable layer. However, the main scalar pass still compiles the 144 Revenue source-cell ASTs individually; aggregations then also construct the row template. Aggregate CSE keys the complete bound AST, so different Summary region/period input references produce separate SUMIFS graphs. A new `row_graph` is allocated inside each aggregate, causing repeated criteria/error/ordinal helper tables. Exact-string CSE exists in each HelperGraph and the scalar pass; it does not merge these per-aggregate row graphs or recognize period/record-equivalent ASTs. The 144-branch original error-order chain is repeated 52 times (7,488 ordinal steps before other predicates). This is the largest source of the explosion.

Recommended next architecture, preserving the source cell/period DAG **before** grouping:

1. Intern the proven source-record Revenue expression once, with five lanes and original row provenance. Route original Revenue source cells to the corresponding live record field; appended records receive the same formula. Preserve any historical override as a separate source exception, never infer it into this template.
2. Intern common typed source columns and stable source error-order evidence once. Reuse them across all criteria; do not replicate 144 comparisons per SUMIFS. Future appended-row error ordering remains a reviewed ordinal obligation until proven.
3. Recognize the finite 16 region-quarter and 36 month-product vertices from their bound ASTs. Keep each exact source selector cell editable, each DATE/EDATE boundary and period dependency explicit, then group native calculations by these reviewed record keys. Use ordinary grouped native aggregates/joins over the shared Revenue rows rather than 52 wide per-cell predicate graphs. Filter each source period exactly (inclusive start, exclusive next-period start); do not collapse exact-month criteria into an entire calendar month.
4. Join the four source region targets while retaining exact-match, missing-key and first-duplicate behavior/error lanes. A natural key declaration is not proof that later user edits cannot create duplicates; source ordering must be preserved or duplicate edits explicitly rejected by a qualified business rule.
5. Route the 214 source-cell mappings to grouped record outputs and a small provenance result table with value/type/error fields; avoid 18 singleton presentation shards. Retain the nine total vertices, source text output, and 214 independent qualification obligations. Business editing tables remain the five named ordinary inputs; charts and presentation come later.

A plausible design is roughly five to eight calculated/grouped tables plus the five editable inputs, subject to real SDK grouping/error-lane qualification. This is a design estimate, not achieved coverage or an approved runtime count. It requires a new intermediate representation for row/period groups and a verified group-aware emitter. Merely deleting helpers or simplifying editable expressions by current values would be invalid.

A discovered correctness defect was fixed during this assessment: the aggregate emitter overwrote the Excel function name with a criteria table's name and emitted CountIf for SUMIFS after materializing helpers. A complete emitted-table-graph test reproduces the historical result 2 versus expected sum 7, then passes the fix. The graph harness also checks multiple rows/criteria, blank numeric summands, zero matches, selected error lanes, excluded error rows, COUNTIFS and IFERROR. All prior 3,515 candidate counts remain construction-only; they were never runtime parity evidence.
