# Native presentation migration

`scripts/build-presentation.py` is the reusable presentation entry point for the
five acceptance fixtures. It extracts native charts, data validation,
conditional formatting, semantic formats, and formula provenance directly from
OOXML. Chart caches are retained only as expected-value evidence; emitted Sigma
charts always use live formulas over runtime-bound elements.

Generate an inspectable logical artifact without touching Sigma:

```bash
python3 scripts/build-presentation.py source.xlsx \
  --fixture-bindings --openapi sigma-openapi-current.json \
  --plan-out presentation-plan.json --out presentation-artifact.json
```

Generate a version-guarded full-contents PUT body after input construction and
GET/readback discovery have supplied actual IDs:

```bash
python3 scripts/build-presentation.py source.xlsx \
  --bindings runtime-bindings.json --contents workbook-readback.json \
  --document-version 7 --openapi sigma-openapi-current.json \
  --plan-out presentation-plan.json --out workbook-put-body.json
```

Runtime bindings must be hash-bound to the approved source, migration plan, and
staging receipt. They provide actual page, element, and column IDs for charts,
views, validation, formats, conditional formats, and editable business inputs.
Logical fixture bindings are deliberately rejected when `--contents` is used.

The emitted artifact creates separate overview and detail table elements so a
single element is never moved between pages. The edit page references the
ordinary input table. A repeat merge replaces only source-owned presentation
pages and elements, preserves unrelated content, and rejects foreign ID
collisions. Existing business input tables use a reconciliation contract rather
than being recreated.

Forecast, sales, inventory, and expense inputs are constructed as ordinary
input tables in published Explore edit mode. Only source-authored literals are seeded;
derived columns stay live. Sales adds live weighted-value and cycle-day columns,
then consumes the reviewed Task 2 AsOfDate join for TimingStatus. Inventory
adds its availability, reorder-point, days-cover, `ROUNDUP`, and status chain.
Expense consumes the reviewed Task 2 policy join for PolicyExcess and Exception.
The artifact exposes this through `businessOperationalWiring`, produced by
`build_shared_input_architecture` and `build_operational_input_elements` rather
than a second formula contract. Research reuses its Task 2 shared input table and
adds the twelve source-authored `FY results!B56:M56` equity inputs in place,
using the stored row ID so Scope and prior edits are preserved.

Local extraction and schema audits do not prove runtime behavior. Completion
still requires SDK compilation, a guarded API dry run, GET/readback inspection,
aggregate parity queries for every chart point, and UI checks for rendering,
editing, ordering, rule precedence, and blank rejection. `allowBlank=false`
remains explicit unresolved evidence because the current input-column schema
does not expose that setting.
