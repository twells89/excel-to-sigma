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
  --bindings runtime-bindings.json \
  --manifest migration-manifest.json \
  --staging-receipt approved-staging-receipt.json \
  --contents workbook-readback.json \
  --document-version 7 --openapi sigma-openapi-current.json \
  --plan-out presentation-plan.json --out workbook-put-body.json
```

Actual emission recomputes the staging receipt hash and immutable input artifact
records, audits the exact source bytes again, checks the complete pending source
issue scope, and requires the manifest target document/version plus the assigned
edit-table columns. Placeholder labels cannot pass this gate. Runtime bindings
provide actual page, element, and column IDs for charts, views, validation,
formats, conditional formats, and editable business inputs. Logical fixture
bindings emit `LOGICAL_FIXTURE_ONLY` artifacts and are rejected with
`--contents`.

The emitted artifact creates separate overview and detail table elements so a
single element is never moved between pages. The edit page references the
ordinary input table. A repeat merge replaces only source-owned presentation
pages and elements, preserves unrelated content, and rejects foreign ID
collisions. Existing business input tables use a reconciliation contract rather
than being recreated.

Forecast, sales, inventory, and expense inputs are constructed as ordinary
input tables in published Explore edit mode. Only source-authored literals are seeded;
derived columns live on native downstream calculation tables owned by the same
immutable architecture. Sales calculates weighted value and cycle days on
`salesInputBase`, then consumes the reviewed Task 2 AsOfDate join. Inventory
owns its availability, reorder-point, days-cover, `ROUNDUP`, and status chain on
`inventory`. Expense consumes the reviewed Task 2 policy join and owns the
numeric-Amount count check on `expenseNumericAmountCount`. Companion settings
and policy inputs are constructed separately, never duplicated by the wiring.
The artifact exposes this through `businessOperationalWiring`, produced by
`build_shared_input_architecture` and `build_operational_input_elements` rather
than a second formula contract. Research reuses its Task 2 shared input table and
adds the twelve source-authored `FY results!B56:M56` equity inputs in place.
The architecture contains `Scope` plus all 145 source inputs. Extension compares
an actual column readback, creates only missing columns, and never writes rows.
After assigned-ID readback, regenerate the 146-column receipt and use the normal
row-ID reconciliation path to seed only missing values; this preserves Scope and
prior user edits after a partial or complete retry.

Use the source-bound input entry point before presentation emission:

```bash
python3 scripts/presentation-inputs.py prepare source.xlsx \
  --connection CONNECTION --folder FOLDER --out-plan input-plan.json \
  --out-manifest input-manifest.json
python3 scripts/input-sync.py stage --source source.xlsx --plan input-plan.json \
  --manifest input-manifest.json --approval reviewed-approval.json \
  --out approved-preconstruction-receipt.json
python3 scripts/presentation-inputs.py emit --manifest input-manifest.json \
  --receipt approved-preconstruction-receipt.json \
  --feature forecastBusinessInput --out construction.json
```

After construction, read the assigned input/column IDs and restage those exact
bindings. The assigned native graph can then be emitted or merged into a complete
GET-back response without manual formula repair:

```bash
python3 scripts/presentation-inputs.py wiring --manifest input-manifest.json \
  --receipt assigned-receipt.json --contents current-get-response.json \
  --out native-wiring-put.json
python3 scripts/presentation-inputs.py bindings source.xlsx \
  --manifest input-manifest.json --receipt assigned-receipt.json \
  --out presentation-bindings.json
```

`--contents` requires `workbookId`, `documentVersion`, and the complete flat
`contents` document. It checks assigned input names, types, and column IDs,
preserves unrelated nodes, rejects foreign target collisions, and returns a
version-guarded contents PUT artifact. Omitting `--contents` emits the assigned
fragments and pending query checks. Apply only `contents` and `documentVersion`
to the API; other artifact fields carry provenance. Refresh/restage the receipt
after the PUT changes the version before emitting presentation bindings.

The forecast graph joins the assigned wide input to `spineWrap` by Line Item.
`Method`, `Monthly Growth`, and each source month column feed `forecast` and
`pnlCombined`; there are no separate rate/manual inputs. Actual amounts are
source literals, and the last-actual base is aggregated by SQL at query time.
`forecastInputsView` joins the same wide row to source line metadata for the
editing/detail presentation. The library forecast builder also accepts
`approved_inputs={"manifest": manifest, "receipt": receipt}` for composing a
workbook draft around this graph in the existing document; its default create
CLI retains the legacy contract.

Unsupported-root review mappings must name an actual emitted target and match
its exact native formula. Inventory uses `logicalElementId=inventory` and
`logicalColumnId=reorder_qty`; expense uses
`logicalCheckId=expense-numeric-amount-count`, which owns the native numeric
Amount count KPI column. These are pending review/runtime targets, not approval
or parity evidence.

Before the presentation PUT, run `presentation-inputs.py bind-target` with the
latest GET/readback document ID and version. Re-run it after a research column
extension changes the version, then approve the assigned post-extension
bindings against that manifest.

Research `prepare` also takes `--research-plan`. Its first `emit` consumes the
approved 134-column binding and an exact `--current-columns` readback, returning
only the twelve missing-column operations. Record their returned IDs with
`checkpoint-extension`, GET/read the exact fields, then stage a new receipt
containing all 146 assigned IDs. `seed-extension` accepts that receipt and proof;
it writes only created fields that were null at the checkpoint and remain null.
A non-null field is retained as a user edit. The checkpoint cannot be replayed,
and a partial retry emits only columns still absent from the latest approved
binding. Unsupported source formulas require explicit reviewed replacements via
`prepare --reviewed-formulas`; they are never silently promoted.
Persist the entire returned `{plan, programs}` file before attempting its
writes and supply that file (or its canonical nested plan) as the next
`--prior-seed`. The public and library boundaries normalize this exact envelope,
reject ambiguous/foreign shapes, and carry checkpoint consumption through later
normal reconciliation. New writes remain pending until an actual matching
readback confirms them; no null baseline is fabricated. Always retain the latest
persisted ledger: deliberately rolling back local state cannot prove that a
creation proof was previously consumed. The writer checks the approved seed-row
hash, reconciliation digest, action source values, and extension target scope.

Research runtime bindings should source the main and detail views from `calc`
(`Calc`): `Year` uses `yr`, and each metric uses the authored `c<row>`
column ID. The main view is Year, Revenue (`c8`), EBITDA (`c18`), EPS (`c37`),
and Net debt / EBITDA (`c57`). The detail view carries every populated FY-results
metric as a separate column. Each column takes its format from the first
formatted source period, preserving currency precision, percentages, and the
ratio `x` suffix. The emitter derives format exactness by comparing the complete
main/detail column contract; it does not trust a caller boolean. A generated runtime template records all
145 logical input IDs and source cells; replace its twelve post-extension column
placeholders with GET/readback IDs before integration.
The edit-table patch keeps `Scope` and `i_model_inputs_c2` through
`i_model_inputs_c7` visible in that order and formats the six growth assumptions
as percentages. The remaining source inputs stay in the same physical table and
remain available in Explore while the main edit grid stays compact.
Horizontal financial conditional rules retain their metric row and source year
scope: B31:M31 targets detail `research_c31`, and B57:M57 targets
`research_c57` on the main/detail views. Unbound Audit Status rules remain
explicitly unresolved; no calculation/status field is invented.

Every chart helper formula records target check ownership. Other source formulas
are collected in one hash-bound unresolved scope per workflow until target
column ownership or an executable parity check is approved.

Local extraction and schema audits do not prove runtime behavior. Completion
still requires SDK compilation, a guarded API dry run, GET/readback inspection,
aggregate parity queries for every chart point, and UI checks for rendering,
editing, ordering, rule precedence, and blank rejection. `allowBlank=false`
remains explicit unresolved evidence because the current input-column schema
does not expose that setting.
