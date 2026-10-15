# Explicitly reviewed custom ports

`custom-ports.py` and `lib/custom_ports.py` author candidates offline. They do not
execute VBA, Python in Sigma, SQL, grants or network calls. A Python availability
probe is not business parity. Production remains **BLOCKED** until actual runtime
acceptance; `validate_completion` checks the additional custom-port case ledger.

The controlled artifacts in `tests/fixtures/task5-controlled-*` are a newly
reviewed source, a real XLSX companion, model data, independent golden values,
and the exact reviewed V3 scalar procedure declaration. The BAS is raw source,
not a compiled XLSM. Do not execute it. The illustrative `Sample Macros.bas` stays
unchanged: missing `PostToGrid`, refresh and external email remain unresolved.
The Hello World binary demonstrates extraction only.

## Public interfaces

- `build_custom_port_plan(base_manifest, base_receipt, module_path=..., fixture_path=...,
  expected_path=..., reviews=..., scope=...)` extends the existing manifest with
  three ordinary input companions, exact source/procedure/helper hashes and named
  custom reviews. It preserves the prior formula mappings and shared inputs.
- `validate_custom_port_plan`, `source_helper_inventory` re-open/hash the source.
  The custom helper inventory distinguishes reviewed VBA intrinsics/labels from
  missing helpers; it does not change automatic native analyzer classifications.
- `build_port_input_specs(forecast_input_template)` adds headcount, overhead and
  one `Goal Seek Invocation` row with `State=UNPREPARED`. All new tables use SDK
  string `editMode='published-explore'`. Existing companion modes stay unchanged.
- `emit_python_programs(model_binding, column_maps=...)` returns standalone source,
  output schema, column order and code hashes. `build_port_fragments` wraps it in
  code elements, output tables and hidden-toast run buttons.
- `prepare_port_update` checks the exact source/title/connection graph, builds an
  append-only new page, and returns a `POST /v2/workbooks` dry-run body.
  `finalize_port_update` requires hash-bound actual HTTP 200/dryRun evidence and
  fresh unchanged contents/version before returning the versioned PUT request.
  Unknown existing IDs/titles are refused. `audit_port_update` compares readback.
- `solve_forecast`, `guard_goal_seek`, `plan_goal_seek_write`,
  `verify_goal_seek_write` expose pure computation/verification. They never mutate
  their inputs or persist assumptions. `build_goal_seek_actions` emits separate
  compute/guard/apply sequences for one exact operation, result and current model.
- `controlled_procedure_contract`, `build_procedure_action`,
  `audit_procedure_result` enforce the exact typed scalar contract and stop-on-error
  semantics. A pending signature cannot produce an action.

The CLI offers `prepare`, `emit-python`, `prepare-update`, `finalize-update` and
`prepare-goal-actions` and `procedure`; `--help` lists mandatory paths. All commands require `--out`. Review
and execution are separate; emitting an artifact does not authorize its submission.

## Construction and review runbook

1. Start with the current source-bound forecast manifest/approved input receipt.
   Obtain fresh staging contents and actual source/input/column/physical-ID rows.
   The scope binds `status=APPROVED_STAGING`, named `owner`, `documentId`, `version`,
   `folderId`, `connectionId`. Only the designated staging workbook is permitted.
2. Review all three mappings by exact `procedureId` and `procedureSha256` from
   `analyze_macro_source`. Each review has `mappingKind=reviewed-custom-port`,
   `reviewStatus=approved`, named `reviewer`, `decisionId`,
   `acknowledgeUnsupportedImperative=true`, `sourceOrigin=new-controlled-fixture`.
   This acknowledges the bounded analyzer's remaining imperative uncertainty; it
   is not automatic native approval. Bind the BAS, XLSX and independent goldens.
3. Build/restage the extended manifest with existing `input-sync.py` artifacts.
   An old receipt's plan hash cannot authorize these new tables. Use the existing
   ordinary-input constructors and resumable natural-key seed path for
   `task5Headcount`, `task5Overhead`, `task5Invocation`. Confirm published explore
   edit mode and preserve all original forecast rows/columns/values/IDs.
4. Add a reviewed **read-only** `Input Row ID` projection to Forecast Inputs View
   using the Task4 identity projection pattern. Do not turn a cached value into
   a live model. Actuals, Forecast Inputs View, Forecast Line Definitions,
   Forecast and PnL Combined must resolve to the same connection as all code
   elements; preserve the current actuals/manual/flat/growth equation graph.
5. Prepare `sources`: exact unique title -> `{elementId, elementHash}` using
   `_digest(actual_element)`. Required titles are Actuals, Forecast Inputs View,
   Forecast Line Definitions, Forecast, PnL Combined, Goal Seek Invocation,
   Task 5 Headcount Inputs and Task 5 Allocation Assumptions. The model binding
   hashes these sources, their full recursive element-source graph, workbook,
   controlled source files and equation version. Recheck that graph before every
   new solve/apply. Any model metadata change requires a new reviewed binding.
6. Inspect actual Snowpark `.columns`. Supply `column_maps` as title ->
   `{logical_display_name: actual_dataframe_column_name}`. Unmapped exact names
   are allowed; missing/duplicate names fail closed. No inferred case folding.
   Call `prepare_port_update`; review its complete diff and hash. It retains all
   original objects/layout and adds a dedicated page. Submit its dry-run body
   through the supported POST path. Capture `{httpStatus:200,dryRun:true,warnings,
   evidenceId,proposalHash}`; independently review nonempty warnings by hash.
7. GET immediately again, call `finalize_port_update`, and durably consume its
   proposalHash **before** sending the PUT exactly once. After timeout inspect
   actual state rather than replaying. This offline library takes a consumed-ID
   ledger; it cannot enforce concurrency against a caller that omits its ledger.
   Capture returned version/IDs and GET readback via `audit_port_update`.
   Server ID rewrites or any unexpected normalization stop for explicit rebinding.
   Full-contents REST may reset ordinary input display; apply final SDK display/
   visibility last, as a separately reviewed step.

## Python execution and stale-output protocol

The runtime uses provided `sigma` and `session`, standard-library math/json/
hashlib/decimal, `sigma.get_element('Unique title')`, `.columns`, `.collect()` and
`Row.as_dict()`. It does not require repository files in the warehouse. It creates
an explicit typed output DataFrame and calls `sigma.output` once unconditionally,
with a literal no-space name. There are no external packages/network integrations.

Allocation computes `(hc[i]/sum(hc))*overhead*(1+0.5*sin(i))` for one-based indices
1..4 in radians. Full doubles are returned, not rounded/normalized. Null/text/
boolean/nonfinite/negative/zero-total cases return one explicit failure row.
Golden values total `128096.06870260219`, intentionally exceeding `100000`.

The forecast solver reads every current keyed row, all six actual months and
June bases, all twelve forecast inputs and section definitions. It uses actual
plus manual/flat/growth forecasts; only Subscription Revenue's growth is replaced
in memory. It validates keys, periods, numeric values and methods, rejects zero
sensitivity/nonfinite endpoints/unbracketed targets, and uses at most 64 bisection
iterations in `[0,.25]`. ToleranceCents is **1** (`$0.01`), growth precision is 12
places using round-to-even, and the rounded solution is re-evaluated before success.
Fingerprints sort required rows and use Double `.17g` canonical strings, a single
zero, explicit null, and SHA-256 over canonical JSON. Numeric Decimal/float values
with the same Double representation bind identically. Output retains operation ID,
model binding, input fingerprint, physical subscription row ID and prior growth.

1. Prepare a durable Task4 operation (or explicitly supplied test UUID). A retry
   reuses the ID and payload; intentional new work uses a new ID. Store it in the
   one-row invocation with RequestKey=goal-seek and State=PENDING **before** Python.
2. Run the solver and read the output. Do not infer success from the run effect.
   Each attempt returns a new-operation failure row when input reads fail. If a
   platform failure leaves old materialized output, the new pending ID blocks it.
3. Set APPLY_PENDING using the same ID and run the separate apply guard. It reads
   the current inputs/invocation/result again and recomputes identity and residual.
   Check the new output/readback. Do not combine these phases until actual Sigma
   tests prove fresh output visibility to following effects.
4. `build_goal_seek_actions` takes actual element names/IDs and columnIds for
   invocation/assumption/guard plus the read-only inputView. Use current model rows
   and successful result; it refuses mismatches. It emits a conditional update of
   only Monthly Growth, keyed by both physical row ID and Subscription Revenue.
   False conditions do no write. Apply always hides the success toast and never
   marks APPLIED. The controller calls `prepare_goal_action_update` (CLI `prepare-goal-actions`) to
   append per-invocation buttons. This checks the immutable initial proposal,
   current full model graph, existing code/output objects and assigned input/
   column bindings, then uses the same dry-run/fresh-version/finalize gates.
5. Compare complete before/after input rows using `verify_goal_seek_write`. Only
   after a verified exact 12-decimal growth and all other values unchanged may the
   controller reconcile invocation State=APPLIED, AppliedOperationId/AppliedGrowth.
   Read back that operation write too before reporting user success. A repeated
   applied operation no-ops only when the original fingerprint restored in memory
   and the current exact applied value both match.

Ordinary Sigma inputs do not give serializable multi-table compare-and-swap. The
controller must recheck model/row data immediately before the staging click and
prevent concurrent edits during the acceptance case. Interleaved input/model
changes after the guard remain a production concurrency prerequisite; the module
never claims that local guards prove transactional isolation.

## Typed procedure

Only `SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO` is supported here; never
put it in writeback schema SIGMA_WRITE_DB.SIGMA_WRITE. The fixture SQL is the
reviewed V3 COMMENT-before-EXECUTE-AS declaration, CALLER rights, scalar-only and
no DML. It must not be reprovisioned by the worker. Parent receipts govern actual
schema/procedure/grant/index state.

Require exact unique Sigma lookup signature, connection/path, ordered declaration
types, return VARCHAR, no overload/default/VARIANT, catalog ID and actual lookup,
caller-grants, scalar mapping and datetime-coercion evidence IDs. `lookup` records
connectionId/path/signature/returnType/parameters/overloads/defaults exactly. The
caller must be SIGMA_SERVICE_USER/SIGMA_ROLE/SIGMA_WH. Ordered action types are
text, number, integer, boolean, datetime; indexes are 0..4. Null, fractional count,
negative count, invalid date, whitespace operation and changed retry inputs reject.
The reviewed formatting/Double-safe caller subset restricts absolute amount and
count below `1e15`, amount to at most two decimal places; this is deliberately
narrower than the warehouse declaration. Datetime is a supplied millisecond NTZ
value. The emitted Z literal requires actual coercion evidence, not an assumption.

No optimistic scalar toast is emitted. Compare the actual scalar with the exact
expected string. A raised call, timeout or cancellation must record null scalar,
`dependentEffectsExecuted=false` and `successToastShown=false`. Same-ID retry is
safe only for the exact deterministic scalar-only input fingerprint. The procedure
itself has no durable deduplication storage. Failure-diagnostic negative/null calls
need their own narrowly reviewed test action, because the normal typed mapper
rejects them before submission. No artificial sleep or infinite-loop SQL is used.

## Acceptance

Every case in `RUNTIME_CASES` must carry actual Sigma runtime execution/evidence ID,
sourceHash/planHash, target workbook/version/element and passed status. Local
unit tests, dry-run validity and `print(4)` cannot satisfy this ledger. Preserve
published owner/editor/explorer/viewer tests, invalid/failure/timeout/halt/retry,
A-success then B-pending/failed stale output, changed model/input/row, exact single
assumption readback, and independent golden output comparisons. Unknown legacy
helpers/email/refresh remain separately BLOCKED after these components pass.
