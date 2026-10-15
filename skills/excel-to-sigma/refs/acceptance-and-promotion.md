# Workflow repair and acceptance controller

All tools here are offline authoring and evidence validators. The approved parent
controller makes actual Sigma calls and captures receipts, rows and browser steps.
A generated fragment, mapping review or successful SDK mutation is not runtime
acceptance. Existing code-rep/SDK input construction and display ledgers still apply.
Do not run this suite against customer data. The integration adapter requires an
explicit synthetic scope and opt-in; credentials stay inside the controller's
environment, never JSON bundles, Git, captures or reports.

## Repair workflow

Use `scripts/migration-acceptance.py COMMAND --bundle arguments.json --out result.json`.
The bundle is a JSON object of the public library's named arguments. `repair-contract`
instead takes `--source SOURCE.xlsx`. All output JSON is deterministic except the
caller's actual evidence. Python modules are in `scripts/lib`.

1. Extract `workflow_repairs.extract_workflow_repairs(source_path)`. This records
   every original formula (1865 across all fixtures), exact source formats,
   secondary literal metadata and explicit unresolved runtime status. It never
   uses formula cached values as inputs. Forecast currency display retains source
   zero decimal precision; Excel's separate zero-as-dash sections remain partial.
2. For Sales, `extend_repair_manifest(base_manifest, receipt, contract, review)`
   introduces exactly four Rep Target rows. Review binds `contractHash`, named
   `owner`, `status=APPROVED_STAGING`. Prepare new artifacts/approval with the
   existing `build_artifact_records`/staging APIs, construct the ordinary input
   using `emit_construction_program`, then capture all assigned IDs and seed/read
   back. Existing receipts cannot authorize an extended plan. Source stage and
   supplier lists are literal reference data; they do not override item inputs.
3. `prepare_repair_update(manifest, receipt, contract, current, bindings,
   scope=scope, review=review)` consumes actual assigned bindings. `scope` names
   documentId, version, connectionId, folderId and actual `folderEvidenceId`.
   Do not infer a folder UUID from the workbook GET's string `path`.
   `bindings.calculation` contains elementId, name and exact elementHash;
   Sales also requires `bindings.repTargets` from the approved receipt.
   Review contains owner, status, contractHash, planHash, beforeHash,
   identical `sourceBindings`, and `inputSystems[feature]` for every input:
   targetBindingHash, columnId and evidenceHash. The columnId is the actual
   code-rep-readable row-ID column/projection established by SDK and query
   evidence, not an invented alias or an assumption that SDK IDs equal REST IDs.
   If a native SDK row ID is not addressable in code-rep, first use the existing
   `native_actions.emit_input_identity_projection`, read back its assigned IDs,
   and review the projection as the source. Do not substitute a normalized REST
   `ID` for a different native SDK row ID without proof.
4. `readonlyFormats[elementId]` supplies exact elementHash and `sourceCells`
   mapping actual view column IDs to source cells. Supported formats and rules
   propagate to ordinary readonly tables; source rule priority is preserved.
   `technicalVisibility[elementId]` supplies elementHash and exact technical
   columnIds to hide after the SDK's additive propagation. Formula/order/chart
   encoding properties remain intact. `cleanup` supplies named owner,
   individually hashed scaffold pages/elements and actual landingPageId;
   foreign placements or active input storage cannot be hidden. New pages
   provide calculated results, reference access, navigation and full input IDs.
5. Expense additionally requires `review.currencyPatch` with elementHash,
   actual Exception columnId, exact beforeFormula and
   decisionHash=`digest(contract.currencyDecision)`. This is an explicit
   sourcePrecision=2 replacement: compare rounded integer cents, preserving
   original stored amounts. It is not claimed to equal raw FLOAT subtraction.
6. Review the generated complete contents. Call actual create dry-run
   `POST /v2/workbooks` with `proposal.dryRunRequest`. Record httpStatus=200,
   dryRun=true, evidenceId, proposalHash, errors/warnings. Then fetch fresh
   original contents and call `finalize_repair_update`; consume proposalHash
   durably before sending its fresh-version PUT. Ambiguous sends require
   readback, never blind replay. Any warning needs exact warning hash review.
7. REST can reset input display even when input JSON is unchanged. Reapply and
   verify the approved SDK display programs after the contents PUT. Rebuild
   receipts against the fresh actual version/IDs; do not merely bump an old
   receipt's version. Capture rows, names/types/classes, options and full display.
   Width/CF evidence requires separate document/browser observation; the compact
   SDK tuples establish format and visibility only.
8. `account_formula_targets(contract, targets, document_id=..., version=...)`
   requires one actual native column or explicit audit check for every source
   formula with sourceFormulaHash. Mapping does not set runtime verification.
   Research's source Audit can legitimately say REVIEW after an input edit.

Existing physical input names stay stable. Research adds a source-derived guide
for raw growth field names and a normalized metadata spine joined to the LIVE
wide input by Scope; Switch reads every numeric input and stored row ID. It
never displays original seed values as a substitute for live edits. Parent must
review/query the actual join and every normalized value after construction.

## Persisted workflow result UI

`operation_outcomes.build_outcome_view(binding)` projects persisted operation
state. Pending/prepared states say "Pending verification". It has no optimistic
action-chain success effect. `reconcile_operation(expected,capture,operations,
completed_at=...,manifest=...,receipt=...)` requires the validated source manifest
and its APPROVED_STAGING receipt for the assigned workflowOperations input.
The complete operations capture must carry sourceHash/planHash/documentId/version,
origin=sigma-query, executionId/evidenceId, exact assigned binding, typed SDK
schema and a rowIdBinding (sourceColumnId, readbackColumn, targetBindingHash,
reviewed evidenceHash). The approved target supplies rowIdColumnId, or the
reviewed systemColumn alias. This separates the physical SDK row ID from the
query label. The reconciler then runs the existing native business audit against complete
before/after rows, exact operation/context/source/plan/document/version and a
fresh sigma-query evidence reference. It emits a narrow SDK follow-up setting
CONFIRMED only for the original pending operation. Execute it serially per
workbook, query every operations row again, and use `verify_outcome_row`. Its
sealed plan binds the original capture, assigned destination, complete intended
rows, exact outcome text and completion timestamp. Verification requires a new
query/evidence identity, the same target/version/schema/row-ID binding, and
exact intended fields with every unrelated row preserved. The SDK has no
compare-and-swap/concurrent uniqueness guarantee; cross-effect partial failure,
wrong-row prevention and recovery still require published UI acceptance. Failed
or unconfirmed business reads produce no CONFIRMED write.

## Mandatory acceptance and reports

`acceptance.build_acceptance_matrix(manifests, targets, receipts=..., oracles=...)` requires all
five validated source manifests and their APPROVED_STAGING assigned receipts.
Each manifest must name its current target document/version. Targets record actual staging documentId,
version, url, sourceStatus, unresolvedSourceFeatures and isolated warehouse input
identities (connectionId + tableIdentity). inputTables must cover exactly every
feature in the final sharedInputArchitecture, including settings, companions and
operations. Each descriptor names assigned documentId/version/pageId/elementId,
source connectionId, tableIdentity, targetBindingHash and identityEvidenceId.
The sealed matrix retains the approved manifest/receipt; production and isolation
validation recheck coverage and reject omitted, extra, stale, foreign or duplicate
assignments. `assert_isolated_inputs` refuses
original/shared writeback identities; a copied workbook ID alone is insufficient.
All mandatory case templates remain unresolved until evidence is recorded.
Oracles are independently reviewed source equations/goldens, not translator
outputs: origin=independent-oracle, reviewer, sourceHash, keyed checks with
expected/comparison (money, ratio with explicit precision, date, exact).

`record --bundle {matrix,capture} --state ledger.json --out checkpoint.json`
locks and atomically appends the hash chain. Capture includes exact caseId,
evidenceId, sourceHash, planHash, documentId, version, executionId, references,
origin=sigma-query or sigma-browser, actual keyed values, receipt containing
COMPLETED and the same execution/document/version. Browser proof additionally
requires published=true/publishedVersion, steps, screenshots, actual session
memberId/accountType/effectiveGrant/editInputTables, identityEvidenceId and
adminBypass=false. Role cases also require dataEntry.inputMode,
editDataAvailable and either writeCount=0 or validSaved/invalidRejected/
cancelPreserved=true. Real published availability derives from entitlement,
effective inherited grant AND table mode. Explorer is the selected Build user
with Explore document grant, not Analyze-account parity. Do not broaden roles.
Persistence captures contain complete before/changed/reopened/restored rows,
rowIdColumn, keyColumns, allowedChanges and reopen/restore evidence IDs. At
least one final business value must differ from before; same-value and cancelling
edits fail. Each allowedChanges entry may declare comparison/precision (default
exact); changes below its declared money/ratio/date precision do not prove an edit.
Reopened, restored and allowed-delta snapshots still compare complete typed rows.

`reports --bundle {matrix,ledger,final_review} --out reports.json` renders portable
per-workflow structured results and Markdown. Each report explains workflow
rules and shows genuinely passed, failed and unresolved mandatory denominators,
source formulas, URL/version and evidence references. Final review binds matrix,
ledger head and exact branchCommit. Source REVIEW_REQUIRED/unresolved features
cannot become VERIFIED_COMPLETE. Reports call
`validate_production_matrix` before any completion verdict. Reduced component
reports retain their own denominators, but say coverageStatus=COMPONENT_ONLY,
verdict=BLOCKED and explain completionBlockers; final approval cannot bypass
the full mandatory suite. Promotion retains the same production gate.
`run_integration(matrix,adapter,scope,enabled=True)` is opt-in controller injection:
adapter.capture receives case/target/designated synthetic scope and returns
actual evidence. Before the first adapter call, scope.stagingInputs must equal
the exact matrix union keyed workflow:feature, with no missing/extra table,
shared original storage, duplicate storage or connection outside scope. It does not log credentials or invent an HTTP/SDK execution API.

## Original URL preserving promotion and rollback

`promotion.build_promotion_backup(scope,current,rows,sdk,data_models)` requires
named owner/sourceHash/planHash, original documentId/version, full contents,
every affected input row with stable ID and unique business key, complete typed
SDK schema and display/visibility/options tuples, and complete affected DM
contents/versions. Hashes bind the entire backup. Capture before any warehouse
write. System audit timestamps cannot be recreated; do not claim otherwise.

`prepare_promotion(backup,fresh,candidate,review,acceptance)` preserves the original
ID and URL. Review names owner/backupHash/candidateHash/targetDocumentId,
replacementIds and allowNewIds. Acceptance is the full {matrix,ledger,finalReview}
bundle; a boolean or "complete" string is rejected. All five complete mandatory
suites and final reviewed exact candidate/original target are required.
`finalize_mutation` additionally requires actual hash-bound dry-run and fresh
version, then returns a consume-before-send PUT. Checkpoint each real outcome
with `checkpoint --bundle {backup,event} --state ... --out ...`; no successful
step replay. Inspect actual state after ambiguous/failed calls before resuming.

`prepare_rollback(backup,fresh,current_rows,owned_added_rows=...,fresh_data_models=...)`
restores original business values by immutable row ID, affected DMs with fresh
versions, original workbook document, then SDK display LAST, then complete
readback. It refuses missing original IDs and unrelated extra rows. Only
explicitly owned added row IDs can be deleted. Its narrow SDK programs guard
page/table/complete classes/names/types/options before mutation. Reverting REST
alone cannot undo warehouse values or restore display. Run `verify_restoration`
with the complete actual document, rows, SDK display and DMs after rollback.
Backup/promotion/checkpoint JSON cannot authenticate itself: the designated
controller captures real calls; independent review verifies their provenance.

`secondary-input-hosts --bundle {contract}` emits dedicated new assumption pages.
New Rep Targets placement starts below its source guide; bind the new input to
that dedicated reviewed page, never to an old page with unreviewed overlapping
frames. `published-mode` consumes manifest, receipt, feature and exact reviewed
restricted published-mode decision and returns the SDK program. Existing primary
input modes are preserved; only reviewed SalesSettings/ExpensePolicy are changed.
The repair proposal automatically includes a persisted Workflow results view
when the receipt contains workflowOperations.

`rebind-repair` calls `rebind_after_repair(manifest,receipt,proposal,readback,
sdk_observations,review)`. SDK observations have documentId, version, evidenceId,
and tables[feature] with elementId and actual column descriptors. Review binds
contentsHash/sdkHash and systemColumns[feature].observationHash,
rowIdEvidenceHash and, for a FormulaColumn normalized alias, approvedAlias using
the existing complete row-ID contract. Source-list conversions require fresh
columnContracts. Exact added contents must round-trip; unknown normalization
stops. The result is a fresh assigned receipt/target version, never an edit of
an old receipt's version field. Formula accounting also requires fresh
current_contents and actual elementHash; explicit audit checks additionally
name independently reviewed source equations.

The five functions in `fixture_regression.py` are reusable independent hand
source equations for the designated fixtures, including edits keyed by exact
source cell. They read literals only and reject a formula cell in the literal
path. They cover Forecast144 monthly values and234 P&L/FY values,132 expanded
records plus secondary rows, and420 Research values including all equity inputs.
Research's archived JSON golden rounded EPS before some downstream ratios;
source39/44 formulas instead divide by live EPS. Tests keep this distinction
explicit and compare source equations rather than freezing that rounded cache.
The existing native/AST regression suite remains in place. Actual null entry,
validation, cent boundary and all published behavior still require actual tests.

## Separate decimal text transport adapter

Parent actual provisioning established the distinct
`SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO_TEXT_AMOUNT` wrapper; exact
reference source is `tests/fixtures/task6-decimal-text-procedure.sql`.
`decimal_procedure_adapter.controlled_decimal_text_contract(connection_id)`
starts with signature=null and BLOCKED. It binds wrapper body trimmed SHA256
`1df68a1cc86017ef93cebcc5e554873894080ff6f029ba46630393d3217b9abc`
and unchanged original callee
`f544541aa79453bcd828baa304c6c1ecc4c7249bf1f181fd3dfbe8642c1960ed`.
It rejects binary floating amount transport. The exact reviewed ASCII text
amount is checked before fixed-point conversion by the wrapper. All five
original business arguments remain ordered and typed; count is integer and
NTZ timestamp uses the independently date-validated native Datetime formula.

`decimal-adapter --bundle {contract,arguments,prior_invocation}` emits an action
only after actual catalogObjectId, lookupEvidenceId, callerEvidenceId,
scalarMappingEvidenceId, datetimeEvidenceId, provisioningEvidenceId and
signatureEvidenceId, plus exact ordered lookup metadata, are present. No guessed
or PENDING signature is accepted. `build_decimal_text_port_fragment` provides a
new, guarded-merge-compatible isolated button page. Same-operation retries bind
the exact text spelling/contract. Verify returned exact scalar, invalid/NULL
cases, no dependent effects after error/timeout and actual roles. Provisioning
is not native scalar parity. Original NUMBER transport remains separate and
blocked: actual Sigma lost1234.56 cents while direct typed SQL preserved them.

Warehouse Python runtime maps are exact per element/logical column. When actual
metadata requires distinct Snowpark identifier and collected Row key, pass
`{"identifier":"\"Line Item\"","rowKey":"Line Item"}`; a string means both
are identical. The emitted program validates both, with no casefold, unquoting
or fallback. Every Sigma accessor/output name is a literal. Native MCP action
programs use individual isinstance-narrowed local column handles, not dictionary
subscript narrowing or unsupported Python annotations.

## Mandatory SDK action backups after actual REST omission findings

Public REST GET has actually omitted persisted SDK-native Prepare and selection
sequences. It is not a complete action backup. Before every repair/promotion PUT,
provide a fresh SDK action inventory covering **every document element**, including
explicit SDK observations of zero sequences. `action_inventory.empty_action_capture`
is only a `TEMPLATE_NOT_RUNTIME` shape helper and cannot pass validation.

Inventory: `{documentId, version, evidenceId, origin:"sigma-sdk-readonly",
complete:true, elements:{elementId:{elementId, complete:true,
inspection:"list_action_sequences"|"sdk-type-no-action-events", sequences:[...]}}}`.
Each sequence records actual id/name/trigger/disabled/condition/ordered blocks/
successToast. Typed blocks record all table/control/page targets, values,
selectors and insert return-variable bindings. Toasts include title/message/
formulas. Unknown or incomplete block types stop. Known API-created Python or
procedure blocks may use an explicitly reviewed dual REST/SDK artifact when the
SDK cannot represent that effect; REST absence alone never establishes absence.

Current SDK Formula has no expression getter. Store opaque observations verbatim
and bind them to the exact known construction artifact and assigned IDs; never
decompile or evaluate their repr. For nonempty sequences, `semanticHash` is the
canonical digest of `configuration(element)`, which preserves all observable
configuration except assigned sequence IDs. It is recomputed, so changing a
condition, disabled flag, block or toast without updating the actual reviewed
configuration cannot pass preservation.

Every action-bearing element requires `restorationArtifact`: elementId,
documentId, reviewer, sourceObservationHash (digest of the element observation
excluding that artifact), semanticHash, program/programHash, constructionArtifactHash,
bindingHash, guardPolicy exactly
`missing-sequences-only; exact-current-target-and-bindings; no-constructor-replay`.
Its `restoreProof` must record actual PASSED_ISOLATED_RESTORATION,
beforeMissingSequences=true, evidenceId, isolated documentId,
sdkReadbackEvidenceId and identical programHash/bindingHash/semanticHash.
These are reviewed per-action restoration programs; the tool does not generate
a fictitious generic SDK action decompiler. Opaque/unknown actions without that
actual proof remain BLOCKED_ACTION_RESTORATION. Full constructor replay after
partial persistence is prohibited. Only reviewed, current-state, missing-sequence
restoration is eligible for execution, followed by full fresh SDK readback.

`build_promotion_backup` additionally takes `sdk_actions` (or the explicit
`scope.sdkActionInventory` capture). Repair review also requires
`sdkActionInventory`. Both document finalizers require `fresh_sdk_actions`.
`rebind_after_repair` requires post-REST `sdk_actions`; `verify_restoration`
requires complete final `sdk_actions`. `verify_action_preservation` compares all
old elements' observed configuration before/after independently of REST JSON.
Rollback order is rows → data models → document → reviewed missing SDK actions →
SDK display → actual readback. When an edit mode was explicitly changed,
`changed_edit_modes` names the expected current mode; only those modes are
restored. Recreated sequence IDs and immutable warehouse audit timestamps are
qualified separately rather than falsely claimed identical.

The parent observed a partially persisted native constructor at Operational v7:
SDK typing succeeded, but a BoolFormula row selector failed `invalid_formula`
after five controls, two buttons and Prepare/selection sequences persisted.
Consumed attempts cannot be replayed. The selector/recovery and zero/one/multiple
selection behavior remain actual runtime acceptance gates. A global
CountDistinct(column) is not assumed to equal the selected-row count and is
never silently replaced with constant1. Forecast source inspection subsequently
proved quoted identifiers differ from collected Row keys at v9; explicit paired
runtime maps address that compiler/runtime boundary.
