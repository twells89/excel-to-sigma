# Migration 01 Regional complete qualification — phased packet

**Offline review packet. Input staging is preparable; full native execution is blocked by scaling review.** Root owns every live call and the move into Conversion testing. No production operation is included. The target title is `Migration 01 Regional complete qualification`.

The source is `work/broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx`, SHA256 `4b46bfa6190d1ef12ea35713e3932aa77f19fb964730b4b3bbb0851294219e2c`. Approved input decisions are `work/complete-source-record-bindings-plans-20261020/01-regional-sales-forecast-xlsx.json`; approved native source plan is `work/complete-native-exact-replay-r1-20261020/01-regional-sales-forecast-xlsx.json`. Those offline plans use review-placeholder scope, which is not a staging authorization. Rebuild them with the actual approved staging connection/folder; never edit a hash by hand.

## Stop before a bulk native build

The current lowering emits **309 tables / 13,749 columns** for 214 formulas. `scaling.json` inventories all 13 sources, record profiles, AST/function counts, table categories and width histograms. Do not stage hundreds of helper elements. The 64-column cap is already used (157 Regional tables have 64 columns); an arbitrary 16-column setting is not the cause. See `scaling-assessment.md` for the grouped architecture needed next. No giant MCP program or fabricated actual target identity is supplied.

## 0. Root creates the disposable shell; retain actual returned identities

Root begins the approved staging Sigma session and verifies its write-enabled connection. Root uses the MCP builder with `target.type = new` exactly once and the title above. Subsequent calls use the returned workbookId and `target.type = existing`. A minimal shell program is:

```python
for name in ['Forecast inputs', 'Regional targets', 'Summary regions', 'Summary months', 'Summary products', 'Calculation qualification']:
    p = sigma.workbook.add_page(name)
    print({'pageName': name, 'pageId': p.id()})
```

Keep the existing initial page; no foreign page is deleted. Root moves the disposable workbook into Conversion testing separately. Save the complete fresh GET as `run/00-current.json`, including workbookId, documentVersion, contents, pages, XML layout, elements and privateState. Keep the exact returned workbook URL. SDK frames must stay within 24 layout columns; this packet uses a separate page per ordinary input table, with width 24.

## 1. Prepare and review the input-only scope, then create/seed one table

Set `repo` to the **integrated checkout** and `task` to the task root. `run` is a new local receipts directory. Set `actual_connection` and `actual_folder` only from root's approved staging observations.

```sh
python3 "$repo/work/regional-01-execution-packet-20261020/prepare_input_review.py" \
  --repo "$repo" --task-root "$task" --connection "$actual_connection" \
  --folder "$actual_folder" --current "$run/00-current.json" --out "$run"
```

This writes actual-scope input/native plans, the source inventory, input-only review plan/manifest, the source audit and a **pending request**. It creates no approval. Root reviews the exact source audit issues, all 214 pending formula IDs, the literal schemas/rows/natural keys and the actual page bindings. In a separate `input-approval.json`, supply `status: approved_for_staging`, named owner/reviewer, `scope: input-tables-only`, exact hashes/reviewItems/pendingFormulaIds/mappedFeatures and targetBindings for all five features. Each binding contains the fresh actual documentId, integer version, actual pageId. For audited issues, add `sourceReviewApproval: {status: approved_for_staging, issues: <exact source-gate issues>}`. Do not mark any source issue verified.

```sh
python3 "$repo/skills/excel-to-sigma/scripts/input-sync.py" stage --inputs-only \
  --input-plan "$run/input-plan.json" --plan "$run/input-review-plan.json" \
  --manifest "$run/input-manifest.json" --source "$task/work/broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx" \
  --approval "$run/input-approval.json" --out "$run/input-staging.json"
python3 "$repo/skills/excel-to-sigma/scripts/input-sync.py" emit-create \
  --manifest "$run/input-manifest.json" --receipt "$run/input-staging.json" \
  --feature sales-forecast --out "$run/01-sales-forecast.py"
```

Root reviews/executes that **single** MCP program, then saves the full GET and describes/queries the actual new input element. It creates only literal editable columns, adds the system row ID, uses the actual returned column handles to seed 144 typed literal rows, and prints actual element/column IDs. Revenue is not seeded or frozen. Do not retry creation after an interrupted or partial result: GET and reconcile first.

Then repeat individually for `regions`, `summary-regions`, `summary-months`, `summary-products`, after refreshing current document/version and reissuing the exact input-only receipt. The input schema is:

| Feature | Rows | Editable fields | Natural key |
|---|---:|---|---|
| sales-forecast | 144 | Region, Product Line, Month, Units, Avg Price | Region, Product Line, Month |
| regions | 4 | Region, Region Manager, FY26 Target | Region |
| summary-regions | 4 | Region | Region |
| summary-months | 12 | Month | Month |
| summary-products | 3 | Product Line | Product Line |

The five tables retain 751 literal source-cell bindings and 11 editable field definitions. Source scalar IDs and source coordinates remain provenance; physical row IDs are never natural business keys. For resumption, query actual rows and use `input-sync.py plan` with the actual system row-ID alias, fresh snapshot ID/age and prior seed ledger; its receipt accepts only the approved literal schema/rows. Never seed calculated fields or overwrite a user's conflicting edit.

## 2. Bind actual IDs/rows and request a separate native review

Save `run/inputs-current.json` and `run/observed-rows.json` (feature → complete typed rows, each including its observed system row-ID alias). Normalize nothing silently: preserve null, text, Boolean, numeric and datetime identities. Root records actual elementId, the exact editable name→columnId map and actual rowIdColumnId/system alias in all five targetBindings, at the fresh version. Refresh the input manifest target and input staging receipt with those real identities.

```sh
python3 "$repo/skills/excel-to-sigma/scripts/native-calculations.py" bind-input-rows \
  --input-plan "$run/input-plan.json" --manifest "$run/input-manifest.json" \
  --staging-receipt "$run/input-staging.json" --current "$run/inputs-current.json" \
  --observed-rows "$run/observed-rows.json" --out "$run/assigned-rows.json"
python3 "$repo/skills/excel-to-sigma/scripts/native-calculations.py" lower-candidate \
  --native-plan "$run/native-plan.json" --assigned-rows "$run/assigned-rows.json" \
  --max-columns 64 --out "$run/native-construction-review.json"
```

The second command is **offline construction review only**, not an instruction to execute its 309 tables. The separate native review plan must include the actual inputPlanHash/sharedInputArchitecture/nativeCalculationPlan and named `nativeReviewApproval` bound to sourceSha256/nativePlanHash. Its formulaCells use the actual lowering's logical five-lane targetMapping, all 214 retained; source Audit/verification obligations are not reclassified as passed. Root rebuilds the manifest from the original source and passes the existing full `input-sync.py stage` gate with a newly reviewed approval. The input-only receipt cannot authorize a native formula scaffold, formula REST update, or completion. Focused negative tests enforce this separation.

## 3. Grouped SUMIFS canary — requires a small reviewed construction first

**BLOCKED pending grouped architecture implementation/review.** Do not use an arbitrary first scalar helper as evidence for Regional SUMIFS. The relevant canary is the exact source vertex `Summary!B4`: live Revenue by source record, followed by Region=`Summary!A4` (initially West), Month ≥ DATE(2026,1,1), Month < EDATE(DATE(2026,1,1),3), then one typed sum/error output. Its complete source range remains the formal table, including appended records; do not truncate to the initial nine matching records.

Use a few ordinary calculated tables: one shared row Revenue/error layer, one criteria/group layer and one five-lane group output (additional helpers only if explicitly reviewed). Root's actual singleton seed must be a saved SQL `SELECT 1 AS C0` element on the approved connection, with fresh document/version/element receipt and queriedRows exactly `[{C0: 1}]`. SDK v2 does not permit in-program SQL; create this tiny saved seed through root's documented guarded structural path and GET/query it, rather than inventing a query API or seed ID. If no observed compatible seed exists, stop before native formula construction.

After a bounded grouped canary exists and its plan is reviewed, use the existing protocol **one dependency vertex at a time**: `native-calculations.py backup`, `prepare-scaffold`, execute its MCP program, full GET, `confirm-scaffold`, append its actual receipt to native-bindings, then `prepare-update --native-bindings`, guarded REST PUT with the fresh documentVersion, full GET, `confirm-readback`, query five lanes. A contiguous confirmed scaffold prefix is now allowed, so all later tables need not be built to test the current vertex. Unknown/holed receipts, missing prior saved formula readbacks, or changed unrelated pages/elements/private state still fail closed. The final formulas come only from fully parenthesized REST text; never round-trip them through SDK serialization.

For each command pass the reviewed native/input plans, manifest, APPROVED_STAGING receipt, current full GET, complete observed rows and exact backup. `prepare-scaffold` also needs observed seed-binding/page-id/shard-index; `prepare-update` needs the actual native-bindings map and prior-readbacks array. Each command's `--out` is a distinct immutable proposal/receipt. Execute only the proposal's `program` or `dryRunRequest` after review; an emitted candidate is not evidence of saved or calculated behavior.

Hand-check the useful canary edits: first West/Hardware/2026-01-01 Units 493→494 changes Revenue 152830→153140 and Q1 West by +310; clearing that Units field gives Revenue 0 and Q1 West −152830 from baseline. Restore the exact literal, reopen and re-query. An appended natural key West/Hardware/2026-01-15 with Units=2 and Avg Price=10 should add +20 to Q1 West, while the source's exact-month Jan-01 summary stays unchanged. Remove the appended row by its observed physical ID and confirm restore. These are expected deltas for future qualification, not recorded runtime results.

## 4. Explicit remaining gates

All 214 formulas remain required: 144 row Revenue formulas, 52 SUMIFS, four VLOOKUP, four IFERROR ratios, nine totals and the source `="Total"` expression. Regional has no Audit worksheet; retain source totals and all 214 provenance obligations without inventing an Audit pass. Source inventory counts two charts; charts, formatting, layout and business-facing calculated-table presentation are later phases.

`regional-excel-review.json` records actual source capture: G4:G7 are **VLOOKUP**, with missing AppleScript values, unavailable serials and displayed `#NAME?`; H4:H7 actually evaluate to zero through IFERROR and differ from local intended candidates. Root must review/reconcile these eight results before source parity can pass. F8 is numeric 21521822.24 with `#########` display overflow. The 144 Revenue formula2 strings use `[@...]` syntax; their source strings differ and each rebound AST is checked separately, with review status retained. No raw-null equivalence, cache fallback or implicit lookup replacement is approved.

`native-semantics-conformance.json` separately verifies all 30 actual vector captures against exact workbook/source/formula2 identity (30 integrity PASS, zero formula2 differences), and all six negative/positive ROUNDUP cases against both independent source evaluation and the corrected emitted scalar reference. This is six-case Excel conformance, not live Sigma qualification. The 24 Boolean criteria vectors remain captured; unsupported Boolean criteria lowering stays unchanged.

Completion still requires source decisions, saved exact formulas, actual value/type/error parity, blank/error and missing/duplicate lookup cases, natural-key uniqueness, relevant period boundaries, input edit/reopen/restore, append behavior, two charts and presentation. No completion or production approval is supplied.
