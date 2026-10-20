# Regional packet R1 — input-only scope escape

**The reported scope escape is fixed for independent review.** This patch does not begin the grouped Regional build. The 309-table candidate remains blocked. No live operation, Excel access, credential use, production write, nested agent or full-suite run occurred.

The reviewer reproduced the defect with `Sheet!A1=1`, `Sheet!A2 =A1+1`: `executionScope: input-tables-only` plus `mappingStatus: live` passed full staging. Adding claimed targetBinding and verification made completion pass. The input-only gate also accepted live/cached/verification statuses because it checked only the pending-ID list. Receipt-name checks in native construction did not protect the other full gates.

## Correction

- Common `validate_manifest` now rejects the input-only execution scope by default. The explicit `allow_input_only=True` is used only by the literal-input staging gate and read-only row-binding CLI. This opt-in requires every formula's canonical mappingStatus to be **unresolved**, with no logical/actual target mapping, verification, evidence, runtime or execution claims. Input-only Audit formulas also default to unresolved; full-plan Audit defaults remain verification obligations.
- Scope is a checked enum. Full/legacy plans omit executionScope; its only current explicit value is `input-tables-only`. Unknown/null values fail closed. The normal full staging, completion, native and custom-port consumers use default validation and cannot ignore the scope.
- `_approved_feature` independently rejects input-only scope when a caller supplies an alleged `APPROVED_STAGING` receipt, including ordinary full wiring and native assigned-row consumers. Only genuine `APPROVED_INPUT_STAGING` plus matching input-only scope can create/seed literal inputs.
- Hash-only acceptance matrix, assigned-input inventory and report paths check embedded manifest scopes. Production/integration paths inherit that rejection. Promotion finalization also checks its embedded reviewed acceptance matrix, closing the path that otherwise relies on an already-prepared plan hash. Rollback remains available; it does not promote input-only evidence.
- The operator packet now explicitly requires a **new**, separately reviewed full native plan/manifest with no input-only execution scope after actual binding. The first-phase pending manifest is retained unchanged.

## Red/green counterexample and entrypoints

The new `test_input_scope_escape.py` reproduced 29 failing subcases against 0744c13, including the exact reported staging/completion escape and missing pending-status checks. Its nine end-to-end tests now pass, including an additional red/green approval/receipt relabeling counterexample. They exercise 19 full/native/production entrypoints: both manifest validators and migration_gate wrappers, full input staging/completion, legacy native update, exact native scaffold/update, ordinary native binding, full wiring, literal creation with an alleged full receipt, acceptance validation, production validation, reports, integration preflight, acceptance construction, promotion preparation and promotion finalization. They also cover both public staging CLI modes, unknown scope values, the Audit default and genuine read-only assigned-row binding.

The same one-cell declared-live fixture without executionScope still passes the prior full staging/completion gates with its unit-test claimed evidence. This control confirms that R1 closes scope handling without pretending offline evidence is cryptographically authenticated or changing the legacy evidence contract. Genuine unresolved input-only creation/seeding and observed-row binding still pass.

## Verification and preserved scope

- R1 counterexample/CLI tests: **9 passed**.
- Focused native: **93 passed**.
- Adjacent acceptance: **25 passed**; formula provenance: **25 passed**; source staging: **9 passed**; input sync: **23 passed**.
- Packaging: **3 passed**. Changed-file compilation and diff checks passed.
- Promotion focused: **6 passed** after the final approval/receipt relabeling guard.

`work/regional-01-execution-packet-20261020/r1-source-identity.json` verifies all 13 original source SHA256s and approved native plan hashes unchanged. Formula accounting remains **3,651 = 1,865 original five + 1,786 broader eight**: 3,515 ready candidates, two reviewed replacements, 134 unresolved. This scope-only patch does not alter lowering, coverage partitions, actual Excel captures or native runtime status.
