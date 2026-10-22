# Source Workbook Model Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a deterministic source-bound inventory shared by future builders without dropping formulas in mixed, hidden, or scratch areas.

**Architecture:** Keep inventory separate from approval, migration manifests, and Sigma construction. Read workbook facts and use the existing parsed expression model for dependency analysis; retain unresolved features explicitly. A thin CLI writes the validated JSON model.

**Tech Stack:** Python, openpyxl, standard-library unittest/JSON/hashlib, existing `excel_expression` parser.

**Spec:** `docs/superpowers/specs/2026-10-20-source-workbook-model.md`

## Global Constraints

- `modelVersion` is `1`; `sourceSha256` hashes the unchanged source bytes; `sourceName` is its basename. Do not include absolute local paths, wall-clock timestamps, credentials, or Sigma runtime IDs.
- Use the shared `excel_expression` parsed expression model for reference analysis.
- No cache becomes a literal input or completed live calculation.
- The decision changes only role metadata. It must not exclude any cell or formula.
- Source files must not be saved or modified; embedded VBA must not run.
- The five original fixtures, existing tests, migration-manifest version, `BLOCKED` behavior, and strict production-write gates remain unchanged.

## Review Focus

- Formula-looking literal text and explicit false/zero values remain typed inputs; leading `=` alone must not create a formula.
- Local defined names, quoted sheets, and blank range interiors either resolve correctly or have explicit unresolved evidence.
- Mixed tables and totals rows retain per-cell provenance; totals must not become data-entry records.
- Styled empty regions, hidden/protected scratch work, and referenced scratch cells must not be excluded by source-role decisions.
- Tampered identities/counts or stale source-bound decisions fail validation before the CLI overwrites any output.

---

### Task 1: Build and expose the source workbook model

**Files:**
- Create: `skills/excel-to-sigma/scripts/lib/source_workbook.py`
- Create: `skills/excel-to-sigma/scripts/source-workbook.py`
- Create: `skills/excel-to-sigma/tests/test_source_workbook.py`
- Create: `skills/excel-to-sigma/refs/source-workbook-model.md`
- Modify: `skills/excel-to-sigma/SKILL.md` (short inventory-CLI pointer with scope limits)

**Interfaces:**
- Consumes: `excel_expression.parse_expression(formula, sheet=...)`, its AST node classes, and bounded-range expansion. Avoid a separate regex formula compiler.
- Produces: `build_source_model(source_path, *, block_decisions=None) -> dict` and `validate_source_model(model, source_path=None) -> None`, as specified in the companion spec.
- Produces: CLI `source-workbook.py SOURCE --out MODEL.json [--block-decisions DECISIONS.json]` with nonzero failure and no failed-run output overwrite.

- [ ] **Step 1: Add failing source-fact and identity tests.** Generate temporary workbooks in unittest, without relying on external fixtures. Pin exact literal types, exact source formula text, deterministic output, all formulas inside mixed/hidden/scratch cells, table body/totals separation, protection/merges, and chart counts. Assert the source SHA256 before and after is unchanged. Add formula-looking string `=not a formula` with string data type and assert `kind == "literal"`.
- [ ] **Step 2: Run the focused tests RED.** Run `python3 -m unittest discover -s skills/excel-to-sigma/tests -p test_source_workbook.py -v`; capture the intended missing-module/interface failure.
- [ ] **Step 3: Implement source inventory and stable IDs.** Read with openpyxl without saving, preserve JSON types, table column cell provenance and table bounds, and source metadata. Use source locations rather than hardcoded names. Represent unsupported source expressions without interpreting cached values as entered data.
- [ ] **Step 4: Add and implement dependency tests.** Assert `SUM(B1:B3)` includes B2; direct prior-period references retain their exact target; escaped sheet apostrophes resolve; blank targets remain identifiable; supported cycles preserve both directions. Exercise local/workbook defined-name shadowing and unsupported structured/external references. An unresolved reference must have an explicit reason and must not be represented as a verified no-dependency formula. No test evaluates source macros.
- [ ] **Step 5: Add and implement reviewed-role tests.** Hash-bound scratch decisions preserve all cells and edges. Reject changed source hashes, unknown sheets/roles, unbounded or invalid ranges, and overlapping reviewed ranges. Keep table column intent unreviewed/ambiguous unless metadata was explicitly reviewed; do not infer editability from formula ratios.
- [ ] **Step 6: Add and implement validation/CLI tests.** Reject duplicate or malformed identities, wrong version/counts, misclassified formula records, broken dependency bindings, and changed source bytes. Invoke the CLI as a subprocess for successful JSON output and failure preserving preexisting output. Ensure `json.dumps` needs no nonstandard encoder.
- [ ] **Step 7: Run GREEN and existing tests.** Run the focused command, `python3 -m unittest discover -s skills/excel-to-sigma/tests -v`, `python3 -m unittest discover -s tests -v`, and `git diff --check`. Do not alter existing acceptance denominators or expected fixtures to make tests pass.
- [ ] **Step 8: Replay the eight unchanged sources.** Use the local external originals at `/Users/tjwells/Documents/Codex/2026-10-15/can-you-try-this-pr-out-3/work/broader-eight-source-replay-20261020/source-originals/`. Compare every source SHA256 with `../source-discovery-replay.json`; formula counts must equal `[214, 222, 132, 228, 109, 166, 625, 90]` and chart counts `[2, 1, 1, 2, 1, 1, 1, 1]`. Assert main pipeline six literal/three formula columns and twelve cash roll-forward references. Save results outside the repo in `work/source-workbook-model-eight-replay-20261020.json`; never commit the user workbooks.
- [ ] **Step 9: Document and commit.** Explain source inventory versus completed migration and the unresolved dependency contract. Commit the component, tests, and these spec/plan docs. Report status, exact commits, RED/GREEN commands/results, source replay results, and any concerns. Do not push, merge, publish, authenticate to Sigma, modify row data, or spawn agents.
