# Planning App Versions Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Snapshot in the planning app freezes the forecast's inputs, and the Versions & approvals page compares any saved version to today's forecast by line item and month.

**Architecture:** Snapshot's one `insert-rows` gains one column per driver (effective growth rate per rate line, effective manual value per manual line × forecast month) plus `Inputs Frozen = true`. A "Compare to" control picks one frozen version; a one-row `versionPick` element exposes its inputs; the version's forecast is rebuilt by the same builder function that emits the live forecast stages, fed the version's inputs instead of the live control / Parameters / input tables. A compare-facts union of today and the version feeds the compare section.

**Tech Stack:** Python 3 (skill venv), unittest, Sigma workbook code-rep specs (`POST /v2/workbooks`), Sigma MCP Code Mode, the local spec evaluator `tests/planning_app_eval.py`.

**Spec:** `docs/superpowers/specs/2026-11-10-planning-app-versions-design.md`

## Global Constraints

- Scope: only the `--workflows` (planning app) path. The plain forecast build's spec stays byte-identical to `main`.
- Driver columns exist only when the plan is editable / what-if (the planning app has `rateInput` / `manualInput`); a read-only planning app's Snapshot is unchanged except `Inputs Frozen` is not written (nothing to compare).
- No warehouse views, no views over writeback tables, no stored procedures, no visible workarounds.
- Agents never type APPROVE; staging posts need the named human reviewer's `--approve-for-staging` at their own terminal.
- The setup program is run unchanged as one `build_workbook` program.
- Migrated (seeded) Versions rows never get `Inputs Frozen`; they are never offered in Compare to.
- Page copy, verbatim: `Numbers not recorded in the source` (migrated rows); `Rebuilt from the saved inputs with actuals as they are now` (compare section note); `No saved versions to compare yet` (compare section, no frozen version).
- A fresh build never holds a Sigma snapshot saved before this feature, so the spec's second label ("Saved before inputs were frozen") is not built; every row without `Inputs Frozen` uses the migrated label. The spec is amended to match.
- Live probes post into TJ's home folder (`57e59735-86b9-40b0-b029-217205406f57`), are named `… (delete me)`, and are moved to trash when done.
- Commit trailer: `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

1. A rate line nobody edited (no `rateInput` row) → the version stores the line's default rate, not NULL. Test in Task 2.
2. A version saved in Upside, compared while the live switch is Base → the rebuild uses Upside's spread. Test in Task 4.
3. Roll forward after a snapshot → the version keeps its own start month: its first forecast month is still forecast while today's is actual. Test in Task 4.
4. No frozen versions yet (only migrated history) → the compare section shows an empty state and its KPIs are blank, not errors. Test in Task 5.
5. A plan with a `pct_of_total` line → the rebuild includes the second pass, and a share line reads the version's own rebuilt base section. Test in Task 4.

---

### Task 1: Live probe — one Snapshot pulls every driver value

Throwaway, no repo code. Decides the driver-pull formula every later task uses.

**Files:**
- Create (scratchpad only, not committed): `probe_drivers.py`

**Interfaces:**
- Produces: `PULL_RATE` and `PULL_MANUAL`, the two formula templates that passed, recorded in this plan file under "Probe result" (Step 5) and used by Task 2.

- [ ] **Step 1: Build the probe workbook.** Mirror the planning app's names: input table `Growth rates` (`LINE_ITEM` text "Line item", `RATE` number "Monthly growth rate"), input table `Manual forecast` (`LINE_ITEM`, `MONTH_NUM` number, `AMOUNT` number), input table `Versions` with 2 rate columns and 2 manual columns plus `Inputs Frozen` (boolean), and one button whose `insert-rows` writes all of them. Write each rate column with candidate A, `Coalesce(Max(If([Growth rates/Line item] = "<line>", [Growth rates/Monthly growth rate])), <default>)`, and duplicate columns with candidate B, `Coalesce(SumIf([Growth rates/Monthly growth rate], [Growth rates/Line item] = "<line>"), <default>)`. Do the same for manual cells, keyed on line and `[Manual forecast/Month Num] = <n>`. Use one line with an input row and one line without, so the default branch is exercised.
- [ ] **Step 2: Dry-run, post, then run the generated setup program** (manual commit, seed 1 rate row and 1 manual row).
- [ ] **Step 3: Click the button in the browser pane** (TJ signs in if asked). Query `Versions` via Sigma MCP `query`.
- [ ] **Step 4: Pass condition.** The edited line holds its input value, the unedited line holds its default, and `Inputs Frozen` is true, for at least one candidate. If neither candidate passes, **stop the plan and report back** (spec call 7).
- [ ] **Step 5: Record the passing templates** in this file under "Probe result" below, trash the workbook, commit the plan edit.

```bash
git add docs/superpowers/plans/2026-11-10-planning-app-versions.md
git commit -m "docs(excel-to-sigma): record the version driver-pull probe result"
```

**Probe result:** _(filled in by Step 5)_

---

### Task 2: Snapshot writes the driver columns

**Files:**
- Modify: `skills/excel-to-sigma/scripts/lib/demo_workflows.py` (`TABLES[VERSIONS]`, `input_table`, `action_effects`, `workflow_elements`, `source_history`)
- Modify: `skills/excel-to-sigma/scripts/build-forecast-model.py` (pass drivers into `demo_workflows.workflow_elements`)
- Test: `skills/excel-to-sigma/tests/test_demo_workflows.py`

**Interfaces:**
- Consumes: Task 1's `PULL_RATE`, `PULL_MANUAL`.
- Produces:
  - `DriverColumn = NamedTuple("DriverColumn", [("cid", str), ("name", str), ("kind", str), ("line", str), ("month_num", int | None), ("default", float)])`. `kind` is `"rate"` or `"manual"`.
  - `version_drivers(P) -> list[DriverColumn]` in `demo_workflows.py`:
    - One `rate` column per plan line whose kind is in `RATE_KINDS` (`growth`, `yoy`, `pct_of_total`), in plan order: `cid = "ve_r" + str(i)`, `name = f"{line} rate"`, `default = line["rate"]`.
    - One `manual` column per manual line × forecast month: `cid = f"ve_m{i}_{month_num}"`, `name = f"{line} {month_label}"`, `default` = the plan's manual value.
  - `INPUTS_FROZEN = ("ve_frozen", "Inputs Frozen", "boolean")`.
  - `workflow_elements(..., drivers: list[DriverColumn] | None = None)`. With `drivers`, the Versions table declares those columns plus `ve_frozen`, and Snapshot's `insert-rows` writes each driver through its template and `ve_frozen` as the constant `true`.

- [ ] **Step 1: Write the failing tests**

```python
def test_version_drivers_cover_every_rate_line_and_manual_cell(self):  # demo plan: 10 growth, 1 manual × 6 months
    drivers = dw.version_drivers(self.plan)
    self.assertEqual([d.kind for d in drivers].count("rate"), 10)
    self.assertEqual([d.kind for d in drivers].count("manual"), 6)
    self.assertEqual(len({d.cid for d in drivers}), 16)

def test_snapshot_writes_every_driver_and_inputs_frozen(self):
    wf = dw.workflow_elements(self.manifest, "conn-1", "FY27 forecast", FIGURES, drivers=dw.version_drivers(self.plan))
    values = snapshot_values(wf)   # the Snapshot button's insert-rows values
    for d in dw.version_drivers(self.plan):
        self.assertIn(d.cid, values)
        self.assertIn(repr(d.default), values[d.cid]["formula"])   # unedited line falls back to its default
    self.assertEqual(values["ve_frozen"], {"type": "constant", "value": {"type": "boolean", "value": True}})

def test_migrated_history_never_sets_inputs_frozen(self):
    seeds = dict((el, (cols, rows)) for el, cols, rows in dw.source_history(DEMO_KIT_SOURCE))
    self.assertNotIn("ve_frozen", seeds[dw.VERSIONS][0])

def test_no_drivers_keeps_todays_versions_table(self):
    self.assertEqual(dw.workflow_elements(self.manifest, "conn-1", "FY27 forecast", FIGURES),
                     PRE_CHANGE_ELEMENTS)   # captured from main before editing
```

- [ ] **Step 2: Run, expect failures** (`version_drivers` not defined)

Run: `.venv/bin/python -m unittest tests.test_demo_workflows -v`

- [ ] **Step 3: Implement** `DriverColumn`, `version_drivers`, `INPUTS_FROZEN` and the `drivers=` path. In `build_wb_spec`, pass `drivers=demo_workflows.version_drivers(P)` only when `editable and workflows is not None and approved_inputs is None`. Add a disposition note to the Snapshot workflow: "The source froze the Forecast grid as values; the migration freezes the inputs and rebuilds the version, so a restated closed month changes a rebuilt version."
- [ ] **Step 4: Run** the module, then `tests.test_builders` and `tests.test_forecast_app`. Expected: PASS.
- [ ] **Step 5: Commit** `feat(excel-to-sigma): Snapshot freezes the forecast's driver values`

---

### Task 3: One function emits the forecast stages

Refactor only: the live specs must not change.

**Files:**
- Modify: `skills/excel-to-sigma/scripts/build-forecast-model.py` (extract from `build_wb_spec` lines ~853–915 and `apply_boundary`)
- Create: `skills/excel-to-sigma/tests/fixtures/forecast-golden/` (`plain-what-if.json`, `app-what-if.json`, `app-read-only.json`, `app-pct.json`)
- Test: `skills/excel-to-sigma/tests/test_forecast_stages.py`

**Interfaces:**
- Produces:

```python
@dataclass(frozen=True)
class StageInputs:
    rate: str       # a line's base rate before the scenario spread, in the stage's row context
    manual: str     # a manual line's value for the row's month
    scenario: str   # the active scenario name ("" when the plan has no spreads)
    start_key: str  # month key of the first forecast month
    source: dict    # the stage-1 element source (the live join, or the version's spine source)
    prefix: str     # the column prefix stage-1 formulas read through ("FJ", "Forecast Spine", …)

def forecast_stages(P, inputs: StageInputs, *, ids: dict[str, str], names: dict[str, str],
                    actuals_name: str) -> list[dict]
```

  - `ids` / `names` map the role keys `stage1`, `ref`, `final` to element ids / names.
  - The function returns `[stage1]`, or `[stage1, ref, final]` with `pct_of_total` lines.
  - The boundary columns (Seed, Window, N, Amount) are applied inside it.
  - `live_stage_inputs(P, editable, spreads) -> StageInputs` reproduces today's formulas exactly.

- [ ] **Step 1: Capture goldens from `main`'s code before editing.** Use a script in the test's `setUpClass` that builds four specs and writes them to the fixture directory:
  - the plain what-if build (demo plan),
  - `--workflows` what-if,
  - `--workflows` read-only,
  - `--workflows` with a `pct_of_total` line. Use `test_forecast_general.py`'s mutation of `plan["lines"][0]` to `{"kind": "pct_of_total", "of": "COGS", "rate": 0.1}`.

  Commit the goldens on their own first.
- [ ] **Step 2: Write `test_live_specs_unchanged`**: each of the four builds equals its golden, compared with `json.dumps(..., sort_keys=True)`.
- [ ] **Step 3: Extract `forecast_stages`, `StageInputs`, `live_stage_inputs`.** `build_wb_spec` calls them. No formula text changes.
- [ ] **Step 4: Run** `tests.test_forecast_stages`, `tests.test_builders`, `tests.test_forecast_app`. Expected: PASS, goldens unchanged.
- [ ] **Step 5: Commit** `refactor(excel-to-sigma): one function emits the forecast stages`

---

### Task 4: Rebuild the picked version

**Files:**
- Modify: `skills/excel-to-sigma/scripts/build-forecast-model.py` (version elements, `--workflows` path only)
- Modify: `skills/excel-to-sigma/tests/planning_app_eval.py` (support the Task 1 pull function if it is `SumIf`)
- Test: `skills/excel-to-sigma/tests/test_forecast_app.py` (new `VersionRebuildTests`)

**Interfaces:**
- Consumes: `forecast_stages`, `StageInputs` (Task 3); `version_drivers`, `INPUTS_FROZEN` (Task 2).
- Produces:
  - **Elements**, on the hidden Model page:
    - `versionPick` ("Version Pick"): Versions filtered to `Inputs Frozen = true` and `Version = [Compare To]`. One row.
    - `versionActuals` ("Version Actuals"): the actuals before the version's start key.
    - `versionForecast` ("Version Forecast"), plus `versionForecastRef` and `versionForecastFinal` when there are pct lines.
    - `versionCombined` ("Version Combined"): the union shaped like `pnlCombined`.
  - **Control:** `CompareTo` (controlId), a list control over `versionPick`'s source Versions, frozen rows only, defaulting to the highest `Log Order`.
  - **`version_stage_inputs(P, drivers) -> StageInputs`:**
    - `rate` = `Switch([<prefix>/Line Item], "<line>", Max([Version Pick/<line> rate]), …)`.
    - `manual` = the same, keyed on line and month number.
    - `scenario` = `Max([Version Pick/Scenario])`.
    - `start_key` = the month key of `Max([Version Pick/Forecast Start Month])`.

- [ ] **Step 1: Write the failing tests.** Run them through `planning_app_eval`, using the demo plan with landed actuals as in the existing `test_forecast_app` setup.

```python
def test_rebuild_of_a_fresh_snapshot_equals_today_cell_by_cell(self):
    # live inputs = version inputs, same scenario/start
    today = grid(evaluate("pnlCombined", self.state)); version = grid(evaluate("versionCombined", self.state_with_v1))
    self.assertEqual(today, version)                     # (line, month) -> amount, to the cent

def test_rebuild_uses_the_versions_scenario_not_the_switch(self):  # v1 saved Upside, live control Base
    self.assertEqual(fy(evaluate("versionCombined", s)), FY_UPSIDE)  # FY_UPSIDE = live FY with control Upside

def test_rebuild_keeps_the_versions_start_month_after_roll_forward(self):
    # v1 start Jul-27, Parameters moved to Aug-27, July actuals landed
    self.assertEqual(period(evaluate("versionCombined", s), "North America Revenue", "Jul-27"), "Forecast")
    self.assertEqual(period(evaluate("pnlCombined", s), "North America Revenue", "Jul-27"), "Actual")

def test_rebuild_runs_the_pct_second_pass(self):   # app-pct plan from Task 3
    self.assertEqual(grid(evaluate("versionCombined", s_pct)), grid(evaluate("pnlCombined", s_pct)))
```

- [ ] **Step 2: Run, expect failures** (unknown element `versionCombined`)
- [ ] **Step 3: Implement `version_stage_inputs` and the version elements** via `forecast_stages(..., ids={"stage1": "versionForecast", ...}, actuals_name="Version Actuals")`. Only the `--workflows` editable path emits them.
- [ ] **Step 4: Run** `tests.test_forecast_app` and `tests.test_forecast_stages`. Expected: PASS. Regenerate only the `app-*` goldens, and say so in the commit message.
- [ ] **Step 5: Commit** `feat(excel-to-sigma): rebuild the picked version from its frozen inputs`

---

### Task 5: Compare section, agent source, docs

**Files:**
- Modify: `skills/excel-to-sigma/scripts/lib/forecast_app.py` (Versions & approvals page `pgVersions`, layout)
- Modify: `skills/excel-to-sigma/scripts/lib/agent_enhance.py` (`GROUNDING`, instructions)
- Modify: `skills/excel-to-sigma/SKILL.md` (planning-app section, smoke pass step 3), `skills/excel-to-sigma/refs/agent-enhance.md`
- Test: `skills/excel-to-sigma/tests/test_forecast_app.py`, `skills/excel-to-sigma/tests/test_agent_enhance.py`

**Interfaces:**
- Consumes: `CompareTo`, `versionCombined`, `pnlCombined` (Task 4).
- Produces (only when the build has driver columns, i.e. editable / what-if):
  - **Elements on `pgVersions`, above the logs:**
    - `compareFacts` (hidden): a union of `pnlCombined` (Which = "Today") and `versionCombined` (Which = "Version").
    - `compareHead`: a text element with the Compare To control and the note `Rebuilt from the saved inputs with actuals as they are now`.
    - `compareKpis`: three text cards, FY revenue / gross margin / EBITDA, each showing Today, Version and Change. They use the same section and derived-row formulas as the live KPIs, scoped by Which.
    - `compareLines`: each line item's FY Today, Version and Change, sorted by absolute change, descending.
    - `compareGrid`: a pivot of line items by months, with the `CompareShow` segmented control (Today / Version / Change).
  - **The Versions log view** shows `Numbers not recorded in the source` in the figure cells of rows without `Inputs Frozen`.
  - **Agent:** `GROUNDING` gains `"compareFacts"`.

- [ ] **Step 1: Write the failing tests**

```python
def test_compare_section_sits_above_the_logs_on_the_versions_page(self):
    order = element_ids_on_page(self.spec, "pgVersions")
    self.assertLess(order.index("compareGrid"), order.index("versionsView"))

def test_compare_kpis_match_the_live_kpis_when_version_equals_today(self): ...  # evaluator, to the cent

def test_no_frozen_version_shows_empty_state(self):   # only migrated rows
    self.assertIsNone(evaluate_text("compareKpis", s_migrated_only).value)
    self.assertIn("No saved versions to compare yet", evaluate_text("compareHead", s_migrated_only).body)

def test_migrated_rows_say_numbers_not_recorded(self): ...

def test_agent_reads_compare_facts(self):
    self.assertIn("compareFacts", [d["elementId"] for d in frag["agents"][0]["dataSources"]])
```

- [ ] **Step 2: Run, expect failures**
- [ ] **Step 3: Implement** the elements and layout, following `forecast_app`'s card and band style. Update `SKILL.md`'s smoke pass step 3: "Snapshot: one Versions row with every driver column and Inputs Frozen; pick it in Compare to; the compare grid's Change is 0 for every cell and the Version KPIs equal the live ones." Update `refs/agent-enhance.md`.
- [ ] **Step 4: Run the full suite** `.venv/bin/python -m unittest discover -s tests`. Expected: OK, with the same skip count as `main` (17).
- [ ] **Step 5: Commit** `feat(excel-to-sigma): compare a saved version to today`

---

### Task 6: Live end-to-end on a staging build

**Files:** none in the repo, unless the run finds a bug. In that case: fix it, add a test, then repeat this task.

- [ ] **Step 1: Build the demo `.xlsm`** with `--workflows --approvers tj+1@sigmacomputing.com` into a staging folder through the router's `stagingAlternative` commands. **User handoff:** TJ types APPROVE at his own terminal.
- [ ] **Step 2: Run the setup program unchanged**, as one `build_workbook` program.
- [ ] **Step 3: Snapshot in the browser pane, then pick it in Compare to.** Query `compareFacts`. Expected: every Change is 0, and the Version FY revenue, gross margin and EBITDA equal the stored FY figures to the cent.
- [ ] **Step 4: Change one growth rate and switch to Upside.** Expected: Today moves and Version doesn't. Then Roll forward. Expected: the version's first forecast month still reads as Forecast.
- [ ] **Step 5: Clean up.** Screenshot the Versions & approvals page for TJ, then move the staging workbook and data model to trash. Run `superpowers:requesting-code-review` (Opus, whole branch), then open the PR.
