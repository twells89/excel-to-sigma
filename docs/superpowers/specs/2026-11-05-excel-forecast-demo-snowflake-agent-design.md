# Excel → Sigma demo: a macro-heavy FP&A forecast, landed on Snowflake, with a Sigma Agent

Date: 2026-11-05 · Owner: TJ Wells · Status: draft for approval

## Intent and the why

A LinkedIn video, narrated by TJ over the screen, showing the excel-to-sigma
skill migrate a realistic, macro-heavy finance spreadsheet into Sigma, on
camera, end to end.

The point it makes, without naming any other product: the answer to "spreadsheets
are the #1 BI tool" is not to pipe governed data *into* a file. That still
leaves the logic, assumptions and macros in a file on a laptop, emailed around,
forked into `v7_FINAL`. Governing the data that flows in does nothing for what
happens to it afterwards. The Sigma answer: **don't bring the warehouse to the
spreadsheet, bring the spreadsheet to the warehouse.** The same grid, formulas
and workflows, running on the CDW in a governed workspace, with one copy,
permissions and history. Then AI that acts *inside* that governed workspace,
instead of AI that helps you build a better ungoverned file.

The current positioning trigger is a public post about a CDW vendor's AI
assistant coming to Excel and Sheets. The video must never name or show that
product, vendor feature, or post.

## Goals

1. A believable `.xlsm` FP&A rolling forecast that a finance viewer recognises
   immediately (macro buttons, a version-named file, an "email the CFO" step).
2. The excel-to-sigma skill migrates it for real: inventory → macro disposition
   gate → plan → land data in Snowflake → data model → workbook with editable
   assumptions → macros as native workflows → layout → parity.
3. Two skill enhancements, shipped as PR(s) on `twells89/excel-to-sigma`:
   - **Snowflake landing** — an optional step that uses the Snowflake CLI
     (`snow`) to land the file's data tabs as real Snowflake tables on the
     write-enabled `ymb68310` connection (`cb2f5180`), so the Sigma build
     reads governed tables, not pasted cells.
   - **Enhance: Sigma Agent recommendation** — a final, optional phase that reads
     the macro inventory, recommends a workbook agent, and on approval authors
     `document.agents` + a chat element whose action tools call the same
     workflows the macros became (with approval before any write).
4. A finished Sigma workbook that is sleek and makes visual sense to a finance
   viewer, approved by TJ before recording.
5. A ~3–4 min, sped-up, edited cut of a real on-camera run, in the same style as
   TJ's previous videos (`~/sigma-git-e2e-video`): lower-thirds from the Sigma
   lower-thirds kit, styled terminal shots, a silent preview, a scratch-voice
   guide copy, and a timed `TALK_TRACK.md` TJ reads in Screen Studio with his
   camera bottom-right.

## Non-goals

- A general-purpose hardening of every excel-to-sigma path. Only the two
  enhancements above, plus fixes the demo run genuinely exposes.
- Emailing anyone for real. "Submit for approval" changes status and logs it.
- Executing VBA. The skill never runs macros; it classifies and rebuilds them.
- Production customer data. Everything is synthetic.

## Anti-goals (a failure even if it technically works)

- **Any visible hack or workaround.** No Snowflake views over Sigma writeback
  tables, no warehouse-view detours, no hand-edited JSON on screen, no "ignore
  this error", no retry loops shown. Everything on screen is the clean, native
  path. If a step only works through a workaround, redesign the build so the
  step isn't needed.
- Naming, showing, or obviously parodying the competing product or post.
- A read-only port: the assumptions must be editable and the workflows must work.
- A generic-looking dashboard. It should look like a finance team's planning app.

## Constraints and environment facts (verified 2026-11-05)

- Sigma org: `https://app.sigmacomputing.com/tj-wells-1989`, API host
  `aws-api.sigmacomputing.com`, Sigma CLI profile **`tj-wells-1989 (1)`**
  (OAuth, signs in fine). Use `SIGMA_CLI_PROFILE='tj-wells-1989 (1)'`.
- Warehouse (decided by TJ): Sigma connection **`ymb68310`**, id
  `cb2f5180-641f-47bd-8efa-da9d590d855a`, Snowflake account `ymb68310`,
  writeback enabled; the skill's verified write connection. Do not use the
  per-user `OAuth Snowflake` connection (`bc0319f8…`).
- Sigma MCP is connected (`sigma_aws_us_west_2_production` plugin); confirm in
  the first session it is the same org before relying on it.
- Workbook agents are enabled in the org (`sigma api workbook-agents list`
  returns agents). Prior art in-org: "Planning copilot" (workbook
  `85e92ad3-346a-41e5-89e5-89b5849f4953`) sets drivers, switches scenarios,
  navigates, saves/submits with approval. Read its spec before authoring ours.
- **No warehouse-view endpoint exists** in the Sigma CLI (all 37 prefixes
  checked); the skill's own notes list it as UI-only. So the architecture must
  not need one (see below).
- Snowflake CLI `snow` 3.16.0 is installed but currently crashes
  (`ModuleNotFoundError: _cffi_backend`). Fix: `brew reinstall snowflake-cli`
  (or repair the cffi dependency). Connections `default`, `tj` (key-pair
  `~/.snowflake/tj_rsa_key.p8`), `looker_svc` exist in
  `~/.snowflake/config.toml`. Confirm which role/database/schema the
  `cb2f5180` connection reads and land the `snow` tables there, so they are
  visible to Sigma without new grants. If they aren't visible, stop and ask TJ;
  granting access is TJ's call.
- Skill source of truth: `twells89/excel-to-sigma` `origin/main` (870e4e3).
  Work happens on branch `feat/snowflake-landing-agent-enhance` in worktree
  `~/excel-to-sigma-agent-demo`. Follow that repo's `AGENTS.md` and existing
  gates (macro disposition gate, staging approval, acceptance evidence). Never
  renumber its phases; add new ones as sub-phases or a trailing phase.
- Org rule: no non-`@sigmacomputing` package versions released in the last
  3 days.
- Excel for Mac is installed. Do **not** change Excel's Trust Center / "trust
  access to the VBA project object model" setting to script VBA import. If no
  other path produces the `.xlsm` with its VBA project, TJ does a one-time
  import of the `.bas` modules (Developer → Visual Basic → File → Import) and
  saves as `.xlsm` — ~2 minutes, off camera.

## The source spreadsheet (decided: FP&A rolling forecast)

Fictional company (builder's call on the name; must not be a real company or a
well-known vendor sample like Northwind/Contoso). File named like a real one,
e.g. `<Company>_FY27_Rolling_Forecast_v7_FINAL.xlsm`.

Tabs (the builder may refine, keeping the spirit):

| Tab | What it is | Becomes in Sigma |
|---|---|---|
| Instructions | "Click Refresh Actuals first, then…" | (cut — the workbook is the instructions) |
| Actuals | pasted export: month × region × product line, revenue/units/COGS, 18–24 months | **Snowflake table** |
| Plan | annual budget by month × region × product line | **Snowflake table** |
| Dimensions | region, product line lists | **Snowflake tables** / DM relationships |
| Assumptions | drivers by scenario × region (× product line where it helps): growth %, price change %, churn %, and a scenario selector (data-validation dropdown) | **input table** (editable) + scenario **control** |
| Forecast | 12-month forward forecast, cross-sheet formulas (SUMIFS, INDEX/MATCH, IF on scenario) | workbook formulas over DM + input table |
| Variance | forecast vs plan vs prior year, with a chart | KPIs + charts |
| Versions | snapshot log written by a macro | **input table** written by an action |
| Approvals | status log written by a macro | **input table** written by an action |

Macros (as `.bas` modules, classified by the skill's existing taxonomy so the
inventory moment on camera shows several routes):

| Macro | Class | Sigma target |
|---|---|---|
| `RefreshActuals_Click` (ADODB/QueryTable style) | EXTERNAL_DATA_REFRESH | gone — data is live on Snowflake |
| `SetScenario` (writes the selector, `Select Case`) | WHATIF_SCENARIO | scenario control |
| `RollForwardMonth_Click` (moves the actuals/forecast boundary) | ACTION / TRANSFORM | native action (builder decides exact mechanics) |
| `SnapshotVersion_Click` (copy-paste values into Versions) | ACTION_WRITEBACK | insert-rows action into the Versions input table |
| `SubmitForApproval_Click` (locks sheet, Outlook email to CFO) | ACTION_WRITEBACK + email | status change + Approvals log row; lock = edit permissions |
| `RecalcAll`, `FormatReport` | RECALC / FORMAT_COSMETIC | dropped / styling (shown as "auto-handled") |

## Target architecture (decided)

- **Snowflake tables** (landed by the new `snow` step): actuals, plan,
  dimensions. A **Sigma data model** sits on them as the governed layer.
- **Input tables** hold only what people edit or what workflows write:
  assumptions, versions, approvals. They live in the workbook; the forecast
  reads the assumptions input table directly in the workbook. Nothing routes
  input tables back through a data model, so no warehouse view is ever needed.
- **Macros → native workflows** in the workbook (buttons + actions). The agent
  calls those same actions.
- **Sigma Agent** (Enhance phase): an Act-profile agent with action tools for
  roll forward / snapshot / submit and set-scenario / set-driver, plus analyze
  Q&A over the forecast and variance; `requiresApproval` on every write; a
  chat element on the main page or a side panel.

## The Sigma workbook (look and feel)

Builder's call on details, within these rules: a finance planning app, light
theme, Sigma brand-aligned, not a generic dashboard. Rough page plan:

1. **Forecast** — header with scenario control and status (Draft/Submitted),
   KPI strip (forecast revenue, vs plan, vs prior year, gross margin),
   actual+forecast trend chart, variance by region, and the workflow buttons
   (Roll forward, Snapshot, Submit for approval).
2. **Assumptions** — the editable drivers grid by scenario × region, with the
   forecast reacting.
3. **Versions & approvals** — the logs the macros used to write, now governed
   rows with who/when.
4. Agent chat available on the Forecast page (panel or element).

Use the repo's sigma-authoring skills (`sigma-workbooks`, `sigma-data-models`)
and the skill's existing presentation/layout phase; layout last.

## The video (decided parts)

- Real run, recorded once rehearsal is approved, into a fresh Sigma folder and
  fresh Snowflake schema so nothing on camera is pre-baked.
- Edited, sped-up cut, ~3–4 minutes. Lower-thirds per phase, e.g. *Inventory →
  Macro gate → Land in Snowflake → Build → Workflows → Agent*.
- Open on the Excel file (buttons, tabs, the `v7_FINAL` name). Close on the
  agent doing a real action with approval, then the governed workbook.
- Reuse `~/sigma-git-e2e-video` pipeline (Playwright scene recorder
  `rec.mjs`, `overlay.mjs`, `shellify.py`/`shell_render.mjs` for terminal
  shots, `assemble.py`, `vo.py`/`talk_track.py`), copied into a new project
  dir `~/sigma-excel-migration-video`.
- Deliverables: silent preview MP4, guide MP4 with scratch voice, and
  `TALK_TRACK.md` with timed lines (same format as the git-e2e one).

## Sequence and approval gates

1. Fix `snow`; confirm the Snowflake role/db/schema the Sigma connection reads.
2. Build the source `.xlsm` (+ `.bas`). TJ glances at it.
3. Implement the two skill enhancements on the branch, with tests per the repo's
   conventions; open PR(s).
4. **Rehearsal run** of the full migration into a rehearsal folder/schema.
5. **TJ approves the built workbook** (the hard gate before recording).
6. Draft the talk track + scene list; TJ OKs.
7. Clean on-camera run + capture; assemble preview, guide, talk track.

## Calls made after the builder check (Claude's calls; TJ to check)

- **Feasibility spike first.** Before the full build, prove in a scratch
  workbook that an assumptions input table drives forecast formulas over
  data-model rows (scenario × region lookup), with no warehouse view. If it
  doesn't work, stop and bring the redesign to TJ rather than shipping a
  workaround.
- **Roll Forward = moving a boundary, not rewriting data.** A one-row
  Settings input table holds the forecast start month. The Roll Forward
  action advances it, and the forecast and actuals split on it. Actuals in
  Snowflake are never written by Sigma. No stored procedures, which the skill
  still lists as UI-finish.
- **Submit for approval = status plus a log row,** not a permissions change.
  Status shows Submitted in the header, and an Approvals row records who and
  when. Edit rights stay as Sigma's normal sharing. The talk track can say so.
- **The agent's moment on camera:** a plain-English request such as *"Model a
  downside where West growth drops 2 points, snapshot it, and submit it for
  approval."* The agent sets the driver, shows the impact, and asks approval
  before each write. Takes can be retried; the edit shows one clean take.
- **The VBA macros only need to classify correctly.** They don't have to run
  in Excel (the skill never executes them). The opening scene shows the
  buttons without clicking them. The `.xlsm` gets its VBA project through
  whatever path works without changing Excel's trust settings; TJ's one-time
  import is the fallback, scheduled right after the source build.
- **One PR** on `twells89/excel-to-sigma` for both enhancements. The on-camera
  run uses the branch, so recording doesn't wait on a merge.
- **Video tooling:** copy the existing lower-thirds kit, brand assets and
  scratch-voice tooling from `~/sigma-git-e2e-video` as-is. Its node
  dependencies are already installed, so nothing new gets pulled.

## Left to the builder

- Company name, exact numbers, row counts, chart choices, colours within the brand.
- Exact VBA bodies (they must classify into the routes above).
- How Roll Forward is represented natively.
- Agent name, greeting, instructions, and tool list.
- Where the new Snowflake landing step sits in the skill's phase numbering
  (sub-phase, never a renumber) and its script/CLI interface.
- How Excel-window footage is captured for the opening scene (window capture
  vs. TJ recording 15 s himself).
