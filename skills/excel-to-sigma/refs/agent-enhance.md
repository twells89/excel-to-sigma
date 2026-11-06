# Phase 6 — Enhance: recommend a Sigma Agent

After the workbook is built, read what the macros became and offer an agent that can
run those same workflows. This is a recommendation, not a default: nothing is added
until the user says yes.

```bash
.venv/bin/python scripts/macro-classify.py --json book.xlsm > macros.json
.venv/bin/python scripts/macro-actions.py demo-manifest book.xlsm --out workflows.json
.venv/bin/python scripts/recommend-agent.py --macros macros.json --workflows workflows.json
# on approval:
.venv/bin/python scripts/recommend-agent.py --macros macros.json --workflows workflows.json \
    --merge-into wb_spec.json --out wb_with_agent.json     # or --emit fragment.json
```

The printed text is the on-camera recommendation. It names the workflows actually found
and says writes ask for approval first.

## Decision rule

- Any `ACTION_WRITEBACK` / `EXTERNAL_IO` / `WHATIF_SCENARIO` macro that became a workflow ->
  profile `act`; every tool changes the workbook, so every tool is `requiresApproval: true`
  (including the scenario switch). A `set-driver` tool (edit one row of the Growth Rates
  input table, `rateInput`, new `RATE`) is added when that table exists. It picks the row in
  two steps: `set-control-value` puts the agent's `line_item` into a hidden text control
  (`DriverLineItem`, on the hidden model page), then `update-rows` with
  `whichRows {type: formula, formula: "[Line item] = [DriverLineItem]"}`.
- Only `RECALC` / `FORMAT_COSMETIC` / `EXTERNAL_DATA_REFRESH` -> no recommendation and no
  tools. The skill never invents a write tool.
- No macros (plain `.xlsx`) -> no recommendation.

## What gets emitted

`document.agents[0]` (`id fpa-agent`, `@dataSource` / `@tool` instructions, `dataSources` =
forecast + Parameters + Versions + Approvals + Growth rates + `agentFacts`, each only when
the workbook has it, one `action` tool per workflow) and a
`chat` element placed in a new full-width band at the bottom of the page that owns
`forecast`, plus `controlElements` (the set-driver row-picker control) that `--merge-into`
places on the workbook's hidden page; a hand merge of `--emit` output must add them too.

`agentFacts` ("Monthly P&L facts", hidden, planning app only) is the agent's fiscal-year
source: every month's Amount by Fiscal year, Period (Actual / Forecast / Plan), Section,
Line item and Region, read from the same facts as the KPI cards, so it follows the forecast
boundary. The instructions say a fiscal year is its Actual + Forecast rows and its plan is
its Plan rows, and that the `forecast` element holds the forecast months only (live: grounded
in `forecast` alone, the agent called the forecast-window total "FY27 revenue" and said the
workbook had no plan). The greeting suggests "how does revenue compare with plan?" only when
a landed plan (`planLines`) is in the workbook. Tool steps are copied from the workbook's own buttons (`btnRoll`,
`btnSnapshot`, `btnSubmit`), so a button and its agent tool cannot drift. `--merge-into`
does this reuse automatically.

## Gotchas (live-proven)

- An input table's only primary key is its system row ID. A `single-row` selector keyed on a
  business column (`primaryKeys: {LINE_ITEM: ...}`) is rejected at run time with "Primary key:
  ID not in update row definition", even though a dry run accepts it. `column-match` takes only
  a literal value and a row formula cannot read an agent input, so an agent tool that picks a row
  by a business key writes that key into a hidden control first and selects rows with a formula
  that reads the control. `merge_fragment` refuses a `single-row` agent selector on any other
  column, and a `set-control-value` step on a control the workbook does not have. (The rejection
  is live-proven; the control-backed selector must be proven by the rehearsal's `call_agent`
  driver change plus a re-query of the Growth rates table.)

- Never name input-table system columns (`CREATED_BY`, `CREATED_AT`, bare `ID`) in the agent's
  instructions or data sources; agent SQL on them fails with "Ambiguous column reference".
- Agent writes only work on **manual-commit-enabled** input tables. A REST spec POST never
  sets that flag; run the post-create program from `build-forecast-model.py
  --workflows workflows.json --emit-manual-commit-program` once, before the agent's first write.
- `requiresApproval` is a chat-UI behavior. It does not gate the REST `run_agent` call, so
  approval can only be demonstrated in the chat panel.
- After any write, approved or declined, the agent must re-query the affected table and
  report the data, not the tool result.
- No email is ever sent by a tool. The Submit tool appends an Approvals row only.

## Validation

`POST /v2/workbooks` with `dryRun: true` validates the merged spec's agent and chat schema
(a bad tool kind is rejected by name). Live behavior (driver change, Versions row,
Approvals row) is checked after publish by querying those tables.

## Planning-app workbooks

When the workbook has the planning app's scenario switch (a `segmented` control with
`controlId: Scenario`), the set-scenario tool is a `set-control-value` step on that control
(agent input `scenario`) instead of a Parameters update, and the agent reads the active scenario
from the forecast's `Active Scenario` column. The chat is placed inside the reserved
`agentPanel` container on the Forecast page, under its header; the agent carries a static
greeting naming what it can do. `build-forecast-model.py --agent-macros macros.json` merges the
agent in the same run, before the workbook POST; the merge is first run against the local spec,
so one that cannot apply stops the run (`AGENT ERROR`) before the data model is created.
