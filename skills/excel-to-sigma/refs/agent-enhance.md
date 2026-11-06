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
  input table, `rateInput`, selected by `single-row` primaryKeys on `LINE_ITEM`, new `RATE`)
  is added when that table exists; `column-match` cannot bind an agent input.
- Only `RECALC` / `FORMAT_COSMETIC` / `EXTERNAL_DATA_REFRESH` -> no recommendation and no
  tools. The skill never invents a write tool.
- No macros (plain `.xlsx`) -> no recommendation.

## What gets emitted

`document.agents[0]` (`id fpa-agent`, `@dataSource` / `@tool` instructions, `dataSources` =
forecast + Parameters + Versions + Approvals, one `action` tool per workflow) and a
`chat` element placed in a new full-width band at the bottom of the page that owns
`forecast`. Tool steps are copied from the workbook's own buttons (`btnRoll`,
`btnSnapshot`, `btnSubmit`), so a button and its agent tool cannot drift. `--merge-into`
does this reuse automatically.

## Gotchas (live-proven)

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
