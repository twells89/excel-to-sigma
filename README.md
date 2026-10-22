# Excel → Sigma

A coding-agent skill that migrates **Excel (`.xlsx`) planning / budget / forecast
models** to Sigma — preserving the part that matters most about a spreadsheet:
**data entry**.

The repository is a Claude Code plugin marketplace and also contains a standalone,
agent-neutral `SKILL.md` with its own scripts, refs, and Python requirements.

## Install

### Claude Code

In Claude Code, add this repository as a marketplace and install its plugin:

```text
/plugin marketplace add twells89/excel-to-sigma
/plugin install excel-to-sigma@excel-to-sigma
```

Invoke `/excel-to-sigma:excel-to-sigma` or ask Claude to migrate an Excel model.
The plugin contains the same skill directory used by every other agent. If you
are developing a local checkout, `claude --plugin-dir /path/to/excel-to-sigma`
loads it for one session. An installed plugin does not install Python dependencies
or connect to Sigma for you; complete the requirements below before posting.

### Codex, Cursor, Cortex Code, and other coding agents

Clone this repository and point your agent at
[`skills/excel-to-sigma/SKILL.md`](skills/excel-to-sigma/SKILL.md). The root
[`AGENTS.md`](AGENTS.md) routes Excel migration requests to that skill. For a
project or user skill installation, copy the **entire**
`skills/excel-to-sigma/` directory into the agent's skill location, keeping
`SKILL.md`, `refs/`, `scripts/`, and `requirements.txt` together. For example:

```bash
git clone https://github.com/twells89/excel-to-sigma.git
cp -R excel-to-sigma/skills/excel-to-sigma /path/to/agent/skills/
```

In a full clone, read `AGENTS.md`; with a standalone skill installation, open
its `SKILL.md` directly. Run commands from the skill directory so `scripts/`
and `refs/` resolve the same way in either form. In Cortex Code, a clone can
also be registered with `cortex skill add excel-to-sigma/skills/excel-to-sigma`.

## What it does

Excel planning models aren't dashboards — they're data-entry tools. The faithful
Sigma translation keeps forecasters typing numbers, which is what **input tables**
are for. The skill covers:

1. **Discovery** — inventory formal Excel Tables, pivots, charts, and a formula
   census (flagging the untranslatable).
2. **Read/output half** — formal Tables → Sigma data-model elements; the
   income-statement SUMIFS become grouped DM metrics + workbook charts.
3. **Data-entry half (validated)** — the formal Table forecasters edit becomes a
   Sigma **input table**; the data model re-points to read it through a
   **warehouse view**.
4. **Fail-loud complex-workbook audit** — inventories formulas, charts, pivots,
   validation, protection, links/connections, VBA, add-ins, and embedded
   content; unsupported behavior receives an explicit route instead of vanishing.

```
Excel formal Table → Sigma input table → (publish) → SIGDS_ writeback
   → warehouse view → data model (FROM view) → workbook charts/metrics
```

## Status

- ✅ **Input-table data-entry path: validated end-to-end** (2026-06-08) with exact
  parity on a 540-row forecast model.
- ✅ Read/output DM + workbook build: validated (2026-06-03).
- ✅ Excel chart discovery and line/bar/doughnut/scatter migrations validated on
  deterministic sales, expense, and headcount fixtures.
- ✅ Current Sigma `contents` API, dry-run validation, Actions, Python, stored
  procedures, input validation, and automation are OpenAPI-audited in CI.
- ✅ Cell-drawn actuals + driver-growth forecast models: `migrate.py plan` → `draft-forecast-plan.py`
  → `build-forecast-model.py` drafts, verifies against the workbook's cached values, and checks local
  parity; unrecognised shapes fail loudly.
- 🚧 Discovery + Excel-formula translation for other archetypes: spiked, not yet a one-shot converter.

See `skills/excel-to-sigma/SKILL.md` for the full workflow and `QUICKSTART.md` for
a runnable walkthrough on the bundled sample fixture.

## Layout

```
.claude-plugin/plugin.json          Claude Code plugin manifest
.claude-plugin/marketplace.json     Claude Code marketplace entry
AGENTS.md                          agent-neutral request routing
skills/excel-to-sigma/
  SKILL.md                          phased workflow (read first)
  QUICKSTART.md                     runnable end-to-end on the sample fixture
  requirements.txt                  pinned Python dependencies
  refs/
    input-tables.md                 the validated input-table workflow + API/UI split
    excel-translation.md            translation surface, formal-Table rule, converter design
    sigma-build-gotchas.md          DM/workbook spec rules
  scripts/
    make-sample-forecast.py         generate a synthetic .xlsx fixture (no customer data)
    make-chart-examples.py          generate sales/expense/headcount chart fixtures
    xlsx-discover.py                Phase 0 — inventory tables/visuals/formulas
    migrate.py                      router: audit + discover + archetype, prints the next steps as JSON
    audit-xlsx.py                   fail-loud workbook capability/routing manifest
    draft-forecast-plan.py          draft + verify a forecast model plan from a cell-drawn P&L
    audit-openapi.py                current Sigma code-rep drift gate
    xlsx-to-input-csv.py            Phase 1 — extract a Table to a paste-ready CSV
    build-input-table-wb.py         Phase 3 — POST the input-table workbook (API)
    build-dm-on-view.py             Phase 4 — POST the DM reading FROM the view (API)
```

## Requirements

- Sigma client credentials: `SIGMA_BASE_URL`, `SIGMA_CLIENT_ID`, and
  `SIGMA_CLIENT_SECRET` in `~/.sigma-migration/env` for the bundled REST builders.
- A **write-enabled** Sigma connection for the data-entry half (input tables write
  back to a `SIGDS_` schema).
- Python 3 and the pinned dependencies. From the skill directory, run
  `python3 -m venv .venv && .venv/bin/pip install -r requirements.txt`.
- Sigma MCP Code Mode for the documented programmatic input-table seeding and
  query/parity steps, or the documented UI fallback where available. REST scripts
  handle workbook/data-model construction; the skill does not bundle an MCP
  server. The optional `sigma-workbooks` companion skill helps with advanced
  workbook authoring and is located through the agent's skill loader.

## License

MIT
