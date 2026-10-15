# Macros (VBA) — detect, classify, route, gate

`.xlsm`/`.xlsb`/`.xls` carry VBA. **Macros are not one thing** — like cells, you
classify them by *intent* and route each class. Actions, Python code elements,
stored-procedure calls, API connectors, and automated actions are now
code-representable. Representable does not mean automatically safe. The
non-negotiable rule:

> **Never silent.** Every macro is surfaced with a disposition. A dropped
> "Commit Forecast" button the user assumes still works is the FP&A equivalent of
> silently dropping row-level security. Macros are a **STOP/flag gate**, not a
> quiet delete.

## Extraction

VBA lives in `xl/vbaProject.bin` (an OLE compound stream) — `openpyxl` only sees
*that it exists*, not the code. Use **`olevba`** (oletools) to pull module source:

```bash
pip install oletools          # stable; not in the org's last-3-days window
python scripts/macro-classify.py path/to/model.xlsm          # human report
python scripts/macro-classify.py --json path/to/model.xlsm   # machine contract
```

`macro-classify.py` accepts a binary office file (extracts via olevba) **or** a
raw VBA module (`.bas/.cls/.frm/.vba/.txt`) for testing. It strips comments before
classifying (comments must never drive a classification — they lie).

## The taxonomy → Sigma mapping

| Class | VBA signals | Sigma target | Status |
|---|---|---|---|
| **RECALC** | `.Calculate(Full)`, `ScreenUpdating` | none after formula parity | ✅ AUTO (drop) |
| **EXTERNAL_DATA_REFRESH** | `RefreshAll`, QueryTables, ADODB, SQL/CommandText | governed warehouse/API source, DM, or stored procedure | 🔌 DATA_REVIEW |
| **TRANSFORM_ROLLUP** | `For…To` loop + `Cells(..).Value =`, `WorksheetFunction`, in-code SUMIFS | proven equivalent formulas/DM; otherwise Python | 🐍 PYTHON_REVIEW |
| **NAVIGATION** | `.Activate`, `.Select`, `.Goto`, `.Visible =` | page navigation / visibility | ✅ AUTO (map or drop) |
| **FORMAT_COSMETIC** | `.Font`, `.Interior`, `.Borders`, `AutoFit`, `.NumberFormat` | conditional formatting / styling | 🐍 PYTHON_REVIEW / render review |
| **WHATIF_SCENARIO** | `GoalSeek`, `Solver`, writes into assumption cells, `Select Case scenario` | controls / parameters / editable assumptions table | 🎛 CONTROL |
| **ACTION_WRITEBACK** | `PasteSpecial xlPasteValues`, `ListRows.Add`, append patterns/name verbs | `insert-rows`/`update-rows` input-table Action | ⚡ ACTION |
| **API_CALL** | `MSXML2`, `WinHttp`, HTTP verbs | connector-backed `call-api` Action | ⚡ ACTION + runtime review |
| **STORED_PROCEDURE** | `adCmdStoredProc`, CommandText `EXEC`/`CALL` | `call-stored-procedure` Action | 🗄 STORED_PROC |
| **PYTHON_AUTOMATION** | `RunPython`, xlwings, PyXLL | `code` element + `run-python-element` Action | 🐍 PYTHON_REVIEW |
| **EXTERNAL_IO** | Outlook/email, `Shell`, PDF/file I/O | scheduled delivery / external workflow | 🚩 FLAG |
| **OPAQUE** | bespoke math writing cells (`Sin/Cos/Rnd/…`), `Do While`, `GoTo` | — | 🚩 FLAG (human review) |
| **event-handler** | name `Workbook_*` / `Worksheet_*` (e.g. `Workbook_Open`) | decompose into its called actions; the auto-run wrapper itself drops | (follows contents) |

**Precedence** (when a macro matches several) favors external I/O, explicit
Python/procedure/API/data-source signatures, then writeback/control/transform
routes. Bespoke math (`Sin/Rnd/…`) overrides weak/AUTO matches to **OPAQUE**.
Action verbs are matched on the **procedure name** only, so a called helper like
`PostToGrid` does not trip “Post”; generic VBA `Call Helper` is never treated as
a stored procedure.

## Status meanings

- **AUTO** — reproduced by the data/formula conversion, or safely dropped after parity.
- **CONTROL** — becomes controls / parameters / an editable input table (the driver story in `refs/forecast-recipes.md`).
- **ACTION** — becomes an element Action; test arguments, permissions, side effects, retries, and failure UX.
- **STORED_PROC** — candidate for `call-stored-procedure`; confirm connection, path, signature, typed arguments, and transaction semantics.
- **PYTHON_REVIEW** — candidate for formulas/DM or a Python element; confirm packages, credentials, outputs, and parity.
- **DATA_REVIEW** — refresh/query logic must be rebuilt as a governed source and data parity-tested.
- **FLAG** — must be reviewed by a human (external I/O or opaque logic). Surface before build; never auto-port.

## Actions, Python, and procedure wiring

The workbook code representation now supports element Actions including
`call-api`, `call-stored-procedure`, input-table insert/update/delete,
`run-python-element`, navigation, controls, export, and refresh. A build step
consumes the classifier JSON and emits only reviewed routes:

```
CommitForecast_Click  → Action(on button): insert current forecast rows → "Snapshots" input table (+ timestamp)
SaveScenario_Click    → Action(on button): insert current assumption rates → "Scenarios" input table
ApplyScenario(name)   → Action(on select): set assumption controls to the chosen scenario's values
```

The classifier output is the routing contract, not proof that the emitted
behavior is correct. Every non-AUTO route remains gated on a live side-effect or
parity test.

## Worked example (`scripts/Sample Macros.bas` — the FP&A model's VBA)

The sample intentionally exercises automatic, control, action, data-review,
Python-review, and flag routes:

| Macro | Class | Status |
|---|---|---|
| `Workbook_Open` | RECALC + event | ✅ AUTO |
| `RefreshActuals_Click` | EXTERNAL_DATA_REFRESH | 🔌 DATA_REVIEW |
| `RebuildPnL` | TRANSFORM_ROLLUP | 🐍 PYTHON_REVIEW |
| `GoToForecast_Click` | NAVIGATION | ✅ AUTO |
| `FormatStatement` | FORMAT_COSMETIC | 🐍 PYTHON_REVIEW |
| `ApplyScenario` | WHATIF_SCENARIO | 🎛 CONTROL |
| `SolveForTarget` (GoalSeek) | WHATIF_SCENARIO | 🎛 CONTROL |
| `CommitForecast_Click` | ACTION_WRITEBACK | ⏳ ACTION |
| `SaveScenario_Click` | ACTION_WRITEBACK | ⏳ ACTION |
| `EmailReport_Click` | EXTERNAL_IO | 🚩 FLAG |
| `AllocateOverheadByHeadcount` | OPAQUE | 🚩 FLAG |

Extraction validated end-to-end through olevba on a real `.xlsm`; classification
validated on the full taxonomy above (2026-06-17).
