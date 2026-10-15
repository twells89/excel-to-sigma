# Complex-workbook coverage contract

“Any Excel file” is not a technically honest promise. Excel can contain an
arbitrary cell program, VBA, Power Query, external applications, COM/ActiveX,
embedded binaries, and credentials. This skill instead promises:

1. inventory every material workbook feature before build;
2. migrate supported behavior with data, formula, interaction, and render parity;
3. route representable-but-not-yet-automatic behavior to a named Sigma primitive;
4. block completion when no faithful route has been implemented;
5. never silently drop workbook behavior.

Run the hard gate first:

```bash
python scripts/audit-xlsx.py model.xlsx --out migration-audit.json
python scripts/macro-classify.py --json model.xlsm > macro-audit.json
python scripts/audit-openapi.py --out sigma-code-rep-audit.json
```

`audit-xlsx.py` exits `0` only for `AUTO`, `1` for `REVIEW_REQUIRED`, and `2`
for `BLOCKED`. `--allow-review` is for inventory/CI smoke only; it must not be
used to declare a customer migration complete.

## Coverage

| Excel surface | Route | Completion gate |
|---|---|---|
| Formal Tables / typed data | input table or data-model table | row/grain + aggregate parity |
| Common formulas | Sigma formulas / data-model calculations | cell or aggregate parity |
| Unsupported formulas / dynamic arrays | rewritten grain, SQL, or Python | explicit formula manifest + parity |
| Charts | matching Sigma chart kind | encoding/filter/aggregation + screenshot review |
| Pivots | Sigma pivot table | rows/columns/values/sorts/subtotals + query parity |
| Slicers/timelines | controls and target filters | interaction test |
| Data validation | input-column `values`, `valuesFrom`, `range`, or file constraints | invalid/valid input test |
| Conditional formatting | Sigma formatting rules | render review |
| Editable cells | input tables and actions | draft/published permission + writeback test |
| VBA writeback/API/navigation | element actions | runtime side-effect test |
| VBA stored procedure | `call-stored-procedure` action | signature/arguments/transaction review |
| VBA/Python integration | `code` + `code-output` + `run-python-element` | package/runtime/output parity |
| Scheduled/webhook automation | document `automatedActions` where supported | owner/failure/retry review |
| Power Query / connections | governed warehouse, DM, or API source | source and refresh parity |
| External workbook links | land as governed sources | source ownership + parity |
| Protection | Sigma sharing plus RLS/CLS policy where required | security review |
| `.xls` / `.xlsb`, add-ins, ActiveX, embedded binaries | convert/extract or external workflow | blocked until dispositioned |

## Hard completion gates

A migration is not complete until:

- all audit issues have an owner and disposition;
- all workbook creates pass `POST /v2/workbooks` with `dryRun: true`;
- server-assigned data-model ids are read back before workbook wiring;
- editable inputs work in the intended draft/published mode;
- formulas and aggregate outputs meet the agreed parity target;
- actions, Python, procedures, source refresh, and writeback are runtime-tested;
- chart/pivot/control behavior and layout are visually reviewed;
- security and unsupported/external behavior are explicitly signed off.
