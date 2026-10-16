# Corrected Forecast snapshot component

`lib/forecast_snapshot.py` implements the explicitly reviewed replacement for
`CommitForecast_Click`: all 18 financial lines in `Monthly Model!I:N`
(July–December), 108 numbers, with one separate frozen timestamp. It does not
claim literal parity with the broken `P&L!H7:M33` copy/timestamp overwrite.

`build_snapshot_definition(source, plan_hash=..., procedure_hash=..., decision=...)`
reads formula-bearing source bytes (`data_only=False`). The decision must name
an approved 18-line/108-value July–December scope and separate timestamp policy.
The definition contains financial source identities and 96 editable input field
identities; it contains no cached formula results or fallback values.

`render_snapshot_candidate(definition, current, bindings, phase="prepare")`
returns a disabled Code Mode constructor and exact controller precondition.
Bindings contain four roles (`primary`, `finance`, `operations`, `snapshots`),
actual assigned element/page IDs, complete SDK schemas, and the twelve stored
primary row IDs. Its single Prepare insert freezes 108 live numbers and 96
editable values, then carries the insert's native returned operation ID into a
hidden control. Timestamp metadata is captured once with `now()`.

After actual assigned-control readback, use `phase="confirm"`, the actual
`operation_control_id`, and the unchanged `prepared_contract` returned by
Prepare. Confirm checks the original input identities/96 values, complete
source/plan/decision metadata, and every existing snapshot key before effects.
Its 108 missing-key branches read only frozen operation columns. Existing
conflicts, duplicate keys, and foreign keys block the entire action. Both
constructors remain disabled and remove custom success toasts. Neither writes
primary inputs or automatically marks an operation successful.

`reconcile_snapshot_rows` is a pure component planner; its unit data is not live
evidence. `plan_snapshot_reconciliation` additionally requires complete actual
`origin: sigma-query` captures bound to the same document/version/source/plan,
actual table IDs and native row-ID column IDs. Each operation and stored snapshot
row includes `nativeRowId`. Same-operation retries reuse that exact ID and frozen
payload. A deliberate new commit requires a new stored operation ID; old rows
remain intact. Reconciliation returns only missing rows and component status.

The parent controller must freshly verify the returned public source hash and
actual rows before submitting a constructor, independently review backend type
admission, and capture resulting assigned state. Stop on any partial constructor
failure; do not replay it. A successful local compile is not backend admission.
Native actions have no multi-effect transaction or version CAS. Actual published
execution, recovery, roles, concurrency, persisted 108-value readback, outcome
reconciliation, rollback, and full workflow acceptance remain required. A new
intent must not clear a pending operation; clear/recover only after actual
reconciliation or an explicitly reviewed abandonment.
