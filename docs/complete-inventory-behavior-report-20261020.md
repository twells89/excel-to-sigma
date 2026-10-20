# Inventory source behavior port handoff — 2026-10-20

Status: **offline implementation for review; no live migration or completion claim**.

The exact input is 08 Inventory Replenishment Plan.xlsm (SHA256
461bee7b71f30822ea4d263ef529723d89dc9f11e7365e6434cefe89619c2f7e)
and companion BAS (SHA256
181ff80f68b9924844cea64d3c964f36a720bc2944e3566dbde2395946428ae0).
The workbook was inspected read only as data. VBA was never executed.
The [machine coverage](../skills/excel-to-sigma/refs/inventory-source-behavior-coverage.json)
records all nine exact procedure hashes, actual inputs, guards, affected rows,
helpers, side effects, and remaining evidence gates.

The [pure port](../skills/excel-to-sigma/scripts/lib/inventory_behavior.py) handles:

* CommitForecast: an immutable 72-row snapshot payload with one operation ID,
  timestamp, authenticated invoker, SKU, week start, and the then-current units.
  A same-ID retry can resume only absent source rows after exact payload
  comparison; a new ID creates a separate batch. The target remains an
  ordinary independent Sigma Snapshots input table, with native append and
  exact post-write readback. Concurrent native insert safety remains a live
  controller/action binding gate.
* ApplyDemandAdjustment: one finite percentage parsed before any write,
  applied to every row in source order with Double arithmetic and ties-to-even
  zero-digit rounding. Source has no percentage bounds. Blank numeric cells
  coerce to zero in arithmetic; blank percentage fails.
* ReforecastExpSmoothing: rows 3–73, alpha 0.3, adjacent SKU equality, and
  the **previous updated row**. The first row of each contiguous block stays
  unchanged. Null first-row units remain null; a later arithmetic read coerces
  them to zero.
* RefreshOnHand: the user-approved repair updates SKUs!C by exact SKU,
  preserving Description and every other field. The source
  CopyFromRecordset starts at SKUs!A2 and wrongly overwrites Description
  in column B. The replacement rejects missing, duplicate, and unknown
  synthetic ERP keys before any write.
  The ERP capture must report the same snapshot date as the warehouse's
  current date in that query.
* CreatePurchaseOrders: only Projection rows 4–9 with literal ORDER,
  taking same-row SKU and whole suggested quantity. A selected zero or
  negative whole quantity is retained because the source has no positivity
  guard; nonwhole and unrepresentable quantities fail the typed port.

The [controller](../skills/excel-to-sigma/scripts/lib/inventory_controller.py)
and [offline CLI](../skills/excel-to-sigma/scripts/inventory-actions.py) use
source hash, current document/version, complete query evidence, actual
physical row and column IDs, operation IDs, plan hashes, full before/after
fingerprints, and exact inverse values for updates. They return blocked plans,
not success receipts. A fresh full readback is required to confirm a
component write. Procedure scalar outputs are type checked and cannot alone
confirm staging. No actual input bindings have been compiled into actions.

The [SQL proposal](../skills/excel-to-sigma/refs/inventory-isolated-integrations.sql)
creates only SIGMA_WRITE_DB.SIGMA_MIGRATION8_INV_TEST, six controlled
ERP snapshot SKUs, and synthetic PO staging. Forecast snapshots stay in
the ordinary independent Sigma input table and have no warehouse procedure.
No ERP_PROD reference, broad grant, credential, or live write is included.
The PO procedure starts and ends its own transaction, updates a
dedicated lock row, then checks the operation/payload fingerprint before
inserting. An interrupted intent resumes only missing rows after exact
readback. Standard Snowflake tables do not enforce unique keys; the
transaction/lock behavior and rollback must be tested with two concurrent
sessions before provisioning is accepted. See [Snowflake constraints](https://docs.snowflake.com/en/sql-reference/constraints)
and [transactions](https://docs.snowflake.com/en/sql-reference/transactions).
The SQL has **not been compiled or executed**; actual stored procedure catalog
signature, owner/caller privileges, typed VARIANT mapping through Sigma, and
concurrency/error paths remain review gates. The root task alone provisions.

The [native surface contract](../skills/excel-to-sigma/refs/inventory-native-surface.json)
records Projection as default/live-calculation page, the GoToProjection A3
navigation, and TidyLayout's B:M width 9 / A3:P3 bold formatting. Actual
page/element IDs and render proof are pending. EmailBuyers remains explicitly
unexecuted external delivery.

The [narrow empty log precheck](../skills/excel-to-sigma/scripts/lib/inventory_completion_gate.py)
accepts a zero baseline only with the exact append-log role and schema,
physical IDs/hashes, published insert/persist, same-ID retry, distinct new
intent, restored zero state, role sessions, screenshots, and all production
gates. It returns only review eligibility and does not override the general
migration completion gate.

Offline tests pin all 72 source Demand rows and independent expected
adjustment/smoothing arrays; they also cover ties, null/zero/negative inputs,
bad percent, nonfinite/overflow, wrong/missing/duplicate physical rows,
ERP key mismatch, changed PO selection, partial retry, uncertain procedure
output, exact readback, and the narrow empty-log exception. The tests and
fixture are under skills/excel-to-sigma/tests/. No published buttons,
actual persisted rows, role sessions, screenshots, or restoration evidence
have been captured in this branch.

Verification: 19 focused Inventory tests passed. The broad repository suite
was stopped before verdict at root coordination request; the root task will
run one integrated suite after merging components. The SQL proposal remains
uncompiled and unexecuted.
Six adjacent existing suites also passed (114 tests across custom ports,
native actions, operation outcomes, migration gate, macro behavior, and
forecast snapshot).
