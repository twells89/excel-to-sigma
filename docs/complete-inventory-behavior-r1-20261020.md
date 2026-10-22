# Inventory PO SQL R1 correction — 2026-10-20

The independent review at work/complete-inventory-behavior-independent-review-20261020.md
found a blocking same-intent retry path. Snowflake COUNT_IF returns NULL when
no row matches. With a persisted row for the same operation and source row
but a changed SKU or quantity, the prior count was one and the match count
was NULL. The SQL IF condition became UNKNOWN, so the procedure could return
RETRY_PRESENT for changed data.

The revised isolated STAGE_PO proposal normalizes the match aggregate with
COALESCE(COUNT_IF(...), 0). With one prior row, changed SKU or quantity now
makes the mismatch guard true. Same payload still reaches RETRY_PRESENT;
no prior row reaches STAGED; duplicate/corrupt rows and a different batch
fingerprint fail. COUNT(*) supplies zero directly for the other aggregate
counts. The argument guard also explicitly rejects NULL source row and
batch fingerprint, avoiding an UNKNOWN validation condition.

The new offline SQL contract test first failed against the reviewed bytes
for both changed-row cases and passed after the correction. An independent
controller test also rejects a changed source SKU for the same staged intent.
These tests inspect the final SQL aggregate shape and apply Snowflake's
documented three-valued COUNT_IF/IF semantics to generated cases. They
do **not** compile or execute Snowflake SQL. The SQL remains create-only,
isolated to SIGMA_WRITE_DB.SIGMA_MIGRATION8_INV_TEST, and unexecuted.

Remaining gates before root provisioning: independent re-review of the exact
SQL bytes; isolated compilation and catalog/caller checks; two-session
same-intent contention, timeout, and rollback proof; changed-SKU and
changed-quantity calls; same payload retry, new intent, partial recovery,
and duplicate-row readback; Sigma typed VARIANT mapping and actual native
binding. No staging receipt or migration-complete claim is produced here.
