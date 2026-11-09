#!/usr/bin/env python3
"""DEMO KIT ONLY (not part of the excel-to-sigma migration path): simulate the Jul-27 month-end ERP load.

The source .xlsm's actuals end Jun-27. Before recording, this lands plausible Jul-27 actuals into the
landed actuals table named by --table (DB.SCHEMA.TABLE, required), so pressing Roll forward on camera shows
a real month close: July turns into landed actuals and Aug-Dec re-forecast from it.

Deterministic: one row per Jun-27 (line, region, product) row, grown by the line's plan growth rate
(flat lines carry, the manual line follows its typed Jul-27 value split by June's product mix), times a
seeded +/-1.2% noise. Idempotent: deletes the month's rows first. Dry run unless --execute.
"""
import argparse, csv, json, random, subprocess, sys

ap = argparse.ArgumentParser()
ap.add_argument("--plan", required=True)
ap.add_argument("--actuals-csv", required=True, help="the landed actuals CSV (landing.json csv)")
ap.add_argument("--table", required=True, help="the landed actuals table, DB.SCHEMA.TABLE (land-in-snowflake.py)")
ap.add_argument("--snow-connection", required=True, help="the snow CLI connection name")
ap.add_argument("--from-month", default="Jun-27")
ap.add_argument("--month", default="Jul-27")
ap.add_argument("--posting-date", default="2027-07-01")
ap.add_argument("--seed", type=int, default=20270731)
ap.add_argument("--out", required=True, help="CSV of the generated rows")
ap.add_argument("--execute", action="store_true")
a = ap.parse_args()

plan = json.load(open(a.plan))
lines = {ln["line"]: ln for ln in plan["lines"]}
rows = [r for r in csv.DictReader(open(a.actuals_csv)) if r["MONTH"] == a.from_month]
if not rows:
    sys.exit(f"no {a.from_month} rows in {a.actuals_csv}")
by_line = {}
for r in rows:
    by_line.setdefault(r["LINE_ITEM"], 0.0)
    by_line[r["LINE_ITEM"]] += float(r["AMOUNT"])
rng = random.Random(a.seed)
out = []
for r in rows:
    ln, amt = lines[r["LINE_ITEM"]], float(r["AMOUNT"])
    if ln["kind"] == "manual":
        value = ln["manual"][0] * amt / by_line[r["LINE_ITEM"]]
    elif ln["kind"] == "flat":
        value = amt
    else:
        value = amt * (1 + ln["rate"])
    value = round(value * (1 + rng.uniform(-0.012, 0.012)), 2)
    out.append([a.posting_date, a.month, r["LINE_ITEM"], r["REGION"], r["PRODUCT_LINE"], f"{value:.2f}"])
with open(a.out, "w", newline="") as f:
    w = csv.writer(f); w.writerow(["POSTING_DATE", "MONTH", "LINE_ITEM", "REGION", "PRODUCT_LINE", "AMOUNT"]); w.writerows(out)
q = lambda s: "'" + s.replace("'", "''") + "'"
values = ",\n".join(f"({q(d)}::date, {q(m)}, {q(l)}, {q(rg)}, {q(p)}, {v})" for d, m, l, rg, p, v in out)
sql = (f"BEGIN; DELETE FROM {a.table} WHERE MONTH = {q(a.month)}; "
       f"INSERT INTO {a.table} (POSTING_DATE, MONTH, LINE_ITEM, REGION, PRODUCT_LINE, AMOUNT) VALUES\n{values}; COMMIT;")
rev = sum(float(v) for *_x, l, _r, _p, v in [(o[0], o[1], o[2], o[3], o[4], o[5]) for o in out] if lines[l]["section"] == "Revenue")
print(f"{len(out)} {a.month} rows -> {a.out}; revenue {rev:,.2f}")
if not a.execute:
    print("DRY RUN; re-run with --execute to land them in", a.table)
    sys.exit(0)
subprocess.run(["snow", "sql", "--connection", a.snow_connection, "-q", sql], check=True)
subprocess.run(["snow", "sql", "--connection", a.snow_connection, "--format", "json", "-q",
                f"SELECT MONTH, COUNT(*) AS N, SUM(AMOUNT) AS AMT FROM {a.table} WHERE MONTH IN ({q(a.from_month)}, {q(a.month)}) GROUP BY 1 ORDER BY 1"], check=True)
