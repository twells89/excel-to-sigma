#!/usr/bin/env python3
"""DEMO KIT ONLY (not part of the excel-to-sigma skill): emit a Sigma MCP Code Mode program that restores the
Vantora forecast demo's seed state, so the next take starts clean.

Restores: Parameters Forecast Start Month -> the plan's first forecast month; rateInput rates -> the build's
rate seed (build/rate_seed.csv); deletes Versions rows other than the migrated v1..v7 and Approvals rows
other than the 3 migrated ones (FY27 v3 / v5 / v7). Hidden-page row ids are not readable, so every change is a
sigma.data.sql.execute DML statement (4 of the 5 allowed per program).

The scenario switch is a workbook control, not a table row; it is not persisted state, so it is not reset here
(a fresh page load shows Base).

Usage: reset-demo.py --workbook <id> --rate-seed <build/rate_seed.csv> [--out program.py]
Run the program via MCP build_workbook with target {"type": "existing", "workbookId": "<id>"}.
"""
import argparse, csv, sys

ap = argparse.ArgumentParser()
ap.add_argument("--workbook", required=True, help="workbook id (recorded in the program header only)")
ap.add_argument("--rate-seed", required=True, help="the build's rate_seed.csv (LINE_ITEM,RATE)")
ap.add_argument("--start-month", default="2027-07-01")
ap.add_argument("--model-page", default="model")
ap.add_argument("--inputs-page", default="pgAssumptions")
ap.add_argument("--versions", default="v1,v2,v3,v4,v5,v6,v7")
ap.add_argument("--approvals", default="FY27 v3,FY27 v5,FY27 v7")
ap.add_argument("--out")
a = ap.parse_args()

# SQL string literal, written so it can sit inside a single-quoted Python f-string in the emitted program
q = lambda s: "\\'" + s.replace("'", "''") + "\\'"
rates = [(r["LINE_ITEM"], float(r["RATE"])) for r in csv.DictReader(open(a.rate_seed))]
if not rates:
    sys.exit("empty rate seed")
case = " ".join(f"WHEN {q(li)} THEN {r}" for li, r in rates)
vers = ", ".join(q(v) for v in a.versions.split(","))
apps = ", ".join(q(v) for v in a.approvals.split(","))
n_ver, n_app = len(a.versions.split(",")), len(a.approvals.split(","))

prog = f'''# Reset the Vantora demo to seed state (workbook {a.workbook}). Demo kit only.
def_page = sigma.workbook.get_page({a.model_page!r})
inp_page = sigma.workbook.get_page({a.inputs_page!r})

params = def_page.get_element('parameters')
if not isinstance(params, InputTable):
    raise RuntimeError('parameters is not an InputTable')
p_src = sigma.data.source.element(params.id())
p_start = params.get_column('pa_start').id()
sigma.data.sql.execute(f'UPDATE {{p_src.ref()}} SET "{{p_start}}" = CAST({q(a.start_month)} AS timestamp)')

rates = inp_page.get_element('rateInput')
if not isinstance(rates, InputTable):
    raise RuntimeError('rateInput is not an InputTable')
r_src = sigma.data.source.element(rates.id())
r_li = rates.get_column('LINE_ITEM').id()
r_rate = rates.get_column('RATE').id()
sigma.data.sql.execute(f'UPDATE {{r_src.ref()}} SET "{{r_rate}}" = CASE "{{r_li}}" {case} ELSE "{{r_rate}}" END')

versions = def_page.get_element('versions')
if not isinstance(versions, InputTable):
    raise RuntimeError('versions is not an InputTable')
v_src = sigma.data.source.element(versions.id())
v_ver = versions.get_column('ve_ver').id()
sigma.data.sql.execute(f'DELETE FROM {{v_src.ref()}} WHERE "{{v_ver}}" IS NULL OR "{{v_ver}}" NOT IN ({vers})')

approvals = def_page.get_element('approvals')
if not isinstance(approvals, InputTable):
    raise RuntimeError('approvals is not an InputTable')
a_src = sigma.data.source.element(approvals.id())
a_req = approvals.get_column('ap_req').id()
sigma.data.sql.execute(f'DELETE FROM {{a_src.ref()}} WHERE "{{a_req}}" IS NULL OR "{{a_req}}" NOT IN ({apps})')
'''
(open(a.out, "w") if a.out else sys.stdout).write(prog)
