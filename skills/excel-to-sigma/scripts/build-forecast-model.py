#!/usr/bin/env python3
"""
build-forecast-model.py — one-shot builder for an actuals + driver-grown + manual
planning model (the architecture in refs/forecast-recipes.md). Consumes a model-plan
JSON (drafted from xlsx-discover.py's model map + the intent answers) and emits:
  - a Sigma data model (GL actuals VALUES + assumptions + raw forecast spine)
  - a workbook (union[actuals, forecast] → KPIs, P&L pivot w/ Actual|Forecast cols,
    monthly summary, trend, GL detail, assumptions; editable rate+manual input tables
    when intent is editable/what-if)
  - seed CSVs for the editable tables (the same rows the --emit-manual-commit-program
    setup program upserts through Sigma MCP Code Mode)
  - a LOCAL parity check (when the plan carries `parityTargets`, fails the run if the
    plan's recomputed total / derived rows differ from them: Revenue / Gross Profit /
    EBITDA / Net Income for the classic four-section plan, every total and profit row
    for a discovered-section plan)

Default = dry-run (writes specs to --out-dir, default /tmp/fcst-build; runs parity, no API;
pending review issues are reported as DRAFT MODE and do not block). `--post` builds live and
keeps the strict source-bound `--review-dispositions` contract.

Intent (plan.intent):
  read-only          → forecast baked from default rates (no input tables)
  editable | what-if → rates + manual cells are editable input tables (Coalesce override)
"""
import argparse, sys, os, re, json, csv, urllib.request, urllib.parse, urllib.error
from pathlib import Path

_LIB = Path(__file__).resolve().parent / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
import workbook_wire  # noqa: E402
import sigma_auth  # noqa: E402
from lib import forecast_calc as calc  # noqa: E402
from lib import demo_workflows  # noqa: E402
import forecast_app  # noqa: E402
import code_rep  # noqa: E402
from lib.migration_gate import (MigrationGateError, gate_plan_source, gate_plan_staging_source,  # noqa: E402
                                source_issue_hash)
import agent_enhance  # noqa: E402

CUR = lambda p: {"kind": "number", "formatString": p["currencyFormat"]}
PCT = lambda p: {"kind": "number", "formatString": p.get("pctFormat", ".1%")}
DATEF = {"kind": "datetime", "formatString": "%b-%y"}
def q(s): return "'" + str(s).replace("'", "''") + "'"


KPI_IDS = ("kpiRev", "kpiEbitda", "kpiNI", "kpiGM")


def forecast_layout(editable, extra_model_ids=(), chart_ids=(), kpi_ids=KPI_IDS, workflow=None):
    """Dense four-page layout: overview, detail, editable inputs, hidden plumbing.

    Source-chart ids sit beside the monthly trend (first) and then full width below it.
    """
    charts = [("monthlyTrend", 1, 25 if not chart_ids else 13, 10, 25)]
    if chart_ids:
        charts.append((chart_ids[0], 13, 25, 10, 25))
    for i, element_id in enumerate(chart_ids[1:]):
        charts.append((element_id, 1, 25, 25 + i * 15, 40 + i * 15))
    pivot_row = 25 + max(len(chart_ids) - 1, 0) * 15
    width = 24 // len(kpi_ids)
    kpi_cells = [(element_id, 1 + i * width, 1 + (i + 1) * width, 4, 10)
                 for i, element_id in enumerate(kpi_ids)]
    # With workflows the header row also carries the Parameters picker, the latest approval
    # status and the three workflow buttons; Versions / Approvals sit under the inputs.
    header_cells = ([("pnlHeader", 1, 6, 1, 4), ("parameters", 6, 11, 1, 4), ("approvalStatus", 11, 15, 1, 4),
                     ("btnRoll", 15, 18, 1, 4), ("btnSnapshot", 18, 21, 1, 4), ("btnSubmit", 21, 25, 1, 4)]
                    if workflow else [("pnlHeader", 1, 25, 1, 4)])
    workflow_inputs = [("versions", 1, 25, 49, 61), ("approvals", 1, 25, 61, 73)] if workflow else []
    pages = [
        ("pnl", [
            *header_cells,
            *kpi_cells,
            *charts,
            ("pnlPivot", 1, 25, pivot_row, pivot_row + 30),
        ]),
        ("detail", [("glTable", 1, 25, 1, 37)]),
        ("inputs", [
            ("assumpTable", 1, 25, 1, 13),
            *([("rateInput", 1, 25, 13, 29), ("manualInput", 1, 25, 29, 49)] if editable else []),
            *workflow_inputs,
        ]),
        ("model", [
            ("actuals", 1, 25, 1, 13),
            ("spineWrap", 1, 25, 13, 25),
            ("forecast", 1, 25, 25, 37),
            ("pnlCombined", 1, 25, 37, 49),
            *[(element_id, 1, 25, 49+i*12, 61+i*12) for i, element_id in enumerate(extra_model_ids)],
        ]),
    ]
    blocks = []
    for page_id, placements in pages:
        lines = [
            f'<Page type="grid" gridTemplateColumns="repeat(24, 1fr)" '
            f'gridTemplateRows="auto" id="{page_id}">'
        ]
        for element_id, col_start, col_end, row_start, row_end in placements:
            lines.append(
                f'  <Element elementId="{element_id}" gridColumn="{col_start} / {col_end}" '
                f'gridRow="{row_start} / {row_end}"/>'
            )
        lines.append("</Page>")
        blocks.append("\n".join(lines))
    return '<?xml version="1.0" encoding="utf-8"?>\n' + "\n".join(blocks)

def env():
    """Resolve Sigma auth (lib/sigma_auth.py: env token, auth.json, CLI profile, OAuth/client creds)."""
    base, tok, method = sigma_auth.get_session()
    return {"SIGMA_BASE_URL": base, "SIGMA_API_TOKEN": tok, "SIGMA_AUTH_METHOD": method}
def token(e): return e["SIGMA_API_TOKEN"]
def api(base, tok, method, path, body=None):
    return sigma_auth.api(method, path, body, base_url=base)

# ---------- DM ----------
def values_sql(rows, cols):
    body = ",\n".join("(" + ",".join(rows[i]) + ")" for i in range(len(rows)))
    return f"select * from (values\n{body}\n) as t({', '.join(cols)})"

def plan_model(P):
    """Normalise a plan to ``(periods, n_actual, sections)`` whatever vintage wrote it.

    Plans from the discovered-section drafter carry ``periods`` / ``sections``; older
    plans carry ``actualMonths`` / ``forecastMonths`` and the four classic sections.
    """
    periods, n_act = calc.plan_periods(P)
    if P.get("sections"):
        sections = [{"name": s["name"], "sign": s["sign"]}
                    for s in sorted(P["sections"], key=lambda s: s["order"])]
    else:
        sections = [{"name": name, "sign": CLASSIC_SIGNS[name]} for name in SECTIONS]
    return periods, n_act, sections


def sql_date(value):
    return f"{q(value)}::date" if value else "NULL::date"


def spine_terms(P, ev):
    """``(kind, rate, manual, base, exponent)`` per line and forecast period.

    ``growth`` and ``yoy`` both evaluate ``base * (1 + rate) ^ exponent`` (a year-over-year
    cell restarts from the same period last year); ``pct`` is resolved in a second pass.
    """
    periods, n_act = calc.plan_periods(P)
    out = {}
    for ln in P["lines"]:
        rows = []
        for i in range(len(periods) - n_act):
            idx = n_act + i
            n = periods[idx]["num"] - periods[n_act - 1]["num"]
            kind = ln["kind"]
            if kind == "yoy":
                hops, at = calc.yoy_terms(idx, n_act - 1, ln["lag"])
                rows.append(("growth", ln["rate"], None, ev.matrix[ln["line"]][at], hops))
            elif kind == "growth":
                rows.append(("growth", ln["rate"], None, ln["juneBase"], n))
            elif kind == "pct_of_total":
                rows.append(("pct", ln["rate"], None, None, n))
            elif kind == "flat":
                rows.append(("flat", None, None, ln["juneBase"], n))
            else:
                rows.append(("manual", None, ln["manual"][i], None, n))
        out[ln["line"]] = rows
    return out


def evaluation(P):
    """The plan's recomputed values; ``PlanError`` when its shape is unsupported."""
    plan = P
    if not P.get("sections"):
        periods, _, sections = plan_model(P)
        plan = {**P, "sections": [{"name": s["name"], "order": i} for i, s in enumerate(sections, 1)]}
    try:
        return calc.Evaluation(plan)
    except (calc.ExprError, ZeroDivisionError) as error:
        raise PlanError(str(error)) from None


# Landed-table column roles: DM output name -> accepted source headers (lower case).
LANDED_ALIASES = {"Txn Date": ("posting date", "txn date", "date"), "Month": ("month",),
                  "Line Item": ("line item",), "Amount": ("amount",),
                  "Department": ("department", "region"), "Vendor": ("vendor", "product line")}


LANDED_TYPES = {"Txn Date": "DATE", "Month": "VARCHAR", "Line Item": "VARCHAR", "Amount": "NUMBER",
                "Department": "VARCHAR", "Vendor": "VARCHAR"}


def landed_actuals_source(landing, actuals_table, P):
    """Resolve ``--actuals-table DB.SCHEMA.TABLE`` against landing.json.

    Returns ``{"path": [DB, SCHEMA, TABLE], "columns": {output name: landed column name}}``;
    column names come from landing.json (the same identifiers the landing step created).
    """
    parts = actuals_table.split(".")
    if len(parts) != 3:
        raise PlanError(f"--actuals-table must be DB.SCHEMA.TABLE, got {actuals_table!r}")
    db, schema, table = (x.upper() for x in parts)
    if (landing.get("database", "").upper(), landing.get("schema", "").upper()) != (db, schema):
        raise PlanError(f"landing.json database/schema do not match {actuals_table}")
    entry = next((t for t in landing["tables"] if t["snowflakeTable"].upper() == table), None)
    if entry is None:
        raise PlanError(f"{actuals_table} is not in landing.json")
    by_header = {c["header"].strip().lower(): c for c in entry["columns"]}
    columns, labels = {}, {}
    for out, aliases in LANDED_ALIASES.items():
        col = next((by_header[a] for a in aliases if a in by_header), None)
        if col:
            columns[out] = col["name"]
            labels[out] = col["header"].strip() if out in ("Department", "Vendor") else out
            expected = LANDED_TYPES[out]
            actual = str(col.get("type", "")).upper()
            if not actual.startswith(expected):
                raise PlanError(f"landed column {col['name']} ({out}) must be {expected} but landing.json "
                                f"says {actual or 'unknown'}; the model's Switch/Date formulas need it")
    for required in ("Txn Date", "Month", "Line Item", "Amount"):
        if required not in columns:
            raise PlanError(f"landed table {actuals_table} has no column for {required!r} "
                            f"(accepted headers: {', '.join(LANDED_ALIASES[required])})")
    landed = {"path": [db, schema, table], "columns": columns, "labels": labels}
    _cross_check_landing(entry, columns, P, actuals_table)
    return landed


def _cross_check_landing(entry, columns, P, actuals_table):
    """Landed CSV (blank Amount = 0) must carry the plan's row count and actuals total. A landing.json
    whose CSV is gone cannot be cross-checked: say so instead of passing silently."""
    path = entry.get("csv")
    if not path or not os.path.exists(path):
        print(f"WARNING: landed CSV for {actuals_table} " + (f"({path}) " if path else "")
              + "not found — the landed actuals were NOT cross-checked against the plan's actuals total; "
              "re-run land-in-snowflake.py (dry run) to rewrite it")
        return
    total, rows = 0.0, 0
    with open(path, encoding="utf-8", newline="") as stream:
        for row in csv.DictReader(stream):
            rows += 1
            total += float(row[columns["Amount"]] or 0)
    plan_total = sum(float(r["amt"]) for r in P["actuals"])
    if abs(total - plan_total) > 0.005:
        raise PlanError(f"landed {actuals_table} ({rows} rows, Amount {total:,.2f}) does not match the plan's "
                        f"actuals total {plan_total:,.2f}; re-land or re-draft the plan")


PLAN_ALIASES = {"Month": ("month",), "Line Item": ("line item",), "Region": ("region",),
                "Plan Amount": ("plan amount", "plan", "budget amount", "budget")}


def landed_plan_source(landing, P):
    """The landed plan table (``tblPlan``) from landing.json, or ``None`` when none was landed.

    A plan table is the excel table named like ``*plan*`` / ``*budget*`` or any landed table with a
    ``Plan Amount`` header; it needs Month, Line Item and a plan amount column. Returns
    ``{"path", "columns": {role: landed column}, "regions": {line item: region}}``; every month must
    be a plan period and every line item a plan line (else ``PlanError``, never a silent zero).
    """
    def is_plan(entry):
        headers = {c["header"].strip().lower() for c in entry["columns"]}
        named = re.search(r"plan|budget", entry.get("excelTable", ""), re.I)
        return bool(headers & set(PLAN_ALIASES["Plan Amount"])) and (named or "plan amount" in headers)
    entry = next((t for t in landing.get("tables", []) if is_plan(t)), None)
    if entry is None:
        return None
    by_header = {c["header"].strip().lower(): c for c in entry["columns"]}
    columns = {}
    for role, aliases in PLAN_ALIASES.items():
        col = next((by_header[a] for a in aliases if a in by_header), None)
        if col:
            columns[role] = col["name"]
    for required in ("Month", "Line Item", "Plan Amount"):
        if required not in columns:
            raise PlanError(f"landed plan table {entry['snowflakeTable']} has no {required!r} column")
    regions = {}
    path = entry.get("csv")
    if path and os.path.exists(path):
        periods, _n, _s = plan_model(P)
        labels = {p["label"] for p in periods}
        lines = {ln["line"] for ln in P["lines"]}
        with open(path, encoding="utf-8", newline="") as stream:
            for row in csv.DictReader(stream):
                month, line = row[columns["Month"]], row[columns["Line Item"]]
                if month not in labels:
                    raise PlanError(f"plan table month {month!r} is not a period of the forecast")
                if line not in lines:
                    raise PlanError(f"plan table line {line!r} is not a forecast line")
                if "Region" in columns and row[columns["Region"]]:
                    regions.setdefault(line, row[columns["Region"]])
    db, schema = landing["database"].upper(), landing["schema"].upper()
    return {"path": [db, schema, entry["snowflakeTable"].upper()], "columns": columns, "regions": regions}


def _make_date(iso):
    y, m, d = (int(x) for x in iso[:10].split("-"))
    return f"MakeDate({y}, {m}, {d})"


def landed_plan_detail(P, plan_src):
    """planDetail: the landed plan table on the plan's connection, with the forecast's Section /
    order / Month Num / Month End attached by Switch (the landed table carries only names)."""
    periods, _n, _s = plan_model(P)
    t = plan_src["path"][-1]; c = plan_src["columns"]
    ref = lambda role: f"[{t}/{c[role]}]"
    lines = {ln["line"]: ln for ln in P["lines"]}
    cols = [("Month", ref("Month")), ("Line Item", ref("Line Item")),
            ("Region", ref("Region") if "Region" in c else '""'),
            ("Plan Amount", ref("Plan Amount")),
            ("Month Num", _switch(ref("Month"), {p["label"]: p["num"] for p in periods}, 0)),
            ("Month End", _switch(ref("Month"), {p["label"]: _make_date(p["end"]) for p in periods if p.get("end")}, "Null")),
            ("Section", _switch(ref("Line Item"), {k: f'"{v["section"]}"' for k, v in lines.items()}, '""')),
            ("Section Order", _switch(ref("Line Item"), {k: v["secOrder"] for k, v in lines.items()}, 0)),
            ("Line Order", _switch(ref("Line Item"), {k: v["lineOrder"] for k, v in lines.items()}, 0))]
    return {"id": "planDetail", "kind": "table", "name": "Plan Detail",
            "source": {"kind": "warehouse-table", "connectionId": P["connectionId"], "path": plan_src["path"]},
            "columns": [{"id": f"c{i}", "name": nm, "formula": f} for i, (nm, f) in enumerate(cols, 1)]}


def summary_metrics(P):
    """P&L summary rows ``[(label, {section: weight})]``: the revenue section, then each amount row,
    each but the first preceded by the sections it adds (Revenue, Gross margin, Opex, EBITDA)."""
    _periods, _n, sections = plan_model(P)
    first = next((s["name"] for s in sections if s["sign"] > 0), sections[0]["name"])
    out, seen = [(first, {first: 1})], {first}
    rows = [r for r in amount_rows(P, sections) if not r.get("lineCoefficients")]
    for i, row in enumerate(rows):
        added = [s["name"] for s in sections if s["name"] in row["coefficients"] and s["name"] not in seen]
        if i:
            out += [(name, {name: 1}) for name in added]
        seen |= set(row["coefficients"])
        out.append((row["label"], dict(row["coefficients"])))
    return out


def metric_map_element(P):
    """metricMap: one row per (summary metric, section, weight) for the P&L summary join."""
    rows = []
    for order, (label, coeffs) in enumerate(summary_metrics(P), 1):
        for section, weight in coeffs.items():
            rows.append([q(forecast_app.sentence(label)), str(order), q(section), f"{float(weight)}"])
    return {"id": "metricMap", "kind": "table", "source": {"kind": "sql", "connectionId": P["connectionId"],
            "statement": values_sql(rows, ["METRIC", "METRIC_ORDER", "SECTION", "WEIGHT"])},
            "columns": dmcols([("METRIC", "Metric"), ("METRIC_ORDER", "Metric Order"), ("SECTION", "Section"),
                               ("WEIGHT", "Weight")])}


def scenario_spreads(P):
    """``[(scenario, column name, {line: spread})]`` for every non-default driver scenario: the
    scenario's rate minus the default scenario's rate, per growth line (0 for other lines)."""
    scn = P.get("scenarios")
    if not scn:
        return []
    out = []
    for name in scn["names"]:
        if name == scn["default"]:
            continue
        if '"' in name or "\\" in name or "/" in name or "[" in name or "]" in name:
            raise PlanError(f"scenario name {name!r} cannot be carried into a Sigma column name")
        spread = {}
        for ln in P["lines"]:
            rates = scn["rates"].get(ln.get("driver") or "")
            spread[ln["line"]] = (rates[name] - rates[scn["default"]]) if rates and ln["kind"] in RATE_KINDS else 0.0
        out.append((name, f"{name} Spread", spread))
    return out


def _switch(key, mapping, default):
    for k in mapping:
        if '"' in str(k) or "\\" in str(k):
            raise PlanError(f"value {k!r} contains a quote or backslash the Sigma formulas cannot carry")
    body = ", ".join(f'"{k}", {v}' for k, v in mapping.items())
    return f"Switch({key}, {body}, {default})"


def landed_gl_detail(P, landed, all_periods=False):
    """glDetail as a warehouse-table source on the plan's connection (no inline VALUES).

    Landed columns pass through; Month Num / Month End / Section / Section Order / Line Order,
    which the landed table does not carry, are Switch / date formulas built from the plan.
    With ``all_periods`` (the planning app, where Roll forward closes a month) Month Num maps every
    period, so actuals landed for a month the source still forecast keep their period number.
    """
    periods, n_act, _ = plan_model(P)
    views = calc.month_views(periods) if all_periods else calc.month_views(periods)[:n_act]
    t = landed["path"][-1]; c = landed["columns"]
    ref = lambda out: f"[{t}/{c[out]}]"
    lab = lambda out: landed.get("labels", {}).get(out, out)  # user-visible name = source header
    lines = {}
    for r in P["actuals"]:
        lines.setdefault(r["line"], r)
    cols = [
        ("Txn Date", ref("Txn Date")), ("Month", ref("Month")),
        ("Month Num", _switch(ref("Month"), {m["month"]: m["num"] for m in views}, 0)),
        ("Month End", f'DateAdd("day", -1, DateAdd("month", 1, DateTrunc("month", {ref("Txn Date")})))'),
        ("Line Item", ref("Line Item")),
        ("Section", _switch(ref("Line Item"), {k: f'"{v["section"]}"' for k, v in lines.items()}, '""')),
        ("Section Order", _switch(ref("Line Item"), {k: v["secOrder"] for k, v in lines.items()}, 0)),
        ("Line Order", _switch(ref("Line Item"), {k: v["lineOrder"] for k, v in lines.items()}, 0)),
        (lab("Department"), ref("Department") if "Department" in c else '""'),
        (lab("Vendor"), ref("Vendor") if "Vendor" in c else '""'),
        ("Amount", ref("Amount")),
    ]
    return {"id": "glDetail", "kind": "table", "name": "GL Detail",
            "source": {"kind": "warehouse-table", "connectionId": P["connectionId"], "path": landed["path"]},
            "columns": [{"id": f"c{i}", "name": nm, "formula": f} for i, (nm, f) in enumerate(cols, 1)]}


def build_dm_spec(P, landed=None, plan_src=None, app=False):
    conn = P["connectionId"]
    periods, n_act, _ = plan_model(P)
    views = calc.month_views(periods)
    actual_views, forecast_views = views[:n_act], views[n_act:]
    has_pct = any(ln["kind"] == "pct_of_total" for ln in P["lines"])
    ev = evaluation(P) if any(ln["kind"] == "yoy" for ln in P["lines"]) else None
    # GL detail (actuals, transaction grain)
    gl = []
    for r in P["actuals"]:
        gl.append([sql_date(r['date']), q(r['month']), q(r['line']), q(r['section']),
                   str(r['secOrder']), str(r['lineOrder']), q(r['dept']), q(r['vendor']), f"{float(r['amt']):.2f}"])
    mnum = {m["month"]: m["num"] for m in actual_views}; mend = {m["month"]: m["end"] for m in actual_views}
    gl_cols = ["TXN_DATE", "MONTH", "LINE_ITEM", "SECTION", "SEC_ORDER", "LINE_ORDER", "DEPARTMENT", "VENDOR", "AMOUNT"]
    # add MONTH_NUM / MONTH_END derived
    for row, r in zip(gl, P["actuals"]):
        row.insert(2, str(mnum[r['month']])); row.insert(3, sql_date(mend[r['month']]))
    gl_cols[2:2] = ["MONTH_NUM", "MONTH_END"]
    glDetail = landed_gl_detail(P, landed, all_periods=app) if landed else {"id": "glDetail", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(gl, gl_cols)},
                "columns": dmcols([("TXN_DATE","Txn Date"),("MONTH","Month"),("MONTH_NUM","Month Num"),("MONTH_END","Month End"),("LINE_ITEM","Line Item"),("SECTION","Section"),("SEC_ORDER","Section Order"),("LINE_ORDER","Line Order"),("DEPARTMENT","Department"),("VENDOR","Vendor"),("AMOUNT","Amount")])}
    # assumptions
    arows = [[q(a), f"{b}", q(c)] for a, b, c in P["assumptions"]]
    rate_name = "Monthly Rate" if periods[0]["grain"] == "month" else "Rate"
    assumptions = {"id": "assumptions", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(arows, ["DRIVER","MONTHLY_RATE","BASIS"])},
                   "columns": dmcols([("DRIVER","Driver"),("MONTHLY_RATE",rate_name),("BASIS","Basis")])}
    # raw forecast spine
    terms = spine_terms(P, ev)
    srows = []
    spreads = scenario_spreads(P) if app else []
    for ln in P["lines"]:
        for i, m in enumerate(forecast_views):
            kind, rate, manual, base, n = terms[ln["line"]][i]
            if app and ln["kind"] == "yoy":
                kind = "yoy"   # the planning app re-seeds growth lines at Roll forward, not year-over-year ones
            row = [q(ln["line"]), q(ln["section"]), str(ln["secOrder"]), str(ln["lineOrder"]),
                   q(m["month"]), str(m["num"]), sql_date(m["end"]), str(n), q(kind),
                   f"{rate}" if rate is not None else "NULL",
                   f"{manual:.2f}" if manual is not None else "NULL",
                   f"{base:.2f}" if base is not None else "NULL"]
            if has_pct:
                row.append(q(ln["of"]) if ln["kind"] == "pct_of_total" else "NULL")
            row += [f"{spread[ln['line']]:.6f}" for _n, _c, spread in spreads]
            srows.append(row)
    scols = ["LINE_ITEM","SECTION","SEC_ORDER","LINE_ORDER","MONTH","MONTH_NUM","MONTH_END","N","KIND","DEFAULT_RATE","MANUAL_VALUE","JUNE_BASE"]
    spine_pairs = [("LINE_ITEM","Line Item"),("SECTION","Section"),("SEC_ORDER","Section Order"),("LINE_ORDER","Line Order"),("MONTH","Month"),("MONTH_NUM","Month Num"),("MONTH_END","Month End"),("N","N"),("KIND","Kind"),("DEFAULT_RATE","Default Rate"),("MANUAL_VALUE","Manual Value"),("JUNE_BASE","June Base")]
    if has_pct:
        scols.append("OF_SECTION"); spine_pairs.append(("OF_SECTION", "Of Section"))
    for i, (_name, col_name, _spread) in enumerate(spreads, 1):
        scols.append(f"SPREAD_{i}"); spine_pairs.append((f"SPREAD_{i}", col_name))
    spine = {"id": "spine", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(srows, scols)},
             "columns": dmcols(spine_pairs)}
    return {"name": P["name"] + " — Model", "schemaVersion": 1, "folderId": P["folderId"],
            "pages": [{"id": "model", "name": "Model", "elements": [glDetail, assumptions, spine,
                                                                     *([landed_plan_detail(P, plan_src)] if plan_src and app else []),
                                                                     *([metric_map_element(P)] if app else [])]}]}

def dmcols(pairs):
    return [{"id": f"c{i}", "name": nm, "formula": f"[Custom SQL/{al}]"} for i, (al, nm) in enumerate(pairs, 1)]

def map_dm_elements(spec):
    """Identify element ids by a signature column name (names round-trip; ids reassign)."""
    out = {}
    for el in spec["pages"][0]["elements"]:
        names = {c["name"] for c in el["columns"]}
        if "Txn Date" in names: out["glDetail"] = el["id"]
        elif "Kind" in names: out["spine"] = el["id"]
        elif "Driver" in names: out["assumptions"] = el["id"]
        elif "Plan Amount" in names: out["planDetail"] = el["id"]
        elif "Weight" in names: out["metricMap"] = el["id"]
    return out

# ---------- parity (local) ----------
SECTIONS = ("Revenue", "Cost of Revenue", "Operating Expenses", "Below the Line")
CLASSIC_SIGNS = {"Revenue": 1, "Cost of Revenue": -1, "Operating Expenses": -1, "Below the Line": -1}
PARITY_METRICS = ("Revenue", "Gross Profit", "EBITDA", "Net Income")


class PlanError(ValueError):
    """The model plan has a shape this builder does not support."""


def _section_sum_key(section, where):
    if section not in SECTIONS:
        raise PlanError(
            f"unsupported section {section!r} in {where}; this builder supports only "
            + ", ".join(SECTIONS) + " (unknown sections need explicit review, not coercion)")
    return section


def classic_parity(P):
    """Revenue / Gross Profit / EBITDA / Net Income for a plan without discovered sections."""
    secsum = {s: 0.0 for s in SECTIONS}
    # actuals
    for i, r in enumerate(P["actuals"]):
        secsum[_section_sum_key(r["section"], f"actuals[{i}] ({r.get('line')!r})")] += r["amt"]
    # forecast
    periods, n_act = calc.plan_periods(P)
    last = periods[n_act - 1]["num"]
    for ln in P["lines"]:
        _section_sum_key(ln["section"], f"line {ln['line']!r}")
        base = ln.get("juneBase", 0)
        for i, m in enumerate(periods[n_act:]):
            n = m["num"] - last
            v = ln["manual"][i] if ln["kind"] == "manual" else (base if ln["kind"] == "flat" else base * (1 + ln["rate"]) ** n)
            secsum[ln["section"]] += v
    rev, cogs, opex, below = secsum["Revenue"], secsum["Cost of Revenue"], secsum["Operating Expenses"], secsum["Below the Line"]
    return {"Revenue": rev, "Gross Profit": rev - cogs, "EBITDA": rev - cogs - opex, "Net Income": rev - cogs - opex - below}


def parity(P):
    """Recomputed total / derived rows keyed like ``parityTargets``.

    Plans from the discovered-section drafter evaluate every section total and derived
    row over the periods the FY column covers; older plans keep the classic four.
    """
    if not P.get("sections"):
        return classic_parity(P)
    ev = evaluation(P)
    fy = P.get("fyPeriods")
    idx = [i for i, label in enumerate(ev.labels) if fy is None or label in fy]
    rows = []
    for meta in P["sections"]:
        if meta.get("totalLabel"):
            rows.append((meta.get("totalRow", 0), meta["totalLabel"], ev.section_total(meta["name"], idx)))
    for d in P.get("derivedRows", []):
        try:
            rows.append((d.get("sourceRow", 0), d["label"], ev.aggregate(d["label"], idx)))
        except calc.ExprError:
            continue  # e.g. a ratio over a zero FY total: no value, so a target on it fails loudly
    out = {}
    for _, label, value in sorted(rows, key=lambda r: r[0]):
        out.setdefault(calc.TARGET_LABELS.get(label.lower(), label), value)
    return out


def check_parity(P, computed, tolerance):
    """Compare computed totals to plan['parityTargets']; print one line per metric.

    Returns True when every present target matches: within ``tolerance`` (absolute) when one
    is given, else within float noise (``calc.close``).
    Plans without targets are reported as unchecked and do not fail.
    """
    targets = P.get("parityTargets") or {}
    present = [m for m, v in targets.items() if v is not None]
    if not present:
        print("LOCAL PARITY: not checked — plan has no parityTargets")
        return True
    ok = True
    for metric in present:
        want = float(targets[metric])
        if metric not in computed:
            ok = False
            print(f"PARITY FAIL {metric}: plan has no row to recompute it; source={want:,.2f}")
            continue
        got = computed[metric]
        passed = calc.close(got, want, tolerance)
        ok = ok and passed
        print(f"PARITY {'PASS' if passed else 'FAIL'} {metric}: plan={got:,.2f} source={want:,.2f} "
              f"diff={got - want:,.6f}")
    return ok

# ---------- seed CSVs ----------
RATE_KINDS = ("growth", "yoy", "pct_of_total")  # lines whose rate is an editable driver


RATE_SEED_COLUMNS = ("LINE_ITEM", "RATE")
MANUAL_SEED_COLUMNS = ("LINE_ITEM", "MONTH", "MONTH_NUM", "MONTH_END", "AMOUNT")


def seed_rows(P):
    """Starting rows for the editable ``rateInput`` and ``manualInput`` tables, one list per
    row in ``RATE_SEED_COLUMNS`` / ``MANUAL_SEED_COLUMNS`` order (a blank month end is None).

    The single source for both the seed CSVs and the setup program's upsert_rows seeds;
    AMOUNT is the CSV's 2-decimal value."""
    rate = [[ln["line"], ln["rate"]] for ln in P["lines"] if ln["kind"] in RATE_KINDS]
    manual = []
    manual_lines = [ln for ln in P["lines"] if ln["kind"] == "manual"]
    if manual_lines:
        periods, n_act = calc.plan_periods(P)
        months = calc.month_views(periods)[n_act:]
        for ln in manual_lines:
            for i, m in enumerate(months):
                manual.append([ln["line"], m["month"], m["num"], m["end"] or None,
                               float(f"{ln['manual'][i]:.2f}")])
    return rate, manual


def seed_csvs(P, outdir):
    rate_p = os.path.join(outdir, "rate_seed.csv"); man_p = os.path.join(outdir, "manual_seed.csv")
    rate, manual = seed_rows(P)
    with open(rate_p, "w", newline="") as f:
        w = csv.writer(f); w.writerow(RATE_SEED_COLUMNS)
        w.writerows(rate)
    with open(man_p, "w", newline="") as f:
        w = csv.writer(f); w.writerow(MANUAL_SEED_COLUMNS)
        for line, month, num, end, amount in manual:
            w.writerow([line, month, num, end or "", f"{amount:.2f}"])
    return rate_p, man_p

# ---------- workbook ----------
LEGACY_AMT = {"Revenue": ("c_rev", "Revenue Amt"), "Cost of Revenue": ("c_cogs", "COGS Amt"),
              "Operating Expenses": ("c_opex", "OpEx Amt"), "Below the Line": ("c_below", "Below Amt")}


def section_columns(sections):
    """``{section: (column id, column name, string literal)}`` for the combined table.

    The four classic sections keep their historical column names; discovered sections
    get ``<name> Amt`` with characters Sigma column paths cannot carry (slash, brackets)
    replaced by spaces.
    """
    classic = all(sec["name"] in LEGACY_AMT for sec in sections)
    out = {}
    for i, sec in enumerate(sections, 1):
        if classic:
            col_id, col_name = LEGACY_AMT[sec["name"]]
        else:
            col_id, col_name = f"c_s{i}", re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9&]+", " ", sec["name"])).strip() + " Amt"
        out[sec["name"]] = (col_id, col_name, '"' + sec["name"] + '"')
    return out


def amount_rows(P, sections):
    """Amount derived rows ``[{label, coefficients}]`` driving KPIs, the trend and charts.

    Plans without derived rows (older plans) get the classic Gross Profit / EBITDA / Net
    Income; a discovered-section plan with none gets one signed bottom line.
    """
    if P.get("derivedRows"):
        rows = [{"label": d["label"], "coefficients": d["coefficients"],
                 **({"lineCoefficients": d["lineCoefficients"]} if d.get("lineCoefficients") else {})}
                for d in P["derivedRows"] if d.get("kind") == "amount"]
        if rows:
            return rows
    if not P.get("sections"):
        names = [s["name"] for s in sections]
        return [{"label": label, "coefficients": {n: CLASSIC_SIGNS[n] for n in names[:count]}}
                for label, count in (("Gross Profit", 2), ("EBITDA", 3), ("Net Income", 4))]
    return [{"label": "Net Result", "coefficients": {s["name"]: s["sign"] for s in sections}}]


def kpi_rows(P, rows):
    """``(middle, bottom)`` indexes into ``rows`` for the two profit KPI cards and the trend.

    Chosen by role, not by position in the sheet: the bottom line is the amount row that
    depends on the most sections (every section when the P&L has a true bottom line; a row
    labelled Net Income wins a tie), the latest such row on a further tie. The middle card
    is the row the bottom line reads that depends on the most sections (EBITDA for a classic
    P&L, Operating Profit for a retail budget), else the best remaining row. An extra
    amount row such as ``Contribution = Subscription - COGS`` touches few sections, so it
    never takes the bottom-line card.
    """
    line_section = {ln["line"]: ln["section"] for ln in P.get("lines", [])}
    exprs = {d["label"]: d["expr"] for d in P.get("derivedRows", []) if "expr" in d}
    cover = [len(set(r["coefficients"]) | {line_section[l] for l in r.get("lineCoefficients", {})
                                           if l in line_section}) for r in rows]
    best = [i for i, c in enumerate(cover) if c == max(cover)]
    named = [i for i in best if rows[i]["label"].strip().lower() == "net income"]
    bottom = (named or best)[-1]
    others = [i for i in range(len(rows)) if i != bottom]
    if not others:
        return None, bottom
    read, todo = set(), [rows[bottom]["label"]]
    while todo:
        label = todo.pop()
        for ref in calc.derived_references(exprs[label]) if label in exprs else ():
            if ref not in read:
                read.add(ref)
                todo.append(ref)
    middle = max(others, key=lambda i: (rows[i]["label"] in read, cover[i], i))
    return middle, bottom


def section_sum(table, sec_cols, fy_nums, name, fy=False):
    """Sum of one section's amount column over ``pnlCombined`` (the FY periods only with ``fy``)."""
    col = f"[{table}/{sec_cols[name][1]}]"
    if fy and fy_nums:
        return f"SumIf({col}, [{table}/Month Num] >= {fy_nums[0]} and [{table}/Month Num] <= {fy_nums[-1]})"
    return f"Sum({col})"


def amount_formula(row, table, sec_cols, fy_nums, fy=False):
    """Sigma formula for an amount row over ``pnlCombined``: section totals and single lines.

    ``row`` is ``{coefficients: {section: weight}, lineCoefficients: {line: weight}}``. A
    section term sums that section's amount column; a line term sums the Amount column for
    that Line Item only (``SumIf``), so a row such as ``Subscription - COGS`` is not widened
    to the whole revenue section. With ``fy`` the sums cover only the FY column's periods.
    """
    window = ""
    if fy and fy_nums:
        window = f" and [{table}/Month Num] >= {fy_nums[0]} and [{table}/Month Num] <= {fy_nums[-1]}"
    terms = [(section_sum(table, sec_cols, fy_nums, name, fy), weight)
             for name, weight in row["coefficients"].items()]
    for label in row.get("lineCoefficients", {}):
        if '"' in label or "\\" in label:
            raise PlanError(f"line {label!r} contains a quote or backslash the Sigma formulas cannot carry")
    terms += [(f'SumIf([{table}/Amount], [{table}/Line Item] = "{label}"{window})', weight)
              for label, weight in row.get("lineCoefficients", {}).items()]
    out = ""
    for i, (term, weight) in enumerate(terms):
        term = term if abs(weight) == 1 else f"{abs(weight):g} * {term}"
        out += (("-" if weight < 0 else "") + term) if i == 0 else (" - " if weight < 0 else " + ") + term
    return out


# Planning-app input tables: the dense grid (presentation rows clip a 10-row driver table), autofit off and
# each column sized to its content by the setup program (demo_workflows.input_column_widths) -- a number
# column stretched across the table shows its edit pencil far left and its right-aligned name far right,
# which reads as an empty unlabelled column (autofit and the presentation preset both stretch, live d2).
APP_INPUT_STYLE = {"preset": "spreadsheet", "gridLines": "horizontal", "autofitColumns": "hidden"}


def cost_lines(P):
    """P&L summary labels where lower is better (every section they add is a cost section): their
    variance to plan is favourable when negative."""
    _periods, _n, sections = plan_model(P)
    sign = {s["name"]: s["sign"] for s in sections}
    return [forecast_app.sentence(label) for label, coeffs in summary_metrics(P)
            if coeffs and all(sign.get(sec, 1) < 0 and w > 0 for sec, w in coeffs.items())]


def month_key(date_expr):
    """Integer month index (year * 12 + month) of a date formula: the boundary's comparison key."""
    return f'DatePart("year", {date_expr}) * 12 + DatePart("month", {date_expr})'


def forecast_boundary(P, parameters=True):
    """The planning app's actual / forecast boundary: Parameters.Forecast Start Month (the source's first
    forecast month until the Parameters row is seeded), defined ONCE as the hidden ``Forecast Boundary``
    element (forecast_app.boundary_element). ``{"element", "start_key", "last_key"}``: that element, the
    formula every other element compares month keys with, and the source's last actual month key (where
    the plan's June Base sits)."""
    periods, n_act, _ = plan_model(P)
    if periods[0]["grain"] != "month" or not periods[n_act].get("end") or not periods[n_act - 1].get("end"):
        raise PlanError("the planning-app layout needs dated periods (month grain)")
    first, last = periods[n_act]["end"], periods[n_act - 1]["end"]
    return {"element": forecast_app.boundary_element(P["connectionId"], (int(first[:4]), int(first[5:7])),
                                                     parameters=parameters),
            "start_key": forecast_app.START_KEY, "last_key": int(last[:4]) * 12 + int(last[5:7])}


def apply_boundary(actuals, stage1, prefix, bnd, has_pct, read_only_cols=None):
    """Roll forward = a month just closed: months before Parameters.Forecast Start Month read the landed
    actuals, months from it on are the driver forecast re-seeded off the last closed month.

    ``actuals`` keeps only rows before the boundary and gains the ``Month Key`` the forecast looks up;
    ``stage1`` (the forecast's first element, columns read through ``prefix``) keeps the rows from the
    boundary on, plus any closed month whose actuals have not landed for that line (``Landed`` is NULL):
    that month keeps its forecast value, labelled Period ``Awaiting actuals``, never blank or dropped.
    A growth / flat line seeds from the line's actual in the month before the boundary (its frozen
    forecast value while that month's actuals have not landed), with N counted from there; manual lines
    keep their typed values and year-over-year lines their source seeding.
    ``read_only_cols`` adds the Kind / June Base / Eff Rate / Eff Manual columns a read-only forecast
    does not carry. The actuals keep their constant Period label.
    """
    S = bnd["start_key"]
    actuals["columns"] += [
        {"id": "a_mk", "name": "Month Key", "formula": month_key(f"[{actuals_prefix(actuals)}/Month End]")},
        {"id": "a_win", "name": "Window", "formula": f'If([Month Key] < {S}, "Actual", "Forecast")'}]
    actuals["filters"] = [{"id": "a_wf", "columnId": "a_win", "kind": "list", "mode": "include", "values": ["Actual"]}]
    cols = stage1["columns"]
    names = {c["name"] for c in cols}
    at = next(i for i, c in enumerate(cols) if c["name"] == "Amount")
    extra = [c for c in (read_only_cols or []) if c["name"] not in names]
    extra += [] if "N" in names else [{"id": "f_n", "name": "N", "formula": f"[{prefix}/N]"}]
    extra += [
        {"id": "f_mk", "name": "Month Key", "formula": month_key(f"[{prefix}/Month End]")},
        {"id": "f_sk", "name": "Start Key", "formula": S},
        {"id": "f_bk", "name": "Base Key", "formula": "[Start Key] - 1"},
        {"id": "f_land", "name": "Landed",
         "formula": "Lookup(Sum([Actuals/Amount]), [Line Item], [Actuals/Line Item], [Month Key], [Actuals/Month Key])"},
        {"id": "f_win", "name": "Window", "formula": (f'If([Month Key] >= [Start Key], "Forecast", IsNull([Landed]), '
                                                       f'"{forecast_app.AWAITING}", "Actual")')},
        {"id": "f_la", "name": "Last Actual",
         "formula": "Lookup(Sum([Actuals/Amount]), [Line Item], [Actuals/Line Item], [Base Key], [Actuals/Month Key])"},
        {"id": "f_seed", "name": "Seed",
         "formula": ('Coalesce([Last Actual], If([Kind] = "flat", [June Base], '
                     f'[June Base] * (1 + [Eff Rate]) ^ ([Base Key] - {bnd["last_key"]})))')}]
    cols[at:at] = extra
    for c in cols:
        if c["name"] == "N":
            c["formula"] = f'If([Kind] = "yoy", [{prefix}/N], [Month Key] - [Base Key])'
        if c["name"] == "Amount":
            c["formula"] = ('If([Kind] = "manual", [Eff Manual], [Kind] = "flat", [Seed], '
                            + ('[Kind] = "pct", 0, ' if has_pct else "")
                            + '[Kind] = "yoy", [June Base] * (1 + [Eff Rate]) ^ [N], [Seed] * (1 + [Eff Rate]) ^ [N])')
        if c["name"] == "Period":
            c["formula"] = f'If([Window] = "{forecast_app.AWAITING}", "{forecast_app.AWAITING}", "Forecast")'
    stage1["filters"] = [{"id": "f_wf", "columnId": "f_win", "kind": "list", "mode": "include",
                          "values": ["Forecast", forecast_app.AWAITING]}]


def actuals_prefix(actuals):
    """The source prefix the actuals wrap reads its columns through (``GL Detail`` / ``Custom SQL``)."""
    return actuals["columns"][0]["formula"][1:].split("/", 1)[0]


def build_wb_spec(P, dm_id, els, *, approved_inputs=None, landed=None, workflows=None, plan_src=None, history=None):
    conn = P["connectionId"]; editable = P["intent"] in ("editable", "what-if"); DMP = "Custom SQL"; C = "PnL Combined"
    GLP = "GL Detail" if landed else DMP  # a DM table element is referenced by its name
    DEPT = (landed or {}).get("labels", {}).get("Department", "Department")
    VEND = (landed or {}).get("labels", {}).get("Vendor", "Vendor")
    U = "Union of 2 Sources"
    periods, n_act, sections = plan_model(P)
    grain_is_month = periods[0]["grain"] == "month"
    grain_adj = {"month": "Monthly", "quarter": "Quarterly", "week": "Weekly"}.get(periods[0]["grain"], "Period")
    sec_cols = section_columns(sections)
    has_pct = any(ln["kind"] == "pct_of_total" for ln in P["lines"])
    rate_name = "Monthly Rate" if grain_is_month else "Rate"
    if approved_inputs is not None:
        if has_pct:
            raise ValueError("approved forecast adapter inputs do not support percentage-of-total lines")
        manifest, receipt = approved_inputs["manifest"], approved_inputs["receipt"]
        architecture = manifest.get("planProjection", {}).get("sharedInputArchitecture", {})
        if (P.get("sourceSha256") != manifest.get("sourceHash")
                or architecture.get("presentationInputFeature") != "forecastBusinessInput"
                or architecture.get("targetScope", {}).get("connectionId") != conn):
            raise ValueError("approved forecast adapter source/connection is stale")
        from input_sync import emit_workbook_wiring
        approved_graph = emit_workbook_wiring(manifest, receipt)["wiring"]
        editable = False  # Assigned wide input is the only editable surface.
    cur = CUR(P)
    # With workflows the workbook is a planning app (lib/forecast_app.py): a scenario switch whose
    # value the forecast reads directly (a control's controlId is a formula handle).
    spreads = scenario_spreads(P) if (workflows is not None and approved_inputs is None) else []
    scenario_term = ""
    if spreads:
        scenario_prefix = "FJ" if editable else "Forecast Spine"
        scenario_term = (f" + Switch([{forecast_app.SCENARIO_CONTROL}], "
                         + ", ".join(f'"{name}", [{scenario_prefix}/{col}]' for name, col, _s in spreads) + ", 0)")
    def dm(elid): return {"kind": "data-model", "dataModelId": dm_id, "elementId": els[elid]}
    elements = []
    # actuals wrap (live)
    actuals = {"id": "actuals", "kind": "table", "name": "Actuals", "source": dm("glDetail"),
        "columns": [{"id":"a_sec","name":"Section","formula":f"[{GLP}/Section]"},{"id":"a_so","name":"Section Order","formula":f"[{GLP}/Section Order]"},
                    {"id":"a_line","name":"Line Item","formula":f"[{GLP}/Line Item]"},{"id":"a_lo","name":"Line Order","formula":f"[{GLP}/Line Order]"},
                    {"id":"a_m","name":"Month","formula":f"[{GLP}/Month]"},{"id":"a_mn","name":"Month Num","formula":f"[{GLP}/Month Num]"},
                    {"id":"a_me","name":"Month End","formula":f"[{GLP}/Month End]"},{"id":"a_amt","name":"Amount","formula":f"[{GLP}/Amount]"},
                    {"id":"a_p","name":"Period","formula":'"Actual"'}]}
    spine_names = [("sec","Section"),("so","Section Order"),("line","Line Item"),("lo","Line Order"),("m","Month"),("mn","Month Num"),("me","Month End"),("n","N"),("k","Kind"),("dr","Default Rate"),("mv","Manual Value"),("jb","June Base")]
    if has_pct:
        spine_names.append(("of", "Of Section"))
    spine_names += [(f"sp{i}", col) for i, (_n, col, _s) in enumerate(spreads, 1)]
    spineWrap = {"id":"spineWrap","kind":"table","name":"Forecast Spine","source":dm("spine"),
        "columns":[{"id":f"s_{k}","name":nm,"formula":f"[{DMP}/{nm}]"} for k,nm in spine_names]}
    # forecast element(s). Without percentage-of-total lines there is one "forecast" table;
    # with them, "forecastCalc" computes every other line and "forecast" adds the lines
    # that are a share of another section's forecast total (a second pass over per-month totals).
    stage1_id, stage1_name = ("forecastCalc", "Forecast Base") if has_pct else ("forecast", "Forecast")
    pct_branch = '[Kind] = "pct", 0, ' if has_pct else ""
    if editable:
        # The planning app (workflows) labels the editable tables in sentence case; join keys and the
        # forecast formulas follow the names (a join column is referenced by its display name).
        app = workflows is not None
        RL, RR = ("Line item", "Monthly growth rate") if app else ("Line Item", "Rate")
        rate_title, manual_title = ("Growth rates", "Manual forecast") if app else ("Growth Rates (editable)", "Manual Forecast (editable)")
        rateInput = {"id":"rateInput","kind":"input-table","name":rate_title,"inputMode":"edit","source":{"kind":"empty","connectionId":conn},
            "columns":[{"id":"ID",**({"name":"Row ID","hidden":True} if app else {})},{"id":"LINE_ITEM","type":"text","name":RL},{"id":"RATE","type":"number","name":RR,"format":PCT(P)}],"order":["LINE_ITEM","RATE"],
            **({"tableComponents":{"summaryBar":"hidden"},"tableStyle":dict(APP_INPUT_STYLE)} if app else {})}
        manualInput = {"id":"manualInput","kind":"input-table","name":manual_title,"inputMode":"edit","source":{"kind":"empty","connectionId":conn},
            **({"tableComponents":{"summaryBar":"hidden"},"tableStyle":dict(APP_INPUT_STYLE)} if app else {}),
            "columns":[{"id":"ID",**({"name":"Row ID","hidden":True} if app else {})},{"id":"LINE_ITEM","type":"text","name":RL},{"id":"MONTH","type":"text","name":"Month"},
                       {"id":"MONTH_NUM","type":"number","name":"Month Num",**({"hidden":True} if app else {})},
                       {"id":"MONTH_END","type":"datetime","name":"Month End",**({"hidden":True} if app else {})},
                       {"id":"AMOUNT","type":"number","name":"Amount","format":cur}],"order":["LINE_ITEM","MONTH","AMOUNT"]}
        forecast = {"id":stage1_id,"kind":"table","name":stage1_name,
            "source":{"kind":"join","name":"FJ","primarySource":{"kind":"table","elementId":"spineWrap"},
                "joins":[{"name":"Rates","joinType":"left-outer","left":{"kind":"table","elementId":"spineWrap"},"right":{"kind":"table","elementId":"rateInput"},"columns":[{"left":"[Line Item]","right":f"[{RL}]"}]},
                         {"name":"Manual","joinType":"left-outer","left":{"kind":"table","elementId":"spineWrap"},"right":{"kind":"table","elementId":"manualInput"},"columns":[{"left":"[Line Item]","right":f"[{RL}]"},{"left":"[Month Num]","right":"[Month Num]"}]}]},
            "columns":[{"id":"f_sec","name":"Section","formula":"[FJ/Section]"},{"id":"f_so","name":"Section Order","formula":"[FJ/Section Order]"},
                {"id":"f_line","name":"Line Item","formula":"[FJ/Line Item]"},{"id":"f_lo","name":"Line Order","formula":"[FJ/Line Order]"},
                {"id":"f_m","name":"Month","formula":"[FJ/Month]"},{"id":"f_mn","name":"Month Num","formula":"[FJ/Month Num]"},{"id":"f_me","name":"Month End","formula":"[FJ/Month End]"},
                {"id":"f_n","name":"N","formula":"[FJ/N]"},{"id":"f_k","name":"Kind","formula":"[FJ/Kind]"},{"id":"f_jb","name":"June Base","formula":"[FJ/June Base]"},
                *([{"id":"f_of","name":"Of Section","formula":"[FJ/Of Section]"}] if has_pct else []),
                {"id":"f_er","name":"Eff Rate","formula":f"Coalesce([Rates/{RR}], [FJ/Default Rate])" + scenario_term},
                *([{"id":"f_scn","name":"Active Scenario","formula":f'Coalesce([{forecast_app.SCENARIO_CONTROL}], "{P["scenarios"]["default"]}")'}] if spreads else []),
                {"id":"f_em","name":"Eff Manual","formula":"Coalesce([Manual/Amount], [FJ/Manual Value])"},
                {"id":"f_amt","name":"Amount","formula":'If([Kind] = "manual", [Eff Manual], [Kind] = "flat", [June Base], ' + pct_branch + '[June Base] * (1 + [Eff Rate]) ^ [N])',"format":cur},
                {"id":"f_p","name":"Period","formula":'"Forecast"'}]}
    else:
        forecast = {"id":stage1_id,"kind":"table","name":stage1_name,"source":{"kind":"table","elementId":"spineWrap"},
            "columns":[{"id":"f_sec","name":"Section","formula":"[Forecast Spine/Section]"},{"id":"f_so","name":"Section Order","formula":"[Forecast Spine/Section Order]"},
                {"id":"f_line","name":"Line Item","formula":"[Forecast Spine/Line Item]"},{"id":"f_lo","name":"Line Order","formula":"[Forecast Spine/Line Order]"},
                {"id":"f_m","name":"Month","formula":"[Forecast Spine/Month]"},{"id":"f_mn","name":"Month Num","formula":"[Forecast Spine/Month Num]"},{"id":"f_me","name":"Month End","formula":"[Forecast Spine/Month End]"},
                *([{"id":"f_k","name":"Kind","formula":"[Forecast Spine/Kind]"},{"id":"f_of","name":"Of Section","formula":"[Forecast Spine/Of Section]"},
                   {"id":"f_er","name":"Eff Rate","formula":"[Forecast Spine/Default Rate]" + scenario_term}]
                  if has_pct else []),
                {"id":"f_amt","name":"Amount","formula":'If([Forecast Spine/Kind] = "manual", [Forecast Spine/Manual Value], [Forecast Spine/Kind] = "flat", [Forecast Spine/June Base], ' + ('[Forecast Spine/Kind] = "pct", 0, ' if has_pct else "") + '[Forecast Spine/June Base] * (1 + [Forecast Spine/Default Rate]) ^ [Forecast Spine/N])',"format":cur},
                {"id":"f_p","name":"Period","formula":'"Forecast"'}]}
    if workflows is not None and approved_inputs is None:
        # The planning app: Roll forward moves the actual / forecast split (Parameters.Forecast Start Month).
        ro = None if editable else [
            {"id": "f_k", "name": "Kind", "formula": "[Forecast Spine/Kind]"},
            {"id": "f_jb", "name": "June Base", "formula": "[Forecast Spine/June Base]"},
            {"id": "f_er", "name": "Eff Rate", "formula": "[Forecast Spine/Default Rate]" + scenario_term},
            *([{"id": "f_scn", "name": "Active Scenario",
                "formula": f'Coalesce([{forecast_app.SCENARIO_CONTROL}], "{P["scenarios"]["default"]}")'}]
              if spreads else []),
            {"id": "f_em", "name": "Eff Manual", "formula": "[Forecast Spine/Manual Value]"}]
        boundary = forecast_boundary(P)
        apply_boundary(actuals, forecast, "FJ" if editable else "Forecast Spine", boundary, has_pct, ro)
    forecast_stages = [forecast]
    if has_pct:
        # Second pass without a join: a join to a grouped table also matches that table's
        # detail rows (a section with N lines matches N+1 times and multiplies every share
        # line). "Forecast Base Ref" is a plain copy of the base rows; each share line sums
        # the base section's amounts for its own month with SumIf, one branch per
        # (base section, forecast month) pair.
        ref_name = "Forecast Base Ref"
        forecastRef = {"id":"forecastRef","kind":"table","name":ref_name,"source":{"kind":"table","elementId":stage1_id},
            "columns":[{"id":"r_sec","name":"Section","formula":f"[{stage1_name}/Section]"},
                       {"id":"r_mn","name":"Month Num","formula":f"[{stage1_name}/Month Num]"},
                       {"id":"r_amt","name":"Amount","formula":f"[{stage1_name}/Amount]","format":cur}]}
        pairs = sorted({(ln["of"], v["num"]) for ln in P["lines"] if ln["kind"] == "pct_of_total"
                        for v in calc.month_views(periods)[n_act:]}, key=lambda t: (t[1], t[0]))
        branches = ", ".join(
            f'[{stage1_name}/Of Section] = "{of}" and [{stage1_name}/Month Num] = {num}, '
            f'SumIf([{ref_name}/Amount], [{ref_name}/Section] = "{of}" and [{ref_name}/Month Num] = {num})'
            for of, num in pairs)
        final = {"id":"forecast","kind":"table","name":"Forecast","source":{"kind":"table","elementId":stage1_id},
            "columns":[{"id":"g_sec","name":"Section","formula":f"[{stage1_name}/Section]"},{"id":"g_so","name":"Section Order","formula":f"[{stage1_name}/Section Order]"},
                {"id":"g_line","name":"Line Item","formula":f"[{stage1_name}/Line Item]"},{"id":"g_lo","name":"Line Order","formula":f"[{stage1_name}/Line Order]"},
                {"id":"g_m","name":"Month","formula":f"[{stage1_name}/Month]"},{"id":"g_mn","name":"Month Num","formula":f"[{stage1_name}/Month Num]"},{"id":"g_me","name":"Month End","formula":f"[{stage1_name}/Month End]"},
                {"id":"g_amt","name":"Amount","formula":f'If([{stage1_name}/Kind] = "pct", [{stage1_name}/Eff Rate] * If({branches}, 0), [{stage1_name}/Amount])',"format":cur},
                {"id":"g_p","name":"Period","formula":f"[{stage1_name}/Period]" if workflows is not None else '"Forecast"'}]}
        forecast_stages = [forecast, forecastRef, final]
        forecast = final
    def m(nm): return {"outputColumnName": nm, "sourceColumns": [f"[{nm}]", f"[{nm}]"]}
    pnlCombined = {"id":"pnlCombined","kind":"table","name":C,
        "source":{"kind":"union","sources":[{"kind":"table","elementId":"actuals"},{"kind":"table","elementId":"forecast"}],
                  "matches":[m("Section"),m("Section Order"),m("Line Item"),m("Line Order"),m("Month"),m("Month Num"),m("Month End"),m("Amount"),m("Period")]},
        "columns":[{"id":"c_sec","name":"Section","formula":f"[{U}/Section]"},{"id":"c_so","name":"Section Order","formula":f"[{U}/Section Order]"},
            {"id":"c_line","name":"Line Item","formula":f"[{U}/Line Item]"},{"id":"c_lo","name":"Line Order","formula":f"[{U}/Line Order]"},
            {"id":"c_m","name":"Month","formula":f"[{U}/Month]"},{"id":"c_mn","name":"Month Num","formula":f"[{U}/Month Num]"},{"id":"c_me","name":"Month End","formula":f"[{U}/Month End]"},
            {"id":"c_amt","name":"Amount","formula":f"[{U}/Amount]","format":cur},{"id":"c_p","name":"Period","formula":f"[{U}/Period]"},
            *[{"id":col_id,"name":col_name,"formula":f'If([{U}/Section] = {literal}, [{U}/Amount], 0)'}
              for col_id, col_name, literal in sec_cols.values()]]}
    if approved_inputs is not None:
        approved_nodes = {node["id"]: node for node in approved_graph["elements"]}
        actuals, spineWrap, forecast, pnlCombined = (
            approved_nodes[key] for key in ("actuals", "spineWrap", "forecast", "pnlCombined"))
        forecast_stages = [forecast]
    # When the workbook's FY total covers only some periods, KPI cards cover exactly those
    # periods (a conditional sum on Month Num); trends and the pivot keep every period.
    fy_nums = sorted(p["num"] for p in periods if p["label"] in (P.get("fyPeriods") or []))
    if fy_nums and fy_nums != list(range(fy_nums[0], fy_nums[-1] + 1)):
        raise PlanError("fyPeriods must be a contiguous run of periods")
    fy_prefix = re.sub(r"\s+total$", "", P.get("fyLabel") or "FY", flags=re.I) if fy_nums else ""
    def sec_sum(name, fy=False):
        return section_sum(C, sec_cols, fy_nums, name, fy)
    def kpi_title(name):
        if not fy_nums:
            return name
        return f"{fy_prefix} {name}"
    def signed(row, fy=False):
        return amount_formula(row, C, sec_cols, fy_nums, fy)
    rows = amount_rows(P, sections)
    first_positive = next((sec["name"] for sec in sections if sec["sign"] > 0), sections[0]["name"])
    REV = sec_sum(first_positive)
    REV_FY = sec_sum(first_positive, True)
    rev_label = "Total Revenue" if first_positive == "Revenue" else f"Total {first_positive}"
    card_style = {"backgroundColor":"#FFFFFF","borderColor":{"kind":"theme","ref":"colors-border"},"borderWidth":1,"borderRadius":"round"}
    def kpi(eid,name,f,fmt): return {"id":eid,"kind":"kpi-chart","name":name,"source":{"kind":"table","elementId":"pnlCombined"},"columns":[{"id":"v","formula":f,"format":fmt}],"value":{"columnId":"v"},"style":dict(card_style)}
    # KPI cards: first positive section, the next-to-last and last amount rows, and the
    # first amount row as a share of that section (Gross Margin %).
    middle, bottom = kpi_rows(P, rows)
    slots = [("kpiRev", kpi_title(rev_label), REV_FY, cur)]
    if middle is not None:
        slots.append(("kpiEbitda", kpi_title(rows[middle]["label"]), signed(rows[middle], True), cur))
    slots.append(("kpiNI", kpi_title(rows[bottom]["label"]), signed(rows[bottom], True), cur))
    if len(rows) >= 2:
        first = rows[0]
        ratio_name = "Gross Margin %" if first["label"].lower() == "gross profit" else f"{first['label']} %"
        slots.append(("kpiGM", kpi_title(ratio_name), f"({signed(first, True)}) / {REV_FY}", PCT(P)))
    kpis = [kpi(*slot) for slot in slots]
    kpis[0]["style"]["borderColor"] = "#2563EB"
    header={"id":"pnlHeader","kind":"text","body":f"# {P['name']}\n{grain_adj} actuals and forecast performance",
        "style":{"backgroundColor":"#F8FAFC","color":"#0F172A","borderRadius":"round"}}
    tail = [rows[i] for i in sorted({middle, bottom} - {None})]
    trend_series = [(first_positive, REV)] + [(r["label"], signed(r)) for r in tail]
    trend_names = [name for name, _ in trend_series]
    trend_label = ", ".join(trend_names[:-1]) + " & " + trend_names[-1] if len(trend_names) > 1 else trend_names[0]
    series_cols = [{"id":"t_" + re.sub(r"\W+", "_", name.lower()).strip("_"),"name":name,"formula":f,"format":cur}
                   for name, f in trend_series]
    if grain_is_month:
        x_cols = [{"id":"t_month","name":"Month","formula":f'DateTrunc("month", [{C}/Month End])',"format":DATEF}]
        x_axis = {"columnId":"t_month","sort":{"direction":"ascending"}}
    else:
        x_cols = [{"id":"t_month","name":"Period","formula":f"[{C}/Month]"},{"id":"t_ord","name":"Period Order","formula":f"[{C}/Month Num]"}]
        x_axis = {"columnId":"t_month","sort":{"direction":"ascending","by":"t_ord","aggregation":"min"}}
    trend={"id":"monthlyTrend","kind":"line-chart","name":f"{grain_adj} {trend_label}",
        "source":{"kind":"table","elementId":"pnlCombined"},
        "columns":x_cols + series_cols,
        "xAxis":x_axis,
        "yAxis":{"columnIds":[c["id"] for c in series_cols],"format":{"marks":"grid","scale":{"type":"linear","zero":False}}},
        "legend":{"position":"bottom"},"lineAreaStyle":{"interpolation":"monotone"},
        "style":{"backgroundColor":"#FFFFFF","borderColor":{"kind":"theme","ref":"colors-border"},"borderWidth":1,"borderRadius":"round"}}
    if grain_is_month:
        p_period = [{"id":"p_m","name":"Month","formula":f"[{C}/Month End]","format":DATEF}]
        p_cols_by = [{"columnId":"p_per","sort":{"direction":"ascending"}},{"columnId":"p_m","sort":{"direction":"ascending"}}]
    else:
        p_period = [{"id":"p_m","name":"Period Label","formula":f"[{C}/Month]"},{"id":"p_mn","name":"Period Order","formula":f"[{C}/Month Num]"}]
        p_cols_by = [{"columnId":"p_per","sort":{"direction":"ascending"}},{"columnId":"p_m","sort":{"direction":"ascending","by":"p_mn","aggregation":"min"}}]
    pivot={"id":"pnlPivot","kind":"pivot-table","name":"P&L Statement","source":{"kind":"table","elementId":"pnlCombined"},
        "columns":[{"id":"p_sec","name":"Section","formula":f"[{C}/Section]"},{"id":"p_so","name":"Section Order","formula":f"[{C}/Section Order]"},
            {"id":"p_line","name":"Line Item","formula":f"[{C}/Line Item]"},{"id":"p_lo","name":"Line Order","formula":f"[{C}/Line Order]"},
            {"id":"p_per","name":"Period","formula":f"[{C}/Period]"},*p_period,
            {"id":"p_amt","name":"Amount","formula":f"Sum([{C}/Amount])","format":cur}],
        "values":["p_amt"],
        "rowsBy":[{"columnId":"p_sec","sort":{"direction":"ascending","by":"p_so","aggregation":"min"}},{"columnId":"p_line","sort":{"direction":"ascending","by":"p_lo","aggregation":"min"}}],
        "columnsBy":p_cols_by,
        "axisTotals":[{"axisId":"grand-row-total","visibility":"hidden"}],
        "style":{"backgroundColor":"#FFFFFF","borderColor":{"kind":"theme","ref":"colors-border"},"borderWidth":1,"borderRadius":"round"}}
    glTable={"id":"glTable","kind":"table","name":"GL Transaction Detail (Actuals)","source":dm("glDetail"),
        "columns":[{"id":"g1","name":"Txn Date","formula":f"[{GLP}/Txn Date]"},{"id":"g2","name":"Month","formula":f"[{GLP}/Month]"},{"id":"g3","name":"Line Item","formula":f"[{GLP}/Line Item]"},{"id":"g4","name":DEPT,"formula":f"[{GLP}/{DEPT}]"},{"id":"g5","name":VEND,"formula":f"[{GLP}/{VEND}]"},{"id":"g6","name":"Amount","formula":f"[{GLP}/Amount]","format":cur}]}
    assumpTable={"id":"assumpTable","kind":"table","name":"Forecast Assumptions","source":dm("assumptions"),
        "columns":[{"id":"d1","name":"Driver","formula":f"[{DMP}/Driver]"},{"id":"d2","name":rate_name,"formula":f"[{DMP}/{rate_name}]","format":PCT(P)},{"id":"d3","name":"Basis","formula":f"[{DMP}/Basis]"}]}
    # Source-workbook charts the drafter mapped and verified (plan.sourceCharts). Series are
    # named like the plan's section totals / amount rows (classic rows use Revenue, Gross
    # Profit, EBITDA, Net Income).
    metric_formula = {}
    for sec in sections:
        label = next((m_["totalLabel"] for m_ in P.get("sections", []) if m_["name"] == sec["name"] and m_.get("totalLabel")), None)
        if label:
            metric_formula[calc.TARGET_LABELS.get(label.lower(), label)] = sec_sum(sec["name"])
    for row in rows:
        metric_formula[calc.TARGET_LABELS.get(row["label"].lower(), row["label"])] = signed(row)
    if not P.get("sections"):
        metric_formula["Revenue"] = REV
    source_charts = []
    for i, sc in enumerate(P.get("sourceCharts") or [], 1):
        if sc["grain"] == "quarter":
            x = {"id": f"s{i}_x", "name": "Quarter",
                 "formula": f'Concat(Text(DatePart("year", [{C}/Month End])), " Q", Text(DatePart("quarter", [{C}/Month End])))'}
        elif sc["grain"] == "period":
            x = {"id": f"s{i}_x", "name": "Period", "formula": f"[{C}/Month]"}
        else:
            x = {"id": f"s{i}_x", "name": "Month", "formula": f'DateTrunc("month", [{C}/Month End])', "format": DATEF}
        ys = [{"id": f"s{i}_y{j}", "name": metric, "formula": metric_formula[metric], "format": cur}
              for j, metric in enumerate(sc["series"], 1)]
        is_bar = sc.get("kind") == "bar"
        if sc["grain"] == "period":
            # Plan periods are labels, not dates: order them by their sequence number.
            extra = {"id": f"s{i}_o", "name": "Period Order", "formula": f"[{C}/Month Num]"}
            filters, x_sort = {}, {"direction": "ascending", "by": extra["id"], "aggregation": "min"}
        else:
            # Keep the source chart's period range (a Jan-Mar chart stays Jan-Mar).
            extra = {"id": f"s{i}_d", "name": "Period Date", "formula": f"[{C}/Month End]"}
            filters = {"filters": [{"id": f"s{i}_range", "columnId": extra["id"], "kind": "date-range",
                                    "mode": "between", "startDate": sc["periodStart"], "endDate": sc["periodEnd"]}]}
            x_sort = {"direction": "ascending"}
        chart = {"id": f"sourceChart{i}", "kind": "bar-chart" if is_bar else "line-chart", "name": sc["title"],
            "source": {"kind": "table", "elementId": "pnlCombined"}, "columns": [x] + ys + [extra],
            **filters,
            "xAxis": {"columnId": x["id"], "sort": x_sort},
            "yAxis": {"columnIds": [y["id"] for y in ys], "format": {"marks": "grid", "scale": {"type": "linear", "zero": is_bar}}},
            "legend": {"position": "bottom"},
            "style": {"backgroundColor": "#FFFFFF", "borderColor": {"kind": "theme", "ref": "colors-border"}, "borderWidth": 1, "borderRadius": "round"}}
        source_charts.append(chart)
    wf = None
    if workflows is not None:
        # Snapshot freezes the KPI tiles' own formulas (revenue, first amount row, middle amount row).
        figures = {"revenue": REV_FY}
        by_label = {r["label"].lower(): r for r in rows}
        gm = by_label.get("gross margin") or by_label.get("gross profit") or (rows[0] if len(rows) >= 2 else None)
        ebitda = by_label.get("ebitda") or (rows[bottom] if len(rows) >= 2 else None)
        if gm is not None:
            figures["grossMargin"] = signed(gm, True)
        if ebitda is not None:
            figures["ebitda"] = signed(ebitda, True)
        wf = demo_workflows.workflow_elements(
            workflows, conn, f"{P.get('fyLabel') or 'FY'} forecast", figures,
            scenario_control=forecast_app.SCENARIO_CONTROL if spreads else None, no_status="Draft",
            start=forecast_app.START_MONTH if approved_inputs is None else demo_workflows.DEFAULT_START)
        return app_spec(P, dm, els, wf=wf, rows=rows, kpi={"middle": middle, "bottom": bottom},
                        first_positive=first_positive, fy_nums=fy_nums, periods=periods, n_act=n_act,
                        hidden=[actuals, spineWrap, *forecast_stages, pnlCombined], pivot=pivot, gl=glTable,
                        assump=assumpTable, rate=rateInput if editable else None,
                        manual=manualInput if editable else None, plan_src=plan_src, spreads=spreads,
                        cur=cur, landed=landed, history=history)
    main_els=[header]+([wf["tables"][0], wf["status"], *wf["buttons"]] if wf else [])+kpis+[trend]+source_charts+[pivot]
    detail_els=[glTable]
    input_els=[assumpTable]
    hidden_els=[actuals,spineWrap,*forecast_stages,pnlCombined]
    extra_model_ids = [e["id"] for e in forecast_stages[:-1]] if has_pct else []
    if approved_inputs is not None:
        extra = [e for e in approved_graph["elements"] if e["id"] not in {n["id"] for n in hidden_els}]
        hidden_els += extra
        extra_model_ids = [e["id"] for e in extra]
    if editable: input_els+=[rateInput,manualInput]
    if wf: input_els+=wf["tables"][1:]
    pages=[{"id":"pnl","name":"P&L Overview","elements":main_els},
           {"id":"detail","name":"Detail","elements":detail_els},
           {"id":"inputs","name":"Inputs","elements":input_els},
           {"id":"model","name":"Model (sources)","visibility":"hidden","elements":hidden_els}]
    # Nested draft; wire_workbook flattens + wraps at the POST / dry-run boundary.
    return {"name":P["name"],"schemaVersion":1,"kind":"workbook","folderId":P["folderId"],
            "pages":pages,"layout":forecast_layout(editable, extra_model_ids, [c["id"] for c in source_charts], [k["id"] for k in kpis], bool(wf))}

def _app_titles(P):
    """``(company, title)`` for the app header: ``Vantora FY27 Rolling Forecast — run tag`` ->
    ``("Vantora", "FY27 rolling forecast")``; the run tag stays in the document name only."""
    base = forecast_app.display_name(re.split(r"\s+[—–-]\s+", P["name"], maxsplit=1)[0].strip())
    fy = P.get("fyLabel") or ""
    fy = re.sub(r"\s+total$", "", fy, flags=re.I)
    if fy and f" {fy} " in f" {base} ":
        company, rest = base.split(fy, 1)
        return company.strip(), forecast_app.sentence((fy + rest).strip())
    return "", forecast_app.sentence(base)


def app_spec(P, dm, els, *, wf, rows, kpi, first_positive, fy_nums, periods, n_act, hidden, pivot, gl, assump,
             rate, manual, plan_src, spreads, cur, landed, history=None):
    """The planning-app workbook (lib/forecast_app.py) for a forecast built with --workflows."""
    by_num = {p["num"]: p for p in periods}
    fy_label = re.sub(r"\s+total$", "", P.get("fyLabel") or "FY", flags=re.I)
    prior_nums = [n - len(fy_nums) for n in fy_nums] if fy_nums else []
    if not prior_nums or not all(n in by_num and by_num[n].get("role", "actual") == "actual"
                                 and n <= periods[n_act - 1]["num"] for n in prior_nums):
        prior_nums = []
    digits = re.search(r"(\d+)$", fy_label)
    prior_label = (fy_label[:digits.start()] + str(int(digits.group(1)) - 1).zfill(len(digits.group(1)))
                   if digits else "Prior year")
    fy_last = by_num[fy_nums[-1]] if fy_nums else periods[-1]
    end = fy_last.get("end") or ""
    if not end:
        raise PlanError("the planning-app layout needs dated periods (month grain)")
    company, title = _app_titles(P)
    # Display-ready copies of the builder's own detail elements.
    # The monthly P&L shows the FY's months (actual | forecast) in $000s so every column is readable; the
    # cells carry the K, so the title does not repeat the unit.
    fy_first = by_num[fy_nums[0]] if fy_nums else periods[0]
    fy_start = (fy_first.get("end") or "")[:8] + "01"
    thousands = {"kind": "number", "formatString": ",.0f", "prefix": "$", "suffix": "K"}
    pivot = {**pivot, "name": f"Monthly P&L — {fy_label}", "tableStyle": {"preset": "presentation"},
             "columns": [{**c, "formula": c["formula"] + " / 1000", "format": thousands} if c["id"] == "p_amt" else c
                         for c in pivot["columns"]],
             "filters": [{"id": "p_fy", "columnId": "p_m", "kind": "date-range", "mode": "between",
                          "startDate": fy_start, "endDate": end}]}
    # no row count in the title: a month close lands more rows than the source had
    gl = {**gl, "name": "GL actuals",
          "tableComponents": {"summaryBar": "hidden"}, "tableStyle": {"preset": "presentation"},
          # a dimension the landed table does not carry (an empty Department / Vendor) is not shown
          "columns": [{**c, "name": forecast_app.sentence(c["name"])} for c in gl["columns"]
                      if c["name"] != "Txn Date" and c["formula"] not in ("[GL Detail/Department]", "[GL Detail/Vendor]")]
          # the one boundary: a landed month that is still a forecast month (before roll forward) is not
          # listed as an actual; the same hidden window the Actuals wrap filters on
          + [{"id": "g_mk", "name": "Month key", "hidden": True,
              "formula": month_key(f"[{actuals_prefix(gl)}/Month End]")},
             {"id": "g_win", "name": "Window", "hidden": True,
              "formula": f'If([Month key] < {forecast_app.START_KEY}, "Actual", "Forecast")'}],
          "filters": [{"id": "g_wf", "columnId": "g_win", "kind": "list", "mode": "include", "values": ["Actual"]}]}
    ctx = {"P": P, "wf": wf, "rows": rows, "kpi": kpi, "first_positive": first_positive, "fy_nums": fy_nums,
           "prior_nums": prior_nums, "fy_label": fy_label, "prior_label": prior_label,
           "fy_end": (int(end[:4]), int(end[5:7])), "n_fy": len(fy_nums) or len(periods),
           "hidden": hidden, "pivot": pivot, "gl": gl, "assump": assump, "rate": rate, "manual": manual,
           "dm": dm, "plan": plan_src, "region_of": (plan_src or {}).get("regions") or {},
           "scenario": ({"names": P["scenarios"]["names"], "default": P["scenarios"]["default"]} if spreads else None),
           "title": title, "company": company,
           # without a Parameters table (no set-scenario / roll-forward) the boundary is the source's own
           "boundary": forecast_boundary(P, parameters=any(t["id"] == demo_workflows.PARAMETERS
                                                           for t in wf["tables"]))["element"],
           "landed": landed is not None,
           "cost_lines": cost_lines(P),
           # migrated log rows per table (the setup program seeds them from the source's own logs)
           "history": {el: len(rows) for el, _cols, rows in (history or []) if rows}}
    app = forecast_app.build(ctx)
    return {"name": P["name"], "schemaVersion": 1, "kind": "workbook", "folderId": P["folderId"],
            "pages": app["pages"], "layout": app["layout"], "settings": app["settings"]}


# What a failed agent recommendation / merge raises (duplicate chat element, layout without the
# anchor page, malformed macros or workflows): reported as AGENT ERROR, never a traceback.
AGENT_ERRORS = (ValueError, KeyError, TypeError, AttributeError)


def driver_example(P):
    """The first growth-rate line of the plan (the set-driver tool's example line item), or None."""
    rate, _manual = seed_rows(P)
    return rate[0][0] if rate else None


def merge_agent(wb_spec, macros, workflows, name="FP&A forecast agent", quiet=False, driver_example=None):
    """Phase 6 in the same run: the recommended workbook agent + chat merged into the spec (unchanged
    when no macros were given or no agent is recommended). Prints the on-camera pitch unless quiet.
    ``driver_example`` is a real line item of the rate table, quoted in the set-driver tool's description."""
    if macros is None:
        return wb_spec
    rec, fragment = agent_enhance.recommend_for_workbook(macros, workflows, wb_spec, name,
                                                         driver_example=driver_example)
    if not quiet:
        print("AGENT: " + (rec["pitch"] if rec["recommend"] else "no agent recommended"))
    return agent_enhance.merge_fragment(wb_spec, fragment) if fragment else wb_spec


def setup_program(P, wb_spec, workflows, history=None):
    """``(program text, seeded summary)`` of the post-create setup program for a wired workbook spec, or
    ``None`` when the workbook has no input tables (a read-only plan without ``--workflows``): manual
    commit on every input table, then the seeds. The one-row Parameters seed is emitted only when the
    workbook has the Parameters table (``--workflows`` with a workflow that needs it); the rate / manual
    input tables are seeded from ``seed_rows`` (the seed CSVs' rows) whenever the plan is editable or
    what-if; ``history`` (demo_workflows.source_history) seeds the Versions / Approvals logs with the
    source's own. Raises ``demo_workflows.WorkflowError`` when the program cannot be built."""
    tables = {el for _pg, el in demo_workflows._input_tables(wb_spec)}
    if not tables:
        return None
    seed = None
    if workflows is not None and demo_workflows.PARAMETERS in tables:
        periods, n_act, _sections = plan_model(P)
        seed = {"start": periods[n_act]["end"][:8] + "01"}
        has_scenario_control = any(
            el.get("kind") == "control" and el.get("controlId") == forecast_app.SCENARIO_CONTROL
            for el in code_rep.workbook_elements(wb_spec))
        if not has_scenario_control:
            seed["scenario"] = "Base"   # no scenario switch: the Parameters row holds the scenario
    table_seeds = []
    if P["intent"] in ("editable", "what-if"):
        rate, manual = seed_rows(P)
        table_seeds = [("rateInput", RATE_SEED_COLUMNS, rate), ("manualInput", MANUAL_SEED_COLUMNS, manual)]
    if workflows is not None:
        table_seeds += [h for h in (history or []) if h[0] in tables]   # only the logs this workbook has
    program = demo_workflows.emit_manual_commit_program(
        wb_spec, seed, table_seeds, series_colors=forecast_app.SERIES_COLORS,
        column_widths=demo_workflows.input_column_widths(wb_spec, table_seeds))
    seeded = (["parameters 1 row"] if seed else []) + [f"{el} {len(rows)} rows" for el, _c, rows in table_seeds if rows]
    return program, seeded


def emit_workflow_outputs(P, wb_spec, outdir, workflows, program_path, history=None):
    """Write the setup program (``setup_program``) to ``program_path`` when asked for; returns the path
    written, or ``None`` (not asked for, or skipped with a printed note: no input tables)."""
    if not program_path:
        return None
    built = setup_program(P, wb_spec, workflows, history)
    if built is None:
        print(f"SETUP PROGRAM: skipped — the workbook has no input tables (a {P['intent']} plan without "
              f"--workflows), so there is nothing to enable or seed; {program_path} was not written")
        return None
    program, seeded = built
    with open(program_path, "w", encoding="utf-8") as stream:
        stream.write(program)
    print(f"SETUP PROGRAM: {program_path} (manual commit on every input table"
          + ("; seeds " + ", ".join(seeded) if seeded else "") + "; run once on the new workbook)")
    return program_path


LOCAL_DM_IDS = {"glDetail": "glDetail", "spine": "spine", "assumptions": "assumptions",
                "planDetail": "planDetail", "metricMap": "metricMap"}


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", help="model-plan JSON")
    parser.add_argument("--source", help="source XLSX used to produce the plan")
    parser.add_argument("--review-dispositions", help="source-bound review dispositions JSON")
    parser.add_argument("--post", action="store_true", help="create the data model and workbook")
    sigma_auth.add_cli_args(parser)
    parser.add_argument("--out-dir", default="/tmp/fcst-build",
                        help="where specs and seed CSVs are written (default: /tmp/fcst-build)")
    parser.add_argument("--actuals-table", help="DB.SCHEMA.TABLE landed by land-in-snowflake.py; GL actuals "
                        "become a warehouse-table source instead of inline VALUES (needs --landing)")
    parser.add_argument("--landing", help="landing.json written by land-in-snowflake.py")
    parser.add_argument("--workflows", help="workflows.json from macro-actions.py demo-manifest: adds the "
                        "Parameters / Versions / Approvals input tables, the Roll forward / Snapshot version / "
                        "Submit for approval buttons and a latest-approval-status KPI")
    parser.add_argument("--approvers", metavar="EMAILS", help="with --workflows: comma-separated emails allowed to "
                        "Approve / Send back a Submitted request. Baked in at build time: the buttons' row "
                        "filter only matches when the signed-in user's email (case-insensitive) is listed. "
                        "Without it (or without a Submit workflow) the Approve / Send back buttons are omitted "
                        "and the workflows manifest says so")
    parser.add_argument("--emit-manual-commit-program", metavar="OUT.py",
                        help="write the post-create Code Mode setup program: enables manual commit on every "
                             "input table (needed before backend actions / agent tools can write), then seeds "
                             "the rate / manual input tables from the seed CSV rows (editable / what-if plans) "
                             "and, with --workflows, the one Parameters row; run it once on a fresh workbook")
    parser.add_argument("--staging-approval", metavar="APPROVAL.json",
                        help="with --post: build in an isolated staging folder while source review issues are "
                             "still pending, under a named staging approval (owner, reviewer, folderId, "
                             "sourceSha256, sourceReviewApproval {status: approved_for_staging, issues: the exact "
                             "audited issue objects}). Never asserts verified dispositions")
    parser.add_argument("--staging-approval-template", metavar="OUT.json",
                        help="write the staging approval to review (exact audited issues, blank owner/reviewer, "
                             "status pending) and stop")
    parser.add_argument("--owner", help="with --staging-approval-template: the named owner of the staging build")
    parser.add_argument("--reviewer", help="with --staging-approval-template: the named reviewer who checked "
                        "every audited issue")
    parser.add_argument("--approve-for-staging", action="store_true",
                        help="with --staging-approval-template, --owner and --reviewer: the named reviewer "
                             "runs this in their own terminal; it lists every audited issue and waits for "
                             "them to type APPROVE, then records status approved_for_staging. It refuses "
                             "without an interactive terminal. Policy: the approval must be given by the "
                             "named human reviewer at their own terminal; the agent must not approve. "
                             "Never marks an issue verified")
    parser.add_argument("--agent-macros", metavar="MACROS.json",
                        help="with --workflows: merge the Phase 6 workbook agent (recommend-agent.py) into the "
                             "workbook before it is validated and created, so one POST carries it")
    parser.add_argument("--parity-tolerance", type=float, default=None,
                        help="absolute tolerance for plan.parityTargets (default: float noise only, "
                             "max(1e-6, 1e-9 * |source value|))")
    return parser.parse_args(argv)


class StagingError(ValueError):
    """A staging approval that does not cover the exact audited source issues and folder."""


APPROVAL_WORD = "APPROVE"
TERMINAL_CONFIRMATION = "reviewer typed APPROVE at an interactive terminal"


def staging_template(P, report, owner="", reviewer="", approve=False):
    """The staging approval for these source bytes and the plan's folder. ``approve`` records the named
    owner's and reviewer's approval of exactly the audited issues (only after
    ``confirm_staging_approval``); it never marks anything verified."""
    if approve and not (str(owner or "").strip() and str(reviewer or "").strip()):
        raise StagingError("--approve-for-staging needs --owner and --reviewer")
    template = {"scope": "staging build only; production --post keeps the verified-disposition gate",
                "folderId": P.get("folderId"), "sourceSha256": report["sourceSha256"],
                "owner": owner or "", "reviewer": reviewer or "",
                "sourceReviewApproval": {"status": "approved_for_staging" if approve else "pending",
                                         "issues": report["issues"]}}
    if approve:
        template["sourceReviewApproval"]["confirmation"] = TERMINAL_CONFIRMATION
    return template


def confirm_staging_approval(P, report, owner, reviewer, stdin=None, prompt=None):
    """The named human reviewer's step: at their own interactive terminal they read every audited issue and
    type APPROVE (policy: the agent must not approve). Refuses when stdin is not a terminal (an agent's
    shell, a pipe, CI), so a copy-pasted command does not record an approval by itself."""
    stdin = stdin or sys.stdin
    prompt = prompt or sys.stderr
    if not stdin.isatty():
        raise StagingError("--approve-for-staging needs the reviewer at an interactive terminal (stdin is not "
                           "a terminal); the named human reviewer runs this command in their own terminal, "
                           "after reading every audited issue — the agent must not approve for them")
    lines = [f"Staging approval for folder {P.get('folderId')} — owner {owner}, reviewer {reviewer}.",
             f"{len(report['issues'])} audited issue(s) stay pending (nothing is marked verified):"]
    lines += [f"  - {i.get('code', '?')} [{i.get('severity', '?')}]: {i.get('message', '')}"
              for i in report["issues"]]
    lines.append(f"{reviewer}: type {APPROVAL_WORD} to approve exactly these issues for this staging folder: ")
    prompt.write("\n".join(lines))
    prompt.flush()
    answer = stdin.readline().strip()
    if answer != APPROVAL_WORD:
        raise StagingError(f"not approved (expected {APPROVAL_WORD!r}); nothing was written")


def check_staging_approval(P, report, approval):
    """The approval must name an owner and reviewer, carry the terminal confirmation that
    ``--approve-for-staging`` records, approve exactly the audited issues of these source bytes for staging,
    and be bound to the plan's folder. It never marks anything verified."""
    if not str(approval.get("owner") or "").strip() or not str(approval.get("reviewer") or "").strip():
        raise StagingError("staging approval needs a named owner and reviewer")
    review = approval.get("sourceReviewApproval") or {}
    if review.get("status") != "approved_for_staging":
        raise StagingError("sourceReviewApproval.status must be approved_for_staging")
    if review.get("confirmation") != TERMINAL_CONFIRMATION:
        raise StagingError("sourceReviewApproval.confirmation is missing or not the terminal confirmation: the "
                           "approval must be recorded by --approve-for-staging, run by the named reviewer at "
                           "their own terminal (a hand-written approval is not accepted)")
    if approval.get("sourceSha256") != report["sourceSha256"]:
        raise StagingError("staging approval is for different source bytes; re-approve")
    if approval.get("folderId") != P.get("folderId"):
        raise StagingError("staging approval is bound to a different folder than the plan's")
    want = sorted(source_issue_hash(i) for i in report["issues"])
    got = sorted(source_issue_hash(i) for i in review.get("issues") or [])
    if want != got:
        raise StagingError("approved issues are not exactly the audited issues; re-approve")


def main(argv=None):
    args = parse_args(argv)
    with open(args.plan, encoding="utf-8") as stream:
        P = json.load(stream)
    staging = bool(args.staging_approval or args.staging_approval_template)
    if args.staging_approval and not args.post:
        print("STAGING: --staging-approval goes with --post", file=sys.stderr)
        return 1
    if (args.owner or args.reviewer or args.approve_for_staging) and not args.staging_approval_template:
        print("STAGING: --owner / --reviewer / --approve-for-staging go with --staging-approval-template",
              file=sys.stderr)
        return 1
    if args.approve_for_staging and not (str(args.owner or "").strip() and str(args.reviewer or "").strip()):
        print("STAGING: --approve-for-staging needs --owner and --reviewer (the named people who checked "
              "every audited issue)", file=sys.stderr)
        return 1
    try:
        if staging:
            source_path, gate_report = gate_plan_staging_source(P, source_arg=args.source, plan_path=args.plan)
        else:
            source_path, gate_report = gate_plan_source(
                P,
                source_arg=args.source,
                dispositions_path=args.review_dispositions,
                plan_path=args.plan,
                draft=not args.post,
            )
    except MigrationGateError as error:
        print(f"MIGRATION GATE: {error}", file=sys.stderr)
        return error.exit_code
    if args.staging_approval_template:
        if args.approve_for_staging:
            try:
                confirm_staging_approval(P, gate_report, args.owner, args.reviewer)
            except StagingError as error:
                print(f"STAGING: {error}", file=sys.stderr)
                return 1
        template = staging_template(P, gate_report, args.owner, args.reviewer, args.approve_for_staging)
        with open(args.staging_approval_template, "w", encoding="utf-8") as stream:
            json.dump(template, stream, indent=1)
        codes = ", ".join(i.get("code", "?") for i in gate_report["issues"]) or "none"
        if args.approve_for_staging:
            print(f"STAGING APPROVAL: {args.staging_approval_template} — APPROVED FOR STAGING by owner "
                  f"{args.owner} and reviewer {args.reviewer} for folder {P.get('folderId')}; issues stay "
                  f"pending (not verified): {codes}. Post with --post --staging-approval "
                  f"{args.staging_approval_template}")
        else:
            print(f"STAGING APPROVAL TEMPLATE: {args.staging_approval_template} — issues {codes}; NOT approved. "
                  "The named reviewer reviews every issue and, in their own terminal, re-runs this command "
                  "with --owner <name> --reviewer <name> --approve-for-staging (it waits for them to type "
                  "APPROVE; the agent must not answer it). Then post with --post --staging-approval")
        return 0
    if args.staging_approval:
        try:
            with open(args.staging_approval, encoding="utf-8") as stream:
                check_staging_approval(P, gate_report, json.load(stream))
        except (OSError, ValueError) as error:
            print(f"STAGING GATE: {error}", file=sys.stderr)
            return 1
        print("STAGING GATE: approved_for_staging by the named owner/reviewer for folder "
              f"{P.get('folderId')}; pending issues stay pending: "
              + ", ".join(i.get("code", "?") for i in gate_report["issues"]))
    agent_macros = None
    if args.agent_macros:
        if not args.workflows:
            print("PLAN ERROR: --agent-macros needs --workflows", file=sys.stderr)
            return 1
        with open(args.agent_macros, encoding="utf-8") as stream:
            agent_macros = json.load(stream)
    print(
        f"MIGRATION GATE: {gate_report['status']} "
        f"source={source_path.name} sha256={gate_report['sourceSha256']}"
    )
    if gate_report.get("pendingReview"):
        print("DRAFT MODE: review pending for " + ", ".join(gate_report["pendingReview"])
              + " — fine for a local build; resolve before --post.")
    do_post = args.post
    outdir = args.out_dir
    landed = None
    if bool(args.actuals_table) != bool(args.landing):
        print("PLAN ERROR: --actuals-table and --landing go together", file=sys.stderr)
        return 1
    plan_src = None
    if args.actuals_table:
        try:
            with open(args.landing, encoding="utf-8") as stream:
                landing = json.load(stream)
            landed = landed_actuals_source(landing, args.actuals_table, P)
            # Only the planning app (--workflows) compares to plan; a plain run never reads the plan table.
            plan_src = landed_plan_source(landing, P) if args.workflows else None
        except (PlanError, OSError, ValueError) as error:
            print(f"PLAN ERROR: {error}", file=sys.stderr)
            return 1
        print("ACTUALS SOURCE: warehouse table " + ".".join(landed["path"]))
        if plan_src:
            print("PLAN SOURCE: warehouse table " + ".".join(plan_src["path"]))
    workflows = None
    if args.workflows:
        try:
            with open(args.workflows, encoding="utf-8") as stream:
                workflows = json.load(stream)
            workflows = demo_workflows.with_approvers(workflows, args.approvers)
            demo_workflows.workflow_elements(workflows, P["connectionId"], "check")
        except (OSError, ValueError, KeyError) as error:
            print(f"PLAN ERROR: --workflows {error}", file=sys.stderr)
            return 1
        print("WORKFLOWS: " + ", ".join(w["actionId"] for w in workflows["workflows"]))
        if workflows.get("approvers"):
            print("APPROVERS: " + ", ".join(workflows["approvers"]) + " (Approve / Send back on Versions & approvals)")
        else:
            print("APPROVERS: none; " + next((d["note"] for d in workflows.get("dispositions", [])
                                                if d.get("disposition") == "approve-omitted-no-approvers"),
                                               "no Submit workflow, so no Approve / Send back"))
    history = []
    if workflows is not None:
        try:   # the source's own version / approval logs seed Versions / Approvals (migrated history)
            history = demo_workflows.source_history(source_path)
        except (OSError, ValueError) as error:
            print(f"PLAN ERROR: source history {error}", file=sys.stderr)
            return 1
        if history:
            print("HISTORY: " + ", ".join(f"{el} {len(rows)} rows" for el, _c, rows in history)
                  + " from the source workbook's own logs (seeded by the setup program)")
    # parity — a mismatch stops the run before any spec is written or posted
    try:
        par=parity(P)
    except PlanError as error:
        print(f"PLAN ERROR: {error}", file=sys.stderr)
        return 1
    shown = {k: v for k, v in par.items() if k in (P.get("parityTargets") or {})} or par
    print("LOCAL PARITY (from plan):", "  ".join(f"{k}={v:,.2f}" for k,v in shown.items()))
    if not check_parity(P, par, args.parity_tolerance):
        print("LOCAL PARITY FAILED — no specs written; fix the plan (or the source) before building.",
              file=sys.stderr)
        return 1
    try:   # a plan the workbook cannot express fails before anything is written or posted
        preflight = build_wb_spec(P,"<dm-id>",LOCAL_DM_IDS,landed=landed,workflows=workflows,plan_src=plan_src,history=history)
    except PlanError as error:
        print(f"PLAN ERROR: {error}", file=sys.stderr)
        return 1
    wired_preflight = workbook_wire.wire_workbook(preflight)
    if args.emit_manual_commit_program:
        try:   # ... and so does a setup program that cannot be built (same input tables, local ids)
            setup_program(P, wired_preflight, workflows, history)
        except demo_workflows.WorkflowError as error:
            print(f"PLAN ERROR: --emit-manual-commit-program {error} — nothing was written or posted",
                  file=sys.stderr)
            return 1
    try:   # ... and so does a Phase 6 agent merge that cannot apply (same spec shape, local ids)
        merge_agent(wired_preflight, agent_macros, workflows, quiet=True, driver_example=driver_example(P))
    except AGENT_ERRORS as error:
        print(f"AGENT ERROR: {error} — nothing was written or posted", file=sys.stderr)
        return 1
    os.makedirs(outdir, exist_ok=True)
    rate_csv,man_csv=seed_csvs(P,outdir)
    print(f"seed CSVs: {rate_csv}, {man_csv}")
    dm_spec=build_dm_spec(P,landed,plan_src,app=workflows is not None)
    with open(f"{outdir}/dm_spec.json", "w", encoding="utf-8") as stream:
        json.dump(dm_spec, stream, indent=1)
    if not do_post:
        # map ids from the local spec (ids as-authored) for a dry-run wb spec
        els=dict(LOCAL_DM_IDS)
        wb_spec=workbook_wire.wire_workbook(build_wb_spec(P,"<dm-id>",els,landed=landed,workflows=workflows,plan_src=plan_src,history=history))
        try:
            emit_workflow_outputs(P, wb_spec, outdir, workflows, args.emit_manual_commit_program, history)
        except demo_workflows.WorkflowError as error:
            print(f"PLAN ERROR: --emit-manual-commit-program {error}", file=sys.stderr)
            return 1
        try:
            wb_spec = merge_agent(wb_spec, agent_macros, workflows, driver_example=driver_example(P))
        except AGENT_ERRORS as error:
            print(f"AGENT ERROR: {error}", file=sys.stderr)
            return 1
        with open(f"{outdir}/wb_spec.json", "w", encoding="utf-8") as stream:
            json.dump(wb_spec, stream, indent=1)
        print(f"DRY RUN — specs written to {outdir}/ (dm_spec.json, wb_spec.json). Re-run with --post to build live.")
        return 0
    sigma_auth.apply_cli_args(args)
    e=env(); base=e["SIGMA_BASE_URL"]; tok=token(e)
    dm=api(base,tok,"POST","/v2/dataModels/spec",dm_spec); dm_id=dm["dataModelId"]
    print("dataModelId:",dm_id)
    full=api(base,tok,"GET",f"/v2/dataModels/{dm_id}/spec")
    els=map_dm_elements(full)
    # Layout is the last write before the current contents API validates + creates.
    wb_spec=workbook_wire.wire_workbook(build_wb_spec(P,dm_id,els,landed=landed,workflows=workflows,plan_src=plan_src,history=history))
    try:   # preflighted above; a failure here still names the data model it leaves behind
        written = emit_workflow_outputs(P, wb_spec, outdir, workflows, args.emit_manual_commit_program, history)
    except demo_workflows.WorkflowError as error:
        print(f"PLAN ERROR: --emit-manual-commit-program {error} — data model {dm_id} was created; no workbook "
              "was posted", file=sys.stderr)
        return 1
    try:   # preflighted above; a failure here still names the data model it leaves behind
        wb_spec = merge_agent(wb_spec, agent_macros, workflows, driver_example=driver_example(P))
    except AGENT_ERRORS as error:
        print(f"AGENT ERROR: {error} — data model {dm_id} was created; no workbook was posted", file=sys.stderr)
        return 1
    with open(f"{outdir}/wb_spec.posted.json", "w", encoding="utf-8") as stream:
        json.dump(wb_spec, stream, indent=1)
    checked=api(base,tok,"POST","/v2/workbooks",workbook_wire.dry_run_request(wb_spec))
    workbook_wire.assert_dry_run_success(checked)
    wb=api(base,tok,"POST","/v2/workbooks",wb_spec)
    print("workbookId:",wb["workbookId"],"| url:",wb.get("url"))
    with open(f"{outdir}/posted-ids.json", "w", encoding="utf-8") as stream:
        json.dump({"dataModelId": dm_id, "dmElements": els, "workbookId": wb["workbookId"], "url": wb.get("url"),
                   "staging": bool(args.staging_approval)}, stream, indent=1)
    if written:
        print(f"NEXT: run {written} once on this workbook through Sigma MCP Code Mode "
              "build_workbook (target existing) before anyone edits it or the agent writes. Pass the file "
              "unchanged as one program (do not shorten, split or rewrite it); if it fails, report the error "
              "and stop.")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
