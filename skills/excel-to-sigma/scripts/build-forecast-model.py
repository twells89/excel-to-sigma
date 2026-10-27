#!/usr/bin/env python3
"""
build-forecast-model.py — one-shot builder for an actuals + driver-grown + manual
planning model (the architecture in refs/forecast-recipes.md). Consumes a model-plan
JSON (drafted from xlsx-discover.py's model map + the intent answers) and emits:
  - a Sigma data model (GL actuals VALUES + assumptions + raw forecast spine)
  - a workbook (union[actuals, forecast] → KPIs, P&L pivot w/ Actual|Forecast cols,
    monthly summary, trend, GL detail, assumptions; editable rate+manual input tables
    when intent is editable/what-if)
  - paste-ready seed CSVs for the editable tables
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
from lib import forecast_presentation as fp  # noqa: E402
from lib import presentation_style as ps  # noqa: E402
from lib.migration_gate import MigrationGateError, gate_plan_source  # noqa: E402

CUR = lambda p: {"kind": "number", "formatString": p["currencyFormat"]}
PCT = lambda p: {"kind": "number", "formatString": p.get("pctFormat", ".1%")}
DATEF = {"kind": "datetime", "formatString": "%b-%y"}
def q(s): return "'" + str(s).replace("'", "''") + "'"


KPI_IDS = ("kpiRev", "kpiEbitda", "kpiNI", "kpiGM")


def forecast_layout(editable, extra_model_ids=(), chart_ids=(), kpi_ids=KPI_IDS):
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
    pages = [
        ("pnl", [
            ("pnlHeader", 1, 25, 1, 4),
            *kpi_cells,
            *charts,
            ("pnlPivot", 1, 25, pivot_row, pivot_row + 30),
        ]),
        ("detail", [("glTable", 1, 25, 1, 37)]),
        ("inputs", [
            ("assumpTable", 1, 25, 1, 13),
            *([("rateInput", 1, 25, 13, 29), ("manualInput", 1, 25, 29, 49)] if editable else []),
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


def build_dm_spec(P):
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
    glDetail = {"id": "glDetail", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(gl, gl_cols)},
                "columns": dmcols([("TXN_DATE","Txn Date"),("MONTH","Month"),("MONTH_NUM","Month Num"),("MONTH_END","Month End"),("LINE_ITEM","Line Item"),("SECTION","Section"),("SEC_ORDER","Section Order"),("LINE_ORDER","Line Order"),("DEPARTMENT","Department"),("VENDOR","Vendor"),("AMOUNT","Amount")])}
    # assumptions
    arows = [[q(a), f"{b}", q(c)] for a, b, c in P["assumptions"]]
    rate_name = "Monthly Rate" if periods[0]["grain"] == "month" else "Rate"
    assumptions = {"id": "assumptions", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(arows, ["DRIVER","MONTHLY_RATE","BASIS"])},
                   "columns": dmcols([("DRIVER","Driver"),("MONTHLY_RATE",rate_name),("BASIS","Basis")])}
    # raw forecast spine
    terms = spine_terms(P, ev)
    srows = []
    for ln in P["lines"]:
        for i, m in enumerate(forecast_views):
            kind, rate, manual, base, n = terms[ln["line"]][i]
            row = [q(ln["line"]), q(ln["section"]), str(ln["secOrder"]), str(ln["lineOrder"]),
                   q(m["month"]), str(m["num"]), sql_date(m["end"]), str(n), q(kind),
                   f"{rate}" if rate is not None else "NULL",
                   f"{manual:.2f}" if manual is not None else "NULL",
                   f"{base:.2f}" if base is not None else "NULL"]
            if has_pct:
                row.append(q(ln["of"]) if ln["kind"] == "pct_of_total" else "NULL")
            srows.append(row)
    scols = ["LINE_ITEM","SECTION","SEC_ORDER","LINE_ORDER","MONTH","MONTH_NUM","MONTH_END","N","KIND","DEFAULT_RATE","MANUAL_VALUE","JUNE_BASE"]
    spine_pairs = [("LINE_ITEM","Line Item"),("SECTION","Section"),("SEC_ORDER","Section Order"),("LINE_ORDER","Line Order"),("MONTH","Month"),("MONTH_NUM","Month Num"),("MONTH_END","Month End"),("N","N"),("KIND","Kind"),("DEFAULT_RATE","Default Rate"),("MANUAL_VALUE","Manual Value"),("JUNE_BASE","June Base")]
    if has_pct:
        scols.append("OF_SECTION"); spine_pairs.append(("OF_SECTION", "Of Section"))
    spine = {"id": "spine", "kind": "table", "source": {"kind": "sql", "connectionId": conn, "statement": values_sql(srows, scols)},
             "columns": dmcols(spine_pairs)}
    return {"name": P["name"] + " — Model", "schemaVersion": 1, "folderId": P["folderId"],
            "pages": [{"id": "model", "name": "Model", "elements": [glDetail, assumptions, spine]}]}

def dmcols(pairs):
    return [{"id": f"c{i}", "name": nm, "formula": f"[Custom SQL/{al}]"} for i, (al, nm) in enumerate(pairs, 1)]

def map_dm_elements(spec):
    """Identify element ids by a signature column name (names round-trip; ids reassign)."""
    out = {}
    for el in spec["pages"][0]["elements"]:
        names = {c["name"] for c in el["columns"]}
        if "Vendor" in names: out["glDetail"] = el["id"]
        elif "Kind" in names: out["spine"] = el["id"]
        elif "Driver" in names: out["assumptions"] = el["id"]
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


def seed_csvs(P, outdir):
    rate_p = os.path.join(outdir, "rate_seed.csv"); man_p = os.path.join(outdir, "manual_seed.csv")
    with open(rate_p, "w", newline="") as f:
        w = csv.writer(f); w.writerow(["LINE_ITEM", "RATE"])
        for ln in P["lines"]:
            if ln["kind"] in RATE_KINDS: w.writerow([ln["line"], ln["rate"]])
    with open(man_p, "w", newline="") as f:
        w = csv.writer(f); w.writerow(["LINE_ITEM", "MONTH", "MONTH_NUM", "MONTH_END", "AMOUNT"])
        for ln in P["lines"]:
            if ln["kind"] == "manual":
                periods, n_act = calc.plan_periods(P)
                for i, m in enumerate(calc.month_views(periods)[n_act:]):
                    w.writerow([ln["line"], m["month"], m["num"], m["end"] or "", f"{ln['manual'][i]:.2f}"])
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


def metric_key_of(P, section):
    """The name a source chart gives a section's total (its total label, mapped like the parity targets)."""
    label = next((m["totalLabel"] for m in P.get("sections", []) if m["name"] == section and m.get("totalLabel")), None)
    return calc.TARGET_LABELS.get(label.lower(), label) if label else "Revenue"


def build_wb_spec(P, dm_id, els, *, approved_inputs=None):
    conn = P["connectionId"]; editable = P["intent"] in ("editable", "what-if"); DMP = "Custom SQL"; C = "PnL Combined"
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
    pres = P.get("presentation")      # None: a plan drafted before the presentation port -> the legacy page
    if pres is not None and fp.check_look(P):
        raise PlanError(fp.check_look(P))
    look = (pres or {}).get("look")
    fm = fp.Formats(P, cur, PCT(P))   # source number formats per element; CLI defaults for General / legacy plans
    accent = (pres or {}).get("accent") or ps.accent_from_fill(None)
    def dm(elid): return {"kind": "data-model", "dataModelId": dm_id, "elementId": els[elid]}
    elements = []
    # actuals wrap (live)
    actuals = {"id": "actuals", "kind": "table", "name": "Actuals", "source": dm("glDetail"),
        "columns": [{"id":"a_sec","name":"Section","formula":f"[{DMP}/Section]"},{"id":"a_so","name":"Section Order","formula":f"[{DMP}/Section Order]"},
                    {"id":"a_line","name":"Line Item","formula":f"[{DMP}/Line Item]"},{"id":"a_lo","name":"Line Order","formula":f"[{DMP}/Line Order]"},
                    {"id":"a_m","name":"Month","formula":f"[{DMP}/Month]"},{"id":"a_mn","name":"Month Num","formula":f"[{DMP}/Month Num]"},
                    {"id":"a_me","name":"Month End","formula":f"[{DMP}/Month End]"},{"id":"a_amt","name":"Amount","formula":f"[{DMP}/Amount]"},
                    {"id":"a_p","name":"Period","formula":'"Actual"'}]}
    spine_names = [("sec","Section"),("so","Section Order"),("line","Line Item"),("lo","Line Order"),("m","Month"),("mn","Month Num"),("me","Month End"),("n","N"),("k","Kind"),("dr","Default Rate"),("mv","Manual Value"),("jb","June Base")]
    if has_pct:
        spine_names.append(("of", "Of Section"))
    spineWrap = {"id":"spineWrap","kind":"table","name":"Forecast Spine","source":dm("spine"),
        "columns":[{"id":f"s_{k}","name":nm,"formula":f"[{DMP}/{nm}]"} for k,nm in spine_names]}
    # forecast element(s). Without percentage-of-total lines there is one "forecast" table;
    # with them, "forecastCalc" computes every other line and "forecast" adds the lines
    # that are a share of another section's forecast total (a second pass over per-month totals).
    stage1_id, stage1_name = ("forecastCalc", "Forecast Base") if has_pct else ("forecast", "Forecast")
    pct_branch = '[Kind] = "pct", 0, ' if has_pct else ""
    if editable:
        rateInput = {"id":"rateInput","kind":"input-table","name":"Growth Rates (editable)","inputMode":"edit","source":{"kind":"empty","connectionId":conn},
            "columns":[{"id":"ID"},{"id":"LINE_ITEM","type":"text","name":"Line Item"},{"id":"RATE","type":"number","name":"Rate","format":fm.drivers}],"order":["LINE_ITEM","RATE"]}
        manualInput = {"id":"manualInput","kind":"input-table","name":"Manual Forecast (editable)","inputMode":"edit","source":{"kind":"empty","connectionId":conn},
            "columns":[{"id":"ID"},{"id":"LINE_ITEM","type":"text","name":"Line Item"},{"id":"MONTH","type":"text","name":"Month"},{"id":"MONTH_NUM","type":"number","name":"Month Num"},{"id":"MONTH_END","type":"datetime","name":"Month End"},{"id":"AMOUNT","type":"number","name":"Amount","format":fm.money}],"order":["LINE_ITEM","MONTH","AMOUNT"]}
        forecast = {"id":stage1_id,"kind":"table","name":stage1_name,
            "source":{"kind":"join","name":"FJ","primarySource":{"kind":"table","elementId":"spineWrap"},
                "joins":[{"name":"Rates","joinType":"left-outer","left":{"kind":"table","elementId":"spineWrap"},"right":{"kind":"table","elementId":"rateInput"},"columns":[{"left":"[Line Item]","right":"[Line Item]"}]},
                         {"name":"Manual","joinType":"left-outer","left":{"kind":"table","elementId":"spineWrap"},"right":{"kind":"table","elementId":"manualInput"},"columns":[{"left":"[Line Item]","right":"[Line Item]"},{"left":"[Month Num]","right":"[Month Num]"}]}]},
            "columns":[{"id":"f_sec","name":"Section","formula":"[FJ/Section]"},{"id":"f_so","name":"Section Order","formula":"[FJ/Section Order]"},
                {"id":"f_line","name":"Line Item","formula":"[FJ/Line Item]"},{"id":"f_lo","name":"Line Order","formula":"[FJ/Line Order]"},
                {"id":"f_m","name":"Month","formula":"[FJ/Month]"},{"id":"f_mn","name":"Month Num","formula":"[FJ/Month Num]"},{"id":"f_me","name":"Month End","formula":"[FJ/Month End]"},
                {"id":"f_n","name":"N","formula":"[FJ/N]"},{"id":"f_k","name":"Kind","formula":"[FJ/Kind]"},{"id":"f_jb","name":"June Base","formula":"[FJ/June Base]"},
                *([{"id":"f_of","name":"Of Section","formula":"[FJ/Of Section]"}] if has_pct else []),
                {"id":"f_er","name":"Eff Rate","formula":"Coalesce([Rates/Rate], [FJ/Default Rate])"},
                {"id":"f_em","name":"Eff Manual","formula":"Coalesce([Manual/Amount], [FJ/Manual Value])"},
                {"id":"f_amt","name":"Amount","formula":'If([Kind] = "manual", [Eff Manual], [Kind] = "flat", [June Base], ' + pct_branch + '[June Base] * (1 + [Eff Rate]) ^ [N])',"format":fm.money},
                {"id":"f_p","name":"Period","formula":'"Forecast"'}]}
    else:
        forecast = {"id":stage1_id,"kind":"table","name":stage1_name,"source":{"kind":"table","elementId":"spineWrap"},
            "columns":[{"id":"f_sec","name":"Section","formula":"[Forecast Spine/Section]"},{"id":"f_so","name":"Section Order","formula":"[Forecast Spine/Section Order]"},
                {"id":"f_line","name":"Line Item","formula":"[Forecast Spine/Line Item]"},{"id":"f_lo","name":"Line Order","formula":"[Forecast Spine/Line Order]"},
                {"id":"f_m","name":"Month","formula":"[Forecast Spine/Month]"},{"id":"f_mn","name":"Month Num","formula":"[Forecast Spine/Month Num]"},{"id":"f_me","name":"Month End","formula":"[Forecast Spine/Month End]"},
                *([{"id":"f_k","name":"Kind","formula":"[Forecast Spine/Kind]"},{"id":"f_of","name":"Of Section","formula":"[Forecast Spine/Of Section]"},
                   {"id":"f_er","name":"Eff Rate","formula":"[Forecast Spine/Default Rate]"}] if has_pct else []),
                {"id":"f_amt","name":"Amount","formula":'If([Forecast Spine/Kind] = "manual", [Forecast Spine/Manual Value], [Forecast Spine/Kind] = "flat", [Forecast Spine/June Base], ' + ('[Forecast Spine/Kind] = "pct", 0, ' if has_pct else "") + '[Forecast Spine/June Base] * (1 + [Forecast Spine/Default Rate]) ^ [Forecast Spine/N])',"format":fm.money},
                {"id":"f_p","name":"Period","formula":'"Forecast"'}]}
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
                       {"id":"r_amt","name":"Amount","formula":f"[{stage1_name}/Amount]","format":fm.money}]}
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
                {"id":"g_amt","name":"Amount","formula":f'If([{stage1_name}/Kind] = "pct", [{stage1_name}/Eff Rate] * If({branches}, 0), [{stage1_name}/Amount])',"format":fm.money},
                {"id":"g_p","name":"Period","formula":'"Forecast"'}]}
        forecast_stages = [forecast, forecastRef, final]
        forecast = final
    def m(nm): return {"outputColumnName": nm, "sourceColumns": [f"[{nm}]", f"[{nm}]"]}
    pnlCombined = {"id":"pnlCombined","kind":"table","name":C,
        "source":{"kind":"union","sources":[{"kind":"table","elementId":"actuals"},{"kind":"table","elementId":"forecast"}],
                  "matches":[m("Section"),m("Section Order"),m("Line Item"),m("Line Order"),m("Month"),m("Month Num"),m("Month End"),m("Amount"),m("Period")]},
        "columns":[{"id":"c_sec","name":"Section","formula":f"[{U}/Section]"},{"id":"c_so","name":"Section Order","formula":f"[{U}/Section Order]"},
            {"id":"c_line","name":"Line Item","formula":f"[{U}/Line Item]"},{"id":"c_lo","name":"Line Order","formula":f"[{U}/Line Order]"},
            {"id":"c_m","name":"Month","formula":f"[{U}/Month]"},{"id":"c_mn","name":"Month Num","formula":f"[{U}/Month Num]"},{"id":"c_me","name":"Month End","formula":f"[{U}/Month End]"},
            {"id":"c_amt","name":"Amount","formula":f"[{U}/Amount]","format":fm.money},{"id":"c_p","name":"Period","formula":f"[{U}/Period]"},
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
    slots = [("kpiRev", kpi_title(rev_label), REV_FY, fm.section(first_positive))]
    if middle is not None:
        slots.append(("kpiEbitda", kpi_title(rows[middle]["label"]), signed(rows[middle], True), fm.derived(rows[middle]["label"])))
    slots.append(("kpiNI", kpi_title(rows[bottom]["label"]), signed(rows[bottom], True), fm.derived(rows[bottom]["label"])))
    if len(rows) >= 2:
        first = rows[0]
        ratio_name = "Gross Margin %" if first["label"].lower() == "gross profit" else f"{first['label']} %"
        slots.append(("kpiGM", kpi_title(ratio_name), f"({signed(first, True)}) / {REV_FY}", fm.ratio(ratio_name)))
    if pres is None:
        kpis = [kpi(*slot) for slot in slots]
        kpis[0]["style"]["borderColor"] = "#2563EB"
        header={"id":"pnlHeader","kind":"text","body":f"# {P['name']}\n{grain_adj} actuals and forecast performance",
            "style":{"backgroundColor":"#F8FAFC","color":"#0F172A","borderRadius":"round"}}
    else:       # equal peer cards (table-report recipe): the same formulas, source formats, no accent on the first
        kpis = [ps.kpi_card(eid, name.upper(), {"kind": "table", "elementId": "pnlCombined"}, f, fmt, look)
                for eid, name, f, fmt in slots]
    tail = [rows[i] for i in sorted({middle, bottom} - {None})]
    trend_series = [(first_positive, REV)] + [(r["label"], signed(r)) for r in tail]
    trend_names = [name for name, _ in trend_series]
    trend_label = ", ".join(trend_names[:-1]) + " & " + trend_names[-1] if len(trend_names) > 1 else trend_names[0]
    trend_formats = [fm.section(first_positive)] + [fm.derived(r["label"]) for r in tail]
    series_cols = [{"id":"t_" + re.sub(r"\W+", "_", name.lower()).strip("_"),"name":name,"formula":f,"format":fmt}
                   for (name, f), fmt in zip(trend_series, trend_formats)]
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
    if pres is not None:
        trend.pop("style"); style = ps.chart_card(look)
        if style: trend["style"] = style
        ps.flat_labels(trend["xAxis"], len(periods))
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
            {"id":"p_amt","name":"Amount","formula":f"Sum([{C}/Amount])","format":fm.money}],
        "values":["p_amt"],
        "rowsBy":[{"columnId":"p_sec","sort":{"direction":"ascending","by":"p_so","aggregation":"min"}},{"columnId":"p_line","sort":{"direction":"ascending","by":"p_lo","aggregation":"min"}}],
        "columnsBy":p_cols_by,
        "axisTotals":[{"axisId":"grand-row-total","visibility":"hidden"}],
        "style":{"backgroundColor":"#FFFFFF","borderColor":{"kind":"theme","ref":"colors-border"},"borderWidth":1,"borderRadius":"round"}}
    glTable={"id":"glTable","kind":"table","name":"GL Transaction Detail (Actuals)","source":dm("glDetail"),
        "columns":[{"id":"g1","name":"Txn Date","formula":f"[{DMP}/Txn Date]",**({"format":fm.gl_date} if pres is not None and fm.gl_date else {})},{"id":"g2","name":"Month","formula":f"[{DMP}/Month]"},{"id":"g3","name":"Line Item","formula":f"[{DMP}/Line Item]"},{"id":"g4","name":"Department","formula":f"[{DMP}/Department]"},{"id":"g5","name":"Vendor","formula":f"[{DMP}/Vendor]"},{"id":"g6","name":"Amount","formula":f"[{DMP}/Amount]","format":fm.gl}]}
    assumpTable={"id":"assumpTable","kind":"table","name":"Forecast Assumptions","source":dm("assumptions"),
        "columns":[{"id":"d1","name":"Driver","formula":f"[{DMP}/Driver]"},{"id":"d2","name":rate_name,"formula":f"[{DMP}/{rate_name}]","format":fm.drivers},{"id":"d3","name":"Basis","formula":f"[{DMP}/Basis]"}]}
    # Source-workbook charts the drafter mapped and verified (plan.sourceCharts). Series are
    # named like the plan's section totals / amount rows (classic rows use Revenue, Gross
    # Profit, EBITDA, Net Income).
    metric_formula, metric_format = {}, {}
    for sec in sections:
        label = next((m_["totalLabel"] for m_ in P.get("sections", []) if m_["name"] == sec["name"] and m_.get("totalLabel")), None)
        if label:
            metric_formula[calc.TARGET_LABELS.get(label.lower(), label)] = sec_sum(sec["name"])
            metric_format[calc.TARGET_LABELS.get(label.lower(), label)] = fm.section(sec["name"])
    for row in rows:
        metric_formula[calc.TARGET_LABELS.get(row["label"].lower(), row["label"])] = signed(row)
        metric_format[calc.TARGET_LABELS.get(row["label"].lower(), row["label"])] = fm.derived(row["label"])
    if not P.get("sections"):
        metric_formula["Revenue"] = REV
        metric_format["Revenue"] = fm.section(first_positive)
    source_charts = []
    for i, sc in enumerate(P.get("sourceCharts") or [], 1):
        if sc["grain"] == "quarter":
            x = {"id": f"s{i}_x", "name": "Quarter",
                 "formula": f'Concat(Text(DatePart("year", [{C}/Month End])), " Q", Text(DatePart("quarter", [{C}/Month End])))'}
        elif sc["grain"] == "period":
            x = {"id": f"s{i}_x", "name": "Period", "formula": f"[{C}/Month]"}
        else:
            x = {"id": f"s{i}_x", "name": "Month", "formula": f'DateTrunc("month", [{C}/Month End])', "format": DATEF}
        ys = [{"id": f"s{i}_y{j}", "name": metric, "formula": metric_formula[metric], "format": metric_format[metric]}
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
        if pres is not None:      # the table-report chart recipe: card chrome, flat labels, one colour for one series
            chart.pop("style"); style = ps.chart_card(look)
            if style: chart["style"] = style
            ps.flat_labels(chart["xAxis"], len(sc["categories"]))
            if len(ys) == 1:
                chart["legend"] = {"visibility": "hidden"}
                one = ps.single_color((sc.get("colors") or [None])[0] or accent["hex"]) if look is None else None
                if one: chart["color"] = one
        source_charts.append(chart)
    hidden_els=[actuals,spineWrap,*forecast_stages,pnlCombined]
    extra_model_ids = [e["id"] for e in forecast_stages[:-1]] if has_pct else []
    if approved_inputs is not None:
        extra = [e for e in approved_graph["elements"] if e["id"] not in {n["id"] for n in hidden_els}]
        hidden_els += extra
        extra_model_ids = [e["id"] for e in extra]
    if pres is not None:
        # The generated trend repeats a mapped source chart that shows the same series at the plan's grain.
        trend_keys = [metric_key_of(P, first_positive)] + [calc.TARGET_LABELS.get(r["label"].lower(), r["label"]) for r in tail]
        charts = ([] if fp.trend_covered(P, trend_keys, periods[0]["grain"]) else [trend]) + source_charts
        page = fp.compose(P, kpis=kpis, charts=charts, pivot=pivot, gl=glTable, assump=assumpTable if P.get("assumptions") else None,
                          inputs=[rateInput, manualInput] if editable else [], hidden=hidden_els,
                          hidden_ids=[e["id"] for e in hidden_els], fm=fm, n_pivot_rows=len(P["lines"]) + len(sections))
        spec = {"name":P["name"],"schemaVersion":1,"kind":"workbook","folderId":P["folderId"],
                "pages":page["pages"],"layout":page["layout"]}
        if page.get("settings"):
            spec["settings"] = page["settings"]
        return spec
    main_els=[header]+kpis+[trend]+source_charts+[pivot]
    detail_els=[glTable]
    input_els=[assumpTable]
    if editable: input_els+=[rateInput,manualInput]
    pages=[{"id":"pnl","name":"P&L Overview","elements":main_els},
           {"id":"detail","name":"Detail","elements":detail_els},
           {"id":"inputs","name":"Inputs","elements":input_els},
           {"id":"model","name":"Model (sources)","visibility":"hidden","elements":hidden_els}]
    # Nested draft; wire_workbook flattens + wraps at the POST / dry-run boundary.
    return {"name":P["name"],"schemaVersion":1,"kind":"workbook","folderId":P["folderId"],
            "pages":pages,"layout":forecast_layout(editable, extra_model_ids, [c["id"] for c in source_charts], [k["id"] for k in kpis])}

def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", help="model-plan JSON")
    parser.add_argument("--source", help="source XLSX used to produce the plan")
    parser.add_argument("--review-dispositions", help="source-bound review dispositions JSON")
    parser.add_argument("--post", action="store_true", help="create the data model and workbook")
    sigma_auth.add_cli_args(parser)
    parser.add_argument("--out-dir", default="/tmp/fcst-build",
                        help="where specs and seed CSVs are written (default: /tmp/fcst-build)")
    parser.add_argument("--parity-tolerance", type=float, default=None,
                        help="absolute tolerance for plan.parityTargets (default: float noise only, "
                             "max(1e-6, 1e-9 * |source value|))")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    with open(args.plan, encoding="utf-8") as stream:
        P = json.load(stream)
    try:
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
    print(
        f"MIGRATION GATE: {gate_report['status']} "
        f"source={source_path.name} sha256={gate_report['sourceSha256']}"
    )
    if gate_report.get("pendingReview"):
        print("DRAFT MODE: review pending for " + ", ".join(gate_report["pendingReview"])
              + " — fine for a local build; resolve before --post.")
    do_post = args.post
    outdir = args.out_dir
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
        build_wb_spec(P,"<dm-id>",{"glDetail":"glDetail","spine":"spine","assumptions":"assumptions"})
    except PlanError as error:
        print(f"PLAN ERROR: {error}", file=sys.stderr)
        return 1
    os.makedirs(outdir, exist_ok=True)
    rate_csv,man_csv=seed_csvs(P,outdir)
    print(f"seed CSVs: {rate_csv}, {man_csv}")
    dm_spec=build_dm_spec(P)
    with open(f"{outdir}/dm_spec.json", "w", encoding="utf-8") as stream:
        json.dump(dm_spec, stream, indent=1)
    if not do_post:
        # map ids from the local spec (ids as-authored) for a dry-run wb spec
        els={"glDetail":"glDetail","spine":"spine","assumptions":"assumptions"}
        wb_spec=workbook_wire.wire_workbook(build_wb_spec(P,"<dm-id>",els))
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
    wb_spec=workbook_wire.wire_workbook(build_wb_spec(P,dm_id,els))
    checked=api(base,tok,"POST","/v2/workbooks",workbook_wire.dry_run_request(wb_spec))
    workbook_wire.assert_dry_run_success(checked)
    wb=api(base,tok,"POST","/v2/workbooks",wb_spec)
    print("workbookId:",wb["workbookId"],"| url:",wb.get("url"))
    print("NOTE: POST uses a stacked notebook-flow layout via code_rep; refine with wb-rep push if needed.")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
