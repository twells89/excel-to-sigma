"""Planning-app presentation for a forecast built with workflows (``build-forecast-model.py --workflows``).

The plain forecast builder lays out a dashboard; once the macros became workflows the workbook is a
finance team's planning app, so it is laid out like one (the approved design target):

  Forecast            header band (title, scenario switch, approval status pill, Roll forward /
                      Snapshot / Submit for approval), four KPI cards, actual-forecast-plan trend +
                      agent panel, forecast vs plan by region + P&L summary
  Assumptions         editable growth drivers and manual lines, with the forecast reacting beside them
  P&L detail          the monthly P&L pivot (Actual | Forecast) and the GL actuals behind it
  Versions & approvals read-only, newest-first views of the governed logs
  Model (hidden)      every helper: wraps, union, Parameters / Versions / Approvals input tables, status,
                      and the one forecast boundary (``Forecast Boundary``) every surface reads

Everything here is a pure function of the plan and the builder's own elements; nothing contacts Sigma.
Every formula reads elements by their display name (``[Forecast Facts/Amount]``) like the rest of the
builder, and every text span uses only the properties Sigma accepts (color, background-color,
font-size, font-family).
"""
from __future__ import annotations

import re

import demo_workflows as dw

# ------------------------------------------------------------------ palette (neutral Sigma look)
INK, MUTED, HAIRLINE, CANVAS = "#0F172A", "#64748B", "#E2E8F0", "#FFFFFF"
BAND = "#F8FAFC"
PRIMARY = "#2563EB"
FORECAST_TINT = "#93C5FD"
PLAN_GREY = "#94A3B8"
GOOD, BAD = "#15803D", "#B91C1C"
PILLS = {"Draft": ("#FEF3C7", "#B45309"), "Submitted": ("#DBEAFE", "#1D4ED8"), "Pending": ("#DBEAFE", "#1D4ED8"),
         "Approved": ("#DCFCE7", "#15803D"), "Returned": ("#FEE2E2", "#B91C1C")}
CARD = {"backgroundColor": CANVAS, "borderColor": HAIRLINE, "borderWidth": 1, "borderRadius": "round"}

SCENARIO_CONTROL = "Scenario"          # controlId: the forecast's Eff Rate reads [Scenario]
FACTS, PLAN_LINES, SUMMARY, METRICS = "Forecast Facts", "Plan Lines", "Summary Facts", "Metric Map"
COMBINED, STATUS = "PnL Combined", "Latest approval status"
# The actual / forecast boundary, defined once (``boundary_element``) and read everywhere by these refs.
BOUNDARY, BOUNDARY_ID = "Forecast Boundary", "boundary"
START_MONTH = f"Max([{BOUNDARY}/Start Month])"
START_KEY = f"Max([{BOUNDARY}/Start Key])"
LAST_CLOSED = f"Max([{BOUNDARY}/Last Closed Month])"
AWAITING = "Awaiting actuals"           # a closed month whose actuals have not landed: forecast values, flagged
AMBER = "#B45309"
BUTTON_TEXT = {"roll-forward": "Roll forward", "snapshot-version": "Snapshot",
               "submit-approval": "Submit for approval"}
PAGES = (("pgForecast", "Forecast"), ("pgAssumptions", "Assumptions"), ("pgDetail", "P&L detail"),
         ("pgVersions", "Versions & approvals"))
PAGE_NAMES = dict(PAGES)


def _join_and(items):
    items = list(items)
    return ", ".join(items[:-1]) + " and " + items[-1] if len(items) > 1 else "".join(items)
MONEY_M = "$,.1f"                       # applied to value / 1e6, then "M"
CUR0 = {"kind": "number", "formatString": "$,.0f"}
CUR_SI = {"kind": "number", "formatString": "$.3~s"}
MILLIONS = {"kind": "number", "formatString": ",.1f", "prefix": "$", "suffix": "M"}   # on value / 1e6
PCT1 = {"kind": "number", "formatString": ".1%"}
SIGNED_PCT = {"kind": "number", "formatString": "+.1%"}
MONTH_FMT = {"kind": "datetime", "formatString": "%b-%y"}
STAMP_FMT = {"kind": "datetime", "formatString": "%b %-d, %Y"}   # migrated log rows carry a date only
NULL_DASH = "\u2014"                    # a log figure the source never froze reads as a dash, not "null"


# a version / status tail a file name carries but an app title does not: v7, FINAL, draft, copy, (2)
_NAME_TAIL = re.compile(r"(?:\s+(?:v\d+(?:\.\d+)*|final|draft|copy|latest|wip|\(\d+\)))+$", re.I)


def display_name(raw):
    """A file stem as a title: ``Vantora_FY27_Rolling_Forecast_v7_FINAL`` -> ``Vantora FY27 Rolling Forecast``.

    Underscores and hyphen-joined words become spaces, a trailing ``- draft`` style suffix and
    version / status words (v7, FINAL, draft, copy, (2)) are dropped. Never returns an empty name."""
    text = re.sub(r"[_]+", " ", str(raw))
    text = re.sub(r"(?<=\w)-(?=\w)", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s+[—–-]\s+(?:draft|copy|final|wip|latest)(?:\s.*)?$", "", text, flags=re.I)
    cleaned = _NAME_TAIL.sub("", text).strip(" -—–")
    cleaned = _NAME_TAIL.sub("", cleaned).strip(" -—–")
    return cleaned or text


def sentence(label):
    """``Gross Margin`` -> ``Gross margin``; all-caps words (EBITDA, FY27) keep their case."""
    words = str(label).split()
    out = []
    for i, word in enumerate(words):
        keep = word.isupper() or any(ch.isdigit() for ch in word) or i == 0
        out.append(word if keep else word.lower())
    return " ".join(out)


def esc(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("{", "(").replace("}", ")").replace("*", "\\*"))


def span(text, color=None, size=None, background=None):
    parts = ([f"color: {color}"] if color else []) + ([f"background-color: {background}"] if background else []) \
        + ([f"font-size: {size}px"] if size else [])
    return f'<span style="{"; ".join(parts)}">{text}</span>' if parts else text


def money_m(expr):
    """``$145.4M`` from a dollar expression."""
    return f"{{{{({expr}) / 1000000 | {MONEY_M}}}}}M"


def signed(expr, unit, suffix, scale=100):
    """Green ``▲ 2.3% vs plan`` / red ``▼ 1.1% vs plan``: two spans, only the matching one non-empty."""
    up = f'{{{{If(({expr}) >= 0, Concat("▲ ", Text(Round(({expr}) * {scale}, 1)), "{unit}"), "")}}}}'
    down = f'{{{{If(({expr}) < 0, Concat("▼ ", Text(Round(Abs({expr}) * {scale}, 1)), "{unit}"), "")}}}}'
    return span(up, GOOD) + span(down, BAD) + span(f" {esc(suffix)}", MUTED)


def month_key(date_expr):
    return f'DatePart("year", {date_expr}) * 12 + DatePart("month", {date_expr})'


def month_label(expr):
    return f'DateFormat({expr}, "%b-%y")'


def boundary_element(connection_id, first_forecast, parameters=True):
    """The planning app's one actual / forecast boundary: ``Parameters.Forecast Start Month``, or the
    source's first forecast month (``first_forecast`` = ``(year, month)``) while the Parameters row is
    not seeded. One row whatever Parameters holds (a one-row VALUES query anchors it), so every other
    element reads it as ``Max([Forecast Boundary/<column>])`` and nothing repeats the fallback.
    Columns: Start Month, Start Key (year * 12 + month, the comparison key), Last Closed Month.
    Without the Parameters table (no set-scenario / roll-forward workflow) the boundary is the source's
    first forecast month."""
    year, month = first_forecast
    first = f"MakeDate({int(year)}, {int(month)}, 1)"
    start = (f"Coalesce(Max([{dw.TABLES[dw.PARAMETERS][0]}/Forecast Start Month]), {first})"
             if parameters else first)
    return {"id": BOUNDARY_ID, "kind": "table", "name": BOUNDARY,
            "source": {"kind": "sql", "connectionId": connection_id,
                       "statement": "select * from (values\n(1)\n) as t(BOUNDARY_ROW)"},
            "columns": [{"id": "bd_start", "name": "Start Month", "formula": start, "format": MONTH_FMT},
                        {"id": "bd_key", "name": "Start Key", "formula": month_key("[Start Month]")},
                        {"id": "bd_last", "name": "Last Closed Month",
                         "formula": 'DateAdd("month", -1, [Start Month])', "format": MONTH_FMT}]}


class Facts:
    """SumIf formulas over the ``Forecast Facts`` union (actual + forecast + plan rows)."""

    def __init__(self, fy_nums, prior_nums):
        self.fy, self.prior = fy_nums, prior_nums

    @staticmethod
    def _window(nums):
        return f" and [{FACTS}/Month Num] >= {nums[0]} and [{FACTS}/Month Num] <= {nums[-1]}" if nums else ""

    def cond(self, which):
        if which == "fc":
            return f'[{FACTS}/Period] != "Plan"' + self._window(self.fy)
        if which == "plan":
            return f'[{FACTS}/Period] = "Plan"' + self._window(self.fy)
        return f'[{FACTS}/Period] = "Actual"' + self._window(self.prior)

    def row(self, coefficients, which):
        out = ""
        for i, (section, weight) in enumerate(coefficients.items()):
            term = f'SumIf([{FACTS}/Amount], [{FACTS}/Section] = "{section}" and {self.cond(which)})'
            term = term if abs(weight) == 1 else f"{abs(weight):g} * {term}"
            out += (("-" if weight < 0 else "") + term) if i == 0 else (" - " if weight < 0 else " + ") + term
        return out


def variance_favourability(line, variance, cost_lines):
    """A variance whose sign means good (+) / bad (-): flipped for lines where lower is better."""
    if not cost_lines:
        return variance
    test = " or ".join(f'{line} = "{esc(name)}"' for name in cost_lines)
    return f"If({test}, -1, 1) * {variance}"


def text(element_id, body, align="center"):
    return {"id": element_id, "kind": "text", "body": body, "verticalAlign": align}


def card(element_id, style=None):
    return {"id": element_id, "kind": "container", "style": dict(style or CARD)}


def kpi_body(title, value, *lines):
    body = span(esc(title), MUTED, 13) + "\n\n" + span(f"**{value}**", INK, 28)
    for line in lines:
        if line:
            body += "\n\n" + span(line, None, 12)
    return body


def build(ctx):
    """Pages, flat elements, layout and theme for the planning app.

    ``ctx`` carries the builder's own pieces: ``P``, ``name``, ``first_positive``, ``rows`` (amount
    rows), ``kpi`` (indexes ``middle`` / ``bottom``), ``fy_nums``, ``prior_nums``, ``fy_label``,
    ``prior_label``, ``fy_end`` (ISO date of the FY's last period), ``n_fy``, ``wf`` (workflow
    elements), ``hidden`` (model elements), ``pivot``, ``gl``, ``assump``, ``rate``, ``manual``,
    ``dm`` (data-model source builder), ``plan`` (landed plan or None), ``summary`` (metric rows),
    ``scenario`` (``{"names", "default"}`` or None), ``region_of`` ({line: region}).
    """
    P, wf = ctx["P"], ctx["wf"]
    facts = Facts(ctx["fy_nums"], ctx["prior_nums"])
    has_plan, has_prior = ctx["plan"] is not None, bool(ctx["prior_nums"])
    fy, prior = ctx["fy_label"], ctx["prior_label"]
    rev = {ctx["first_positive"]: 1}
    rows = ctx["rows"]
    gm_row = rows[0] if len(rows) >= 2 else None
    eb_row = rows[ctx["kpi"]["bottom"]]
    hidden = list(ctx["hidden"])
    elements = []

    # ---------------------------------------------------------------- hidden facts: plan + union + summary
    region_of = ctx.get("region_of") or {}
    region_formula = ('Switch([Union of 2 Sources/Line Item], '
                      + ", ".join(f'"{k}", "{v}"' for k, v in sorted(region_of.items()))
                      + ', "")') if region_of else '""'
    fact_cols = ["Section", "Section Order", "Line Item", "Month Num", "Month End", "Amount", "Period"]
    if has_plan:
        plan_wrap = {"id": "planLines", "kind": "table", "name": PLAN_LINES, "source": ctx["dm"]("planDetail"),
                     "columns": [{"id": "pl_sec", "name": "Section", "formula": "[Plan Detail/Section]"},
                                 {"id": "pl_so", "name": "Section Order", "formula": "[Plan Detail/Section Order]"},
                                 {"id": "pl_line", "name": "Line Item", "formula": "[Plan Detail/Line Item]"},
                                 {"id": "pl_mn", "name": "Month Num", "formula": "[Plan Detail/Month Num]"},
                                 {"id": "pl_me", "name": "Month End", "formula": "[Plan Detail/Month End]"},
                                 {"id": "pl_amt", "name": "Amount", "formula": "[Plan Detail/Plan Amount]", "format": CUR0},
                                 {"id": "pl_p", "name": "Period", "formula": '"Plan"'}]}
        sources = [{"kind": "table", "elementId": "pnlCombined"}, {"kind": "table", "elementId": "planLines"}]
        hidden.append(plan_wrap)
    else:
        sources = [{"kind": "table", "elementId": "pnlCombined"}]
    union_name = f"Union of {len(sources)} Sources"
    report_facts = {"id": "reportFacts", "kind": "table", "name": FACTS,
                    "source": ({"kind": "union", "sources": sources,
                                "matches": [{"outputColumnName": c, "sourceColumns": [f"[{c}]"] * len(sources)}
                                            for c in fact_cols]}
                               if len(sources) > 1 else sources[0]),
                    "columns": [{"id": "rf_" + re.sub(r"\W+", "_", c.lower()), "name": c,
                                 "formula": f"[{union_name if len(sources) > 1 else COMBINED}/{c}]",
                                 **({"format": CUR0} if c == "Amount" else {})} for c in fact_cols]}
    region_src = union_name if len(sources) > 1 else COMBINED
    report_facts["columns"].append({"id": "rf_region", "name": "Region",
                                    "formula": region_formula.replace("Union of 2 Sources", region_src)})
    # Actual / Forecast by month against the one boundary, plan rows included, so a chart coloured by it
    # gives every month one colour that moves at Roll forward. A closed month whose actuals have not
    # landed still shows its forecast values, so it keeps the forecast colour.
    mk = month_key(f"[{region_src}/Month End]")
    report_facts["columns"].append({"id": "rf_phase", "name": "Phase", "formula": (
        f'If({mk} >= {START_KEY}, "Forecast", '
        f'IsNull(Lookup(Max([Actuals/Month Key]), {mk}, [Actuals/Month Key])), "Forecast", "Actual")')})
    metric_wrap = {"id": "metricWrap", "kind": "table", "name": METRICS, "source": ctx["dm"]("metricMap"),
                   "columns": [{"id": "mm_m", "name": "Metric", "formula": "[Custom SQL/Metric]"},
                               {"id": "mm_o", "name": "Metric Order", "formula": "[Custom SQL/Metric Order]"},
                               {"id": "mm_s", "name": "Section", "formula": "[Custom SQL/Section]"},
                               {"id": "mm_w", "name": "Weight", "formula": "[Custom SQL/Weight]"}]}
    summary_facts = {"id": "summaryFacts", "kind": "table", "name": SUMMARY,
                     "source": {"kind": "join", "name": "SJ", "primarySource": {"kind": "table", "elementId": "reportFacts"},
                                "joins": [{"name": "Metrics", "joinType": "inner",
                                           "left": {"kind": "table", "elementId": "reportFacts"},
                                           "right": {"kind": "table", "elementId": "metricWrap"},
                                           "columns": [{"left": "[Section]", "right": "[Section]"}]}]},
                     "columns": [{"id": "sf_m", "name": "Metric", "formula": "[Metrics/Metric]"},
                                 {"id": "sf_o", "name": "Metric Order", "formula": "[Metrics/Metric Order]"},
                                 {"id": "sf_p", "name": "Period", "formula": "[SJ/Period]"},
                                 {"id": "sf_mn", "name": "Month Num", "formula": "[SJ/Month Num]"},
                                 {"id": "sf_v", "name": "Signed Amount", "formula": "[SJ/Amount] * [Metrics/Weight]"}]}
    agent_facts = agent_facts_element(ctx, has_plan, has_prior)
    hidden += [ctx["boundary"], report_facts, metric_wrap, summary_facts, agent_facts, *wf["tables"],
               *([wf["status"]] if wf["status"] else [])]

    # ---------------------------------------------------------------- A. header band
    start = START_MONTH
    status = f"Max([{STATUS}/{STATUS}])"
    # "actuals through" is the last month whose actuals landed; a closed month still waiting for them is
    # flagged beside it (its figures stay the forecast's until the load arrives).
    awaiting = (f'{{{{If(Max([Actuals/Month Key]) < {START_KEY} - 1, Concat(" · ", '
                f'{month_label(LAST_CLOSED)}, " {AWAITING.lower()}"), "")}}}}')
    title = text("hdrTitle",
                 "## " + span(" V ", CANVAS, None, PRIMARY) + " " + esc(ctx["title"]) + "\n"
                 + span((esc(ctx["company"]) + " · actuals through " if ctx["company"] else "Actuals through ")
                        + f"{{{{{month_label('Max([Actuals/Month End])')}}}}}", MUTED, 13)
                 + span(awaiting, AMBER, 13)
                 + (span(" · live on Snowflake", MUTED, 13) if ctx.get("landed") else ""))
    pill = "".join(span(f'{{{{If({status} = "{name}", "  {name}  ", "")}}}}', fg, 14, bg)
                   for name, (bg, fg) in PILLS.items())
    # the approval status pill exists only with a Submit workflow (the Approvals log)
    status_text = text("hdrStatus", span("Status", MUTED, 12) + "\n\n" + pill) if wf["status"] else None
    buttons = []
    for b in wf["buttons"]:
        action = next(a for a, eid in dw.BUTTON_IDS.items() if eid == b["id"])
        primary = action == "submit-approval"
        # stretched to fill their cells in the action group, so the three read as one toolbar
        buttons.append({**b, "text": BUTTON_TEXT.get(action, b["text"]),
                        "appearance": "filled" if primary else "outline", "size": "medium", "align": "stretch",
                        "borderRadius": "round", **({"fillColor": PRIMARY} if primary else {})})
    controls = []
    if ctx["scenario"]:
        controls.append({"id": "ctlScenario", "kind": "control", "controlId": SCENARIO_CONTROL, "name": "Scenario",
                         "controlType": "segmented", "selectionMode": "single", "value": ctx["scenario"]["default"],
                         "source": {"kind": "manual", "valueType": "text", "values": list(ctx["scenario"]["names"])}})
    band = card("hdrBand", {"backgroundColor": BAND, "borderColor": HAIRLINE, "borderWidth": 1, "borderRadius": "round"})
    elements += [band, title, *controls, *([status_text] if status_text else []), *buttons]

    # ---------------------------------------------------------------- B. KPI strip
    rev_fc, rev_plan, rev_prior = (facts.row(rev, w) for w in ("fc", "plan", "prior"))
    cards = []
    lines = []
    if has_plan:
        lines.append(signed(f"({rev_fc}) / ({rev_plan}) - 1", "%", "vs plan"))
    if has_prior:
        lines.append(signed(f"({rev_fc}) / ({rev_prior}) - 1", "%", f"vs {prior} actual"))
    cards.append(("kpiRevenue", kpi_body(f"{fy} revenue", money_m(rev_fc), *lines)))
    if gm_row is not None:
        gm_fc, gm_plan = facts.row(gm_row["coefficients"], "fc"), facts.row(gm_row["coefficients"], "plan")
        lines = [span(money_m(gm_fc), INK) + span(" " + esc(sentence(gm_row["label"]).lower()), MUTED)]
        if has_plan:
            lines.append(signed(f"({gm_fc}) / ({rev_fc}) - ({gm_plan}) / ({rev_plan})", " pts", "vs plan"))
        cards.append(("kpiMargin", kpi_body(sentence(gm_row["label"]),
                                            f"{{{{({gm_fc}) / ({rev_fc}) | .1%}}}}", *lines)))
    eb_fc, eb_plan = facts.row(eb_row["coefficients"], "fc"), facts.row(eb_row["coefficients"], "plan")
    lines = [span(f"{{{{({eb_fc}) / ({rev_fc}) | .1%}}}}", INK) + span(" margin", MUTED)]
    if has_plan:
        lines.append(signed(f"({eb_fc}) / ({eb_plan}) - 1", "%", "vs plan"))
    cards.append(("kpiProfit", kpi_body(sentence(eb_row["label"]), money_m(eb_fc), *lines)))
    end = ctx["fy_end"]
    left = f'DateDiff("month", {start}, MakeDate({end[0]}, {end[1]}, 1)) + 1'
    cards.append(("kpiMonths", kpi_body("Forecast months left", f"{{{{{left}}}}} of {ctx['n_fy']}",
                                        span(f"{{{{{month_label(start)}}}}} → "
                                             + f"{{{{{month_label(f'MakeDate({end[0]}, {end[1]}, 1)')}}}}}", MUTED))))
    for cid, body in cards:
        elements += [card(cid), text(cid + "Text", body)]

    # ---------------------------------------------------------------- C. trend + agent panel
    # One bar series (actual + forecast months) coloured by a per-month Actual / Forecast category, and the
    # plan as a dashed line layer (a combo chart colours only its bar layers by the category).
    rev_cond = f'[{FACTS}/Section] = "{ctx["first_positive"]}"'
    columns = [{"id": "tr_m", "name": "Month", "formula": f'DateTrunc("month", [{FACTS}/Month End])',
                "format": MONTH_FMT},
               {"id": "tr_cat", "name": "Period",
                "formula": f"[{FACTS}/Phase]"},
               {"id": "tr_rev", "name": sentence(ctx["first_positive"]),
                "formula": f'Sum(If({rev_cond} and [{FACTS}/Period] != "Plan", [{FACTS}/Amount], Null))',
                "format": CUR_SI}]
    y = [{"columnId": "tr_rev", "type": "bar"}]
    if has_plan:
        columns.append({"id": "tr_plan", "name": "Plan", "format": CUR_SI,
                        "formula": f'Sum(If({rev_cond} and [{FACTS}/Period] = "Plan", [{FACTS}/Amount], Null))'})
        y.append({"columnId": "tr_plan", "type": "line"})
    trend = {"id": "revTrend", "kind": "combo-chart",
             "name": f"{sentence(ctx['first_positive'])}: actual vs forecast" + (" vs plan" if has_plan else ""),
             "source": {"kind": "table", "elementId": "reportFacts"}, "columns": columns,
             "xAxis": {"columnId": "tr_m", "sort": {"direction": "ascending"}},
             "yAxis": {"columnIds": y, "format": {"marks": "grid"}},
             "color": {"by": "category", "column": "tr_cat", "scheme": [PRIMARY, FORECAST_TINT]},
             "legend": {"position": "top"}, "style": dict(CARD)}
    if has_plan:
        trend["seriesLineAreaStyle"] = [{"columnId": "tr_plan", "style": {"line": {"style": "dashed", "width": 2},
                                                                          "points": {"visibility": "hidden"},
                                                                          "missing": "hide"}}]
    agent_panel = card("agentPanel")
    can = (["switch scenario"] if ctx["scenario"] else []) + ["change a driver"] + [
        verb for action, verb in (("roll-forward", "roll forward"), ("snapshot-version", "snapshot"),
                                  ("submit-approval", "submit for approval"))
        if any(b["id"] == dw.BUTTON_IDS[action] for b in wf["buttons"])]
    agent_head = text("agentHead", "**Forecast agent**\n\n" + span(
        "Ask about the forecast, or have it " + (", ".join(can[:-1]) + " or " + can[-1] if len(can) > 1 else can[0])
        + ". Every change asks you first.", MUTED, 12), "top")
    elements += [trend, agent_panel, agent_head]

    # ---------------------------------------------------------------- D. region + P&L summary
    bottom = []
    if region_of and has_plan:
        fc_r = (f'SumIf([{FACTS}/Amount], [{FACTS}/Section] = "{ctx["first_positive"]}" and {facts.cond("fc")})')
        pl_r = (f'SumIf([{FACTS}/Amount], [{FACTS}/Section] = "{ctx["first_positive"]}" and {facts.cond("plan")})')
        var = f"({fc_r}) / ({pl_r}) - 1"
        bottom.append({"id": "regionVariance", "kind": "bar-chart",
                       "name": f"{fy} {sentence(ctx['first_positive']).lower()} forecast vs plan by region",
                       "source": {"kind": "table", "elementId": "reportFacts"},
                       "columns": [{"id": "rv_r", "name": "Region", "formula": f"[{FACTS}/Region]"},
                                   {"id": "rv_v", "name": "Vs plan", "formula": var, "format": SIGNED_PCT},
                                   {"id": "rv_c", "name": "Direction", "formula": var, "format": SIGNED_PCT}],
                       "xAxis": {"columnId": "rv_r", "sort": {"direction": "descending", "by": "rv_v"}},
                       "yAxis": {"columnIds": ["rv_v"], "format": {"marks": "grid"}},
                       "color": {"by": "scale", "column": "rv_c", "scheme": [BAD, HAIRLINE, GOOD],
                                 "domain": {"min": -0.002, "mid": 0, "max": 0.002}},
                       "orientation": "horizontal",
                       "dataLabel": {"labels": "shown", "fontSize": 12, "color": CANVAS, "fontWeight": "bold"},
                       "legend": {"visibility": "hidden"},
                       "style": dict(CARD)})
    S = SUMMARY
    def s_sum(which):
        window = facts._window(ctx["fy_nums"] if which != "prior" else ctx["prior_nums"]).replace(FACTS, S)
        period = {"fc": f'[{S}/Period] != "Plan"', "plan": f'[{S}/Period] = "Plan"',
                  "prior": f'[{S}/Period] = "Actual"'}[which]
        return f"SumIf([{S}/Signed Amount], {period}{window}) / 1000000"
    sum_cols = [{"id": "ps_m", "name": "Line", "formula": f"[{S}/Metric]"},
                {"id": "ps_o", "name": "Metric Order", "formula": f"[{S}/Metric Order]", "hidden": True}]
    values = []
    if has_prior:
        sum_cols.append({"id": "ps_prior", "name": f"{prior} actual", "formula": s_sum("prior"), "format": MILLIONS})
        values.append("ps_prior")
    sum_cols.append({"id": "ps_fc", "name": f"{fy} forecast", "formula": s_sum("fc"), "format": MILLIONS})
    values.append("ps_fc")
    if has_plan:
        sum_cols += [{"id": "ps_plan", "name": "Plan", "formula": s_sum("plan"), "format": MILLIONS},
                     {"id": "ps_vs", "name": "Vs plan", "formula": f"({s_sum('fc')}) / ({s_sum('plan')}) - 1",
                      "format": SIGNED_PCT}]
        values += ["ps_plan", "ps_vs"]
    summary = {"id": "plSummary", "kind": "pivot-table", "name": "P&L summary",
               "source": {"kind": "table", "elementId": "summaryFacts"}, "columns": sum_cols, "values": values,
               "rowsBy": [{"columnId": "ps_m", "sort": {"direction": "ascending", "by": "ps_o", "aggregation": "min"}}],
               "axisTotals": [{"axisId": "grand-row-total", "visibility": "hidden"}],
               "tableStyle": {"preset": "presentation"}, "style": dict(CARD)}
    if has_plan:
        # Coloured by favourability, not sign: a cost line under plan is good news (green).
        favourable = variance_favourability("[Line]", "[Vs plan]", ctx.get("cost_lines") or [])
        summary["conditionalFormats"] = [
            {"type": "single", "columnIds": ["ps_vs"], "condition": "formula", "formula": f"{favourable} < 0",
             "style": {"color": BAD, "backgroundColor": "none"}},
            {"type": "single", "columnIds": ["ps_vs"], "condition": "formula", "formula": f"{favourable} > 0",
             "style": {"color": GOOD, "backgroundColor": "none"}}]
    bottom.append(summary)
    elements += bottom

    # ---------------------------------------------------------------- Assumptions page
    rate, manual = ctx["rate"], ctx["manual"]
    others = [n for n in (ctx["scenario"] or {}).get("names", []) if n != (ctx["scenario"] or {}).get("default")]
    moved = (" and ".join(esc(n) for n in others) + (" move" if len(others) > 1 else " moves")
             + " every growth rate by the source workbook's spread") if others else ""
    scen_line = (span("Scenario: ", MUTED, 13) + span(f"**{{{{[{SCENARIO_CONTROL}]}}}}**", INK, 13) + span(
        " (switch it in the Forecast header" + (f"; {moved}" if moved else "") + ").", MUTED, 13)
        ) if ctx["scenario"] else ""
    manual_note = (" Manual forecast lines are typed in by hand, as in the source workbook, rather than grown "
                   "from a driver." if ctx["manual"] is not None else "")
    asm_head = text("asmHead", "## Growth drivers\n" + span(
        "Monthly growth rate per forecast line. Edit a rate and the forecast recalculates." + manual_note, MUTED, 13)
        + ("\n" + scen_line if scen_line else ""), "top")
    impact = ["**Forecast impact**", span("Recalculates as you edit.", MUTED, 12),
              kpi_body(f"{fy} revenue", money_m(rev_fc), *([signed(f"({rev_fc}) / ({rev_plan}) - 1", "%", "vs plan")]
                                                            if has_plan else [])),
              kpi_body(f"{fy} {sentence(eb_row['label'])}", money_m(eb_fc), *([signed(f"({eb_fc}) / ({eb_plan}) - 1", "%",
                                                                                      "vs plan")] if has_plan else []))]
    asm_impact = text("asmImpact", "\n\n".join(impact), "top")
    asm_els = [asm_head, card("asmImpactCard"), asm_impact]
    if rate is not None:
        asm_els.append(rate)
    if manual is not None:
        asm_els.append(manual)
    # ---------------------------------------------------------------- P&L detail page
    detail_els = [ctx["pivot"], ctx["gl"]]
    # ---------------------------------------------------------------- Versions & approvals page
    figure = {**CUR_SI, "displayNullAs": NULL_DASH}
    fmt = {"FY Revenue": figure, "FY Gross Margin": figure, "FY EBITDA": figure,
           "Saved At": STAMP_FMT, "Submitted At": STAMP_FMT, "Decided At": STAMP_FMT}
    # a migrated row has no forecast start / comment: show a dash / nothing, never the word "null"
    shown = {"Forecast Start Month": lambda ref: f'Coalesce(DateFormat({ref}, "%b-%y"), "{NULL_DASH}")',
             "Comment": lambda ref: f'Coalesce({ref}, "")'}
    # newest first by entry order (dw.LOG_ORDER): migrated history keeps the source's dates
    label = {"Forecast Start Month": "Forecast start", "FY Revenue": f"{fy} revenue",
             "FY Gross Margin": "Gross margin", "FY EBITDA": "EBITDA"}
    def log_view(view_id, table, names, title):
        cols = [{"id": f"{view_id}_{i}", "name": label.get(n, sentence(n)),
                 "formula": shown.get(n, lambda ref: ref)(f"[{table['name']}/{n}]"),
                 **({"format": fmt[n]} if n in fmt else {})}
                for i, n in enumerate(names) if any(c["name"] == n for c in table["columns"])]
        cols.append({"id": f"{view_id}_order", "name": dw.LOG_ORDER, "formula": f"[{table['name']}/{dw.LOG_ORDER}]",
                     "hidden": True})
        return {"id": view_id, "kind": "table", "name": title, "source": {"kind": "table", "elementId": table["id"]},
                "columns": cols, "sort": [{"columnId": f"{view_id}_order", "direction": "descending"}],
                "tableStyle": {"preset": "presentation"}, "tableComponents": {"summaryBar": "hidden"},
                "style": dict(CARD)}
    logs = {t["id"]: t for t in wf["tables"]}
    log_views = []
    if dw.VERSIONS in logs:
        log_views.append(log_view("versionsView", logs[dw.VERSIONS],
                                  ["Version", "Scenario", "Forecast Start Month", "FY Revenue", "FY Gross Margin",
                                   "FY EBITDA", "Saved By", "Saved At", "Comment"], "Versions"))
    if dw.APPROVALS in logs:
        log_views.append(log_view("approvalsView", logs[dw.APPROVALS],
                                  ["Status", "Request", "Submitted By", "Submitted At", "Decided By", "Decided At",
                                   "Comment"], "Approvals"))
    history = {el: n for el, n in (ctx.get("history") or {}).items() if el in logs}
    both = dw.VERSIONS in logs and dw.APPROVALS in logs
    counts = _join_and([f"{history[el]} {noun}" for el, noun in ((dw.VERSIONS, "versions"),
                                                                   (dw.APPROVALS, "approvals")) if el in history])
    origin = ((f" The first {counts} are the history migrated from the source workbook"
               + (" (it did not freeze the figures, so those cells are blank)." if dw.VERSIONS in history else "."))
              if history else (" Both logs start" if both else " The log starts") + " empty in a new workbook.")
    what = [s for el, s in ((dw.VERSIONS, "Snapshot freezes the headline numbers as a version"),
                            (dw.APPROVALS, "Submit for approval logs a request")) if el in logs]
    log_title = " & ".join(t for el, t in ((dw.VERSIONS, "Versions"), (dw.APPROVALS, "approvals" if both else "Approvals"))
                           if el in logs)
    log_head = text("logHead", f"## {log_title}\n" + span(
        "; ".join(what) + ". " + ("Both are" if both else "It is") + " written by the "
        + ("buttons" if both else "button") + " on the Forecast page (or by the forecast agent, after you "
        "approve), with who and when, newest first." + origin, MUTED, 13), "top") if logs else None

    # Approve / Send back: beside the Approvals table (only with --approvers), same size, Approve primary.
    decision_els = []
    if wf.get("decisions") and dw.APPROVALS in logs:
        for b in wf["decisions"]:
            primary = b["id"] == dw.DECISION_IDS["approve"]
            decision_els.append({**b, "size": "medium", "align": "stretch", "borderRadius": "round",
                                 "appearance": "filled" if primary else "outline",
                                 **({"fillColor": PRIMARY} if primary else {})})
        decision_els.append(text("decisionNote", span("Only approvers can approve: "
                                                      + esc(", ".join(wf["approvers"])), MUTED, 12)))
    pages = [{"id": "pgForecast", "name": "Forecast", "elements": elements},
             {"id": "pgAssumptions", "name": "Assumptions", "elements": asm_els},
             {"id": "pgDetail", "name": "P&L detail", "elements": detail_els},
             *([{"id": "pgVersions", "name": PAGE_NAMES["pgVersions"] if log_title == "Versions & approvals"
                  else log_title, "elements": [log_head, *log_views, *decision_els]}] if log_views else []),
             {"id": "model", "name": "Model (sources)", "visibility": "hidden",
              "elements": hidden + ([ctx["assump"]] if ctx["assump"] else [])}]
    layout = app_layout(pages, controls=bool(controls), cards=[c for c, _ in cards], bottom=[b["id"] for b in bottom],
                        rate=rate is not None, manual=manual is not None,
                        buttons=[b["id"] for b in buttons], status=status_text is not None,
                        logs=[v["id"] for v in log_views], decisions=[e["id"] for e in decision_els])
    theme = {"theme": {"overrides": {"categoricalScheme": [PRIMARY, FORECAST_TINT, PLAN_GREY, "#14B8A6", "#F59E0B",
                                                           "#8B5CF6"]}}}
    return {"pages": pages, "layout": layout, "settings": theme}


AGENT_FACTS_ID = "agentFacts"           # lib/agent_enhance.py AGENT_FACTS: the agent's fiscal-year source


def agent_facts_element(ctx, has_plan, has_prior):
    """The forecast agent's one fiscal-year source (hidden): every month's amount by Fiscal year, Period
    (Actual / Forecast / Plan), Section, Line item and Region, read from ``Forecast Facts`` so it follows the
    one boundary like the KPI cards. A fiscal year's figure is its Actual + Forecast rows (exactly the KPI
    card's sum); its plan is its Plan rows. A closed month still awaiting actuals counts as Forecast (its
    values are the forecast's). Without it the agent saw only the forecast window and called that total
    the fiscal year, and had no plan to compare with."""
    num = f"[{FACTS}/Month Num]"
    windows = [(ctx["fy_nums"], ctx["fy_label"])] + ([(ctx["prior_nums"], ctx["prior_label"])] if has_prior else [])
    branches = ", ".join(f'{num} >= {nums[0]} and {num} <= {nums[-1]}, "{esc(label)}"'
                         for nums, label in windows if nums)
    fiscal = (f"If({branches}, Null)" if branches
              else f'Text(DatePart("year", [{FACTS}/Month End]))')
    period = f'If([{FACTS}/Period] = "Plan", "Plan", [{FACTS}/Period] = "Actual", "Actual", "Forecast")'
    columns = [{"id": "af_fy", "name": "Fiscal year", "formula": fiscal},
               {"id": "af_m", "name": "Month", "formula": f'DateTrunc("month", [{FACTS}/Month End])',
                "format": MONTH_FMT},
               {"id": "af_p", "name": "Period", "formula": period},
               {"id": "af_sec", "name": "Section", "formula": f"[{FACTS}/Section]"},
               {"id": "af_line", "name": "Line item", "formula": f"[{FACTS}/Line Item]"},
               {"id": "af_region", "name": "Region", "formula": f"[{FACTS}/Region]"},
               {"id": "af_amt", "name": "Amount", "formula": f"[{FACTS}/Amount]", "format": CUR0}]
    if ctx["scenario"]:
        columns.append({"id": "af_scn", "name": "Active scenario",
                        "formula": f'Coalesce([{SCENARIO_CONTROL}], "{esc(ctx["scenario"]["default"])}")'})
    return {"id": AGENT_FACTS_ID, "kind": "table", "name": "Monthly P&L facts",
            "source": {"kind": "table", "elementId": "reportFacts"}, "columns": columns}


def _el(element_id, c0, c1, r0, r1, indent="  "):
    return f'{indent}<Element elementId="{element_id}" gridColumn="{c0} / {c1}" gridRow="{r0} / {r1}"/>'


def _container(element_id, c0, c1, r0, r1, children):
    head = (f'  <Container elementId="{element_id}" type="grid" gridColumn="{c0} / {c1}" gridRow="{r0} / {r1}" '
            f'gridTemplateColumns="repeat(24, 1fr)" gridTemplateRows="auto">')
    return "\n".join([head, *[_el(*child, indent="    ") for child in children], "  </Container>"])


AGENT_PANEL = "agentPanel"
# Header action group: adjacent cells of the band's grid, right-aligned, sized by label length, primary
# last and flush with the band's right edge. (A container nested in the band is not used: the server's
# layout check misplaces a container inside a container and moves elements.)
ACTION_SPANS = {"btnRoll": (18, 20), "btnSnapshot": (20, 22), "btnSubmit": (22, 25)}
# Series colours a workbook POST cannot carry (the spec has no per-series colour on a combo chart):
# (elementId, series columnId, hex), applied by the setup program through Code Mode set_series_color.
SERIES_COLORS = (("revTrend", "tr_plan", "#475569"),)
AGENT_CHAT_ROWS = (4, 22)        # the chat's rows inside the agent panel (merge_fragment fills the slot)


def app_layout(pages, *, controls, cards, bottom, rate, manual, buttons, status=True,
               logs=("versionsView", "approvalsView"), decisions=()):
    blocks = []
    def page(page_id, lines):
        blocks.append("\n".join([f'<Page type="grid" gridTemplateColumns="repeat(24, 1fr)" gridTemplateRows="auto" '
                                 f'id="{page_id}">', *lines, "</Page>"]))
    # Forecast: header (rows 1-5), KPIs (5-12), trend + agent (12-34), region + summary (34-52)
    head = [("hdrTitle", 1, 8, 1, 5)]
    status_at = 8
    if controls:
        head.append(("ctlScenario", 8, 13, 1, 5))
        status_at = 13
    if status:
        head.append(("hdrStatus", status_at, status_at + 3, 1, 5))   # the pill sits right after the switch
    head += [(b, *ACTION_SPANS[b], 1, 5) for b in buttons]   # one right-aligned toolbar
    lines = [_container("hdrBand", 1, 25, 1, 5, head)]
    width = 24 // len(cards)
    for i, cid in enumerate(cards):
        lines.append(_container(cid, 1 + i * width, 1 + (i + 1) * width, 5, 11, [(cid + "Text", 1, 25, 1, 6)]))
    lines.append(_el("revTrend", 1, 17, 11, 33))
    lines.append(_container(AGENT_PANEL, 17, 25, 11, 33, [("agentHead", 1, 25, 1, AGENT_CHAT_ROWS[0])]))
    if len(bottom) == 2:
        lines += [_el(bottom[0], 1, 13, 33, 49), _el(bottom[1], 13, 25, 33, 49)]
    else:
        lines += [_el(bottom[0], 1, 25, 33, 49)]
    page("pgForecast", lines)
    # Assumptions: the content-sized input tables side by side, then the impact card, one row
    lines = [_el("asmHead", 1, 25, 1, 5)]
    at = 1
    for element_id, span_cols, present in (("rateInput", 9, rate), ("manualInput", 8, manual)):
        if present:
            lines.append(_el(element_id, at, at + span_cols, 5, 19))
            at += span_cols
    lines.append(_container("asmImpactCard", at, 25, 5, 19, [("asmImpact", 1, 25, 1, 14)]))
    page("pgAssumptions", lines)
    page("pgDetail", [_el("pnlPivot", 1, 25, 1, 41), _el("glTable", 1, 25, 41, 63)])
    # tall enough for the migrated history plus a few new rows without an inner scroll
    if logs:
        rows = {"versionsView": 24, "approvalsView": 14}
        lines, at = [_el("logHead", 1, 25, 1, 5)], 5
        for view in logs:
            if view == "approvalsView" and decisions:    # Approve / Send back + caption, above the Approvals table
                buttons_at = {"btnApprove": (1, 4), "btnReturn": (4, 7), "decisionNote": (7, 25)}
                lines += [_el(d, *buttons_at[d], at, at + 3) for d in decisions]
                at += 3
            lines.append(_el(view, 1, 25, at, at + rows[view]))
            at += rows[view]
        page("pgVersions", lines)
    hidden = next(p for p in pages if p["id"] == "model")
    page("model", [_el(e["id"], 1, 25, 1 + i * 12, 13 + i * 12) for i, e in enumerate(hidden["elements"])])
    return '<?xml version="1.0" encoding="utf-8"?>\n' + "\n".join(blocks)
