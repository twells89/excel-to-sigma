"""Demo macros -> native Sigma workflows (input tables written by button actions).

``build_workflow_manifest`` classifies a macro module (``.bas`` source or a real
``.xlsm``; VBA is only read, never executed) and routes each macro to a native
workflow or records a disposition. ``workflow_elements`` turns the manifest into
workbook elements: only the Parameters / Versions / Approvals input tables its
workflows target, one button per button workflow and (with Submit) a
latest-approval-status table. ``emit_manual_commit_program`` writes the
post-create Code Mode program that enables manual commit on every input table
(backend actions, e.g. a Sigma Agent tool, only write to such tables; a REST
spec POST never sets the flag) and seeds the Parameters row and any other
editable tables it is given.

Nothing here sends email: the email step of a Submit macro is replaced by an
Approvals row and recorded under ``dispositions``. The Approvals log is an editable
input table (a governed log, not an access control).
"""
from __future__ import annotations

import importlib.util
import re
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]

PARAMETERS, VERSIONS, APPROVALS = "parameters", "versions", "approvals"
STATUS_KPI = "approvalStatus"
BUTTON_IDS = {"roll-forward": "btnRoll", "snapshot-version": "btnSnapshot", "submit-approval": "btnSubmit"}
EFFECTS = ("update-rows", "insert-rows", "control")

# elementId -> (display name, [(columnId, display name, type)])
# System columns (CREATED_BY / CREATED_AT) are deliberately not declared: an agent
# query that names them bare hits "Ambiguous column reference". ID is hidden and
# renamed "Row ID" so no bare ``ID`` is exposed either.
# ``Log Order`` (hidden) numbers the log rows in the order they were written: the migrated history
# carries the source's own dates (which can lie after today), so "newest first" and the latest
# approval status follow the entry order, not the clock.
LOG_ORDER = "Log Order"
TABLES = {
    PARAMETERS: ("Parameters", [("pa_scn", "Selected Scenario", "text"),
                                ("pa_start", "Forecast Start Month", "datetime")]),
    VERSIONS: ("Versions", [("ve_ver", "Version", "text"), ("ve_scn", "Scenario", "text"),
                            ("ve_start", "Forecast Start Month", "datetime"),
                            ("ve_rev", "FY Revenue", "number"), ("ve_gm", "FY Gross Margin", "number"),
                            ("ve_ebitda", "FY EBITDA", "number"),
                            ("ve_by", "Saved By", "text"), ("ve_at", "Saved At", "datetime"),
                            ("ve_note", "Comment", "text"), ("ve_seq", LOG_ORDER, "number")]),
    APPROVALS: ("Approvals", [("ap_req", "Request", "text"), ("ap_by", "Submitted By", "text"),
                              ("ap_to", "Approver", "text"), ("ap_status", "Status", "text"),
                              ("ap_at", "Submitted At", "datetime"), ("ap_note", "Comment", "text"),
                              ("ap_seq", LOG_ORDER, "number")]),
}
HIDDEN_COLUMNS = {"ve_seq", "ap_seq"}
COLUMN_IDS = {(el, name): cid for el, (_n, cols) in TABLES.items() for cid, name, _t in cols}
# The forecast start month a workflow reads when the workbook defines no single boundary element.
DEFAULT_START = "Max([Parameters/Forecast Start Month])"

ROUTES = (
    # (name regex, classes allowed, workflow key)
    (re.compile(r"roll\s*forward", re.I), ("ACTION_WRITEBACK", "TRANSFORM_ROLLUP"), "roll-forward"),
    (re.compile(r"snapshot", re.I), ("ACTION_WRITEBACK",), "snapshot-version"),
    (re.compile(r"submit|approval", re.I), ("ACTION_WRITEBACK", "EXTERNAL_IO"), "submit-approval"),
    (re.compile(r"scenario", re.I), ("WHATIF_SCENARIO",), "set-scenario"),
)
WORKFLOWS = {
    "set-scenario": ("Set scenario", "control", PARAMETERS, ["Selected Scenario"], False),
    "roll-forward": ("Roll forward one month", "update-rows", PARAMETERS, ["Forecast Start Month"], True),
    "snapshot-version": ("Snapshot version", "insert-rows", VERSIONS,
                         ["Version", "Scenario", "Forecast Start Month", "FY Revenue", "FY Gross Margin",
                          "FY EBITDA", "Saved By", "Saved At"], True),
    "submit-approval": ("Submit for approval", "insert-rows", APPROVALS,
                        ["Request", "Submitted By", "Approver", "Status", "Submitted At"], True),
}
# The refresh disposition names the Snowflake landing only when this run landed the source Tables.
REFRESH_NOTES = {
    False: "The macro's external data reload is replaced by the live warehouse data Sigma reads on every "
           "query; no refresh button is needed.",
    True: "The macro's external data reload is replaced by the source Tables landed in Snowflake (Phase 1b), "
          "which Sigma reads live; no refresh button is needed.",
}
DISPOSITIONS = {
    "EXTERNAL_DATA_REFRESH": ("replaced-by-live-warehouse", REFRESH_NOTES[False]),
    "RECALC": ("replaced-by-live-calculation",
        "Sigma recalculates on read; there is nothing to port."),
    "FORMAT_COSMETIC": ("styling",
        "Cell formatting is carried by workbook styling and number formats, not an action."),
}
NOTES = {
    "roll-forward": "Advances Parameters.Forecast Start Month by one month (update-rows, DateAdd): a month "
                    "close. The actual / forecast split follows it, so the closed month reads the landed "
                    "actuals and the remaining months re-forecast from it. Replaces the VBA that copied the "
                    "first forecast month into the Actuals table.",
    "snapshot-version": "Appends ONE Versions row freezing the headline numbers at click time (version label, "
                        "scenario, forecast start month, FY revenue / gross margin / EBITDA taken from the same "
                        "formulas the KPI tiles use, user, time). The VBA pasted the whole forecast grid as "
                        "values; the full line x month grid is not reproduced.",
    "submit-approval": "Appends an Approvals row (status Submitted, user, time). The VBA also protected the "
                       "sheet and sent an Outlook email; neither is reproduced.",
    "set-scenario": "The planning app's Downside / Base / Upside switch is a workbook control the forecast "
                    "formulas read (set-scenario sets it; Snapshot records its value). Without the planning "
                    "app, the one-row Parameters input table holds Selected Scenario instead.",
}


class WorkflowError(ValueError):
    pass


def _classifier():
    spec = importlib.util.spec_from_file_location("macro_classify_for_workflows", SCRIPTS / "macro-classify.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def classify_macros(path):
    """Per-procedure ``{name, class, signals}`` for a .bas / .xlsm source (read only)."""
    classifier = _classifier()
    code = classifier.extract_vba(str(path))
    if not code.strip():
        raise WorkflowError(f"no VBA macros found in {path}")
    out = []
    for name, _kind, body in classifier.split_procedures(code):
        cls, _status, _target, _note, signals = classifier.classify(name, body)
        out.append({"name": name, "class": cls, "signals": list(signals)})
    return out


def build_workflow_manifest(path, landed=False):
    """Route every macro to a native workflow or a recorded disposition (never silent). ``landed``: this
    run landed the source Tables in Snowflake (Phase 1b), so the refresh disposition can say so."""
    workflows, dispositions = [], []
    for macro in classify_macros(path):
        name, cls = macro["name"], macro["class"]
        route = next((key for rx, classes, key in ROUTES if rx.search(name) and cls in classes), None)
        if route:
            label, effect, element, columns, writes = WORKFLOWS[route]
            workflows.append({"macro": name, "class": cls, "actionId": route, "label": label,
                              "effect": effect, "target": {"elementId": element, "columns": list(columns)},
                              "writes": writes, "note": NOTES[route]})
            if route == "submit-approval" and any(re.search(r"outlook|sendmail", s, re.I) for s in macro["signals"]):
                dispositions.append({"macro": name, "class": cls, "disposition": "email-replaced-never-sent",
                    "note": "The email the macro sent is replaced by the Approvals row and the latest-status "
                            "header; no email is sent by the workflow, the button or the agent."})
        elif cls in DISPOSITIONS:
            kind, note = DISPOSITIONS[cls]
            if cls == "EXTERNAL_DATA_REFRESH":
                note = REFRESH_NOTES[bool(landed)]
            dispositions.append({"macro": name, "class": cls, "disposition": kind, "note": note})
        else:
            dispositions.append({"macro": name, "class": cls, "disposition": "unmapped-needs-review",
                "note": "No native workflow mapping; a human must review before migration is complete."})
    return {"version": 1, "workflows": workflows, "dispositions": dispositions}


APPROVER = "Approver"   # generic: the source's approver role is not inferred from the macro
STATUS_FORMULA = (f'Coalesce(Lookup([Approvals/Status], Max([Approvals/{LOG_ORDER}]), '
                  f'[Approvals/{LOG_ORDER}]), "Not submitted")')


def next_entry(table):
    """The next ``Log Order`` of a log table: one past the last entry (1 for an empty log)."""
    return f"Coalesce(Max([{table}/{LOG_ORDER}]), 0) + 1"


def _const(text):
    return {"type": "constant", "value": {"type": "text", "value": text}}


def _formula(text):
    return {"type": "formula", "formula": text}


def _values(element, formulas):
    return {COLUMN_IDS[(element, name)]: value for name, value in formulas.items()}


def action_effects(action_id, request_label, figures=None, scenario_control=None, start=DEFAULT_START):
    """The effect list for one workflow (live-proven working shapes).

    With ``scenario_control`` (the controlId of the workbook's scenario switch) Snapshot records the
    control's current value; otherwise the Parameters row's Selected Scenario. ``start`` is the formula
    of the current forecast start month (the planning app passes its one boundary element's)."""
    if action_id == "roll-forward":
        return [{"effect": "update-rows", "tableElementId": PARAMETERS,
                 "whichRows": {"type": "formula", "formula": "true"},
                 "values": _values(PARAMETERS, {"Forecast Start Month": _formula(
                     f'DateAdd("month", 1, {start})')})}]
    if action_id == "snapshot-version":
        if not figures:
            raise WorkflowError("snapshot needs the KPI figure formulas (figures=)")
        values = {
            # Log Order is a number column: Text() of the bare sum renders "v8.0000000000000000000".
            "Version": _formula(f'Concat("v", Text(Int({next_entry("Versions")})))'),
            "Scenario": ({"type": "control", "control": scenario_control} if scenario_control
                         else _formula("Max([Parameters/Selected Scenario])")),
            "Forecast Start Month": _formula(start),
            "Saved By": _formula("CurrentUserEmail()"), "Saved At": _formula("Now()"),
            LOG_ORDER: _formula(next_entry("Versions"))}
        names = {"revenue": "FY Revenue", "grossMargin": "FY Gross Margin", "ebitda": "FY EBITDA"}
        for key, formula in figures.items():
            values[names[key]] = _formula(formula)
        return [{"effect": "insert-rows", "tableElementId": VERSIONS, "values": _values(VERSIONS, values)}]
    if action_id == "submit-approval":
        return [{"effect": "insert-rows", "tableElementId": APPROVALS, "values": _values(APPROVALS, {
            "Request": _const(request_label), "Submitted By": _formula("CurrentUserEmail()"),
            "Approver": _const(APPROVER), "Status": _const("Submitted"), "Submitted At": _formula("Now()"),
            LOG_ORDER: _formula(next_entry("Approvals"))})}]
    raise WorkflowError(f"no button action for {action_id}")


def input_table(element, connection_id, mode, figures=None, scenario_control=None):
    name, cols = TABLES[element]
    if element == PARAMETERS and scenario_control:
        cols = [c for c in cols if c[1] != "Selected Scenario"]   # the scenario switch holds it
    if element == VERSIONS and figures is not None:
        keep = {"FY Revenue": "revenue", "FY Gross Margin": "grossMargin", "FY EBITDA": "ebitda"}
        cols = [c for c in cols if c[1] not in keep or keep[c[1]] in figures]
    return {"id": element, "kind": "input-table", "name": name, "inputMode": mode,
            "source": {"kind": "empty", "connectionId": connection_id},
            "columns": [{"id": "ID", "name": "Row ID", "hidden": True},
                        *[{"id": cid, "type": typ, "name": nm, **({"hidden": True} if cid in HIDDEN_COLUMNS else {})}
                          for cid, nm, typ in cols]]}


def workflow_elements(manifest, connection_id, request_label, figures=None, scenario_control=None,
                      no_status="Not submitted", start=DEFAULT_START):
    """Workbook elements for a manifest: ``{tables, buttons, status}``.

    ``figures`` maps revenue / grossMargin / ebitda to the workbook's KPI formulas;
    Snapshot freezes exactly those into its Versions row. ``scenario_control`` is the controlId of
    a scenario switch that the forecast reads: Parameters then carries only the forecast start
    month and Snapshot records the switch's value. ``no_status`` is the status shown before any
    request was submitted. ``start`` is the forecast start month formula Roll forward advances and
    Snapshot records.
    """
    if not isinstance(manifest, dict) or manifest.get("version") != 1:
        raise WorkflowError("workflows manifest version 1 required (a JSON object from demo-manifest)")
    if not isinstance(manifest.get("workflows", []), list) or not all(
            isinstance(w, dict) and "actionId" in w and "effect" in w for w in manifest.get("workflows", [])):
        raise WorkflowError("workflows manifest: 'workflows' must be a list of workflow objects")
    if figures is None and any(w["actionId"] == "snapshot-version" for w in manifest.get("workflows", [])):
        figures = {"revenue": "Sum([PnL Combined/Amount])"}  # structural default (checks); builder passes real KPIs
    items = {w["actionId"]: w for w in manifest.get("workflows", [])}
    for w in items.values():
        if w["effect"] not in EFFECTS:
            raise WorkflowError(f"unsupported workflow effect {w['effect']}")
    buttons = []
    for action_id, element_id in BUTTON_IDS.items():
        if action_id not in items:
            continue
        w = items[action_id]
        buttons.append({"id": element_id, "kind": "button", "text": w["label"],
            "appearance": "outline" if action_id == "submit-approval" else "filled",
            "actions": [{"id": "a-" + action_id, "trigger": "on-click",
                         "effects": action_effects(action_id, request_label, figures, scenario_control,
                                                   start)}]})
    # Only the tables the workflows target: Parameters for set-scenario / roll-forward (or when an effect
    # reads its row: Snapshot without a scenario switch or boundary), Versions for Snapshot, Approvals and
    # the status for Submit.
    effects_text = repr([b["actions"] for b in buttons])
    wanted = {PARAMETERS: bool({"set-scenario", "roll-forward"} & set(items))
                          or f"[{TABLES[PARAMETERS][0]}/" in effects_text,
              VERSIONS: "snapshot-version" in items, APPROVALS: "submit-approval" in items}
    tables = [input_table(PARAMETERS, connection_id, "edit", scenario_control=scenario_control)
              if el == PARAMETERS else
              input_table(VERSIONS, connection_id, "view", figures) if el == VERSIONS else
              input_table(APPROVALS, connection_id, "view")
              for el in (PARAMETERS, VERSIONS, APPROVALS) if wanted[el]]
    status = None
    if wanted[APPROVALS]:
        # Latest status = Status of the Approvals row with the last Log Order, held by a one-row table
        # (over the one-row Parameters table when the workbook has it: the shape proven live; otherwise
        # over a one-row VALUES query, the shape of the planning app's boundary element).
        source = ({"kind": "table", "elementId": PARAMETERS} if wanted[PARAMETERS] else
                  {"kind": "sql", "connectionId": connection_id,
                   "statement": "select * from (values\n(1)\n) as t(STATUS_ROW)"})
        status = {"id": STATUS_KPI, "kind": "table", "name": "Latest approval status",
            "source": source,
            "columns": [{"id": "st_v", "name": "Latest approval status",
                         "formula": STATUS_FORMULA.replace('"Not submitted"', '"' + no_status + '"')}]}
    return {"tables": tables, "buttons": buttons, "status": status}


# Migrated history: the source workbook's version / approval log tables -> the Versions / Approvals
# input tables (seeded once by the setup program). Workbook column -> accepted source headers.
HISTORY_TABLES = {
    VERSIONS: (("tblversions", "versions"),
               {"Version": ("version",), "Scenario": ("scenario",), "Saved By": ("saved by", "author"),
                "Saved At": ("saved at", "saved on", "date"), "Comment": ("comment", "comments", "notes")},
               ("Version", "Saved At")),
    APPROVALS: (("tblapprovals", "approvals"),
                {"Request": ("request", "version"), "Submitted By": ("submitted by",),
                 "Approver": ("approver",), "Status": ("status",),
                 "Submitted At": ("submitted at", "submitted on", "date"),
                 "Comment": ("comment", "comments", "notes")},
                ("Status", "Submitted At")),
}


def _seed_value(value, typ):
    from datetime import date, datetime
    if value in (None, ""):
        return None
    if typ == "datetime":
        if isinstance(value, datetime):
            return value.date().isoformat() if value.time() == datetime.min.time() else value.isoformat(" ")
        if isinstance(value, date):
            return value.isoformat()
        raise WorkflowError(f"history date {value!r} is not a date")
    if typ == "number":
        return float(value)
    return str(value).strip()


def source_history(path):
    """``[(elementId, [columnId, ...], rows)]`` seeds for Versions / Approvals from the source workbook's own
    version and approval log tables (``tblVersions`` / ``tblApprovals``), so the migrated workbook keeps its
    history instead of starting with empty logs. Rows keep the source's order (by date, stable) and get
    ``Log Order`` 1..n; a column the source does not record (the frozen FY figures) stays unset. A source
    without such a table yields no seed; a table whose required columns cannot be found raises."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)    # read only: the source is never saved
    found = {}
    for ws in wb.worksheets:
        for name, ref in ws.tables.items():
            ref = ref if isinstance(ref, str) else ref.ref
            found[name.lower()] = [[c.value for c in row] for row in ws[ref]]
    out = []
    for element, (names, aliases, required) in HISTORY_TABLES.items():
        grid = next((found[n] for n in names if n in found), None)
        if not grid:
            continue
        headers = [str(h or "").strip().lower() for h in grid[0]]
        where = {}
        for column, accepted in aliases.items():
            at = next((headers.index(a) for a in accepted if a in headers), None)
            if at is not None:
                where[column] = at
        missing = [c for c in required if c not in where]
        if missing:
            raise WorkflowError(f"source {element} table has no {', '.join(missing)} column "
                                f"(headers: {', '.join(h for h in headers if h)})")
        types = {nm: typ for _cid, nm, typ in TABLES[element][1]}
        stamp = required[1]
        body = [r for r in grid[1:] if any(v not in (None, "") for v in r)]
        body = sorted(enumerate(body), key=lambda ir: (_seed_value(ir[1][where[stamp]], "datetime") or "", ir[0]))
        columns = [COLUMN_IDS[(element, c)] for c in where] + [COLUMN_IDS[(element, LOG_ORDER)]]
        rows = [[_seed_value(r[where[c]], types[c]) for c in where] + [i]
                for i, (_at, r) in enumerate(body, 1)]
        out.append((element, columns, rows))
    return out


def _input_tables(spec):
    """``(pageId, elementId)`` for every input table of a (nested or wired) workbook spec."""
    import code_rep
    doc = code_rep.document(spec)
    owners = code_rep.workbook_page_by_element(doc)
    tables = [e["id"] for e in code_rep.workbook_elements(doc) if e.get("kind") == "input-table"]
    missing = [t for t in tables if t not in owners]
    if missing:
        raise WorkflowError("input tables not placed on a page by the layout: " + ", ".join(missing))
    return [(owners[t]["id"], t) for t in tables]


def input_column_display(spec):
    """``{elementId: [(columnId, d3 pattern or None, hidden)]}`` for typed input-table columns.

    A workbook POST keeps an input table's typed columns but drops their ``format`` and ``hidden``
    (live readback, 2026-11); the setup program re-applies them through Code Mode
    (``set_format_pattern`` / ``set_visible(False)``) so a rate shows as 1.0%, not 0.01."""
    import code_rep
    out = {}
    for el in code_rep.workbook_elements(code_rep.document(spec)):
        if el.get("kind") != "input-table":
            continue
        for col in el.get("columns", []):
            if not col.get("type"):
                continue        # system columns (ID, ...) keep their spec fields
            fmt = (col.get("format") or {}).get("formatString")
            if fmt or col.get("hidden"):
                out.setdefault(el["id"], []).append((col["id"], fmt, bool(col.get("hidden"))))
    return out


def input_table_style_lines(spec):
    """``{elementId: [program lines]}`` re-applying an input table's spec ``tableStyle`` through Code Mode
    (``table_style()``): preset first, since switching preset resets grid lines and autofit."""
    import code_rep
    out = {}
    for el in code_rep.workbook_elements(code_rep.document(spec)):
        style = el.get("tableStyle") if el.get("kind") == "input-table" else None
        if not style:
            continue
        lines = ["s = t.table_style()"]
        if style.get("preset"):
            lines.append(f"s.set_preset({style['preset']!r})")
        if style.get("gridLines"):
            lines.append(f"s.set_grid_lines({style['gridLines']!r})")
        if style.get("autofitColumns"):
            lines.append(f"s.set_autofit_columns({style['autofitColumns'] == 'shown'})")
        out[el["id"]] = lines
    return out


# Content-sized input-table columns (an input table with autofit off): about 7 px per character of the
# longest seed value or the header, plus room for the header's edit pencil and the cell padding.
CHAR_PX, HEADER_EXTRA_CHARS, PAD_PX = 7, 3, 28
MIN_COLUMN_PX, MAX_COLUMN_PX = 120, 360


def _shown(value, fmt):
    """Roughly how a seed value displays under its d3 format (only its length matters)."""
    if isinstance(value, (int, float)) and fmt:
        if "%" in fmt:
            return f"{value * 100:.1f}%"
        if "$" in fmt:
            return f"${value:,.0f}"
    return "" if value is None else str(value)


def input_column_widths(spec, table_seeds=()):
    """``{elementId: {columnId: px}}`` for the visible typed columns of every input table whose spec
    ``tableStyle`` turns autofit off: sized to the longest of the header (plus its edit pencil) and the
    seed values, so a number column is not stretched across the table (its pencil far left, its
    right-aligned name far right reads as an empty unlabelled column). ``table_seeds`` is
    ``[(elementId, [columnId, ...], rows)]`` with each row a list in that column order."""
    import code_rep
    seeds = {el: (list(cols), list(rows)) for el, cols, rows in table_seeds}
    out = {}
    for el in code_rep.workbook_elements(code_rep.document(spec)):
        if el.get("kind") != "input-table" or (el.get("tableStyle") or {}).get("autofitColumns") != "hidden":
            continue
        cols, rows = seeds.get(el["id"], ([], []))
        widths = {}
        for col in el.get("columns", []):
            if not col.get("type") or col.get("hidden"):
                continue
            fmt = (col.get("format") or {}).get("formatString")
            chars = len(col.get("name") or col["id"]) + HEADER_EXTRA_CHARS
            if col["id"] in cols:
                i = cols.index(col["id"])
                chars = max([chars] + [len(_shown(row[i], fmt)) for row in rows])
            widths[col["id"]] = max(MIN_COLUMN_PX, min(MAX_COLUMN_PX, chars * CHAR_PX + PAD_PX))
        out[el["id"]] = widths
    return out


MAX_SEED_ROWS = 2000  # rows per input table in one program's single upsert_rows batch


def _table_seed_lines(page_id, element_id, columns, rows):
    """One typed-checked ``upsert_rows`` call carrying every row; a ``None`` value is left unset."""
    var = {cid: "c_" + re.sub(r"\W", "_", cid).lower() for cid in columns}
    out = [f"# Seed {element_id}: {len(rows)} row(s), one batch (fresh workbook only).",
           f"t = sigma.workbook.get_page('{page_id}').get_element('{element_id}')",
           "if not isinstance(t, InputTable):",
           f"    raise RuntimeError('{element_id} is not an InputTable')"]
    for cid in columns:
        out += [f"{var[cid]} = t.get_column('{cid}')",
                f"if not isinstance({var[cid]}, LiteralEditableColumn):",
                f"    raise RuntimeError('{element_id}.{cid} is not an editable column')"]
    out.append("t.upsert_rows([")
    for row in rows:
        if len(row) != len(columns):
            raise WorkflowError(f"{element_id} seed row {row!r} does not match columns {list(columns)}")
        cells = ", ".join(f"{var[cid]}: {value!r}" for cid, value in zip(columns, row) if value is not None)
        out.append(f"    {{{cells}}},")
    return out + ["])", ""]


def _series_color_lines(spec, series_colors):
    import code_rep
    doc = code_rep.document(spec)
    owners = code_rep.workbook_page_by_element(doc)
    by_id = {e["id"]: e for e in code_rep.workbook_elements(doc)}
    kinds = {k: e.get("kind") for k, e in by_id.items()}
    cls = {"combo-chart": "ComboChart", "bar-chart": "BarChart", "line-chart": "LineChart"}
    out = []
    for element_id, column_id, color in series_colors:
        if element_id not in owners or kinds.get(element_id) not in cls:
            continue
        if not any(c.get("id") == column_id for c in by_id[element_id].get("columns", [])):
            continue
        out += [f"# Series colour the spec cannot carry: {element_id}.{column_id}",
                f"ch = sigma.workbook.get_page('{owners[element_id]['id']}').get_element('{element_id}')",
                f"if not isinstance(ch, {cls[kinds[element_id]]}):",
                f"    raise RuntimeError('{element_id} is not a {cls[kinds[element_id]]}')",
                "for s in ch.list_series():",
                f"    if s.id() == '{column_id}':",
                f"        ch.set_series_color(s, '{color}')", ""]
    return out


def emit_manual_commit_program(spec, seed=None, table_seeds=(), series_colors=(), column_widths=None):
    """Post-create Code Mode program text: re-set each input table's edit mode so the
    host publishes it with manual commit enabled, then seed rows.

    ``seed`` (``{"scenario", "start"}``) writes the single Parameters row. ``table_seeds``
    is ``[(elementId, [columnId, ...], rows)]``: each table gets one ``upsert_rows`` call
    with all of its rows (at most ``MAX_SEED_ROWS``); a table with no rows is skipped.
    ``column_widths`` (``{elementId: {columnId: px}}``, see ``input_column_widths``) sizes input-table
    columns after their style.
    upsert_rows neither queries first nor enforces keys, so the program is run once, on a
    fresh workbook."""
    tables = _input_tables(spec)
    if not tables:
        raise WorkflowError("workbook has no input tables")
    pages = {el: pg for pg, el in tables}
    table_seeds = [(el, list(cols), list(rows)) for el, cols, rows in table_seeds if rows]
    for element_id, _cols, rows in table_seeds:
        if element_id not in pages:
            raise WorkflowError(f"seed requested for {element_id}, which is not an input table of the workbook")
        if len(rows) > MAX_SEED_ROWS:
            raise WorkflowError(f"{element_id} seed has {len(rows):,} rows; one program seeds at most "
                                f"{MAX_SEED_ROWS:,} rows per table")
    lines = ["# Enable manual commit on every input table (required for backend actions / agent tools);",
             "# first re-apply the typed columns' display formats / hidden flags, the table style and column widths (a workbook POST may drop them).",
             "# Run once through Sigma MCP Code Mode after the workbook is created."]
    if seed or table_seeds:
        seeded = ([PARAMETERS] if seed else []) + [el for el, _c, _r in table_seeds]
        lines.append("# The seeds at the end write rows into " + ", ".join(seeded) + ": run once on a fresh "
                     "workbook (the writes do not query first and keys are not enforced).")
        lines.append("# Every call above the seeds only sets formats, styles and edit modes, so it is safe to re-run.")
    lines.append("")
    display = input_column_display(spec)
    styles = input_table_style_lines(spec)
    for page_id, element_id in tables:
        lines += [f"t = sigma.workbook.get_page('{page_id}').get_element('{element_id}')",
                  "if not isinstance(t, InputTable):",
                  f"    raise RuntimeError('{element_id} is not an InputTable')"]
        for column_id, fmt, hidden in display.get(element_id, []):
            lines += [f"c = t.get_column('{column_id}')",
                      "if not isinstance(c, LiteralEditableColumn):",
                      f"    raise RuntimeError('{element_id}.{column_id} is not an editable column')"]
            if fmt:
                lines.append(f"c.set_format_pattern({fmt!r})")
            if hidden:
                lines.append("c.set_visible(False)")
        lines += styles.get(element_id, [])
        lines += [f"t.set_column_width(t.get_column({cid!r}), {px})"
                  for cid, px in ((column_widths or {}).get(element_id) or {}).items()]
        lines += ["t.set_edit_mode('draft')", "t.set_edit_mode('published-all-users')", ""]
    # Author-only settings (series colours) precede every seed: everything above the seeds is safe
    # to re-run; the seeds are the program's only non-repeatable writes, so they go last.
    lines += _series_color_lines(spec, series_colors)
    if seed:
        page = pages.get(PARAMETERS)
        if page is None:
            raise WorkflowError("seed requested but the workbook has no parameters input table")
        lines += ["# Seed the single Parameters row (fresh workbook only).",
                  f"t = sigma.workbook.get_page('{page}').get_element('{PARAMETERS}')",
                  "if not isinstance(t, InputTable):",
                  f"    raise RuntimeError('{PARAMETERS} is not an InputTable')",
                  *([f"c_scn = t.get_column('{COLUMN_IDS[(PARAMETERS, 'Selected Scenario')]}')",
                     "if not isinstance(c_scn, LiteralEditableColumn):",
                     "    raise RuntimeError('Selected Scenario is not an editable column')"]
                    if seed.get("scenario") else []),
                  f"c_start = t.get_column('{COLUMN_IDS[(PARAMETERS, 'Forecast Start Month')]}')",
                  "if not isinstance(c_start, LiteralEditableColumn):",
                  "    raise RuntimeError('Forecast Start Month is not an editable column')",
                  "t.upsert_rows([{" + (f"c_scn: {seed['scenario']!r}, " if seed.get("scenario") else "")
                  + f"c_start: {seed['start']!r}}}])", ""]
    for element_id, columns, rows in table_seeds:
        lines += _table_seed_lines(pages[element_id], element_id, columns, rows)
    program = "\n".join(lines)
    import code_mode_check
    errors, _calls = code_mode_check.check_program(program)
    if errors:
        raise WorkflowError("setup program does not match the recorded Sigma SDK signatures "
                            f"({code_mode_check.SDK.SOURCE}): " + "; ".join(errors))
    return program
