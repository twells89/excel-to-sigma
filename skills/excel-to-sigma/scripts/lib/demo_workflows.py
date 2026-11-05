"""Demo macros -> native Sigma workflows (input tables written by button actions).

``build_workflow_manifest`` classifies a macro module (``.bas`` source or a real
``.xlsm``; VBA is only read, never executed) and routes each macro to a native
workflow or records a disposition. ``workflow_elements`` turns the manifest into
workbook elements: Parameters / Versions / Approvals input tables, three buttons
and a latest-approval-status KPI. ``emit_manual_commit_program`` writes the
post-create Code Mode program that enables manual commit on every input table
(backend actions, e.g. a Sigma Agent tool, only write to such tables; a REST
spec POST never sets the flag).

Nothing here sends email: the Outlook step of the Submit macro is replaced by an
Approvals row and recorded under ``dispositions``.
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
TABLES = {
    PARAMETERS: ("Parameters", [("pa_scn", "Selected Scenario", "text"),
                                ("pa_start", "Forecast Start Month", "datetime")]),
    VERSIONS: ("Versions", [("ve_ver", "Version", "text"), ("ve_scn", "Scenario", "text"),
                            ("ve_start", "Forecast Start Month", "datetime"),
                            ("ve_rev", "FY Revenue", "number"), ("ve_gm", "FY Gross Margin", "number"),
                            ("ve_ebitda", "FY EBITDA", "number"),
                            ("ve_by", "Saved By", "text"), ("ve_at", "Saved At", "datetime")]),
    APPROVALS: ("Approvals", [("ap_req", "Request", "text"), ("ap_by", "Submitted By", "text"),
                              ("ap_to", "Approver", "text"), ("ap_status", "Status", "text"),
                              ("ap_at", "Submitted At", "datetime")]),
}
COLUMN_IDS = {(el, name): cid for el, (_n, cols) in TABLES.items() for cid, name, _t in cols}

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
DISPOSITIONS = {
    "EXTERNAL_DATA_REFRESH": ("replaced-by-live-warehouse",
        "The ODBC reload of Actuals is replaced by the live Snowflake landing table; no refresh button is needed."),
    "RECALC": ("replaced-by-live-calculation",
        "Sigma recalculates on read; there is nothing to port."),
    "FORMAT_COSMETIC": ("styling",
        "Cell formatting is carried by workbook styling and number formats, not an action."),
}
NOTES = {
    "roll-forward": "Advances Parameters.Forecast Start Month by one month (update-rows, DateAdd). "
                    "Replaces the VBA that copied the first forecast month into the Actuals table.",
    "snapshot-version": "Appends ONE Versions row freezing the headline numbers at click time (version label, "
                        "scenario, forecast start month, FY revenue / gross margin / EBITDA taken from the same "
                        "formulas the KPI tiles use, user, time). The VBA pasted the whole forecast grid as "
                        "values; the full line x month grid is not reproduced.",
    "submit-approval": "Appends an Approvals row (status Submitted, user, time). The VBA also protected the "
                       "sheet and sent an Outlook email; neither is reproduced.",
    "set-scenario": "Parameters is a one-row input table (a control filters but cannot feed a formula "
                    "parameter); the scenario is picked in the Parameters cell, no button action writes it.",
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


def build_workflow_manifest(path):
    """Route every macro to a native workflow or a recorded disposition (never silent)."""
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
                    "note": "The Outlook email to the CFO is replaced by the Approvals row and the latest-status "
                            "header; no email is sent by the workflow, the button or the agent."})
        elif cls in DISPOSITIONS:
            kind, note = DISPOSITIONS[cls]
            dispositions.append({"macro": name, "class": cls, "disposition": kind, "note": note})
        else:
            dispositions.append({"macro": name, "class": cls, "disposition": "unmapped-needs-review",
                "note": "No native workflow mapping; a human must review before migration is complete."})
    return {"version": 1, "workflows": workflows, "dispositions": dispositions}


STATUS_FORMULA = ('Coalesce(Lookup([Approvals/Status], Max([Approvals/Submitted At]), '
                  '[Approvals/Submitted At]), "Not submitted")')


def _const(text):
    return {"type": "constant", "value": {"type": "text", "value": text}}


def _formula(text):
    return {"type": "formula", "formula": text}


def _values(element, formulas):
    return {COLUMN_IDS[(element, name)]: value for name, value in formulas.items()}


def action_effects(action_id, request_label, figures=None):
    """The effect list for one workflow (the working shapes from the Task 2 spike)."""
    if action_id == "roll-forward":
        return [{"effect": "update-rows", "tableElementId": PARAMETERS,
                 "whichRows": {"type": "formula", "formula": "true"},
                 "values": _values(PARAMETERS, {"Forecast Start Month": _formula(
                     'DateAdd("month", 1, Max([Parameters/Forecast Start Month]))')})}]
    if action_id == "snapshot-version":
        if not figures:
            raise WorkflowError("snapshot needs the KPI figure formulas (figures=)")
        values = {
            "Version": _formula('Concat("v", Text(Count([Versions/Saved At]) + 1))'),
            "Scenario": _formula("Max([Parameters/Selected Scenario])"),
            "Forecast Start Month": _formula("Max([Parameters/Forecast Start Month])"),
            "Saved By": _formula("CurrentUserEmail()"), "Saved At": _formula("Now()")}
        names = {"revenue": "FY Revenue", "grossMargin": "FY Gross Margin", "ebitda": "FY EBITDA"}
        for key, formula in figures.items():
            values[names[key]] = _formula(formula)
        return [{"effect": "insert-rows", "tableElementId": VERSIONS, "values": _values(VERSIONS, values)}]
    if action_id == "submit-approval":
        return [{"effect": "insert-rows", "tableElementId": APPROVALS, "values": _values(APPROVALS, {
            "Request": _const(request_label), "Submitted By": _formula("CurrentUserEmail()"),
            "Approver": _const("CFO"), "Status": _const("Submitted"), "Submitted At": _formula("Now()")})}]
    raise WorkflowError(f"no button action for {action_id}")


def input_table(element, connection_id, mode, figures=None):
    name, cols = TABLES[element]
    if element == VERSIONS and figures is not None:
        keep = {"FY Revenue": "revenue", "FY Gross Margin": "grossMargin", "FY EBITDA": "ebitda"}
        cols = [c for c in cols if c[1] not in keep or keep[c[1]] in figures]
    return {"id": element, "kind": "input-table", "name": name, "inputMode": mode,
            "source": {"kind": "empty", "connectionId": connection_id},
            "columns": [{"id": "ID", "name": "Row ID", "hidden": True},
                        *[{"id": cid, "type": typ, "name": nm} for cid, nm, typ in cols]]}


def workflow_elements(manifest, connection_id, request_label, figures=None):
    """Workbook elements for a manifest: ``{tables, buttons, status}``.

    ``figures`` maps revenue / grossMargin / ebitda to the workbook's KPI formulas;
    Snapshot freezes exactly those into its Versions row.
    """
    if figures is None and any(w["actionId"] == "snapshot-version" for w in manifest.get("workflows", [])):
        figures = {"revenue": "Sum([PnL Combined/Amount])"}  # structural default (checks); builder passes real KPIs
    if not isinstance(manifest, dict) or manifest.get("version") != 1:
        raise WorkflowError("workflows manifest version 1 required")
    items = {w["actionId"]: w for w in manifest.get("workflows", [])}
    for w in items.values():
        if w["effect"] not in EFFECTS:
            raise WorkflowError(f"unsupported workflow effect {w['effect']}")
    tables = [input_table(PARAMETERS, connection_id, "edit"),
              input_table(VERSIONS, connection_id, "view", figures),
              input_table(APPROVALS, connection_id, "view")]
    buttons = []
    for action_id, element_id in BUTTON_IDS.items():
        if action_id not in items:
            continue
        w = items[action_id]
        buttons.append({"id": element_id, "kind": "button", "text": w["label"],
            "appearance": "outline" if action_id == "submit-approval" else "filled",
            "actions": [{"id": "a-" + action_id, "trigger": "on-click",
                         "effects": action_effects(action_id, request_label, figures)}]})
    # Latest status = Status of the Approvals row with the max Submitted At. A one-row table
    # over Parameters holds the cross-table Lookup (the shape proven for the forecast lookup).
    status = {"id": STATUS_KPI, "kind": "table", "name": "Latest approval status",
        "source": {"kind": "table", "elementId": PARAMETERS},
        "columns": [{"id": "st_v", "name": "Latest approval status", "formula": STATUS_FORMULA}]}
    return {"tables": tables, "buttons": buttons, "status": status}


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


def emit_manual_commit_program(spec, seed=None):
    """Post-create Code Mode program text: re-set each input table's edit mode so the
    host publishes it with manual commit enabled."""
    tables = _input_tables(spec)
    if not tables:
        raise WorkflowError("workbook has no input tables")
    lines = ["# Enable manual commit on every input table (required for backend actions / agent tools).",
             "# Run once through Sigma MCP Code Mode after the workbook is created.", ""]
    if seed:
        lines[1] += "\n# The Parameters seed at the end writes one row: run once on a fresh workbook (it cannot query first)."
    for page_id, element_id in tables:
        lines += [f"t = sigma.workbook.get_page('{page_id}').get_element('{element_id}')",
                  "if not isinstance(t, InputTable):",
                  f"    raise RuntimeError('{element_id} is not an InputTable')",
                  "t.set_edit_mode('draft')", "t.set_edit_mode('published-all-users')", ""]
    if seed:
        page = next((pg for pg, el in tables if el == PARAMETERS), None)
        if page is None:
            raise WorkflowError("seed requested but the workbook has no parameters input table")
        lines += ["# Seed the single Parameters row (fresh workbook only).",
                  f"t = sigma.workbook.get_page('{page}').get_element('{PARAMETERS}')",
                  "if not isinstance(t, InputTable):",
                  f"    raise RuntimeError('{PARAMETERS} is not an InputTable')",
                  f"c_scn = t.get_column('{COLUMN_IDS[(PARAMETERS, 'Selected Scenario')]}')",
                  "if not isinstance(c_scn, LiteralEditableColumn):",
                  "    raise RuntimeError('Selected Scenario is not an editable column')",
                  f"c_start = t.get_column('{COLUMN_IDS[(PARAMETERS, 'Forecast Start Month')]}')",
                  "if not isinstance(c_start, LiteralEditableColumn):",
                  "    raise RuntimeError('Forecast Start Month is not an editable column')",
                  f"t.upsert_rows([{{c_scn: {seed['scenario']!r}, c_start: {seed['start']!r}}}])", ""]
    return "\n".join(lines)
