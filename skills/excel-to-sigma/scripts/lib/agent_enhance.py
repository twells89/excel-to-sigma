"""Phase 6 -- Enhance: recommend a Sigma workbook agent from the macro inventory.

``recommend_agent`` reads the macro classes and the workflows the macros became
(``workflows.json``) and decides whether a workbook agent is worth offering. Write
workflows become approval-gated action tools; the scenario switch is a no-approval
tool; recalculation/formatting macros never produce a tool. ``agent_document_fragment``
emits the ``document.agents`` entry and the ``chat`` element in the shapes proven
in the Task 2 spike. Nothing here contacts Sigma or sends email.
"""
from __future__ import annotations

import copy

import demo_workflows as dw

ACTION_CLASSES = ("ACTION_WRITEBACK", "WHATIF_SCENARIO", "EXTERNAL_IO")
SCENARIO_ID = "set-scenario"
SCENARIO_INPUT = "scenario"
DRIVER_ID = "set-driver"          # growth driver edit; not a macro workflow, offered with the act profile
DRIVER_TABLE = "rateInput"        # build-forecast-model.py "Growth Rates (editable)" input table
DRIVER_KEY_COL, DRIVER_VALUE_COL = "LINE_ITEM", "RATE"
# Plain nouns for the on-camera sentence, keyed by actionId.
SPOKEN = {"roll-forward": "roll forward a month", "snapshot-version": "snapshot a version",
          "submit-approval": "submit for approval", SCENARIO_ID: "switch scenario",
          DRIVER_ID: "change a growth driver"}


def _macro_list(macros):
    if isinstance(macros, dict):
        macros = macros.get("macros", [])
    return [m for m in (macros or []) if isinstance(m, dict)]


def _join(items):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " and " + items[-1]


def recommend_agent(macros, workflows, drivers=True):
    macro_list = _macro_list(macros)
    wfs = [w for w in (workflows or []) if isinstance(w, dict)]
    none = {"recommend": False, "profile": None, "reasons": [], "tools": [], "pitch": ""}
    if not macro_list:
        none["reasons"] = ["No macros were found, so there are no manual workflows for an agent to run."]
        return none
    has_action_class = any(m.get("class") in ACTION_CLASSES for m in macro_list)
    tools = []
    if has_action_class:  # never invent a tool for a macro the inventory did not classify as an action
        for w in wfs:
            if w.get("class") not in ACTION_CLASSES:
                continue
            # every tool changes the workbook (the scenario switch edits Parameters), so every one asks first
            tools.append({"actionId": w["actionId"], "label": w["label"], "requiresApproval": True})
        if drivers and tools:
            tools.insert(1 if tools[0]["actionId"] == SCENARIO_ID else 0,
                         {"actionId": DRIVER_ID, "label": "Set growth driver", "requiresApproval": True})
    total = len(macro_list)
    if not tools:
        classes = sorted({m.get("class", "?") for m in macro_list})
        return {"recommend": False, "profile": None, "tools": [],
                "reasons": [f"The {total} macros are {_join(classes)}; none is a workflow people trigger by hand. "
                            "Sigma recalculates and styles the workbook on its own, so an agent has nothing to run."],
                "pitch": ""}
    writes = [t for t in tools if t["requiresApproval"]]
    names = [SPOKEN.get(t["actionId"], t["label"].lower()) for t in tools]
    flows = [t for t in tools if t["actionId"] != DRIVER_ID]  # driver edit is offered, not a macro
    reasons = [f"{len(flows)} of {total} macros were workflows people trigger by hand: "
               + _join(t["label"] for t in flows) + "."]
    if writes:
        reasons.append(f"{len(writes)} of them change the workbook, so each asks for approval before it runs.")
    pitch = (f"{len(flows)} of your {total} macros were workflows people trigger by hand. "
             f"A Sigma Agent can run them on request: {_join(names)}. "
             + ("Each one changes the workbook, so each asks for your approval first. " if len(writes) == len(tools)
                else f"The {len(writes)} that write ask for your approval first. " if writes else "")
             + "It reads the forecast and the log tables to check what changed, and it never sends email.")
    return {"recommend": True, "profile": "act" if writes else "analyze", "reasons": reasons,
            "tools": tools, "pitch": pitch}


def _tool(entry, effects, label):
    action_id = entry["actionId"]
    steps = [{"kind": "effect", **copy.deepcopy(e)} for e in effects]
    tool = {"toolId": action_id, "kind": "action", "name": label,
            "description": TOOL_DESCRIPTIONS.get(action_id, label + "."), "steps": steps}
    if entry["requiresApproval"]:
        tool["requiresApproval"] = True
    return tool


TOOL_DESCRIPTIONS = {
    "roll-forward": "Advances Forecast Start Month in Parameters by one month. Takes no inputs.",
    "snapshot-version": "Appends one Versions row freezing the current scenario, start month and headline figures. Takes no inputs.",
    "submit-approval": "Appends a Submitted row to Approvals for the current forecast. Takes no inputs. Does not send email.",
    DRIVER_ID: "Changes the monthly growth Rate of one row in the Growth Rates table. Inputs line_item (text): "
               "the exact Line Item, for example North America Revenue; rate (number): the new monthly rate as a "
               "decimal, for example 0.008 for 0.8%. Read the current rate first and compute the new value.",
    SCENARIO_ID: "Sets the Selected Scenario in Parameters. Input scenario (text): the scenario name, for example Base, Upside or Downside.",
}


def _scenario_effects():
    return [{"effect": "update-rows", "tableElementId": dw.PARAMETERS,
             "whichRows": {"type": "formula", "formula": "true"},
             "values": {dw.COLUMN_IDS[(dw.PARAMETERS, "Selected Scenario")]:
                        {"type": "agent-input", "inputName": SCENARIO_INPUT}}}]


def _driver_effects():
    return [{"effect": "update-rows", "tableElementId": DRIVER_TABLE,
             # single-row + primaryKeys on the input table's own column ids is the documented agent-tool
             # row selector (column-match takes only a literal string value, so it cannot bind an agent input)
             "whichRows": {"type": "single-row",
                           "primaryKeys": {DRIVER_KEY_COL: {"type": "agent-input", "inputName": "line_item"}}},
             "values": {DRIVER_VALUE_COL: {"type": "agent-input", "inputName": "rate"}}}]


def _instructions(rec, name):
    ids = {t["actionId"] for t in rec["tools"]}
    lines = [f"You operate the {name} workbook that replaced a macro-driven Excel forecast. "
             'Read @dataSource("parameters"), @dataSource("versions") and @dataSource("approvals") '
             "to see the current scenario, forecast start month, saved versions and approval requests."]
    if "roll-forward" in ids:
        lines.append('To roll the forecast forward one month call @tool("roll-forward").')
    if "snapshot-version" in ids:
        lines.append('To save the current numbers as a version call @tool("snapshot-version").')
    if "submit-approval" in ids:
        lines.append('To submit the forecast for approval call @tool("submit-approval"). Never send email.')
    if DRIVER_ID in ids:
        lines.append('To change a growth driver call @tool("set-driver") with the Line Item and the new monthly '
                     'rate; read @dataSource("rateInput") first and compute the new rate from the current one.')
    if SCENARIO_ID in ids:
        lines.append(f'To switch scenario call @tool("{SCENARIO_ID}") with the scenario name.')
    lines.append("Writes ask the user for approval first. After any write, approved or declined, re-query "
                 "the affected table and report the row count and values as the data shows them; never "
                 "claim a change from the tool result alone. Refer to columns by their display names only.")
    return " ".join(lines)


def agent_document_fragment(rec, name, grounding_element_ids, effects_by_action=None, request_label="Forecast",
                            figures=None, agent_id="fpa-agent", chat_id="chat"):
    """``{"agents": [agent], "chatElement": {...}}`` for a recommendation.

    ``effects_by_action`` (actionId -> effect list) lets the caller reuse the exact effects of the
    workbook's own buttons so button and agent tool can never drift; otherwise the effects come
    from the workflow builder.
    """
    if not rec.get("recommend") or not rec.get("tools"):
        raise ValueError("no agent recommended; nothing to emit")
    effects_by_action = effects_by_action or {}
    tools = []
    for entry in rec["tools"]:
        action_id = entry["actionId"]
        if action_id == DRIVER_ID:
            if DRIVER_TABLE not in grounding_element_ids:
                continue  # workbook has no growth-rate table; never emit a tool that targets nothing
            effects = _driver_effects()
        elif action_id == SCENARIO_ID:
            effects = _scenario_effects()
        elif action_id in effects_by_action:
            effects = effects_by_action[action_id]
        else:
            effects = dw.action_effects(action_id, request_label, figures)
        tools.append(_tool(entry, effects, entry["label"]))
    if DRIVER_TABLE not in grounding_element_ids:
        rec = {**rec, "tools": [t for t in rec["tools"] if t["actionId"] != DRIVER_ID]}
    agent = {"id": agent_id, "name": name, "instructions": _instructions(rec, name),
             "dataSources": [{"kind": "table", "elementId": e} for e in grounding_element_ids],
             "tools": tools}
    return {"agents": [agent], "chatElement": {"id": chat_id, "kind": "chat", "agentId": agent_id}}


def button_effects(spec):
    """actionId -> effects of the workbook's own workflow buttons (single source of truth)."""
    import code_rep
    out = {}
    by_id = {e.get("id"): e for e in code_rep.workbook_elements(spec)}
    for action_id, element_id in dw.BUTTON_IDS.items():
        el = by_id.get(element_id)
        if el and el.get("actions"):
            out[action_id] = el["actions"][0]["effects"]
    return out


def merge_fragment(spec, fragment, chat_anchor="forecast", columns=(1, 25), rows=24):
    """Return a copy of a (wired or nested) workbook spec with the agent and chat added.

    The chat is placed on the page that owns ``chat_anchor`` (falling back to the first page),
    in a new band below that page's last element.
    """
    import re
    import code_rep
    out = copy.deepcopy(spec)
    doc = out["contents"] if isinstance(out.get("contents"), dict) else out.get("document", out)
    doc["agents"] = fragment["agents"]
    chat = fragment["chatElement"]
    elements = doc.setdefault("elements", [])
    if any(e.get("id") == chat["id"] for e in elements):
        raise ValueError(f"workbook already has an element {chat['id']}")
    elements.append(chat)
    owners = code_rep.workbook_page_by_element(doc)
    page_id = (owners.get(chat_anchor) or doc["pages"][0])["id"]
    layout = doc["layout"]
    block = re.search(r'(<Page\b[^>]*\bid="%s"[^>]*>)(.*?)(</Page>)' % re.escape(page_id), layout, re.S)
    if not block:
        raise ValueError(f"page {page_id} not found in layout")
    ends = [int(m.group(1)) for m in re.finditer(r'gridRow="\d+ / (\d+)"', block.group(2))]
    start = max(ends, default=1)
    line = (f'  <Element elementId="{chat["id"]}" gridColumn="{columns[0]} / {columns[1]}" '
            f'gridRow="{start} / {start + rows}"/>\n')
    doc["layout"] = layout[:block.end(2)] + line + layout[block.end(2):]
    return out
