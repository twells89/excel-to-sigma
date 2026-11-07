"""Phase 6 -- Enhance: recommend a Sigma workbook agent from the macro inventory.

``recommend_agent`` reads the macro classes and the workflows the macros became
(``workflows.json``) and decides whether a workbook agent is worth offering. Every
workflow tool, the scenario switch included, changes the workbook and is
approval-gated; recalculation/formatting macros never produce a tool.
``agent_document_fragment`` emits the ``document.agents`` entry and the ``chat``
element in live-proven shapes. Nothing here contacts Sigma or sends email.
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
DRIVER_KEY_NAME = "Line item"     # the planning app's display name for LINE_ITEM (read from the workbook when given)
# The set-driver tool's row picker: a hidden text control the tool fills with the agent's line item.
DRIVER_CONTROL, DRIVER_CONTROL_ELEMENT = "DriverLineItem", "ctlDriverLineItem"
# The agent's fiscal-year source: lib/forecast_app.py's monthly Actual / Forecast / Plan facts.
AGENT_FACTS, PLAN_ELEMENT = "agentFacts", "planLines"
AGENT_PANEL = "agentPanel"        # lib/forecast_app.py reserves this container for the chat
SCENARIO_CONTROL = "Scenario"     # lib/forecast_app.py scenario switch controlId
# Plain nouns for the on-camera sentence, keyed by actionId.
SPOKEN = {"roll-forward": "roll forward a month", "snapshot-version": "snapshot a version",
          "submit-approval": "submit for approval", SCENARIO_ID: "switch scenario",
          DRIVER_ID: "change a growth driver"}


def _macro_list(macros):
    if isinstance(macros, dict):
        macros = macros.get("macros", [])
    return [m for m in (macros or []) if isinstance(m, dict)]


EXTRA_NOUN = {DRIVER_ID: "a growth-driver tool"}   # agent tools that are not one of the macros
_ORDINALS = ("first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth")


def _ordinal(n):
    return _ORDINALS[n - 1] if 1 <= n <= len(_ORDINALS) else f"{n}th"


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
    spoken = lambda t: SPOKEN.get(t["actionId"], t["label"].lower())  # noqa: E731
    flows = [t for t in tools if t["actionId"] != DRIVER_ID]  # driver edit is offered, not a macro
    extras = [t for t in tools if t["actionId"] == DRIVER_ID]
    # One count per noun: macros -> workflows, then actions (workflows + extra tools); never "4 ... 5 of them".
    n, nf = len(tools), len(flows)
    reasons = [f"{nf} of {total} macros were workflows people trigger by hand: "
               + _join(t["label"] for t in flows) + "."]
    gets = (f"The agent gets {n} actions: those {nf} workflows plus "
            + _join(EXTRA_NOUN.get(t["actionId"], f"a {t['label'].lower()} tool") for t in extras) + "."
            if extras else f"The agent gets {n} actions, one per workflow.")
    if len(writes) == n:
        gets += f" All {n} change the workbook, so each asks for approval before it runs."
    elif writes:
        gets += f" {len(writes)} of the {n} change the workbook, so each of those asks for approval before it runs."
    reasons.append(gets)
    if not extras:
        more = ""
    elif len(extras) == 1:
        more = f", plus a {_ordinal(n)} action: {spoken(extras[0])}"
    else:
        more = f", plus {len(extras)} more actions: " + _join(spoken(t) for t in extras)
    pitch = (f"{nf} of your {total} macros were workflows people trigger by hand: "
             f"{_join(spoken(t) for t in flows)}. A Sigma Agent can run all {nf} on request{more}. "
             + (f"Each of the {n} actions changes the workbook, so each asks for your approval first. "
                if len(writes) == n
                else f"The {len(writes)} of the {n} actions that change the workbook ask for your approval first. "
                if writes else "")
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
    "roll-forward": "Advances Forecast Start Month in Parameters by one month. Takes no inputs: call it with an empty inputs object, inputs: {}.",
    "snapshot-version": "Appends one Versions row freezing the current scenario, start month and headline figures. Takes no inputs: call it with an empty inputs object, inputs: {}.",
    "submit-approval": "Appends a Submitted row to Approvals for the current forecast. Takes no inputs: call it with an empty inputs object, inputs: {}. Does not send email.",
    DRIVER_ID: "Changes the monthly growth Rate of one row in the Growth Rates table. Inputs line_item (text): "
               "the exact Line Item, for example North America Revenue; rate (number): the new monthly rate as a "
               "decimal, for example 0.008 for 0.8%. Read the current rate first and compute the new value.",
    SCENARIO_ID: "Sets the Selected Scenario in Parameters. Input scenario (text): the scenario name, for example Base, Upside or Downside.",
}


def _scenario_effects(scenario_control=None):
    if scenario_control:
        # the workbook's scenario switch is what the forecast reads: the tool sets that control
        return [{"effect": "set-control-value", "control": scenario_control,
                 "value": {"type": "agent-input", "inputName": SCENARIO_INPUT}}]
    return [{"effect": "update-rows", "tableElementId": dw.PARAMETERS,
             "whichRows": {"type": "formula", "formula": "true"},
             "values": {dw.COLUMN_IDS[(dw.PARAMETERS, "Selected Scenario")]:
                        {"type": "agent-input", "inputName": SCENARIO_INPUT}}}]


def driver_control():
    """The hidden text control set-driver selects its row through (no filter targets: it filters nothing)."""
    return {"id": DRIVER_CONTROL_ELEMENT, "kind": "control", "controlId": DRIVER_CONTROL, "controlType": "text",
            "name": "Driver line item", "mode": "equals"}


def _driver_effects(key_name=DRIVER_KEY_NAME):
    # An input table's only primary key is its system row ID: single-row primaryKeys on LINE_ITEM is refused
    # live ("Primary key: ID not in update row definition"), column-match takes only a literal, and a row
    # formula cannot read an agent input. So the tool puts the line item into a hidden text control, then
    # updates the rows whose line item equals that control (a whichRows formula can read a control).
    return [{"effect": "set-control-value", "control": DRIVER_CONTROL,
             "value": {"type": "agent-input", "inputName": "line_item"}},
            {"effect": "update-rows", "tableElementId": DRIVER_TABLE,
             "whichRows": {"type": "formula", "formula": f"[{key_name}] = [{DRIVER_CONTROL}]"},
             "values": {DRIVER_VALUE_COL: {"type": "agent-input", "inputName": "rate"}}}]


def _instructions(rec, name, scenario_control=None, grounding=(), plan=False):
    ids = {t["actionId"] for t in rec["tools"]}
    lines = [f"You operate the {name} workbook that replaced a macro-driven Excel forecast. "
             'Read @dataSource("parameters"), @dataSource("versions") and @dataSource("approvals") '
             + ("to see the forecast start month, saved versions and approval requests; the Active Scenario "
                'column of @dataSource("forecast") is the scenario the forecast is on.' if scenario_control else
                "to see the current scenario, forecast start month, saved versions and approval requests.")]
    if AGENT_FACTS in grounding:
        # The forecast element holds only the forecast months: totalled, it is not a fiscal year.
        lines.append(f'@dataSource("{AGENT_FACTS}") holds every month\'s amounts by Fiscal year, Period, Section, '
                     "Line item and Region. A fiscal year's figure (for example its revenue) is the sum of its "
                     "Actual and Forecast rows: the months already closed plus the months still forecast."
                     + (" Its plan is the sum of its Plan rows; compare the two for any vs-plan question."
                        if plan else "")
                     + f' Answer fiscal-year{" and plan" if plan else ""} questions from @dataSource("{AGENT_FACTS}"). '
                     '@dataSource("forecast") holds the forecast months only; never report its total as a '
                     "fiscal-year figure.")
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
    if ids & {"roll-forward", "snapshot-version", "submit-approval"}:
        lines.append("Those three tools take no inputs, but every tool call must still include the inputs "
                     "argument: pass inputs: {} (an empty object).")
    lines.append("Answer in plain finance language with figures formatted like $145.4M and 65.8%.")
    lines.append("Writes ask the user for approval first. After any write, approved or declined, re-query "
                 "the affected table and report the row count and values as the data shows them; never "
                 "claim a change from the tool result alone. Refer to columns by their display names only.")
    return " ".join(lines)


GREETING = ("Hi, I'm the forecast agent. I can explain the numbers or run the forecast workflows for you: "
            "{actions}. Anything that changes the workbook asks you first. Try asking: {example}")


def agent_document_fragment(rec, name, grounding_element_ids, effects_by_action=None, request_label="Forecast",
                            figures=None, agent_id="fpa-agent", chat_id="chat", scenario_control=None,
                            driver_key_name=DRIVER_KEY_NAME, plan=False):
    """``{"agents": [agent], "chatElement": {...}, "controlElements": [...]}`` for a recommendation.

    ``controlElements`` are the hidden controls the tools set (the set-driver row picker); ``merge_fragment``
    places them on the workbook's hidden page. ``plan`` says the grounded facts carry Plan rows.

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
            effects = _driver_effects(driver_key_name)
        elif action_id == SCENARIO_ID:
            effects = _scenario_effects(scenario_control)
        elif action_id in effects_by_action:
            effects = effects_by_action[action_id]
        else:
            effects = dw.action_effects(action_id, request_label, figures)
        tools.append(_tool(entry, effects, entry["label"]))
    if DRIVER_TABLE not in grounding_element_ids:
        rec = {**rec, "tools": [t for t in rec["tools"] if t["actionId"] != DRIVER_ID]}
    actions = _join(SPOKEN.get(t["actionId"], t["label"].lower()) for t in rec["tools"])
    # Suggest the plan question only when the agent can read fiscal-year plan figures to answer it.
    plan = bool(plan and AGENT_FACTS in grounding_element_ids)
    example = "how does revenue compare with plan?" if plan else "summarise the forecast."
    agent = {"id": agent_id, "name": name,
             "instructions": _instructions(rec, name, scenario_control, grounding_element_ids, plan),
             "greeting": {"mode": "static", "message": GREETING.format(actions=actions, example=example)},
             "dataSources": [{"kind": "table", "elementId": e} for e in grounding_element_ids],
             "tools": tools}
    controls = [driver_control()] if any(t["toolId"] == DRIVER_ID for t in tools) else []
    return {"agents": [agent], "chatElement": {"id": chat_id, "kind": "chat", "agentId": agent_id},
            "controlElements": controls}


def check_agent_steps(doc):
    """Refuse agent tool steps the live API rejects, before anything is posted: a ``single-row`` selector
    keyed on anything but the target input table's system row ID ("Primary key: ID not in update row
    definition"), and a ``set-control-value`` on a control the workbook does not have."""
    by_id = {e.get("id"): e for e in doc.get("elements", [])}
    control_ids = {e.get("controlId") for e in doc.get("elements", []) if e.get("kind") == "control"}
    for agent in doc.get("agents") or []:
        for tool in agent.get("tools") or []:
            for step in tool.get("steps") or []:
                rows = step.get("whichRows") or {}
                if rows.get("type") == "single-row":
                    table = by_id.get(step.get("tableElementId")) or {}
                    if table.get("kind") != "input-table":
                        raise ValueError(f"tool {tool.get('toolId')}: single-row selector on "
                                         f"{step.get('tableElementId')}, which is not an input table")
                    keys = {c["id"] for c in table.get("columns", []) if c.get("id") == "ID" or "type" not in c}
                    extra = sorted(set(rows.get("primaryKeys") or {}) - keys)
                    if extra or not rows.get("primaryKeys"):
                        raise ValueError(f"tool {tool.get('toolId')}: single-row primaryKeys {extra or '{}'} on "
                                         f"{table['id']}; its only primary key is the system row ID")
                if step.get("effect") == "set-control-value" and step.get("control") not in control_ids:
                    raise ValueError(f"tool {tool.get('toolId')}: sets control {step.get('control')}, which the "
                                     "workbook does not have")


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
    for el in [chat, *fragment.get("controlElements", [])]:
        if any(e.get("id") == el["id"] for e in elements):
            raise ValueError(f"workbook already has an element {el['id']}")
    elements.append(chat)
    for control in fragment.get("controlElements", []):
        # tool plumbing: on the hidden page, never on a page people see
        owners = code_rep.workbook_page_by_element(doc)
        hidden = next((p for p in doc.get("pages", []) if p.get("visibility") == "hidden"), None)
        page_id = (hidden or owners.get(DRIVER_TABLE) or doc["pages"][0])["id"]
        elements.append(control)
        doc["layout"] = _append_to_page(doc["layout"], page_id, control["id"], 2)
    check_agent_steps(doc)
    layout = doc["layout"]
    # A planning-app workbook reserves an agent panel: the chat fills the panel below its header.
    panel = re.search(r'(<Container\b[^>]*\belementId="%s"[^>]*>)(.*?)(</Container>)' % re.escape(AGENT_PANEL),
                      layout, re.S)
    if panel:
        ends = [int(m.group(1)) for m in re.finditer(r'gridRow="\d+ / (\d+)"', panel.group(2))]
        start = max(ends, default=1)
        rows_end = int(re.search(r'gridRow="(\d+) / (\d+)"', panel.group(1)).group(2)) - \
            int(re.search(r'gridRow="(\d+) / (\d+)"', panel.group(1)).group(1)) + 1
        line = (f'      <Element elementId="{chat["id"]}" gridColumn="1 / 25" '
                f'gridRow="{start} / {max(rows_end, start + 6)}"/>\n')
        doc["layout"] = layout[:panel.end(2)] + line + layout[panel.end(2):]
        return out
    owners = code_rep.workbook_page_by_element(doc)
    page_id = (owners.get(chat_anchor) or doc["pages"][0])["id"]
    doc["layout"] = _append_to_page(layout, page_id, chat["id"], rows, columns)
    return out


def _append_to_page(layout, page_id, element_id, rows, columns=(1, 25)):
    """``layout`` with ``element_id`` in a new band below the last element of page ``page_id``."""
    import re
    block = re.search(r'(<Page\b[^>]*\bid="%s"[^>]*>)(.*?)(</Page>)' % re.escape(page_id), layout, re.S)
    if not block:
        raise ValueError(f"page {page_id} not found in layout")
    ends = [int(m.group(1)) for m in re.finditer(r'gridRow="\d+ / (\d+)"', block.group(2))]
    start = max(ends, default=1)
    line = (f'  <Element elementId="{element_id}" gridColumn="{columns[0]} / {columns[1]}" '
            f'gridRow="{start} / {start + rows}"/>\n')
    return layout[:block.end(2)] + line + layout[block.end(2):]


GROUNDING = ("forecast", "parameters", "versions", "approvals", "rateInput", AGENT_FACTS)


def recommend_for_workbook(macros, workflows, wb=None, name="FP&A forecast agent", fragment=True):
    """``(recommendation, fragment or None)`` for a workbook spec (or the default grounding).

    Grounds the agent in the workbook's real elements, reuses its button effects and, when the
    workbook has the planning app's scenario switch, makes set-scenario set that control."""
    import code_rep
    wf = workflows.get("workflows", []) if isinstance(workflows, dict) else workflows
    elements = code_rep.workbook_elements(wb) if wb else []
    have = {e.get("id") for e in elements} if wb else set(GROUNDING)
    rec = recommend_agent(macros, wf, drivers=DRIVER_TABLE in have)
    if not rec["recommend"] or not fragment:
        return rec, None
    ids = [i for i in GROUNDING if i in have]
    effects = button_effects(wb) if wb else None
    scenario = next((e.get("controlId") for e in elements
                     if e.get("kind") == "control" and e.get("controlId") == SCENARIO_CONTROL), None)
    rate = next((e for e in elements if e.get("id") == DRIVER_TABLE), {})
    key_name = next((c.get("name") for c in rate.get("columns", []) if c.get("id") == DRIVER_KEY_COL),
                    None) or DRIVER_KEY_NAME
    return rec, agent_document_fragment(rec, name, ids, effects_by_action=effects, scenario_control=scenario,
                                        driver_key_name=key_name, plan=PLAN_ELEMENT in have)

