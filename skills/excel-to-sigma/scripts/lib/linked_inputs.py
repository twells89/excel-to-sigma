"""Sigma MCP Code Mode programs that create, seed and edit a LINKED input table.

Validated live 2026-10-22 (AWS US West production MCP).  The proven recipe is three
``build_workbook`` calls:

1. base table plus a grouped one-row-per-key table on a hidden data page,
2. ``add_linked_input_table`` with its editable and formula columns (the grouped
   source must come from an EARLIER call),
3. ``update_rows`` keyed by the linked key column.

This module only emits text.  It never calls Sigma; the caller runs each program
through Code Mode, copies the ids the program prints into ``ids.json`` and queries
the readback SQL.  Programs obey the Code Mode dialect (no ``def``/``class``/
``lambda``/``while``, no slices, no f-strings/``.format``, no ``%``/``**``/``//``,
single-clause comprehensions, ``+`` concatenation) and are linted before returning.
"""
from __future__ import annotations

import ast
import datetime as dt
import io
import math
import re
import tokenize

MAX_ROWS = 2000
EDIT_MODES = ("draft", "published-explore", "published-all-users")
COLUMN_TYPES = ("number", "text", "datetime", "boolean")
AGGREGATES = {"max": "max", "min": "min", "sum": "sum", "avg": "avg"}
FORMULA_BINARY = {"add": "add", "subtract": "subtract", "multiply": "multiply",
                  "divide": "divide", "coalesce": "coalesce"}
FORMULA_UNARY = {"number": "number"}
MAX_FORMULA_DEPTH = 8

ALLOWED_BUILTINS = {"print", "len", "any", "enumerate", "zip", "range", "list",
                    "reversed", "sorted", "min", "max", "str", "isinstance",
                    "datetime", "RuntimeError"}
ID_KEYS = ("dataPageId", "pageId", "planBaseId", "linkedTableId")
PLACEHOLDERS = {"dataPageId": "<DATA PAGE ID>", "pageId": "<PAGE ID>",
                "planBaseId": "<PLAN BASE ID>", "linkedTableId": "<LINKED TABLE ID>"}


class LinkedInputError(ValueError):
    pass


# --------------------------------------------------------------------------- spec


def _name(value, what):
    if not isinstance(value, str) or not value.strip() or value != value.strip():
        raise LinkedInputError(what + " must be a non-empty string without surrounding whitespace")
    if any(ord(ch) < 32 for ch in value):
        raise LinkedInputError(what + " must not contain control characters")
    return value


def _frame(value, default):
    if value is None:
        return list(default)
    if (not isinstance(value, (list, tuple)) or len(value) != 4
            or any(isinstance(v, bool) or not isinstance(v, int) or v < 1 for v in value)):
        raise LinkedInputError("frame must be four positive integers [x, y, width, height]")
    return list(value)


def _formula(node, columns, path, depth=0):
    """Validate a formula descriptor; return a normalized copy."""
    if depth > MAX_FORMULA_DEPTH:
        raise LinkedInputError(path + ": formula nests deeper than " + str(MAX_FORMULA_DEPTH))
    if not isinstance(node, dict):
        raise LinkedInputError(path + ": formula node must be an object")
    kinds = [k for k in ("col", "const", "op") if k in node]
    if len(kinds) != 1:
        raise LinkedInputError(path + ": node needs exactly one of col / const / op")
    kind = kinds[0]
    extra = set(node) - {kind, "args"}
    if extra or (kind != "op" and "args" in node):
        raise LinkedInputError(path + ": unexpected keys " + ", ".join(sorted(extra or {"args"})))
    if kind == "col":
        name = node["col"]
        if name not in columns:
            raise LinkedInputError(path + ": formula references unknown or later column " + repr(name))
        return {"col": name}
    if kind == "const":
        value = node["const"]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise LinkedInputError(path + ": const must be a finite number")
        return {"const": value}
    op = node["op"]
    if op in FORMULA_BINARY:
        arity = 2
    elif op in FORMULA_UNARY:
        arity = 1
    else:
        allowed = sorted(list(FORMULA_BINARY) + list(FORMULA_UNARY))
        raise LinkedInputError(path + ": unsupported op " + repr(op) + " (allowed: " + ", ".join(allowed) + ")")
    args = node.get("args")
    if not isinstance(args, list) or len(args) != arity:
        raise LinkedInputError(path + ": op " + op + " takes exactly " + str(arity) + " args")
    return {"op": op, "args": [_formula(a, columns, path + "." + op + "[" + str(i) + "]", depth + 1)
                               for i, a in enumerate(args)]}


def _numeric(node, types):
    """True when the descriptor certainly yields a number."""
    if "const" in node:
        return True
    if "col" in node:
        return types.get(node["col"]) == "number"
    if node["op"] in ("add", "subtract", "multiply", "divide", "number"):
        return True
    return all(_numeric(a, types) for a in node["args"])


def _datetime_value(value, path):
    if isinstance(value, dt.datetime):
        parsed = value
    elif isinstance(value, dt.date):
        parsed = dt.datetime(value.year, value.month, value.day, 12)
    elif isinstance(value, str):
        try:
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                day = dt.date.fromisoformat(value)
                parsed = dt.datetime(day.year, day.month, day.day, 12)   # noon: date-only values render a day early otherwise
            else:
                parsed = dt.datetime.fromisoformat(value)
        except ValueError:
            raise LinkedInputError(path + ": datetime must be YYYY-MM-DD or ISO datetime, got " + repr(value))
    else:
        raise LinkedInputError(path + ": datetime must be an ISO string")
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return parsed


def _cell(value, ctype, path):
    if ctype == "number":
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise LinkedInputError(path + ": expected a finite number, got " + repr(value))
        return value
    if ctype == "text":
        if not isinstance(value, str):
            raise LinkedInputError(path + ": expected text, got " + repr(value))
        return value
    if ctype == "boolean":
        if not isinstance(value, bool):
            raise LinkedInputError(path + ": expected true/false, got " + repr(value))
        return value
    return _datetime_value(value, path)


def normalize_spec(spec):
    """Validate a spec and return a normalized copy; raise :class:`LinkedInputError`."""
    if not isinstance(spec, dict):
        raise LinkedInputError("spec must be an object")
    out = {}
    warnings = []

    wb = spec.get("workbook")
    if not isinstance(wb, dict):
        raise LinkedInputError("workbook must be {id} or {name, description}")
    if wb.get("id"):
        out["workbook"] = {"id": _name(wb["id"], "workbook.id")}
    else:
        out["workbook"] = {"name": _name(wb.get("name"), "workbook.name"),
                           "description": _name(wb.get("description"),
                                                "workbook.description (never omitted for a new workbook)")}
    existing = "id" in out["workbook"]

    data_page = spec.get("dataPage") or {}
    out["dataPage"] = {"name": _name(data_page.get("name", "Data"), "dataPage.name")}
    page = spec.get("page")
    if not isinstance(page, dict) or bool(page.get("id")) == bool(page.get("name")):
        raise LinkedInputError("page must be exactly one of {name} (created in call 1) or {id} (existing page)")
    if page.get("id"):
        if not existing:
            raise LinkedInputError("page.id needs an existing workbook; a new workbook has no pages yet")
        out["page"] = {"id": _name(page["id"], "page.id")}
    else:
        out["page"] = {"name": _name(page["name"], "page.name")}
    if out["page"].get("name") and out["page"]["name"].lower() == out["dataPage"]["name"].lower():
        raise LinkedInputError("page.name must differ from dataPage.name (the data page is hidden)")

    src = spec.get("source")
    if not isinstance(src, dict):
        raise LinkedInputError("source must be an object")
    if src.get("kind") == "dataModel":
        out["source"] = {"kind": "dataModel", "inodeId": _name(src.get("inodeId"), "source.inodeId"),
                         "memberColumn": _name(src.get("memberColumn"),
                                               "source.memberColumn (selects the data model member)")}
    elif src.get("kind") == "element":
        out["source"] = {"kind": "element", "elementId": _name(src.get("elementId"), "source.elementId")}
    else:
        raise LinkedInputError("source.kind must be 'dataModel' or 'element'")
    src_cols = spec.get("sourceColumns")
    if src_cols is not None:
        if not isinstance(src_cols, list) or not all(isinstance(c, str) for c in src_cols):
            raise LinkedInputError("sourceColumns must be a list of column names (from describe)")
        out["sourceColumns"] = list(src_cols)
    src_lower = set(c.lower() for c in src_cols or [])
    if (out["source"]["kind"] == "dataModel" and src_cols is not None
            and out["source"]["memberColumn"] not in src_cols):
        raise LinkedInputError("source.memberColumn " + repr(out["source"]["memberColumn"])
                               + " is not in sourceColumns")

    key = spec.get("key")
    if not isinstance(key, list) or not key:
        raise LinkedInputError("key must be a non-empty list of column names")
    for index, k in enumerate(key):
        _name(k, "key[" + str(index) + "]")
    if len(set(k.lower() for k in key)) != len(key):
        raise LinkedInputError("key columns must be unique")
    if src_cols is not None:
        for k in key:
            if k not in src_cols:
                raise LinkedInputError("key column " + repr(k) + " is not in sourceColumns")
    out["key"] = list(key)
    if len(key) > 1:
        warnings.append("multi-column key: the spike validated a single key; dry-run call 1 and 2 before seeding")

    names = {}          # lower-name -> type for linked-table columns
    types = {}
    for k in key:
        names[k.lower()] = k
        types[k] = None

    inherited = spec.get("inherited") or []
    if not isinstance(inherited, list):
        raise LinkedInputError("inherited must be a list")
    out["inherited"] = []
    calc_names = set()
    for index, item in enumerate(inherited):
        path = "inherited[" + str(index) + "]"
        if not isinstance(item, dict):
            raise LinkedInputError(path + " must be an object")
        calc = _name(item.get("name"), path + ".name")
        of = _name(item.get("of"), path + ".of")
        agg = item.get("aggregate")
        if agg not in AGGREGATES:
            raise LinkedInputError(path + ".aggregate must be one of " + ", ".join(sorted(AGGREGATES)))
        linked = _name(item.get("as", calc), path + ".as")
        if calc.lower() in calc_names:
            raise LinkedInputError(path + ": duplicate calculation name " + repr(calc))
        calc_names.add(calc.lower())
        reserved = set(k.lower() for k in key) | set([of.lower()]) | src_lower
        if calc.lower() in reserved:
            raise LinkedInputError(path + ": calculation name " + repr(calc)
                                   + " collides with a source column; Sigma silently renames it to '"
                                   + calc + " (1)' - pick a distinct name such as 'Baseline " + of + "'")
        if src_cols is not None and of not in src_cols:
            raise LinkedInputError(path + ".of " + repr(of) + " is not in sourceColumns")
        if linked.lower() in names:
            raise LinkedInputError(path + ": linked column name " + repr(linked) + " is already used")
        names[linked.lower()] = linked
        types[linked] = "number"
        out["inherited"].append({"name": calc, "aggregate": agg, "of": of, "as": linked})

    editable = spec.get("editable")
    if not isinstance(editable, list) or not editable:
        raise LinkedInputError("editable must list at least one editable column")
    out["editable"] = []
    for index, item in enumerate(editable):
        path = "editable[" + str(index) + "]"
        if not isinstance(item, dict):
            raise LinkedInputError(path + " must be an object")
        ename = _name(item.get("name"), path + ".name")
        etype = item.get("type")
        if etype not in COLUMN_TYPES:
            raise LinkedInputError(path + ".type must be one of " + ", ".join(COLUMN_TYPES))
        if ename.lower() in names:
            raise LinkedInputError(path + ": column name " + repr(ename)
                                   + " collides with another column (key, inherited, editable or formula)")
        names[ename.lower()] = ename
        types[ename] = etype
        out["editable"].append({"name": ename, "type": etype})

    out["formulas"] = []
    for index, item in enumerate(spec.get("formulas") or []):
        path = "formulas[" + str(index) + "]"
        if not isinstance(item, dict):
            raise LinkedInputError(path + " must be an object")
        fname = _name(item.get("name"), path + ".name")
        if fname.lower() in names:
            raise LinkedInputError(path + ": formula column name " + repr(fname)
                                   + " collides with another column; a collision silently becomes '"
                                   + fname + " (1)'")
        tree = _formula(item.get("formula"), types, path + ".formula")
        names[fname.lower()] = fname
        types[fname] = "number" if _numeric(tree, types) else None
        out["formulas"].append({"name": fname, "formula": tree})

    table = spec.get("table") or {}
    mode = table.get("editMode", "draft")
    if mode not in EDIT_MODES:
        raise LinkedInputError("table.editMode must be one of " + ", ".join(EDIT_MODES))
    out["table"] = {"name": _name(table.get("name", "Linked input"), "table.name"), "editMode": mode,
                    "frame": _frame(table.get("frame"), (1, 3, 24, 16)),
                    "baselineName": _name(table.get("baselineName", table.get("name", "Linked input") + " baseline"),
                                          "table.baselineName")}
    out["baseFrame"] = _frame(spec.get("baseFrame"), (1, 1, 24, 12))
    out["planBaseFrame"] = _frame(spec.get("planBaseFrame"), (1, 13, 24, 12))

    rows = spec.get("rows") or []
    if not isinstance(rows, list):
        raise LinkedInputError("rows must be a list")
    if len(rows) > MAX_ROWS:
        raise LinkedInputError("more than " + str(MAX_ROWS) + " seed rows; split them across update calls")
    edit_types = {e["name"]: e["type"] for e in out["editable"]}
    out["rows"] = []
    seen = {}
    for index, row in enumerate(rows):
        path = "rows[" + str(index) + "]"
        if not isinstance(row, dict) or set(row) - {"key", "values"}:
            raise LinkedInputError(path + " must be {key: {...}, values: {...}}")
        rkey = row.get("key")
        if not isinstance(rkey, dict) or set(rkey) != set(key):
            raise LinkedInputError(path + ".key must have exactly the key columns " + ", ".join(key))
        ident = []
        for k in key:
            value = rkey[k]
            if value is None or isinstance(value, bool) or not isinstance(value, (str, int, float)):
                raise LinkedInputError(path + ".key." + k + " must be a non-empty string or number")
            if isinstance(value, str) and not value.strip():
                raise LinkedInputError(path + ".key." + k + " is empty")
            if isinstance(value, float) and not math.isfinite(value):
                raise LinkedInputError(path + ".key." + k + " must be finite")
            ident.append((type(value).__name__, value))
        ident = tuple(ident)
        if ident in seen:
            raise LinkedInputError(path + ": duplicate seed key " + repr(rkey) + " (first at rows["
                                   + str(seen[ident]) + "])")
        seen[ident] = index
        values = row.get("values")
        if not isinstance(values, dict) or not values:
            raise LinkedInputError(path + ".values must set at least one editable column")
        clean = {}
        for column, value in values.items():
            if column not in edit_types:
                raise LinkedInputError(path + ".values." + column + " is not an editable column")
            clean[column] = _cell(value, edit_types[column], path + ".values." + column)
        out["rows"].append({"key": dict(rkey), "values": clean})
    out["warnings"] = warnings
    out["columnTypes"] = types
    return out


# ------------------------------------------------------------------------ programs


def _lit(value):
    if isinstance(value, dt.datetime):
        parts = [value.year, value.month, value.day, value.hour]
        if value.minute or value.second or value.microsecond:
            parts += [value.minute, value.second]
        if value.microsecond:
            parts.append(value.microsecond)
        return "datetime(" + ", ".join(str(p) for p in parts) + ")"
    return repr(value)


def _id(ids, key, spec=None):
    if ids and ids.get(key):
        return ids[key]
    if key == "pageId" and spec and spec["page"].get("id"):
        return spec["page"]["id"]
    return PLACEHOLDERS[key]


def _formula_expr(node, var):
    if "col" in node:
        return var[node["col"]]
    if "const" in node:
        return "sigma.formula.const_num(" + repr(node["const"]) + ")"
    fn = FORMULA_BINARY.get(node["op"]) or FORMULA_UNARY[node["op"]]
    return "sigma.formula." + fn + "(" + ", ".join(_formula_expr(a, var) for a in node["args"]) + ")"


def _target(spec, existing_id=None):
    if existing_id:
        return {"type": "existing", "workbookId": existing_id}
    wb = spec["workbook"]
    if "id" in wb:
        return {"type": "existing", "workbookId": wb["id"]}
    return {"type": "new", "name": wb["name"], "description": wb["description"]}


def call1_program(spec):
    new_wb = "id" not in spec["workbook"]
    src = spec["source"]
    key = spec["key"]
    lines = ["# excel-to-sigma linked input table - call 1 of 3: base + one-row-per-key grouped table",
             "# target: " + ("new workbook " + repr(spec["workbook"]["name"]) if new_wb
                             else "existing workbook " + spec["workbook"]["id"])]
    if new_wb:
        lines.append("old_pages = sigma.workbook.list_pages()")
    lines += ["data_page = sigma.workbook.add_page(" + repr(spec["dataPage"]["name"]) + ")",
              "data_page.set_visibility(False)"]
    if "name" in spec["page"]:
        lines.append("page = sigma.workbook.add_page(" + repr(spec["page"]["name"]) + ")")
    if src["kind"] == "dataModel":
        sel = repr(src["memberColumn"])
        lines += ["dm = sigma.data.source.data_model(" + repr(src["inodeId"]) + ")",
                  "members = [t for t in dm.list_tables() if any([c.name() == " + sel + " for c in t.columns()])]",
                  "if len(members) != 1:",
                  "    raise RuntimeError('expected exactly one data model member with column '"
                  " + " + sel + " + ', found ' + str(len(members)))",
                  "base = data_page.add_table(members[0])"]
    else:
        lines.append("base = data_page.add_table(sigma.data.source.element(" + repr(src["elementId"]) + "))")
    bf, pf = spec["baseFrame"], spec["planBaseFrame"]
    lines += ["base.set_name(" + repr(spec["table"]["baselineName"] + " source") + ")",
              "base.set_frame(" + ", ".join(str(v) for v in bf) + ")",
              "plan_base = data_page.add_table(base.to_source())",
              "plan_base.set_name(" + repr(spec["table"]["baselineName"]) + ")",
              "plan_base.set_frame(" + ", ".join(str(v) for v in pf) + ")"]
    keys = ", ".join("plan_base.get_column_by_name(" + repr(k) + ")" for k in key)
    if spec["inherited"]:
        lines.append("g = plan_base.set_groupings([" + keys + "])[0]")
        for item in spec["inherited"]:
            lines.append("g.set_calculation(" + repr(item["name"]) + ", sigma.formula."
                         + AGGREGATES[item["aggregate"]] + "(plan_base.get_column_by_name(" + repr(item["of"]) + ")))")
    else:
        lines.append("plan_base.set_groupings([" + keys + "])")
    if new_wb:
        lines += ["for p in old_pages:", "    p.delete()"]
    lines += ["print('dataPageId=' + data_page.id())",
              "print('pageId=' + page.id())" if "name" in spec["page"] else
              "print('pageId=' + " + repr(spec["page"]["id"]) + ")",
              "print('planBaseId=' + plan_base.id())"]
    return "\n".join(lines) + "\n"


def _lt_vars(spec):
    var = {}
    for i, k in enumerate(spec["key"]):
        var[k] = "k" + str(i)
    for i, item in enumerate(spec["inherited"]):
        var[item["as"]] = "i" + str(i)
    for i, item in enumerate(spec["editable"]):
        var[item["name"]] = "e" + str(i)
    for i, item in enumerate(spec["formulas"]):
        var[item["name"]] = "f" + str(i)
    return var


def call2_program(spec, ids=None):
    var = _lt_vars(spec)
    t = spec["table"]
    lookup = "{" + ", ".join(repr(k) + ": " + repr(k) for k in spec["key"]) + "}"
    lines = ["# excel-to-sigma linked input table - call 2 of 3: linked input table + columns",
             "# target: existing workbook (the id call 1 returned)",
             "data_page = sigma.workbook.get_page(" + repr(_id(ids, "dataPageId")) + ")",
             "page = sigma.workbook.get_page(" + repr(_id(ids, "pageId", spec)) + ")",
             "if not (isinstance(data_page, Page) and isinstance(page, Page)):",
             "    raise RuntimeError('re-fetch pages')",
             "plan_base = data_page.get_element(" + repr(_id(ids, "planBaseId")) + ")",
             "if not isinstance(plan_base, LevelTable):",
             "    raise RuntimeError('re-fetch grouped table')"]
    call = "lt = page.add_linked_input_table(plan_base.to_source(), lookup_columns=" + lookup
    if spec["inherited"]:
        call += ", source_columns={" + ", ".join(repr(i["name"]) + ": " + repr(i["as"]) for i in spec["inherited"]) + "}"
    lines += [call + ")",
              "lt.set_name(" + repr(t["name"]) + ")",
              "lt.set_edit_mode(" + repr(t["editMode"]) + ")",
              "lt.set_frame(" + ", ".join(str(v) for v in t["frame"]) + ")"]
    for k in spec["key"]:
        lines.append(var[k] + " = lt.get_column_by_name(" + repr(k) + ")")
    for item in spec["inherited"]:
        lines.append(var[item["as"]] + " = lt.get_column_by_name(" + repr(item["as"]) + ")")
    for item in spec["editable"]:
        lines.append(var[item["name"]] + " = lt.add_editable_column(" + repr(item["name"]) + ", " + repr(item["type"]) + ")")
    for item in spec["formulas"]:
        lines.append(var[item["name"]] + " = lt.add_formula_column(" + repr(item["name"]) + ", "
                     + _formula_expr(item["formula"], var) + ")")
    lines += ["pk = lt.list_primary_key_columns()",
              "print('primaryKeyColumns=' + str(len(pk)))",
              "print('linkedTableId=' + lt.id())"]
    for name, v in var.items():
        lines.append("print(" + repr("column:" + name + "=") + " + " + v + ".id())")
    return "\n".join(lines) + "\n"


def call3_program(spec, ids=None):
    rows = spec["rows"]
    if not rows:
        raise LinkedInputError("no seed rows: nothing for call 3 to write")
    used = []
    for e in spec["editable"]:
        if any(e["name"] in r["values"] for r in rows):
            used.append(e["name"])
    kvar = {k: "k" + str(i) for i, k in enumerate(spec["key"])}
    evar = {n: "v" + str(i) for i, n in enumerate(used)}
    lines = ["# excel-to-sigma linked input table - call 3 of 3: seed / update rows by key",
             "# target: existing workbook (the id call 1 returned)",
             "page = sigma.workbook.get_page(" + repr(_id(ids, "pageId", spec)) + ")",
             "if not isinstance(page, Page):",
             "    raise RuntimeError('re-fetch page')",
             "lt = page.get_element(" + repr(_id(ids, "linkedTableId")) + ")",
             "if not isinstance(lt, LinkedInputTable):",
             "    raise RuntimeError('expected linked input table')"]
    for k, v in kvar.items():
        lines.append(v + " = lt.get_column_by_name(" + repr(k) + ")")
    for n, v in evar.items():
        lines += [v + " = lt.get_column_by_name(" + repr(n) + ")",
                  "if not isinstance(" + v + ", LiteralEditableColumn):",
                  "    raise RuntimeError(" + repr("expected editable column: " + n) + ")"]
    lines.append("lt.update_rows([")
    for r in rows:
        sel = "{" + ", ".join(kvar[k] + ": " + _lit(r["key"][k]) for k in spec["key"]) + "}"
        val = "{" + ", ".join(evar[n] + ": " + _lit(r["values"][n]) for n in used if n in r["values"]) + "}"
        lines.append("    (" + sel + ", " + val + "),")
    lines += ["])", "print('updatedKeys=" + str(len(rows)) + "')"]
    return "\n".join(lines) + "\n"


# ------------------------------------------------------------------------ readback


def _q(name):
    return '"' + name.replace('"', '""') + '"'


def _sql_lit(value):
    if isinstance(value, str):
        return "'" + value.replace("'", "''") + "'"
    return repr(value)


def readback_queries(spec, ids=None):
    """Readback SQL for the workbook `query` tool (type "workbook").

    ``ids["describe"]`` = {"elementId": ..., "columns": {name: columnId}} as printed by
    ``describe`` of the linked element; missing ids stay as ``<...>`` placeholders.
    """
    desc = (ids or {}).get("describe") or {}
    columns = desc.get("columns") or {}
    element = '"workbook".' + _q(desc.get("elementId") or "<ELEMENT ID>")

    def col(name):
        return _q(columns.get(name) or "<COLUMN ID: " + name + ">")

    types = spec["columnTypes"]
    keys = spec["key"]
    all_names = list(keys) + [i["as"] for i in spec["inherited"]] + [e["name"] for e in spec["editable"]] \
        + [f["name"] for f in spec["formulas"]]
    queries = [{"name": "rows",
                "sql": "SELECT " + ", ".join(col(n) + " AS " + _q(n) for n in all_names)
                + " FROM " + element + " ORDER BY 1"}]
    if len(keys) == 1:
        distinct = "COUNT(DISTINCT " + col(keys[0]) + ")"
    else:
        distinct = "COUNT(DISTINCT CONCAT(" + ", '|', ".join("CAST(" + col(k) + " AS string)" for k in keys) + "))"
    queries.append({"name": "distinct_keys",
                    "sql": "SELECT COUNT(*) AS \"row_count\", " + distinct + " AS \"distinct_keys\" FROM " + element})
    editables = [e["name"] for e in spec["editable"]]
    queries.append({"name": "non_null_editable",
                    "sql": "SELECT " + ", ".join("COUNT(" + col(n) + ") AS " + _q(n + " non-null") for n in editables)
                    + " FROM " + element})
    sums = [e["name"] for e in spec["editable"] if e["type"] == "number"]
    sums += [i["as"] for i in spec["inherited"]]
    sums += [f["name"] for f in spec["formulas"] if types.get(f["name"]) == "number"]
    if sums:
        queries.append({"name": "sums",
                        "sql": "SELECT " + ", ".join("SUM(" + col(n) + ") AS " + _q(n + " sum") for n in sums)
                        + " FROM " + element})
    if spec["rows"] and len(keys) == 1:
        keyed = ", ".join(_sql_lit(r["key"][keys[0]]) for r in spec["rows"])
        queries.append({"name": "seeded_rows",
                        "sql": "SELECT " + ", ".join(col(n) + " AS " + _q(n) for n in all_names)
                        + " FROM " + element + " WHERE " + col(keys[0]) + " IN (" + keyed + ") ORDER BY 1"})
    expected = {"seededKeys": len(spec["rows"]),
                "nonNullEditable": {n: sum(1 for r in spec["rows"] if n in r["values"]) for n in editables},
                "numericEditableSums": {n: round(sum(r["values"][n] for r in spec["rows"] if n in r["values"]), 12)
                                        for n in sums if n in editables},
                "oneRowPerKey": "row_count must equal distinct_keys; both equal the source's distinct key count",
                "note": "expected sums assume the editable columns were empty before the seed"}
    return {"queries": queries, "expected": expected}


# ------------------------------------------------------------------------- linting


def lint_program(program):
    """Return a list of Code Mode dialect violations for a program text."""
    problems = []
    tree = ast.parse(program)
    assigned = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            assigned.add(node.id)
    for node in ast.walk(tree):
        line = getattr(node, "lineno", 0)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda, ast.While)):
            problems.append("line " + str(line) + ": " + type(node).__name__ + " is not in the dialect")
        elif isinstance(node, (ast.Import, ast.ImportFrom, ast.JoinedStr)):
            problems.append("line " + str(line) + ": " + type(node).__name__ + " is not in the dialect")
        elif isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Mod, ast.Pow, ast.FloorDiv)):
            problems.append("line " + str(line) + ": operator " + type(node.op).__name__ + " is not in the dialect")
        elif isinstance(node, ast.Subscript) and isinstance(node.slice, ast.Slice):
            problems.append("line " + str(line) + ": slices are not in the dialect")
        elif isinstance(node, (ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)) and len(node.generators) > 1:
            problems.append("line " + str(line) + ": single-clause comprehensions only")
        elif isinstance(node, ast.Attribute) and node.attr == "format":
            problems.append("line " + str(line) + ": .format is not in the dialect")
        elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            name = node.id
            if (name not in assigned and name not in ALLOWED_BUILTINS and name != "sigma"
                    and not (name[:1].isupper() and any(ch.islower() for ch in name))):
                problems.append("line " + str(line) + ": name " + name + " is not assigned or exposed")
    previous = None
    for token in tokenize.generate_tokens(io.StringIO(program).readline):
        if token.type in (tokenize.NL, tokenize.NEWLINE, tokenize.COMMENT, tokenize.INDENT, tokenize.DEDENT):
            if token.type in (tokenize.NEWLINE, tokenize.INDENT, tokenize.DEDENT):
                previous = None
            continue
        if token.type == tokenize.STRING and previous == tokenize.STRING:
            problems.append("line " + str(token.start[0]) + ": adjacent string literals; use +")
        previous = token.type
    return problems


# --------------------------------------------------------------------------- emit


def emit(spec, ids=None):
    """Return {programs, readback, manifest, warnings} for a raw spec."""
    spec = normalize_spec(spec)
    ids = dict(ids or {})
    programs = {"call1": call1_program(spec), "call2": call2_program(spec, ids)}
    if spec["rows"]:
        programs["call3"] = call3_program(spec, ids)
    for name, text in programs.items():
        problems = lint_program(text)
        if problems:
            raise LinkedInputError("internal: " + name + " violates the Code Mode dialect: " + "; ".join(problems))
    readback = readback_queries(spec, ids)
    calls = []
    for index, name in enumerate(sorted(programs), start=1):
        text = programs[name]
        calls.append({"call": index, "file": name + ".py",
                      "target": _target(spec) if index == 1 else _target(spec, ids.get("workbookId") or "<WORKBOOK ID>"),
                      "unresolvedPlaceholders": sorted(set(re.findall(r"<[A-Z ]+ ID>", text)))})
    manifest = {"calls": calls,
                "idsToCapture": {"call1": ["dataPageId", "pageId", "planBaseId"],
                                 "call2": ["linkedTableId"], "describe": ["elementId", "columns"]},
                "unvalidated": ["published-mode edit by a non-editor user",
                                "warehouse view over a linked input table"] + spec["warnings"]}
    return {"programs": programs, "readback": readback, "manifest": manifest, "warnings": spec["warnings"]}


def parse_printed_ids(text):
    """Collect the key=value lines the programs print (dataPageId=..., column:Name=...)."""
    ids = {}
    columns = {}
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith("column:") and "=" in line:
            name, _, value = line[len("column:"):].rpartition("=")
            columns[name] = value.strip()
            continue
        for key in ID_KEYS + ("primaryKeyColumns",):
            if line.startswith(key + "="):
                ids[key] = line[len(key) + 1:].strip()
    if columns:
        ids["printedColumnIds"] = columns
    return ids
