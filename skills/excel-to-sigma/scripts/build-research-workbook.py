#!/usr/bin/env python3
"""
build-research-workbook.py — MULTI-SHEET, MULTI-TABLE driver for the equity-research
archetype. Converts a whole broker model (many analysis sheets) into ONE Sigma workbook:

  * one PAGE per source sheet, holding a pivot per sub-table on that sheet
  * each sub-table converted on its OWN year axis (detect_tables) — no cross-table year
    shift; col_ids are namespaced per sub-table so nothing collides across sheets
  * a hidden `plumbing` page with the calc / transpose / join elements
  * the FY-results model can carry an EDITABLE input layer (--editable): the analyst's
    hardcoded inputs become an input table that Coalesce-overrides the inline-VALUES
    defaults (so parity holds before seeding; edit to override). Seed CSV emitted for the
    one-time UI paste.

Every workbook POST goes through workbook_wire (released document/layout contract).

Usage:
  build-research-workbook.py <file.xlsx> --conn <id> --folder <id> [--name "..."]
        [--sheets "FY results,Geo segments,..."]   # default: auto-pick financial sheets
        [--editable-sheet "FY results"] [--post]   # default dry-run to /tmp
"""
import sys, os, json, csv, re, importlib.util, warnings
warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
def _load(fn, name):
    s = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(s); s.loader.exec_module(m); return m
INF = _load("infer-canonical-formulas.py", "inf")
BRM = _load("build-research-model.py", "brm")

NUM = {"kind": "number", "formatString": "#,##0.0"}
DMP = "Custom SQL"


def abbr(title):
    a = re.sub(r"[^a-z0-9]", "", title.lower())[:6] or "s"
    return a


def namespace_plan(plan, pre):
    """Rewrite col_ids -> <pre>c<row> and every [cN] ref in canonicals, so a sub-table's
    columns are globally unique across the workbook."""
    ren = {ln["col_id"]: f"{pre}{ln['col_id']}" for ln in plan["lines"]}
    for ln in plan["lines"]:
        ln["col_id"] = ren[ln["col_id"]]
        if ln.get("canonical"):
            ln["canonical"] = re.sub(r"\[c(\d+)\]", lambda m: f"[{pre}c{m.group(1)}]", ln["canonical"])
    return plan


def subtable_plans(path, sheet):
    """All sub-tables on a sheet, each an inferred + namespaced plan (skips empties/failures)."""
    wf, wv = INF.load(path);
    if sheet not in wf.sheetnames:
        return []
    tables = INF.detect_tables(wf[sheet], wv[sheet]) or []
    if not tables:                                       # single-table sheet: normal detect
        try:
            p = INF.infer(path, sheet=sheet)
            return [(f"{abbr(sheet)}0", p, None)]
        except Exception:
            return []
    out = []
    for i, t in enumerate(tables):
        try:
            p = INF.infer(path, sheet=sheet, axis_override=t["axis"], header_row=t["header_row"],
                          first_row=t["first_row"], last_row=t["last_row"])
        except Exception:
            continue
        if not p["lines"] or not p["periods"]:
            continue
        pre = f"{abbr(sheet)}{i}"
        out.append((pre, namespace_plan(p, pre), t))
    return out


def dm_elements(pre, plan):
    """DM `<pre>_data` (inline VALUES, editable-default) + `<pre>_dim` (line dim)."""
    data_lines = [ln for ln in plan["lines"] if ln["typed"]]
    years = [p["year"] for p in plan["periods"]]
    cols = ["YR"] + [ln["col_id"] for ln in data_lines]
    rows = []
    for y in years:
        cells = [str(y)] + ["NULL" if (ln["typed"].get(str(y), ln["typed"].get(y)) is None)
                            else f"{float(ln['typed'].get(str(y), ln['typed'].get(y))):.6f}"
                            for ln in data_lines]
        rows.append("(" + ",".join(cells) + ")")
    inner = "select * from (values\n" + ",\n".join(rows) + "\n) as t(" + ", ".join(cols) + ")"
    casts = ["YR::int as YR"] + [f"{ln['col_id']}::float as {ln['col_id']}" for ln in data_lines]
    data_sql = f"select {', '.join(casts)} from (\n{inner}\n) s"
    dim_sql = ("select * from (values\n" +
               ",\n".join("(" + ",".join([BRM.q(ln["col_id"]), BRM.q(ln["label"] or ln["col_id"]),
                                          BRM.q(plan["sheet"]), str(ln["sec_order"]), str(ln["line_order"])]) + ")"
                          for ln in plan["lines"]) +
               "\n) as t(LINE_ID, LABEL, SECTION, SEC_ORDER, LINE_ORDER)")
    data_el = {"id": f"{pre}_data", "kind": "table",
               "source": {"kind": "sql", "connectionId": CONN, "statement": data_sql},
               "columns": [BRM.dmcol(0, "YR", "Year")] +
                          [BRM.dmcol(i + 1, ln["col_id"], ln["col_id"]) for i, ln in enumerate(data_lines)]}
    dim_el = {"id": f"{pre}_dim", "kind": "table",
              "source": {"kind": "sql", "connectionId": CONN, "statement": dim_sql},
              "columns": [BRM.dmcol(0, "LINE_ID", "Line Id"), BRM.dmcol(1, "LABEL", "Label"),
                          BRM.dmcol(2, "SECTION", "Section"), BRM.dmcol(3, "SEC_ORDER", "Sec Order"),
                          BRM.dmcol(4, "LINE_ORDER", "Line Order")]}
    return data_el, dim_el, data_lines


def wb_elements(pre, plan, dm_id, data_el_id, dim_el_id, data_lines, title, editable=False):
    """calc / transpose(long) / dimw / labeled / pivot for one sub-table.
    editable=True: the analyst INPUT lines become an editable input table that
    Coalesce-overrides the inline-VALUES defaults (parity holds until seeded)."""
    dcols = {ln["col_id"] for ln in data_lines}
    def dm(eid): return {"kind": "data-model", "dataModelId": dm_id, "elementId": eid}
    extra, calc_src = [], dm(data_el_id)
    input_cids = [ln["col_id"] for ln in data_lines if ln["kind"] == "input"]
    if editable and input_cids:
        # wrap the data DM element (join leg) + an empty editable input table, joined on Year
        dw = {"id": f"{pre}_dw", "kind": "table", "name": f"{pre}DW", "visibleAsSource": False,
              "source": dm(data_el_id),
              "columns": [{"id": f"{pre}_wyr", "name": "Year", "formula": f"[{DMP}/Year]"}] +
                         [{"id": f"{pre}_w_{c}", "name": c, "formula": f"[{DMP}/{c}]"} for c in dcols]}
        inp = {"id": f"{pre}_inp", "kind": "input-table", "name": f"{title} — inputs (editable)",
               "inputMode": "edit", "source": {"kind": "empty", "connectionId": CONN},
               "columns": [{"id": "ID"}, {"id": f"{pre}IY", "type": "number", "name": "Year"}] +
                          [{"id": f"{pre}_i_{c}", "type": "number", "name": c} for c in input_cids],
               "order": [f"{pre}IY"] + [f"{pre}_i_{c}" for c in input_cids]}
        calc_src = {"kind": "join", "name": f"{pre}JN",
                    "primarySource": {"kind": "table", "elementId": f"{pre}_dw"},
                    "joins": [{"name": f"{pre}INP", "joinType": "left-outer",
                               "left": {"kind": "table", "elementId": f"{pre}_dw"},
                               "right": {"kind": "table", "elementId": f"{pre}_inp"},
                               "columns": [{"left": "[Year]", "right": "[Year]"}]}]}
        extra = [dw, inp]
    src_pref = f"{pre}JN" if (editable and input_cids) else DMP
    calc_cols = [{"id": f"{pre}_yr", "name": "Year", "formula": f"[{src_pref}/Year]"}]
    merge = []
    for ln in plan["lines"]:
        cid = ln["col_id"]; merge.append(cid); has = cid in dcols
        if editable and input_cids and ln["kind"] == "input" and has:
            f = f"Coalesce([{pre}INP/{cid}], [{src_pref}/{cid}])"     # analyst edit wins, else default
        elif ln["kind"] == "input" or not ln.get("canonical"):
            f = f"[{src_pref}/{cid}]" if has else "Null"
        elif has:
            f = f"Coalesce([{src_pref}/{cid}], {ln['canonical']})"
        else:
            f = ln["canonical"]
        calc_cols.append({"id": cid, "name": cid, "formula": f})
    calcname = f"{pre}Calc"
    calc = {"id": f"{pre}_calc", "kind": "table", "name": calcname, "visibleAsSource": False,
            "source": calc_src, "columns": calc_cols}
    TR = f"Transpose of {calcname}"
    long_el = {"id": f"{pre}_long", "kind": "table", "visibleAsSource": False,
               "source": {"kind": "transpose", "source": {"kind": "table", "elementId": f"{pre}_calc"},
                          "direction": "column-to-row", "columnsToMerge": merge,
                          "columnLabelForMergedColumns": "LineId", "columnLabelForValues": "Value"},
               "columns": [{"id": f"{pre}_lyr", "name": "Year", "formula": f"[{TR}/Year]"},
                           {"id": f"{pre}_lid", "name": "LineId", "formula": f"[{TR}/LineId]"},
                           {"id": f"{pre}_lv", "name": "Value", "formula": f"[{TR}/Value]"}]}
    dimw = {"id": f"{pre}_dimw", "kind": "table", "name": f"{pre}Dim", "visibleAsSource": False,
            "source": dm(dim_el_id),
            "columns": [{"id": f"{pre}_did", "name": "Line Id", "formula": f"[{DMP}/Line Id]"},
                        {"id": f"{pre}_dlab", "name": "Label", "formula": f"[{DMP}/Label]"},
                        {"id": f"{pre}_dso", "name": "Sec Order", "formula": f"[{DMP}/Sec Order]"},
                        {"id": f"{pre}_dlo", "name": "Line Order", "formula": f"[{DMP}/Line Order]"}]}
    labeled = {"id": f"{pre}_labeled", "kind": "table", "name": f"{pre}Labeled", "visibleAsSource": False,
               "source": {"kind": "join", "name": f"{pre}J",
                          "primarySource": {"kind": "table", "elementId": f"{pre}_long"},
                          "joins": [{"name": f"{pre}D", "joinType": "left-outer",
                                     "left": {"kind": "table", "elementId": f"{pre}_long"},
                                     "right": {"kind": "table", "elementId": f"{pre}_dimw"},
                                     "columns": [{"left": "[LineId]", "right": "[Line Id]"}]}]},
               "columns": [{"id": f"{pre}_xyr", "name": "Year", "formula": f"[{pre}J/Year]"},
                           {"id": f"{pre}_xv", "name": "Value", "formula": f"[{pre}J/Value]"},
                           {"id": f"{pre}_xlab", "name": "Line Item", "formula": f"[{pre}D/Label]"},
                           {"id": f"{pre}_xso", "name": "Sec Order", "formula": f"[{pre}D/Sec Order]"},
                           {"id": f"{pre}_xlo", "name": "Line Order", "formula": f"[{pre}D/Line Order]"}]}
    pivot = {"id": f"{pre}_pivot", "kind": "pivot-table", "name": title,
             "source": {"kind": "table", "elementId": f"{pre}_labeled"},
             "columns": [{"id": f"{pre}_pl", "name": "Line Item", "formula": f"[{pre}Labeled/Line Item]"},
                         {"id": f"{pre}_plo", "name": "Line Order", "formula": f"[{pre}Labeled/Line Order]"},
                         {"id": f"{pre}_py", "name": "Year", "formula": f"[{pre}Labeled/Year]"},
                         {"id": f"{pre}_pv", "name": "Value", "formula": f"Sum([{pre}Labeled/Value])"}],
             "values": [f"{pre}_pv"],
             "rowsBy": [{"id": f"{pre}_pl", "sort": {"direction": "ascending", "by": f"{pre}_plo", "aggregation": "min"}}],
             "columnsBy": [{"id": f"{pre}_py", "sort": {"direction": "ascending"}}]}
    inputs = [el for el in extra if el["kind"] == "input-table"]
    plumb = [el for el in extra if el["kind"] != "input-table"] + [calc, long_el, dimw, labeled]
    return {"pivot": pivot, "plumb": plumb, "inputs": inputs}


def main():
    global CONN
    a = [x for x in sys.argv[1:] if not x.startswith("--")]
    opt = {x.split("=")[0]: (x.split("=", 1)[1] if "=" in x else True) for x in sys.argv[1:] if x.startswith("--")}
    if not a:
        sys.exit(__doc__)
    path = a[0]; CONN = opt.get("--conn", "<conn>"); folder = opt.get("--folder", "<folder>")
    name = opt.get("--name") or (os.path.splitext(os.path.basename(path))[0][:50] + " (auto)")
    do_post = "--post" in opt
    wf, wv = INF.load(path)
    if opt.get("--sheets"):
        sheets = [s.strip() for s in opt["--sheets"].split(",")]
    else:                                                # auto: sheets with a detectable financial table
        sheets = []
        for ws in wf.worksheets:
            if ws.sheet_state != "visible" or ws.title.startswith("__") or ws.title.endswith(">>"):
                continue
            try:
                if INF.detect_tables(wf[ws.title], wv[ws.title]):
                    sheets.append(ws.title)
            except Exception:
                pass
    print("sheets to convert:", sheets)

    dm_data, dm_dim, dm_els = [], [], {}
    pages, plumb_all, verify = [], [], []
    for sheet in sheets:
        subs = subtable_plans(path, sheet)
        page_pivots = []
        for j, (pre, plan, tbl) in enumerate(subs):
            de, di, dl = dm_elements(pre, plan)
            if not dl:                                   # nothing typed -> skip (all-derived/empty)
                continue
            dm_data.append(de); dm_dim.append(di)
            dm_els[pre] = {"data_cols": {c["name"] for c in de["columns"]},
                           "dim_cols": {c["name"] for c in di["columns"]}, "plan": plan, "dl": dl,
                           "title": f"{sheet}" + (f" — table {j+1}" if len(subs) > 1 else "")}
            page_pivots.append(pre)
        if page_pivots:
            pages.append((sheet, page_pivots))
    if not dm_data:
        sys.exit("no convertible tables found")

    dm_spec = {"name": name + " — Model", "schemaVersion": 1, "folderId": folder,
               "pages": [{"id": "m", "name": "Model", "elements": dm_data + dm_dim}]}
    os.makedirs("/tmp/research-wb", exist_ok=True)
    json.dump(dm_spec, open("/tmp/research-wb/dm.json", "w"), indent=1)
    print(f"DM: {len(dm_data)} data + {len(dm_dim)} dim elements across {len(pages)} pages")
    if not do_post:
        print("DRY RUN — /tmp/research-wb/dm.json written. Re-run with --post --conn --folder to build.")
        return

    e = BRM.env(); base = e["SIGMA_BASE_URL"]; tok = BRM.token(e)
    dm = BRM.api(base, tok, "POST", "/v2/dataModels/spec", dm_spec); dm_id = dm["dataModelId"]
    print("dataModelId:", dm_id)
    full = BRM.api(base, tok, "GET", f"/v2/dataModels/{dm_id}/spec")
    # map reassigned ids by column-name signature (col_ids are namespaced -> unique)
    id_by_cols = {frozenset(c["name"] for c in el["columns"]): el["id"] for el in full["pages"][0]["elements"]}
    edit_sheet = opt.get("--editable-sheet", "FY results")
    plumb_all, main_pages, seeds = [], [], []
    for sheet, pres in pages:
        pv = []
        for k, pre in enumerate(pres):
            info = dm_els[pre]
            de_id = id_by_cols.get(frozenset(info["data_cols"]))
            di_id = id_by_cols.get(frozenset(info["dim_cols"]))
            editable = (sheet == edit_sheet and k == 0)          # primary table of the model sheet
            els = wb_elements(pre, info["plan"], dm_id, de_id, di_id, info["dl"], info["title"], editable=editable)
            pv.append(els["pivot"]); plumb_all += els["plumb"]
            if els.get("inputs"):
                pv += els["inputs"]                               # editable input table shown on the page
                seeds.append((pre, info["plan"], info["dl"]))
        main_pages.append({"id": abbr(sheet), "name": sheet[:40], "elements": pv})
    main_pages.append({"id": "plumbing", "name": "plumbing", "visibility": "hidden", "elements": plumb_all})
    for pre, plan, dl in seeds:                                   # seed CSV for the one-time UI paste
        icols = [ln for ln in dl if ln["kind"] == "input"]
        sp = f"/tmp/research-wb/{pre}_inputs_seed.csv"
        with open(sp, "w", newline="") as fh:
            w = csv.writer(fh); w.writerow(["Year"] + [ln["col_id"] for ln in icols])
            for p in plan["periods"]:
                y = p["year"]; row = [y]
                for ln in icols:
                    v = ln["typed"].get(str(y), ln["typed"].get(y))
                    row.append("" if v is None else round(v, 6))
                w.writerow(row)
        with open(f"/tmp/research-wb/{pre}_inputs_labels.csv", "w", newline="") as fh:
            w = csv.writer(fh); w.writerow(["col_id", "label"])
            for ln in icols: w.writerow([ln["col_id"], ln["label"]])
        print(f"seed CSV: {sp}  (+ {pre}_inputs_labels.csv for the col_id->label map)")
    wb_spec = {"name": name, "schemaVersion": 1, "kind": "workbook", "folderId": folder, "pages": main_pages}
    wb_spec = _load("lib/workbook_wire.py", "ww").wire_workbook(wb_spec) if os.path.exists(os.path.join(HERE, "lib/workbook_wire.py")) else wb_spec
    json.dump(wb_spec, open("/tmp/research-wb/wb.json", "w"), indent=1)
    wb = BRM.api(base, tok, "POST", "/v2/workbooks/spec", wb_spec)
    print("workbookId:", wb.get("workbookId"), "| url:", wb.get("url"))


if __name__ == "__main__":
    main()
