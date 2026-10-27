#!/usr/bin/env python3
"""A reference look ("make it look like workbook X"): extraction from synthetic specs and its application to the
table-report Report page. The fixtures reproduce the SHAPES of three real org workbooks; no real spec is copied."""
from __future__ import annotations

import copy
import importlib.util
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(ROOT / "tests"))

import presentation_style as ps  # noqa: E402
import reference_look as rl  # noqa: E402
import table_report_model as trm  # noqa: E402
from test_table_report_presentation import (PIPELINE, SOURCES, draft, element, layout_tree,  # noqa: E402
                                            wb_spec)

OURS = ("#0F172A", "#64748B", "#F0F7FF", "#E2E8F0", "#FFFFFF", "#2563EB")


def kpi(index, name=None, value=None, style=None):
    return {"id": f"k{index}", "kind": "kpi-chart", "name": name if name is not None else f"KPI {index}",
            "value": {"columnId": f"v{index}", **(value or {})}, **({"style": style} if style else {})}


NAMED = "0a1b2c3d-1111-4222-8333-444455556666"       # an org theme id (shape (a))
SPEC_A = {"workbookId": "aaaaaaaa-1111-4222-8333-444455556666", "name": "Repo Insights",
          "document": {"settings": {"theme": {"name": NAMED}},
                       "elements": [{"id": "t1", "kind": "text", "body": "{{[Repeated/Repo Name]}}"},
                                    kpi(1), kpi(2), {"id": "tb", "kind": "table"},
                                    {"id": "bc", "kind": "bar-chart"}]}}

THEME_B = {"name": "Light", "overrides": {
    "colorOverrides": [{"name": "backgroundCanvas", "color": "#f6f6f6"}],
    "categoricalScheme": ["#82badf", "#8bc34a", "#f3a24f", "#d95c59", "#c8e5a3", "#7fb4d3", "#f8dfa0", "#8bbf78"],
    "titleFont": {"fontSize": 10, "fontWeight": "bold"}}}
CARD_B = {"backgroundColor": "#ffffff", "borderColor": "#d9dee5", "borderRadius": "round"}
SPEC_B = {"workbookId": "bbbbbbbb-1111-4222-8333-444455556666", "name": "Executive Dashboard",
          "document": {"settings": {"theme": THEME_B},
                       "elements": [{"id": "h", "kind": "text", "body": "## Executive Dashboard"},
                                    {"id": "h2", "kind": "text", "body": "### Executive KPIs"},
                                    *[{"id": f"c{i}", "kind": "container", "style": dict(CARD_B)} for i in range(16)],
                                    {"id": "c-bare", "kind": "container"},
                                    *[kpi(i, value={"fontSize": 48} if i < 6 else None) for i in range(18)],
                                    *[{"id": f"b{i}", "kind": "bar-chart"} for i in range(9)],
                                    {"id": "tb", "kind": "table"}]}}

THEME_C = {"overrides": {
    "colors": {"text": "#121b39", "highlight": "#4e79a7", "surface": "#ffffff"},
    "colorOverrides": [{"name": "backgroundCanvas", "color": "#ffffff"}],
    "categoricalScheme": ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1",
                          "#ff9da7", "#9c755f", "#bab0ac"],
    "borderRadius": "pill", "elementBorder": {"width": 1, "color": "#d0d5dd"},
    "tableStyles": {"gridLines": "vertical", "cellSpacing": "medium"}}}
SPEC_C = {"workbookId": "cccccccc-1111-4222-8333-444455556666", "name": "FP&A Overview",
          "document": {"settings": {"theme": THEME_C},
                       "elements": [
                           {"id": "t1", "kind": "text",
                            "body": '<span style="font-size: 30px; color: #121b39">FP&A Executive Overview</span>'},
                           {"id": "t2", "kind": "text",
                            "body": '<span style="font-size: 13px; color: #667085">FY2026 actuals through June</span>'},
                           *[kpi(i, name={"text": f"K{i}", "color": "#121b39", "fontWeight": "bold", "fontSize": 13},
                                 value={"fontSize": 30}, style={"borderRadius": "pill"}) for i in range(4)],
                           *[{"id": f"l{i}", "kind": "line-chart", "style": {"padding": "none"}} for i in range(4)],
                           *[{"id": f"b{i}", "kind": "bar-chart", "style": {"padding": "none"}} for i in range(3)],
                           {"id": "p1", "kind": "pivot-table", "style": {"borderRadius": "pill"},
                            "tableStyle": {"preset": "presentation"}},
                           {"id": "p2", "kind": "pivot-table", "style": {"borderRadius": "pill"},
                            "tableStyle": {"preset": "presentation"}},
                           *[{"id": f"tb{i}", "kind": "table"} for i in range(4)]]}}


def with_source(spec):
    look = rl.extract_look(spec)
    look["source"] = {"workbookId": spec["workbookId"], "name": spec["name"]}
    return look


class TestExtractLook(unittest.TestCase):
    def test_a_named_org_theme_is_kept_as_a_name_and_nothing_else_is_invented(self):
        look = rl.extract_look(SPEC_A)
        self.assertEqual(look["theme"], {"name": NAMED})
        self.assertEqual(look["styles"], {"container": None, "kpi-chart": None, "chart": None, "table": None})
        self.assertEqual(look["kpi"], {"name": None, "value": None})
        self.assertEqual(look["text"], {"title": None, "subline": None, "color": None})
        self.assertIsNone(look["tablePreset"])

    def test_a_light_theme_with_overrides_is_copied_exactly(self):
        look = rl.extract_look(SPEC_B)
        self.assertEqual(look["theme"], THEME_B)
        self.assertIsNot(look["theme"], SPEC_B["document"]["settings"]["theme"])      # a copy
        self.assertEqual(look["styles"]["container"], CARD_B)                           # 16 of 17 containers
        self.assertIsNone(look["styles"]["kpi-chart"])
        self.assertEqual(look["kpi"], {"name": None, "value": None})        # fontSize 48 on 6 of 18: a minority
        self.assertEqual(look["text"]["title"], {"heading": 2})             # "## Title"
        self.assertIsNone(look["text"]["subline"])                          # "### ..." carries no span

    def test_an_overrides_only_theme_gives_kpi_and_text_type_and_a_preset(self):
        look = rl.extract_look(SPEC_C)
        self.assertEqual(look["theme"], THEME_C)
        self.assertEqual(look["styles"]["kpi-chart"], {"borderRadius": "pill"})
        self.assertEqual(look["styles"]["chart"], {"padding": "none"})
        self.assertIsNone(look["styles"]["table"])                           # plain tables have none
        self.assertEqual(look["kpi"]["name"], {"color": "#121b39", "fontSize": 13, "fontWeight": "bold"})
        self.assertEqual(look["kpi"]["value"], {"fontSize": 30})
        self.assertEqual(look["text"], {"title": {"size": 30, "color": "#121b39"},
                                        "subline": {"size": 13, "color": "#667085"}, "color": "#121b39"})
        self.assertEqual(look["tablePreset"], "presentation")

    def test_a_response_a_document_and_contents_are_all_accepted(self):
        document = SPEC_C["document"]
        self.assertEqual(rl.extract_look(document), rl.extract_look(SPEC_C))
        self.assertEqual(rl.extract_look({"contents": document}), rl.extract_look(SPEC_C))
        self.assertEqual(rl.extract_look({})["theme"], None)

    def test_workbook_ids_come_from_a_url_or_a_bare_id(self):
        wid = "fa2aeeb4-995d-4940-9dbb-e1cd3119d760"
        self.assertEqual(rl.workbook_id(wid), wid)
        self.assertEqual(rl.workbook_id(f"https://app.sigmacomputing.com/org/workbook/Golden-{wid}?:link=1"), wid)


class TestStyleReferenceCli(unittest.TestCase):
    def load(self):
        spec = importlib.util.spec_from_file_location("style_reference_cli", SCRIPTS / "style-reference.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def test_it_makes_one_read_only_get_and_writes_the_look_with_its_source(self):
        cli = self.load()
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(cli.sigma_auth, "api",
                                                                      return_value=SPEC_C) as api:
            out = Path(tmp) / "look.json"
            self.assertEqual(cli.main([f"https://x.example/workbook/FP-A-{SPEC_C['workbookId']}", "--out", str(out)]),
                             0)
            api.assert_called_once_with("GET", f"/v2/workbooks/{SPEC_C['workbookId']}/spec")
            look = json.loads(out.read_text())
        self.assertEqual(look["theme"], THEME_C)
        self.assertEqual(look["source"], {"workbookId": SPEC_C["workbookId"], "name": "FP&A Overview"})

    def test_a_saved_response_needs_no_network(self):
        cli = self.load()
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(cli.sigma_auth, "api") as api:
            saved = Path(tmp) / "spec.json"
            saved.write_text(json.dumps(SPEC_B))
            out = Path(tmp) / "look.json"
            self.assertEqual(cli.main(["--spec", str(saved), "--out", str(out)]), 0)
            api.assert_not_called()
            self.assertEqual(json.loads(out.read_text())["styles"]["container"], CARD_B)

    def test_it_refuses_to_run_without_a_workbook(self):
        cli = self.load()
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(cli.main(["--out", str(Path(tmp) / "x.json")]), 2)


URL_ID = "7C3Hwd6PvaK9h9FY53ztJe"


class TestWorkbookIds(unittest.TestCase):
    def test_a_url_gives_the_id_not_the_slug(self):
        base = "https://app.sigmacomputing.com/org/workbook/Golden-Orders-Executive-Dashboard-" + URL_ID
        self.assertEqual(rl.workbook_id(base), URL_ID)
        self.assertEqual(rl.workbook_id(base + "/page/aBcD123?x=1"), URL_ID)
        self.assertEqual(rl.workbook_id(base + "?:link_source=share#frag"), URL_ID)
        self.assertEqual(rl.workbook_id("https://app.sigmacomputing.com/org/workbook/" + URL_ID), URL_ID)

    def test_a_uuid_url_a_bare_uuid_and_a_bare_id_pass_through(self):
        uuid = "fa2aeeb4-995d-4940-9dbb-e1cd3119d760"
        self.assertEqual(rl.workbook_id(f"https://x.example/org/workbook/Golden-{uuid}/page/p1"), uuid)
        self.assertEqual(rl.workbook_id(uuid), uuid)
        self.assertEqual(rl.workbook_id(uuid.upper()), uuid)
        self.assertEqual(rl.workbook_id(URL_ID), URL_ID)
        self.assertIsNone(rl.workbook_id(""))


class TestExtractRobustness(unittest.TestCase):
    def test_hidden_pages_do_not_outvote_visible_ones(self):
        plan, _ = draft(PIPELINE)
        spec = wb_spec(plan)                                  # Report (visible), Inputs, hidden Model (sources)
        look = rl.extract_look(spec)
        self.assertEqual(look["styles"]["table"], ps.card_style())          # the Report tables, not 6 bare model ones
        self.assertEqual(look["styles"]["kpi-chart"], ps.card_style())
        self.assertEqual(look["tablePreset"], "presentation")

    def test_when_nothing_visible_is_styled_the_styled_elements_are_used(self):
        document = {"pages": [{"id": "p1", "name": "Main"}, {"id": "p2", "name": "Data", "visibility": "hidden"}],
                    "layout": '<Page id="p1"><Element elementId="a"/></Page><Page id="p2">'
                              '<Element elementId="b"/><Element elementId="c"/></Page>',
                    "elements": [{"id": "a", "kind": "table"},
                                 {"id": "b", "kind": "table", "style": {"borderRadius": "pill"}},
                                 {"id": "c", "kind": "table"}]}
        self.assertEqual(rl.extract_look({"document": document})["styles"]["table"], {"borderRadius": "pill"})

    def test_unexpected_shapes_do_not_crash(self):
        odd = [
            {"document": {"elements": [{"id": "t", "kind": "text"}, {"id": "t2", "kind": "text", "body": None}]}},
            {"document": {"elements": [{"id": "x", "kind": "table", "tableStyle": "presentation"},
                                       {"id": "y", "kind": "pivot-table", "tableStyle": None}]}},
            {"document": {"settings": None, "elements": []}},
            {"document": {"settings": {"theme": "Light"}}},
            {"document": {}},
            {"document": {"elements": None, "pages": "x", "layout": 5}},
            {"document": {"elements": [{"id": "k", "kind": "kpi-chart", "name": 5, "value": "big", "style": "x"},
                                       "not an element", None]}},
            {"document": {"settings": {"theme": {"overrides": {"colors": "dark"}}}}},
            {"elements": [{"id": "c", "kind": "container", "style": ["x"]}]},
            [], None, "text",
        ]
        for spec in odd:
            look = rl.extract_look(spec)
            self.assertEqual(look["version"], 1, spec)
            json.dumps(look)

    def test_look_values_that_go_into_span_styles_are_validated(self):
        spec = {"document": {"elements": [
            {"id": "k", "kind": "kpi-chart", "name": {"text": "n", "color": "red;} <b>", "fontSize": "huge",
                                                      "fontWeight": "heavy"},
             "value": {"columnId": "v", "fontSize": 30, "color": "#12345"}}]}}
        look = rl.extract_look(spec)
        self.assertEqual(look["kpi"]["name"], None)                                 # nothing valid survived
        self.assertEqual(look["kpi"]["value"], {"fontSize": 30, "color": "#12345"})
        hostile = {"version": 1, "theme": None, "styles": {}, "tablePreset": None,
                   "kpi": {"name": {"color": 'red"><script>', "fontSize": "9px; x", "fontWeight": "bold"},
                           "value": {"fontSize": "40;evil", "color": "#fff"}},
                   "text": {"title": {"size": "30px", "color": "blue;"}, "subline": None, "color": None}}
        _, text = ps.kpi_text_card("k", "REV", "Sum([T/x])", None, hostile)
        self.assertEqual(ps.disallowed_span_styles(text["body"]), [])
        self.assertNotIn("script", text["body"])
        self.assertNotIn("evil", text["body"])
        band, header = ps.header_band("h", "T", "S", {"tint": "#fff"}, hostile)
        self.assertNotIn("blue;", header["body"])
        self.assertEqual(ps.disallowed_span_styles(header["body"]), [])


class TestStyleReferenceErrors(unittest.TestCase):
    def load(self):
        spec = importlib.util.spec_from_file_location("style_reference_cli_err", SCRIPTS / "style-reference.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def run_with(self, error):
        import contextlib
        import io
        cli = self.load()
        out, err = io.StringIO(), io.StringIO()
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(cli.sigma_auth, "api", side_effect=error), \
                contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = cli.main([URL_ID, "--out", str(Path(tmp) / "look.json")])
            self.assertFalse((Path(tmp) / "look.json").exists())
        return code, out.getvalue() + err.getvalue()

    def test_an_http_error_is_one_line_and_exit_1(self):
        import urllib.error
        code, text = self.run_with(urllib.error.HTTPError("https://x/v2/workbooks", 400, "Bad Request", {}, None))
        self.assertEqual(code, 1)
        self.assertIn("400", text)
        self.assertIn(URL_ID, text)
        self.assertEqual(len(text.strip().splitlines()), 1)
        self.assertNotIn("Traceback", text)

    def test_an_auth_error_is_one_line_and_exit_1_without_a_token(self):
        code, text = self.run_with(self.load().sigma_auth.SigmaAuthError("no credentials: set SIGMA_CLIENT_ID"))
        self.assertEqual(code, 1)
        self.assertIn("no credentials", text)
        self.assertEqual(len(text.strip().splitlines()), 1)
        self.assertNotIn("Traceback", text)


# --------------------------------------------------------------------------- applying a look to 04
def report_ids(spec):
    return [i for node in layout_tree(spec) for i in [node[0]] + [c[0] for c in node[5]]]


def report_text(spec):
    """JSON of every Report-page element (the hidden Model / Inputs pages are not presentation)."""
    ids = set(report_ids(spec))
    return json.dumps([e for e in spec["contents"]["elements"] if e["id"] in ids], sort_keys=True)


def arrangement(spec):
    """The layout minus row numbers: element order, column spans and containers (roomy looks lengthen tables)."""
    return [[(n[0], n[1], n[2]) for n in layout_tree(spec, page)] for page in ("report", "inputs", "model")]


def span_lint(spec):
    return [(e["id"], bad) for e in spec["contents"]["elements"] if e["kind"] == "text"
            for bad in [ps.disallowed_span_styles(e["body"])] if bad]


class TestApplyLook(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base_plan, _ = draft(PIPELINE)
        cls.base = wb_spec(cls.base_plan)
        cls.looks = {"a": with_source(SPEC_A), "b": with_source(SPEC_B), "c": with_source(SPEC_C)}
        cls.plans, cls.specs = {}, {}
        for key, look in cls.looks.items():
            cls.plans[key], _ = draft(PIPELINE, look=look)
            cls.specs[key] = wb_spec(cls.plans[key])

    def test_the_plan_records_the_look_and_without_one_nothing_changes(self):
        for key, look in self.looks.items():
            self.assertEqual(self.plans[key]["presentation"]["look"], look)
        self.assertNotIn("look", self.base_plan["presentation"])
        self.assertIn("#0F172A", report_text(self.base))                  # the default look is unchanged
        self.assertIn("#F0F7FF", report_text(self.base))

    def test_none_of_our_colours_leaks_onto_a_reference_look(self):
        for key, spec in self.specs.items():
            text = report_text(spec).lower()
            supplied = json.dumps(self.looks[key]).lower()
            for colour in OURS:
                if colour.lower() not in supplied:
                    self.assertNotIn(colour.lower(), text, (key, colour))

    def test_the_workbook_theme_is_the_references_and_nothing_is_layered_on_it(self):
        for key, spec in self.specs.items():
            self.assertEqual(spec["contents"]["settings"], {"theme": self.looks[key]["theme"]}, key)
        self.assertEqual(self.specs["a"]["contents"]["settings"]["theme"], {"name": NAMED})   # by name, not inlined
        text = json.dumps(self.specs["c"]["contents"]["settings"])
        self.assertNotIn(ps.categorical_scheme(ps.ROLES["primary"])[1], text)     # no scheme of ours

    def test_no_chart_carries_a_per_chart_colour_override(self):
        for key, spec in self.specs.items():
            for el in spec["contents"]["elements"]:
                if el["kind"].endswith("-chart") and el["kind"] != "kpi-chart":
                    self.assertNotIn("color", {k: v for k, v in el.items() if k == "color" and "by" in (v or {})}, key)

    def test_kpi_cards_take_the_references_kpi_type(self):
        total = lambda key: element(self.specs[key], "kpi_total_weighted")           # noqa: E731
        self.assertEqual(total("c")["name"], {"text": "TOTAL WEIGHTED", "color": "#121b39", "fontSize": 13,
                                              "fontWeight": "bold"})
        self.assertEqual(total("c")["value"], {"columnId": "kpi_total_weighted_v", "fontSize": 30})
        self.assertEqual(total("c")["style"], {"borderRadius": "pill"})
        for key in ("a", "b"):
            self.assertEqual(total(key)["name"], "TOTAL WEIGHTED")                   # theme default
            self.assertEqual(total(key)["value"], {"columnId": "kpi_total_weighted_v"})
            self.assertNotIn("style", total(key))
        self.assertEqual(self.specs["b"]["contents"]["elements"][0].get("kind") is not None, True)

    def test_the_text_card_kpi_mirrors_the_reference_kpi_type(self):
        card, text = element(self.specs["c"], "kpi_coverage"), element(self.specs["c"], "kpi_coverage_t")
        self.assertEqual(card["style"], {"borderRadius": "pill"})
        label, value = text["body"].split("\n\n")
        self.assertEqual(label, '<span style="font-size: 13px; color: #121b39">**COVERAGE**</span>')
        self.assertTrue(value.startswith('<span style="font-size: 30px">**{{If('))
        self.assertTrue(value.endswith("| .1%}}**</span>"))
        _, plain = element(self.specs["a"], "kpi_coverage"), element(self.specs["a"], "kpi_coverage_t")
        self.assertEqual(plain["body"].split("\n\n")[0], '<span style="font-size: 18px">**COVERAGE**</span>')
        self.assertTrue(plain["body"].split("\n\n")[1].startswith('<span style="font-size: 40px">**{{If('))

    def test_the_header_band_and_text_follow_the_reference(self):
        band_b = element(self.specs["b"], "trHeaderBand")
        self.assertEqual(band_b["style"], CARD_B)                                    # the reference's container
        self.assertTrue(element(self.specs["b"], "trHeader")["body"].startswith("## "))
        self.assertNotIn("style", element(self.specs["a"], "trHeaderBand"))
        self.assertNotIn("style", element(self.specs["c"], "trHeaderBand"))
        body_c = element(self.specs["c"], "trHeader")["body"]
        title, subline = body_c.split("\n\n")
        self.assertEqual(title, '<span style="font-size: 30px; color: #121b39">H2 FY26 Weighted Forecast</span>')
        self.assertTrue(subline.startswith('<span style="font-size: 13px; color: #667085">'))
        self.assertEqual(element(self.specs["a"], "trHeader")["body"].split("\n")[0], "# H2 FY26 Weighted Forecast")
        label = element(self.specs["c"], "trDetail")["body"]
        self.assertIn("color: #667085", label)                                        # the reference's muted colour
        self.assertNotIn("color", element(self.specs["a"], "trDetail")["body"])

    def test_card_styles_per_kind_and_the_table_preset(self):
        for key, kind in (("c", "bar-chart"),):
            charts = [e for e in self.specs[key]["contents"]["elements"] if e["kind"] == kind]
            self.assertTrue(charts and all(c["style"] == {"padding": "none"} for c in charts))
        for el in self.specs["a"]["contents"]["elements"]:
            if el["id"].startswith(("f_", "ch_")) and el["kind"] != "text":
                self.assertNotIn("style", el, el["id"])
        summary = element(self.specs["c"], "f_s_h2_fy26_weighted_forecast")
        self.assertEqual(summary["tableStyle"]["preset"], "presentation")
        self.assertNotIn("style", summary)                  # tables had no card style in that reference
        self.assertNotIn("preset", element(self.specs["a"], "f_s_h2_fy26_weighted_forecast").get("tableStyle", {}))

    def test_under_a_look_the_theme_styles_the_headers_and_only_bold_survives(self):
        for key, spec in self.specs.items():
            for eid in ("f_s_h2_fy26_weighted_forecast", "f_s_open_deals_by_stage"):
                header = element(spec, eid)["tableStyle"]["textStyles"]["header"]
                self.assertEqual(header, {"fontWeight": "bold"}, (key, eid))        # no Excel fill or colour
            self.assertNotIn("#DDE7F5", report_text(spec).upper(), key)
        # without a look the Excel header fill is exactly as before
        for eid in ("f_s_h2_fy26_weighted_forecast", "f_s_open_deals_by_stage"):
            self.assertEqual(element(self.base, eid)["tableStyle"]["textStyles"]["header"],
                             {"backgroundColor": "#DDE7F5", "fontWeight": "bold"})

    def test_a_pivot_under_a_look_keeps_only_bold_too(self):
        style = ps.table_style({"fill": "#DDE7F5", "bold": True, "color": "#1F3864"}, "pivot-table",
                               self.looks["c"])
        self.assertEqual(style, {"preset": "presentation", "textStyles": {"header": {"fontWeight": "bold"},
                                                                          "columnHeader": {"fontWeight": "bold"}}})
        plain = ps.table_style({"fill": "#DDE7F5", "color": "#1F3864"}, "table", self.looks["a"])
        self.assertIsNone(plain)                          # not bold, no preset: nothing of Excel's, nothing to say

    def test_the_preset_is_behind_a_toggle(self):
        self.addCleanup(ps.SURFACES.update, dict(ps.SURFACES))
        ps.SURFACES["look_table_preset"] = False
        spec = wb_spec(draft(PIPELINE, look=self.looks["c"])[0])
        self.assertNotIn("preset", element(spec, "f_s_h2_fy26_weighted_forecast")["tableStyle"])

    def test_values_formats_formulas_and_layout_are_identical_to_the_default_look(self):
        presentation_only = {"style", "name", "value", "body", "tableStyle", "color", "legend"}
        base = {e["id"]: e for e in self.base["contents"]["elements"]}
        for key, spec in self.specs.items():
            self.assertEqual(arrangement(spec), arrangement(self.base), key)
            now = {e["id"]: e for e in spec["contents"]["elements"]}
            self.assertEqual(set(now), set(base), key)
            for eid, el in now.items():
                for field in set(el) | set(base[eid]):
                    if field in presentation_only:
                        continue
                    if field == "columns" and el["kind"] == "container":
                        continue
                    self.assertEqual(el.get(field), base[eid].get(field), (key, eid, field))
            self.assertEqual(span_lint(spec), [], key)
        # the number formats on every column survive
        for eid, el in base.items():
            for column, other in zip(el.get("columns", []), {e["id"]: e for e in self.specs["c"]["contents"][
                    "elements"]}[eid].get("columns", [])):
                self.assertEqual(column.get("format"), other.get("format"))

    def test_the_default_look_still_passes_the_span_lint(self):
        self.assertEqual(span_lint(self.base), [])

    def test_an_old_plan_and_a_bad_look_are_handled(self):
        plan = copy.deepcopy(self.plans["c"])
        del plan["presentation"]["look"]
        self.assertEqual(wb_spec(plan)["contents"]["layout"], self.base["contents"]["layout"])
        plan["presentation"]["look"] = {"version": 9}
        with self.assertRaises(trm.PlanError):
            trm.validate_plan(plan)

    def test_every_fixture_builds_with_a_look_and_keeps_its_layout(self):
        for number, path in SOURCES.items():
            plain, _ = draft(path)
            styled, _ = draft(path, look=self.looks["c"])
            self.assertEqual(arrangement(wb_spec(styled)), arrangement(wb_spec(plain)), number)
            self.assertEqual(span_lint(wb_spec(styled)), [], number)


if __name__ == "__main__":
    unittest.main()
