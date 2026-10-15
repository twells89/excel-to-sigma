#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import hashlib
import io
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "lib"))


def load_script(filename: str, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, SCRIPTS / filename)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


FORECAST = load_script("build-forecast-model.py", "forecast_builder")
RESEARCH = load_script("build-research-model.py", "research_builder")


def forecast_plan():
    return {
        "name": "Forecast fixture",
        "intent": "read-only",
        "connectionId": "connection",
        "folderId": "folder",
        "currencyFormat": "$,.0f",
        "pctFormat": ".1%",
        "actualMonths": [{"month": "Jan-26", "num": 1, "end": "2026-01-31"}],
        "forecastMonths": [{"month": "Feb-26", "num": 2, "end": "2026-02-28"}],
        "lines": [{
            "line": "Revenue", "section": "Revenue", "secOrder": 1,
            "lineOrder": 1, "kind": "growth", "rate": 0.1, "juneBase": 100.0,
        }],
        "assumptions": [["Revenue growth", 0.1, "monthly"]],
        "actuals": [{
            "date": "2026-01-10", "month": "Jan-26", "line": "Revenue",
            "section": "Revenue", "secOrder": 1, "lineOrder": 1,
            "dept": "Sales", "vendor": "Customer", "amt": 100.0,
        }],
    }


def research_plan():
    named = [
        (8, "Revenue", "input", {"2024": 100.0}, None),
        (9, "% growth", "input", {"2024": 0.125}, None),
        (13, "Gross profit", "derived", {}, "[c8]"),
        (14, "% margin", "ratio", {}, "[c13]/[c8]"),
        (18, "EBITDA", "derived", {}, "[c13]"),
        (19, "% margin", "ratio", {}, "[c18]/[c8]"),
        (22, "EBIT", "derived", {}, "[c18]"),
        (23, "% margin", "ratio", {}, "[c22]/[c8]"),
        (31, "Net income", "derived", {}, "[c22]"),
        (32, "% growth", "input", {"2024": 0.25}, None),
    ]
    named.extend(
        (100 + i, f"Unique line {i:02d}", "input", {"2024": float(i)}, None)
        for i in range(1, 26)
    )
    lines = []
    for order, (row, label, kind, typed, canonical) in enumerate(named, 1):
        lines.append({
            "row": row,
            "col_id": f"c{row}",
            "label": label,
            "section": "PROFIT AND LOSS",
            "sec_order": 1,
            "line_order": order,
            "kind": kind,
            "typed": typed,
            "canonical": canonical,
        })
    return {"periods": [{"year": 2024}], "lines": lines}


def all_dicts(value):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from all_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from all_dicts(child)


class TestWorkbookBuilders(unittest.TestCase):
    def build_forecast_workbook(self):
        return FORECAST.build_wb_spec(
            forecast_plan(), "dm", {"glDetail": "gl", "spine": "sp", "assumptions": "as"})

    def test_forecast_builder_emits_current_pivot_axes_without_forbidden_visibility(self):
        workbook = self.build_forecast_workbook()

        pivot = next(node for node in all_dicts(workbook) if node.get("kind") == "pivot-table")
        self.assertEqual(
            pivot["rowsBy"],
            [
                {"columnId": "p_sec", "sort": {"direction": "ascending", "by": "p_so", "aggregation": "min"}},
                {"columnId": "p_line", "sort": {"direction": "ascending", "by": "p_lo", "aggregation": "min"}},
            ],
        )
        self.assertEqual(
            pivot["columnsBy"],
            [
                {"columnId": "p_per", "sort": {"direction": "ascending"}},
                {"columnId": "p_m", "sort": {"direction": "ascending"}},
            ],
        )
        self.assertFalse(any("visibleAsSource" in node for node in all_dicts(workbook)))

    def test_forecast_builder_adds_monthly_profit_trend_with_chronological_axis(self):
        workbook = self.build_forecast_workbook()

        charts = [node for node in all_dicts(workbook) if node.get("kind") == "line-chart"]
        self.assertEqual(len(charts), 1)
        chart = charts[0]
        self.assertEqual(chart["id"], "monthlyTrend")
        self.assertEqual(
            chart["columns"],
            [
                {"id": "t_month", "name": "Month", "formula": 'DateTrunc("month", [PnL Combined/Month End])', "format": FORECAST.DATEF},
                {"id": "t_revenue", "name": "Revenue", "formula": "Sum([PnL Combined/Revenue Amt])", "format": FORECAST.CUR(forecast_plan())},
                {"id": "t_ebitda", "name": "EBITDA", "formula": "Sum([PnL Combined/Revenue Amt]) - Sum([PnL Combined/COGS Amt]) - Sum([PnL Combined/OpEx Amt])", "format": FORECAST.CUR(forecast_plan())},
                {"id": "t_net_income", "name": "Net Income", "formula": "Sum([PnL Combined/Revenue Amt]) - Sum([PnL Combined/COGS Amt]) - Sum([PnL Combined/OpEx Amt]) - Sum([PnL Combined/Below Amt])", "format": FORECAST.CUR(forecast_plan())},
            ],
        )
        self.assertEqual(
            chart["xAxis"],
            {"columnId": "t_month", "sort": {"direction": "ascending"}},
        )
        self.assertEqual(chart["yAxis"]["columnIds"], ["t_revenue", "t_ebitda", "t_net_income"])

    def test_forecast_pivot_hides_only_the_invalid_cross_section_total_row(self):
        workbook = self.build_forecast_workbook()

        pivot = next(node for node in all_dicts(workbook) if node.get("id") == "pnlPivot")
        self.assertEqual(
            pivot["axisTotals"],
            [{"axisId": "grand-row-total", "visibility": "hidden"}],
        )
        self.assertNotIn("totals", pivot)

    def test_forecast_builder_uses_designed_pages_with_unique_nonoverlapping_layout(self):
        workbook = self.build_forecast_workbook()

        self.assertEqual([page["id"] for page in workbook["pages"]], ["pnl", "detail", "inputs", "model"])
        self.assertEqual(workbook["pages"][-1]["visibility"], "hidden")
        page_ids = {
            page["id"]: [element["id"] for element in page.get("elements", [])]
            for page in workbook["pages"]
        }
        self.assertEqual(
            page_ids["pnl"],
            ["pnlHeader", "kpiRev", "kpiEbitda", "kpiNI", "kpiGM", "monthlyTrend", "pnlPivot"],
        )
        self.assertEqual(page_ids["detail"], ["glTable"])
        self.assertEqual(page_ids["inputs"], ["assumpTable"])
        self.assertEqual(page_ids["model"], ["actuals", "spineWrap", "forecast", "pnlCombined"])
        all_element_ids = [element_id for ids in page_ids.values() for element_id in ids]
        self.assertEqual(len(all_element_ids), len(set(all_element_ids)))

        placements = {}
        for page_match in re.finditer(r'<Page\b[^>]*\bid="([^"]+)"[^>]*>(.*?)</Page>', workbook["layout"], re.S):
            page_id, body = page_match.groups()
            spans = []
            for element_id, col_start, col_end, row_start, row_end in re.findall(
                r'<Element elementId="([^"]+)" gridColumn="(\d+) / (\d+)" gridRow="(\d+) / (\d+)"/>', body
            ):
                placement = tuple(map(int, (col_start, col_end, row_start, row_end)))
                placements[element_id] = (page_id, placement)
                spans.append((element_id, placement))
            for index, (left_id, left) in enumerate(spans):
                for right_id, right in spans[index + 1:]:
                    overlaps = left[0] < right[1] and right[0] < left[1] and left[2] < right[3] and right[2] < left[3]
                    self.assertFalse(overlaps, f"{page_id}: {left_id} overlaps {right_id}")

        self.assertEqual(set(placements), set(all_element_ids))
        self.assertEqual(placements["pnlHeader"], ("pnl", (1, 25, 1, 4)))
        self.assertEqual(placements["kpiRev"], ("pnl", (1, 7, 4, 10)))
        self.assertEqual(placements["kpiGM"], ("pnl", (19, 25, 4, 10)))
        self.assertEqual(placements["monthlyTrend"], ("pnl", (1, 25, 10, 25)))
        self.assertEqual(placements["pnlPivot"], ("pnl", (1, 25, 25, 55)))

    def test_research_builder_keeps_all_35_display_rows_distinct_and_nonadditive(self):
        plan = research_plan()
        dm_spec, data_lines = RESEARCH.build_dm(plan, "connection", "folder", "Research")
        workbook = RESEARCH.build_wb(
            plan, "dm", {"data": "data", "linedim": "linedim"}, data_lines,
            "folder", "Research")

        linedim = next(
            element for element in dm_spec["pages"][0]["elements"]
            if any(column["name"] == "Label" for column in element["columns"])
        )
        labels = re.findall(
            r"\('c\d+','([^']*)','PROFIT AND LOSS',",
            linedim["source"]["statement"],
        )
        self.assertEqual(len(labels), 35)
        self.assertEqual(len(set(labels)), 35)
        self.assertEqual(
            set(labels) & {"Revenue growth", "Gross profit margin", "EBITDA margin", "EBIT margin", "Net income growth"},
            {"Revenue growth", "Gross profit margin", "EBITDA margin", "EBIT margin", "Net income growth"},
        )

        pivot = next(node for node in all_dicts(workbook) if node.get("kind") == "pivot-table")
        self.assertEqual(
            pivot["rowsBy"],
            [
                {"columnId": "pv_sec", "sort": {"direction": "ascending", "by": "pv_so", "aggregation": "min"}},
                {"columnId": "pv_line", "sort": {"direction": "ascending", "by": "pv_lo", "aggregation": "min"}},
            ],
        )
        self.assertEqual(
            pivot["columnsBy"],
            [{"columnId": "pv_yr", "sort": {"direction": "ascending"}}],
        )
        self.assertEqual(
            pivot["axisTotals"],
            [
                {"axisId": "pv_sec", "visibility": "hidden"},
                {"axisId": "pv_yr", "visibility": "hidden"},
                {"axisId": "grand-row-total", "visibility": "hidden"},
                {"axisId": "grand-column-total", "visibility": "hidden"},
            ],
        )
        self.assertNotIn("totals", pivot)
        self.assertFalse(any("visibleAsSource" in node for node in all_dicts(workbook)))

    def test_research_ratio_inputs_remain_decimal_values(self):
        plan = research_plan()
        data_sql = RESEARCH.data_sql(plan, [plan["lines"][1]])

        self.assertIn("0.125000", data_sql)
        self.assertNotIn("12.500000", data_sql)

    def test_research_builder_uses_period_formulas_and_only_literal_overrides(self):
        plan = research_plan()
        line = plan["lines"][2]
        line["typed"] = {"2024": 999.1234567890123}
        line["literalInputs"] = {}
        line["periodFormulas"] = {"2024": "[c8] * 2"}
        dm_spec, data_lines = RESEARCH.build_dm(plan, "connection", "folder", "Research")
        workbook = RESEARCH.build_wb(
            plan, "dm", {"data": "data", "linedim": "linedim"}, data_lines,
            "folder", "Research")
        calc = next(node for node in all_dicts(workbook) if node.get("id") == "calc")
        formula = next(column["formula"] for column in calc["columns"] if column["id"] == "c13")

        self.assertEqual(formula, 'If([Year] = 2024, [c8] * 2, Null)')
        self.assertNotIn("999.1234567890123", dm_spec["pages"][0]["elements"][0]["source"]["statement"])

    def test_research_builder_coalesces_real_literal_years_with_period_formulas(self):
        plan = research_plan()
        plan["periods"] = [{"year": 2024}, {"year": 2025}]
        line = plan["lines"][2]
        line["literalInputs"] = {"2024": 602}
        line["periodFormulas"] = {"2025": "[c8] + [c11]"}
        _dm_spec, data_lines = RESEARCH.build_dm(plan, "connection", "folder", "Research")
        workbook = RESEARCH.build_wb(
            plan, "dm", {"data": "data", "linedim": "linedim"}, data_lines,
            "folder", "Research")
        calc = next(node for node in all_dicts(workbook) if node.get("id") == "calc")

        self.assertEqual(
            next(column["formula"] for column in calc["columns"] if column["id"] == "c13"),
            "Coalesce([Custom SQL/c13], If([Year] = 2025, [c8] + [c11], Null))",
        )

    def test_research_data_sql_preserves_source_precision(self):
        plan = research_plan()
        plan["lines"][0]["typed"] = {"2024": 1.23456789012345}

        sql = RESEARCH.data_sql(plan, [plan["lines"][0]])

        self.assertIn("1.23456789012345", sql)
        self.assertNotIn("1.234568", sql)

    def test_research_builder_exposes_source_bound_cross_sheet_input_defaults(self):
        plan = research_plan()
        plan["sourceInputs"] = [{
            "id": "i_model_inputs_c2", "sourceCellId": "Model Inputs!C2",
            "defaultValue": 0.05, "values": {"2024": 0.05}, "mappingStatus": "live",
        }]
        plan["lines"][2]["periodFormulas"] = {
            "2024": "[c8] * (1 + [i_model_inputs_c2])",
        }
        dm_spec, data_lines = RESEARCH.build_dm(plan, "connection", "folder", "Research")
        workbook = RESEARCH.build_wb(
            plan, "dm", {"data": "data", "linedim": "linedim"}, data_lines,
            "folder", "Research")
        sql = dm_spec["pages"][0]["elements"][0]["source"]["statement"]
        calc = next(node for node in all_dicts(workbook) if node.get("id") == "calc")

        self.assertIn("i_model_inputs_c2", sql)
        self.assertIn("0.050000", sql)
        self.assertEqual(
            next(column["formula"] for column in calc["columns"]
                 if column["id"] == "i_model_inputs_c2"),
            "[Custom SQL/i_model_inputs_c2]",
        )

    def test_forecast_cli_rejects_a_plan_without_a_resolvable_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            Workbook().save(source)
            plan = forecast_plan()
            plan["sourceSha256"] = hashlib.sha256(source.read_bytes()).hexdigest()
            plan_path = Path(tmp) / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")

            with mock.patch.object(sys, "argv", ["build-forecast-model.py", str(plan_path)]):
                exit_code = FORECAST.main()

        self.assertEqual(exit_code, 1)

    def test_forecast_cli_reports_a_successful_source_gate_in_dry_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            Workbook().save(source)
            plan = forecast_plan()
            plan.update({
                "sourcePath": str(source),
                "sourceSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            })
            plan_path = Path(tmp) / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")

            output = io.StringIO()
            with mock.patch("sys.stdout", output):
                exit_code = FORECAST.main([str(plan_path)])

        self.assertEqual(exit_code, 0)
        self.assertIn("MIGRATION GATE: AUTO", output.getvalue())

    def test_research_cli_rejects_a_plan_without_a_resolvable_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            Workbook().save(source)
            plan = research_plan()
            plan.update({
                "source": "source.xlsx",
                "sourceSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                "summary": {"input": 27, "derived": 4, "ratio": 4, "needs_review": 0},
                "parity": {"live": 35, "frozen": 0},
                "anchors": {},
            })
            plan_path = Path(tmp) / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")

            with mock.patch.object(sys, "argv", [
                "build-research-model.py", str(plan_path), "--conn", "conn", "--folder", "folder",
            ]):
                exit_code = RESEARCH.main()

        self.assertEqual(exit_code, 1)

    def test_research_cli_accepts_spaced_options_and_reports_source_gate(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            Workbook().save(source)
            plan = research_plan()
            plan.update({
                "source": "source.xlsx",
                "sourcePath": str(source),
                "sourceSha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                "summary": {"input": 27, "derived": 4, "ratio": 4, "needs_review": 0},
                "parity": {"live": 35, "frozen": 0},
                "anchors": {},
            })
            plan_path = Path(tmp) / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")

            output = io.StringIO()
            with mock.patch("sys.stdout", output):
                exit_code = RESEARCH.main([
                    str(plan_path), "--conn", "connection", "--folder", "folder",
                ])

        self.assertEqual(exit_code, 0)
        self.assertIn("MIGRATION GATE: AUTO", output.getvalue())


if __name__ == "__main__":
    unittest.main()
