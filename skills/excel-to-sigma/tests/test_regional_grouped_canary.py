"""Focused source and emitted-graph oracle for the Regional01 B4 canary."""
from datetime import datetime
from decimal import Decimal
import json
import os
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.regional_grouped_canary import (build_candidate, source_rows, reference_b4,
    bind_assigned_formula)
from native_graph_reference import NativeGraphReference


ROOT = Path(os.environ.get("REGIONAL_PACKET_ROOT", Path(__file__).resolve().parents[5]))
WORK = ROOT / "work"
SOURCE = WORK / "broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx"
INPUT = WORK / "complete-source-record-bindings-plans-20261020/01-regional-sales-forecast-xlsx.json"
NATIVE = WORK / "complete-native-exact-replay-r1-20261020/01-regional-sales-forecast-xlsx.json"
ORACLES = WORK / "six-excel-authoritative-oracles-20261020/excel-oracles.jsonl"


def evaluate_graph(candidate, sales, summary_region, *, injected_errors=None):
    """Execute every emitted row, group and output column in dependency order.

    Group Sum/Min use the operation declared in the graph. DatePart quarter is
    evaluated here because the shared scalar test interpreter lacks quarters.
    This is an offline formula interpreter, never a claim about Sigma runtime.
    """
    sales = [dict(row, **{"Row ID": f"test-only-sales-{i}"}) for i, row in enumerate(sales)]
    for row in sales:
        if isinstance(row["Month"], str):
            row["Month"] = datetime.fromisoformat(row["Month"])
    graph = NativeGraphReference({"sales": ("Sales", sales),
                                  "summary": ("Summary Regions", [{"Region": summary_region,
                                    "Source Key": "Summary!row:4", "Row ID": "test-only-summary-0"}])})
    row_element, grouped_element, output_element = candidate["elements"]
    graph.tables["Rows"] = [{"origin": ("sales", i), "values": dict(row)}
                            for i, row in enumerate(sales)]
    graph.elements["rows"] = "Rows"
    for column in row_element["columns"]:
        for i, row in enumerate(graph.tables["Rows"]):
            formula = column["formula"]
            if formula == 'DatePart("year", [Sales/Month])':
                value = row["values"]["Month"].year if row["values"]["Month"] else None
            elif formula == 'DatePart("quarter", [Sales/Month])':
                value = (row["values"]["Month"].month - 1) // 3 + 1 if row["values"]["Month"] else None
            else:
                value = graph.evaluate(formula, "Rows", i)
            row["values"][column["name"]] = value
    for ordinal, error in (injected_errors or {}).items():
        row = graph.tables["Rows"][ordinal - 1]["values"]
        row["Revenue Error"] = error
        row["First Error Ordinal"] = ordinal
    keys = grouped_element["keys"]
    assert keys == ["Group Key"]
    grouped = {}
    for row in graph.tables["Rows"]:
        value = row["values"]
        group = grouped.setdefault(value["Group Key"], {"Group Key": value["Group Key"], "rows": []})
        group["rows"].append(value)
    for group in grouped.values():
        for calc in grouped_element["calculations"]:
            values = [r[calc["column"]] for r in group["rows"] if r[calc["column"]] is not None]
            group[calc["name"]] = (sum(values) if calc["aggregate"] == "Sum" else min(values)) if values else None
    graph.tables["Grouped"] = [{"origin": ("group", i), "values": {k: v for k, v in group.items() if k != "rows"}}
                               for i, group in enumerate(grouped.values())]
    graph.elements["grouped"] = "Grouped"
    graph.tables["Summary"] = [{"origin": ("summary", 0), "values": {"Region": summary_region,
        "Source Key": "Summary!row:4", "Row ID": "test-only-summary-0"}}]
    graph.elements["output"] = "Summary"
    result = graph.tables["Summary"][0]["values"]
    for column in output_element["columns"]:
        result[column["name"]] = graph.evaluate(column["formula"], "Summary", 0)
    return {lane: result[f"B4 {lane.title()}"] for lane in
            ("number", "text", "logical", "kind", "error")}


@unittest.skipUnless(SOURCE.exists() and INPUT.exists() and NATIVE.exists() and ORACLES.exists(),
                     "source-bound Regional replay packet is not installed")
class RegionalGroupedCanaryTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.inputs = json.loads(INPUT.read_text())
        cls.native = json.loads(NATIVE.read_text())
        cls.oracle = next(json.loads(line) for line in ORACLES.read_text().splitlines()
                          if '"sourceFile": "01 Regional Sales Forecast.xlsx"' in line
                          and '"sheet": "Summary", "address": "B4"' in line)
        cls.candidate = build_candidate(SOURCE, cls.inputs, cls.native, cls.oracle)

    def test_baseline_matches_independent_excel_and_decimal(self):
        rows = source_rows(self.inputs)
        result = evaluate_graph(self.candidate, rows, "West")
        self.assertEqual(Decimal(str(result["number"])), Decimal(self.oracle["numericValue"]))
        self.assertEqual(reference_b4(rows, "West"), Decimal(self.oracle["numericValue"]))
        self.assertEqual((result["kind"], result["error"]), ("n", ""))
        self.assertEqual(self.candidate["constructionBudget"]["tables"], 3)
        self.assertLessEqual(self.candidate["constructionBudget"]["formulaColumns"], 128)
        self.assertEqual(len(self.candidate["sourceMappings"]), 145)

    def test_edits_zero_null_no_match_and_first_error(self):
        baseline = source_rows(self.inputs)
        edited = [dict(row) for row in baseline]
        edited[0]["Units"] += 10
        edited[0]["Avg Price"] += 5
        edited[1]["Month"] = "2026-04-01"
        edited[2]["Region"] = "East"
        edited[3]["Units"] = 0
        edited[4]["Avg Price"] = None
        result = evaluate_graph(self.candidate, edited, "West")
        self.assertEqual(Decimal(str(result["number"])), reference_b4(edited, "West"))
        self.assertEqual(evaluate_graph(self.candidate, edited, "No such region")["number"], 0)
        errors = evaluate_graph(self.candidate, edited, "West", injected_errors={14: "#VALUE!", 13: "#DIV/0!"})
        self.assertEqual(errors["error"], "#DIV/0!")
        self.assertEqual(errors["kind"], "n")

    def test_fails_closed_on_wrong_b4_source_formula(self):
        changed = json.loads(json.dumps(self.native))
        changed["formulas"]["Summary!B4"]["originalFormula"] = "=0"
        with self.assertRaisesRegex(ValueError, "stale|changed"):
            build_candidate(SOURCE, self.inputs, changed, self.oracle)

    def test_assigned_binding_and_exact_patch_contract(self):
        with self.assertRaisesRegex(ValueError, "unassigned"):
            bind_assigned_formula("[Unknown/X]", namespaces={}, local_columns={})


if __name__ == "__main__":
    unittest.main()
