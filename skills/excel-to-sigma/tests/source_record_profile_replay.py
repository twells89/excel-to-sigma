"""Offline source-specific reviewed input profiles; no target calls or writes.

Run with --originals, --coverage and --out. The builder in source_inputs remains
workbook-agnostic; this file holds the source-specific review decisions.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl.utils.cell import get_column_letter

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts" / "lib"))
from source_inputs import build_source_input_plan, validate_source_input_plan  # noqa: E402


def table(fid, sheet, name, keys, types, *, role="input"):
    return {"featureId": fid, "sheet": sheet, "table": name, "role": role,
            "reviewed": True, "keyColumns": keys, "literalColumns": list(types),
            "columnTypes": types}


def block(fid, sheet, bounds, headers, keys, types):
    return {"featureId": fid, "sheet": sheet, "range": bounds,
            "headers": headers, "role": "input", "reviewed": True,
            "keyColumns": keys, "literalColumns": list(types), "columnTypes": types}


def records(fid, keys, types, bindings, *, role="input", sha=None, basis=None, identity=False):
    result = {"featureId": fid, "role": role, "reviewed": True,
              "keyColumns": keys, "literalColumns": list(types), "columnTypes": types,
              "recordBindings": [{"sourceCells": cells, **({"sourceRecordId": record_id} if record_id else {})}
                                 for cells, record_id in bindings]}
    if role == "reference":
        result["immutableReview"] = {"reviewed": True, "sourceSha256": sha, "basis": basis}
    if identity:
        result["recordIdentity"] = {"kind": "frozen-source-record", "column": keys[0],
                                    "reviewed": True, "sourceSha256": sha, "reviewBasis": basis}
    return result


def cells(fid, sheet, addresses, field, kind, *, sha, basis):
    return records(fid, [field], {field: kind},
                   [({field: f"{sheet}!{address}"}, None) for address in addresses],
                   role="reference", sha=sha, basis=basis)


def wide(fid, sheet, mapping, types, key, *, role="input", sha=None, basis=None):
    return records(fid, [key], types,
                   [({field: f"{sheet}!{address}" for field, address in mapping.items()}, None)],
                   role=role, sha=sha, basis=basis)


def source_profiles(sha):
    """Explicit source shapes; every address is checked against current source bytes."""
    m = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return {
        "01 Regional Sales Forecast.xlsx": [
            table("sales-forecast", "Forecast Input", "tblSalesForecast",
                  ["Region", "Product Line", "Month"],
                  {"Region": "text", "Product Line": "text", "Month": "datetime",
                   "Units": "number", "Avg Price": "number"}),
            table("regions", "Regions", "tblRegions", ["Region"],
                  {"Region": "text", "Region Manager": "text", "FY26 Target": "number"}, role="reference"),
            cells("summary-regions", "Summary", [f"A{r}" for r in range(4, 8)], "Region", "text",
                  sha=sha["01 Regional Sales Forecast.xlsx"], basis="Source row labels identify regions"),
            cells("summary-months", "Summary", [f"A{r}" for r in range(12, 24)], "Month", "datetime",
                  sha=sha["01 Regional Sales Forecast.xlsx"], basis="Source rows define the fixed FY26 month axis"),
            cells("summary-products", "Summary", [f"{c}11" for c in "BCD"], "Product Line", "text",
                  sha=sha["01 Regional Sales Forecast.xlsx"], basis="Source header labels identify product lines"),
        ],
        "02 SaaS Driver Forecast FY26.xlsx": [
            block("actuals", "Data", "A2:F157",
                  ["Date", "Month", "Line Item", "Department", "Vendor", "Amount"],
                  ["Date", "Line Item", "Department", "Vendor"],
                  {"Date": "datetime", "Month": "text", "Line Item": "text",
                   "Department": "text", "Vendor": "text", "Amount": "number"}),
            block("drivers", "Assumptions", "A5:C10", ["Driver", "Monthly Rate", "Basis"],
                  ["Driver"], {"Driver": "text", "Monthly Rate": "number", "Basis": "text"}),
            wide("data-headers", "Data", {"Month": "B1", "Line Item": "C1", "Amount": "F1"},
                 {"Month": "text", "Line Item": "text", "Amount": "text"}, "Month",
                 role="reference", sha=sha["02 SaaS Driver Forecast FY26.xlsx"],
                 basis="Source header cells define the historical table columns"),
            cells("pnl-line-labels", "P&L", [f"A{r}" for r in (7, 12, 13, 19, 20, 21, 22, 27)],
                  "Line Item", "text", sha=sha["02 SaaS Driver Forecast FY26.xlsx"],
                  basis="Source P&L row labels identify computed business lines"),
            cells("pnl-period-labels", "P&L", [f"{c}4" for c in "BCDEFG"],
                  "Period", "text", sha=sha["02 SaaS Driver Forecast FY26.xlsx"],
                  basis="Source P&L headers define historical periods"),
            records("manual-forecast-lines", ["Line Item"],
                    {"Line Item": "text", **{f"{month}-26": "number" for month in m[6:]}},
                    [({"Line Item": f"P&L!A{r}",
                       **{f"{month}-26": f"P&L!{col}{r}" for month, col in zip(m[6:], "HIJKLM")}}, None)
                     for r in (8, 28)]),
        ],
        "03 Retail Quarterly Budget.xlsx": [
            block("gl-detail", "GL Detail", "A2:E125",
                  ["Period", "Store", "Account", "Account Group", "Amount"],
                  ["Period", "Store", "Account"],
                  {"Period": "text", "Store": "text", "Account": "text",
                   "Account Group": "text", "Amount": "number"}),
            block("drivers", "Drivers", "A4:C8", ["Driver", "Value", "Notes"],
                  ["Driver"], {"Driver": "text", "Value": "number"}),
            wide("gl-headers", "GL Detail", {"Period": "A1", "Account": "C1", "Amount": "E1"},
                 {"Period": "text", "Account": "text", "Amount": "text"}, "Period",
                 role="reference", sha=sha["03 Retail Quarterly Budget.xlsx"],
                 basis="Source GL header cells define transaction fields"),
            cells("budget-accounts", "Budget", [f"A{r}" for r in (7, 8, 12, 13, 18, 20, 21, 26)],
                  "Account", "text", sha=sha["03 Retail Quarterly Budget.xlsx"],
                  basis="Source row labels identify computed budget accounts"),
            cells("budget-historical-periods", "Budget", [f"{c}4" for c in "BCDE"],
                  "Period", "text", sha=sha["03 Retail Quarterly Budget.xlsx"],
                  basis="Source historical period headers are immutable labels"),
            wide("manual-marketing", "Budget",
                 {"Account": "A19", "Q3 FY26 Bud": "F19", "Q4 FY26 Bud": "G19"},
                 {"Account": "text", "Q3 FY26 Bud": "number", "Q4 FY26 Bud": "number"}, "Account"),
            wide("manual-interest-income", "Budget",
                 {"Account": "A25", "Q3 FY26 Bud": "F25", "Q4 FY26 Bud": "G25"},
                 {"Account": "text", "Q3 FY26 Bud": "number", "Q4 FY26 Bud": "number"}, "Account"),
            {"featureId": "corporate-allocations", "role": "replacement", "reviewed": True,
             "displayName": "Corporate Allocations",
             "replacementFor": "FY26 Corporate Allocations external workbook",
             "keyColumns": ["Period"], "literalColumns": ["Period", "Allocation"],
             "columnTypes": {"Period": "text", "Allocation": "number"},
             "suppliedRows": [{"Period": "FY26-Q3", "Allocation": 61000},
                              {"Period": "FY26-Q4", "Allocation": 64000}],
             "replacementBindings": [
                 {"key": {"Period": "FY26-Q3"}, "sourceCells": {"Period": "Budget!F4", "Allocation": "Budget!F21"},
                  "sourceLabel": "Q3 FY26 Bud", "originalFormula": "='[FY26 Corporate Allocations.xlsx]Alloc'!$C$4"},
                 {"key": {"Period": "FY26-Q4"}, "sourceCells": {"Period": "Budget!G4", "Allocation": "Budget!G21"},
                  "sourceLabel": "Q4 FY26 Bud", "originalFormula": "='[FY26 Corporate Allocations.xlsx]Alloc'!$D$4"}],
             "replacementReview": {"approved": True,
                                   "basis": "Previously reviewed persisted editable replacement of two external allocations"}},
        ],
        "04 Sales Pipeline Forecast.xlsx": [
            table("pipeline", "Pipeline", "tblPipeline", ["Opp ID"],
                  {"Opp ID": "text", "Account": "text", "Owner": "text", "Stage": "text",
                   "Close Date": "datetime", "Amount": "number"}),
            table("stages", "Stages", "tblStages", ["Stage"],
                  {"Stage": "text", "Probability": "number", "Forecast Category": "text"}, role="reference"),
            cells("forecast-stages", "Forecast", [f"A{r}" for r in range(12, 18)], "Stage", "text",
                  sha=sha["04 Sales Pipeline Forecast.xlsx"], basis="Source forecast stage labels"),
            records("owner-targets", ["Owner"], {"Owner": "text", "Target": "number"},
                    [({"Owner": f"Forecast!A{r}", "Target": f"Forecast!J{r}"}, None)
                     for r in range(4, 8)]),
        ],
        "05 Statistical Demand Forecast.xlsx": [
            {"featureId": "history-long", "role": "input", "reviewed": True,
             "keyColumns": ["Month", "SKU"], "literalColumns": ["Month", "SKU", "Units"],
             "columnTypes": {"Month": "datetime", "SKU": "text", "Units": "number"},
             "reshape": {"kind": "unpivot", "sheet": "History", "table": "tblHistory",
                         "keyColumns": ["Month"], "valueColumns": ["SKU-A100", "SKU-B200", "SKU-C300"],
                         "axisName": "SKU", "valueName": "Units"}},
            wide("forecast-parameters", "Params",
                 {"Title": "A1", "Selected SKU": "B3", "Seasonality Months": "B4",
                  "Service Level": "B5", "Lead Time Months": "B6"},
                 {"Title": "text", "Selected SKU": "text", "Seasonality Months": "number",
                  "Service Level": "number", "Lead Time Months": "number"}, "Title"),
            cells("forecast-months", "Forecast", [f"A{r}" for r in range(4, 16)], "Month", "datetime",
                  sha=sha["05 Statistical Demand Forecast.xlsx"],
                  basis="Source forecast horizon dates are immutable calendar axis"),
            cells("history-month-header", "History", ["A1"], "Header", "text",
                  sha=sha["05 Statistical Demand Forecast.xlsx"],
                  basis="Source history table Month header label"),
        ],
        "06 13-Week Cash Flow.xlsx": [
            table("receivables", "AR Aging", "tblAR", ["Invoice"],
                  {"Invoice": "text", "Customer": "text", "Amount": "number",
                   "Due Date": "datetime", "Expected Week": "number"}),
            records("payables", ["Source Record ID"],
                    {"Source Record ID": "text", "Vendor": "text", "Amount": "number",
                     "Pay Week": "number", "Category": "text"},
                    [({"Vendor": f"AP Schedule!A{r}", "Amount": f"AP Schedule!B{r}",
                       "Pay Week": f"AP Schedule!C{r}", "Category": f"AP Schedule!D{r}"},
                      f"AP Schedule!row:{r}") for r in range(2, 32)],
                    sha=sha["06 13-Week Cash Flow.xlsx"], identity=True,
                    basis="AP Vendor/Pay Week/Category repeats; each frozen physical source row is distinct and source drift requires review"),
            wide("cash-settings", "Inputs",
                 {"Title": "A1", "Scenario": "B3", "Opening Cash": "B4",
                  "Minimum Cash": "B5", "Week Ending": "B6", "New Sales Receipts": "B7",
                  "Growth": "B8", "Payroll": "B9"},
                 {"Title": "text", "Scenario": "text", "Opening Cash": "number",
                  "Minimum Cash": "number", "Week Ending": "datetime",
                  "New Sales Receipts": "number", "Growth": "number", "Payroll": "number"}, "Title"),
            block("collection-scenarios", "Inputs", "A13:B15", ["Scenario", "Collection Rate"],
                  ["Scenario"], {"Scenario": "text", "Collection Rate": "number"}),
            records("manual-cash", ["Week"], {"Week": "number", "Rent": "number", "Debt": "number"},
                    [({"Week": f"13WCF!{col}4", "Rent": f"13WCF!{col}16",
                       "Debt": f"13WCF!{col}18"}, None)
                     for col in "BCDEFGHIJKLMN"]),
        ],
        "07 Headcount and Opex Plan.xlsx": [
            table("roster", "Roster", "tblRoster", ["Emp ID"],
                  {"Emp ID": "text", "Name": "text", "Department": "text", "Role": "text",
                   "Start Date": "datetime", "End Date": "datetime", "Annual Salary": "number",
                   "Status": "text"}),
            {"featureId": "headcount-budget", "role": "input", "reviewed": True,
             "keyColumns": ["Department", "Month"], "literalColumns": ["Department", "Month", "Budget"],
             "columnTypes": {"Department": "text", "Month": "datetime", "Budget": "number"},
             "reshape": {"kind": "unpivot", "sheet": "Dept Budget", "table": "tblDeptBudget",
                         "keyColumns": ["Department"], "valueColumns": m,
                         "axisName": "Month", "valueName": "Budget",
                         "axisValues": {name: f"2026-{i+1:02d}-01" for i, name in enumerate(m)}}},
            wide("cost-assumptions", "Assumptions",
                 {"Title": "A1", "Benefits Load": "B3", "Merit Increase": "B4",
                  "Merit Effective Month": "B5"},
                 {"Title": "text", "Benefits Load": "number", "Merit Increase": "number",
                  "Merit Effective Month": "number"}, "Title"),
            cells("monthly-cost-row-axis", "Monthly Cost", [f"A{r}" for r in range(5, 45)],
                  "Source Row", "number", sha=sha["07 Headcount and Opex Plan.xlsx"],
                  basis="Source sequence numbers identify Monthly Cost rows"),
            wide("monthly-cost-period-axis", "Monthly Cost",
                 {"Label": "C4", **{f"2026-{i+1:02d}": f"{get_column_letter(i+4)}4"
                                       for i in range(12)}, "Total": "P4"},
                 {"Label": "text", **{f"2026-{i+1:02d}": "datetime" for i in range(12)},
                  "Total": "text"}, "Label", role="reference",
                 sha=sha["07 Headcount and Opex Plan.xlsx"],
                 basis="Source Monthly Cost header dates and total label are fixed axes"),
            cells("variance-departments", "Variance", [f"A{r}" for r in range(4, 9)],
                  "Department", "text", sha=sha["07 Headcount and Opex Plan.xlsx"],
                  basis="Source variance row labels identify departments"),
        ],
        "08 Inventory Replenishment Plan.xlsm": [
            table("sku", "SKUs", "tblSKU", ["SKU"],
                  {"SKU": "text", "Description": "text", "On Hand": "number",
                   "Lead Time Wks": "number", "Safety Stock": "number",
                   "Order Multiple": "number", "Unit Cost": "number"}),
            table("demand", "Demand", "tblDemand", ["SKU", "Week Start"],
                  {"SKU": "text", "Week Start": "datetime", "Forecast Units": "number",
                   "Planned Receipts": "number"}),
            table("snapshots", "Snapshots", "tblSnapshots", ["SKU"],
                  {"Committed At": "datetime", "SKU": "text", "Week Start": "datetime",
                   "Forecast Units": "number", "Committed By": "text"}),
            cells("projection-skus", "Projection", [f"A{r}" for r in range(4, 10)], "SKU", "text",
                  sha=sha["08 Inventory Replenishment Plan.xlsm"],
                  basis="Source projection row labels identify SKUs"),
            cells("projection-weeks", "Projection", [f"{c}3" for c in "BCDEFGHIJKLM"],
                  "Week Start", "datetime", sha=sha["08 Inventory Replenishment Plan.xlsm"],
                  basis="Source projection weekly calendar axis"),
        ],
        "fixture-expense-approval.xlsx": [
            table("expense-register", "Expense Register", "tblExpenseRegister", ["Expense ID"],
                  {"Expense ID": "text", "Date": "datetime", "Employee": "text",
                   "Department": "text", "Category": "text", "Vendor": "text",
                   "Amount": "number", "Receipt Status": "text", "Approval Status": "text",
                   "Card Amount": "number", "Ledger Amount": "number"}),
            table("expense-policy", "Policy", "tblExpensePolicy", ["Category"],
                  {"Category": "text", "Per-item Limit": "number", "Receipt Threshold": "number",
                   "Policy Owner": "text"}),
        ],
        "fixture-forecast-model.xlsx": [
            table("actuals-detail", "Actuals Detail", "tblFixtureActuals",
                  ["Date", "Line Item", "Department", "Counterparty"],
                  {"Date": "datetime", "Month": "text", "Line Item": "text", "Section": "text",
                   "Section Order": "number", "Line Order": "number", "Department": "text",
                   "Counterparty": "text", "Amount": "number"}),
            table("forecast-inputs", "Forecast Inputs", "tblFixtureForecastInputs", ["Line Item"],
                  {"Line Item": "text", "Method": "text", "Monthly Growth": "number",
                   **{f"{name}-26": "number" for name in m[6:]}}),
        ],
        "fixture-inventory-reorder.xlsx": [
            table("inventory", "Inventory", "tblInventory", ["SKU"],
                  {"SKU": "text", "Product": "text", "Category": "text", "Supplier": "text",
                   "Lead Days": "number", "On Hand": "number", "Allocated": "number",
                   "Unit Cost": "number", "Avg Daily Use": "number", "Safety Stock": "number"}),
        ],
        "fixture-research-model.xlsx": [
            table("model-inputs", "Model Inputs", "tblResearchInputs", ["Metric"],
                  {"Metric": "text", "Unit": "text", "Forecast Value": "number", "Description": "text"}),
            table("segment-detail", "Segment Detail", "tblResearchSegments", ["Year", "Segment"],
                  {"Year": "number", "Segment": "text", "Sort Order": "number",
                   "Revenue (USD m)": "number"}),
            records("audit-tolerances", ["Check"], {"Check": "text", "Tolerance": "number"},
                    [({"Check": f"Audit!A{r}", "Tolerance": f"Audit!E{r}"}, None)
                     for r in range(5, 9)]),
        ],
        "fixture-sales-pipeline.xlsx": [
            table("pipeline", "Pipeline", "tblSalesPipeline", ["Opportunity ID"],
                  {"Opportunity ID": "text", "Created Date": "datetime",
                   "Expected Close": "datetime", "Account": "text", "Region": "text",
                   "Segment": "text", "Rep": "text", "Stage": "text", "Probability": "number",
                   "ACV": "number", "Forecast Category": "text"}),
            table("rep-targets", "Rep Targets", "tblRepTargets", ["Rep"],
                  {"Rep": "text", "Annual Target": "number"}),
            wide("as-of-setting", "Assumptions", {"Setting": "E2", "As-of Date": "F2"},
                 {"Setting": "text", "As-of Date": "datetime"}, "Setting"),
        ],
    }


def add_research_historical_rows(features, source, direct_cells):
    """Group source-entered historical/forecast literals by each row's true schema."""
    by_row = defaultdict(list)
    for cell_id in direct_cells:
        if cell_id.startswith("FY results!"):
            address = cell_id.rsplit("!", 1)[1]
            col, row = re.fullmatch(r"([A-Z]+)(\d+)", address).groups()
            by_row[int(row)].append((col, cell_id))
    for row, columns in sorted(by_row.items()):
        fields = {"Line Item": f"FY results!A{row}"}
        for col, cell_id in sorted(columns):
            year = 2019 + ord(col) - ord("B")
            fields[f"FY{year}"] = cell_id
        features.append(records(f"research-history-{row}", ["Line Item"],
                                {"Line Item": "text", **{name: "number" for name in fields if name != "Line Item"}},
                                [(fields, None)]))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--originals", required=True, type=Path)
    parser.add_argument("--coverage", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--profiles-out", type=Path)
    parser.add_argument("--plans-dir", type=Path)
    args = parser.parse_args(argv)
    coverage = json.loads(args.coverage.read_text())
    sha = {name: info["sourceSha256"] for name, info in coverage.items()}
    profiles = source_profiles(sha)
    fixtures = HERE / "fixtures"
    output, reviewed_profiles = {}, {}
    if args.plans_dir:
        args.plans_dir.mkdir(parents=True, exist_ok=True)
    for name, features in profiles.items():
        source = (fixtures if name.startswith("fixture-") else args.originals) / name
        if hashlib.sha256(source.read_bytes()).hexdigest() != sha[name]:
            raise ValueError(f"source hash changed: {name}")
        if name == "fixture-research-model.xlsx":
            groups = coverage[name]["proposedInputDecisions"]["candidateGroups"]
            direct = [cell for group in groups for cell in group["sourceCellIds"]]
            add_research_historical_rows(features, source, direct)
        if name == "08 Inventory Replenishment Plan.xlsm":
            features[2]["emptySchema"] = True
            features[2]["reviewedBlankColumns"] = list(features[2]["columnTypes"])
            features[2]["typeReviewEvidence"] = "Snapshot macro append schema reviewed; source contains no rows"
        decision = {"sourceSha256": sha[name], "features": features}
        reviewed_profiles[name] = decision
        plan = build_source_input_plan(source, decision,
                                       connection_id="review-placeholder-connection",
                                       folder_id="review-placeholder-folder")
        validate_source_input_plan(plan, source)
        if args.plans_dir:
            slug = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
            (args.plans_dir / f"{slug}.json").write_text(
                json.dumps(plan, indent=2, ensure_ascii=False) + "\n")
        selected = Counter(cell for spec in plan["sharedInputArchitecture"]["inputs"].values()
                           for row in spec["sourceCells"].values() for cell in row.values())
        groups = coverage[name]["proposedInputDecisions"]["candidateGroups"]
        direct = {cell for group in groups for cell in group["sourceCellIds"]}
        missing = sorted(cell for cell in direct if selected[cell] == 0)
        ambiguous = sorted(cell for cell in direct if selected[cell] > 1)
        output[name] = {"sourceSha256": sha[name], "inputPlanHash": plan["inputPlanHash"],
                        "formulaObligations": len(plan["formulaObligations"]),
                        "expectedFormulaObligations": coverage[name]["formulaCount"],
                        "directLiteralCells": len(direct), "boundLiteralCells": len(direct) - len(missing),
                        "missingLiteralCells": missing, "ambiguousLiteralCells": ambiguous,
                        "features": {fid: {"rows": len(spec["rows"]),
                                            "columns": [c["name"] for c in spec["columns"]],
                                            "keyColumns": spec["keyColumns"]}
                                     for fid, spec in plan["sharedInputArchitecture"]["inputs"].items()}}
        if output[name]["formulaObligations"] != output[name]["expectedFormulaObligations"]:
            raise ValueError(f"formula obligations changed: {name}")
    args.out.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n")
    if args.profiles_out:
        args.profiles_out.write_text(json.dumps(reviewed_profiles, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps({name: {key: info[key] for key in ("directLiteralCells", "boundLiteralCells",
                                                       "missingLiteralCells", "ambiguousLiteralCells")}
                      for name, info in output.items()}, indent=2))
    if any(info["missingLiteralCells"] or info["ambiguousLiteralCells"] for info in output.values()):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
