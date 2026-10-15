#!/usr/bin/env python3
from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

from xlsx_audit import audit_workbook  # noqa: E402


class TestXlsxAudit(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def save(self, workbook, name="book.xlsx"):
        path = Path(self.tmp.name) / name
        workbook.save(path)
        return path

    def test_plain_workbook_is_automatic(self):
        wb = Workbook()
        wb.active["A1"] = "value"
        report = audit_workbook(self.save(wb))
        self.assertEqual(report["status"], "AUTO")
        self.assertEqual(report["issues"], [])

    def test_stored_excel_error_cells_require_review_without_counting_as_formulas(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "#DIV/0!"
        ws["A2"] = "#REF!"
        ws["A3"] = "#BUSY!"
        ws["A3"].data_type = "e"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["status"], "REVIEW_REQUIRED")
        self.assertEqual(report["inventory"]["formulaCells"], 0)
        self.assertEqual(
            report["inventory"]["formulaErrorCells"],
            ["Sheet!A1", "Sheet!A2", "Sheet!A3"],
        )
        issues = {item["code"]: item for item in report["issues"]}
        self.assertEqual(
            issues["FORMULA_ERRORS"]["evidence"],
            ["Sheet!A1", "Sheet!A2", "Sheet!A3"],
        )

    def test_plain_strings_containing_excel_errors_are_benign(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "#DIV/0!"
        ws["A1"].data_type = "s"
        ws["A2"] = "Error labels: #REF! and #BUSY!"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["status"], "AUTO")
        self.assertEqual(report["inventory"]["formulaErrorCells"], [])
        self.assertEqual(report["inventory"]["formulaCells"], 0)

    def test_unknown_and_volatile_formulas_require_review(self):
        wb = Workbook()
        wb.active["A1"] = "=OFFSET(B1,0,0)"
        wb.active["A2"] = "=MYSTERYFUNC(1)"
        report = audit_workbook(self.save(wb))
        codes = {item["code"] for item in report["issues"]}
        self.assertEqual(report["status"], "REVIEW_REQUIRED")
        self.assertIn("VOLATILE_FORMULAS", codes)
        self.assertIn("UNKNOWN_FORMULAS", codes)

    def test_array_formula_is_inventoried_and_routed_for_review(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Special"
        ws["A1"] = ArrayFormula(
            ref="A1:A3",
            text="=_xlfn._xlws.FILTER(B1:B3,B1:B3>0)",
        )

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]
        codes = {item["code"] for item in report["issues"]}

        self.assertEqual(inventory["formulaCells"], 1)
        self.assertEqual(
            inventory["arrayFormulaCells"],
            [{"cell": "Special!A1", "ref": "A1:A3"}],
        )
        self.assertEqual(inventory["functions"], {"FILTER": 1})
        self.assertEqual(inventory["dynamicFormulaCells"], ["Special!A1"])
        self.assertIn("ARRAY_FORMULAS_PRESENT", codes)
        self.assertIn("DYNAMIC_ARRAY_FORMULAS", codes)

    def test_two_variable_data_table_is_inventoried_and_routed_for_review(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "What If"
        ws["D1"] = DataTableFormula(
            ref="D1:E3",
            dt2D=True,
            r1="C1",
            r2="C2",
        )

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]
        codes = {item["code"] for item in report["issues"]}

        self.assertEqual(inventory["formulaCells"], 1)
        self.assertEqual(
            inventory["dataTableFormulaCells"],
            [{
                "cell": "What If!D1",
                "ref": "D1:E3",
                "twoVariable": True,
                "rowInput": "C1",
                "columnInput": "C2",
            }],
        )
        self.assertIn("WHAT_IF_DATA_TABLES_PRESENT", codes)

    def test_defined_name_expressions_are_inventoried_with_scope_and_reviewed(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Local"
        wb.defined_names.add(DefinedName(
            "ExternalInput",
            attr_text="'[ExternalInputs.xlsx]Rates'!$A$1",
        ))
        wb.defined_names.add(DefinedName(
            "ComputedInput",
            attr_text="OFFSET(Local!$A$1,0,0)",
        ))
        wb.defined_names.add(DefinedName(
            "ScopedInput",
            localSheetId=0,
            attr_text="MYSTERYFUNC(Local!$A$1)",
        ))
        wb.defined_names.add(DefinedName(
            "DynamicInput",
            localSheetId=0,
            attr_text="_xlfn._xlws.FILTER(Local!$A$1:$A$3,Local!$A$1:$A$3>0)",
        ))
        wb.defined_names.add(DefinedName(
            "SpillInput",
            localSheetId=0,
            attr_text="Local!$A$1#",
        ))
        ws["B1"] = "=ExternalInput"
        ws["B2"] = "=ComputedInput"
        ws["B3"] = "=ScopedInput"
        ws["B4"] = "=DynamicInput"
        ws["B5"] = "='[Direct.xlsx]Rates'!$A$1"
        ws["B6"] = "=SpillInput"

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]
        codes = {item["code"] for item in report["issues"]}

        names = {item["name"]: item for item in inventory["definedNames"]}
        self.assertEqual(
            names["ExternalInput"]["attrText"],
            "'[ExternalInputs.xlsx]Rates'!$A$1",
        )
        self.assertEqual(names["ExternalInput"]["scope"], "workbook")
        self.assertEqual(names["ScopedInput"]["scope"], "Local")
        self.assertEqual(names["ExternalInput"]["referencedBy"], ["Local!B1"])
        self.assertEqual(names["ScopedInput"]["referencedBy"], ["Local!B3"])
        self.assertEqual(inventory["externalFormulaCells"], ["Local!B5"])
        self.assertEqual(inventory["externalDefinedNames"], ["ExternalInput"])
        self.assertEqual(
            inventory["dynamicDefinedNames"],
            ["Local!DynamicInput", "Local!SpillInput"],
        )
        self.assertEqual(
            inventory["functions"],
            {"FILTER": 1, "MYSTERYFUNC": 1, "OFFSET": 1},
        )
        self.assertIn("EXTERNAL_WORKBOOK_LINKS", codes)
        self.assertIn("DYNAMIC_ARRAY_FORMULAS", codes)
        self.assertIn("VOLATILE_FORMULAS", codes)
        self.assertIn("UNKNOWN_FORMULAS", codes)

    def test_quoted_formula_text_does_not_create_formula_features(self):
        wb = Workbook()
        wb.active["A1"] = '=\"MYSTERYFUNC(1) OFFSET( [Budget.xlsx] # FILTER(\"'

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]

        self.assertEqual(report["status"], "AUTO")
        self.assertEqual(inventory["formulaCells"], 1)
        self.assertEqual(inventory["functions"], {})
        self.assertEqual(inventory["externalFormulaCells"], [])
        self.assertEqual(inventory["dynamicFormulaCells"], [])
        self.assertEqual(inventory["formulaErrorCells"], [])

    def test_structured_selectors_and_error_literals_are_not_spill_references(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=SUM(Table1[#All])"
        ws["A2"] = "=ROWS(Table1[#Headers])"
        ws["A3"] = "=#N/A"
        ws["A4"] = "=IFERROR(B4,#SPILL!)"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["inventory"]["dynamicFormulaCells"], [])
        self.assertEqual(
            report["inventory"]["formulaErrorCells"],
            ["Sheet!A3", "Sheet!A4"],
        )
        self.assertNotIn(
            "DYNAMIC_ARRAY_FORMULAS",
            {item["code"] for item in report["issues"]},
        )
        self.assertIn(
            "FORMULA_ERRORS",
            {item["code"] for item in report["issues"]},
        )

    def test_spill_reference_and_prefixed_dynamic_function_require_review(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=B1#"
        ws["A2"] = "=_xlfn._xlws.FILTER(B1:B3,B1:B3>0)"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["inventory"]["functions"], {"FILTER": 1})
        self.assertEqual(
            report["inventory"]["dynamicFormulaCells"],
            ["Sheet!A1", "Sheet!A2"],
        )
        self.assertIn(
            "DYNAMIC_ARRAY_FORMULAS",
            {item["code"] for item in report["issues"]},
        )

    def test_formula_issue_inventory_is_complete_beyond_twenty_five_cells(self):
        wb = Workbook()
        ws = wb.active
        for row in range(1, 31):
            ws.cell(row=row, column=1, value=f"='[Inputs.xlsx]Rates'!A{row}")
            ws.cell(row=row, column=2, value=f"=A{row}#")
            ws.cell(row=row, column=3, value="=#REF!")

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]

        self.assertEqual(len(inventory["externalFormulaCells"]), 30)
        self.assertEqual(inventory["externalFormulaCells"][-1], "Sheet!A30")
        self.assertEqual(len(inventory["dynamicFormulaCells"]), 30)
        self.assertEqual(inventory["dynamicFormulaCells"][-1], "Sheet!B30")
        self.assertEqual(len(inventory["formulaErrorCells"]), 30)
        self.assertEqual(inventory["formulaErrorCells"][-1], "Sheet!C30")
        evidence = {
            item["code"]: item.get("evidence", []) for item in report["issues"]
        }
        self.assertIn("Sheet!A30", evidence["EXTERNAL_WORKBOOK_LINKS"])
        self.assertIn("Sheet!B30", evidence["DYNAMIC_ARRAY_FORMULAS"])
        self.assertIn("Sheet!C30", evidence["FORMULA_ERRORS"])

    def test_indexed_external_ranges_are_detected_without_flagging_structured_refs(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "='[1]Sheet1'!$A$1"
        ws["A2"] = "=SUM(Table1[#All])"
        ws["A3"] = "=IndexedInput"
        wb.defined_names.add(DefinedName(
            "IndexedInput",
            attr_text="'[2]Sheet2'!$B$2",
        ))

        report = audit_workbook(self.save(wb))
        inventory = report["inventory"]

        self.assertEqual(inventory["externalFormulaCells"], ["Sheet!A1"])
        self.assertEqual(inventory["externalDefinedNames"], ["IndexedInput"])
        self.assertEqual(inventory["dynamicFormulaCells"], [])
        self.assertIn(
            "EXTERNAL_WORKBOOK_LINKS",
            {item["code"] for item in report["issues"]},
        )

    def test_modern_error_literals_require_review_without_becoming_spills(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=#BUSY!"
        ws["A2"] = "=B2+#PYTHON!"
        ws["A3"] = "=#DATA!"
        ws["A4"] = '=\"#BUSY! #PYTHON! #DATA!\"'

        report = audit_workbook(self.save(wb))

        self.assertEqual(
            report["inventory"]["formulaErrorCells"],
            ["Sheet!A1", "Sheet!A2", "Sheet!A3"],
        )
        self.assertEqual(report["inventory"]["dynamicFormulaCells"], [])
        self.assertIn(
            "FORMULA_ERRORS",
            {item["code"] for item in report["issues"]},
        )

    def test_function_like_quoted_sheet_name_is_a_range_not_a_function(self):
        wb = Workbook()
        wb.active["A1"] = "='MYSTERYFUNC(1)'!A1"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["inventory"]["functions"], {})
        self.assertNotIn(
            "UNKNOWN_FORMULAS",
            {item["code"] for item in report["issues"]},
        )

    def test_structured_reference_escape_does_not_hide_following_formula_error(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["#OfItems", "Result"])
        ws.append([1, None])
        ws.add_table(Table(displayName="Table1", ref="A1:A2"))
        ws["B2"] = "=SUM(Table1['#OfItems])+#REF!"

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["inventory"]["functions"], {"SUM": 1})
        self.assertEqual(report["inventory"]["formulaErrorCells"], ["Sheet!B2"])
        self.assertEqual(report["inventory"]["dynamicFormulaCells"], [])
        self.assertIn(
            "FORMULA_ERRORS",
            {item["code"] for item in report["issues"]},
        )

    def test_escaped_table_headers_do_not_hide_following_spill_or_functions(self):
        for header, reference in [
            ("Customer#", "Customer'#"),
            ("#OfItems", "'#OfItems"),
            ("Customer[Code", "Customer'[Code"),
            ("Customer]Code", "Customer']Code"),
            ('Customer"Name', 'Customer"Name'),
        ]:
            with self.subTest(header=header):
                wb = Workbook()
                ws = wb.active
                ws.append([header])
                ws.append([1])
                ws.add_table(Table(displayName="Table1", ref="A1:A2"))
                ws["B2"] = f"=SUM(Table1[{reference}])+D1#"
                ws["D1"] = ArrayFormula(ref="D1:D2", text="=SUM(1,2)")

                report = audit_workbook(self.save(wb))

                self.assertEqual(report["inventory"]["functions"], {"SUM": 2})
                self.assertEqual(
                    report["inventory"]["dynamicFormulaCells"], ["Sheet!B2"],
                )
                issues = {item["code"]: item for item in report["issues"]}
                self.assertEqual(
                    issues["DYNAMIC_ARRAY_FORMULAS"]["evidence"], ["Sheet!B2"],
                )

    def test_table_header_characters_do_not_create_spills_or_functions(self):
        for header, reference in [
            ("Customer#", "Customer'#"),
            ("Customer[Code", "Customer'[Code"),
            ("Customer]Code", "Customer']Code"),
            ('Customer"Name', 'Customer"Name'),
            ("MYSTERYFUNC(1)#REF!", "MYSTERYFUNC(1)'#REF!"),
        ]:
            with self.subTest(header=header):
                wb = Workbook()
                ws = wb.active
                ws.append([header])
                ws.append([1])
                ws.add_table(Table(displayName="Table1", ref="A1:A2"))
                ws["B2"] = f"=SUM(Table1[{reference}])"

                report = audit_workbook(self.save(wb))

                expected_status = (
                    "REVIEW_REQUIRED" if header == "Customer[Code" else "AUTO"
                )
                self.assertEqual(report["status"], expected_status)
                self.assertEqual(report["inventory"]["functions"], {"SUM": 1})
                self.assertEqual(report["inventory"]["dynamicFormulaCells"], [])
                self.assertEqual(report["inventory"]["formulaErrorCells"], [])

    def test_tokenizer_fallback_keeps_nested_functions_and_ignores_quoted_names(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Customer[Code"])
        ws.append([1])
        ws.add_table(Table(displayName="Table1", ref="A1:A2"))
        wb.create_sheet("SIN(1)")["E1"] = 2
        ws["B2"] = (
            '=MYSTERYFUNC(SUM(Table1[Customer\'[Code]))'
            '+\'SIN(1)\'!E1+IF(1,"OFFSET(1)",0)+D1#'
        )

        report = audit_workbook(self.save(wb))

        self.assertEqual(
            report["inventory"]["functions"],
            {"IF": 1, "MYSTERYFUNC": 1, "SUM": 1},
        )
        self.assertEqual(report["inventory"]["dynamicFormulaCells"], ["Sheet!B2"])
        codes = {item["code"] for item in report["issues"]}
        self.assertIn("UNKNOWN_FORMULAS", codes)
        self.assertIn("DYNAMIC_ARRAY_FORMULAS", codes)
        self.assertNotIn("VOLATILE_FORMULAS", codes)

    def test_unparsed_formulas_cannot_hide_external_links_behind_automatic_status(self):
        for workbook_reference in ["[1]", "[Inputs.xlsx]"]:
            with self.subTest(workbook_reference=workbook_reference):
                wb = Workbook()
                ws = wb.active
                ws.append(["Customer[Code"])
                ws.append([1])
                ws.add_table(Table(displayName="Table1", ref="A1:A2"))
                ws["B2"] = (
                    "=SUM(Table1[Customer'[Code])+"
                    f"'{workbook_reference}Sheet1'!A1"
                )

                report = audit_workbook(self.save(wb))

                self.assertEqual(report["status"], "REVIEW_REQUIRED")
                self.assertEqual(report["inventory"]["functions"], {"SUM": 1})
                self.assertEqual(
                    report["inventory"]["unparsedFormulaCells"], ["Sheet!B2"],
                )
                issues = {item["code"]: item for item in report["issues"]}
                self.assertEqual(
                    issues["UNPARSED_FORMULAS"]["evidence"], ["Sheet!B2"],
                )
                self.assertIn(
                    "formula-review-or-python", report["routing"]["review"],
                )

    def test_unparsed_defined_name_expressions_require_review_with_scope(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Local"
        ws.append(["Customer[Code"])
        ws.append([1])
        ws.add_table(Table(displayName="Table1", ref="A1:A2"))
        wb.defined_names.add(DefinedName(
            "ExternalInput",
            attr_text="SUM(Table1[Customer'[Code])+'[Inputs.xlsx]Sheet1'!A1",
        ))
        ws.defined_names.add(DefinedName(
            "LocalInput",
            localSheetId=0,
            attr_text="SUM(Table1[Customer'[Code])+'[1]Sheet1'!A1",
        ))

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["status"], "REVIEW_REQUIRED")
        self.assertEqual(report["inventory"]["functions"], {"SUM": 2})
        self.assertEqual(
            report["inventory"]["unparsedDefinedNames"],
            ["ExternalInput", "Local!LocalInput"],
        )
        issues = {item["code"]: item for item in report["issues"]}
        self.assertEqual(
            issues["UNPARSED_FORMULAS"]["evidence"],
            ["defined-name:ExternalInput", "defined-name:Local!LocalInput"],
        )

    def test_external_reference_text_literals_do_not_require_tokenizer_review(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "'[1]Sheet1'!A1"
        ws["A2"] = '=\"\'[1]Sheet1\'!A1\"'
        ws["A3"] = '=\"SUM(Table1[Customer\'[Code])+\'[Inputs.xlsx]Sheet1\'!A1\"'

        report = audit_workbook(self.save(wb))

        self.assertEqual(report["status"], "AUTO")
        self.assertEqual(report["issues"], [])

    def test_local_defined_name_shadows_global_and_explicit_scope_resolves_local(self):
        wb = Workbook()
        first = wb.active
        first.title = "S1"
        second = wb.create_sheet("S2")
        wb.defined_names.add(DefinedName("Rate", attr_text="S2!$B$1"))
        first.defined_names.add(DefinedName(
            "Rate",
            localSheetId=0,
            attr_text="S1!$B$1",
        ))
        first["A1"] = "=Rate"
        second["A1"] = "=Rate"
        second["A2"] = "=S1!Rate"

        report = audit_workbook(self.save(wb))
        rate_names = [
            item for item in report["inventory"]["definedNames"]
            if item["name"] == "Rate"
        ]
        by_scope = {item["scope"]: item["referencedBy"] for item in rate_names}

        self.assertEqual(by_scope["workbook"], ["S2!A1"])
        self.assertEqual(by_scope["S1"], ["S1!A1", "S2!A2"])

    def test_chart_validation_and_protection_are_routed(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Category", "Value"])
        ws.append(["A", 1])
        ws.append(["B", 2])
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
        ws.add_chart(chart, "D2")
        validation = DataValidation(type="list", formula1='"A,B"')
        ws.add_data_validation(validation)
        validation.add(ws["A2"])
        ws.protection.sheet = True

        report = audit_workbook(self.save(wb))
        codes = {item["code"] for item in report["issues"]}
        self.assertTrue({
            "CHARTS_PRESENT", "DATA_VALIDATION_PRESENT", "WORKBOOK_PROTECTION"
        }.issubset(codes))

    def test_connection_and_custom_xml_are_not_silent(self):
        wb = Workbook()
        path = self.save(wb)
        with zipfile.ZipFile(path, "a") as archive:
            archive.writestr("xl/connections.xml", "<connections/>")
            archive.writestr("customXml/item1.xml", "<business-metadata/>")
        report = audit_workbook(path)
        codes = {item["code"] for item in report["issues"]}
        self.assertIn("POWER_QUERY_OR_CONNECTIONS", codes)
        self.assertIn("CUSTOM_XML", codes)

    def test_legacy_binary_is_blocked_before_parsing(self):
        path = Path(self.tmp.name) / "book.xls"
        path.write_bytes(b"not an OOXML package")
        report = audit_workbook(path)
        self.assertEqual(report["status"], "BLOCKED")
        self.assertEqual(report["issues"][0]["code"], "LEGACY_OR_UNSUPPORTED_FORMAT")


if __name__ == "__main__":
    unittest.main()
