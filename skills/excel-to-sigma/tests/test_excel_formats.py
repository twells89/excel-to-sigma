#!/usr/bin/env python3
"""Excel number-format codes -> Sigma format objects (table-driven)."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

import excel_formats as ef  # noqa: E402


def num(string, **extra):
    return {"kind": "number", "formatString": string, **extra}


def date(string):
    return {"kind": "datetime", "formatString": string}


# code -> (category, Sigma format, negative_red)
EXACT = [
    ("General", "general", None, False),
    ("", "general", None, False),
    ("@", "text", None, False),
    ("0", "number", num(".0f"), False),
    ("0.0", "number", num(".1f"), False),
    ("0.00", "number", num(".2f"), False),
    ("#,##0", "number", num(",.0f"), False),
    ("#,##0.0", "number", num(",.1f"), False),
    ("#,##0.00", "number", num(",.2f"), False),
    ("#,##0.##", "number", num(",.2~f"), False),
    ("$#,##0", "currency", num("$,.0f"), False),
    ("$#,##0.00", "currency", num("$,.2f"), False),
    ('"$"#,##0', "currency", num("$,.0f"), False),
    ("[$$-409]#,##0", "currency", num("$,.0f"), False),
    ("[$$-409]#,##0.00", "currency", num("$,.2f"), False),
    ("€#,##0", "currency", num(",.0f", prefix="€"), False),
    ("£#,##0.00", "currency", num(",.2f", prefix="£"), False),
    ("[$€-407] #,##0", "currency", num(",.0f", prefix="€ "), False),
    ("[$£-809]#,##0.00", "currency", num(",.2f", prefix="£"), False),
    ("0%", "percent", num(".0%"), False),
    ("0.0%", "percent", num(".1%"), False),
    ("0.00%", "percent", num(".2%"), False),
    ("#,##0%", "percent", num(",.0%"), False),
    # negatives: d3's "(" sign is parentheses; [Red] is a conditional-format rule, not part of the format
    ("#,##0;(#,##0)", "number", num("(,.0f"), False),
    ("#,##0;[Red](#,##0)", "number", num("(,.0f"), True),
    ("#,##0_);(#,##0)", "number", num("(,.0f"), False),
    ("#,##0_);[Red](#,##0)", "number", num("(,.0f"), True),
    ("#,##0.00_);[Red](#,##0.00)", "number", num("(,.2f"), True),
    ("#,##0;-#,##0", "number", num(",.0f"), False),
    ("#,##0.00;[Red]-#,##0.00", "number", num(",.2f"), True),
    ("$#,##0_);($#,##0)", "currency", num("($,.0f"), False),
    ("$#,##0;[Red]($#,##0)", "currency", num("($,.0f"), True),
    ("0.0%;(0.0%)", "percent", num("(.1%"), False),
    ('0 "units"', "number", num(".0f", suffix=" units"), False),
    # dates
    ("yyyy-mm-dd", "date", date("%Y-%m-%d"), False),
    ("yyyy-mm-dd;@", "date", date("%Y-%m-%d"), False),
    ("m/d/yyyy", "date", date("%-m/%-d/%Y"), False),
    ("mm/dd/yyyy", "date", date("%m/%d/%Y"), False),
    ("mm-dd-yy", "date", date("%-m/%-d/%Y"), False),      # openpyxl's name for built-in 14
    ("mmm-yy", "date", date("%b-%y"), False),
    ("mmm d, yyyy", "date", date("%b %-d, %Y"), False),
    ("d-mmm-yy", "date", date("%-d-%b-%y"), False),
    ("mmmm yyyy", "date", date("%B %Y"), False),
    ("dddd, mmmm dd, yyyy", "date", date("%A, %B %d, %Y"), False),
    ("[$-F800]dddd, mmmm dd, yyyy", "date", date("%A, %B %d, %Y"), False),
    ("h:mm AM/PM", "date", date("%-I:%M %p"), False),
    ("h:mm:ss", "date", date("%-H:%M:%S"), False),
    ("yyyy-mm-dd hh:mm:ss", "date", date("%Y-%m-%d %H:%M:%S"), False),
    ("m/d/yy h:mm", "date", date("%-m/%-d/%y %-H:%M"), False),
]


class TestTranslate(unittest.TestCase):
    def test_exact_translations(self):
        for code, category, fmt, red in EXACT:
            with self.subTest(code=code):
                result = ef.translate(code)
                self.assertEqual(result.category, category)
                self.assertEqual(result.format, fmt)
                self.assertEqual(result.negative_red, red)
                self.assertEqual(result.approximated, [], code)

    def test_accounting_format_is_translated_and_the_zero_dash_is_reported(self):
        result = ef.translate('_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)')
        self.assertEqual(result.format, num("($,.0f"))
        self.assertEqual(result.category, "currency")
        self.assertEqual(result.approximated, ["the zero section is shown with the positive format"])

    def test_parentheses_fall_back_to_a_minus_sign_when_not_native(self):
        result = ef.translate("#,##0;[Red](#,##0)", native_parens=False)
        self.assertEqual(result.format, num(",.0f"))
        self.assertTrue(result.negative_red)
        self.assertEqual(result.approximated, ["negatives in parentheses are shown with a minus sign"])

    def test_scientific_notation_maps_with_a_reported_exponent_difference(self):
        result = ef.translate("0.00E+00")
        self.assertEqual(result.format, num(".2e"))
        self.assertEqual(result.approximated, ["scientific exponent is shown as e+N, not E+NN"])

    def test_unknown_codes_are_reported_never_dropped(self):
        for code, fragment in (("# ?/?", "fraction"), ("[h]:mm", "elapsed"), ("[>100]0;0", "conditional"),
                               ("#,##0,", "divide by 1000"), ("000", "leading zeros")):
            with self.subTest(code=code):
                self.assertTrue(any(fragment in note for note in ef.translate(code).approximated))

    def test_the_minutes_token_follows_excel_s_rule(self):
        self.assertEqual(ef.translate("h:mm").format["formatString"], "%-H:%M")
        self.assertEqual(ef.translate("mm:ss").format["formatString"], "%M:%S")
        self.assertEqual(ef.translate("mm/dd").format["formatString"], "%m/%d")

    def test_format_strings_are_d3_never_excel_masks(self):
        for code, _, fmt, _ in EXACT:
            if fmt and fmt["kind"] == "number":
                self.assertNotIn("#", fmt["formatString"], code)       # Sigma rejects "$#,##0"
                self.assertRegex(fmt["formatString"], r"^\(?\$?,?\.\d(~?[f%])$", code)


class TestTranslateFor(unittest.TestCase):
    def test_general_numbers_become_a_plain_number_with_the_decimals_the_values_need(self):
        fmt, red, reasons, _ = ef.translate_for("General", kind="number", values=[1, 2, 3])
        self.assertEqual((fmt, red, reasons), (num(",.0f"), False, []))
        fmt, _, _, _ = ef.translate_for("General", kind="number", values=[1, 2.5, 3])
        self.assertEqual(fmt, num(",.1f"))
        fmt, _, _, _ = ef.translate_for(None, kind="number", values=[0.12345678])
        self.assertEqual(fmt, num(",.4f"))

    def test_float_noise_does_not_add_decimals(self):
        fmt, *_ = ef.translate_for("General", kind="number", values=[11600.000000000002, 114799.99999999999])
        self.assertEqual(fmt, num(",.0f"))

    def test_unknown_numeric_codes_become_plain_numbers_and_say_why(self):
        fmt, red, reasons, _ = ef.translate_for("# ?/?", kind="number", values=[1])
        self.assertEqual(fmt, num(",.0f"))
        self.assertEqual(reasons, ["fraction formats are not supported"])

    def test_a_date_mask_on_a_number_column_is_not_used(self):
        fmt, _, reasons, _ = ef.translate_for("yyyy-mm-dd", kind="number", values=[1.5])
        self.assertEqual(fmt["kind"], "number")

    def test_datetime_columns_keep_their_mask_or_get_iso_dates(self):
        self.assertEqual(ef.translate_for("mmm-yy", kind="datetime")[0], date("%b-%y"))
        self.assertEqual(ef.translate_for("General", kind="datetime")[0], date("%Y-%m-%d"))

    def test_percent_detection_covers_legacy_and_current_plans(self):
        self.assertTrue(ef.is_percent_format("percent"))
        self.assertTrue(ef.is_percent_format(num(".1%")))
        self.assertFalse(ef.is_percent_format(num(",.0f")))
        self.assertFalse(ef.is_percent_format(None))


if __name__ == "__main__":
    unittest.main()
