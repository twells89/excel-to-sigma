#!/usr/bin/env python3
"""The setup program's Code Mode calls are bound against the recorded SDK signatures
(scripts/lib/sdk_signatures.py, transcribed from the Sigma MCP server's sdk://sdk/sigma/*.py stubs).
A fake host accepts any call; this checker does not, so a wrong method name, arity or Literal value
fails when the program is emitted instead of halfway through its one-time run."""
import inspect
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

import code_mode_check  # noqa: E402
import sdk_signatures as SDK  # noqa: E402

PROGRAM = """\
t = sigma.workbook.get_page('pgAssumptions').get_element('rateInput')
if not isinstance(t, InputTable):
    raise RuntimeError('rateInput is not an InputTable')
c = t.get_column('RATE')
if not isinstance(c, LiteralEditableColumn):
    raise RuntimeError('rateInput.RATE is not an editable column')
c.set_format_pattern('.1%')
c.set_visible(False)
s = t.table_style()
s.set_preset('spreadsheet')
s.set_grid_lines('horizontal')
s.set_autofit_columns(False)
t.set_column_width(t.get_column('LINE_ITEM'), 231)
t.set_edit_mode('draft')
t.set_edit_mode('published-all-users')
ch = sigma.workbook.get_page('pgForecast').get_element('revTrend')
if not isinstance(ch, ComboChart):
    raise RuntimeError('revTrend is not a ComboChart')
for s in ch.list_series():
    if s.id() == 'plan':
        ch.set_series_color(s, '#475569')
t = sigma.workbook.get_page('pgAssumptions').get_element('rateInput')
if not isinstance(t, InputTable):
    raise RuntimeError('rateInput is not an InputTable')
c_l = t.get_column('LINE_ITEM')
if not isinstance(c_l, LiteralEditableColumn):
    raise RuntimeError('x')
c_r = t.get_column('RATE')
if not isinstance(c_r, LiteralEditableColumn):
    raise RuntimeError('x')
t.upsert_rows([
    {c_l: 'West Revenue', c_r: 0.012},
    {c_l: 'East Revenue', c_r: 1, },
])
"""


def errors(program):
    return code_mode_check.check_program(program)[0]


class RecordedSignatureTests(unittest.TestCase):
    def test_representative_program_binds_cleanly(self):
        errs, calls = code_mode_check.check_program(PROGRAM)
        self.assertEqual(errs, [])
        self.assertGreaterEqual(calls, 20)

    def test_checker_runs_on_python_without_type_alias_type(self):
        # typing.TypeAliasType is Python 3.12+; the checker runs on every emit, so it must not need it.
        # (types.UnionType, 3.10+, is read the same way; deleting it here would break typing itself.)
        import importlib
        import typing
        saved = [(typing, "TypeAliasType", typing.TypeAliasType)] if hasattr(typing, "TypeAliasType") else []
        try:
            for mod, name, _value in saved:
                delattr(mod, name)
            module = importlib.reload(code_mode_check)
            errs, calls = module.check_program(PROGRAM)
            self.assertEqual(errs, [])
            self.assertGreaterEqual(calls, 20)
        finally:
            for mod, name, value in saved:
                setattr(mod, name, value)
            importlib.reload(code_mode_check)

    def test_wrong_calls_are_rejected(self):
        cases = {
            ("s.set_autofit_columns(False)", "s.set_autofit_columns('hidden')"): "argument show",
            ("s.set_preset('spreadsheet')", "s.set_preset('compact')"): "'compact' is not one of",
            ("s.set_grid_lines('horizontal')", "s.set_gridlines('horizontal')"): "no method set_gridlines",
            ("s = t.table_style()", "s = t.table_style(True)"): "too many positional",
            ("t.set_column_width(t.get_column('LINE_ITEM'), 231)",
             "t.set_column_width('LINE_ITEM', 231)"): "argument column",
            ("t.set_column_width(t.get_column('LINE_ITEM'), 231)",
             "t.set_column_width(t.get_column('LINE_ITEM'))"): "missing a required argument",
            ("t.set_column_width(t.get_column('LINE_ITEM'), 231)",
             "t.set_column_width(t.get_column('LINE_ITEM'), px=231)"): "missing a required argument: 'width'",
            ("t.set_edit_mode('draft')", "t.set_edit_mode('edit')"): "'edit' is not one of",
            ("c.set_visible(False)", "c.set_visible('no')"): "argument visible",
            ("ch.set_series_color(s, '#475569')", "ch.set_series_color('plan', '#475569')"): "argument series",
            ("c_r: 0.012}", "'RATE': 0.012}"): "argument rows",
            ("t.upsert_rows([", "t.upsert_rows(rows=[], index='end', data=["): "unexpected keyword",
            ("if not isinstance(ch, ComboChart):", "if not isinstance(ch, Gauge):"): "not a recorded SDK class",
        }
        for (old, new), expected in cases.items():
            with self.subTest(new=new):
                self.assertIn(old, PROGRAM)
                errs = errors(PROGRAM.replace(old, new, 1))
                self.assertTrue(errs, new)
                self.assertIn(expected, " ".join(errs))

    def test_unnarrowed_handle_is_rejected(self):
        program = PROGRAM.replace("if not isinstance(t, InputTable):\n    raise RuntimeError('rateInput is not an InputTable')\n"
                                  "c = t.get_column('RATE')", "c = t.get_column('RATE')", 1)
        # get_element returns any element: a chart has get_column but no table_style
        self.assertIn("ComboChart has no method table_style()", " ".join(errors(program)))

    def test_program_outside_the_subset_is_reported(self):
        self.assertIn("unsupported statement", " ".join(errors("import os\n")))
        self.assertIn("is not defined", " ".join(errors("x.set_visible(True)\n")))
        self.assertIn("syntax error", " ".join(errors("t = (\n")))

    def test_recorded_stub_pins_the_signatures_the_setup_program_uses(self):
        def params(fn):
            return [(p.name, p.default) for p in inspect.signature(fn).parameters.values()][1:]
        empty = inspect.Parameter.empty
        self.assertEqual(params(SDK.SharedInputTableFeatures.table_style), [])
        self.assertTrue(issubclass(SDK.InputTable, SDK.SharedInputTableFeatures))
        self.assertTrue(issubclass(SDK.InputTableStyle, SDK.AutofitColumns))
        self.assertEqual(params(SDK.TableStyle.set_preset), [("preset", empty)])
        self.assertEqual(params(SDK.TableStyle.set_grid_lines), [("lines", empty)])
        self.assertEqual(params(SDK.AutofitColumns.set_autofit_columns), [("show", empty)])
        self.assertEqual(params(SDK.ResizableColumns.set_column_width), [("column", empty), ("width", empty)])
        self.assertTrue(issubclass(SDK.InputTable, SDK.ResizableColumns))
        self.assertEqual(params(SDK.InputTableEditMode.set_edit_mode), [("mode", empty)])
        self.assertEqual(params(SDK.InputTableFeatures.upsert_rows),
                         [("rows", empty), ("index", -1), ("sort_keys", None)])
        self.assertEqual(params(SDK.Formattable.set_format_pattern), [("pattern", empty)])
        self.assertEqual(params(SDK.Visibility.set_visible), [("visible", empty)])
        self.assertEqual(params(SDK.CartesianChart.set_series_color), [("series", empty), ("color", empty)])
        import typing
        self.assertEqual(typing.get_args(SDK.TableGridPreset), ("spreadsheet", "presentation"))
        self.assertEqual(typing.get_args(SDK.TableGridLines), ("none", "vertical", "horizontal", "all"))
        self.assertEqual(typing.get_args(SDK.InputTableWritebackMode),
                         ("draft", "published-explore", "published-all-users"))
        self.assertIn("2026-11-06", SDK.SOURCE)


if __name__ == "__main__":
    unittest.main()
