"""Recorded Sigma Code Mode SDK signatures for the calls the setup program makes.

Transcribed from the SDK stubs the Sigma MCP server publishes (``list_resources`` /
``read_resource``), read on 2026-11-06 from the aws-us-west-2 production server:

    sdk://sdk/sigma/inputTable.py      InputTable, LiteralEditableColumn, InputTableEditMode,
                                       SharedInputTableFeatures.table_style, InputTableFeatures
    sdk://sdk/sigma/tableStyle.py      TableStyle, AutofitColumns, InputTableStyle (+ Literals)
    sdk://sdk/sigma/sheetElement.py    SheetFeatures.get_column, FormulaSheet, ResizableColumns,
                                       VisibleColumns, FreezableColumns, TableComponentsVisibility
    sdk://sdk/sigma/format.py          Formattable.set_format / set_format_pattern
    sdk://sdk/sigma/common.py          Visibility, ConstVal, HexColor
    sdk://sdk/sigma/cartesianChart.py  CartesianChart.list_series / set_series_color
    sdk://sdk/sigma/comboChart.py      ComboChart, ComboChartFeatures.set_series_color
    sdk://sdk/sigma/barChart.py        BarChart
    sdk://sdk/reference/sigma/{workbook,page,column,lineChart}.md
                                       get_page, ElementHost.get_element, ColumnBase.id, LineChart

Only the names, parameters, defaults, annotations and class bases are kept (docstrings and
the mixins no setup-program call reaches are dropped). ``code_mode_check`` binds every call
of an emitted program against these, so a wrong method name, arity, keyword or Literal value
fails when the program is emitted -- not halfway through its one-time run on a fresh
workbook. Refresh this file from the server when the SDK changes (the test suite pins it).
"""
import datetime
from typing import Literal, Union

SOURCE = "Sigma MCP sdk://sdk/sigma/*.py (aws-us-west-2 production), read 2026-11-06"

# --- sigma.common -------------------------------------------------------------------------
HexColor = str
ConstVal = Union[str, float, bool, datetime.datetime, None]

# --- sigma.tableStyle ---------------------------------------------------------------------
TableGridPreset = Literal["spreadsheet", "presentation"]
TableCellSpacing = Literal["extra-small", "small", "medium", "large"]
TableGridLines = Literal["none", "vertical", "horizontal", "all"]

# --- sigma.inputTable ---------------------------------------------------------------------
InputTableWritebackMode = Literal["draft", "published-explore", "published-all-users"]

# --- sigma.format -------------------------------------------------------------------------
FormatKind = Literal["number", "percent", "currency", "currency_si", "datetime", "automatic"]

# --- sigma.chart --------------------------------------------------------------------------
AxisSide = Literal["primary", "secondary"]


class Visibility:
    def is_visible(self) -> bool: ...
    def set_visible(self, visible: bool) -> None: ...


class Formattable:
    def set_format(self, kind: FormatKind, decimals: Union[int, None] = None,
                   currency: Union[str, None] = None) -> None: ...
    def set_format_pattern(self, pattern: str) -> None: ...


class ColumnBase:
    def id(self) -> str: ...
    def element_id(self) -> str: ...


class FormulaColumn(ColumnBase, Visibility, Formattable):
    def name(self) -> str: ...
    def formula(self) -> str: ...


class LiteralEditableColumn(ColumnBase, Visibility, Formattable):
    def name(self) -> str: ...


class SingleSelectEditableColumn(ColumnBase, Visibility, Formattable):
    def name(self) -> str: ...


class MultiSelectEditableColumn(ColumnBase, Visibility, Formattable):
    def name(self) -> str: ...


class FileEditableColumn(ColumnBase, Visibility, Formattable):
    def name(self) -> str: ...


class InputTableIDColumn(ColumnBase, Visibility, Formattable):
    pass


class InputTableBlameColumn(ColumnBase, Visibility, Formattable):
    pass


class LinkedInputTablePrimaryKeyColumn(ColumnBase, Visibility, Formattable):
    pass


EditableColumn = Union[LiteralEditableColumn, SingleSelectEditableColumn, MultiSelectEditableColumn,
                       FileEditableColumn]
EditableValue = Union[ConstVal, list[ConstVal]]
Column = Union[FormulaColumn, EditableColumn, InputTableIDColumn, InputTableBlameColumn,
               LinkedInputTablePrimaryKeyColumn]


class TableStyle:
    def preset(self) -> TableGridPreset: ...
    def set_preset(self, preset: TableGridPreset) -> None: ...
    def set_cell_spacing(self, spacing: TableCellSpacing) -> None: ...
    def set_grid_lines(self, lines: TableGridLines) -> None: ...
    def set_banding(self, show: bool) -> None: ...
    def set_banding_color(self, color: HexColor) -> None: ...
    def set_outer_border(self, show: bool) -> None: ...
    def set_header_divider_color(self, color: HexColor) -> None: ...


class AutofitColumns:
    def autofit_columns(self) -> bool: ...
    def set_autofit_columns(self, show: bool) -> None: ...


class InputTableStyle(TableStyle, AutofitColumns):
    pass


class ElementBase:
    def id(self) -> str: ...


class SheetFeatures:
    def get_column(self, id: str) -> Column: ...
    def get_column_by_name(self, name: str) -> Column: ...
    def has_column(self, name: str) -> bool: ...
    def list_columns(self) -> list[Column]: ...


class FormulaSheet(SheetFeatures):
    def get_column(self, id: str) -> FormulaColumn: ...
    def get_column_by_name(self, name: str) -> FormulaColumn: ...
    def list_columns(self) -> list[FormulaColumn]: ...


class VisibleColumns:
    def set_visible_columns(self, columns: list[Column]) -> None: ...


class FreezableColumns:
    def set_frozen_column(self, column: Union[Column, None]) -> None: ...


class ResizableColumns:
    def set_column_width(self, column: Column, width: Union[float, None]) -> None: ...


class TableComponentsVisibility:
    def set_show_table(self, show: bool = True) -> None: ...
    def set_summary_bar(self, show: bool = True) -> None: ...


class InputTableEditMode:
    def edit_mode(self) -> InputTableWritebackMode: ...
    def set_edit_mode(self, mode: InputTableWritebackMode) -> None: ...


class SharedInputTableFeatures:
    def table_style(self) -> InputTableStyle: ...


class InputTableFeatures:
    def update_rows(self, rows: list[tuple[str, dict[EditableColumn, EditableValue]]]) -> None: ...
    def delete_rows(self, row_ids: list[str]) -> None: ...
    def upsert_rows(self, rows: list[dict[EditableColumn, EditableValue]], index: int = -1,
                    sort_keys: Union[list, None] = None) -> list[str]: ...


class InputTable(ElementBase, SheetFeatures, InputTableEditMode, VisibleColumns, FreezableColumns,
                 ResizableColumns, TableComponentsVisibility, SharedInputTableFeatures, InputTableFeatures):
    pass


class CartesianChart(FormulaSheet):
    def list_series(self) -> list[FormulaColumn]: ...
    def set_series_color(self, series: FormulaColumn, color: HexColor) -> None: ...
    def get_series_color(self, series: FormulaColumn) -> Union[HexColor, None]: ...
    def set_series_axis(self, series: FormulaColumn, axis: AxisSide) -> None: ...


class ComboChart(ElementBase, CartesianChart):
    pass


class BarChart(ElementBase, CartesianChart):
    pass


class LineChart(ElementBase, CartesianChart):
    pass


Element = Union[InputTable, ComboChart, BarChart, LineChart]


class ElementHost:
    def get_element(self, element_id: str) -> Element: ...


class PageBase(ElementHost):
    def name(self) -> str: ...


class workbook:
    """The ``sigma.workbook`` module's free functions (called as ``sigma.workbook.<fn>``)."""

    @staticmethod
    def get_page(page_id: str) -> PageBase: ...

    @staticmethod
    def get_element(element_id: str) -> Element: ...


# Names a program may use bare (``isinstance(t, InputTable)``): the setup program's globals.
CLASSES = {c.__name__: c for c in (
    InputTable, ComboChart, BarChart, LineChart, LiteralEditableColumn, SingleSelectEditableColumn,
    MultiSelectEditableColumn, FileEditableColumn, FormulaColumn, InputTableIDColumn,
    InputTableBlameColumn, LinkedInputTablePrimaryKeyColumn, InputTableStyle)}
