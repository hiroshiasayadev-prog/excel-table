"""
Table formatting and schema models.

This module contains two families of classes:

**Formatted table classes** (Write side)
    Wrap a concrete table instance with rendering metadata.
    Used as items in :class:`~excel_table.writer.SheetWriteSchema`.

**Schema classes** (Read side)
    Describe the expected type and layout of a table to be read from a sheet.
    Used as items in :class:`~excel_table.reader.SheetReadSchema`.

Conditional format dict syntax
-------------------------------
``value_conditional_formats`` accepts a list of xlsxwriter
``conditional_format`` parameter dicts. Each dict is passed verbatim to
``worksheet.conditional_format(range, dict)`` after one substitution:

- If the dict contains a ``"cell_format"`` key whose value is a plain
  ``dict``, the writer replaces it with a ``workbook.add_format(...)``
  object under the key ``"format"`` before calling ``conditional_format``.

Examples::

    # 3-color gradient (no cell_format needed)
    value_conditional_formats=[
        {"type": "3_color_scale",
         "min_color": "#0000FF", "mid_color": "#FFFFFF", "max_color": "#FF0000"},
    ]

    # Gradient + outlier highlighting
    value_conditional_formats=[
        {"type": "2_color_scale",
         "min_color": "#FFFFFF", "max_color": "#FF0000"},
        {"type": "cell", "criteria": ">", "value": 1000,
         "cell_format": {"bg_color": "#000000", "font_color": "#FFFFFF"}},
        {"type": "cell", "criteria": "<", "value": 0,
         "cell_format": {"bg_color": "#FF0000"}},
    ]
"""
from __future__ import annotations

from typing import Literal, Type

from pydantic import BaseModel

from .table_base import Table1D, Table2D, TableKeyValue


# ---------------------------------------------------------------------------
# Write-side: Formatted table classes
# ---------------------------------------------------------------------------

class FormattedTable1D(BaseModel):
    """
    A :class:`~excel_table.models.table_base.Table1D` bundled with rendering
    metadata.

    Orientation
    -----------
    ``orientation="horizontal"`` (default) renders the table with column
    headers across the top::

        (0,0) title
        (1,0) column_label  [merged, direction controlled by column_label_direction]
        (2,0) col1  (2,1) col2  (2,2) col3
        (3,0) val   (3,1) val   (3,2) val

    ``orientation="vertical"`` rotates the table 90 degrees so that column
    headers run down the left side::

        (0,0) title
        (1,0) column_label  (1,1) col1  (1,2) val
        (2,0)               (2,1) col2  (2,2) val
        (3,0)               (3,1) col3  (3,2) val

    Attributes:
        table: The underlying data table.
        orientation: Table layout direction. ``"horizontal"`` (default) places
            column headers across the top row; ``"vertical"`` places them down
            the left column.
        column_color: Hex background color for column header cells.
        value_conditional_formats: List of xlsxwriter ``conditional_format``
            parameter dicts applied to the data cell range. See module
            docstring for the ``cell_format`` substitution rule.
        column_label_direction: Text orientation of the ``column_label`` cell.
            ``"horizontal"`` (default) renders text normally;
            ``"vertical"`` rotates text 90 degrees.
    """

    table: Table1D
    orientation: Literal["horizontal", "vertical"] = "horizontal"
    column_color: str
    value_conditional_formats: list[dict] = []
    column_label_direction: Literal["horizontal", "vertical"] = "horizontal"


class FormattedTable2D(BaseModel):
    """
    A :class:`~excel_table.models.table_base.Table2D` bundled with rendering
    metadata.

    Layout variants
    ---------------
    The four ``column_location`` / ``row_location`` combinations produce the
    following layouts (all coordinates 0-indexed from the title anchor):

    **column_location="top", row_location="left"** (default)::

        (0,0) title
        (1,0) [2×2 blank]      (1,2) column_label [merged]
        (2,0)                  (2,2) col1  col2  col3
        (3,0) row_label [↓]    (3,1) row1  val   val
        (4,0)                  (4,1) row2  val   val

    **column_location="top", row_location="right"**::

        (0,0) title
        (1,0) column_label [merged]       (1,n) [2×2 blank]
        (2,0) col1  col2  col3            (2,n)
        (3,0) row1  val   val   (3,n-1)   (3,n) row_label [↓]
        (4,0) row2  val   val   (4,n-1)

    **column_location="bottom", row_location="left"**::

        (0,0) row_label [↓]    (0,1) row1  val   val
        (1,0)                  (1,1) row2  val   val
        (2,0)                  (2,2) col1  col2  col3
        (3,0) [2×2 blank]      (3,2) column_label [merged]
        (4,0) title

    **column_location="bottom", row_location="right"**::

        (0,0) row1  val   val   (0,n-1)   (0,n) row_label [↓]
        (1,0) row2  val   val   (1,n-1)
        (2,0) col1  col2  col3
        (3,0) column_label [merged]       (3,n) [2×2 blank]
        (4,0) title                       (4,n)

    Attributes:
        table: The underlying data table.
        column_location: Which edge the column headers appear on.
            ``"top"`` (default) or ``"bottom"``.
        row_location: Which edge the row headers appear on.
            ``"left"`` (default) or ``"right"``.
        column_color: Hex background color for column header cells.
        row_color: Hex background color for row header cells (including
            the ``row_label`` cell).
        value_conditional_formats: List of xlsxwriter ``conditional_format``
            parameter dicts applied to the data cell range. See module
            docstring for the ``cell_format`` substitution rule.
        column_label_direction: Text orientation of the ``column_label`` cell.
            ``"horizontal"`` (default) or ``"vertical"`` (90-degree rotation).
        row_label_direction: Text orientation of the ``row_label`` cell.
            ``"horizontal"`` (default) or ``"vertical"`` (90-degree rotation).
    """

    table: Table2D
    column_location: Literal["top", "bottom"] = "top"
    row_location: Literal["left", "right"] = "left"
    column_color: str = "white"
    row_color: str = "white"
    value_conditional_formats: list[dict] = []
    column_label_direction: Literal["horizontal", "vertical"] = "horizontal"
    row_label_direction: Literal["horizontal", "vertical"] = "horizontal"


# ---------------------------------------------------------------------------
# Read-side: Schema classes
# ---------------------------------------------------------------------------

class FormattedTable2DSchema(BaseModel):
    """
    Read schema for a :class:`~excel_table.models.table_base.Table2D`.

    Describes the expected type and layout of a ``Table2D`` to be located and
    read from a worksheet. Used as an item in
    :class:`~excel_table.reader.SheetReadSchema`.

    The ``column_location`` and ``row_location`` fields must match the actual
    layout of the table in the Excel file so that the reader can compute the
    correct cell offsets from the title anchor.

    Attributes:
        title: Exact title string to search for in the sheet.
        table_type: Concrete subclass to instantiate via ``model_validate``.
            May be any typed or domain subclass of
            :class:`~excel_table.models.table_base.Table2D`.
        column_location: Edge where column headers appear. Must match the
            written layout. Default ``"top"``.
        row_location: Edge where row headers appear. Must match the written
            layout. Default ``"left"``.
    """

    title: str
    table_type: Type[Table2D]
    column_location: Literal["top", "bottom"] = "top"
    row_location: Literal["left", "right"] = "left"


class FormattedTable1DSchema(BaseModel):
    """
    Read schema for a :class:`~excel_table.models.table_base.Table1D`.

    Attributes:
        title: Exact title string to search for in the sheet.
        table_type: Concrete subclass to instantiate via ``model_validate``.
        orientation: Must match the orientation used when the table was
            written. Default ``"horizontal"``.
    """

    title: str
    table_type: Type[Table1D]
    orientation: Literal["horizontal", "vertical"] = "horizontal"


class TableKeyValueSchema(BaseModel):
    """
    Read schema for a :class:`~excel_table.models.table_base.TableKeyValue`.

    :class:`~excel_table.models.table_base.TableKeyValue` has no
    ``location`` or ``orientation`` variants; it always uses the fixed
    layout with the title in the top-left cell.

    Attributes:
        title: Exact title string to search for in the sheet.
        table_type: Concrete class to instantiate. Defaults to
            :class:`~excel_table.models.table_base.TableKeyValue` itself,
            but may be a domain subclass.
    """

    title: str
    table_type: Type[TableKeyValue] = TableKeyValue