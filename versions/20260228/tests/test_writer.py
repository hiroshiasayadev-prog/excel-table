"""
Excel sheet writer.

Renders :class:`FormattedTable2D` and :class:`ChartConfig` instances into
an Excel sheet using xlsxwriter. Tables are arranged in a grid defined by
:class:`SheetWriteSchema`, with configurable spacing between cells.

Usage::

    schema = SheetWriteSchema(rows=[
        [formatted_loss_map, chart_config],
    ])
    write_sheet("output.xlsx", "Results", schema)
"""
from __future__ import annotations
import math
from pathlib import Path
from typing import Union

import xlsxwriter

from excel_table.models.format import FormattedTable2D, ChartConfig, ColorScale
from excel_table.models.table import Table2D, Table1D
from excel_table.constants import TABLE_MARGIN_ROWS, TABLE_MARGIN_COLS


SheetWriteItem = Union[FormattedTable2D, ChartConfig]


class SheetWriteSchema:
    """
    Describes the layout of tables and charts to write into a sheet.

    Items in each inner list of ``rows`` are placed left-to-right. Each outer
    list entry starts a new row of tables below the previous one. Spacing
    between items is controlled by :data:`TABLE_MARGIN_COLS` and
    :data:`TABLE_MARGIN_ROWS`.

    Attributes:
        rows: Grid of :class:`FormattedTable2D` or :class:`ChartConfig` items.
        col_width: Reference column width in pixels, used to compute chart cell footprint.
        row_height: Reference row height in pixels, used to compute chart cell footprint.
    """

    def __init__(
        self,
        rows: list[list[SheetWriteItem]],
        col_width: int = 64,
        row_height: int = 20,
    ):
        self.rows = rows
        self.col_width = col_width
        self.row_height = row_height


def _hex(hex_color: str) -> str:
    """
    Normalize a hex color string for xlsxwriter (strips leading ``#``).

    Args:
        hex_color: Hex color with or without leading ``#``.

    Returns:
        Uppercase hex string without ``#``.
    """
    return hex_color.lstrip("#").upper()


def _interpolate_color(color_min: str, color_max: str, t: float) -> str:
    """
    Linearly interpolate between two hex colors.

    Args:
        color_min: Hex color at ``t=0``.
        color_max: Hex color at ``t=1``.
        t: Interpolation factor in ``[0, 1]``.

    Returns:
        Interpolated hex color string (with leading ``#``).
    """
    def parse(h: str) -> tuple[int, int, int]:
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    r1, g1, b1 = parse(color_min)
    r2, g2, b2 = parse(color_max)
    r = int(r1 + (r2 - r1) * t)
    g = int(g1 + (g2 - g1) * t)
    b = int(b1 + (b2 - b1) * t)
    return f"#{r:02X}{g:02X}{b:02X}"


def _colorscale_color(scale: ColorScale, t: float) -> str:
    """
    Map a normalized value to a hex color using a :class:`ColorScale`.

    For 2-color scales, interpolates directly between ``min_color`` and
    ``max_color``. For 3-color scales, the midpoint color is at ``t=0.5``.

    Args:
        scale: The color scale definition.
        t: Normalized value in ``[0, 1]``.

    Returns:
        Hex color string (with leading ``#``).
    """
    if scale.mid_color is None:
        return _interpolate_color(scale.min_color, scale.max_color, t)
    if t <= 0.5:
        return _interpolate_color(scale.min_color, scale.mid_color, t * 2)
    return _interpolate_color(scale.mid_color, scale.max_color, (t - 0.5) * 2)


def _write_table2d(
    ws,
    workbook: xlsxwriter.Workbook,
    fmt: FormattedTable2D,
    origin_row: int,
    origin_col: int,
) -> tuple[int, int]:
    """
    Write a single :class:`FormattedTable2D` to the worksheet.

    Cell structure written (0-indexed offsets from origin)::

        (0,0) title [single cell]
        (1,0) [2×2 merged blank]            (1,2) column_label [merged across cols]
        (2,0)                               (2,2) col1  (2,3) col2 ...
        (3,0) row_label [merged down]       (3,1) row1  (3,2) val  ...
        (4,0)                               (4,1) row2  (4,2) val  ...

    Data cells are colored according to ``fmt.value_colorscale``, normalized
    over the min/max of all non-``None`` values in ``fmt.table.values``.

    Args:
        ws: xlsxwriter worksheet object.
        workbook: Parent xlsxwriter workbook (used to create formats).
        fmt: Formatted table to render.
        origin_row: Top-left row index (0-indexed).
        origin_col: Top-left column index (0-indexed).

    Returns:
        ``(used_rows, used_cols)`` — cell footprint of the written table,
        excluding margin.
    """
    table = fmt.table
    r, c = origin_row, origin_col
    n_col = len(table.column)
    n_row = len(table.row)

    col_fmt = workbook.add_format({
        "bg_color": _hex(fmt.column_color),
        "bold": True,
        "align": "center",
    })
    row_fmt_props: dict = {
        "bg_color": _hex(fmt.row_color),
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    }
    if fmt.row_label_direction == "vertical":
        row_fmt_props["rotation"] = 90
    row_fmt = workbook.add_format(row_fmt_props)
    header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})

    # Layout (0-indexed offsets from origin):
    #   (0,0) title [single cell]
    #   (1,0) [2×2 merged blank]          (1,2) column_label [merged across cols]
    #   (2,0)                              (2,2) col1  (2,3) col2 ...
    #   (3,0) row_label [merged down]      (3,1) row1  (3,2) val  ...
    #   (4,0)                              (4,1) row2  (4,2) val  ...

    # title (single cell, row 0)
    ws.write(r, c, table.title, header_fmt)

    # 2×2 blank merged block below title
    ws.merge_range(r + 1, c, r + 2, c + 1, None)

    # column_label (merged across columns, row 1)
    ws.merge_range(r + 1, c + 2, r + 1, c + 1 + n_col, table.column_label, col_fmt)

    # column header values (row 2)
    for ci, col_val in enumerate(table.column):
        ws.write(r + 2, c + 2 + ci, col_val, col_fmt)

    # row_label (merged down, starting row 3)
    ws.merge_range(r + 3, c, r + 2 + n_row, c, table.row_label, row_fmt)

    # normalize values for colorscale
    scale = fmt.value_colorscale
    flat = [v for row in table.values for v in row if v is not None]
    v_min = min(flat) if flat else 0.0
    v_max = max(flat) if flat else 1.0
    v_range = v_max - v_min or 1.0

    # row header values and data cells (starting row 3)
    for ri, row_val in enumerate(table.row):
        ws.write(r + 3 + ri, c + 1, row_val, row_fmt)
        for ci, cell_val in enumerate(table.values[ri]):
            if cell_val is None:
                ws.write_blank(r + 3 + ri, c + 2 + ci, None)
            else:
                t = (cell_val - v_min) / v_range
                color = _colorscale_color(scale, t)
                cell_fmt = workbook.add_format({"bg_color": _hex(color)})
                ws.write(r + 3 + ri, c + 2 + ci, cell_val, cell_fmt)

    used_rows = 3 + n_row  # title row + 2 header rows + data rows
    used_cols = 2 + n_col  # row_label col + row_val col + data cols
    return used_rows, used_cols


def write_sheet(
    path: str | Path,
    sheet_name: str,
    schema: SheetWriteSchema,
):
    """
    Write tables and charts to a new Excel file.

    Items within each row of ``schema.rows`` are placed left-to-right, separated
    by :data:`TABLE_MARGIN_COLS` blank columns. Rows of items are stacked
    top-to-bottom, separated by :data:`TABLE_MARGIN_ROWS` blank rows. The row
    height is determined by the tallest item in each row.

    :class:`ChartConfig` items are fully rendered using
    :func:`~excel_table.chart.render_chart`. Each chart's ``source_block``
    references must resolve to a unique :class:`FormattedTable2D` title that
    appears in the same schema.

    Args:
        path: Destination ``.xlsx`` file path. Created or overwritten.
        sheet_name: Name of the worksheet to create.
        schema: Layout and content definition.

    Raises:
        ValueError: If a chart's ``source_block`` title is not found in the
            schema, is duplicated, or if any filter expression is invalid.
    """
    from .chart import render_chart

    workbook = xlsxwriter.Workbook(str(path))
    ws = workbook.add_worksheet(sheet_name)

    # Collect all FormattedTable2D items and compute their origins in one pass.
    # We need origins before rendering charts, so we do a dry-run first.
    all_formatted_tables: list[FormattedTable2D] = []
    table_origins: dict[str, tuple[int, int]] = {}

    dry_row = 0
    for row_items in schema.rows:
        dry_col = 0
        row_height_used = 0
        for item in row_items:
            if isinstance(item, FormattedTable2D):
                all_formatted_tables.append(item)
                table_origins[item.table.title] = (dry_row, dry_col)
                n_col = len(item.table.column)
                n_row = len(item.table.row)
                used_rows = 3 + n_row
                used_cols = 2 + n_col
                row_height_used = max(row_height_used, used_rows)
                dry_col += used_cols + TABLE_MARGIN_COLS
            elif isinstance(item, ChartConfig):
                cols, rows = item.occupied_cells(schema.col_width, schema.row_height)
                row_height_used = max(row_height_used, rows)
                dry_col += cols + TABLE_MARGIN_COLS
        dry_row += row_height_used + TABLE_MARGIN_ROWS

    # Actual rendering pass
    current_row = 0
    for row_items in schema.rows:
        current_col = 0
        row_height_used = 0

        for item in row_items:
            if isinstance(item, FormattedTable2D):
                used_rows, used_cols = _write_table2d(
                    ws, workbook, item, current_row, current_col
                )
                current_col += used_cols + TABLE_MARGIN_COLS
                row_height_used = max(row_height_used, used_rows)

            elif isinstance(item, ChartConfig):
                used_cols, used_rows = render_chart(
                    ws=ws,
                    workbook=workbook,
                    config=item,
                    tables=all_formatted_tables,
                    origin_row=current_row,
                    origin_col=current_col,
                    schema=schema,
                    table_origins=table_origins,
                    sheet_name=sheet_name,
                )
                current_col += used_cols + TABLE_MARGIN_COLS
                row_height_used = max(row_height_used, used_rows)

        current_row += row_height_used + TABLE_MARGIN_ROWS

    workbook.close()