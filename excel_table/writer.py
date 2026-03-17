"""
Excel sheet writer.

Renders :class:`~excel_table.models.table_format.FormattedTable2D`,
:class:`~excel_table.models.table_format.FormattedTable1D`,
:class:`~excel_table.models.table_base.TableKeyValue`, and
:class:`~excel_table.models.chart_format.ChartConfig` instances into an Excel
sheet using xlsxwriter. Tables are arranged in a grid defined by
:class:`SheetWriteSchema`, with configurable spacing between items.

Usage::

    from excel_table import write_sheet
    from excel_table.writer import SheetWriteSchema

    schema = SheetWriteSchema(rows=[
        [formatted_value_map, formatted_x_points],
        [formatted_conditions, chart_config],
    ])
    write_sheet("output.xlsx", "Results", schema)

Border style
------------
Each table is rendered with two border weights:

- **Thick border**: outer perimeter of the table body (excluding title),
  horizontal boundary between ``column_label`` and ``col_values`` rows,
  vertical boundary between ``row_values`` and ``values`` columns.
- **Thin border**: vertical boundary between ``row_label`` and ``row_values``,
  horizontal boundaries between individual ``row_values`` cells,
  vertical boundaries between individual ``col_values`` cells,
  all boundaries between ``values`` cells.
"""
from __future__ import annotations

from pathlib import Path
from typing import Union
import io

import xlsxwriter

from .models.table_base import Table2D, Table1D, TableKeyValue
from .models.table_format import FormattedTable2D, FormattedTable1D
from .models.chart_format import ChartConfig
from .constants import TABLE_MARGIN_ROWS, TABLE_MARGIN_COLS

THICK = 2
THIN = 1

SheetWriteItem = Union[FormattedTable2D, FormattedTable1D, TableKeyValue, ChartConfig]


class SheetWriteSchema:
    """
    Describes the layout of tables and charts to write into a sheet.

    Items in each inner list of ``rows`` are placed left-to-right. Each outer
    list entry starts a new row of tables below the previous one. Spacing
    between items is controlled by :data:`~excel_table.constants.TABLE_MARGIN_COLS`
    and :data:`~excel_table.constants.TABLE_MARGIN_ROWS`.

    Column structure constraint
    ---------------------------
    All rows must have the same number of items, and each column position must
    hold the same item type across all rows. This mirrors the read-side
    :class:`~excel_table.reader.SheetReadSchema` contract where ``columns``
    defines a fixed structure repeated for each row.

    Raises:
        ValueError: At construction time if ``rows`` is empty, if row lengths
            differ, or if any column position has inconsistent item types
            across rows, or if any table title is duplicated within the same
            row.

    Attributes:
        rows: Grid of write items. All rows must have equal length and
            consistent per-column types.
        col_width: Reference column width in pixels, used to compute chart
            cell footprint. Default ``64``.
        row_height: Reference row height in pixels, used to compute chart
            cell footprint. Default ``20``.
        sheet_margin_rows: Number of blank rows between the sheet edge and
            the first row of tables. Default ``1``.
        sheet_margin_cols: Number of blank columns between the sheet edge and
            the first column of tables. Default ``1``.
    """

    def __init__(
        self,
        rows: list[list[SheetWriteItem]],
        col_width: int = 64,
        row_height: int = 20,
        sheet_margin_rows: int = 1,
        sheet_margin_cols: int = 1,
    ):
        if not rows:
            raise ValueError("SheetWriteSchema.rows must not be empty.")

        n_cols = len(rows[0])
        for i, row in enumerate(rows):
            if len(row) != n_cols:
                raise ValueError(
                    f"All rows must have the same number of items. "
                    f"Row 0 has {n_cols} items, row {i} has {len(row)}."
                )

        for col_idx in range(n_cols):
            col_types = [type(row[col_idx]) for row in rows]
            if len(set(col_types)) > 1:
                raise ValueError(
                    f"Column {col_idx} has inconsistent item types across rows: "
                    f"{[t.__name__ for t in col_types]}. "
                    "All rows must use the same type for each column position."
                )

        for row_idx, row in enumerate(rows):
            seen_titles: set[str] = set()
            for item in row:
                if isinstance(item, (FormattedTable2D, FormattedTable1D, TableKeyValue)):
                    if isinstance(item, (FormattedTable2D, FormattedTable1D)):
                        title = item.table.title
                    elif isinstance(item, TableKeyValue):
                        title = item.title
                    else:
                        continue
                    if title in seen_titles:
                        raise ValueError(
                            f"Duplicate table title {title!r} in row {row_idx}. "
                            "Table titles must be unique within the same row."
                        )
                    seen_titles.add(title)

        self.rows = rows
        self.col_width = col_width
        self.row_height = row_height
        self.sheet_margin_rows = sheet_margin_rows
        self.sheet_margin_cols = sheet_margin_cols


# ---------------------------------------------------------------------------
# Helpers: color / format
# ---------------------------------------------------------------------------

def _hex(hex_color: str) -> str:
    """Strip leading ``#`` and uppercase — xlsxwriter format."""
    return hex_color.lstrip("#").upper()


def _apply_conditional_formats(ws, workbook, first_row, first_col, last_row, last_col, fmts):
    """
    Apply a list of xlsxwriter conditional_format dicts to a cell range.

    If a dict contains a ``"cell_format"`` key whose value is a plain ``dict``,
    it is replaced with a ``workbook.add_format(...)`` object stored under the
    key ``"format"`` before the call.

    Args:
        ws: xlsxwriter worksheet.
        workbook: Parent xlsxwriter workbook.
        first_row: 0-indexed top row of the range.
        first_col: 0-indexed left column of the range.
        last_row: 0-indexed bottom row of the range.
        last_col: 0-indexed right column of the range.
        fmts: List of conditional_format parameter dicts.
    """
    for raw in fmts:
        fmt_dict = dict(raw)
        if "cell_format" in fmt_dict and isinstance(fmt_dict["cell_format"], dict):
            fmt_dict["format"] = workbook.add_format(fmt_dict.pop("cell_format"))
        ws.conditional_format(first_row, first_col, last_row, last_col, fmt_dict)


# ---------------------------------------------------------------------------
# Helpers: border format cache
# ---------------------------------------------------------------------------

class _FmtCache:
    """
    Caches xlsxwriter format objects to avoid creating duplicates.

    xlsxwriter has a limit on the number of unique formats. Reusing format
    objects for identical style combinations keeps the count low.
    """

    def __init__(self, workbook):
        self._wb = workbook
        self._cache: dict[tuple, object] = {}

    def get(self, props: dict):
        key = tuple(sorted(props.items()))
        if key not in self._cache:
            self._cache[key] = self._wb.add_format(props)
        return self._cache[key]


# ---------------------------------------------------------------------------
# Border drawing helpers
# ---------------------------------------------------------------------------

def _border_props(
    top: int = 0,
    bottom: int = 0,
    left: int = 0,
    right: int = 0,
    bg_color: str | None = None,
    bold: bool = False,
    align: str | None = None,
    valign: str | None = None,
    rotation: int | None = None,
) -> dict:
    props: dict = {
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
    }
    if bg_color:
        props["bg_color"] = _hex(bg_color)
    if bold:
        props["bold"] = True
    if align:
        props["align"] = align
    if valign:
        props["valign"] = valign
    if rotation is not None:
        props["rotation"] = rotation
    return props


# ---------------------------------------------------------------------------
# Table2D writer
# ---------------------------------------------------------------------------

def _write_table2d(
    ws,
    workbook,
    fmt: FormattedTable2D,
    origin_row: int,
    origin_col: int,
    cache: _FmtCache,
) -> tuple[int, int]:
    """
    Write a :class:`~excel_table.models.table_format.FormattedTable2D`.

    Layout variants are controlled by ``fmt.column_location`` and
    ``fmt.row_location``. All coordinates below are 0-indexed offsets from
    ``(origin_row, origin_col)``.

    **top / left** (default)::

        (0,0) title
        (1,0) [2×2 blank]         (1,2) column_label [merged n_c cols]
        (2,0)                     (2,2) col1  col2  ...
        (3,0) row_label [↓ n_r]   (3,1) row1  val   ...
        (4,0)                     (4,1) row2  val   ...

    **top / right**::

        (0,0) title
        (1,0) column_label [merged n_c cols]   (1,n_c+1) [2×2 blank]
        (2,0) col1  col2  ...                  (2,n_c+1)
        (3,0) row1  val   ...   (3,n_c-1)      (3,n_c+1) row_label [↓ n_r]
        (4,0) row2  val   ...

    **bottom / left**::

        (0,0) row_label [↓ n_r]   (0,1) row1  val   ...
        (1,0)                     (1,1) row2  val   ...
        (n_r,0)                   (n_r,2) col1  col2  ...
        (n_r+1,0) [2×2 blank]    (n_r+1,2) column_label [merged n_c cols]
        (n_r+2,0) title

    **bottom / right**::

        (0,0) row1  val   ...   (0,n_c-1)      (0,n_c+1) row_label [↓ n_r]
        (1,0) row2  val   ...
        (n_r,0) col1  col2  ...
        (n_r+1,0) column_label [merged n_c cols]   (n_r+1,n_c+1) [2×2 blank]
        (n_r+2,0) title                            (n_r+2,n_c+1)

    Args:
        ws: xlsxwriter worksheet.
        workbook: Parent xlsxwriter workbook.
        fmt: Formatted table to render.
        origin_row: Top-left row (0-indexed).
        origin_col: Top-left column (0-indexed).
        cache: Format object cache.

    Returns:
        ``(used_rows, used_cols)`` cell footprint excluding margin.
    """
    table = fmt.table
    r, c = origin_row, origin_col
    n_c = len(table.column)
    n_r = len(table.row)
    col_loc = fmt.column_location
    row_loc = fmt.row_location

    col_label_rot = 90 if fmt.column_label_direction == "vertical" else None
    row_label_rot = 90 if fmt.row_label_direction == "vertical" else None

    # ------------------------------------------------------------------
    # Compute anchor offsets for each region (0-indexed from origin)
    # ------------------------------------------------------------------
    if col_loc == "top" and row_loc == "left":
        title_r, title_c = 0, 0
        blank_r, blank_c = 1, 0          # 2×2 blank
        col_label_r, col_label_c = 1, 2
        col_vals_r, col_vals_c = 2, 2
        row_label_r, row_label_c = 3, 0
        row_vals_r, row_vals_c = 3, 1
        data_r, data_c = 3, 2
        used_rows = 3 + n_r
        used_cols = 2 + n_c

    elif col_loc == "top" and row_loc == "right":
        title_r, title_c = 0, 0
        blank_r, blank_c = 1, n_c + 1   # 2×2 blank
        col_label_r, col_label_c = 1, 0
        col_vals_r, col_vals_c = 2, 0
        row_label_r, row_label_c = 3, n_c + 1
        row_vals_r, row_vals_c = 3, n_c
        data_r, data_c = 3, 0
        used_rows = 3 + n_r
        used_cols = n_c + 2

    elif col_loc == "bottom" and row_loc == "left":
        title_r, title_c = n_r + 2, 0
        blank_r, blank_c = n_r + 1, 0   # 2×2 blank
        col_label_r, col_label_c = n_r + 1, 2
        col_vals_r, col_vals_c = n_r, 2
        row_label_r, row_label_c = 0, 0
        row_vals_r, row_vals_c = 0, 1
        data_r, data_c = 0, 2
        used_rows = n_r + 3
        used_cols = 2 + n_c

    else:  # bottom / right
        title_r, title_c = n_r + 2, 0
        blank_r, blank_c = n_r + 1, n_c + 1  # 2×2 blank
        col_label_r, col_label_c = n_r + 1, 0
        col_vals_r, col_vals_c = n_r, 0
        row_label_r, row_label_c = 0, n_c + 1
        row_vals_r, row_vals_c = 0, n_c
        data_r, data_c = 0, 0
        used_rows = n_r + 3
        used_cols = n_c + 2

    # ------------------------------------------------------------------
    # Write title
    # ------------------------------------------------------------------
    ws.write(r + title_r, c + title_c, table.title,
             cache.get({"bold": True, "align": "left", "valign": "vcenter"}))

    # ------------------------------------------------------------------
    # Write 2×2 blank merge
    # ------------------------------------------------------------------
    blank_props = _border_props(top=THICK, bottom=THICK, left=THICK, right=THICK)
    ws.merge_range(r + blank_r, c + blank_c,
                   r + blank_r + 1, c + blank_c + 1, None, cache.get(blank_props))

    # ------------------------------------------------------------------
    # Write column_label (merged across n_c cols)
    # ------------------------------------------------------------------
    col_label_props = _border_props(
        top=THICK, bottom=THICK, left=THICK, right=THICK,
        bg_color=fmt.column_color, bold=True, align="center", valign="vcenter",
        rotation=col_label_rot,
    )
    ws.merge_range(
        r + col_label_r, c + col_label_c,
        r + col_label_r, c + col_label_c + n_c - 1,
        table.column_label, cache.get(col_label_props),
    )

    # ------------------------------------------------------------------
    # Write col_values
    # ------------------------------------------------------------------
    for ci, col_val in enumerate(table.column):
        is_first = ci == 0
        is_last = ci == n_c - 1
        if col_loc == "top":
            top_w = THICK   # border with column_label above
            bottom_w = THICK  # border with data below
        else:
            top_w = THICK   # border with data above
            bottom_w = THICK  # border with column_label below
        props = _border_props(
            top=top_w, bottom=bottom_w,
            left=THICK if is_first else THIN,
            right=THICK if is_last else THIN,
            bg_color=fmt.column_color, bold=True, align="center",
        )
        ws.write(r + col_vals_r, c + col_vals_c + ci, col_val, cache.get(props))

    # ------------------------------------------------------------------
    # Write row_label (merged down n_r rows)
    # ------------------------------------------------------------------
    row_label_props = _border_props(
        top=THICK, bottom=THICK, left=THICK, right=THICK,
        bg_color=fmt.row_color, bold=True, align="center", valign="vcenter",
        rotation=row_label_rot,
    )
    ws.merge_range(
        r + row_label_r, c + row_label_c,
        r + row_label_r + n_r - 1, c + row_label_c,
        table.row_label, cache.get(row_label_props),
    )

    # ------------------------------------------------------------------
    # Write row_values and data
    # ------------------------------------------------------------------
    for ri in range(n_r):
        is_first_row = ri == 0
        is_last_row = ri == n_r - 1

        # row_value cell
        rv_props = _border_props(
            top=THICK if is_first_row else THIN,
            bottom=THICK if is_last_row else THIN,
            left=THIN,   # thin border with row_label
            right=THICK,  # thick border with data
            bg_color=fmt.row_color, bold=True, align="center",
        )
        ws.write(r + row_vals_r + ri, c + row_vals_c, table.row[ri], cache.get(rv_props))

        # data cells
        for ci in range(n_c):
            is_first_col = ci == 0
            is_last_col = ci == n_c - 1
            d_props = _border_props(
                top=THICK if is_first_row else THIN,
                bottom=THICK if is_last_row else THIN,
                left=THICK if is_first_col else THIN,
                right=THICK if is_last_col else THIN,
            )
            val = table.values[ri][ci]
            if val is None:
                ws.write_blank(r + data_r + ri, c + data_c + ci, None, cache.get(d_props))
            else:
                ws.write(r + data_r + ri, c + data_c + ci, val, cache.get(d_props))

    # ------------------------------------------------------------------
    # Apply conditional formats to data range
    # ------------------------------------------------------------------
    if fmt.value_conditional_formats:
        _apply_conditional_formats(
            ws, workbook,
            r + data_r, c + data_c,
            r + data_r + n_r - 1, c + data_c + n_c - 1,
            fmt.value_conditional_formats,
        )

    return used_rows, used_cols


# ---------------------------------------------------------------------------
# Table1D writer
# ---------------------------------------------------------------------------

def _write_table1d(
    ws,
    workbook,
    fmt: FormattedTable1D,
    origin_row: int,
    origin_col: int,
    cache: _FmtCache,
) -> tuple[int, int]:
    """
    Write a :class:`~excel_table.models.table_format.FormattedTable1D`.

    **horizontal** (default)::

        (0,0) title
        (1,0) column_label [merged n_c cols]
        (2,0) col1  col2  ...
        (3,0) val1  val2  ...

    **vertical**::

        (0,0) title
        (1,0) column_label [merged n_c rows]  (1,1) col1  (1,2) val
        (2,0)                                  (2,1) col2  (2,2) val
        ...

    Args:
        ws: xlsxwriter worksheet.
        workbook: Parent xlsxwriter workbook.
        fmt: Formatted table to render.
        origin_row: Top-left row (0-indexed).
        origin_col: Top-left column (0-indexed).
        cache: Format object cache.

    Returns:
        ``(used_rows, used_cols)`` cell footprint excluding margin.
    """
    table = fmt.table
    r, c = origin_row, origin_col
    n_c = len(table.column)
    col_label_rot = 90 if fmt.column_label_direction == "vertical" else None

    # title
    ws.write(r, c, table.title,
             cache.get({"bold": True, "align": "left", "valign": "vcenter"}))

    if fmt.orientation == "horizontal":
        # column_label merged across n_c cols
        cl_props = _border_props(
            top=THICK, bottom=THICK, left=THICK, right=THICK,
            bg_color=fmt.column_color, bold=True, align="left", valign="vcenter",
            rotation=col_label_rot,
        )
        ws.merge_range(r + 1, c, r + 1, c + n_c - 1, table.column_label, cache.get(cl_props))

        # col_values
        for ci, col_val in enumerate(table.column):
            props = _border_props(
                top=THICK, bottom=THICK,
                left=THICK if ci == 0 else THIN,
                right=THICK if ci == n_c - 1 else THIN,
                bg_color=fmt.column_color, bold=True, align="center",
            )
            ws.write(r + 2, c + ci, col_val, cache.get(props))

        # data values (single row)
        vals = table.values[0]
        for ci, val in enumerate(vals):
            props = _border_props(
                top=THICK, bottom=THICK,
                left=THICK if ci == 0 else THIN,
                right=THICK if ci == n_c - 1 else THIN,
            )
            if val is None:
                ws.write_blank(r + 3, c + ci, None, cache.get(props))
            else:
                ws.write(r + 3, c + ci, val, cache.get(props))

        if fmt.value_conditional_formats:
            _apply_conditional_formats(
                ws, workbook, r + 3, c, r + 3, c + n_c - 1,
                fmt.value_conditional_formats,
            )

        return 4, n_c

    else:  # vertical
        # column_label merged down n_c rows
        cl_props = _border_props(
            top=THICK, bottom=THICK, left=THICK, right=THICK,
            bg_color=fmt.column_color, bold=True, align="center", valign="vcenter",
            rotation=col_label_rot,
        )
        ws.merge_range(r + 1, c, r + n_c, c, table.column_label, cache.get(cl_props))

        for ci in range(n_c):
            is_first = ci == 0
            is_last = ci == n_c - 1

            # col_value
            cv_props = _border_props(
                top=THICK if is_first else THIN,
                bottom=THICK if is_last else THIN,
                left=THICK,   # border with column_label
                right=THICK,  # border with data
                bg_color=fmt.column_color, bold=True, align="center",
            )
            ws.write(r + 1 + ci, c + 1, table.column[ci], cache.get(cv_props))

            # data value (single col)
            val = table.values[ci][0]
            d_props = _border_props(
                top=THICK if is_first else THIN,
                bottom=THICK if is_last else THIN,
                left=THICK, right=THICK,
            )
            if val is None:
                ws.write_blank(r + 1 + ci, c + 2, None, cache.get(d_props))
            else:
                ws.write(r + 1 + ci, c + 2, val, cache.get(d_props))

        if fmt.value_conditional_formats:
            _apply_conditional_formats(
                ws, workbook, r + 1, c + 2, r + n_c, c + 2,
                fmt.value_conditional_formats,
            )

        return 1 + n_c, 3


# ---------------------------------------------------------------------------
# TableKeyValue writer
# ---------------------------------------------------------------------------

def _write_table_key_value(
    ws,
    workbook,
    table: TableKeyValue,
    origin_row: int,
    origin_col: int,
    cache: _FmtCache,
) -> tuple[int, int]:
    """
    Write a :class:`~excel_table.models.table_base.TableKeyValue`.

    Layout::

        (0,0) title
        (1,0) key1  key2  ...
        (2,0) val1  val2  ...

    Args:
        ws: xlsxwriter worksheet.
        workbook: Parent xlsxwriter workbook.
        table: TableKeyValue instance.
        origin_row: Top-left row (0-indexed).
        origin_col: Top-left column (0-indexed).
        cache: Format object cache.

    Returns:
        ``(used_rows, used_cols)`` cell footprint excluding margin.
    """
    r, c = origin_row, origin_col
    n = len(table.column)

    ws.write(r, c, table.title,
             cache.get({"bold": True, "align": "left", "valign": "vcenter"}))

    for i, key in enumerate(table.column):
        props = _border_props(
            top=THICK, bottom=THIN,
            left=THICK if i == 0 else THIN,
            right=THICK if i == n - 1 else THIN,
            bold=True, align="center",
        )
        ws.write(r + 1, c + i, key, cache.get(props))

    for i, val in enumerate(table.value):
        props = _border_props(
            top=THIN, bottom=THICK,
            left=THICK if i == 0 else THIN,
            right=THICK if i == n - 1 else THIN,
        )
        if val is None:
            ws.write_blank(r + 2, c + i, None, cache.get(props))
        elif isinstance(val, str):
            ws.write_string(r + 2, c + i, val, cache.get(props))
        elif isinstance(val, float) and val == int(val):
            ws.write_number(r + 2, c + i, int(val), cache.get(props))
        else:
            ws.write_number(r + 2, c + i, val, cache.get(props))

    return 3, n


# ---------------------------------------------------------------------------
# Footprint calculation (pass 1)
# ---------------------------------------------------------------------------

def _calc_footprint(item: SheetWriteItem, col_width: int, row_height: int) -> tuple[int, int]:
    """Return ``(used_rows, used_cols)`` for *item* without writing anything."""
    if isinstance(item, FormattedTable2D):
        return 3 + len(item.table.row), 2 + len(item.table.column)
    elif isinstance(item, FormattedTable1D):
        n_c = len(item.table.column)
        return (4, n_c) if item.orientation == "horizontal" else (1 + n_c, 3)
    elif isinstance(item, TableKeyValue):
        return 3, len(item.column)
    elif isinstance(item, ChartConfig):
        used_cols, used_rows = item.occupied_cells(col_width, row_height)
        return used_rows, used_cols
    else:
        raise ValueError(f"Unknown item type: {type(item)}")


def _build_grid(schema: SheetWriteSchema):
    """Pass 1: compute origins for every item and build per-row table_origins maps."""
    origins = []
    row_table_origins: list[dict[str, tuple[int, int]]] = []
    current_row = schema.sheet_margin_rows
    for row_items in schema.rows:
        current_col = schema.sheet_margin_cols
        row_height_used = 0
        row_origins = []
        this_row_origins: dict[str, tuple[int, int]] = {}
        for item in row_items:
            row_origins.append((current_row, current_col))
            if isinstance(item, FormattedTable2D):
                this_row_origins[item.table.title] = (current_row, current_col)
            used_rows, used_cols = _calc_footprint(item, schema.col_width, schema.row_height)
            current_col += used_cols + TABLE_MARGIN_COLS
            row_height_used = max(row_height_used, used_rows)
        origins.append(row_origins)
        row_table_origins.append(this_row_origins)
        current_row += row_height_used + TABLE_MARGIN_ROWS
    return origins, row_table_origins


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _write_to_worksheet(workbook, ws, schema: SheetWriteSchema, sheet_name: str) -> None:
    """
    Render all items in *schema* into *ws*.

    Extracted from :func:`write_sheet` so that single-sheet and
    multi-sheet paths can share the same rendering logic.
    Uses the same two-pass architecture as :func:`write_sheet`:
    :func:`_build_grid` computes coordinates first, then items are rendered
    in order with :class:`ChartConfig` delegated to
    :func:`~excel_table.chart.render_chart`.

    Args:
        workbook: xlsxwriter workbook that owns *ws*.
        ws: xlsxwriter worksheet to write into.
        schema: Layout and content definition.
        sheet_name: Worksheet name, passed through to
            :func:`~excel_table.chart.render_chart`.
    """
    from .chart import render_chart

    origins, row_table_origins = _build_grid(schema)
    cache = _FmtCache(workbook)

    for row_idx, row_items in enumerate(schema.rows):
        for col_idx, item in enumerate(row_items):
            origin_row, origin_col = origins[row_idx][col_idx]
            if isinstance(item, FormattedTable2D):
                _write_table2d(ws, workbook, item, origin_row, origin_col, cache)
            elif isinstance(item, FormattedTable1D):
                _write_table1d(ws, workbook, item, origin_row, origin_col, cache)
            elif isinstance(item, TableKeyValue):
                _write_table_key_value(ws, workbook, item, origin_row, origin_col, cache)
            elif isinstance(item, ChartConfig):
                row_origins = row_table_origins[row_idx]
                row_fmt_tables = [
                    ft
                    for col_idx2, ft in enumerate(row_items)
                    if isinstance(ft, FormattedTable2D)
                ]
                render_chart(
                    ws=ws,
                    workbook=workbook,
                    config=item,
                    tables=row_fmt_tables,
                    origin_row=origin_row,
                    origin_col=origin_col,
                    schema=schema,
                    table_origins=row_origins,
                    sheet_name=sheet_name,
                )
            else:
                raise ValueError(f"Unknown item type: {type(item)}")


def write_sheet(path: str | Path, sheet_name: str, schema: SheetWriteSchema) -> None:
    """
    Write tables and charts to a new Excel file.

    Uses two passes:

    1. **Pass 1** (:func:`_build_grid`): compute grid coordinates for every
       item without touching xlsxwriter.
    2. **Pass 2**: render each item. :class:`ChartConfig` items are rendered
       via :func:`~excel_table.chart.render_chart` using the pre-computed
       ``table_origins``, so charts may appear before or after their source
       tables in the grid.

    Args:
        path: Destination ``.xlsx`` file path. Created or overwritten.
        sheet_name: Name of the worksheet to create.
        schema: Layout and content definition.
    """
    workbook = xlsxwriter.Workbook(str(path))
    ws = workbook.add_worksheet(sheet_name)
    _write_to_worksheet(workbook, ws, schema, sheet_name)
    workbook.close()


def write_sheet_bytes(sheet_name: str, schema: SheetWriteSchema) -> bytes:
    """
    Write tables and charts to an in-memory Excel file and return the raw bytes.

    Equivalent to :func:`write_sheet` but writes to a ``BytesIO`` buffer
    instead of disk. Useful for Streamlit ``st.download_button`` or other
    contexts where a file path is not available.

    Args:
        sheet_name: Name of the worksheet to create.
        schema: Layout and content definition.

    Returns:
        Raw ``.xlsx`` file contents as :class:`bytes`.
    """
    buf = io.BytesIO()
    workbook = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = workbook.add_worksheet(sheet_name)
    _write_to_worksheet(workbook, ws, schema, sheet_name)
    workbook.close()
    return buf.getvalue()


def write_workbook(sheets: list[tuple[str, SheetWriteSchema]]) -> bytes:
    """
    Write multiple sheets into a single in-memory Excel workbook.

    Each entry in *sheets* becomes one worksheet. Sheets are created in
    the order they appear in the list.

    Args:
        sheets: Sequence of ``(sheet_name, schema)`` pairs.

    Returns:
        Raw ``.xlsx`` file contents as :class:`bytes`.

    Example::

        data = write_workbook([
            ("Results", results_schema),
            ("Conditions", conditions_schema),
        ])
        with open("output.xlsx", "wb") as f:
            f.write(data)
    """
    buf = io.BytesIO()
    workbook = xlsxwriter.Workbook(buf, {"in_memory": True})
    for sheet_name, schema in sheets:
        ws = workbook.add_worksheet(sheet_name)
        _write_to_worksheet(workbook, ws, schema, sheet_name)
    workbook.close()
    return buf.getvalue()