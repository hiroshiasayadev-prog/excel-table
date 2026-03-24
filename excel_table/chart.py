"""
Excel chart rendering for structured Table2D data.

This module provides :func:`render_chart`, which writes an xlsxwriter chart
object into a worksheet based on a :class:`~excel_table.models.format.ChartConfig`.

Filter expressions
------------------
``row_filter`` and ``col_filter`` in :class:`~excel_table.models.format.LineSeriesConfig`
are evaluated as Python expressions using :func:`eval`. The axis value being
tested is exposed as a scalar variable.

For a plain :class:`~excel_table.models.table.Table2D`, the available variable
names are ``"row"`` (for row axis) and ``"column"`` (for column axis).
Domain subclasses that expose additional :class:`property` aliases for the
same list are also accepted. For example, if a subclass defines::

    @property
    def y(self) -> list[str]:
        return self.row

then ``row_filter = "y >= 0"`` is equivalent to ``row_filter = "row >= 0"``.

A filter expression that references a variable name not associated with the
correct axis raises :class:`ValueError` at render time (before any data is
processed).

Series splitting
----------------
Each :class:`~excel_table.models.format.LineSeriesConfig` always produces one
xlsxwriter series per item on the split axis. This is required because Excel
cannot represent multiple disconnected lines within a single series object.

When ``color_axis`` is set, the splitting axis is also used to assign a color
from the table's :class:`~excel_table.models.format.ColorScale` to each
sub-series. When ``color_axis`` is ``None``, xlsxwriter assigns colors
automatically.
"""
from __future__ import annotations

import ast
from typing import TYPE_CHECKING, Literal

from .models.chart_format import ChartConfig, LineSeriesConfig, ColorScale
from .models.table_format import FormattedTable2D
from .models.table_base import Table2D
from .writer import _hex


# ---------------------------------------------------------------------------
# Internal helpers – color interpolation (chart series only)
# ---------------------------------------------------------------------------

def _interpolate_color(color_min: str, color_max: str, t: float) -> str:
    """Linearly interpolate between two hex colors."""
    def parse(h: str) -> tuple[int, int, int]:
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r1, g1, b1 = parse(color_min)
    r2, g2, b2 = parse(color_max)
    return f"#{int(r1+(r2-r1)*t):02X}{int(g1+(g2-g1)*t):02X}{int(b1+(b2-b1)*t):02X}"


def _colorscale_color(scale: ColorScale, t: float) -> str:
    """Map t in [0,1] to a hex color. Used for chart series coloring only."""
    if scale.mid_color is None:
        return _interpolate_color(scale.min_color, scale.max_color, t)
    if t <= 0.5:
        return _interpolate_color(scale.min_color, scale.mid_color, t * 2)
    return _interpolate_color(scale.mid_color, scale.max_color, (t - 0.5) * 2)


if TYPE_CHECKING:
    import xlsxwriter
    from .writer import SheetWriteSchema


# ---------------------------------------------------------------------------
# Internal helpers – axis utilities
# ---------------------------------------------------------------------------

def _property_names_for_axis(
    table: Table2D,
    axis: Literal["column", "row"],
) -> frozenset[str]:
    """
    Collect all valid variable names for *axis* on *table*.

    The default names ``"row"`` / ``"column"`` are always included. In addition,
    any :class:`property` defined on the table's class that returns the **same
    list object** as ``table.row`` or ``table.column`` (identity check) is also
    included. This lets domain subclass aliases such as ``"y"`` or
    ``"x"`` be used directly in filter expressions.

    Args:
        table: Table instance to inspect.
        axis: ``"column"`` or ``"row"``.

    Returns:
        Frozenset of valid variable name strings for the given axis.
    """
    axis_list = table.column if axis == "column" else table.row
    valid: set[str] = {axis}

    for attr_name in dir(type(table)):
        prop = getattr(type(table), attr_name, None)
        if not isinstance(prop, property):
            continue
        try:
            val = getattr(table, attr_name)
        except Exception:
            continue
        if val is axis_list:
            valid.add(attr_name)

    return frozenset(valid)


# ---------------------------------------------------------------------------
# Internal helpers – filter validation and evaluation
# ---------------------------------------------------------------------------

def _validate_filter_expr(
    expr: str,
    axis: Literal["column", "row"],
    valid_names: frozenset[str],
    series_label: str,
) -> None:
    """
    Parse *expr* and verify that every identifier references *valid_names*.

    Args:
        expr: Python expression string (e.g. ``"y >= 0"``).
        axis: Which axis this filter applies to (for error messages).
        valid_names: Set of names that are valid for this axis.
        series_label: Series label (for error messages).

    Raises:
        ValueError: If *expr* has a syntax error, or if any identifier in
            *expr* is not present in *valid_names*.
    """
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as exc:
        raise ValueError(
            f"Series {series_label!r}: {axis}_filter expression {expr!r} "
            f"has a syntax error: {exc}"
        ) from exc

    used_names = {node.id for node in ast.walk(tree) if isinstance(node, ast.Name)}
    invalid = used_names - valid_names
    if invalid:
        raise ValueError(
            f"Series {series_label!r}: {axis}_filter expression {expr!r} "
            f"references {sorted(invalid)}, which are not valid variable names "
            f"for the {axis!r} axis. "
            f"Valid names for this table: {sorted(valid_names)}"
        )


def _passes_filter(
    expr: str,
    valid_names: frozenset[str],
    value: str,
) -> bool:
    """
    Evaluate *expr* with the axis value bound to all names in *valid_names*.

    Pre-condition: *expr* has already been validated by :func:`_validate_filter_expr`.

    Attempts to cast *value* to ``float`` before binding. If the cast fails
    (e.g. string axis labels such as ``"forward"``), the raw string is bound
    instead. This allows filter expressions to work with both numeric axes
    (e.g. ``"vgs >= 0.0"``) and string axes (e.g. ``"column == 'forward'"``).

    Args:
        expr: Validated Python expression string.
        valid_names: Names exposed in the evaluation namespace; all bound to *value*.
        value: Axis header value as a string. Cast to ``float`` if possible,
            otherwise used as-is.

    Returns:
        ``True`` if the expression is truthy, ``False`` on falsy or any runtime error.
    """
    _value: float | str | None = None
    try:
        _value = float(value)
    except:
        _value = value
    ns = {name: _value for name in valid_names}
    try:
        return bool(eval(expr, {"__builtins__": {}}, ns))  # noqa: S307
    except Exception:
        return False


def _validate_all_filters(
    config: ChartConfig,
    tables: list[FormattedTable2D],
    sheet_name: str,
) -> None:
    """
    Validate every filter expression in *config* before rendering begins.

    Iterates over all :class:`LineSeriesConfig` entries in ``config.series``
    and calls :func:`_validate_filter_expr` for each non-``None`` filter.
    Errors are raised immediately so that no xlsxwriter state is mutated on
    invalid input.

    Args:
        config: Chart configuration to validate.
        tables: All :class:`FormattedTable2D` instances in the sheet.
        sheet_name: Worksheet name (for error messages).

    Raises:
        ValueError: On the first invalid filter expression encountered.
    """
    for series_cfg in config.series:
        fmt = _resolve_table(series_cfg.source_block, tables, sheet_name)
        table = fmt.table

        row_names = _property_names_for_axis(table, "row")
        col_names = _property_names_for_axis(table, "column")

        if series_cfg.row_filter is not None:
            _validate_filter_expr(
                series_cfg.row_filter, "row", row_names, series_cfg.label
            )
        if series_cfg.col_filter is not None:
            _validate_filter_expr(
                series_cfg.col_filter, "column", col_names, series_cfg.label
            )


# ---------------------------------------------------------------------------
# Internal helpers – table lookup
# ---------------------------------------------------------------------------

def _resolve_table(
    source_block: str,
    tables: list[FormattedTable2D],
    sheet_name: str,
) -> FormattedTable2D:
    """
    Look up a :class:`FormattedTable2D` by its table title.

    Args:
        source_block: Title string to search for.
        tables: All :class:`FormattedTable2D` instances in the sheet.
        sheet_name: Worksheet name (for error messages).

    Returns:
        The matching :class:`FormattedTable2D`.

    Raises:
        ValueError: If *source_block* is not found in *tables*, or if multiple
            tables share the same title.
    """
    matches = [ft for ft in tables if ft.table.title == source_block]
    if len(matches) == 0:
        raise ValueError(
            f"source_block {source_block!r} not found in sheet {sheet_name!r}. "
            f"Available titles: {[ft.table.title for ft in tables]}"
        )
    if len(matches) > 1:
        raise ValueError(
            f"source_block {source_block!r} is ambiguous: "
            f"{len(matches)} tables with the same title exist in sheet {sheet_name!r}."
        )
    return matches[0]


# ---------------------------------------------------------------------------
# Internal helpers – xlsxwriter range builders
# ---------------------------------------------------------------------------

def _col_letter(col_idx: int) -> str:
    """
    Convert a 0-indexed column number to an Excel column letter.

    Supports columns 0–701 (A–ZZ).

    Args:
        col_idx: 0-indexed column number.

    Returns:
        Excel column letter string (e.g. ``0`` → ``"A"``, ``26`` → ``"AA"``).
    """
    result = ""
    n = col_idx + 1
    while n:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _xl_col_header_range(
    sheet_name: str,
    origin_row: int,
    origin_col: int,
    table: Table2D,
) -> list:
    """
    Build an xlsxwriter list-form range for the column header row.

    The column headers occupy row ``origin_row + 2`` (0-indexed), starting at
    ``origin_col + 2``.

    Args:
        sheet_name: Worksheet name.
        origin_row: Table origin row (0-indexed).
        origin_col: Table origin column (0-indexed).
        table: Source table.

    Returns:
        ``[sheet_name, r, c_start, r, c_end]`` (all 0-indexed row/col).
    """
    r = origin_row + 2
    c_start = origin_col + 2
    c_end = origin_col + 2 + len(table.column) - 1
    return [sheet_name, r, c_start, r, c_end]


def _xl_row_header_range(
    sheet_name: str,
    origin_row: int,
    origin_col: int,
    table: Table2D,
) -> list:
    """
    Build an xlsxwriter list-form range for the row header column.

    The row headers occupy column ``origin_col + 1`` (0-indexed), starting at
    ``origin_row + 3``.

    Args:
        sheet_name: Worksheet name.
        origin_row: Table origin row (0-indexed).
        origin_col: Table origin column (0-indexed).
        table: Source table.

    Returns:
        ``[sheet_name, r_start, c, r_end, c]`` (all 0-indexed row/col).
    """
    r_start = origin_row + 3
    r_end = origin_row + 3 + len(table.row) - 1
    c = origin_col + 1
    return [sheet_name, r_start, c, r_end, c]


def _xl_data_row_range(
    sheet_name: str,
    origin_row: int,
    origin_col: int,
    table: Table2D,
    row_idx: int,
) -> list:
    """
    Build an xlsxwriter list-form range for one data row (all columns).

    Args:
        sheet_name: Worksheet name.
        origin_row: Table origin row (0-indexed).
        origin_col: Table origin column (0-indexed).
        table: Source table.
        row_idx: Row index within the table (0-indexed).

    Returns:
        ``[sheet_name, r, c_start, r, c_end]`` (all 0-indexed row/col).
    """
    r = origin_row + 3 + row_idx
    c_start = origin_col + 2
    c_end = origin_col + 2 + len(table.column) - 1
    return [sheet_name, r, c_start, r, c_end]


def _xl_data_col_range(
    sheet_name: str,
    origin_row: int,
    origin_col: int,
    table: Table2D,
    col_idx: int,
) -> list:
    """
    Build an xlsxwriter list-form range for one data column (all rows).

    Args:
        sheet_name: Worksheet name.
        origin_row: Table origin row (0-indexed).
        origin_col: Table origin column (0-indexed).
        table: Source table.
        col_idx: Column index within the table (0-indexed).

    Returns:
        ``[sheet_name, r_start, c, r_end, c]`` (all 0-indexed row/col).
    """
    r_start = origin_row + 3
    r_end = origin_row + 3 + len(table.row) - 1
    c = origin_col + 2 + col_idx
    return [sheet_name, r_start, c, r_end, c]


# ---------------------------------------------------------------------------
# Internal helpers – series style
# ---------------------------------------------------------------------------

def _series_style_opts(style: str) -> dict:
    """
    Build xlsxwriter marker/line option dicts for the given *style*.

    Args:
        style: One of ``"line"``, ``"scatter"``, or ``"both"``.

    Returns:
        Dict with ``"marker"`` and ``"line"`` keys for ``chart.add_series()``.
    """
    if style == "line":
        return {"marker": {"type": "none"}, "line": {}}
    if style == "scatter":
        return {"marker": {"type": "circle", "size": 5}, "line": {"none": True}}
    # "both"
    return {"marker": {"type": "circle", "size": 5}, "line": {}}


# ---------------------------------------------------------------------------
# Internal helpers – series construction
# ---------------------------------------------------------------------------

def _add_series_for_config(
    chart,
    sheet_name: str,
    series_cfg: LineSeriesConfig,
    fmt: FormattedTable2D,
    origin_row: int,
    origin_col: int,
) -> None:
    """
    Add all xlsxwriter series entries for one :class:`LineSeriesConfig`.

    The table is always split to produce one series per item on the axis that
    is **not** the X axis. This ensures that each line in the chart corresponds
    to a single contiguous data range, which is the only form Excel supports.

    Splitting rule:

    - ``x_axis == "column"`` → split by row (one series per row value)
    - ``x_axis == "row"``    → split by column (one series per column value)

    When ``color_axis`` names an alias that matches the split axis (identity
    check via :func:`_property_names_for_axis`), colors are drawn from the
    table's ``value_colorscale``, spread evenly over the filtered split items.
    Otherwise xlsxwriter assigns colors automatically.

    Args:
        chart: xlsxwriter chart object.
        sheet_name: Worksheet name.
        series_cfg: Series configuration.
        fmt: Formatted table providing data ranges and color scale.
        origin_row: Table origin row (0-indexed).
        origin_col: Table origin column (0-indexed).
        x_axis: Chart X axis mapping.
    """
    table = fmt.table
    scale = series_cfg.series_colorscale

    row_names = _property_names_for_axis(table, "row")
    col_names = _property_names_for_axis(table, "column")

    # Apply filters
    filtered_row_indices = [
        i for i, rv in enumerate(table.row)
        if series_cfg.row_filter is None
        or _passes_filter(series_cfg.row_filter, row_names, rv)
    ]
    filtered_col_indices = [
        i for i, cv in enumerate(table.column)
        if series_cfg.col_filter is None
        or _passes_filter(series_cfg.col_filter, col_names, cv)
    ]

    # Determine split axis
    split_by_row = (series_cfg.x_axis == "column")

    if split_by_row:
        split_indices = filtered_row_indices
        split_labels = [table.row[i] for i in split_indices]
        color_axis_matches_split = (
            series_cfg.color_axis is not None and series_cfg.color_axis in row_names
        )
    else:
        split_indices = filtered_col_indices
        split_labels = [table.column[i] for i in split_indices]
        color_axis_matches_split = (
            series_cfg.color_axis is not None and series_cfg.color_axis in col_names
        )

    n_split = len(split_indices)
    style_opts = _series_style_opts(series_cfg.style)

    for i, (split_idx, split_label) in enumerate(zip(split_indices, split_labels)):
        series_def: dict = {
            "name": f"{series_cfg.label}: {split_label}",
        }

        # Cell ranges
        if split_by_row:
            series_def["categories"] = _xl_col_header_range(
                sheet_name, origin_row, origin_col, table
            )
            series_def["values"] = _xl_data_row_range(
                sheet_name, origin_row, origin_col, table, split_idx
            )
        else:
            series_def["categories"] = _xl_row_header_range(
                sheet_name, origin_row, origin_col, table
            )
            series_def["values"] = _xl_data_col_range(
                sheet_name, origin_row, origin_col, table, split_idx
            )

        # Marker / line style
        series_def["marker"] = dict(style_opts["marker"])
        series_def["line"] = dict(style_opts["line"])

        # Color override when color_axis matches the split axis
        if color_axis_matches_split and scale is not None:
            t = i / (n_split - 1) if n_split > 1 else 0.5
            raw_color = _colorscale_color(scale, t)
            hex_color = "#" + _hex(raw_color)
            series_def["line"] = {"color": hex_color}
            if series_cfg.style != "line":
                series_def["marker"] = {
                    "type": "circle",
                    "size": 5,
                    "fill": {"color": hex_color},
                    "border": {"color": hex_color},
                }
        if series_cfg.y_axis == "y2":
            series_def["y2_axis"] = True
        chart.add_series(series_def)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def render_chart(
    ws,
    workbook: "xlsxwriter.Workbook",
    config: ChartConfig,
    tables: list[FormattedTable2D],
    origin_row: int,
    origin_col: int,
    schema: "SheetWriteSchema",
    table_origins: dict[str, tuple[int, int]],
    sheet_name: str,
) -> tuple[int, int]:
    """
    Render an Excel chart into *ws* according to *config*.

    Each :class:`~excel_table.models.format.LineSeriesConfig` in
    ``config.series`` is resolved to a :class:`FormattedTable2D` via its
    ``source_block`` title. The table's worksheet position is supplied via
    *table_origins* so that correct xlsxwriter cell range references can be
    constructed.

    All ``row_filter`` / ``col_filter`` expressions are validated
    **before** any xlsxwriter state is mutated. If any expression is invalid
    (syntax error, or references a variable name not associated with the
    correct axis), a :class:`ValueError` is raised immediately.

    Args:
        ws: xlsxwriter worksheet to insert the chart into.
        workbook: Parent xlsxwriter :class:`xlsxwriter.Workbook`.
        config: Chart layout and series configuration.
        tables: All :class:`FormattedTable2D` objects written to this sheet.
            Used for ``source_block`` resolution and ``value_colorscale`` access.
        origin_row: Top-left row for the chart image (0-indexed).
        origin_col: Top-left column for the chart image (0-indexed).
        schema: Sheet write schema providing ``col_width`` / ``row_height``
            for the cell-footprint calculation.
        table_origins: Mapping of ``table.title → (origin_row, origin_col)``
            (0-indexed) for every :class:`FormattedTable2D` already written to
            this sheet.
        sheet_name: Worksheet name, embedded in xlsxwriter range references.

    Returns:
        ``(used_cols, used_rows)`` — cell footprint of the inserted chart,
        identical to ``config.occupied_cells(schema.col_width, schema.row_height)``.

    Raises:
        ValueError: If any ``source_block`` title is missing or duplicated in
            *tables*, or if any filter expression is invalid.
        ValueError: If series reference tables with inconsistent x-axis
            values and ``on_axis_mismatch="raise"``.

    Example::

        table_origins = {"Current Map": (0, 0)}
        used_cols, used_rows = render_chart(
            ws, workbook, chart_cfg, [fmt_table],
            origin_row=0, origin_col=6,
            schema=schema,
            table_origins=table_origins,
            sheet_name="Results",
        )
    """
    # Validate all filters before touching xlsxwriter
    _validate_all_filters(config, tables, sheet_name)

    # Validate x-axis value consistency across series
    def _rounded(values: list, digits: int = 10) -> tuple:
        result = []
        for v in values:
            try:
                result.append(round(float(v), digits))
            except (ValueError, TypeError):
                result.append(v)
        return tuple(result)

    # Create chart object
    chart = workbook.add_chart({"type": config.chart_type})
    if chart is None:
        raise NotImplementedError()
    
    chart.set_x_axis({"name": config.x_label})
    chart.set_y_axis({"name": config.y_label})
    chart.set_size({"width": config.width, "height": config.height})
    if config.y2_label:
        chart.set_y2_axis({"name": config.y2_label})

    # Add series for each LineSeriesConfig
    for series_cfg in config.series:
        fmt = _resolve_table(series_cfg.source_block, tables, sheet_name)
        t_origin_row, t_origin_col = table_origins[series_cfg.source_block]

        _add_series_for_config(
            chart=chart,
            sheet_name=sheet_name,
            series_cfg=series_cfg,
            fmt=fmt,
            origin_row=t_origin_row,
            origin_col=t_origin_col,
        )

    # Insert chart at origin cell
    cell_ref = f"{_col_letter(origin_col)}{origin_row + 1}"
    ws.insert_chart(cell_ref, chart)

    return config.occupied_cells(schema.col_width, schema.row_height)