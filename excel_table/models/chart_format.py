"""
Chart configuration models.

These models describe how Excel charts are constructed from
:class:`~excel_table.models.table_base.Table2D` data.
"""
from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, model_validator


class ColorScale(BaseModel):
    """
    Defines a color gradient used for chart series coloring.

    This model is used exclusively by
    :class:`LineSeriesConfig` to assign colors to per-value sub-series when
    ``color_axis`` is set. It is **not** used for cell background coloring;
    cell colors are controlled via
    :attr:`~excel_table.models.table_format.FormattedTable2D.value_conditional_formats`
    using xlsxwriter's ``conditional_format`` dict syntax.

    Attributes:
        min_color: Hex color string at ``t=0`` (e.g. ``"#0000FF"``).
        max_color: Hex color string at ``t=1``.
        mid_color: Optional midpoint color at ``t=0.5``. ``None`` produces a
            2-color linear scale; providing a value produces a 3-color scale.
    """

    min_color: str
    max_color: str
    mid_color: str | None = None


class LineSeriesConfig(BaseModel):
    """
    Configuration for a single chart series.

    Series color priority
    ---------------------
    The following priority order determines the color applied to each
    sub-series. Higher priority overrides lower:

    1. **``series_colorscale``** — used when ``color_axis`` is set and the
       named axis matches the split axis. Colors are distributed evenly from
       ``min_color`` to ``max_color`` (via ``mid_color`` if present) across
       the filtered split items.
    2. **``series_color``** — a single hex color applied to all sub-series
       when ``color_axis`` is ``None`` or does not match the split axis.
    3. **xlsxwriter automatic** — when both ``series_colorscale`` and
       ``series_color`` are ``None``, xlsxwriter assigns colors automatically.

    Attributes:
        label: Series label shown in the chart legend.
        source_block: Title of the :class:`~excel_table.models.table_base.Table2D`
            this series reads from. Titles must be unique within a sheet.
        style: Rendering style.

            - ``"line"``    — line only, no markers
            - ``"scatter"`` — markers only, no line
            - ``"both"``    — markers and line

        color_axis: Name of the axis (e.g. ``"voltage"``) used to split the
            series into per-value sub-series and assign colors from
            ``series_colorscale``. The name must match either the default axis
            name (``"row"`` or ``"column"``) or a property alias defined on the
            domain subclass. ``None`` disables color-based splitting.
        series_colorscale: Color gradient applied to sub-series when
            ``color_axis`` is set and matches the split axis. See the color
            priority note above.
        series_color: Single hex color (e.g. ``"#FF0000"``) applied uniformly
            to all sub-series. Ignored when ``series_colorscale`` takes
            priority. See the color priority note above.
        row_filter: Optional filter expression evaluated per row value
            (e.g. ``"y >= 0"``). Rows where the expression is ``False``
            are excluded. The axis value is passed as a scalar ``float``.
            Variable names must correspond to the row axis of the source table
            (default ``"row"`` or a property alias). Using a column-axis name
            raises :class:`ValueError` at render time.
        col_filter: Optional filter expression evaluated per column value
            (e.g. ``"x <= 6000"``). Same rules as ``row_filter``, but
            variable names must correspond to the column axis.
    """

    label: str
    source_block: str
    style: Literal["line", "scatter", "both"]
    color_axis: str | None = None
    series_colorscale: ColorScale | None = None
    series_color: str | None = None
    row_filter: str | None = None
    col_filter: str | None = None


class ChartConfig(BaseModel):
    """
    Configuration for an Excel chart embedded in a sheet.

    Multiple :class:`LineSeriesConfig` entries can reference different
    :class:`~excel_table.models.table_base.Table2D` blocks, allowing data
    from separate tables to be overlaid in a single chart.

    .. note::
        Each ``source_block`` value in ``series`` must reference a title that
        is unique within the sheet. Duplicate titles cause a ``ValueError`` at
        chart render time. Within a single :class:`ChartConfig`, duplicate
        ``source_block`` values across series entries are also rejected at
        model construction time.

    Attributes:
        chart_type: Excel chart type: ``"line"``, ``"scatter"``, or ``"bar"``.
        width: Chart width in pixels. Default ``480``.
        height: Chart height in pixels. Default ``288``.
        series: List of series configurations. ``source_block`` values must be
            unique across entries.
        x_label: Axis label for the X axis.
        y_label: Axis label for the Y axis.
        x_axis: Which dimension of the source table maps to the X axis.
            ``"column"`` → column headers become X values;
            ``"row"`` → row headers become X values;
            ``"value"`` → treated as ``"row"`` for header lookup.
        y_axis: Which dimension of the source table maps to the Y axis.
    """

    chart_type: Literal["line", "scatter", "bar"]
    width: int = 480
    height: int = 288
    series: list[LineSeriesConfig]
    x_label: str = ""
    y_label: str = ""
    x_axis: Literal["column", "row", "value"]
    y_axis: Literal["column", "row", "value"]

    @model_validator(mode="after")
    def _check_unique_source_blocks(self) -> "ChartConfig":
        """Validate that all ``source_block`` values within this chart are unique."""
        seen: set[str] = set()
        for s in self.series:
            if s.source_block in seen:
                raise ValueError(
                    f"Duplicate source_block '{s.source_block}' in ChartConfig.series. "
                    "Each series must reference a distinct Table2D title."
                )
            seen.add(s.source_block)
        return self

    def occupied_cells(self, col_width: int, row_height: int) -> tuple[int, int]:
        """
        Return the number of cells the chart occupies.

        Args:
            col_width: Width of one cell in pixels.
            row_height: Height of one cell in pixels.

        Returns:
            ``(cols, rows)`` — the cell footprint of the chart, rounded up.
        """
        import math
        cols = math.ceil(self.width / col_width)
        rows = math.ceil(self.height / row_height)
        return cols, rows