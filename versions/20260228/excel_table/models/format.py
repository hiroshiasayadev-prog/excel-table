"""
Formatting and chart configuration models.

These models describe how tables and charts are rendered into Excel sheets.
They wrap :class:`Table2D` instances with visual metadata rather than
extending the data model itself.
"""
from __future__ import annotations
from typing import Literal

from pydantic import BaseModel, model_validator

from .table import Table2D


class ColorScale(BaseModel):
    """
    Defines a color gradient for cell background coloring.

    Attributes:
        min_color: Hex color string for the minimum value (e.g. ``"#FF0000"``).
        max_color: Hex color string for the maximum value.
        mid_color: Optional midpoint color. ``None`` produces a 2-color scale;
            providing a value produces a 3-color scale with the midpoint at 50%.
    """

    min_color: str
    max_color: str
    mid_color: str | None = None


class FormattedTable2D(BaseModel):
    """
    A :class:`Table2D` bundled with rendering metadata.

    Attributes:
        table: The underlying data table.
        column_color: Hex background color for column header cells.
        row_color: Hex background color for row header cells (including row_label).
        value_colorscale: Color gradient applied to data cells based on their value.
        row_label_direction: Text orientation of the row_label cell.
            ``"horizontal"`` (default) renders text normally;
            ``"vertical"`` rotates text 90 degrees counter-clockwise.
    """

    table: Table2D
    column_color: str
    row_color: str
    value_colorscale: ColorScale
    row_label_direction: Literal["horizontal", "vertical"] = "horizontal"


# ---------------------------------------------------------------------------
# Chart
# ---------------------------------------------------------------------------

class LineSeriesConfig(BaseModel):
    """
    Configuration for a single chart series.

    Attributes:
        label: Series label shown in the chart legend.
        source_block: Title of the :class:`Table2D` this series reads from.
            Titles must be unique within a sheet.
        style: Rendering style: ``"line"``, ``"scatter"``, or ``"both"``.
        color_axis: Name of the axis (e.g. ``"y"``) used to split the series
            into per-value sub-series and assign colors from a :class:`ColorScale`.
            ``None`` disables color splitting.
        row_filter: Optional filter expression evaluated per row value
            (e.g. ``"y >= 0"``). Rows where the expression is ``False`` are excluded.
            The axis value is passed as a scalar ``float``.
        col_filter: Optional filter expression evaluated per column value
            (e.g. ``"x <= 6000"``). Same evaluation rules as ``row_filter``.
    """

    label: str
    source_block: str
    style: Literal["line", "scatter", "both"]
    color_axis: str | None = None
    row_filter: str | None = None
    col_filter: str | None = None


class ChartConfig(BaseModel):
    """
    Configuration for an Excel chart embedded in a sheet.

    Multiple :class:`LineSeriesConfig` entries can reference different
    :class:`Table2D` blocks, allowing data from separate tables to be
    overlaid in a single chart.

    .. note::
        Each ``source_block`` value in ``series`` must reference a title that is
        unique within the sheet. Duplicate titles cause a ``ValueError`` at
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
        y_axis: Which dimension of the source table maps to the Y axis.
    """

    chart_type: Literal["line", "scatter", "bar"]
    width: int = 480
    height: int = 288
    series: list[LineSeriesConfig]
    x_label: str
    y_label: str
    x_axis: Literal["column", "row", "value"]
    y_axis: Literal["column", "row", "value"]

    @model_validator(mode="after")
    def _check_unique_source_blocks(self) -> "ChartConfig":
        """Validate that all source_block values within this chart are unique."""
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