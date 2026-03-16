"""
excel-table

High-level public API for reading, writing and formatting Excel tables.

Public components:
- Table models (1D/2D/KeyValue)
- Formatting & chart configs
- read_sheet / write_sheet
- render_chart
"""

# =========================
# Models
# =========================
from .models import (
    Table1D,
    Table1DFloat,
    Table1DInt,
    Table1DStr,
    Table2D,
    Table2DFloat,
    Table2DInt,
    Table2DStr,
    TableKeyValue,
    ColorScale,
    FormattedTable2D,
    ChartConfig,
    LineSeriesConfig,
)

# =========================
# IO
# =========================
from .reader import SheetReadSchema, read_sheet
from .writer import SheetWriteSchema, write_sheet

# =========================
# Chart
# =========================
from .chart import render_chart


__all__ = [
    # Models
    "Table1D",
    "Table1DFloat",
    "Table1DInt",
    "Table1DStr",
    "Table2D",
    "Table2DFloat",
    "Table2DInt",
    "Table2DStr",
    "TableKeyValue",
    "ColorScale",
    "FormattedTable2D",
    "ChartConfig",
    "LineSeriesConfig",
    # IO
    "SheetReadSchema",
    "read_sheet",
    "SheetWriteSchema",
    "write_sheet",
    # Chart
    "render_chart",
]
