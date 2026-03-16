"""
Public models for excel_table.

Re-exports:
- Table1D / Table2D variants
- TableKeyValue
- Formatting configs (ColorScale, FormattedTable2D, ChartConfig, LineSeriesConfig)
"""

from .table import (
    Table1D,
    Table1DFloat,
    Table1DInt,
    Table1DStr,
    Table2D,
    Table2DFloat,
    Table2DInt,
    Table2DStr,
    TableKeyValue,
)
from .format import (
    ColorScale,
    FormattedTable2D,
    ChartConfig,
    LineSeriesConfig,
)

__all__ = [
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
]