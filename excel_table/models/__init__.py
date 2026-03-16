"""
Public re-exports for ``excel_table.models``.

Import from this package to avoid depending on internal module layout::

    from excel_table.models import Table2DFloat, FormattedTable2D, ChartConfig
"""
from .table_base import Table1D, Table2D, TableKeyValue, KeyValueScalar
from .table_typed import (
    Table1DFloat,
    Table1DInt,
    Table1DStr,
    Table2DFloat,
    Table2DInt,
    Table2DStr,
)
from .table_format import (
    FormattedTable1D,
    FormattedTable2D,
    FormattedTable1DSchema,
    FormattedTable2DSchema,
    TableKeyValueSchema,
)
from .chart_format import ColorScale, LineSeriesConfig, ChartConfig

__all__ = [
    # base
    "Table1D",
    "Table2D",
    "TableKeyValue",
    "KeyValueScalar",
    # typed
    "Table1DFloat",
    "Table1DInt",
    "Table1DStr",
    "Table2DFloat",
    "Table2DInt",
    "Table2DStr",
    # table format (write)
    "FormattedTable1D",
    "FormattedTable2D",
    # table schema (read)
    "FormattedTable1DSchema",
    "FormattedTable2DSchema",
    "TableKeyValueSchema",
    # chart
    "ColorScale",
    "LineSeriesConfig",
    "ChartConfig",
]