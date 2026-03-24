"""
Base table models for structured Excel tables.

All models are immutable Pydantic v2 ``BaseModel`` instances.
Mutation should be done via ``model_copy(update={...})``.

Read / Write contract for ``TableKeyValue.value``
--------------------------------------------------
- **Read** (openpyxl → model): all cell values are cast to ``str`` regardless
  of the Excel cell type. Type interpretation is the caller's responsibility.
- **Write** (model → xlsxwriter): each element is written according to its
  Python type. ``str`` values are written as text cells; ``int`` / ``float``
  values are written as numeric cells so that conditional formatting rules
  based on numeric comparisons work correctly. ``None`` is written as a blank
  cell.

Domain subclass pattern
-----------------------
Inherit from a typed subclass (e.g. :class:`~excel_table.models.table_typed.Table2DFloat`)
to add domain-specific property aliases::

    from excel_table.models.table_typed import Table2DFloat

    class ValueMap(Table2DFloat):
        @property
        def x(self) -> list[str]:
            \"\"\"Alias for :attr:`column` (x axis in rpm).\"\"\"
            return self.column

        @property
        def y(self) -> list[str]:
            \"\"\"Alias for :attr:`row` (y axis in Nm).\"\"\"
            return self.row

    new_map = value_map.model_copy(update={"column": new_x})

``model_copy`` preserves the concrete subclass type and does **not** re-run
validators by default. Pass ``validate=True`` if coercion is needed on the
updated fields.
"""
from __future__ import annotations

from typing import Generic, TypeVar

from pydantic import BaseModel, model_validator

T = TypeVar("T", float, int, str)

KeyValueScalar = str | int | float | None
"""Type alias for a single value in :class:`TableKeyValue`.

- ``str``   — text cell (written as-is by xlsxwriter)
- ``int``   — integer numeric cell
- ``float`` — floating-point numeric cell
- ``None``  — blank cell
"""


class Table1D(BaseModel, Generic[T]):
    """
    A one-dimensional table with a single row of values.

    Layout in Excel (``orientation="horizontal"``, the default)::

        (0,0) title
        (1,0) column_label  [merged across all columns]
        (2,0) col1  (2,1) col2  (2,2) col3
        (3,0) val   (3,1) val   (3,2) val

    Layout in Excel (``orientation="vertical"``)::

        (0,0) title
        (1,0) column_label  (1,1) col1  (1,2) val
        (2,0)               (2,1) col2  (2,2) val
        (3,0)               (3,1) col3  (3,2) val

    Attributes:
        title: Table title. Used as the anchor for sheet scanning.
        column_label: Label for the column axis.
        column: List of column header values. Accepts ``str``, ``float``, or ``int``.
                Numeric values are written as numeric cells in Excel; string values
                are written as text cells.
        values: Nested list of shape ``[1][len(column)]`` for horizontal
            orientation, or ``[len(column)][1]`` for vertical orientation.
            Inner ``None`` represents an empty cell.

    Note:
        The ``orientation`` attribute lives on
        :class:`~excel_table.models.table_format.FormattedTable1D`, not here,
        because orientation is a rendering concern rather than a data concern.
        The data shape in ``values`` is always ``[1][len(column)]`` regardless
        of orientation; the writer transposes as needed.
    """

    title: str
    column_label: str
    column: list[str | float | int]
    values: list[list[T | None]]

    @model_validator(mode="after")
    def _check_shape(self) -> "Table1D[T]":
        """Validate that every row in ``values`` matches ``len(column)``."""
        n = len(self.column)
        for i, row in enumerate(self.values):
            if len(row) != n:
                raise ValueError(
                    f"values[{i}] has {len(row)} elements, expected {n} (len(column))"
                )
        return self


class Table2D(BaseModel, Generic[T]):
    """
    A two-dimensional table with row and column axes.

    Default layout in Excel (``column_location="top"``, ``row_location="left"``)::

        (0,0) title [single cell]
        (1,0) [2×2 merged blank]            (1,2) column_label [merged across columns]
        (2,0)                               (2,2) col1  (2,3) col2  (2,4) col3
        (3,0) row_label [merged down]       (3,1) row1  (3,2) val   (3,3) val
        (4,0)                               (4,1) row2  (4,2) val   (4,3) val

    See :class:`~excel_table.models.table_format.FormattedTable2D` for
    ``column_location`` / ``row_location`` variants.

    Attributes:
        title: Table title. Used as the anchor for sheet scanning.

            .. warning::
                When referencing this table from a
                :class:`~excel_table.models.chart_format.ChartConfig`
                via ``source_block``, the title must be unique within the sheet.
                Duplicate titles cause a ``ValueError`` at chart render time.

        column_label: Label for the column axis.
        row_label: Label for the row axis.
        column: List of column header values. Accepts ``str``, ``float``, or ``int``.
                Numeric values are written as numeric cells in Excel; string values
                are written as text cells.
        row: List of row header values. Same type rules as ``column``.
        values: Nested list of shape ``[len(row)][len(column)]``.
            Inner ``None`` represents an empty cell.
    """

    title: str
    column_label: str
    row_label: str
    column: list[str | float | int]
    row: list[str | float | int]
    values: list[list[T | None]]

    @model_validator(mode="after")
    def _check_shape(self) -> "Table2D[T]":
        """Validate minimum axis size and values shape.

        ``column`` and ``row`` must each have at least 2 elements. A
        single-element axis cannot be reliably distinguished from a label cell
        during sheet scanning (the merge width of ``column_label`` would be 1,
        identical to a single column-value cell). Use
        :class:`~excel_table.models.table_base.Table1D` for single-axis data
        instead.

        Raises:
            ValueError: If ``len(column) < 2``, ``len(row) < 2``, or if
                ``values`` shape does not match ``len(row) × len(column)``.
        """
        n_col = len(self.column)
        n_row = len(self.row)
        if n_col < 2:
            raise ValueError(
                f"Table2D requires at least 2 columns, got {n_col}. "
                "Use Table1D for single-axis data."
            )
        if n_row < 2:
            raise ValueError(
                f"Table2D requires at least 2 rows, got {n_row}. "
                "Use Table1D for single-axis data."
            )
        if len(self.values) != n_row:
            raise ValueError(
                f"values has {len(self.values)} rows, expected {n_row} (len(row))"
            )
        for i, row in enumerate(self.values):
            if len(row) != n_col:
                raise ValueError(
                    f"values[{i}] has {len(row)} elements, expected {n_col} (len(column))"
                )
        return self


class TableKeyValue(BaseModel):
    """
    A flat key-value table with a single header row and a single value row.

    Layout in Excel::

        (0,0) title
        (1,0) key1  (1,1) key2  (1,2) key3
        (2,0) val1  (2,1) val2  (2,2) val3

    Read / Write contract
    ---------------------
    - **Read**: all ``value`` elements are returned as ``str`` regardless of
      the Excel cell type. Type conversion is the caller's responsibility.
    - **Write**: each element is written according to its Python type.

      - ``str``   → text cell (xlsxwriter ``write_string``)
      - ``int``   → integer numeric cell (xlsxwriter ``write_number``)
      - ``float`` → float numeric cell (xlsxwriter ``write_number``).
        If the float is a whole number (e.g. ``2.0``), it is written as ``int``
        to avoid ``"2.0"`` appearing in the cell.
      - ``None``  → blank cell

      Writing numeric types as numeric cells ensures that conditional
      formatting rules using numeric comparisons (``>=``, ``<``, etc.) work
      correctly in Excel.

    Attributes:
        title: Table title.
        column: List of key strings. Always ``str``.
        value: List of values. May contain ``str``, ``int``, ``float``, or
            ``None``. See the Read / Write contract above.
    """

    title: str
    column: list[str]
    value: list[KeyValueScalar]

    @model_validator(mode="after")
    def _check_shape(self) -> "TableKeyValue":
        """Validate that ``column`` and ``value`` have the same length."""
        if len(self.column) != len(self.value):
            raise ValueError(
                f"column length {len(self.column)} != value length {len(self.value)}"
            )
        return self