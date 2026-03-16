"""
Data models for structured Excel tables.

All models are immutable Pydantic BaseModels.
Mutation should be done via ``model_copy(update={...})``.

openpyxl type coercion notes
-----------------------------
openpyxl returns cell values in their native Excel type:
numeric cells as ``float`` or ``int``, text cells as ``str``, empty cells as
``None``.  Pydantic v2 coercion behaviour for each typed subclass is as
follows:

.. list-table::
   :header-rows: 1

   * - Cell type (openpyxl)
     - Model field type
     - Behaviour
   * - ``float``
     - ``float``
     - OK
   * - ``int``
     - ``float``
     - OK (widened automatically)
   * - ``str`` (numeric string)
     - ``float``
     - OK (Pydantic parses numeric strings by default)
   * - ``float``
     - ``int``
     - Truncated via ``field_validator`` (e.g. ``1.7 → 1``)
   * - ``int``
     - ``int``
     - OK
   * - ``str`` (numeric string)
     - ``int``
     - OK (Pydantic parses numeric strings by default)
   * - ``float`` / ``int``
     - ``str``
     - Coerced via ``coerce_numbers_to_str=True``
   * - ``str``
     - ``str``
     - OK

When creating **domain subclasses** that override ``values``, ``column``, or
``row`` with a different element type, replicate the corresponding coercion
strategy shown above.  For ``str`` fields, add::

    from pydantic import ConfigDict
    model_config = ConfigDict(coerce_numbers_to_str=True)

For ``int`` fields, add a ``field_validator`` that truncates floats (see
:class:`Table2DInt` for a reference implementation).

``model_copy`` preserves the subclass type, but does **not** re-run
validators by default.  Pass ``validate=True`` if coercion is needed on the
updated fields, or ensure the supplied values already match the field type.
See the project README for domain subclass patterns.
"""
from __future__ import annotations
from typing import Generic, TypeVar

from pydantic import BaseModel, ConfigDict, field_validator, model_validator

T = TypeVar("T", float, int, str)


class Table1D(BaseModel, Generic[T]):
    """
    A one-dimensional table with a single row of values.

    Layout in Excel::

        (0,0) title
        (1,0) column_label  [merged across all columns]
        (2,0) col1  (2,1) col2  (2,2) col3
        (3,0) val   (3,1) val   (3,2) val

    Attributes:
        title: Table title. Used as the anchor for sheet scanning.
        column_label: Label for the column axis (displayed as a merged header).
        column: List of column header strings.
        values: Nested list of shape ``[1][len(column)]``. Inner ``None`` represents an empty cell.

    See Also:
        Module docstring for openpyxl type coercion behaviour and domain
        subclass patterns.
    """

    title: str
    column_label: str
    column: list[str]
    values: list[list[T | None]]

    @model_validator(mode="after")
    def _check_shape(self) -> "Table1D[T]":
        """Validate that every row in values matches len(column)."""
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

    Layout in Excel::

        (0,0) title [single cell]
        (1,0) [2×2 merged blank]            (1,2) column_label [merged across columns]
        (2,0)                               (2,2) col1  (2,3) col2  (2,4) col3
        (3,0) row_label [merged down]       (3,1) row1  (3,2) val   (3,3) val
        (4,0)                               (4,1) row2  (4,2) val   (4,3) val

    Attributes:
        title: Table title. Used as the anchor for sheet scanning.

            .. warning::
                When referencing this table from :class:`~excel_table.models.format.ChartConfig`
                via ``source_block``, the title must be unique within the sheet.
                Duplicate titles cause a ``ValueError`` at chart render time.

        column_label: Label for the column axis.
        row_label: Label for the row axis.
        column: List of column header strings.
        row: List of row header strings.
        values: Nested list of shape ``[len(row)][len(column)]``. Inner ``None`` represents an empty cell.

    See Also:
        Module docstring for openpyxl type coercion behaviour and domain
        subclass patterns.
    """

    title: str
    column_label: str
    row_label: str
    column: list[str]
    row: list[str]
    values: list[list[T | None]]

    @model_validator(mode="after")
    def _check_shape(self) -> "Table2D[T]":
        """Validate that values shape matches len(row) × len(column)."""
        n_col = len(self.column)
        n_row = len(self.row)
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

    ``column`` and ``value`` are always ``str``; type conversion is the
    caller's responsibility.

    Layout in Excel::

        (0,0) title
        (1,0) key1  (1,1) key2  (1,2) key3
        (2,0) val1  (2,1) val2  (2,2) val3

    Attributes:
        title: Table title.
        column: List of key strings.
        value: List of value strings, same length as ``column``.
    """

    title: str
    column: list[str]
    value: list[str]

    @model_validator(mode="after")
    def _check_shape(self) -> "TableKeyValue":
        """Validate that column and value have the same length."""
        if len(self.column) != len(self.value):
            raise ValueError(
                f"column length {len(self.column)} != value length {len(self.value)}"
            )
        return self


# ---------------------------------------------------------------------------
# Typed subclasses
# ---------------------------------------------------------------------------

class Table1DFloat(Table1D[float]):
    """Table1D with ``float | None`` values.

    ``int`` and numeric ``str`` cell values are coerced to ``float``.
    """
    values: list[list[float | None]]


class Table1DInt(Table1D[int]):
    """Table1D with ``int | None`` values.

    ``float`` cell values are **truncated** to ``int`` (e.g. ``1.7 → 1``).
    Numeric ``str`` cell values are also accepted.
    """
    values: list[list[int | None]]

    @field_validator("values", mode="before")
    @classmethod
    def _truncate_floats(cls, v: list) -> list:
        """Truncate float cell values to int before Pydantic validation."""
        return [
            [int(cell) if isinstance(cell, float) else cell for cell in row]
            for row in v
        ]


class Table1DStr(Table1D[str]):
    """Table1D with ``str | None`` values.

    ``float`` and ``int`` cell values are coerced to ``str`` via
    ``coerce_numbers_to_str=True``.
    """
    model_config = ConfigDict(coerce_numbers_to_str=True)
    values: list[list[str | None]]


class Table2DFloat(Table2D[float]):
    """Table2D with ``float | None`` values.

    ``int`` and numeric ``str`` cell values are coerced to ``float``.
    """
    values: list[list[float | None]]


class Table2DInt(Table2D[int]):
    """Table2D with ``int | None`` values.

    ``float`` cell values are **truncated** to ``int`` (e.g. ``1.7 → 1``).
    Numeric ``str`` cell values are also accepted.
    """
    values: list[list[int | None]]

    @field_validator("values", mode="before")
    @classmethod
    def _truncate_floats(cls, v: list) -> list:
        """Truncate float cell values to int before Pydantic validation."""
        return [
            [int(cell) if isinstance(cell, float) else cell for cell in row]
            for row in v
        ]


class Table2DStr(Table2D[str]):
    """Table2D with ``str | None`` values.

    ``float`` and ``int`` cell values are coerced to ``str`` via
    ``coerce_numbers_to_str=True``.
    """
    model_config = ConfigDict(coerce_numbers_to_str=True)
    values: list[list[str | None]]