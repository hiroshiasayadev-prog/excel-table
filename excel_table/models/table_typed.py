"""
Typed subclasses of the base table models.

Each class fixes the type parameter ``T`` of its base class to ``float``,
``int``, or ``str``, and applies the appropriate Pydantic v2 coercion strategy.

openpyxl type coercion
----------------------
openpyxl returns cell values in their native Excel type: numeric cells as
``float`` or ``int``, text cells as ``str``, empty cells as ``None``.
The coercion behaviour for each typed subclass is as follows:

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
     - OK
   * - ``float`` / ``int``
     - ``str``
     - Coerced via ``coerce_numbers_to_str=True``
   * - ``str``
     - ``str``
     - OK

When creating **domain subclasses** that further override ``values``,
``column``, or ``row``, replicate the coercion strategy of the corresponding
typed subclass shown above.
"""
from __future__ import annotations

from pydantic import ConfigDict, field_validator

from .table_base import Table1D, Table2D


# ---------------------------------------------------------------------------
# Table1D typed subclasses
# ---------------------------------------------------------------------------

class Table1DFloat(Table1D[float]):
    """
    :class:`Table1D` with ``float | None`` values.

    ``int`` and numeric ``str`` cell values are coerced to ``float``.
    """

    values: list[list[float | None]]


class Table1DInt(Table1D[int]):
    """
    :class:`Table1D` with ``int | None`` values.

    ``float`` cell values are **truncated** to ``int`` (e.g. ``1.7 → 1``).
    Numeric ``str`` cell values are also accepted.
    """

    values: list[list[int | None]]

    @field_validator("values", mode="before")
    @classmethod
    def _truncate_floats(cls, v: list) -> list:
        """Truncate ``float`` cell values to ``int`` before Pydantic validation."""
        return [
            [int(cell) if isinstance(cell, float) else cell for cell in row]
            for row in v
        ]


class Table1DStr(Table1D[str]):
    """
    :class:`Table1D` with ``str | None`` values.

    ``float`` and ``int`` cell values are coerced to ``str`` via
    ``coerce_numbers_to_str=True``.
    """

    model_config = ConfigDict(coerce_numbers_to_str=True)
    values: list[list[str | None]]


# ---------------------------------------------------------------------------
# Table2D typed subclasses
# ---------------------------------------------------------------------------

class Table2DFloat(Table2D[float]):
    """
    :class:`Table2D` with ``float | None`` values.

    ``int`` and numeric ``str`` cell values are coerced to ``float``.
    """

    values: list[list[float | None]]


class Table2DInt(Table2D[int]):
    """
    :class:`Table2D` with ``int | None`` values.

    ``float`` cell values are **truncated** to ``int`` (e.g. ``1.7 → 1``).
    Numeric ``str`` cell values are also accepted.
    """

    values: list[list[int | None]]

    @field_validator("values", mode="before")
    @classmethod
    def _truncate_floats(cls, v: list) -> list:
        """Truncate ``float`` cell values to ``int`` before Pydantic validation."""
        return [
            [int(cell) if isinstance(cell, float) else cell for cell in row]
            for row in v
        ]


class Table2DStr(Table2D[str]):
    """
    :class:`Table2D` with ``str | None`` values.

    ``float`` and ``int`` cell values are coerced to ``str`` via
    ``coerce_numbers_to_str=True``.
    """

    model_config = ConfigDict(coerce_numbers_to_str=True)
    values: list[list[str | None]]