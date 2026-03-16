"""
Tests for data model validation (table.py, format.py).

Covers shape validation, type coercion behaviour, domain subclass patterns,
ChartConfig constraints, and ColorScale construction.
"""
import pytest

from excel_table.models.table import (
    Table1DFloat,
    Table2DFloat,
    Table2DInt,
    Table2DStr,
    TableKeyValue,
)
from excel_table.models.format import (
    ChartConfig,
    ColorScale,
    FormattedTable2D,
    LineSeriesConfig,
)


# ---------------------------------------------------------------------------
# Table2D
# ---------------------------------------------------------------------------

class TestTable2DValidation:
    """Shape validation and coercion for Table2D typed subclasses."""

    def test_valid(self):
        """A correctly shaped Table2DFloat is created without error."""
        t = Table2DFloat(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=["1.0", "2.0", "3.0"],
            row=["10.0", "20.0"],
            values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],
        )
        assert t.title == "Current Map"
        assert len(t.values) == 2

    def test_values_row_count_mismatch_raises(self):
        """values with fewer rows than len(row) raises ValueError."""
        with pytest.raises(ValueError, match="values has"):
            Table2DFloat(
                title="Current Map",
                column_label="Voltage",
                row_label="Current",
                column=["1.0", "2.0"],
                row=["10.0", "20.0"],
                values=[[0.1, 0.2]],  # 1 row, expected 2
            )

    def test_values_col_count_mismatch_raises(self):
        """values row with wrong column count raises ValueError."""
        with pytest.raises(ValueError, match="values\\[0\\]"):
            Table2DFloat(
                title="Current Map",
                column_label="Voltage",
                row_label="Current",
                column=["1.0", "2.0"],
                row=["10.0"],
                values=[[0.1, 0.2, 0.3]],  # 3 cols, expected 2
            )

    def test_int_subclass_truncates_float(self):
        """Table2DInt truncates float cell values to int (e.g. 1.7 → 1)."""
        from excel_table.models.table import Table2DInt
        t = Table2DInt(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=["1.0", "2.0"],
            row=["10.0"],
            values=[[1.7, 2.9]],
        )
        assert t.values[0] == [1, 2]

    def test_str_subclass_coerces_numeric(self):
        """Table2DStr coerces float and int cell values to str."""
        t = Table2DStr(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=["1.0", "2.0"],
            row=["10.0"],
            values=[[0.1, 2]],
        )
        assert t.values[0] == ["0.1", "2"]

    def test_none_values_allowed(self):
        """None is a valid cell value (represents an empty cell)."""
        t = Table2DFloat(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=["1.0", "2.0"],
            row=["10.0"],
            values=[[None, 0.2]],
        )
        assert t.values[0][0] is None

    def test_model_copy_returns_same_type(self):
        """model_copy preserves the concrete subclass type."""
        t = Table2DFloat(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=["1.0", "2.0"],
            row=["10.0"],
            values=[[0.1, 0.2]],
        )
        t2 = t.model_copy(update={"column": ["3.0", "4.0"]})
        assert isinstance(t2, Table2DFloat)
        assert t2.column == ["3.0", "4.0"]


# ---------------------------------------------------------------------------
# Table1D
# ---------------------------------------------------------------------------

class TestTable1DValidation:
    """Shape validation for Table1D typed subclasses."""

    def test_valid(self):
        """A correctly shaped Table1DFloat is created without error."""
        t = Table1DFloat(
            title="Voltage Points",
            column_label="Voltage",
            column=["1.0", "2.0", "3.0"],
            values=[[0.1, 0.2, 0.3]],
        )
        assert len(t.column) == 3

    def test_col_count_mismatch_raises(self):
        """values row with wrong column count raises ValueError."""
        with pytest.raises(ValueError, match="values\\[0\\]"):
            Table1DFloat(
                title="Voltage Points",
                column_label="Voltage",
                column=["1.0", "2.0"],
                values=[[0.1, 0.2, 0.3]],  # 3 cols, expected 2
            )


# ---------------------------------------------------------------------------
# TableKeyValue
# ---------------------------------------------------------------------------

class TestTableKeyValueValidation:
    """Shape validation for TableKeyValue."""

    def test_valid(self):
        """A correctly shaped TableKeyValue is created without error."""
        t = TableKeyValue(
            title="Conditions",
            column=["temperature", "frequency"],
            value=["25.0", "50.0"],
        )
        assert t.value[0] == "25.0"

    def test_length_mismatch_raises(self):
        """column and value with different lengths raises ValueError."""
        with pytest.raises(ValueError, match="column length"):
            TableKeyValue(
                title="Conditions",
                column=["temperature"],
                value=["25.0", "50.0"],
            )


# ---------------------------------------------------------------------------
# Domain subclass with overridden column/row types
# ---------------------------------------------------------------------------

class TestDomainSubclass:
    """Domain subclass patterns: overriding column/row types and model_copy."""

    def test_float_column_override(self):
        """A domain subclass can override column/row to list[float]."""
        class CurrentMap(Table2DFloat):
            column: list[float]
            row: list[float]

        m = CurrentMap(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=[1.0, 2.0, 3.0],
            row=[10.0, 20.0],
            values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],
        )
        assert m.column == [1.0, 2.0, 3.0]

    def test_model_copy_preserves_subclass(self):
        """model_copy on a domain subclass returns the same subclass type."""
        class CurrentMap(Table2DFloat):
            column: list[float]
            row: list[float]

        m = CurrentMap(
            title="Current Map",
            column_label="Voltage",
            row_label="Current",
            column=[1.0, 2.0],
            row=[10.0],
            values=[[0.1, 0.2]],
        )
        m2 = m.model_copy(update={"column": [3.0, 4.0]})
        assert isinstance(m2, CurrentMap)
        assert m2.column == [3.0, 4.0]


# ---------------------------------------------------------------------------
# ChartConfig
# ---------------------------------------------------------------------------

class TestChartConfigValidation:
    """Validation for ChartConfig and LineSeriesConfig."""

    def _make_series(self, source_block: str) -> LineSeriesConfig:
        """Helper to create a minimal LineSeriesConfig."""
        return LineSeriesConfig(
            label="series",
            source_block=source_block,
            style="line",
        )

    def test_valid(self):
        """ChartConfig with distinct source_blocks is created without error."""
        cfg = ChartConfig(
            chart_type="line",
            series=[
                self._make_series("Current Map"),
                self._make_series("Voltage Map"),
            ],
            x_label="Voltage",
            y_label="Current",
            x_axis="column",
            y_axis="value",
        )
        assert len(cfg.series) == 2

    def test_duplicate_source_block_raises(self):
        """Two series with the same source_block raises ValueError."""
        with pytest.raises(ValueError, match="Duplicate source_block"):
            ChartConfig(
                chart_type="line",
                series=[
                    self._make_series("Current Map"),
                    self._make_series("Current Map"),
                ],
                x_label="Voltage",
                y_label="Current",
                x_axis="column",
                y_axis="value",
            )

    def test_occupied_cells(self):
        """occupied_cells returns correct (cols, rows) footprint for given px dimensions."""
        cfg = ChartConfig(
            chart_type="line",
            series=[self._make_series("Current Map")],
            x_label="V",
            y_label="I",
            x_axis="column",
            y_axis="value",
            width=480,
            height=288,
        )
        cols, rows = cfg.occupied_cells(col_width=64, row_height=20)
        assert cols == 8   # ceil(480/64)
        assert rows == 15  # ceil(288/20)


# ---------------------------------------------------------------------------
# ColorScale
# ---------------------------------------------------------------------------

class TestColorScale:
    """Construction of 2-color and 3-color scales."""

    def test_two_color(self):
        """ColorScale without mid_color produces a 2-color scale."""
        cs = ColorScale(min_color="#000000", max_color="#FFFFFF")
        assert cs.mid_color is None

    def test_three_color(self):
        """ColorScale with mid_color produces a 3-color scale."""
        cs = ColorScale(min_color="#0000FF", max_color="#FF0000", mid_color="#FFFFFF")
        assert cs.mid_color == "#FFFFFF"