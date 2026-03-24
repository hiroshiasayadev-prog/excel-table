"""
Tests for ChartConfig and LineSeriesConfig model validation.

Covers:
- y2 axis field presence on LineSeriesConfig
- x_axis per-series independence
- occupied_cells footprint calculation
- _passes_filter string axis support
"""
import pytest

from excel_table.models import ColorScale, LineSeriesConfig, ChartConfig
from excel_table.chart import _passes_filter


# ---------------------------------------------------------------------------
# LineSeriesConfig
# ---------------------------------------------------------------------------

class TestLineSeriesConfig:
    """Validation and defaults for LineSeriesConfig."""

    def _make(self, **kwargs) -> LineSeriesConfig:
        defaults = dict(
            label="series",
            source_block="IV Result",
            style="line",
        )
        return LineSeriesConfig(**{**defaults, **kwargs})

    def test_default_y_axis_is_y1(self):
        """y_axis defaults to 'y1'."""
        s = self._make()
        assert s.y_axis == "y1"

    def test_y2_axis(self):
        """y_axis='y2' is accepted."""
        s = self._make(y_axis="y2")
        assert s.y_axis == "y2"

    def test_default_x_axis_is_row(self):
        """x_axis defaults to 'row'."""
        s = self._make()
        assert s.x_axis == "row"

    def test_x_axis_column(self):
        """x_axis='column' is accepted."""
        s = self._make(x_axis="column")
        assert s.x_axis == "column"

    def test_x_axis_value_rejected(self):
        """x_axis='value' is no longer valid — must be 'row' or 'column'."""
        with pytest.raises(Exception):
            self._make(x_axis="value")

    def test_two_series_different_x_axis(self):
        """Two series in the same ChartConfig can have different x_axis values."""
        cfg = ChartConfig(
            chart_type="line",
            x_label="X",
            y_label="Y",
            series=[
                self._make(source_block="IV Result", x_axis="row"),
                self._make(source_block="gm", x_axis="row"),
            ],
        )
        assert cfg.series[0].x_axis == "row"
        assert cfg.series[1].x_axis == "row"

    def test_duplicate_source_block_allowed(self):
        """Two series referencing the same source_block is now valid (e.g. col_filter splits)."""
        cfg = ChartConfig(
            chart_type="line",
            x_label="Vgs [V]",
            y_label="Jd [mA/mm]",
            series=[
                self._make(
                    label="forward",
                    source_block="Transfer Result",
                    col_filter="column == 'forward'",
                ),
                self._make(
                    label="backward",
                    source_block="Transfer Result",
                    col_filter="column == 'backward'",
                ),
            ],
        )
        assert len(cfg.series) == 2


# ---------------------------------------------------------------------------
# ChartConfig
# ---------------------------------------------------------------------------

class TestChartConfig:
    """Validation and defaults for ChartConfig."""

    def _make_series(self, source_block: str = "IV Result") -> LineSeriesConfig:
        return LineSeriesConfig(label="s", source_block=source_block, style="line")

    def test_valid(self):
        """ChartConfig with minimal required fields is created without error."""
        cfg = ChartConfig(
            chart_type="line",
            x_label="Vds [V]",
            y_label="Jd [mA/mm]",
            series=[self._make_series()],
        )
        assert cfg.chart_type == "line"

    def test_y2_label_default_empty(self):
        """y2_label defaults to empty string."""
        cfg = ChartConfig(
            chart_type="line",
            series=[self._make_series()],
        )
        assert cfg.y2_label == ""

    def test_y2_label_set(self):
        """y2_label can be set for dual-axis charts."""
        cfg = ChartConfig(
            chart_type="line",
            y_label="Jd [mA/mm]",
            y2_label="gm [mS/mm]",
            series=[
                self._make_series("Transfer Result"),
                LineSeriesConfig(
                    label="gm",
                    source_block="gm",
                    style="line",
                    y_axis="y2",
                ),
            ],
        )
        assert cfg.y2_label == "gm [mS/mm]"
        assert cfg.series[1].y_axis == "y2"

    def test_occupied_cells(self):
        """occupied_cells returns correct (cols, rows) for given pixel dimensions."""
        cfg = ChartConfig(
            chart_type="line",
            series=[self._make_series()],
            width=480,
            height=288,
        )
        cols, rows = cfg.occupied_cells(col_width=64, row_height=20)
        assert cols == 8   # ceil(480/64)
        assert rows == 15  # ceil(288/20)


# ---------------------------------------------------------------------------
# _passes_filter
# ---------------------------------------------------------------------------

class TestPassesFilter:
    """Unit tests for _passes_filter — numeric and string axis values."""

    def test_numeric_gt(self):
        """Numeric filter: value > threshold."""
        assert _passes_filter("vgs >= 0.0", frozenset(["vgs"]), "0.0") is True
        assert _passes_filter("vgs >= 0.0", frozenset(["vgs"]), "-0.5") is False

    def test_numeric_lt(self):
        """Numeric filter: value < threshold."""
        assert _passes_filter("vds <= 0.5", frozenset(["vds"]), "0.5") is True
        assert _passes_filter("vds <= 0.5", frozenset(["vds"]), "1.0") is False

    def test_string_equality(self):
        """String filter: exact match against sweep direction label."""
        assert _passes_filter("column == 'forward'", frozenset(["column"]), "forward") is True
        assert _passes_filter("column == 'forward'", frozenset(["column"]), "backward") is False

    def test_string_inequality(self):
        """String filter: != operator against string axis."""
        assert _passes_filter("column != 'forward'", frozenset(["column"]), "backward") is True
        assert _passes_filter("column != 'forward'", frozenset(["column"]), "forward") is False

    def test_invalid_expr_returns_false(self):
        """Malformed expression returns False rather than raising."""
        assert _passes_filter("vgs >>>", frozenset(["vgs"]), "0.0") is False

    def test_unknown_variable_returns_false(self):
        """Expression referencing an unknown variable returns False."""
        assert _passes_filter("unknown > 0", frozenset(["vgs"]), "0.0") is False