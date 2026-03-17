"""
Tests for chart.py – render_chart and supporting helpers.

All tests operate on in-memory data structures. No actual Excel file is
written by the unit tests; xlsxwriter interaction is verified via the
existing test_writer.py integration tests.

Domain subclass used throughout
--------------------------------
``CurrentMap`` mirrors the "Current Map" fixture in conftest.py and exercises
the subclass-property alias path in the filter system::

    column = ["1.0", "2.0", "3.0"]   → Voltage alias
    row    = ["10.0", "20.0"]         → Current alias
    values = [[0.1, 0.2, 0.3],
              [0.4, 0.5, 0.6]]
"""
from __future__ import annotations

import pytest

from excel_table.models.table_typed import Table2DFloat
from excel_table.models.chart_format import (
    ColorScale,
    FormattedTable2D,
    LineSeriesConfig,
    ChartConfig,
)
from excel_table.chart import (
    _property_names_for_axis,
    _validate_filter_expr,
    _passes_filter,
    _validate_all_filters,
    _resolve_table,
    _col_letter,
    _xl_col_header_range,
    _xl_row_header_range,
    _xl_data_row_range,
    _xl_data_col_range,
)

# ---------------------------------------------------------------------------
# Domain subclass fixture
# ---------------------------------------------------------------------------

class CurrentMap(Table2DFloat):
    """Domain subclass for motor current-voltage maps.

    Aliases:
        voltage → column
        current → row
    """

    @property
    def voltage(self) -> list[str]:
        """Alias for :attr:`column` (voltage axis)."""
        return self.column

    @property
    def current(self) -> list[str]:
        """Alias for :attr:`row` (current axis)."""
        return self.row


@pytest.fixture()
def current_map() -> CurrentMap:
    """Small CurrentMap table matching the conftest Table2D fixture."""
    return CurrentMap(
        title="Current Map",
        column_label="Voltage",
        row_label="Current",
        column=["1.0", "2.0", "3.0"],
        row=["10.0", "20.0"],
        values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],
    )


@pytest.fixture()
def fmt_current_map(current_map: CurrentMap) -> FormattedTable2D:
    """FormattedTable2D wrapping the CurrentMap fixture."""
    return FormattedTable2D(
        table=current_map,
        column_color="#AAAAAA",
        row_color="#BBBBBB",
        value_colorscale=ColorScale(min_color="#FFFFFF", max_color="#FF0000"),
    )


@pytest.fixture()
def plain_table2d() -> Table2DFloat:
    """Plain Table2DFloat without domain subclass (no property aliases)."""
    return Table2DFloat(
        title="Plain Table",
        column_label="Columns",
        row_label="Rows",
        column=["A", "B"],
        row=["X", "Y", "Z"],
        values=[[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]],
    )


@pytest.fixture()
def fmt_plain(plain_table2d: Table2DFloat) -> FormattedTable2D:
    """FormattedTable2D wrapping the plain Table2DFloat fixture."""
    return FormattedTable2D(
        table=plain_table2d,
        column_color="#CCCCCC",
        row_color="#DDDDDD",
        value_colorscale=ColorScale(min_color="#000000", max_color="#0000FF"),
    )


# ---------------------------------------------------------------------------
# _col_letter
# ---------------------------------------------------------------------------

class TestColLetter:
    """Tests for :func:`_col_letter`."""

    def test_a(self):
        """Column 0 → 'A'."""
        assert _col_letter(0) == "A"

    def test_z(self):
        """Column 25 → 'Z'."""
        assert _col_letter(25) == "Z"

    def test_aa(self):
        """Column 26 → 'AA'."""
        assert _col_letter(26) == "AA"

    def test_az(self):
        """Column 51 → 'AZ'."""
        assert _col_letter(51) == "AZ"

    def test_ba(self):
        """Column 52 → 'BA'."""
        assert _col_letter(52) == "BA"


# ---------------------------------------------------------------------------
# _property_names_for_axis
# ---------------------------------------------------------------------------

class TestPropertyNamesForAxis:
    """Tests for :func:`_property_names_for_axis`."""

    def test_plain_table_column(self, plain_table2d):
        """Plain Table2D: only 'column' is valid for column axis."""
        names = _property_names_for_axis(plain_table2d, "column")
        assert "column" in names
        # No extra aliases expected
        assert "row" not in names

    def test_plain_table_row(self, plain_table2d):
        """Plain Table2D: only 'row' is valid for row axis."""
        names = _property_names_for_axis(plain_table2d, "row")
        assert "row" in names
        assert "column" not in names

    def test_current_map_voltage_alias(self, current_map):
        """CurrentMap: 'voltage' and 'column' are both valid for column axis."""
        names = _property_names_for_axis(current_map, "column")
        assert "column" in names
        assert "voltage" in names

    def test_current_map_current_alias(self, current_map):
        """CurrentMap: 'current' and 'row' are both valid for row axis."""
        names = _property_names_for_axis(current_map, "row")
        assert "row" in names
        assert "current" in names

    def test_aliases_do_not_cross_axes(self, current_map):
        """'voltage' must not appear in row names, and 'current' not in column names."""
        row_names = _property_names_for_axis(current_map, "row")
        col_names = _property_names_for_axis(current_map, "column")
        assert "voltage" not in row_names
        assert "current" not in col_names


# ---------------------------------------------------------------------------
# _validate_filter_expr
# ---------------------------------------------------------------------------

class TestValidateFilterExpr:
    """Tests for :func:`_validate_filter_expr`."""

    def test_valid_default_name(self, plain_table2d):
        """Filter using the default axis name passes validation."""
        names = _property_names_for_axis(plain_table2d, "row")
        # Should not raise
        _validate_filter_expr("row >= 0", "row", names, "MySeries")

    def test_valid_alias(self, current_map):
        """Filter using a property alias passes validation."""
        names = _property_names_for_axis(current_map, "row")
        _validate_filter_expr("current >= 10.0", "row", names, "MySeries")

    def test_wrong_axis_raises(self, current_map):
        """Filter referencing a name from the wrong axis raises ValueError."""
        row_names = _property_names_for_axis(current_map, "row")
        # 'voltage' belongs to column axis, not row
        with pytest.raises(ValueError, match="voltage"):
            _validate_filter_expr("voltage <= 2.0", "row", row_names, "MySeries")

    def test_syntax_error_raises(self, plain_table2d):
        """A syntactically invalid expression raises ValueError."""
        names = _property_names_for_axis(plain_table2d, "row")
        with pytest.raises(ValueError, match="syntax error"):
            _validate_filter_expr("row >=", "row", names, "MySeries")

    def test_unknown_name_raises(self, plain_table2d):
        """An expression using an undefined variable name raises ValueError."""
        names = _property_names_for_axis(plain_table2d, "row")
        with pytest.raises(ValueError, match="unknown_var"):
            _validate_filter_expr("unknown_var > 0", "row", names, "MySeries")

    def test_error_message_includes_series_label(self, plain_table2d):
        """ValueError message includes the series label for easy diagnosis."""
        names = _property_names_for_axis(plain_table2d, "row")
        with pytest.raises(ValueError, match="BadSeries"):
            _validate_filter_expr("bad_var > 0", "row", names, "BadSeries")


# ---------------------------------------------------------------------------
# _passes_filter
# ---------------------------------------------------------------------------

class TestPassesFilter:
    """Tests for :func:`_passes_filter`."""

    def test_passes_when_true(self):
        """Expression evaluating to True returns True."""
        assert _passes_filter("row >= 10.0", frozenset({"row"}), 10.0) is True

    def test_fails_when_false(self):
        """Expression evaluating to False returns False."""
        assert _passes_filter("row >= 10.0", frozenset({"row"}), 5.0) is False

    def test_alias_name_works(self):
        """Alias name exposed in namespace evaluates correctly."""
        names = frozenset({"row", "current"})
        assert _passes_filter("current >= 10.0", names, 20.0) is True

    def test_runtime_error_returns_false(self):
        """A runtime error (e.g. division by zero) returns False."""
        assert _passes_filter("row / 0 > 1", frozenset({"row"}), 5.0) is False

    def test_boundary_value_inclusive(self):
        """Boundary values are handled correctly (inclusive >=)."""
        assert _passes_filter("row >= 10.0", frozenset({"row"}), 10.0) is True
        assert _passes_filter("row > 10.0", frozenset({"row"}), 10.0) is False


# ---------------------------------------------------------------------------
# _resolve_table
# ---------------------------------------------------------------------------

class TestResolveTable:
    """Tests for :func:`_resolve_table`."""

    def test_found(self, fmt_current_map):
        """Resolves correctly when title is unique."""
        result = _resolve_table("Current Map", [fmt_current_map], "Sheet1")
        assert result is fmt_current_map

    def test_not_found_raises(self, fmt_current_map):
        """Raises ValueError when source_block title is not in tables."""
        with pytest.raises(ValueError, match="not found"):
            _resolve_table("Nonexistent", [fmt_current_map], "Sheet1")

    def test_duplicate_raises(self, fmt_current_map):
        """Raises ValueError when two tables share the same title."""
        duplicate = FormattedTable2D(
            table=fmt_current_map.table,
            column_color="#000000",
            row_color="#000000",
            value_colorscale=ColorScale(min_color="#000000", max_color="#FFFFFF"),
        )
        with pytest.raises(ValueError, match="ambiguous"):
            _resolve_table("Current Map", [fmt_current_map, duplicate], "Sheet1")

    def test_error_message_includes_sheet_name(self, fmt_current_map):
        """Error message includes sheet name and available titles."""
        with pytest.raises(ValueError, match="MySheet"):
            _resolve_table("Missing", [fmt_current_map], "MySheet")


# ---------------------------------------------------------------------------
# _validate_all_filters
# ---------------------------------------------------------------------------

class TestValidateAllFilters:
    """Tests for :func:`_validate_all_filters`."""

    def test_no_filters_passes(self, fmt_current_map):
        """Chart with no filters passes validation without error."""
        config = ChartConfig(
            chart_type="scatter",
            series=[
                LineSeriesConfig(
                    label="S1",
                    source_block="Current Map",
                    style="scatter",
                )
            ],
            x_label="V",
            y_label="I",
            x_axis="column",
            y_axis="value",
        )
        # Should not raise
        _validate_all_filters(config, [fmt_current_map], "Sheet1")

    def test_valid_alias_filter_passes(self, fmt_current_map):
        """Chart with a valid alias filter passes validation."""
        config = ChartConfig(
            chart_type="scatter",
            series=[
                LineSeriesConfig(
                    label="S1",
                    source_block="Current Map",
                    style="scatter",
                    row_filter="current >= 10.0",
                )
            ],
            x_label="V",
            y_label="I",
            x_axis="column",
            y_axis="value",
        )
        _validate_all_filters(config, [fmt_current_map], "Sheet1")

    def test_wrong_axis_filter_raises(self, fmt_current_map):
        """row_filter referencing a column alias raises ValueError."""
        config = ChartConfig(
            chart_type="scatter",
            series=[
                LineSeriesConfig(
                    label="S1",
                    source_block="Current Map",
                    style="scatter",
                    row_filter="voltage <= 2.0",  # voltage is column axis
                )
            ],
            x_label="V",
            y_label="I",
            x_axis="column",
            y_axis="value",
        )
        with pytest.raises(ValueError, match="voltage"):
            _validate_all_filters(config, [fmt_current_map], "Sheet1")

    def test_missing_source_block_raises(self, fmt_current_map):
        """source_block that does not exist raises ValueError."""
        config = ChartConfig(
            chart_type="scatter",
            series=[
                LineSeriesConfig(
                    label="S1",
                    source_block="Nonexistent Table",
                    style="scatter",
                )
            ],
            x_label="V",
            y_label="I",
            x_axis="column",
            y_axis="value",
        )
        with pytest.raises(ValueError, match="not found"):
            _validate_all_filters(config, [fmt_current_map], "Sheet1")


# ---------------------------------------------------------------------------
# Range builders
# ---------------------------------------------------------------------------

class TestRangeBuilders:
    """Tests for xlsxwriter range helper functions."""

    def test_col_header_range(self, current_map):
        """Column header range is at origin_row+2, origin_col+2 → origin_col+4."""
        r = _xl_col_header_range("Sheet1", 0, 0, current_map)
        # row2, cols 2–4 (0-indexed)
        assert r == ["Sheet1", 2, 2, 2, 4]

    def test_col_header_range_with_offset(self, current_map):
        """Column header range shifts correctly with non-zero origin."""
        r = _xl_col_header_range("Sheet1", 5, 3, current_map)
        assert r == ["Sheet1", 7, 5, 7, 7]

    def test_row_header_range(self, current_map):
        """Row header range is at col origin_col+1, rows origin_row+3 → origin_row+4."""
        r = _xl_row_header_range("Sheet1", 0, 0, current_map)
        assert r == ["Sheet1", 3, 1, 4, 1]

    def test_data_row_range_first(self, current_map):
        """Data row range for row_idx=0 spans the first data row."""
        r = _xl_data_row_range("Sheet1", 0, 0, current_map, 0)
        assert r == ["Sheet1", 3, 2, 3, 4]

    def test_data_row_range_second(self, current_map):
        """Data row range for row_idx=1 spans the second data row."""
        r = _xl_data_row_range("Sheet1", 0, 0, current_map, 1)
        assert r == ["Sheet1", 4, 2, 4, 4]

    def test_data_col_range_first(self, current_map):
        """Data column range for col_idx=0 spans the first data column."""
        r = _xl_data_col_range("Sheet1", 0, 0, current_map, 0)
        assert r == ["Sheet1", 3, 2, 4, 2]

    def test_data_col_range_last(self, current_map):
        """Data column range for col_idx=2 spans the third data column."""
        r = _xl_data_col_range("Sheet1", 0, 0, current_map, 2)
        assert r == ["Sheet1", 3, 4, 4, 4]


# ---------------------------------------------------------------------------
# Series count: filter and color_axis splitting
# ---------------------------------------------------------------------------

class TestSeriesCount:
    """Integration-level tests for the number of series produced.

    These tests call :func:`_add_series_for_config` via a mock chart that
    records ``add_series`` calls, verifying the series splitting logic.
    """

    class _MockChart:
        """Minimal chart stub that records add_series calls."""

        def __init__(self):
            self.series: list[dict] = []

        def add_series(self, opts: dict) -> None:
            """Record a series definition."""
            self.series.append(opts)

        def set_x_axis(self, *a, **kw):
            """No-op."""

        def set_y_axis(self, *a, **kw):
            """No-op."""

        def set_size(self, *a, **kw):
            """No-op."""

    def _run(self, series_cfg, fmt, x_axis="column", y_axis="value"):
        from excel_table.chart import _add_series_for_config
        chart = self._MockChart()
        _add_series_for_config(
            chart=chart,
            sheet_name="Sheet1",
            series_cfg=series_cfg,
            fmt=fmt,
            origin_row=0,
            origin_col=0,
            x_axis=x_axis,
            y_axis=y_axis,
        )
        return chart.series

    def test_no_filter_splits_by_row_count(self, fmt_current_map):
        """Without filters, one series per row (2 rows → 2 series)."""
        cfg = LineSeriesConfig(label="S", source_block="Current Map", style="line")
        series = self._run(cfg, fmt_current_map)
        assert len(series) == 2

    def test_row_filter_reduces_series(self, fmt_current_map):
        """row_filter="current >= 20.0" excludes row "10.0" → 1 series."""
        cfg = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
            row_filter="current >= 20.0",
        )
        series = self._run(cfg, fmt_current_map)
        assert len(series) == 1

    def test_col_filter_does_not_change_series_count(self, fmt_current_map):
        """col_filter reduces data range but does not change series count when x_axis=column."""
        cfg = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
            col_filter="voltage <= 2.0",
        )
        # Series count is driven by rows, not columns
        series = self._run(cfg, fmt_current_map)
        assert len(series) == 2

    def test_color_axis_does_not_change_series_count(self, fmt_current_map):
        """color_axis only affects colors, not the number of series."""
        cfg_with = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
            color_axis="current",
        )
        cfg_without = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
        )
        n_with = len(self._run(cfg_with, fmt_current_map))
        n_without = len(self._run(cfg_without, fmt_current_map))
        assert n_with == n_without

    def test_x_axis_row_splits_by_column(self, fmt_current_map):
        """When x_axis='row', split by column (3 columns → 3 series)."""
        cfg = LineSeriesConfig(label="S", source_block="Current Map", style="line")
        series = self._run(cfg, fmt_current_map, x_axis="row", y_axis="value")
        assert len(series) == 3

    def test_color_axis_assigns_colors(self, fmt_current_map):
        """When color_axis is set and matches the split axis, each series gets a line color."""
        cfg = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
            color_axis="current",
        )
        series = self._run(cfg, fmt_current_map)
        for s in series:
            assert "color" in s.get("line", {}), "Expected line color to be set"

    def test_no_color_axis_no_explicit_color(self, fmt_current_map):
        """Without color_axis, line color is not explicitly set (xlsxwriter auto-colors)."""
        cfg = LineSeriesConfig(label="S", source_block="Current Map", style="line")
        series = self._run(cfg, fmt_current_map)
        for s in series:
            line = s.get("line", {})
            assert "color" not in line or line == {}, (
                "Expected no explicit color without color_axis"
            )

    def test_all_rows_filtered_out_produces_no_series(self, fmt_current_map):
        """Filter that excludes all rows produces zero series."""
        cfg = LineSeriesConfig(
            label="S",
            source_block="Current Map",
            style="line",
            row_filter="current > 9999",
        )
        series = self._run(cfg, fmt_current_map)
        assert len(series) == 0