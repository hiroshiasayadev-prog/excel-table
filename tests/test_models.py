"""
Tests for data model validation.

Covers shape validation, type coercion, domain subclass patterns,
and TableKeyValue cast contracts.
"""
import pytest

from excel_table.models import (
    Table2DFloat,
    Table2DInt,
    Table2DStr,
    Table1DFloat,
    TableKeyValue,
    FormattedTable2D,
    ColorScale,
)
from excel_table.models.table_base import Table2D


# ---------------------------------------------------------------------------
# Table2D
# ---------------------------------------------------------------------------

class TestTable2DValidation:
    """Shape validation and type coercion for Table2D typed subclasses."""

    def test_valid(self):
        """A correctly shaped Table2DFloat is created without error."""
        t = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["-0.40", "0.00", "0.40"],
            row=["0.00", "0.50"],
            values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],
        )
        assert t.title == "IV Result"
        assert len(t.values) == 2

    def test_single_column_raises(self):
        """Table2D requires at least 2 columns — use Table1D for single-axis data."""
        with pytest.raises(ValueError, match="at least 2 columns"):
            Table2DFloat(
                title="IV Result",
                column_label="Vgs [V]",
                row_label="Vds [V]",
                column=["0.00"],
                row=["0.00", "0.50"],
                values=[[0.1], [0.2]],
            )

    def test_single_row_raises(self):
        """Table2D requires at least 2 rows — use Table1D for single-axis data."""
        with pytest.raises(ValueError, match="at least 2 rows"):
            Table2DFloat(
                title="IV Result",
                column_label="Vgs [V]",
                row_label="Vds [V]",
                column=["0.00", "0.50"],
                row=["0.00"],
                values=[[0.1, 0.2]],
            )

    def test_values_row_count_mismatch_raises(self):
        """values with fewer rows than len(row) raises ValueError."""
        with pytest.raises(ValueError, match="values has"):
            Table2DFloat(
                title="IV Result",
                column_label="Vgs [V]",
                row_label="Vds [V]",
                column=["0.00", "0.50"],
                row=["0.00", "0.50"],
                values=[[0.1, 0.2]],  # 1 row, expected 2
            )

    def test_values_col_count_mismatch_raises(self):
        """values row with wrong column count raises ValueError."""
        with pytest.raises(ValueError, match=r"values\[0\]"):
            Table2DFloat(
                title="IV Result",
                column_label="Vgs [V]",
                row_label="Vds [V]",
                column=["0.00", "0.50"],
                row=["0.00", "0.50"],
                values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],  # 3 cols, expected 2
            )

    def test_column_accepts_str(self):
        """Table2D.column is list[str] — string values pass through as-is."""
        t = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["-0.40", "0.00"],
            row=["0.00", "0.50"],
            values=[[0.1, 0.2], [0.3, 0.4]],
        )
        assert t.column == ["-0.40", "0.00"]

    def test_string_axis_labels_accepted(self):
        """Table2D.column accepts string labels like 'forward'/'backward'."""
        t = Table2DFloat(
            title="Transfer Result",
            column_label="Sweep Direction",
            row_label="Vgs [V]",
            column=["forward", "backward"],
            row=["-0.50", "0.00", "0.50"],
            values=[[0.1, 0.2], [0.3, 0.4], [0.5, 0.6]],
        )
        assert t.column == ["forward", "backward"]

    def test_none_values_allowed(self):
        """None is a valid cell value (blank cell in the template)."""
        t = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["-0.40", "0.00"],
            row=["0.00", "0.50"],
            values=[[None, 0.2], [0.3, None]],
        )
        assert t.values[0][0] is None
        assert t.values[1][1] is None

    def test_int_subclass_truncates_float(self):
        """Table2DInt truncates float cell values to int."""
        t = Table2DInt(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["0.00", "0.50"],
            row=["0.00", "0.50"],
            values=[[1.7, 2.9], [3.1, 4.8]],
        )
        assert t.values[0] == [1, 2]

    def test_str_subclass_coerces_numeric(self):
        """Table2DStr coerces float and int cell values to str."""
        t = Table2DStr(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["0.00", "0.50"],
            row=["0.00", "0.50"],
            values=[[0.1, 2], [3, 0.4]],
        )
        assert isinstance(t.values[0][0], str)

    def test_model_copy_preserves_subclass(self):
        """model_copy preserves the concrete subclass type."""
        t = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=["-0.40", "0.00"],
            row=["0.00", "0.50"],
            values=[[0.1, 0.2], [0.3, 0.4]],
        )
        t2 = t.model_copy(update={"column": ["0.20", "0.40"]})
        assert isinstance(t2, Table2DFloat)
        assert t2.column == ["0.20", "0.40"]


# ---------------------------------------------------------------------------
# Table1D
# ---------------------------------------------------------------------------

class TestTable1DValidation:
    """Shape validation for Table1D typed subclasses."""

    def test_valid(self):
        """A correctly shaped Table1DFloat is created without error."""
        t = Table1DFloat(
            title="Vgs Points",
            column_label="Vgs [V]",
            column=["-0.50", "0.00", "0.50"],
            values=[[0.1, 0.2, 0.3]],
        )
        assert len(t.column) == 3

    def test_col_count_mismatch_raises(self):
        """values row with wrong column count raises ValueError."""
        with pytest.raises(ValueError, match=r"values\[0\]"):
            Table1DFloat(
                title="Vgs Points",
                column_label="Vgs [V]",
                column=["-0.50", "0.00"],
                values=[[0.1, 0.2, 0.3]],  # 3 cols, expected 2
            )


# ---------------------------------------------------------------------------
# TableKeyValue
# ---------------------------------------------------------------------------

class TestTableKeyValueValidation:
    """Shape validation and cast contracts for TableKeyValue."""

    def test_valid_str_values(self):
        """TableKeyValue with string values is created without error."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=["100.0", "1.0"],
        )
        assert t.value == ["100.0", "1.0"]

    def test_valid_numeric_values(self):
        """TableKeyValue accepts int and float values."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=[100.0, 1.0],
        )
        assert t.value[0] == 100.0

    def test_none_value_allowed(self):
        """None is a valid value (blank cell — used in blank templates)."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=[None, None],
        )
        assert t.value[0] is None

    def test_length_mismatch_raises(self):
        """column and value with different lengths raises ValueError."""
        with pytest.raises(ValueError, match="column length"):
            TableKeyValue(
                title="Model Params",
                column=["GateWidth [um]"],
                value=[100.0, 1.0],
            )

    def test_read_contract_values_are_str(self):
        """Read contract: values returned from reader are str regardless of Excel type.
        Simulated here by constructing with str — caller must convert to float."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=["100.0", "1.0"],  # as returned by reader
        )
        # caller's responsibility to convert
        W_um = float(t.value[0])
        L_um = float(t.value[1])
        assert W_um == 100.0
        assert L_um == 1.0


# ---------------------------------------------------------------------------
# ColorScale
# ---------------------------------------------------------------------------

class TestColorScale:
    """Construction of 2-color and 3-color scales."""

    def test_two_color(self):
        """ColorScale without mid_color produces a 2-color scale."""
        cs = ColorScale(min_color="#FFFFFF", max_color="#FF5722")
        assert cs.mid_color is None

    def test_three_color(self):
        """ColorScale with mid_color produces a 3-color scale."""
        cs = ColorScale(min_color="#FFFFFF", mid_color="#FFF176", max_color="#FF5722")
        assert cs.mid_color == "#FFF176"

# ---------------------------------------------------------------------------
# Table2D.to_dataarray
# ---------------------------------------------------------------------------

class TestTable2DToDataarray:
    """to_dataarray conversion for Table2D."""

    def setup_method(self):
        self.table = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=[-0.4, 0.0, 0.4],
            row=[0.0, 0.5],
            values=[[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]],
        )

    def test_dims(self):
        """dims must be ("row", "column")."""
        import xarray as xr
        da = self.table.to_dataarray()
        assert da.dims == ("row", "column")

    def test_coords(self):
        """coords must match row and column axes."""
        import numpy as np
        da = self.table.to_dataarray()
        np.testing.assert_array_equal(da.coords["row"].values, [0.0, 0.5])
        np.testing.assert_array_equal(da.coords["column"].values, [-0.4, 0.0, 0.4])

    def test_attrs(self):
        """attrs must contain title, row_label, column_label."""
        da = self.table.to_dataarray()
        assert da.attrs["title"] == "IV Result"
        assert da.attrs["row_label"] == "Vds [V]"
        assert da.attrs["column_label"] == "Vgs [V]"

    def test_values(self):
        """values must match the original table values."""
        import numpy as np
        da = self.table.to_dataarray()
        np.testing.assert_array_equal(da.values, [[0.1, 0.2, 0.3], [0.4, 0.5, 0.6]])

    def test_dtype_float64_converts_none_to_nan(self):
        """dtype=np.float64 converts None to nan."""
        import numpy as np
        table = Table2DFloat(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=[-0.4, 0.0, 0.4],
            row=[0.0, 0.5],
            values=[[None, 0.2, 0.3], [0.4, None, 0.6]],
        )
        da = table.to_dataarray(dtype=np.float64)
        assert np.isnan(da.values[0][0])
        assert np.isnan(da.values[1][1])
        assert da.dtype == np.float64

    def test_transpose(self):
        """da.T reverses dims to ("column", "row")."""
        da = self.table.to_dataarray()
        assert da.T.dims == ("column", "row")

    def test_string_axis(self):
        """String axis values (e.g. 'forward'/'backward') are supported."""
        import numpy as np
        table = Table2DFloat(
            title="Transfer",
            column_label="Sweep",
            row_label="Vgs [V]",
            column=["forward", "backward"],
            row=[-0.5, 0.0, 0.5],
            values=[[0.1, 0.2], [0.3, 0.4], [0.5, 0.6]],
        )
        da = table.to_dataarray()
        np.testing.assert_array_equal(da.coords["column"].values, ["forward", "backward"])


# ---------------------------------------------------------------------------
# TableKeyValue.to_dict
# ---------------------------------------------------------------------------

class TestTableKeyValueToDict:
    """to_dict conversion for TableKeyValue."""

    def test_basic(self):
        """to_dict returns a dict mapping column keys to values."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=[100.0, 1.0],
        )
        assert t.to_dict() == {"GateWidth [um]": 100.0, "GateLength [um]": 1.0}

    def test_none_value(self):
        """None values are preserved in the dict."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=[None, None],
        )
        assert t.to_dict() == {"GateWidth [um]": None, "GateLength [um]": None}

    def test_str_values(self):
        """str values (as returned by reader) are preserved."""
        t = TableKeyValue(
            title="Model Params",
            column=["GateWidth [um]", "GateLength [um]"],
            value=["100.0", "1.0"],
        )
        d = t.to_dict()
        assert d["GateWidth [um]"] == "100.0"