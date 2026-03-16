"""
Tests for the sheet reader.

Fixtures are built with openpyxl directly (not via the writer)
to keep reader and writer tests fully independent.
"""
import pytest

from excel_table.reader import SheetReadSchema, read_sheet
from excel_table.models.table import (
    Table1DFloat,
    Table2DFloat,
    Table2DStr,
    TableKeyValue,
)
from tests.conftest import (
    TABLE2D_TITLE, TABLE2D_COLUMN_LABEL, TABLE2D_ROW_LABEL,
    TABLE2D_COLUMNS, TABLE2D_ROWS, TABLE2D_VALUES,
    TABLE1D_TITLE, TABLE1D_COLUMN_LABEL, TABLE1D_COLUMNS, TABLE1D_VALUES,
    KV_TITLE, KV_COLUMNS, KV_VALUES,
)


# ---------------------------------------------------------------------------
# Table2D
# ---------------------------------------------------------------------------

class TestReadTable2D:
    """Read accuracy for each field of a Table2D."""

    def _schema(self):
        """Return a minimal SheetReadSchema targeting the standard Table2D fixture."""
        return SheetReadSchema(rows=[[(TABLE2D_TITLE, Table2DFloat)]])

    def test_title(self, table2d_xlsx):
        """title is read correctly from the anchor cell."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].title == TABLE2D_TITLE

    def test_column_label(self, table2d_xlsx):
        """column_label is read from the cell at (r+1, c+2)."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].column_label == TABLE2D_COLUMN_LABEL

    def test_row_label(self, table2d_xlsx):
        """row_label is read from the cell at (r+3, c)."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].row_label == TABLE2D_ROW_LABEL

    def test_columns(self, table2d_xlsx):
        """column values are read right from (r+2, c+2) until None."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].column == TABLE2D_COLUMNS

    def test_rows(self, table2d_xlsx):
        """row values are read down from (r+3, c+1) until None."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].row == TABLE2D_ROWS

    def test_values(self, table2d_xlsx):
        """data values are read from the row×col region starting at (r+3, c+2)."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert result[0][0].values == TABLE2D_VALUES

    def test_returns_correct_type(self, table2d_xlsx):
        """read_sheet returns an instance of the type specified in the schema."""
        result = read_sheet(table2d_xlsx, "Sheet", self._schema())
        assert isinstance(result[0][0], Table2DFloat)

    def test_str_subclass_coerces_float_values(self, table2d_xlsx):
        """Table2DStr coerces float cell values returned by openpyxl to str."""
        schema = SheetReadSchema(rows=[[(TABLE2D_TITLE, Table2DStr)]])
        result = read_sheet(table2d_xlsx, "Sheet", schema)
        assert isinstance(result[0][0].values[0][0], str)

    def test_title_not_found_raises(self, table2d_xlsx):
        """A title absent from the sheet raises ValueError."""
        schema = SheetReadSchema(rows=[[("Nonexistent", Table2DFloat)]])
        with pytest.raises(ValueError, match="not found"):
            read_sheet(table2d_xlsx, "Sheet", schema)

    def test_none_cells_preserved(self, table2d_none_xlsx):
        """Empty cells in the data region are returned as None."""
        schema = SheetReadSchema(rows=[[(TABLE2D_TITLE, Table2DFloat)]])
        result = read_sheet(table2d_none_xlsx, "Sheet", schema)
        values = result[0][0].values
        assert values[0][1] is None  # D4 was empty
        assert values[1][0] is None  # C5 was empty
        assert values[1][2] is None  # E5 was empty

    def test_duplicate_titles_both_read(self, table2d_duplicate_xlsx):
        """Two tables sharing a title can both be read by listing the title twice in the schema."""
        schema = SheetReadSchema(rows=[
            [(TABLE2D_TITLE, Table2DFloat), (TABLE2D_TITLE, Table2DFloat)],
        ])
        result = read_sheet(table2d_duplicate_xlsx, "Sheet", schema)
        assert result[0][0].title == TABLE2D_TITLE
        assert result[0][1].title == TABLE2D_TITLE


# ---------------------------------------------------------------------------
# Table1D
# ---------------------------------------------------------------------------

class TestReadTable1D:
    """Read accuracy for each field of a Table1D."""

    def _schema(self):
        """Return a minimal SheetReadSchema targeting the standard Table1D fixture."""
        return SheetReadSchema(rows=[[(TABLE1D_TITLE, Table1DFloat)]])

    def test_title(self, table1d_xlsx):
        """title is read correctly from the anchor cell."""
        result = read_sheet(table1d_xlsx, "Sheet", self._schema())
        assert result[0][0].title == TABLE1D_TITLE

    def test_column_label(self, table1d_xlsx):
        """column_label is read from the cell at (r+1, c)."""
        result = read_sheet(table1d_xlsx, "Sheet", self._schema())
        assert result[0][0].column_label == TABLE1D_COLUMN_LABEL

    def test_columns(self, table1d_xlsx):
        """column values are read right from (r+2, c) until None."""
        result = read_sheet(table1d_xlsx, "Sheet", self._schema())
        assert result[0][0].column == TABLE1D_COLUMNS

    def test_values(self, table1d_xlsx):
        """data values are read from row (r+3) across all columns."""
        result = read_sheet(table1d_xlsx, "Sheet", self._schema())
        assert result[0][0].values == TABLE1D_VALUES


# ---------------------------------------------------------------------------
# TableKeyValue
# ---------------------------------------------------------------------------

class TestReadTableKeyValue:
    """Read accuracy for TableKeyValue."""

    def _schema(self):
        """Return a minimal SheetReadSchema targeting the standard TableKeyValue fixture."""
        return SheetReadSchema(rows=[[(KV_TITLE, TableKeyValue)]])

    def test_columns(self, table_kv_xlsx):
        """column (key) values are read from row (r+1)."""
        result = read_sheet(table_kv_xlsx, "Sheet", self._schema())
        assert result[0][0].column == KV_COLUMNS

    def test_values(self, table_kv_xlsx):
        """value strings are read from row (r+2)."""
        result = read_sheet(table_kv_xlsx, "Sheet", self._schema())
        assert result[0][0].value == KV_VALUES


# ---------------------------------------------------------------------------
# Error cases
# ---------------------------------------------------------------------------

def test_invalid_sheet_name_raises(table2d_xlsx):
    """Requesting a sheet that does not exist raises KeyError."""
    schema = SheetReadSchema(rows=[[(TABLE2D_TITLE, Table2DFloat)]])
    with pytest.raises(KeyError):
        read_sheet(table2d_xlsx, "DoesNotExist", schema)