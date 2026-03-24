"""
Round-trip integration tests for excel-table.

Write → Read → verify: the core value proposition of excel-table is that
structured tables can be written to Excel and read back with the same
structure and values. These tests verify that contract end-to-end using
the transistor measurement domain from the demo app.

Fixture dependency chain
------------------------
conftest.py generates measurement data programmatically via the transistor
simulator, builds Excel bytes via write_sheet_bytes, and feeds those bytes
to the reader. No static files or manual data entry is involved.

    TransistorModel
        → Analyzer.sweep_IV / sweep_Vgs
        → filled_bytes_1device / filled_bytes_2devices  (writer)
            → read_input()                               (reader)
                → assertions
"""
from __future__ import annotations

import numpy as np
import pytest

from excel_table.reader import SheetReadSchema, read_sheet_bytes
from excel_table.models import Table2DFloat, TableKeyValue
from excel_table.models.table_format import FormattedTable2DSchema, TableKeyValueSchema

from transistor.converter import iv_to_xarray, transfer_to_xarray

from tests.conftest import (
    SHEET_NAME,
    DEVICE_A_W_UM, DEVICE_A_L_UM,
    DEVICE_B_W_UM, DEVICE_B_L_UM,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def read_input(data: bytes) -> list[list]:
    """Parse the transistor input template — mirrors Page 4's read_input."""
    schema = SheetReadSchema(
        columns=[
            TableKeyValueSchema(title="Model Params"),
            FormattedTable2DSchema(title="IV Result", table_type=Table2DFloat),
            FormattedTable2DSchema(title="Transfer Result", table_type=Table2DFloat),
        ]
    )
    return read_sheet_bytes(data, SHEET_NAME, schema)


# ---------------------------------------------------------------------------
# Blank template (write only — no value assertions)
# ---------------------------------------------------------------------------

class TestBlankTemplate:
    """write_sheet_bytes produces a parseable blank template."""

    def test_1device_parses(self, template_bytes_1device):
        """Blank 1-device template can be parsed without error."""
        result = read_input(template_bytes_1device)
        print(type(result[0][1].column[0]))
        [print("foo!") for _ in range(100)]
        assert len(result) == 1

    def test_2devices_parses(self, template_bytes_2devices):
        """Blank 2-device template yields 2 rows."""
        result = read_input(template_bytes_2devices)
        assert len(result) == 2

    def test_blank_iv_values_are_none(self, template_bytes_1device):
        """Blank template IV cells are read back as None."""
        result = read_input(template_bytes_1device)
        iv: Table2DFloat = result[0][1]
        assert all(v is None for row in iv.values for v in row)

    def test_blank_transfer_values_are_none(self, template_bytes_1device):
        """Blank template Transfer cells are read back as None."""
        result = read_input(template_bytes_1device)
        tr: Table2DFloat = result[0][2]
        assert all(v is None for row in tr.values for v in row)


# ---------------------------------------------------------------------------
# Model Params round-trip
# ---------------------------------------------------------------------------

class TestModelParamsRoundTrip:
    """TableKeyValue write → read preserves W and L values."""

    def test_1device_gate_width(self, filled_bytes_1device):
        """GateWidth is read back correctly for device A."""
        result = read_input(filled_bytes_1device)
        params: TableKeyValue = result[0][0]
        assert float(params.value[0]) == pytest.approx(DEVICE_A_W_UM)

    def test_1device_gate_length(self, filled_bytes_1device):
        """GateLength is read back correctly for device A."""
        result = read_input(filled_bytes_1device)
        params: TableKeyValue = result[0][0]
        assert float(params.value[1]) == pytest.approx(DEVICE_A_L_UM)

    def test_2devices_params(self, filled_bytes_2devices):
        """Two devices return correct W/L values in order."""
        result = read_input(filled_bytes_2devices)
        params_a: TableKeyValue = result[0][0]
        params_b: TableKeyValue = result[1][0]
        assert float(params_a.value[0]) == pytest.approx(DEVICE_A_W_UM)
        assert float(params_b.value[0]) == pytest.approx(DEVICE_B_W_UM)


# ---------------------------------------------------------------------------
# IV round-trip
# ---------------------------------------------------------------------------

class TestIVRoundTrip:
    """FormattedTable2D(IV) write → read preserves axes and values."""

    def test_column_axis(self, filled_bytes_1device, iv_a):
        """Vgs column axis is read back with correct values."""
        result = read_input(filled_bytes_1device)
        iv: Table2DFloat = result[0][1]
        expected = [v for v in iv_a.coords["vgs"].values]
        np.testing.assert_allclose(
            np.array(iv.column, dtype=np.float64),
            expected,
            rtol=1e-6,
        )

    def test_row_axis(self, filled_bytes_1device, iv_a):
        """Vds row axis is read back with correct values."""
        result = read_input(filled_bytes_1device)
        iv: Table2DFloat = result[0][1]
        expected = [v for v in iv_a.coords["vds"].values]
        np.testing.assert_allclose(
            np.array(iv.row, dtype=np.float64),
            expected,
            rtol=1e-6,
        )

    def test_values_shape(self, filled_bytes_1device, iv_a):
        """Read-back values have correct shape (n_vds × n_vgs)."""
        result = read_input(filled_bytes_1device)
        iv: Table2DFloat = result[0][1]
        n_vds = len(iv_a.coords["vds"])
        n_vgs = len(iv_a.coords["vgs"])
        assert len(iv.values) == n_vds
        assert len(iv.values[0]) == n_vgs

    def test_values_roundtrip(self, filled_bytes_1device, iv_a):
        """IV values survive write → read with float precision."""
        result = read_input(filled_bytes_1device)
        iv: Table2DFloat = result[0][1]
        recovered = iv_to_xarray(iv)
        np.testing.assert_allclose(
            recovered.values,
            iv_a.values,
            rtol=1e-6,
        )

    def test_2devices_iv_values_independent(self, filled_bytes_2devices, iv_a, iv_b):
        """Two devices' IV values are independent after round-trip."""
        result = read_input(filled_bytes_2devices)
        recovered_a = iv_to_xarray(result[0][1])
        recovered_b = iv_to_xarray(result[1][1])
        np.testing.assert_allclose(recovered_a.values, iv_a.values, rtol=1e-6)
        np.testing.assert_allclose(recovered_b.values, iv_b.values, rtol=1e-6)


# ---------------------------------------------------------------------------
# Transfer round-trip
# ---------------------------------------------------------------------------

class TestTransferRoundTrip:
    """FormattedTable2D(Transfer) write → read preserves axes and values."""

    def test_column_axis_labels(self, filled_bytes_1device):
        """Transfer column axis contains 'forward' and 'backward'."""
        result = read_input(filled_bytes_1device)
        tr: Table2DFloat = result[0][2]
        assert tr.column == ["forward", "backward"]

    def test_row_axis(self, filled_bytes_1device, transfer_a):
        """Vgs row axis is read back with correct values."""
        result = read_input(filled_bytes_1device)
        tr: Table2DFloat = result[0][2]
        expected = [v for v in transfer_a.coords["vgs"].values]
        np.testing.assert_allclose(
            np.array(tr.row, dtype=np.float64),
            expected,
            rtol=1e-6,
        )

    def test_values_shape(self, filled_bytes_1device, transfer_a):
        """Read-back values have correct shape (n_vgs × 2)."""
        result = read_input(filled_bytes_1device)
        tr: Table2DFloat = result[0][2]
        n_vgs = len(transfer_a.coords["vgs"])
        assert len(tr.values) == n_vgs
        assert len(tr.values[0]) == 2

    def test_values_roundtrip(self, filled_bytes_1device, transfer_a):
        """Transfer values survive write → read with float precision."""
        result = read_input(filled_bytes_1device)
        tr: Table2DFloat = result[0][2]
        recovered = transfer_to_xarray(tr)
        np.testing.assert_allclose(
            recovered.values,
            transfer_a.values,
            rtol=1e-6,
        )

    def test_2devices_transfer_values_independent(
        self, filled_bytes_2devices, transfer_a, transfer_b
    ):
        """Two devices' Transfer values are independent after round-trip."""
        result = read_input(filled_bytes_2devices)
        recovered_a = transfer_to_xarray(result[0][2])
        recovered_b = transfer_to_xarray(result[1][2])
        np.testing.assert_allclose(recovered_a.values, transfer_a.values, rtol=1e-6)
        np.testing.assert_allclose(recovered_b.values, transfer_b.values, rtol=1e-6)