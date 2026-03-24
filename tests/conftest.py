"""
Shared pytest fixtures for excel_table tests.

Round-trip philosophy
---------------------
These fixtures are designed around the write → read round-trip that is
excel-table's core value proposition. Rather than building xlsx files
directly with openpyxl (which would test the reader in isolation),
fixtures use the writer to produce xlsx bytes, then feed them to the
reader. This ensures that write and read stay mutually consistent.

Transistor domain
-----------------
The transistor measurement domain (GaAs HEMT, IV + Transfer characteristics)
is used as a realistic stand-in for the "instrument CSV → Excel → excel-table"
use case that the demo app demonstrates. Device parameters and sweep
conditions are defined as module-level constants so that tests can assert
on exact values.
"""
from __future__ import annotations

import numpy as np
import pytest
import xarray as xr

from transistor import Analyzer, TransistorModel
from transistor.converter import iv_to_xarray, transfer_to_xarray

from excel_table.models import (
    Table2D,
    FormattedTable2D,
    TableKeyValue,
)
from excel_table.writer import SheetWriteSchema, write_sheet_bytes

# ---------------------------------------------------------------------------
# Device constants
# ---------------------------------------------------------------------------

DEVICE_A_W_UM = 100.0
DEVICE_A_L_UM = 1.0

DEVICE_B_W_UM = 200.0
DEVICE_B_L_UM = 0.5

SHEET_NAME = "Transistor Input"

# Sweep conditions — kept small for test speed
VDS_FROM   = 0.0
VDS_UNTIL  = 0.5
VDS_STEP   = 0.1
VGS_FROM_IV  = -0.4
VGS_UNTIL_IV = 0.4
VGS_STEP_IV  = 0.2

VGS_FROM_TR  = -0.5
VGS_UNTIL_TR = 0.5
VGS_STEP_TR  = 0.1
VDS_TR       = 1.0
DT           = 1e-4


# ---------------------------------------------------------------------------
# Model helpers
# ---------------------------------------------------------------------------

def _make_model(W_um: float, L_um: float) -> TransistorModel:
    m = TransistorModel()
    m.W = W_um * 1e-6
    m.L = L_um * 1e-6
    return m


def _sweep_iv(model: TransistorModel) -> xr.DataArray:
    return Analyzer.sweep_IV(
        transistor=model,
        vds_from=VDS_FROM, vds_until=VDS_UNTIL, vds_step=VDS_STEP,
        vgs_from=VGS_FROM_IV, vgs_until=VGS_UNTIL_IV, vgs_step=VGS_STEP_IV,
    )


def _sweep_transfer(model: TransistorModel) -> xr.DataArray:
    return Analyzer.sweep_Vgs(
        transistor=model,
        vgs_from=VGS_FROM_TR, vgs_until=VGS_UNTIL_TR, vgs_step=VGS_STEP_TR,
        vds=VDS_TR, dt=DT,
    )


# ---------------------------------------------------------------------------
# Single-device fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def model_a() -> TransistorModel:
    """Device A: W=100um, L=1um."""
    return _make_model(DEVICE_A_W_UM, DEVICE_A_L_UM)


@pytest.fixture
def iv_a(model_a) -> xr.DataArray:
    """IV sweep DataArray for device A."""
    return _sweep_iv(model_a)


@pytest.fixture
def transfer_a(model_a) -> xr.DataArray:
    """Transfer sweep DataArray for device A."""
    return _sweep_transfer(model_a)


# ---------------------------------------------------------------------------
# Two-device fixtures (num_devices=2 round-trip)
# ---------------------------------------------------------------------------

@pytest.fixture
def model_b() -> TransistorModel:
    """Device B: W=200um, L=0.5um."""
    return _make_model(DEVICE_B_W_UM, DEVICE_B_L_UM)


@pytest.fixture
def iv_b(model_b) -> xr.DataArray:
    return _sweep_iv(model_b)


@pytest.fixture
def transfer_b(model_b) -> xr.DataArray:
    return _sweep_transfer(model_b)


# ---------------------------------------------------------------------------
# Excel template bytes fixtures
# ---------------------------------------------------------------------------

def _build_template_row(
    W_um: float,
    L_um: float,
    iv: xr.DataArray,
    transfer: xr.DataArray,
) -> list:
    """Build one row of [TableKeyValue, FormattedTable2D, FormattedTable2D]."""
    vds = [v for v in iv.coords["vds"].values]
    vgs = [v for v in iv.coords["vgs"].values]
    tr_vgs = [v for v in transfer.coords["vgs"].values]

    model_params = TableKeyValue(
        title="Model Params",
        column=["GateWidth [um]", "GateLength [um]"],
        value=[W_um, L_um],
    )
    iv_table = FormattedTable2D(
        table=Table2D(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=vgs,
            row=vds,
            values=np.full((len(vds), len(vgs)), None).tolist(),
        )
    )
    transfer_table = FormattedTable2D(
        table=Table2D(
            title="Transfer Result",
            column_label="Sweep Direction",
            row_label="Vgs [V]",
            column=["forward", "backward"],
            row=tr_vgs,
            values=np.full((len(tr_vgs), 2), None).tolist(),
        )
    )
    return [model_params, iv_table, transfer_table]


@pytest.fixture
def template_bytes_1device(iv_a, transfer_a) -> bytes:
    """Blank input template for 1 device (device A axes, values=None)."""
    row = _build_template_row(DEVICE_A_W_UM, DEVICE_A_L_UM, iv_a, transfer_a)
    schema = SheetWriteSchema(rows=[row])
    return write_sheet_bytes(sheet_name=SHEET_NAME, schema=schema)


@pytest.fixture
def template_bytes_2devices(iv_a, transfer_a, iv_b, transfer_b) -> bytes:
    """Blank input template for 2 devices."""
    row_a = _build_template_row(DEVICE_A_W_UM, DEVICE_A_L_UM, iv_a, transfer_a)
    row_b = _build_template_row(DEVICE_B_W_UM, DEVICE_B_L_UM, iv_b, transfer_b)
    schema = SheetWriteSchema(rows=[row_a, row_b])
    return write_sheet_bytes(sheet_name=SHEET_NAME, schema=schema)


# ---------------------------------------------------------------------------
# Filled template bytes fixtures (values populated for round-trip tests)
# ---------------------------------------------------------------------------

def _fill_template(
    W_um: float,
    L_um: float,
    iv: xr.DataArray,
    transfer: xr.DataArray,
) -> list:
    """Build one row with actual measurement values filled in."""
    vds = [v for v in iv.coords["vds"].values]
    vgs = [v for v in iv.coords["vgs"].values]
    tr_vgs = [v for v in transfer.coords["vgs"].values]

    model_params = TableKeyValue(
        title="Model Params",
        column=["GateWidth [um]", "GateLength [um]"],
        value=[W_um, L_um],
    )
    iv_table = FormattedTable2D(
        table=Table2D(
            title="IV Result",
            column_label="Vgs [V]",
            row_label="Vds [V]",
            column=vgs,
            row=vds,
            values=iv.values.T.tolist(),
        )
    )
    transfer_table = FormattedTable2D(
        table=Table2D(
            title="Transfer Result",
            column_label="Sweep Direction",
            row_label="Vgs [V]",
            column=["forward", "backward"],
            row=tr_vgs,
            values=transfer.values.T.tolist(),
        )
    )
    return [model_params, iv_table, transfer_table]


@pytest.fixture
def filled_bytes_1device(iv_a, transfer_a) -> bytes:
    """Filled input template for 1 device (device A)."""
    row = _fill_template(DEVICE_A_W_UM, DEVICE_A_L_UM, iv_a, transfer_a)
    schema = SheetWriteSchema(rows=[row])
    return write_sheet_bytes(sheet_name=SHEET_NAME, schema=schema)


@pytest.fixture
def filled_bytes_2devices(iv_a, transfer_a, iv_b, transfer_b) -> bytes:
    """Filled input template for 2 devices."""
    row_a = _fill_template(DEVICE_A_W_UM, DEVICE_A_L_UM, iv_a, transfer_a)
    row_b = _fill_template(DEVICE_B_W_UM, DEVICE_B_L_UM, iv_b, transfer_b)
    schema = SheetWriteSchema(rows=[row_a, row_b])
    return write_sheet_bytes(sheet_name=SHEET_NAME, schema=schema)