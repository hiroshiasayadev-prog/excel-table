import numpy as np
import xarray as xr
from excel_table.models import Table2DFloat


def iv_to_xarray(iv: Table2DFloat) -> xr.DataArray:
    return xr.DataArray(
        data=np.array(iv.values, dtype=np.float64).T,
        dims=["vgs", "vds"],
        coords={
            "vgs": np.array(iv.column, dtype=np.float64),
            "vds": np.array(iv.row, dtype=np.float64),
        },
        attrs={"units": "A"},
    )


def transfer_to_xarray(tr: Table2DFloat) -> xr.DataArray:
    return xr.DataArray(
        data=np.stack([
            np.array(tr.values, dtype=np.float64)[:, 0],
            np.array(tr.values, dtype=np.float64)[:, 1],
        ]),
        dims=["sweep", "vgs"],
        coords={
            "sweep": ["forward", "backward"],
            "vgs": np.array(tr.row, dtype=np.float64),
        },
    )


def gm_from_transfer(jd_tr: xr.DataArray) -> xr.DataArray:
    """Compute transconductance from transfer current density.

    Derives gm = d(Jd)/d(Vgs) for each sweep direction.

    Args:
        jd_tr: Transfer current density DataArray in mA/mm.
            dims=["sweep", "vgs"], coords={"sweep": [...], "vgs": [...]}

    Returns:
        DataArray with the same dims/coords as ``jd_tr``, values in mS/mm.
    """
    vgs = jd_tr.coords["vgs"].values
    gm_data = np.stack([
        np.gradient(jd_tr.sel(sweep=s).values, vgs)
        for s in jd_tr.coords["sweep"].values
    ])
    return xr.DataArray(
        data=gm_data,
        dims=jd_tr.dims,
        coords=jd_tr.coords,
    )