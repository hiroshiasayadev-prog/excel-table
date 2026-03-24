"""
Convert measurement DataArrays to CSV-compatible list[list].
1 measurement = 1 file (instrument convention).
"""

import xarray as xr


def iv_to_list(iv: xr.DataArray, W: float, L: float) -> list[list]:
    """
    Convert IV sweep DataArray to a list of rows.

    Expected iv structure:
        dims=["vgs", "vds"]
        coords={"vgs": ..., "vds": ...}
        attrs={"units": "A", ...}
    """
    vgs_arr = iv.coords["vgs"].values
    vds_arr = iv.coords["vds"].values

    rows = []

    # --- header ---
    rows.append(["AnalysisType", "IV"])
    rows.append(["ModelParams"])
    rows.append(["GateWidth", W*1e6, "um"])
    rows.append(["GateLength", L*1e6, "um"])
    rows.append(["Variables"])
    rows.append(["Ids", "A", "target"])
    rows.append(["Vds", "V", "sweep"])
    rows.append(["Vgs", "V", "sweep"])

    # --- data ---
    rows.append(["", "Vgs"] + [""] * (len(vgs_arr) - 1))
    rows.append(["Vds"] + [v for v in vgs_arr])
    for vds in vds_arr:
        row = [vds]
        for vgs in vgs_arr:
            row.append(f"{iv.sel(vgs=vgs, vds=vds).item():.6e}")
        rows.append(row)

    return rows


def transfer_to_list(transfer: xr.DataArray, W: float, L: float) -> list[list]:
    """
    Convert Transfer sweep DataArray to a list of rows.

    Expected transfer structure:
        dims=["sweep", "vgs"]
        coords={"sweep": ["forward", "backward"], "vgs": ...}
        attrs={"vds": float, "dt": float, ...}
    """
    vgs_arr = transfer.coords["vgs"].values
    sweep_labels = transfer.coords["sweep"].values
    vds = transfer.attrs["vds"]

    rows = []

    # --- header ---
    rows.append(["AnalysisType", "Transfer"])
    rows.append(["ModelParams"])
    rows.append(["GateWidth", W*1e6, "um"])
    rows.append(["GateLength", L*1e6, "um"])
    rows.append(["Variables"])
    rows.append(["Ids", "A", "target"])
    rows.append(["Vds", "V", vds, "const"])
    rows.append(["Vgs", "V", "sweep"])

    # --- data ---
    rows.append(["", "Vgs"] + [""] * (len(sweep_labels) - 1))
    rows.append(["Vgs"] + list(sweep_labels))
    for vgs in vgs_arr:
        row = [vgs]
        for sweep in sweep_labels:
            row.append(f"{transfer.sel(sweep=sweep, vgs=vgs).item():.6e}")
        rows.append(row)

    return rows