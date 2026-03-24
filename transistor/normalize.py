
import xarray as xr
import funcexpr_xr as fxr


def to_current_density(data: xr.DataArray, W_um: float) -> xr.DataArray:
    """Convert drain current [A] to current density [mA/mm].

    Uses funcexpr-xr to evaluate the normalization expression,
    keeping the xarray coordinate structure intact.

    The conversion is:
        Jd [mA/mm] = Id [A] * 1e3 / (W [um] * 1e-3)
                   = Id [A] * 1e6 / W [um]

    Args:
        data: Drain current DataArray in amperes [A].
              Any dims/coords are preserved.
        W_um: Gate width in micrometres [um].

    Returns:
        DataArray with the same dims/coords as ``data``,
        values in mA/mm.
    """
    def bypass(arrays, digits):
        return arrays
    fxr.alignment.register(
        "bypass",
        bypass
    )
    return fxr.evaluate(
        "iv_A * 1e3 / (W_um * 1e-3)",
        ctx={"iv_A": data, "W_um": W_um},
        alignment="bypass"
    )