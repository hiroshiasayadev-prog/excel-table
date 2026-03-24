import numpy as np
import xarray as xr

from .model import TransistorModel, TransistorHysteresisModel

class Analyzer:

    @classmethod
    def ID(
        cls,
        transistor: TransistorModel,
        VGS: float | np.ndarray,
        VDS: float | np.ndarray
    ) -> np.ndarray:
        _VGS = VGS[:, None] if isinstance(VGS, np.ndarray) else VGS
        _VDS = VDS[None, :] if isinstance(VDS, np.ndarray) else VDS
        return transistor.Id_v(_VGS, _VDS)
    
    @classmethod
    def sweep_IV(
        cls,
        transistor: TransistorModel,
        vds_from: float = 0.0,
        vds_until: float = 1.0,
        vds_step: float = 0.01,
        vds: float | None = None,
        vgs_from: float = -0.4,
        vgs_until: float = 1.0,
        vgs_step: float = 0.2,
        vgs: float | None = None,
    ) -> xr.DataArray:
        """
        measure Id by sweeping VDS and VGS.
        If a specific value is provided for vds or vgs,
        that variable is fixed instead of swept.

        returns result that the structure is as follows:

        xr.DataArray(
            data=id_data,
            dims=["vgs", "vds"],
            coords={"vgs": vgs, "vds": vds},
            attrs={"units": "A", "description": "Drain current"},
        )
        """

        # define VDS points
        num_vds = round((vds_until - vds_from) / vds_step) + 1
        VDS_arr = np.linspace(vds_from, vds_until, num_vds)

        # define VGS points
        num_vgs = round((vgs_until - vgs_from) / vgs_step) + 1
        VGS_arr = np.linspace(vgs_from, vgs_until, num_vgs)

        # measure
        id_arr = Analyzer.ID(
            transistor,
            VGS=VGS_arr,
            VDS=VDS_arr
        )

        # format
        dims = []
        coords = {}
        for dim, const, arr in (("vgs", vgs, VGS_arr), ("vds", vds, VDS_arr)):
            if const is None:
                dims.append(dim)
                coords[dim] = arr

        return xr.DataArray(
            data=id_arr,
            dims=dims,
            coords=coords,
            attrs={"units": "A", "description": "Drain current"},
        )
    
    @classmethod
    def sweep_Vgs(
        cls,
        transistor: TransistorModel | TransistorHysteresisModel,
        vgs_from: float = -1.0,
        vgs_until: float = 1.0,
        vgs_step: float = 0.01,
        vds: float = 1.5,
        dt: float = 1e-4,
    ) -> xr.DataArray:
        """
        measure Id by sweeping VGS.

        returns result that the structure is as follows:

        xr.DataArray(
            data=np.stack([id_forward, id_backward]),
            dims=["sweep", "vgs"],
            coords={
                "sweep": ["forward", "backward"],
                "vgs": vgs_forward,
            },
        )
        """

        # prepare model if needed
        if isinstance(transistor, TransistorModel):
            model = TransistorHysteresisModel(transistor)
        elif isinstance(transistor, TransistorHysteresisModel):
            model = transistor
        else:
            raise NotImplementedError()

        # define vgs points
        num_vgs = round((vgs_until - vgs_from)/vgs_step) + 1
        vgs_for = np.linspace(vgs_from, vgs_until, num_vgs)
        vgs_bak = np.linspace(vgs_until, vgs_from, num_vgs)

        # measure
        id_for = model.sweep_vgs(vgs_for, VDS=vds, dt=dt, initialize_state=True)
        id_bak = model.sweep_vgs(vgs_bak, VDS=vds, dt=dt, initialize_state=False)
        id_bak = id_bak[::-1]

        # format
        return xr.DataArray(
            data=np.stack([id_for, id_bak]),
            dims=["sweep", "vgs"],
            coords={
                "sweep": ["forward", "backward"],
                "vgs": vgs_for,
            },
            attrs={"vds": vds, "dt": dt},
        )