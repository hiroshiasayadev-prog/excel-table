
import numpy as np
import xarray as xr

import plotly.graph_objects as go
import plotly.io as pio
import plotly.colors as pc

class PlotlyStyle:

    @classmethod
    def _sample_colorscale(cls, scale, val):
        """val: 0.0〜1.0"""
        return pc.sample_colorscale(scale, val)[0]
    
    @classmethod
    def rdylgn_colorscale(cls, axis, value):
        scale = pc.get_colorscale("RdYlGn")
        idx = np.where(axis == value)[0]
        n = len(axis)
        val = idx / (n - 1)
        return pc.sample_colorscale(scale, val)[0]
    
    @classmethod
    def apply_paper_style(cls, fig: go.Figure) -> go.Figure:
        fig.update_layout(
            template="plotly_white",
            xaxis=dict(
                showline=True, linewidth=1, linecolor="black", mirror=True,
                ticks="inside", tickwidth=1, tickcolor="black",
                showgrid=True, gridwidth=1, gridcolor="lightgray",
                zeroline=True, zerolinewidth=1, zerolinecolor="lightgray",
            ),
            yaxis=dict(
                showline=True, linewidth=1, linecolor="black", mirror=True,
                ticks="inside", tickwidth=1, tickcolor="black",
                showgrid=True, gridwidth=1, gridcolor="lightgray",
                zeroline=True, zerolinewidth=1, zerolinecolor="lightgray",
            ),
            legend=dict(
                bordercolor="black", borderwidth=1,
            ),
        )
        return fig

class Display:

    @classmethod
    def generate_IV_map(cls, iv_data: xr.DataArray) -> str:
        """
        generate I-V characteristic map as json for plotly.
        the structure of `iv_data` must be as follows:

        iv_data = xr.DataArray(
            data=id_data,
            dims=["vgs", "vds"],
            coords={"vgs": vgs, "vds": vds},
            attrs={"units": "A", "description": "Drain current"},
        )
        """
        fig = go.Figure()

        vgs_arr = iv_data.coords["vgs"].values
        for vgs in vgs_arr:
            slice_ = iv_data.sel(vgs=vgs)
            fig.add_trace(go.Scatter(
                x=slice_.coords["vds"].values,
                y=slice_.values,
                mode="lines",
                name=f"Vgs={vgs:.1f}V",
                line=dict(color=PlotlyStyle.rdylgn_colorscale(
                    axis=vgs_arr, value=vgs
                )),
            ))

        fig.update_layout(
            #title="I-V characteristics",
            xaxis_title="Vds [V]",
            yaxis_title="Id [A]",
            legend_title="Vgs",
        )

        PlotlyStyle.apply_paper_style(fig)

        fig_json = fig.to_json()
        if not isinstance(fig_json, str):
            raise Exception("Fuck")

        return fig_json
    
    @classmethod
    def generate_Transfer_map(cls, transfer_data: xr.DataArray) -> str:
        """
        generate Transfer characteristic map as json for plotly.
        the structure of `transfer_data` must be as follows:

        xr.DataArray(
            data=np.stack([id_forward, id_backward]),
            dims=["sweep", "vgs"],
            coords={
                "sweep": ["forward", "backward"],
                "vgs": vgs_forward,
            },
        )
        """
        fig = go.Figure()

        style = {
            "forward":  dict(dash="solid"),
            "backward": dict(dash="dash"),
        }

        for sweep in transfer_data.coords["sweep"].values:
            slice_ = transfer_data.sel(sweep=sweep)
            vgs = slice_.coords["vgs"].values
            id_ = slice_.values
            gm = np.gradient(id_, vgs)

            fig.add_trace(go.Scatter(
                x=vgs, y=id_,
                mode="lines",
                name=f"Id ({sweep})",
                line=dict(color="blue", **style[sweep]),
            ))
            fig.add_trace(go.Scatter(
                x=vgs, y=gm,
                mode="lines",
                name=f"gm ({sweep})",
                yaxis="y2",
                line=dict(color="red", **style[sweep]),
            ))

        fig.update_layout(
            xaxis=dict(title="Vgs [V]"),
            yaxis=dict(title="Id [A]"),
            yaxis2=dict(
                title="gm [S]",
                overlaying="y",
                side="right",
            ),
        )

        PlotlyStyle.apply_paper_style(fig)

        fig_json = fig.to_json()
        if not isinstance(fig_json, str):
            raise Exception("Fuck")

        return fig_json

    @classmethod
    def show_IV(cls, iv_data: xr.DataArray):
        """
        show I-V characteristic.
        iv_data's structure must like below.

        iv_data = xr.DataArray(
            data=id_data,
            dims=["vgs", "vds"],
            coords={"vgs": vgs, "vds": vds},
            attrs={"units": "A", "description": "Drain current"},
        )
        """
        fig = pio.from_json(
            cls.generate_IV_map(iv_data)
        )
        fig.show()

    @classmethod
    def show_Transfer(cls, transfer_data: xr.DataArray):
        """
        show Transfer characteristic map as json for plotly.
        the structure of `transfer_data` must be as follows:

        xr.DataArray(
            data=np.stack([id_forward, id_backward]),
            dims=["sweep", "vgs"],
            coords={
                "sweep": ["forward", "backward"],
                "vgs": vgs_forward,
            },
        )
        """
        fig = pio.from_json(
            cls.generate_Transfer_map(transfer_data)
        )
        fig.show()

    @classmethod
    def generate_IV_density_map(cls, jd_data: xr.DataArray) -> str:
        """generate I-V current density map as json for plotly.

        jd_data must have the same structure as generate_IV_map,
        but values in mA/mm (current density normalized by gate width).
        """
        fig = go.Figure()

        vgs_arr = jd_data.coords["vgs"].values
        for vgs in vgs_arr:
            slice_ = jd_data.sel(vgs=vgs)
            fig.add_trace(go.Scatter(
                x=slice_.coords["vds"].values,
                y=slice_.values,
                mode="lines",
                name=f"Vgs={vgs:.1f}V",
                line=dict(color=PlotlyStyle.rdylgn_colorscale(
                    axis=vgs_arr, value=vgs
                )),
            ))

        fig.update_layout(
            title="I-V characteristics (current density)",
            xaxis_title="Vds [V]",
            yaxis_title="Jd [mA/mm]",
            legend_title="Vgs",
        )

        PlotlyStyle.apply_paper_style(fig)

        fig_json = fig.to_json()
        if not isinstance(fig_json, str):
            raise Exception("Fuck")
        return fig_json

    @classmethod
    def generate_Transfer_density_map(cls, jd_data: xr.DataArray) -> str:
        """generate Transfer current density map as json for plotly.

        jd_data must have the same structure as generate_Transfer_map,
        but values in mA/mm. gm is derived as d(Jd)/d(Vgs) [mS/mm].
        """
        fig = go.Figure()

        style = {
            "forward":  dict(dash="solid"),
            "backward": dict(dash="dash"),
        }

        for sweep in jd_data.coords["sweep"].values:
            slice_ = jd_data.sel(sweep=sweep)
            vgs = slice_.coords["vgs"].values
            jd_ = slice_.values
            gm = np.gradient(jd_, vgs)

            fig.add_trace(go.Scatter(
                x=vgs, y=jd_,
                mode="lines",
                name=f"Jd ({sweep})",
                line=dict(color="blue", **style[sweep]),
            ))
            fig.add_trace(go.Scatter(
                x=vgs, y=gm,
                mode="lines",
                name=f"gm ({sweep})",
                yaxis="y2",
                line=dict(color="red", **style[sweep]),
            ))

        fig.update_layout(
            xaxis=dict(title="Vgs [V]"),
            yaxis=dict(title="Jd [mA/mm]"),
            yaxis2=dict(
                title="gm [mS/mm]",
                overlaying="y",
                side="right",
            ),
        )

        PlotlyStyle.apply_paper_style(fig)

        fig_json = fig.to_json()
        if not isinstance(fig_json, str):
            raise Exception("Fuck")
        return fig_json