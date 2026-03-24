from typing import Callable
import numpy as np
import copy


class TransistorModel:
    """
    Simplified depletion-mode GaAs/AlGaAs HEMT model.

    The model uses a compact velocity-saturation style current expression
    driven by gate overdrive and drain-source voltage. It is intended for
    demonstration and testing rather than quantitative device simulation.

    Notes
    -----
    The drain current is evaluated from the gate overdrive
    ``Vov = VGS - Vth``. If ``Vov <= 0``, the device is treated as off and
    the drain current is set to zero.

    The sheet-charge coefficient is defined as

    ``k_2deg = eps * W / (d * L)``

    and the current expression is evaluated as

    ``Id = mu * k_2deg * Vov * VDS / (1 + VDS / (v_sat * L / mu))``

    with a simple drain-voltage clipping at the estimated saturation voltage

    ``Vdsat = (Vov * v_sat * L) / (mu * Vov + v_sat * L)``.

    This formulation loosely mimics a 2DEG-based HEMT with velocity
    saturation, but omits many effects such as series resistance,
    self-heating, channel-length modulation, and trap dynamics.
    """

    def __init__(self):
        """
        Initialize default model parameters.

        Notes
        -----
        The default values correspond to a rough GaAs/AlGaAs HEMT-like setup.
        They are embedded directly for convenience in tests and demos.

        Parameters stored on the instance are:

        - ``mu``: carrier mobility in ``m^2 / (V s)``
        - ``v_sat``: saturation velocity in ``m / s``
        - ``eps``: dielectric permittivity in ``F / m``
        - ``d``: barrier thickness in ``m``
        - ``W``: gate width in ``m``
        - ``L``: gate length in ``m``
        - ``Vth``: threshold voltage in ``V``
        """
        self.mu = 0.6        # 移動度 [m^2/V·s]  GaAsは高い
        self.v_sat = 1.2e5   # 飽和速度 [m/s]
        self.eps = 12.9 * 8.854e-12  # GaAs誘電率
        self.d = 30e-9       # AlGaAsバリア厚 [m]
        self.W = 100e-6      # ゲート幅 [m]
        self.L = 1e-6        # ゲート長 [m]
        self.Vth = -0.5      # 閾値電圧（depletion-mode）

    def get_k_2deg(self) -> float:
        """
        Return the effective sheet-charge coefficient.

        Returns
        -------
        float
            Effective coefficient defined as ``eps * W / (d * L)``.

        Notes
        -----
        This helper isolates the geometric and dielectric term reused by the
        current equation.
        """
        return self.eps * self.W / (self.d * self.L)
    
    def Id(self, VGS: float, VDS: float) -> float:
        """
        Compute drain current for the given bias point.

        Parameters
        ----------
        VGS : float
            Gate-source voltage in volts.
        VDS : float
            Drain-source voltage in volts.

        Returns
        -------
        float
            Drain current in amperes.

        Notes
        -----
        The current is based on a compact velocity-saturation expression.
        First, the gate overdrive is evaluated as ``Vov = VGS - Vth``.
        When ``Vov <= 0``, the method returns ``0.0``.

        For ``Vov > 0``, the unsaturated current is computed as

        ``Id = mu * k_2deg * Vov * VDS / (1 + VDS / (v_sat * L / mu))``

        and ``VDS`` is clipped to the estimated saturation voltage

        ``Vdsat = (Vov * v_sat * L) / (mu * Vov + v_sat * L)``

        when the applied drain voltage exceeds that value.
        """
        Vov = VGS - self.Vth
        if Vov <= 0:
            return 0.0
        # 速度飽和モデル
        num = self.mu * self.get_k_2deg() * Vov * VDS
        den = 1 + VDS / (self.v_sat * self.L / self.mu)  # = 1 + mu*VDS/(v_sat*L)
        # 飽和電圧
        Vdsat = (Vov * self.v_sat * self.L) / (self.mu * Vov + self.v_sat * self.L)
        if VDS >= Vdsat:
            VDS_c = Vdsat
            num = self.mu * self.get_k_2deg() * Vov * VDS_c
            den = 1 + VDS_c / (self.v_sat * self.L / self.mu)
        return num / den
    
    @property
    def Id_v(self) -> Callable[[float | np.ndarray, float | np.ndarray], np.ndarray]:
        """
        Return a vectorized wrapper of :meth:`Id`.

        Returns
        -------
        Callable[[float | numpy.ndarray, float | numpy.ndarray], numpy.ndarray]
            ``numpy.vectorize`` wrapper around :meth:`Id`.

        Notes
        -----
        This is a convenience wrapper for array-oriented evaluation. It does
        not change the underlying scalar model or add extra physical effects.
        """
        return np.vectorize(self.Id)
    

class TransistorHysteresisModel:
    """
    Hysteresis wrapper around :class:`TransistorModel`.

    This model augments a base transistor model with a single internal state
    variable ``z`` representing trap occupancy. The state shifts the effective
    threshold voltage and evolves over time according to a simple first-order
    relaxation model.

    Notes
    -----
    The internal state is constrained to ``0 <= z <= 1``.

    The effective threshold voltage is defined as

    ``Vth_eff = base_model.Vth + alpha_vth * z``

    so a larger trap occupancy increases the threshold voltage and suppresses
    current.

    Time evolution is controlled by different time constants for filling and
    emptying, allowing a simple hysteresis behavior during bias sweeps.
    """

    def __init__(
        self,
        base_model: TransistorModel,
        *,
        alpha_vth: float = 0.25,
        tau_fill: float = 1e-3,
        tau_empty: float = 5e-3,
        z0: float = 0.0,
        copy_base_model: bool = True,
    ):
        """
        Initialize the hysteresis model.

        Parameters
        ----------
        base_model : TransistorModel
            Base transistor model used for the instantaneous current equation.
        alpha_vth : float, default=0.25
            Threshold-voltage shift coefficient applied to the trap state.
        tau_fill : float, default=1e-3
            Time constant used when the trap occupancy is increasing.
        tau_empty : float, default=5e-3
            Time constant used when the trap occupancy is decreasing.
        z0 : float, default=0.0
            Initial trap occupancy. The value is clipped into ``[0, 1]``.
        copy_base_model : bool, default=True
            Whether to deep-copy ``base_model`` before storing it.

        Notes
        -----
        Deep-copying the base model avoids accidental parameter sharing across
        multiple hysteresis model instances.
        """
        self.base_model = copy.deepcopy(base_model) if copy_base_model else base_model

        self.alpha_vth = alpha_vth
        self.tau_fill = tau_fill
        self.tau_empty = tau_empty
        self.z = float(np.clip(z0, 0.0, 1.0))

    def initialize_state(self, z0: float = 0.0) -> None:
        """
        Reset the internal trap state.

        Parameters
        ----------
        z0 : float, default=0.0
            New initial trap occupancy. The value is clipped into ``[0, 1]``.
        """
        self.z = float(np.clip(z0, 0.0, 1.0))

    @property
    def Vth_eff(self) -> float:
        """
        Return the effective threshold voltage.

        Returns
        -------
        float
            Effective threshold voltage defined as
            ``base_model.Vth + alpha_vth * z``.
        """
        return self.base_model.Vth + self.alpha_vth * self.z

    def z_inf(self, VGS: float, VDS: float) -> float:
        """
        Compute the equilibrium trap occupancy for a bias point.

        Parameters
        ----------
        VGS : float
            Gate-source voltage in volts.
        VDS : float
            Drain-source voltage in volts.

        Returns
        -------
        float
            Equilibrium trap occupancy in the range ``[0, 1]``.

        Notes
        -----
        This is a deliberately simple empirical model. Higher ``VGS`` and
        higher non-negative ``VDS`` increase the target trap occupancy through
        a sigmoid mapping.
        """
        x = 4.0 * (VGS - 0.0) + 1.5 * max(VDS, 0.0)
        return float(1.0 / (1.0 + np.exp(-x)))

    def update_state(self, VGS: float, VDS: float, dt: float) -> None:
        """
        Advance the internal trap state toward equilibrium.

        Parameters
        ----------
        VGS : float
            Gate-source voltage in volts.
        VDS : float
            Drain-source voltage in volts.
        dt : float
            Time step used for the state update.

        Notes
        -----
        The state moves toward :meth:`z_inf` using a first-order exponential
        relaxation. ``tau_fill`` is used when the target occupancy is above the
        current state, and ``tau_empty`` is used otherwise.
        """
        z_target = self.z_inf(VGS, VDS)
        tau = self.tau_fill if z_target > self.z else self.tau_empty

        self.z += (z_target - self.z) * (1.0 - np.exp(-dt / tau))
        self.z = float(np.clip(self.z, 0.0, 1.0))

    def Id(self, VGS: float, VDS: float) -> float:
        """
        Compute the instantaneous drain current at the current internal state.

        Parameters
        ----------
        VGS : float
            Gate-source voltage in volts.
        VDS : float
            Drain-source voltage in volts.

        Returns
        -------
        float
            Drain current in amperes.

        Notes
        -----
        This method does not update the hysteresis state. It only evaluates the
        current using the present ``z`` value through ``Vth_eff``. The current
        equation itself matches the compact expression used by the base model,
        with ``Vth`` replaced by ``Vth_eff``.
        """
        Vov = VGS - self.Vth_eff
        if Vov <= 0:
            return 0.0

        m = self.base_model
        num = m.mu * m.get_k_2deg() * Vov * VDS
        den = 1 + VDS / (m.v_sat * m.L / m.mu)

        Vdsat = (Vov * m.v_sat * m.L) / (m.mu * Vov + m.v_sat * m.L)
        if VDS >= Vdsat:
            VDS_c = Vdsat
            num = m.mu * m.get_k_2deg() * Vov * VDS_c
            den = 1 + VDS_c / (m.v_sat * m.L / m.mu)

        return num / den

    def Id_step(self, VGS: float, VDS: float, dt: float) -> float:
        """
        Update the internal state by one time step and then compute current.

        Parameters
        ----------
        VGS : float
            Gate-source voltage in volts.
        VDS : float
            Drain-source voltage in volts.
        dt : float
            Time step used for the state update.

        Returns
        -------
        float
            Drain current after the state update.

        Notes
        -----
        This method is typically used for bias sweeps where hysteresis should
        accumulate over time.
        """
        self.update_state(VGS, VDS, dt)
        return self.Id(VGS, VDS)

    def sweep_vgs(
        self,
        vgs_values: np.ndarray,
        VDS: float,
        dt: float,
        *,
        initialize_state: bool = False,
        z0: float = 0.0,
    ) -> np.ndarray:
        """
        Sweep gate voltage and return the corresponding drain currents.

        Parameters
        ----------
        vgs_values : numpy.ndarray
            Sequence of gate-source voltages.
        VDS : float
            Fixed drain-source voltage in volts.
        dt : float
            Time step used between consecutive bias points.
        initialize_state : bool, default=False
            Whether to reset the internal state before starting the sweep.
        z0 : float, default=0.0
            Initial trap occupancy used when ``initialize_state`` is true.

        Returns
        -------
        numpy.ndarray
            Drain-current array evaluated point by point with hysteresis.

        Notes
        -----
        The state is updated for each element of ``vgs_values`` in order, so
        the result depends on the sweep direction and history.
        """
        if initialize_state:
            self.initialize_state(z0)

        ids = []
        for vgs in vgs_values:
            ids.append(self.Id_step(float(vgs), float(VDS), dt))
        return np.asarray(ids)

    def sweep_vds(
        self,
        vds_values: np.ndarray,
        VGS: float,
        dt: float,
        *,
        initialize_state: bool = False,
        z0: float = 0.0,
    ) -> np.ndarray:
        """
        Sweep drain voltage and return the corresponding drain currents.

        Parameters
        ----------
        vds_values : numpy.ndarray
            Sequence of drain-source voltages.
        VGS : float
            Fixed gate-source voltage in volts.
        dt : float
            Time step used between consecutive bias points.
        initialize_state : bool, default=False
            Whether to reset the internal state before starting the sweep.
        z0 : float, default=0.0
            Initial trap occupancy used when ``initialize_state`` is true.

        Returns
        -------
        numpy.ndarray
            Drain-current array evaluated point by point with hysteresis.

        Notes
        -----
        The state is updated for each element of ``vds_values`` in order, so
        the result depends on the sweep trajectory rather than only the final
        bias points.
        """
        if initialize_state:
            self.initialize_state(z0)

        ids = []
        for vds in vds_values:
            ids.append(self.Id_step(float(VGS), float(vds), dt))
        return np.asarray(ids)

    @property
    def Id_v(self) -> Callable[[float | np.ndarray, float | np.ndarray], np.ndarray]:
        """
        Return a vectorized wrapper of :meth:`Id` without state updates.

        Returns
        -------
        Callable[[float | numpy.ndarray, float | numpy.ndarray], numpy.ndarray]
            ``numpy.vectorize`` wrapper around :meth:`Id`.

        Notes
        -----
        This property does not perform hysteresis evolution. For history-aware
        sweeps, use :meth:`sweep_vgs` or :meth:`sweep_vds` instead.
        """
        return np.vectorize(self.Id)
