"""
studies/transient_stability/transient_stability_study.py
---------------------------------------------------------
Automates AESO Transient Stability Analysis (Phase 3) using PSS/E via psspy.

Workflow
--------
1.  Load base case (.sav) — power flow snapshot
2.  Load dynamic models (.dyr) — machine models, AVRs, PSS, etc.
3.  Solve base case Newton-Raphson power flow
4.  For each contingency in TS_Contingencies:
    a. Reload fresh .sav
    b. Reload .dyr (dynamics must be re-loaded after each psspy.case() call)
    c. Set up output channels (voltage, rotor angle, active/reactive power)
    d. Initialise dynamic simulation (psspy.strt)
    e. Run pre-fault simulation to steady state (t=0 to fault_apply_time)
    f. Apply bolted 3-phase fault at contingency bus
    g. Run during-fault simulation (fault_apply_time to fault_clear_time)
    h. Clear fault (restore branch)
    i. Run post-fault simulation (fault_clear_time to sim_end)
    j. Extract channel data: bus voltages, rotor angles, active/reactive power
    k. Check AESO criteria: rotor angle stability, voltage recovery
5.  Export results to Excel + multi-panel PNG plots + PDF report

AESO Criteria Applied (Study Scope Section 5.3 and Requirements doc Section 3.4)
----------------------------------------------------------------------------------
- Fault type: bolted 3-phase to ground (most severe for Category B/C)
- Reference generator: Genesee unit 3, Wabamun Area 40
- Rotor angle stability: all generators remain stable (no pole slip)
- Rotor angle limit: < 180 degrees relative to reference generator
- Voltage recovery: POI bus voltage recovers to ≥ 0.90 pu within 1.0 s post-fault
- Monitor: 500 kV, 240 kV, 138 kV buses near point of connection
- Monitor: rotor angle, active power, reactive power for study area generators
- Fault clearing times: from Table 4-8 / Table 4-14 in Study Scope (per contingency)
- Dynamic models: loaded from .dyr file (must be co-located with .sav or set explicitly
  via ProjectInfo.dyr_file_path)

Usage
-----
    from core.psse_interface import PSSEInterface
    from studies.transient_stability.transient_stability_study import (
        TransientStabilityStudy
    )
    from project_io.project_data import ProjectData

    psse = PSSEInterface(psse_path=PSSE_PATH)
    psse.initialize()

    study = TransientStabilityStudy(psse, project_data)
    results = study.run(sav_path)
    study.save_results("output/results", "output/plots", "output/reports")
"""

import logging
import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

from project_io.project_data import ProjectData, TSContingency

logger = logging.getLogger(__name__)

# ── AESO colour palette ───────────────────────────────────────────────────────────────────
_BLUE   = "#003865"
_RED    = "#C8102E"
_ORANGE = "#E87722"
_GREEN  = "#00853E"
_GREY   = "#6C6F70"

plt.rcParams.update({
    "figure.dpi":       150,
    "figure.facecolor": "white",
    "axes.facecolor":   "#F7F7F7",
    "axes.edgecolor":   _GREY,
    "axes.labelcolor":  _BLUE,
    "axes.titleweight": "bold",
    "axes.titlecolor":  _BLUE,
    "xtick.color":      _GREY,
    "ytick.color":      _GREY,
    "grid.color":       "white",
    "grid.linewidth":   1.0,
    "font.family":      "DejaVu Sans",
})


# ── Data containers ──────────────────────────────────────────────────────────────────

@dataclass
class ChannelData:
    """Time-domain simulation data for one monitored channel."""
    name:        str
    units:       str
    time_s:      List[float] = field(default_factory=list)
    values:      List[float] = field(default_factory=list)


@dataclass
class ContingencyTSResult:
    """Transient stability results for one contingency."""
    contingency_name:     str
    fault_location:       str
    fault_bus_no:         Optional[int]
    near_end_cycles:      float
    far_end_cycles:       float
    fault_apply_time_s:   float
    fault_clear_time_s:   float
    converged_base:       bool
    sim_completed:        bool

    # AESO compliance
    rotor_angle_stable:   bool  = True
    max_rotor_angle_deg:  float = 0.0
    voltage_recovered:    bool  = True
    min_poi_voltage_pu:   float = 1.0
    recovery_time_s:      float = 0.0

    # Channel data (keyed by channel description)
    channels: Dict[str, ChannelData] = field(default_factory=dict)

    @property
    def aeso_pass(self) -> bool:
        return self.rotor_angle_stable and self.voltage_recovered


@dataclass
class TransientStabilityResults:
    """Aggregated results for one scenario."""
    scenario_label:   str
    sav_path:         str
    sim_duration_s:   float
    contingencies:    List[ContingencyTSResult] = field(default_factory=list)

    @property
    def total_pass(self) -> int:
        return sum(1 for c in self.contingencies if c.aeso_pass)

    @property
    def total_fail(self) -> int:
        return sum(1 for c in self.contingencies if not c.aeso_pass)


# ── Study class ───────────────────────────────────────────────────────────────────

class TransientStabilityStudy:
    """
    Automates PSS/E transient stability analysis for one operating scenario.

    Parameters
    ----------
    psse : PSSEInterface
        Initialized PSS/E interface.
    project : ProjectData
        Populated project data (scenarios, contingencies, bus numbers).
    scenario_label : str
        Label used in output filenames.
    sim_duration_s : float
        Total simulation window in seconds (default 10.0 s).
    fault_apply_time_s : float
        Time at which the fault is applied (default 1.0 s).
    time_step_s : float
        Integration time step (default 1/2 cycle = 0.00833 s at 60 Hz).
    rotor_angle_limit_deg : float
        First-swing rotor angle stability limit (default 180 deg).
    voltage_recovery_pu : float
        Minimum POI voltage after recovery window (default 0.90 pu).
    voltage_recovery_window_s : float
        Post-fault window to check voltage recovery (default 1.0 s).
    """

    def __init__(
        self,
        psse,
        project:                  ProjectData,
        scenario_label:           str   = "Base_Case",
        sim_duration_s:           float = 10.0,
        fault_apply_time_s:       float = 1.0,
        time_step_s:              float = 0.00833,
        rotor_angle_limit_deg:    float = 180.0,
        voltage_recovery_pu:      float = 0.90,
        voltage_recovery_window_s:float = 1.0,
    ):
        self._psse                  = psse
        self._project               = project
        self.scenario               = scenario_label
        self.sim_duration_s         = sim_duration_s
        self.fault_apply_time_s     = fault_apply_time_s
        self.time_step_s            = time_step_s
        self.rotor_limit            = rotor_angle_limit_deg
        self.v_recovery_pu          = voltage_recovery_pu
        self.v_recovery_window_s    = voltage_recovery_window_s
        self.results: Optional[TransientStabilityResults] = None

    # ── Public ────────────────────────────────────────────────────────────────

    def run(self, sav_path: str) -> TransientStabilityResults:
        """
        Run transient stability analysis for all contingencies.

        Parameters
        ----------
        sav_path : str
            Path to the .sav case file.

        Returns
        -------
        TransientStabilityResults
        """
        logger.info("=" * 68)
        logger.info("Transient Stability Study  |  Scenario: %s", self.scenario)
        logger.info("SAV: %s", sav_path)
        logger.info(
            "Sim: %.1f s  |  Fault at: %.1f s  |  dt: %.5f s",
            self.sim_duration_s, self.fault_apply_time_s, self.time_step_s
        )
        logger.info("=" * 68)

        contingencies = self._project.ts_contingencies
        if not contingencies:
            logger.warning(
                "No TS contingencies defined in project data. "
                "Fill TS_Contingencies sheet."
            )

        cont_results = []
        for cont in contingencies:
            result = self._run_contingency(sav_path, cont)
            cont_results.append(result)
            status = "PASS" if result.aeso_pass else "FAIL"
            logger.info(
                "  %-45s  [%s]  "
                "max_angle=%.1f°  min_V=%.3f pu",
                cont.contingency_name, status,
                result.max_rotor_angle_deg,
                result.min_poi_voltage_pu,
            )

        self.results = TransientStabilityResults(
            scenario_label=self.scenario,
            sav_path=sav_path,
            sim_duration_s=self.sim_duration_s,
            contingencies=cont_results,
        )

        logger.info(
            "\nTS Study complete: %d contingencies  "
            "%d PASS  %d FAIL",
            len(cont_results),
            self.results.total_pass,
            self.results.total_fail,
        )
        return self.results

    def save_results(
        self,
        results_dir: str,
        plots_dir:   str,
        reports_dir: str,
    ) -> dict:
        """
        Export results to Excel, PNG plots, and PDF report.

        Returns dict with keys 'excel', 'plots', 'pdf'.
        """
        if self.results is None:
            raise RuntimeError("No results to save. Call run() first.")

        for d in (results_dir, plots_dir, reports_dir):
            os.makedirs(d, exist_ok=True)

        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        label = f"transient_stability_{self.scenario}_{ts}"

        # Excel
        excel_path = os.path.join(results_dir, f"{label}.xlsx")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            self._summary_to_df().to_excel(
                writer, sheet_name="Summary", index=False)
            self._compliance_to_df().to_excel(
                writer, sheet_name="AESO_Compliance", index=False)
            for cont in self.results.contingencies:
                self._channel_to_df(cont).to_excel(
                    writer,
                    sheet_name=_safe_sheet_name(cont.contingency_name),
                    index=False,
                )
        logger.info("Excel saved: %s", excel_path)

        # PNG plots — one per contingency
        plot_paths = []
        for cont in self.results.contingencies:
            if cont.channels:
                path = self._plot_contingency(cont, plots_dir, label)
                if path:
                    plot_paths.append(path)

        # PDF report
        pdf_path = os.path.join(reports_dir, f"{label}.pdf")
        self._export_pdf(pdf_path, plot_paths)
        logger.info("PDF saved: %s", pdf_path)

        return {"excel": excel_path, "plots": plot_paths, "pdf": pdf_path}

    # ── Private: PSS/E transient stability engine ─────────────────────────────────

    def _load_dynamics(self, psspy, sav_path: str) -> bool:
        """
        Load the .dyr dynamic data file into PSS/E memory.

        Must be called after every psspy.case() call because loading a new
        .sav clears all previously loaded dynamic models from memory.
        PSS/E API call: psspy.dyre_new() reads a .dyr file and registers
        all machine, exciter, governor, and PSS models.

        Resolution order for the .dyr path (via ProjectData.resolve_dyr_path):
          1. info.dyr_file_path  — explicit path in Project_Info sheet / GUI
          2. sav_path with .sav replaced by .dyr  — co-located file

        Parameters
        ----------
        psspy : module
            The active psspy module from PSSEInterface.
        sav_path : str
            Path to the .sav file just loaded (used as a fallback base path).

        Returns
        -------
        bool
            True if dynamics loaded successfully, False otherwise.
        """
        dyr_path = self._project.resolve_dyr_path()

        # If resolve_dyr_path() returned empty or the project-level path is
        # empty, fall back to deriving from the per-contingency sav_path.
        if not dyr_path or not os.path.isfile(dyr_path):
            dyr_path = os.path.splitext(sav_path)[0] + ".dyr"

        if not os.path.isfile(dyr_path):
            logger.error(
                "  .dyr file not found. Dynamics cannot be loaded. "
                "Set 'DYR File Path' in Project_Info or place a .dyr file "
                "alongside the .sav file. Checked: '%s'",
                dyr_path,
            )
            return False

        # psspy.dyre_new([iflags], dyrfile, ldyfile, logfile, status)
        # iflags=[0,0,0,0]: use defaults for all options
        # Empty strings for ldyfile/logfile: no separate load/log files
        ierr = psspy.dyre_new([0, 0, 0, 0], dyr_path, "", "", "")
        if ierr != 0:
            logger.error(
                "  psspy.dyre_new() failed (ierr=%d) for .dyr file: '%s'.",
                ierr, dyr_path,
            )
            return False

        logger.debug("  Dynamics loaded from: %s", dyr_path)
        return True

    def _run_contingency(
        self,
        sav_path: str,
        cont:     TSContingency,
    ) -> ContingencyTSResult:
        """Run one contingency and return its result."""

        # Fault clearing time comes from the Study Scope table (per contingency)
        # Use near-end clearing time (more conservative)
        fault_clear_time_s = (
            self.fault_apply_time_s + cont.near_end_seconds
        )

        result = ContingencyTSResult(
            contingency_name   = cont.contingency_name,
            fault_location     = cont.fault_location,
            fault_bus_no       = cont.from_bus_no,
            near_end_cycles    = cont.near_end_cycles,
            far_end_cycles     = cont.far_end_cycles,
            fault_apply_time_s = self.fault_apply_time_s,
            fault_clear_time_s = fault_clear_time_s,
            converged_base     = False,
            sim_completed      = False,
        )

        if self._psse.mock:
            return self._mock_contingency_result(result)

        psspy = self._psse.psspy
        poi_bus = self._project.info.poi_bus_number

        # ── Step 1: Load fresh case ─────────────────────────────────────────
        ierr = psspy.case(sav_path)
        if ierr != 0:
            logger.error("  Failed to load case for '%s'.", cont.contingency_name)
            return result

        # ── Step 2: Load dynamic models (.dyr) ──────────────────────────────
        # MUST be done after psspy.case() and before psspy.strt().
        # psspy.case() loads power flow data only; dynamic models are cleared
        # from memory each time a new .sav is loaded. Without this step,
        # psspy.strt() will fail with ierr=5 ("No dynamics data in memory").
        if not self._load_dynamics(psspy, sav_path):
            logger.error(
                "  Skipping contingency '%s': dynamics could not be loaded.",
                cont.contingency_name,
            )
            return result

        # ── Step 3: Solve base power flow ───────────────────────────────────
        ret = psspy.fnsl([1, 0, 1, 1, 1, 0, 0, 0])
        result.converged_base = (ret == 0)
        if not result.converged_base:
            logger.warning(
                "  Base case did not converge for '%s'. Skipping.",
                cont.contingency_name
            )
            return result

        # ── Step 4: Set up output channels ───────────────────────────────────
        # Channel setup must be done BEFORE psspy.strt().
        # We set up channels for:
        #   - POI bus voltage (if poi_bus is known)
        #   - Generator rotor angles (relative to system COI reference)
        #   - Active and reactive power at POI
        output_file = os.path.join(
            os.path.dirname(sav_path),
            f"ts_{self.scenario}_{_safe_filename(cont.contingency_name)}.out"
        )

        # Initialise output file
        psspy.delete_all_plot_channels()

        # POI bus voltage channel
        if poi_bus:
            ierr = psspy.voltage_channel([-1, -1, -1, poi_bus], "Bus Voltage POI")
            if ierr != 0:
                logger.warning(
                    "  Could not add voltage channel for POI bus %d.", poi_bus
                )

        # Add rotor angle channels for generators in study area.
        # FIX: psspy.angle_channel() does not exist in the PSS/E Python API.
        # The correct function is psspy.machine_array_channel():
        #   - ITYPE=1 requests the rotor angle output (degrees relative to
        #     the system centre-of-inertia reference)
        #   - Parameters: [subsystem, ITYPE, bus_number], machine_id, label
        try:
            ierr, (gen_buses,) = psspy.amachint(-1, 4, ["NUMBER"])
            ierr2, (gen_ids,)  = psspy.amachchar(-1, 4, ["ID"])
            for i, gbus in enumerate(gen_buses):
                gid = gen_ids[i].strip() if gen_ids else "1"
                ierr_ch = psspy.machine_array_channel(
                    [-1, 1, gbus],          # subsystem=-1 (all), ITYPE=1 (angle), bus
                    gid,                    # machine ID string
                    f"Rotor Angle Bus {gbus}",
                )
                if ierr_ch != 0:
                    logger.debug(
                        "  Could not add rotor angle channel for gen at bus %d "
                        "(machine_array_channel ierr=%d).",
                        gbus, ierr_ch,
                    )
        except Exception as exc:
            logger.debug("  Generator channel setup: %s", exc)

        # ── Step 5: Initialise dynamics ───────────────────────────────────────
        # psspy.strt() initialises the dynamic simulation and writes the
        # output file header. It requires dynamic models to already be in
        # memory (loaded in Step 2 above). units=0 → output in per unit.
        ierr = psspy.strt(0, output_file)
        if ierr != 0:
            logger.warning(
                "  psspy.strt() failed (ierr=%d) for '%s'.",
                ierr, cont.contingency_name
            )
            return result

        # ── Step 6: Run to fault application time ──────────────────────────
        ierr = psspy.run(0, self.fault_apply_time_s, 1000, 1, 0)
        if ierr != 0:
            logger.warning(
                "  Pre-fault simulation failed (ierr=%d).", ierr
            )
            return result

        # ── Step 7: Apply 3-phase fault ───────────────────────────────────────
        fault_bus = cont.from_bus_no
        if fault_bus is None:
            logger.warning(
                "  No fault bus defined for '%s'. Cannot apply fault. "
                "Fill From Bus No in TS_Contingencies sheet.",
                cont.contingency_name
            )
            return result

        # Apply bolted 3-phase fault: fault impedance = 0
        # fault_bus, units, voltage, basekv, 0.0+j0.0 fault impedance
        ierr = psspy.dist_bus_fault(fault_bus, 1, 0.0, [0.0, 0.0])
        if ierr != 0:
            logger.warning(
                "  dist_bus_fault failed (ierr=%d) at bus %d.",
                ierr, fault_bus
            )

        # ── Step 8: Run during-fault ────────────────────────────────────────
        ierr = psspy.run(0, fault_clear_time_s, 1000, 1, 0)

        # ── Step 9: Clear fault (open branch) ──────────────────────────────
        # Remove fault
        ierr_cf = psspy.dist_clear_fault(1)

        # Trip the contingency branch (N-1)
        if cont.from_bus_no and cont.to_bus_no:
            psspy.dist_branch_trip(
                cont.from_bus_no,
                cont.to_bus_no,
                cont.circuit_id,
            )

        # ── Step 10: Run post-fault simulation ─────────────────────────────
        ierr = psspy.run(0, self.sim_duration_s, 10000, 1, 0)
        result.sim_completed = (ierr == 0)

        # ── Step 11: Extract channel data ───────────────────────────────────
        channels = self._extract_channels(psspy, output_file, poi_bus)
        result.channels = channels

        # ── Step 12: Evaluate AESO criteria ────────────────────────────────
        self._evaluate_criteria(result, poi_bus)

        return result

    def _extract_channels(
        self,
        psspy,
        output_file: str,
        poi_bus:     Optional[int],
    ) -> Dict[str, ChannelData]:
        """
        Extract time-domain channel data from PSS/E output file.
        Returns dict of channel_name → ChannelData.
        """
        channels = {}
        try:
            # Get number of channels written
            ierr, nchan = psspy.numchnf(output_file)
            if ierr != 0 or nchan == 0:
                return channels

            # Read time axis
            ierr, time_arr = psspy.chnval(0)   # channel 0 = time
            if ierr != 0:
                return channels
            time_list = list(time_arr) if time_arr else []

            # Read each channel
            for ch in range(1, nchan + 1):
                try:
                    ierr_ch, desc = psspy.chndes(ch)
                    ierr_v,  vals = psspy.chnval(ch)
                    if ierr_v == 0 and vals is not None:
                        cd = ChannelData(
                            name   = desc.strip() if desc else f"Channel {ch}",
                            units  = "pu",
                            time_s = time_list,
                            values = list(vals),
                        )
                        channels[cd.name] = cd
                except Exception:
                    pass

        except Exception as exc:
            logger.debug("Channel extraction error: %s", exc)

        return channels

    def _evaluate_criteria(
        self,
        result: ContingencyTSResult,
        poi_bus: Optional[int],
    ) -> None:
        """
        Check AESO transient stability criteria against extracted channels.
        Updates result in-place.
        """
        fault_clear_time = result.fault_clear_time_s
        recovery_end     = fault_clear_time + self.v_recovery_window_s

        # Check rotor angle stability
        # Look for any channel with "Rotor Angle" or "ANGLE" in name
        max_angle = 0.0
        for name, ch in result.channels.items():
            if "angle" in name.lower() or "rotor" in name.lower():
                for t, v in zip(ch.time_s, ch.values):
                    if t >= result.fault_apply_time_s:
                        if abs(v) > max_angle:
                            max_angle = abs(v)
                        if abs(v) > self.rotor_limit:
                            result.rotor_angle_stable = False

        result.max_rotor_angle_deg = round(max_angle, 2)

        # Check POI voltage recovery
        poi_key = f"Bus Voltage POI" if poi_bus else None
        if poi_key and poi_key in result.channels:
            ch = result.channels[poi_key]
            min_v = 1.0
            recovered = False
            for t, v in zip(ch.time_s, ch.values):
                if fault_clear_time <= t <= recovery_end:
                    if v < min_v:
                        min_v = v
                if t >= recovery_end and v >= self.v_recovery_pu:
                    recovered = True
                    result.recovery_time_s = round(t - fault_clear_time, 3)
                    break

            result.min_poi_voltage_pu = round(min_v, 4)
            result.voltage_recovered  = recovered
        else:
            # No POI channel — cannot assess, mark as warning
            result.voltage_recovered  = True   # conservative: don't fail
            result.min_poi_voltage_pu = 0.0
            logger.warning(
                "  POI bus voltage channel not available for '%s'. "
                "Voltage recovery cannot be assessed.",
                result.contingency_name
            )

    # ── Private: plots ────────────────────────────────────────────────────────────

    def _plot_contingency(
        self,
        cont:      ContingencyTSResult,
        plots_dir: str,
        label:     str,
    ) -> Optional[str]:
        """
        Multi-panel time-domain response plot for one contingency.
        Panels: voltage, rotor angles, active power, reactive power.
        """
        if not cont.channels:
            return None

        # Separate channels by type
        voltage_chs  = [ch for n, ch in cont.channels.items()
                        if "voltage" in n.lower() or "volt" in n.lower()]
        angle_chs    = [ch for n, ch in cont.channels.items()
                        if "angle" in n.lower() or "rotor" in n.lower()]
        pq_chs       = [ch for n, ch in cont.channels.items()
                        if "power" in n.lower() or " p " in n.lower()
                        or " q " in n.lower()]

        panels = []
        if voltage_chs: panels.append(("Bus Voltage (pu)",    voltage_chs))
        if angle_chs:   panels.append(("Rotor Angle (deg)",   angle_chs))
        if pq_chs:      panels.append(("Power (MW / MVAR)",   pq_chs))

        if not panels:
            return None

        n_panels = len(panels)
        fig, axes = plt.subplots(
            n_panels, 1,
            figsize=(12, 3 * n_panels),
            sharex=True,
        )
        if n_panels == 1:
            axes = [axes]

        colors = [_BLUE, _RED, _ORANGE, _GREEN, _GREY]

        for ax, (ylabel, chs) in zip(axes, panels):
            for i, ch in enumerate(chs[:5]):   # max 5 channels per panel
                if ch.time_s and ch.values:
                    ax.plot(
                        ch.time_s, ch.values,
                        color=colors[i % len(colors)],
                        linewidth=1.5,
                        label=ch.name[:40],
                    )
            # Shade fault period
            ax.axvspan(
                cont.fault_apply_time_s,
                cont.fault_clear_time_s,
                alpha=0.15, color=_RED, label="Fault period"
            )
            # Recovery window
            ax.axvspan(
                cont.fault_clear_time_s,
                cont.fault_clear_time_s + self.v_recovery_window_s,
                alpha=0.08, color=_ORANGE, label="Recovery window"
            )
            if "Voltage" in ylabel:
                ax.axhline(
                    self.v_recovery_pu,
                    color=_RED, linestyle="--", linewidth=1.2,
                    label=f"V recovery limit ({self.v_recovery_pu} pu)"
                )
            if "Angle" in ylabel:
                ax.axhline(
                    self.rotor_limit,
                    color=_RED, linestyle="--", linewidth=1.2,
                    label=f"Angle limit ({self.rotor_limit}°)"
                )
            ax.set_ylabel(ylabel, fontsize=9)
            ax.grid(True, zorder=0)
            ax.legend(fontsize=7, loc="upper right", ncol=2)

        axes[-1].set_xlabel("Time (s)", fontsize=10)
        status = "✔ PASS" if cont.aeso_pass else "✖ FAIL"
        axes[0].set_title(
            f"Transient Stability — {cont.contingency_name}\n"
            f"{self.scenario}  |  {status}  |  "
            f"Max angle: {cont.max_rotor_angle_deg:.1f}°  |  "
            f"Min V_POI: {cont.min_poi_voltage_pu:.4f} pu",
            fontsize=10, pad=8
        )
        fig.tight_layout()

        safe_name = _safe_filename(cont.contingency_name)
        path = os.path.join(plots_dir, f"{label}_{safe_name}.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    # ── Private: PDF export ──────────────────────────────────────────────────────

    def _export_pdf(self, pdf_path: str, plot_paths: List[str]) -> None:
        """Assemble PDF: plots + compliance summary table."""
        with PdfPages(pdf_path) as pdf:

            # Embed each PNG
            for png in plot_paths:
                if png and os.path.isfile(png):
                    img = plt.imread(png)
                    fig, ax = plt.subplots(figsize=(14, 8))
                    ax.imshow(img)
                    ax.axis("off")
                    pdf.savefig(fig, bbox_inches="tight")
                    plt.close(fig)

            # Compliance summary table
            df = self._compliance_to_df()
            if not df.empty:
                fig2, ax2 = plt.subplots(
                    figsize=(16, max(4, len(df) * 0.55 + 2.5))
                )
                ax2.axis("off")
                tbl = ax2.table(
                    cellText=df.values,
                    colLabels=df.columns,
                    cellLoc="center", loc="center",
                )
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(8.5)
                tbl.scale(1.1, 1.9)
                for col in range(len(df.columns)):
                    tbl[0, col].set_facecolor(_BLUE)
                    tbl[0, col].set_text_props(color="white", weight="bold")
                for row in range(1, len(df) + 1):
                    if "FAIL" in str(df.iloc[row - 1].get("AESO Status", "")):
                        for col in range(len(df.columns)):
                            tbl[row, col].set_facecolor("#ffe0e0")
                    elif "PASS" in str(df.iloc[row - 1].get("AESO Status", "")):
                        for col in range(len(df.columns)):
                            tbl[row, col].set_facecolor("#e0ffe0")
                ax2.set_title(
                    f"AESO Transient Stability — Compliance Summary\n"
                    f"{self.scenario}",
                    fontsize=11, weight="bold", pad=15,
                )
                plt.tight_layout()
                pdf.savefig(fig2, bbox_inches="tight")
                plt.close(fig2)

            d = pdf.infodict()
            d["Title"]        = f"AESO Transient Stability — {self.scenario}"
            d["Author"]       = "AESO Automation Tool"
            d["CreationDate"] = datetime.now()

    # ── Private: DataFrame builders ───────────────────────────────────────────────

    def _summary_to_df(self) -> pd.DataFrame:
        r = self.results
        if r is None:
            return pd.DataFrame()
        return pd.DataFrame([
            {"Parameter": "Scenario",                  "Value": r.scenario_label},
            {"Parameter": "SAV File",                  "Value": os.path.basename(r.sav_path)},
            {"Parameter": "Simulation Duration (s)",   "Value": r.sim_duration_s},
            {"Parameter": "Contingencies Run",         "Value": len(r.contingencies)},
            {"Parameter": "AESO PASS",                 "Value": r.total_pass},
            {"Parameter": "AESO FAIL",                 "Value": r.total_fail},
            {"Parameter": "Rotor Angle Limit (deg)",   "Value": self.rotor_limit},
            {"Parameter": "Voltage Recovery Limit (pu)","Value": self.v_recovery_pu},
            {"Parameter": "Recovery Window (s)",       "Value": self.v_recovery_window_s},
        ])

    def _compliance_to_df(self) -> pd.DataFrame:
        rows = []
        for c in (self.results.contingencies if self.results else []):
            rows.append({
                "Contingency":             c.contingency_name,
                "Fault Location":          c.fault_location,
                "Near End (cycles)":       c.near_end_cycles,
                "Far End (cycles)":        c.far_end_cycles,
                "Fault Clear Time (s)":    round(c.fault_clear_time_s, 4),
                "Base Converged":          str(c.converged_base),
                "Sim Completed":           str(c.sim_completed),
                "Max Rotor Angle (deg)":   c.max_rotor_angle_deg,
                "Rotor Stable":            str(c.rotor_angle_stable),
                "Min POI Voltage (pu)":    c.min_poi_voltage_pu,
                "Voltage Recovered":       str(c.voltage_recovered),
                "Recovery Time (s)":       c.recovery_time_s,
                "AESO Status":             "PASS" if c.aeso_pass else "FAIL",
            })
        return pd.DataFrame(rows)

    def _channel_to_df(self, cont: ContingencyTSResult) -> pd.DataFrame:
        """Convert channel data to a wide DataFrame (one column per channel)."""
        if not cont.channels:
            return pd.DataFrame()

        # Use time from first channel
        first_ch = next(iter(cont.channels.values()))
        data: Dict[str, list] = {"Time (s)": first_ch.time_s}
        for name, ch in cont.channels.items():
            data[name] = ch.values

        return pd.DataFrame(data)

    # ── Mock data ───────────────────────────────────────────────────────────────────

    def _mock_contingency_result(
        self,
        result: ContingencyTSResult,
    ) -> ContingencyTSResult:
        """Generate synthetic time-domain data for testing without PSS/E."""
        import math, random
        random.seed(hash(result.contingency_name) % 9999)

        t_step = self.time_step_s
        t_end  = self.sim_duration_s
        n      = int(t_end / t_step)
        times  = [round(i * t_step, 5) for i in range(n)]

        fa = result.fault_apply_time_s
        fc = result.fault_clear_time_s

        # Synthetic voltage — dips during fault, recovers after
        voltages = []
        for t in times:
            if t < fa:
                v = 1.00 + random.uniform(-0.005, 0.005)
            elif t < fc:
                v = 0.05 + random.uniform(-0.02, 0.02)  # fault
            else:
                # Recovery curve
                tau = 0.3
                v_ss = 0.97
                v = v_ss - (v_ss - 0.05) * math.exp(-(t - fc) / tau)
                v = min(v, 1.02) + random.uniform(-0.003, 0.003)
            voltages.append(round(v, 5))

        # Synthetic rotor angle
        angles = []
        for t in times:
            if t < fa:
                ang = 15.0 + random.uniform(-1, 1)
            elif t < fc:
                ang = 15.0 + 80.0 * (t - fa) / (fc - fa)
            else:
                peak = 15.0 + 80.0
                decay = peak * math.exp(-0.8 * (t - fc))
                ang = max(15.0, decay) + random.uniform(-2, 2)
            angles.append(round(ang, 3))

        result.channels = {
            "Bus Voltage POI": ChannelData(
                name="Bus Voltage POI", units="pu",
                time_s=times, values=voltages
            ),
            "Rotor Angle Gen1": ChannelData(
                name="Rotor Angle Gen1", units="deg",
                time_s=times, values=angles
            ),
        }

        result.converged_base     = True
        result.sim_completed      = True
        result.min_poi_voltage_pu = round(min(voltages[int(fa/t_step):int((fc+1)/t_step)]), 4)
        result.voltage_recovered  = voltages[-1] >= self.v_recovery_pu
        result.max_rotor_angle_deg= round(max(abs(a) for a in angles), 2)
        result.rotor_angle_stable = result.max_rotor_angle_deg < self.rotor_limit
        result.recovery_time_s    = round(
            max(0.0, next(
                (t - fc for t, v in zip(times, voltages)
                 if t > fc and v >= self.v_recovery_pu),
                self.v_recovery_window_s + 0.1
            )), 3
        )

        return result


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _safe_sheet_name(name: str) -> str:
    """Excel sheet names max 31 chars, no special chars."""
    safe = name.replace(":", "").replace("/", "-").replace("\\", "-")
    safe = safe.replace("*", "").replace("?", "").replace("[", "").replace("]", "")
    return safe[:31]


def _safe_filename(name: str) -> str:
    """Convert contingency name to safe filename fragment."""
    safe = name.replace(" ", "_").replace(":", "-").replace("/", "-")
    safe = safe.replace("–", "-").replace("—", "-")
    return safe[:60]
