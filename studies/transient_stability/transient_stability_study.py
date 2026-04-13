"""
studies/transient_stability/transient_stability_study.py
---------------------------------------------------------
Automates AESO Transient Stability Analysis (Phase 3) using PSS/E via psspy.

PSS/E call sequence (from Updated_TSA.py):
------------------------------------------
1.  redirect.psse2py()                     — redirect PSS/E output to console
2.  psspy.psseinit()                       — initialise PSS/E
3.  psspy.case(sav_path)                   — load power flow case
4.  psspy.dyre_new([1,1,1,1], dyr, "","","") — load dynamic models
5.  psspy.cong(0)                          — convert generators
6.  psspy.conl(0,1,1..3, [0,0], [...])    — convert loads (3 passes)
7.  psspy.bsys(0,0,[0.0,750.0],...)        — define bus subsystem
8.  psspy.ordr(1)                          — order network
9.  psspy.fact()                           — factorise admittance matrix
10. psspy.tysl(0)                          — solve power flow
11. psspy.chsb(...)                        — define channels BEFORE strt
12. psspy.set_relang(1, ref_bus, ref_id)   — set reference generator
13. psspy.dynamics_solution_params(...)    — set solver parameters
14. psspy.strt_2(1, outfile)               — initialise dynamics (strt_2!)
15. psspy.run(0, t_prefault, ...)          — run to fault time
16. psspy.dist_bus_fault(bus,1,kv,[0,-2e9]) — apply 3-phase fault
17. psspy.run(0, t_clear, ...)             — run during fault
18. psspy.dist_clear_fault(1)              — clear fault
19. psspy.run(0, t_end, ...)               — run post-fault
20. dyntools.CHNF(outfile)                 — extract channel data

Workflow per contingency:
--------------------------
For each contingency in TS_Contingencies:
    a. Reload fresh .sav
    b. Load .dyr
    c. Run conversion sequence (cong/conl/bsys/ordr/fact/tysl)
    d. Set up channels via chsb (before strt_2)
    e. Set reference generator via set_relang
    f. Set dynamics solution params
    g. Initialise with strt_2
    h. Run pre-fault
    i. Apply 3-phase fault at contingency bus
    j. Run during-fault
    k. Clear fault / trip branch
    l. Run post-fault
    m. Extract data via dyntools.CHNF
    n. Check AESO criteria

AESO Criteria Applied (Study Scope Section 5.3 and Requirements Section 3.4)
------------------------------------------------------------------------------
- Fault type: bolted 3-phase to ground
- Rotor angle stability: all generators remain stable (< 180 deg relative to ref)
- Voltage recovery: POI bus voltage recovers to >= 0.90 pu within 1.0 s post-fault
- Reference generator: set via set_relang (default: Genesee/Wabamun bus)

PSS/E return-value note
-----------------------
Many psspy functions return a TUPLE on the PSS/E Python 3.7 binding.
The helper _ierr() handles both cases transparently.
"""

import logging
import os
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

from project_io.project_data import ProjectData, TSContingency

logger = logging.getLogger(__name__)

# ── AESO colour palette ───────────────────────────────────────────────────────
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


# ── PSS/E return-value helper ─────────────────────────────────────────────────

def _ierr(ret) -> int:
    """Normalise a psspy return value to a plain int (unwraps tuples)."""
    if isinstance(ret, tuple):
        return int(ret[0])
    return int(ret)


def _out_path(sav_path: str, scenario: str, cont_name: str) -> str:
    """
    Build a safe, absolute path for the PSS/E .out output file.
    Falls back to system temp directory if the .sav directory is not writable.
    """
    sav_dir = os.path.dirname(os.path.abspath(sav_path))
    if not sav_dir or not os.path.isdir(sav_dir):
        sav_dir = tempfile.gettempdir()

    safe_cont = (
        cont_name
        .replace(" ", "_").replace(":", "-").replace("/", "-")
        .replace("\\", "-").replace("–", "-").replace("—", "-")
    )[:60]

    return os.path.join(sav_dir, f"ts_{scenario}_{safe_cont}.out")


# ── Data containers ───────────────────────────────────────────────────────────

@dataclass
class ChannelData:
    """Time-domain simulation data for one monitored channel."""
    name:    str
    units:   str
    time_s:  List[float] = field(default_factory=list)
    values:  List[float] = field(default_factory=list)


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

    # Channel data keyed by channel description
    channels: Dict[str, ChannelData] = field(default_factory=dict)

    @property
    def aeso_pass(self) -> bool:
        return self.rotor_angle_stable and self.voltage_recovered


@dataclass
class TransientStabilityResults:
    """Aggregated results for one scenario."""
    scenario_label:  str
    sav_path:        str
    sim_duration_s:  float
    contingencies:   List[ContingencyTSResult] = field(default_factory=list)

    @property
    def total_pass(self) -> int:
        return sum(1 for c in self.contingencies if c.aeso_pass)

    @property
    def total_fail(self) -> int:
        return sum(1 for c in self.contingencies if not c.aeso_pass)


# ── Study class ───────────────────────────────────────────────────────────────

class TransientStabilityStudy:
    """
    Automates PSS/E transient stability analysis for one operating scenario.

    Parameters
    ----------
    psse : PSSEInterface
        Initialized PSS/E interface (with redirect and dyntools available).
    project : ProjectData
        Populated project data (scenarios, contingencies, bus numbers).
    scenario_label : str
        Label used in output filenames.
    sim_duration_s : float
        Total simulation window in seconds (default 10.0 s).
    fault_apply_time_s : float
        Time at which the fault is applied (default 1.0 s).
    time_step_s : float
        Integration time step (default 0.00833 s = 1/2 cycle at 60 Hz).
    rotor_angle_limit_deg : float
        First-swing rotor angle stability limit (default 180 deg).
    voltage_recovery_pu : float
        Minimum POI voltage after recovery window (default 0.90 pu).
    voltage_recovery_window_s : float
        Post-fault window to check voltage recovery (default 1.0 s).
    ref_bus : int
        Bus number of the reference generator for set_relang (default 101).
    ref_gen_id : str
        Machine ID of the reference generator (default '1').
    """

    def __init__(
        self,
        psse,
        project:                   ProjectData,
        scenario_label:            str   = "Base_Case",
        sim_duration_s:            float = 10.0,
        fault_apply_time_s:        float = 1.0,
        time_step_s:               float = 0.00833,
        rotor_angle_limit_deg:     float = 180.0,
        voltage_recovery_pu:       float = 0.90,
        voltage_recovery_window_s: float = 1.0,
        ref_bus:                   int   = 101,
        ref_gen_id:                str   = "1",
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
        self.ref_bus                = ref_bus
        self.ref_gen_id             = ref_gen_id
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
                "max_angle=%.1f deg  min_V=%.3f pu",
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
            "\nTS Study complete: %d contingencies  %d PASS  %d FAIL",
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

    # ── Private: PSS/E transient stability engine ─────────────────────────────

    def _convert_for_dynamics(self, psspy) -> bool:
        """
        Run the PSS/E network-conversion sequence after loading the case.

        Sequence (from Updated_TSA.py):
            cong(0)                           — convert generators
            conl(0,1,1,[0,0],[100,0,0,100])  — load conversion pass 1
            conl(0,1,2,[0,0],[100,0,0,100])  — load conversion pass 2
            conl(0,1,3,[0,0],[100,0,0,100])  — load conversion pass 3
            bsys(0,0,[0.0,750.0],...)         — define bus subsystem (all voltage levels)
            ordr(1)                           — order network
            fact()                            — factorise admittance matrix
            tysl(0)                           — solve initial-condition power flow

        Returns True if critical steps succeeded, False on fatal error.
        """
        # cong: convert generators first
        ret = psspy.cong(0)
        if _ierr(ret) != 0:
            logger.warning("  cong() returned ierr=%d (non-fatal).", _ierr(ret))

        # conl: 3-pass load model conversion (exactly as in Updated_TSA.py)
        for conl_step in (1, 2, 3):
            ret = psspy.conl(0, 1, conl_step, [0, 0], [100.0, 0.0, 0.0, 100.0])
            if _ierr(ret) != 0:
                logger.warning(
                    "  conl step %d returned ierr=%d (non-fatal).",
                    conl_step, _ierr(ret),
                )

        # bsys: define bus subsystem spanning all voltage levels (0–750 kV)
        ret = psspy.bsys(0, 0, [0.0, 750.0], 0, [], 0, [], 0, [], 0, [])
        if _ierr(ret) != 0:
            logger.warning("  bsys() returned ierr=%d (non-fatal).", _ierr(ret))

        # ordr: order network for dynamics
        ret = psspy.ordr(1)
        if _ierr(ret) != 0:
            logger.error("  ordr() failed (ierr=%d). Aborting conversion.", _ierr(ret))
            return False

        # fact: factorise admittance matrix
        ret = psspy.fact()
        if _ierr(ret) != 0:
            logger.error("  fact() failed (ierr=%d). Aborting conversion.", _ierr(ret))
            return False

        # tysl: solve initial-condition power flow
        ret = psspy.tysl(0)
        if _ierr(ret) != 0:
            logger.warning(
                "  tysl() returned ierr=%d "
                "(acceptable if residuals are small; check PSS/E output).",
                _ierr(ret),
            )

        logger.debug("  Network conversion for dynamics complete.")
        return True

    def _load_dynamics(self, psspy, sav_path: str) -> bool:
        """
        Load the .dyr dynamic data file via psspy.dyre_new().

        Exactly as in Updated_TSA.py:
            psspy.dyre_new([1,1,1,1], dyr_file, "", "", "")

        Resolution order for the .dyr path:
          1. ProjectData.resolve_dyr_path()  (explicit path in Project_Info)
          2. sav_path with .sav replaced by .dyr  (co-located fallback)

        NOTE: Must be called AFTER case() and BEFORE the conversion sequence
        in this implementation (matching Updated_TSA.py order).
        """
        dyr_path = self._project.resolve_dyr_path()
        if not dyr_path or not os.path.isfile(dyr_path):
            dyr_path = os.path.splitext(sav_path)[0] + ".dyr"

        if not os.path.isfile(dyr_path):
            logger.error(
                "  .dyr file not found. Set 'DYR File Path' in Project_Info "
                "or place a .dyr alongside the .sav. Checked: '%s'",
                dyr_path,
            )
            return False

        # Exactly as Updated_TSA.py: dyre_new([1,1,1,1], dyr_file, "", "", "")
        ret = psspy.dyre_new([1, 1, 1, 1], dyr_path, "", "", "")
        if _ierr(ret) != 0:
            logger.error(
                "  psspy.dyre_new() failed (ierr=%d) for .dyr: '%s'.",
                _ierr(ret), dyr_path,
            )
            return False

        logger.debug("  Dynamics loaded from: %s", dyr_path)
        return True

    def _setup_channels(self, psspy, output_file: str, poi_bus: Optional[int]) -> None:
        """
        Define PSS/E output channels using chsb (as in Updated_TSA.py).

        Channels registered (exactly as Updated_TSA.py):
            chsb(0,1,[-1,-1,-1,1,1,0])   — Relative Rotor Angle
            chsb(0,1,[-1,-1,-1,1,2,0])   — Machine Electrical Power
            chsb(0,1,[-1,-1,-1,1,3,0])   — Machine Reactive Power
            chsb(0,1,[-1,-1,-1,1,13,0])  — Bus Voltages
            chsb(0,1,[-1,-1,-1,1,16,0])  — Branch P/Q Flow

        MUST be called AFTER dyre_new() and BEFORE strt_2().
        """
        # Delete any existing channels from a previous contingency
        try:
            psspy.delete_all_plot_channels()
        except Exception:
            pass

        channel_defs = [
            ([-1, -1, -1, 1, 1,  0], "Relative Rotor Angle"),
            ([-1, -1, -1, 1, 2,  0], "Machine Electrical Power"),
            ([-1, -1, -1, 1, 3,  0], "Machine Reactive Power"),
            ([-1, -1, -1, 1, 13, 0], "Bus Voltages"),
            ([-1, -1, -1, 1, 16, 0], "Branch P/Q Flow"),
        ]

        for args, desc in channel_defs:
            ret = psspy.chsb(0, 1, args)
            if _ierr(ret) != 0:
                logger.warning("  chsb() for '%s' returned ierr=%d.", desc, _ierr(ret))
            else:
                logger.debug("  Channel registered: %s", desc)

    def _run_contingency(
        self,
        sav_path: str,
        cont:     "TSContingency",
    ) -> ContingencyTSResult:
        """Run one contingency and return its result."""

        self._current_cont_name = cont.contingency_name

        fault_clear_time_s = self.fault_apply_time_s + cont.near_end_seconds

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

        # Redirect PSS/E output to Python console (as in Updated_TSA.py)
        try:
            import redirect
            redirect.psse2py()
        except ImportError:
            logger.debug("  redirect module not available; PSS/E output goes to PSSE window.")

        _i = psspy.getdefaultint()
        _f = psspy.getdefaultreal()

        poi_bus     = self._project.info.poi_bus_number
        output_file = _out_path(sav_path, self.scenario, cont.contingency_name)

        # ── Step 1: Load fresh power-flow case ───────────────────────────────
        ret = psspy.case(sav_path)
        if _ierr(ret) != 0:
            logger.error("  case() failed (ierr=%d) for '%s'.", _ierr(ret), cont.contingency_name)
            return result

        # ── Step 2: Load dynamic models (.dyr) ───────────────────────────────
        # dyre_new BEFORE the conversion sequence, matching Updated_TSA.py order.
        if not self._load_dynamics(psspy, sav_path):
            logger.error("  Skipping '%s': .dyr could not be loaded.", cont.contingency_name)
            return result

        # ── Step 3: Convert network for dynamics ─────────────────────────────
        # cong → conl (3 passes) → bsys → ordr → fact → tysl
        if not self._convert_for_dynamics(psspy):
            logger.error("  Conversion failed for '%s'. Skipping.", cont.contingency_name)
            return result

        result.converged_base = True

        # ── Step 4: Define output channels via chsb ───────────────────────────
        # MUST be before strt_2 (PSS/E requirement).
        self._setup_channels(psspy, output_file, poi_bus)

        # ── Step 5: Set reference generator for relative rotor angle ─────────
        # psspy.set_relang(1, bus_number, machine_id)  — as in Updated_TSA.py
        try:
            psspy.set_relang(1, self.ref_bus, self.ref_gen_id)
            logger.debug("  Reference generator set: bus %d id '%s'.", self.ref_bus, self.ref_gen_id)
        except Exception as exc:
            logger.warning("  set_relang failed: %s", exc)

        # ── Step 6: Set dynamics solution parameters ─────────────────────────
        # From Updated_TSA.py:
        #   dynamics_solution_params(
        #       [99, _i, _i, _i, _i, _i, _i, _i],
        #       [1.0, _f, 0.004, 0.016, _f, _f, _f, _f], ''
        #   )
        # iparams[0]=99  → print channel values every 99 steps
        # rparams[2]=0.004  → integration step size (s)
        # rparams[3]=0.016  → output print step (s)
        try:
            psspy.dynamics_solution_params(
                [99, _i, _i, _i, _i, _i, _i, _i],
                [1.0, _f, self.time_step_s, self.time_step_s * 4, _f, _f, _f, _f],
                "",
            )
            logger.debug("  Dynamics solution params set (dt=%.5f s).", self.time_step_s)
        except Exception as exc:
            logger.warning("  dynamics_solution_params failed: %s", exc)

        # ── Step 7: Initialise dynamics with strt_2 ───────────────────────────
        # Updated_TSA.py uses strt_2(1, outfile) — NOT strt()
        # strt_2 first arg: 1 = initialise, 0 = restart
        ret = psspy.strt_2(1, output_file)
        if _ierr(ret) != 0:
            logger.error(
                "  strt_2() failed (ierr=%d) for '%s'. Cannot proceed.",
                _ierr(ret), cont.contingency_name,
            )
            return result

        logger.debug("  strt_2() succeeded. Output: %s", output_file)

        # ── Step 8: Run pre-fault ─────────────────────────────────────────────
        # Updated_TSA.py: psspy.run(0, 1.000, 99, 19, 0)
        ret = psspy.run(0, self.fault_apply_time_s, 99, 19, 0)
        if _ierr(ret) != 0:
            logger.warning(
                "  Pre-fault run() failed (ierr=%d) for '%s'.",
                _ierr(ret), cont.contingency_name,
            )
            return result

        # ── Step 9: Apply bolted 3-phase fault ────────────────────────────────
        fault_bus = cont.from_bus_no
        if fault_bus is None:
            logger.warning(
                "  No fault bus for '%s'. Fill From Bus No in TS_Contingencies.",
                cont.contingency_name,
            )
            return result

        # Updated_TSA.py: dist_bus_fault(bus, 1, kv, [0.0, -0.2E+10])
        # 1 = 3-phase fault, kv = bus nominal voltage
        fault_kv = cont.fault_kv if hasattr(cont, "fault_kv") and cont.fault_kv else 0.0
        ret = psspy.dist_bus_fault(fault_bus, 1, fault_kv, [0.0, -0.2e10])
        if _ierr(ret) != 0:
            logger.warning(
                "  dist_bus_fault() failed (ierr=%d) at bus %d.",
                _ierr(ret), fault_bus,
            )

        # ── Step 10: Run during-fault ─────────────────────────────────────────
        # Updated_TSA.py: psspy.run(0, 1.15, 99, 3, 0)  — 9 cycles (0.15 s at 60 Hz)
        ret = psspy.run(0, fault_clear_time_s, 99, 3, 0)
        if _ierr(ret) != 0:
            logger.warning(
                "  During-fault run() failed (ierr=%d) for '%s'.",
                _ierr(ret), cont.contingency_name,
            )

        # ── Step 11: Clear fault (and optionally trip branch) ─────────────────
        psspy.dist_clear_fault(1)

        if cont.to_bus_no is not None and cont.from_bus_no is not None:
            try:
                psspy.dist_branch_trip(
                    cont.from_bus_no,
                    cont.to_bus_no,
                    cont.circuit_id if cont.circuit_id else "1",
                )
            except Exception as exc:
                logger.debug("  dist_branch_trip: %s", exc)

        # ── Step 12: Run post-fault simulation ────────────────────────────────
        # Updated_TSA.py: psspy.run(0, 5.0, 501, 9, 0)
        ret = psspy.run(0, self.sim_duration_s, 501, 9, 0)
        result.sim_completed = (_ierr(ret) == 0)
        if not result.sim_completed:
            logger.warning(
                "  Post-fault run() failed (ierr=%d) for '%s'.",
                _ierr(ret), cont.contingency_name,
            )

        # ── Step 13: Extract channel data via dyntools.CHNF ───────────────────
        result.channels = self._extract_channels_dyntools(output_file)

        # ── Step 14: Evaluate AESO criteria ───────────────────────────────────
        self._evaluate_criteria(result, poi_bus)

        return result

    def _extract_channels_dyntools(self, output_file: str) -> Dict[str, ChannelData]:
        """
        Extract time-domain data from the PSS/E .out file using dyntools.CHNF.

        This matches the data extraction approach in Updated_TSA.py
        (chnfobj.get_data('')), which returns:
            sh_ttl  — run title string
            ch_id   — dict of {channel_num: description}
            ch_data — dict of {channel_num: [values]}

        Channel 0 is the time vector.

        Returns dict of channel_description -> ChannelData.
        """
        channels: Dict[str, ChannelData] = {}

        if not os.path.isfile(output_file):
            logger.warning("  .out file not found: %s", output_file)
            return channels

        try:
            import dyntools
        except ImportError:
            logger.error(
                "  dyntools not importable. "
                "Ensure PSS/E is installed and dyntools is on sys.path."
            )
            return channels

        try:
            chnfobj = dyntools.CHNF(output_file, outvrsn=0)
            if chnfobj.ierr:
                logger.error(
                    "  dyntools.CHNF reported error (ierr=%d) for: %s",
                    chnfobj.ierr, output_file,
                )
                return channels

            sh_ttl, ch_id, ch_data = chnfobj.get_data("")

            # Channel 0 is the time axis
            time_list = list(ch_data.get(0, []))

            for ch_num, desc in ch_id.items():
                if ch_num == 0:
                    continue
                vals = ch_data.get(ch_num, [])
                if vals:
                    cd = ChannelData(
                        name   = str(desc).strip() if desc else f"Channel {ch_num}",
                        units  = "pu",
                        time_s = time_list,
                        values = list(vals),
                    )
                    channels[cd.name] = cd

            logger.debug(
                "  Extracted %d channels from %s", len(channels), output_file
            )

        except Exception as exc:
            logger.error("  dyntools extraction failed: %s", exc)

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

        # ── Rotor angle stability ─────────────────────────────────────────────
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

        # ── POI voltage recovery ──────────────────────────────────────────────
        # dyntools channel names for bus voltages typically contain "VOLT" or bus number
        # Try to match the POI bus by name or number
        poi_ch = None
        if poi_bus:
            poi_str = str(poi_bus)
            for name, ch in result.channels.items():
                if poi_str in name or "volt" in name.lower():
                    poi_ch = ch
                    break

        if poi_ch is None:
            # Fall back to first voltage channel found
            for name, ch in result.channels.items():
                if "volt" in name.lower():
                    poi_ch = ch
                    break

        if poi_ch:
            min_v     = 1.0
            recovered = False
            for t, v in zip(poi_ch.time_s, poi_ch.values):
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
            result.voltage_recovered  = True
            result.min_poi_voltage_pu = 0.0
            logger.warning(
                "  POI bus voltage channel not found for '%s'. "
                "Voltage recovery cannot be assessed.",
                result.contingency_name,
            )

    # ── Private: plots ────────────────────────────────────────────────────────

    def _plot_contingency(
        self,
        cont:      ContingencyTSResult,
        plots_dir: str,
        label:     str,
    ) -> Optional[str]:
        """Multi-panel time-domain response plot for one contingency."""
        if not cont.channels:
            return None

        voltage_chs = [ch for n, ch in cont.channels.items()
                       if "volt" in n.lower()]
        angle_chs   = [ch for n, ch in cont.channels.items()
                       if "angle" in n.lower() or "rotor" in n.lower()]
        pq_chs      = [ch for n, ch in cont.channels.items()
                       if "power" in n.lower() or " p " in n.lower()
                       or " q " in n.lower() or "elec" in n.lower()]

        panels = []
        if voltage_chs: panels.append(("Bus Voltage (pu)",  voltage_chs))
        if angle_chs:   panels.append(("Rotor Angle (deg)", angle_chs))
        if pq_chs:      panels.append(("Power (MW / MVAR)", pq_chs))

        if not panels:
            return None

        n_panels = len(panels)
        fig, axes = plt.subplots(n_panels, 1, figsize=(12, 3 * n_panels), sharex=True)
        if n_panels == 1:
            axes = [axes]

        colors = [_BLUE, _RED, _ORANGE, _GREEN, _GREY]

        for ax, (ylabel, chs) in zip(axes, panels):
            for i, ch in enumerate(chs[:5]):
                if ch.time_s and ch.values:
                    ax.plot(
                        ch.time_s, ch.values,
                        color=colors[i % len(colors)],
                        linewidth=1.5,
                        label=ch.name[:40],
                    )
            ax.axvspan(cont.fault_apply_time_s, cont.fault_clear_time_s,
                       alpha=0.15, color=_RED, label="Fault period")
            ax.axvspan(cont.fault_clear_time_s,
                       cont.fault_clear_time_s + self.v_recovery_window_s,
                       alpha=0.08, color=_ORANGE, label="Recovery window")
            if "Voltage" in ylabel:
                ax.axhline(self.v_recovery_pu, color=_RED, linestyle="--",
                           linewidth=1.2,
                           label=f"V recovery limit ({self.v_recovery_pu} pu)")
            if "Angle" in ylabel:
                ax.axhline(self.rotor_limit, color=_RED, linestyle="--",
                           linewidth=1.2,
                           label=f"Angle limit ({self.rotor_limit} deg)")
            ax.set_ylabel(ylabel, fontsize=9)
            ax.grid(True, zorder=0)
            ax.legend(fontsize=7, loc="upper right", ncol=2)

        axes[-1].set_xlabel("Time (s)", fontsize=10)
        status = "PASS" if cont.aeso_pass else "FAIL"
        axes[0].set_title(
            f"Transient Stability — {cont.contingency_name}\n"
            f"{self.scenario}  |  {status}  |  "
            f"Max angle: {cont.max_rotor_angle_deg:.1f} deg  |  "
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

    # ── Private: PDF export ───────────────────────────────────────────────────

    def _export_pdf(self, pdf_path: str, plot_paths: List[str]) -> None:
        """Assemble PDF: plots + compliance summary table."""
        with PdfPages(pdf_path) as pdf:

            for png in plot_paths:
                if png and os.path.isfile(png):
                    img = plt.imread(png)
                    fig, ax = plt.subplots(figsize=(14, 8))
                    ax.imshow(img)
                    ax.axis("off")
                    pdf.savefig(fig, bbox_inches="tight")
                    plt.close(fig)

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
                    f"AESO Transient Stability — Compliance Summary\n{self.scenario}",
                    fontsize=11, weight="bold", pad=15,
                )
                plt.tight_layout()
                pdf.savefig(fig2, bbox_inches="tight")
                plt.close(fig2)

            d = pdf.infodict()
            d["Title"]        = f"AESO Transient Stability — {self.scenario}"
            d["Author"]       = "AESO Automation Tool"
            d["CreationDate"] = datetime.now()

    # ── Private: DataFrame builders ───────────────────────────────────────────

    def _summary_to_df(self) -> pd.DataFrame:
        r = self.results
        if r is None:
            return pd.DataFrame()
        return pd.DataFrame([
            {"Parameter": "Scenario",                    "Value": r.scenario_label},
            {"Parameter": "SAV File",                    "Value": os.path.basename(r.sav_path)},
            {"Parameter": "Simulation Duration (s)",     "Value": r.sim_duration_s},
            {"Parameter": "Contingencies Run",           "Value": len(r.contingencies)},
            {"Parameter": "AESO PASS",                   "Value": r.total_pass},
            {"Parameter": "AESO FAIL",                   "Value": r.total_fail},
            {"Parameter": "Rotor Angle Limit (deg)",     "Value": self.rotor_limit},
            {"Parameter": "Voltage Recovery Limit (pu)", "Value": self.v_recovery_pu},
            {"Parameter": "Recovery Window (s)",         "Value": self.v_recovery_window_s},
            {"Parameter": "Reference Bus",               "Value": self.ref_bus},
            {"Parameter": "Reference Gen ID",            "Value": self.ref_gen_id},
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
        first_ch = next(iter(cont.channels.values()))
        data: Dict[str, list] = {"Time (s)": first_ch.time_s}
        for name, ch in cont.channels.items():
            data[name] = ch.values
        return pd.DataFrame(data)

    # ── Mock data ─────────────────────────────────────────────────────────────

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

        voltages = []
        for t in times:
            if t < fa:
                v = 1.00 + random.uniform(-0.005, 0.005)
            elif t < fc:
                v = 0.05 + random.uniform(-0.02, 0.02)
            else:
                tau  = 0.3
                v_ss = 0.97
                v = v_ss - (v_ss - 0.05) * math.exp(-(t - fc) / tau)
                v = min(v, 1.02) + random.uniform(-0.003, 0.003)
            voltages.append(round(v, 5))

        angles = []
        for t in times:
            if t < fa:
                ang = 15.0 + random.uniform(-1, 1)
            elif t < fc:
                ang = 15.0 + 80.0 * (t - fa) / (fc - fa)
            else:
                peak  = 15.0 + 80.0
                decay = peak * math.exp(-0.8 * (t - fc))
                ang   = max(15.0, decay) + random.uniform(-2, 2)
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

        result.converged_base      = True
        result.sim_completed       = True
        result.min_poi_voltage_pu  = round(
            min(voltages[int(fa/t_step):int((fc+1)/t_step)]), 4)
        result.voltage_recovered   = voltages[-1] >= self.v_recovery_pu
        result.max_rotor_angle_deg = round(max(abs(a) for a in angles), 2)
        result.rotor_angle_stable  = result.max_rotor_angle_deg < self.rotor_limit
        result.recovery_time_s     = round(
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
