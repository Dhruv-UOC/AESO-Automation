"""
studies/power_flow/power_flow_study.py
---------------------------------------
Automates AESO Power Flow Analysis (Phase 1) using PSS/E via psspy.

Workflow
--------
1.  Load base case (.sav)
2.  Solve Newton-Raphson power flow
3.  Extract bus voltages, angles, line flows, system totals
4.  Flag violations against AESO thresholds
5.  Run N-1 contingency analysis (always enabled)
6.  Export results to Excel + PDF report + PNG plots

AESO Criteria Applied
---------------------
Category A (N-0) : V_min = 0.95 pu, V_max = 1.05 pu  (normal operation)
Category B (N-1) : V_min = 0.90 pu, V_max = 1.10 pu  (post-contingency)
Thermal loading  : ≤ 100% of rated MVA (seasonal continuous rating)
Post-contingency voltage deviation:
    ≤ ±10%  within 30 seconds  (post-transient)
    ≤ ±7%   after auto controls (30 sec – 5 min)
    ≤ ±5%   steady state       (post-manual)

Usage
-----
    from core.psse_interface import PSSEInterface
    from studies.power_flow.power_flow_study import PowerFlowStudy
    from config.settings import PSSE_PATH, PSSE_VERSION, DEFAULT_SAV, AESO

    psse = PSSEInterface(psse_path=PSSE_PATH, psse_version=PSSE_VERSION)
    psse.initialize()

    study = PowerFlowStudy(psse, scenario_label="Sample_Case")
    study.run(DEFAULT_SAV)
    study.save_results("output/results", "output/plots", "output/reports")
"""

import os
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING, List, Optional, Tuple

if TYPE_CHECKING:
    from project_io.project_data import ProjectData

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages

logger = logging.getLogger(__name__)

# ── AESO colour palette (consistent across all study modules) ─────────────────
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


# ── Data containers ───────────────────────────────────────────────────────────

@dataclass
class BusResult:
    bus_number:     int
    bus_name:       str
    base_kv:        float
    voltage_pu:     float
    angle_deg:      float
    voltage_kv:     float
    bus_type:       int    # 1=load, 2=generator, 3=slack
    violation:      bool  = False
    violation_type: str   = ""


@dataclass
class BranchResult:
    from_bus:    int
    to_bus:      int
    circuit_id:  str
    mw_flow:     float
    mvar_flow:   float
    mva_flow:    float
    rating_mva:  float
    loading_pct: float
    violation:   bool = False


@dataclass
class ContingencyResult:
    contingency_name:  str
    element_removed:   str
    converged:         bool
    bus_violations:    List[BusResult]   = field(default_factory=list)
    branch_violations: List[BranchResult] = field(default_factory=list)
    max_voltage_pu:    float = 0.0
    min_voltage_pu:    float = 0.0


@dataclass
class PowerFlowResults:
    scenario:            str
    converged:           bool
    total_generation_mw: float
    total_load_mw:       float
    total_losses_mw:     float
    buses:               List[BusResult]       = field(default_factory=list)
    branches:            List[BranchResult]    = field(default_factory=list)
    contingencies:       List[ContingencyResult] = field(default_factory=list)
    bus_violations:      List[BusResult]       = field(default_factory=list)
    branch_violations:   List[BranchResult]    = field(default_factory=list)


# ── Study class ───────────────────────────────────────────────────────────────

class PowerFlowStudy:
    """
    Automates PSS/E power flow analysis for one operating scenario.

    N-1 contingency analysis always runs automatically.

    Parameters
    ----------
    psse : PSSEInterface
        Initialized PSS/E interface object.
    scenario_label : str
        Human-readable label used in output filenames and reports.
    project : ProjectData, optional
        Populated project data. When provided, generator dispatch from
        Tables 4-4, 4-5, 4-6 of the Study Scope is applied before solving
        (calls psspy.machine_chng_2 for each unit with a known bus number).
        If None, the .sav file dispatch is used unchanged.
    season_label : str, optional
        Season label to look up in project.conv_gen and project.renewables
        dispatch_mw dicts (e.g. "2028 SP"). Required when project is given.
    voltage_min : float
        Category A minimum bus voltage (pu). Default = AESO 0.95 pu.
    voltage_max : float
        Category A maximum bus voltage (pu). Default = AESO 1.05 pu.
    voltage_min_contingency : float
        Category B post-contingency minimum voltage (pu). Default = 0.90 pu.
    voltage_max_contingency : float
        Category B post-contingency maximum voltage (pu). Default = 1.10 pu.
    thermal_limit_pct : float
        Branch loading % that triggers a thermal violation. Default = 100.0.
    """

    def __init__(
        self,
        psse,
        scenario_label:          str            = "Base_Case",
        project:                 "ProjectData"  = None,
        season_label:            str            = "",
        voltage_min:             float          = 0.95,
        voltage_max:             float          = 1.05,
        voltage_min_contingency: float          = 0.90,
        voltage_max_contingency: float          = 1.10,
        thermal_limit_pct:       float          = 100.0,
    ):
        self._psse         = psse
        self.scenario      = scenario_label
        self._project      = project
        self._season_label = season_label
        self.v_min         = voltage_min
        self.v_max         = voltage_max
        self.v_min_cont    = voltage_min_contingency
        self.v_max_cont    = voltage_max_contingency
        self.thermal_limit = thermal_limit_pct
        self.results: Optional[PowerFlowResults] = None

    # ── Public ────────────────────────────────────────────────────────────────

    def run(self, sav_path: str) -> PowerFlowResults:
        """
        Run base case power flow + N-1 contingency analysis.

        Parameters
        ----------
        sav_path : str
            Path to the .sav case file.

        Returns
        -------
        PowerFlowResults
        """
        logger.info("=" * 68)
        logger.info("Power Flow Study  |  Scenario: %s", self.scenario)
        logger.info("=" * 68)

        # 1. Load case
        self._psse.load_case(sav_path)

        # 2. Apply generation dispatch from Study Scope Tables 4-4/4-5/4-6
        #    (only if project data and season label are provided)
        self._apply_dispatch()

        # 3. Solve base case power flow
        converged = self._solve_power_flow()
        if not converged:
            logger.warning("Base case power flow did not converge for %s.", self.scenario)

        # 4. Extract results
        buses    = self._extract_bus_results()
        branches = self._extract_branch_results()
        gen_mw, load_mw, loss_mw = self._extract_system_totals()

        # 5. Flag base case violations
        bus_viol    = [b for b in buses    if b.violation]
        branch_viol = [b for b in branches if b.violation]

        logger.info(
            "Base case: converged=%s  gen=%.1f MW  load=%.1f MW  "
            "losses=%.1f MW  bus_viol=%d  branch_viol=%d",
            converged, gen_mw, load_mw, loss_mw,
            len(bus_viol), len(branch_viol),
        )

        # 6. N-1 contingency (always runs)
        contingencies = self._run_n1_contingency(sav_path)

        self.results = PowerFlowResults(
            scenario=self.scenario,
            converged=converged,
            total_generation_mw=gen_mw,
            total_load_mw=load_mw,
            total_losses_mw=loss_mw,
            buses=buses,
            branches=branches,
            contingencies=contingencies,
            bus_violations=bus_viol,
            branch_violations=branch_viol,
        )
        return self.results

    def save_results(
        self,
        results_dir: str,
        plots_dir:   str,
        reports_dir: str,
    ) -> dict:
        """
        Export results to Excel, PNG plots, and a PDF report.

        Parameters
        ----------
        results_dir : str   Path for Excel output.
        plots_dir   : str   Path for PNG plot output.
        reports_dir : str   Path for PDF report output.

        Returns
        -------
        dict with keys 'excel', 'plots' (list), 'pdf'
        """
        if self.results is None:
            raise RuntimeError("No results to save. Call run() first.")

        for d in (results_dir, plots_dir, reports_dir):
            os.makedirs(d, exist_ok=True)

        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        label = f"power_flow_{self.scenario}_{ts}"

        # Excel
        excel_path = os.path.join(results_dir, f"{label}.xlsx")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            self._summary_to_df().to_excel(         writer, sheet_name="Summary",         index=False)
            self._buses_to_df().to_excel(            writer, sheet_name="Bus_Results",     index=False)
            self._branches_to_df().to_excel(         writer, sheet_name="Branch_Results",  index=False)
            self._violations_to_df().to_excel(       writer, sheet_name="Violations",      index=False)
            self._contingency_summary_to_df().to_excel(writer, sheet_name="N1_Contingency", index=False)
        logger.info("Excel saved: %s", excel_path)

        # PNG plots
        plot_paths = []
        plot_paths.append(self._plot_voltage_profile(plots_dir, label))
        plot_paths.append(self._plot_thermal_loading(plots_dir, label))
        if self.results.contingencies:
            plot_paths.append(self._plot_contingency_summary(plots_dir, label))

        # PDF report
        pdf_path = os.path.join(reports_dir, f"{label}.pdf")
        self._export_pdf(pdf_path, plot_paths)
        logger.info("PDF saved: %s", pdf_path)

        return {"excel": excel_path, "plots": plot_paths, "pdf": pdf_path}

    # ── Private: PSS/E calls ──────────────────────────────────────────────────

    def _apply_dispatch(self) -> None:
        """
        Set generator MW outputs in the loaded PSS/E case from project data.

        Reads conv_gen and renewables dispatch from the project's
        Study Scope tables (Tables 4-4, 4-5, 4-6) for the current
        season_label and calls psspy.machine_chng_2() for each unit
        that has a known bus number.

        Skipped entirely if:
          - mock mode is active (no real PSS/E calls)
          - self._project is None (no project data provided)
          - self._season_label is empty
        """
        if self._psse.mock:
            return
        if self._project is None or not self._season_label:
            logger.debug(
                "Dispatch not applied: project=%s season_label='%s'",
                self._project is not None, self._season_label
            )
            return

        psspy = self._psse.psspy
        conv_units, renewables = self._project.get_dispatch_for_season(
            self._season_label
        )
        applied = 0
        skipped = 0

        for gen in conv_units + renewables:
            if gen.bus_no is None:
                skipped += 1
                continue
            mw = gen.dispatch_mw.get(self._season_label, None)
            if mw is None:
                skipped += 1
                continue
            try:
                # machine_chng_2: set PGEN (realar1) and keep unit in-service
                ierr = psspy.machine_chng_2(
                    gen.bus_no,
                    r"""1""",
                    intgar1 = 1,           # machine in-service
                    realar1 = float(mw),   # PGEN in MW
                )
                if ierr == 0:
                    applied += 1
                else:
                    logger.warning(
                        "  machine_chng_2 failed for %s bus %d (ierr=%d).",
                        gen.facility_name, gen.bus_no, ierr
                    )
                    skipped += 1
            except Exception as exc:
                logger.warning(
                    "  Dispatch error for %s bus %d: %s",
                    gen.facility_name, gen.bus_no, exc
                )
                skipped += 1

        logger.info(
            "Dispatch applied: %d units set  |  %d skipped (no bus no or MW)",
            applied, skipped
        )

    def _solve_power_flow(self) -> bool:
        """Run Newton-Raphson power flow. Returns True if converged."""
        if self._psse.mock:
            return True

        psspy = self._psse.psspy

        # solution_parameters_4 signature for PSSE35:
        # intgar: 14-element list
        #   [0] = solution method (0=NR, 1=FDNS, 2=FDXS)
        #   [1] = tap adjustment (0=off, 1=on)
        #   [2] = area interchange control (0=off, 1=on)
        #   [3] = phase shift adjustment (0=off, 1=on)
        #   [4] = dc tap adjustment (0=off, 1=on)
        #   [5] = switched shunt adjustment (0=off, 1=on)
        #   [6] = flat start (0=off, 1=on)
        #   [7] = var limits (0=apply, 1=ignore)
        #   [8] = non-divergent solution (0=off, 1=on)
        #   [9..13] = reserved (0)
        # realar: 6-element list
        #   [0] = power mismatch tolerance (pu)
        #   [1] = voltage mismatch tolerance (pu)
        #   [2..5] = reserved (0.0)
        psspy.solution_parameters_4(
            [0, 1, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
            [0.0001, 0.0001, 0.0, 0.0, 0.0, 0.0],
        )

        # fnsl: Full Newton-Raphson solve
        # Options: [tap, area, phase, dcTap, swShunt, flatStart, varLim, nonDiv]
        ret = psspy.fnsl([1, 0, 1, 1, 1, 0, 0, 0])
        return ret == 0

    def _extract_bus_results(self) -> List[BusResult]:
        """Extract voltage, angle, base kV, and type for every bus."""
        if self._psse.mock:
            return self._mock_bus_results()

        psspy = self._psse.psspy
        results = []

        ierr, (voltages,) = psspy.abusreal(-1, 1, ["PU"])
        ierr, (angles,)   = psspy.abusreal(-1, 1, ["ANGLED"])
        ierr, (base_kvs,) = psspy.abusreal(-1, 1, ["BASE"])
        ierr, (bus_nums,) = psspy.abusint(-1, 1, ["NUMBER"])
        ierr, (bus_types,)= psspy.abusint(-1, 1, ["TYPE"])
        ierr, (bus_names,)= psspy.abuschar(-1, 1, ["NAME"])

        for i, bnum in enumerate(bus_nums):
            v_pu     = voltages[i]
            v_kv     = v_pu * base_kvs[i]
            violation = v_pu < self.v_min or v_pu > self.v_max
            vtype    = ""
            if v_pu < self.v_min:
                vtype = f"LOW  ({v_pu:.4f} pu < {self.v_min} pu)"
            elif v_pu > self.v_max:
                vtype = f"HIGH ({v_pu:.4f} pu > {self.v_max} pu)"

            results.append(BusResult(
                bus_number  = bnum,
                bus_name    = bus_names[i].strip(),
                base_kv     = base_kvs[i],
                voltage_pu  = round(v_pu, 5),
                angle_deg   = round(angles[i], 4),
                voltage_kv  = round(v_kv, 3),
                bus_type    = bus_types[i],
                violation   = violation,
                violation_type = vtype,
            ))

        return results

    def _extract_branch_results(self) -> List[BranchResult]:
        """Extract MW/MVAR/MVA flows and % loading for every in-service branch."""
        if self._psse.mock:
            return self._mock_branch_results()

        psspy = self._psse.psspy
        results = []

        ierr, (from_buses,) = psspy.abrnint(-1, -1, -1, 1, 1, ["FROMNUMBER"])
        ierr, (to_buses,)   = psspy.abrnint(-1, -1, -1, 1, 1, ["TONUMBER"])
        ierr, (ckt_ids,)    = psspy.abrnchar(-1, -1, -1, 1, 1, ["ID"])
        ierr, (mw_flows,)   = psspy.abrnreal(-1, -1, -1, 1, 1, ["P"])
        ierr, (mvar_flows,) = psspy.abrnreal(-1, -1, -1, 1, 1, ["Q"])
        ierr, (mva_flows,)  = psspy.abrnreal(-1, -1, -1, 1, 1, ["MVA"])
        ierr, (ratings,)    = psspy.abrnreal(-1, -1, -1, 1, 1, ["RATEA"])

        for i in range(len(from_buses)):
            rating   = ratings[i] if ratings[i] > 0 else 9999.0
            load_pct = (mva_flows[i] / rating) * 100.0
            violation = load_pct > self.thermal_limit

            results.append(BranchResult(
                from_bus    = from_buses[i],
                to_bus      = to_buses[i],
                circuit_id  = ckt_ids[i].strip(),
                mw_flow     = round(mw_flows[i], 2),
                mvar_flow   = round(mvar_flows[i], 2),
                mva_flow    = round(mva_flows[i], 2),
                rating_mva  = round(rating, 2),
                loading_pct = round(load_pct, 2),
                violation   = violation,
            ))

        return results

    def _extract_system_totals(self) -> Tuple[float, float, float]:
        """
        Return (total_gen_mw, total_load_mw, total_losses_mw).

        Uses psspy.systot() which returns system-wide generation, load,
        and loss totals directly from PSS/E's internal accounting.
        This is the correct method — not Gen minus Load approximation.

        systot(string) returns (ierr, (cmpval,)) where cmpval is a
        complex number: real part = MW, imag part = MVAR.
        Valid string arguments: "GEN", "LOAD", "LOSS", "BUS", "LINE",
        "TRANSFORMER", "CHARGE", "INDMACH"
        """
        if self._psse.mock:
            return (5000.0, 4900.0, 100.0)

        psspy = self._psse.psspy

        try:
            ierr, cmpval = psspy.systot("GEN")
            total_gen_mw = cmpval.real if ierr == 0 else 0.0

            ierr, cmpval = psspy.systot("LOAD")
            total_load_mw = cmpval.real if ierr == 0 else 0.0

            ierr, cmpval = psspy.systot("LOSS")
            total_loss_mw = cmpval.real if ierr == 0 else 0.0

        except Exception as exc:
            logger.warning(
                "systot() failed (%s). Falling back to Gen-Load approximation.", exc
            )
            # Fallback: sum MW from all generator buses
            try:
                ierr, (gen_mws,)  = psspy.agereal(-1, 1, ["PGEN"])
                ierr, (load_mws,) = psspy.aloadreal(-1, 1, ["MVAACT"])
                total_gen_mw  = sum(g for g in gen_mws  if g is not None)
                total_load_mw = sum(abs(l.real) for l in load_mws if l is not None)
                total_loss_mw = total_gen_mw - total_load_mw
            except Exception as exc2:
                logger.warning("Fallback extraction also failed (%s). Using zeros.", exc2)
                total_gen_mw = total_load_mw = total_loss_mw = 0.0

        return (
            round(total_gen_mw, 2),
            round(total_load_mw, 2),
            round(abs(total_loss_mw), 2),
        )

    def _run_n1_contingency(self, base_sav: str) -> List[ContingencyResult]:
        """
        N-1 contingency: open each in-service branch one at a time,
        re-solve power flow, check violations against Category B limits,
        then restore the base case.
        """
        logger.info("Running N-1 contingency analysis ...")

        if self._psse.mock:
            return self._mock_contingency_results()

        psspy = self._psse.psspy
        contingencies = []

        # Get branch list from base case (already loaded)
        ierr, (from_buses,) = psspy.abrnint(-1, -1, -1, 1, 1, ["FROMNUMBER"])
        ierr, (to_buses,)   = psspy.abrnint(-1, -1, -1, 1, 1, ["TONUMBER"])
        ierr, (ckt_ids,)    = psspy.abrnchar(-1, -1, -1, 1, 1, ["ID"])

        total = len(from_buses)
        logger.info("  %d branches to analyse ...", total)

        for i in range(total):
            fb  = from_buses[i]
            tb  = to_buses[i]
            cid = ckt_ids[i].strip()
            name = f"Branch_{fb}_{tb}_{cid}"

            # Reload base case to ensure clean state
            self._psse.load_case(base_sav)

            # Open the branch (status 0 = out of service)
            psspy.branch_chng_3(fb, tb, cid,
                                [0, psspy._i, psspy._i, psspy._i,
                                 psspy._i, psspy._i],
                                [psspy._f, psspy._f, psspy._f,
                                 psspy._f, psspy._f, psspy._f,
                                 psspy._f, psspy._f, psspy._f,
                                 psspy._f, psspy._f, psspy._f])

            # Re-solve with contingency voltage limits
            ret = psspy.fnsl([1, 0, 1, 1, 1, 0, 0, 0])
            converged = (ret == 0)

            # Extract results using contingency (Category B) voltage limits
            buses    = self._extract_bus_results_cont()
            branches = self._extract_branch_results()
            bv       = [b for b in buses    if b.violation]
            brv      = [b for b in branches if b.violation]
            v_all    = [b.voltage_pu for b in buses]

            contingencies.append(ContingencyResult(
                contingency_name  = name,
                element_removed   = name,
                converged         = converged,
                bus_violations    = bv,
                branch_violations = brv,
                max_voltage_pu    = max(v_all) if v_all else 0.0,
                min_voltage_pu    = min(v_all) if v_all else 0.0,
            ))

            logger.debug(
                "  N-1 %-35s  converged=%-5s  viol=%d",
                name, converged, len(bv) + len(brv)
            )

        # Reload base case to leave PSS/E in clean state
        self._psse.load_case(base_sav)
        self._solve_power_flow()

        logger.info(
            "N-1 complete: %d contingencies analysed  "
            "(%d with violations)",
            len(contingencies),
            sum(1 for c in contingencies
                if c.bus_violations or c.branch_violations)
        )
        return contingencies

    def _extract_bus_results_cont(self) -> List[BusResult]:
        """
        Extract bus results using Category B (contingency) voltage limits.
        Used inside N-1 contingency loop only.
        """
        if self._psse.mock:
            return self._mock_bus_results()

        psspy = self._psse.psspy
        results = []

        ierr, (voltages,) = psspy.abusreal(-1, 1, ["PU"])
        ierr, (angles,)   = psspy.abusreal(-1, 1, ["ANGLED"])
        ierr, (base_kvs,) = psspy.abusreal(-1, 1, ["BASE"])
        ierr, (bus_nums,) = psspy.abusint(-1, 1, ["NUMBER"])
        ierr, (bus_types,)= psspy.abusint(-1, 1, ["TYPE"])
        ierr, (bus_names,)= psspy.abuschar(-1, 1, ["NAME"])

        for i, bnum in enumerate(bus_nums):
            v_pu      = voltages[i]
            v_kv      = v_pu * base_kvs[i]
            violation = v_pu < self.v_min_cont or v_pu > self.v_max_cont
            vtype     = ""
            if v_pu < self.v_min_cont:
                vtype = f"LOW  ({v_pu:.4f} pu < {self.v_min_cont} pu Cat-B)"
            elif v_pu > self.v_max_cont:
                vtype = f"HIGH ({v_pu:.4f} pu > {self.v_max_cont} pu Cat-B)"

            results.append(BusResult(
                bus_number     = bnum,
                bus_name       = bus_names[i].strip(),
                base_kv        = base_kvs[i],
                voltage_pu     = round(v_pu, 5),
                angle_deg      = round(angles[i], 4),
                voltage_kv     = round(v_kv, 3),
                bus_type       = bus_types[i],
                violation      = violation,
                violation_type = vtype,
            ))

        return results

    # ── Private: plots ────────────────────────────────────────────────────────

    def _plot_voltage_profile(self, plots_dir: str, label: str) -> str:
        """Bar chart of bus voltage magnitudes with AESO limit lines."""
        buses = self.results.buses
        df = pd.DataFrame([{
            "Bus Name":     b.bus_name,
            "Voltage (pu)": b.voltage_pu,
            "Violation":    b.violation,
        } for b in buses]).sort_values("Voltage (pu)")

        colors = [_RED if v else _GREEN for v in df["Violation"]]
        fig, ax = plt.subplots(figsize=(max(10, len(df) * 0.35), 6))
        ax.bar(range(len(df)), df["Voltage (pu)"], color=colors, width=0.7, zorder=3)
        ax.axhline(self.v_min, color=_RED,    linestyle="--", linewidth=1.5)
        ax.axhline(self.v_max, color=_ORANGE, linestyle="--", linewidth=1.5)
        ax.axhline(1.00,       color=_BLUE,   linestyle=":",  linewidth=1.0, alpha=0.5)
        ax.set_xticks(range(len(df)))
        ax.set_xticklabels(df["Bus Name"], rotation=45, ha="right", fontsize=8)
        ax.set_ylabel("Voltage (pu)")
        ax.set_title(f"Bus Voltage Profile — {self.scenario}")
        ax.set_ylim(
            min(0.85, df["Voltage (pu)"].min() - 0.02),
            max(1.15, df["Voltage (pu)"].max() + 0.02),
        )
        ax.grid(axis="y", zorder=0)
        ax.legend(handles=[
            mpatches.Patch(color=_GREEN,  label="Within Limits"),
            mpatches.Patch(color=_RED,    label="Violation"),
            mpatches.Patch(color=_RED,    label=f"Min {self.v_min} pu"),
            mpatches.Patch(color=_ORANGE, label=f"Max {self.v_max} pu"),
        ], loc="lower right", fontsize=8)
        fig.tight_layout()
        path = os.path.join(plots_dir, f"{label}_voltage_profile.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    def _plot_thermal_loading(self, plots_dir: str, label: str, top_n: int = 30) -> str:
        """Horizontal bar chart of branch thermal loading, top N most loaded."""
        branches = self.results.branches
        df = pd.DataFrame([{
            "Label":             f"{b.from_bus}-{b.to_bus} [{b.circuit_id}]",
            "Loading %":         b.loading_pct,
            "Thermal Violation": b.violation,
        } for b in branches]).sort_values("Loading %", ascending=False).head(top_n)

        colors = [_RED if v else _GREEN for v in df["Thermal Violation"]]
        fig, ax = plt.subplots(figsize=(10, max(6, len(df) * 0.35)))
        ax.barh(range(len(df)), df["Loading %"], color=colors, height=0.7, zorder=3)
        ax.axvline(self.thermal_limit, color=_RED, linestyle="--", linewidth=1.5,
                   label=f"Limit = {self.thermal_limit}%")
        ax.set_yticks(range(len(df)))
        ax.set_yticklabels(df["Label"], fontsize=8)
        ax.invert_yaxis()
        ax.set_xlabel("Thermal Loading (%)")
        ax.set_title(f"Branch Thermal Loading (Top {top_n}) — {self.scenario}")
        ax.grid(axis="x", zorder=0)
        ax.legend(fontsize=9)
        fig.tight_layout()
        path = os.path.join(plots_dir, f"{label}_thermal_loading.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    def _plot_contingency_summary(self, plots_dir: str, label: str) -> str:
        """Scatter: N-1 post-contingency min voltage vs. total violations."""
        import numpy as np
        conts = self.results.contingencies
        df = pd.DataFrame([{
            "Min Voltage (pu)": c.min_voltage_pu,
            "Bus Violations":   len(c.bus_violations),
            "Branch Violations":len(c.branch_violations),
        } for c in conts])
        df["Total Violations"] = df["Bus Violations"] + df["Branch Violations"]

        fig, ax = plt.subplots(figsize=(10, 6))
        sc = ax.scatter(
            df["Min Voltage (pu)"], df["Total Violations"],
            c=df["Total Violations"], cmap="RdYlGn_r",
            s=60, alpha=0.8, zorder=3,
        )
        plt.colorbar(sc, ax=ax, label="Total Violations")
        ax.axvline(self.v_min_cont, color=_RED, linestyle="--", linewidth=1.5,
                   label=f"Cat B Min V = {self.v_min_cont} pu")
        ax.set_xlabel("Post-Contingency Min Bus Voltage (pu)")
        ax.set_ylabel("Total Violations (Bus + Branch)")
        ax.set_title(f"N-1 Contingency Overview — {self.scenario}")
        ax.grid(zorder=0)
        ax.legend(fontsize=9)
        fig.tight_layout()
        path = os.path.join(plots_dir, f"{label}_n1_contingency.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    # ── Private: PDF export ───────────────────────────────────────────────────

    def _export_pdf(self, pdf_path: str, plot_paths: List[str]) -> None:
        """Assemble a PDF report from plots + summary table."""
        with PdfPages(pdf_path) as pdf:

            # Page 1 … N: embed each PNG plot as a full page
            for png in plot_paths:
                if png and os.path.isfile(png):
                    img = plt.imread(png)
                    fig, ax = plt.subplots(figsize=(14, 8))
                    ax.imshow(img)
                    ax.axis("off")
                    pdf.savefig(fig, bbox_inches="tight")
                    plt.close(fig)

            # Final page: summary table
            r = self.results
            summary_data = {
                "Parameter": [
                    "Scenario", "Converged",
                    "Total Generation (MW)", "Total Load (MW)", "Total Losses (MW)",
                    "Bus Violations (Cat A)", "Branch Violations",
                    "N-1 Contingencies Analysed",
                    "N-1 With Violations",
                ],
                "Value": [
                    r.scenario, str(r.converged),
                    f"{r.total_generation_mw:.2f}",
                    f"{r.total_load_mw:.2f}",
                    f"{r.total_losses_mw:.2f}",
                    len(r.bus_violations),
                    len(r.branch_violations),
                    len(r.contingencies),
                    sum(1 for c in r.contingencies
                        if c.bus_violations or c.branch_violations),
                ],
            }
            df_sum = pd.DataFrame(summary_data)
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.axis("off")
            tbl = ax.table(
                cellText=df_sum.values,
                colLabels=df_sum.columns,
                cellLoc="left", loc="center",
            )
            tbl.auto_set_font_size(False)
            tbl.set_fontsize(10)
            tbl.scale(1.2, 2.0)
            for col in range(len(df_sum.columns)):
                tbl[0, col].set_facecolor(_BLUE)
                tbl[0, col].set_text_props(color="white", weight="bold")
            ax.set_title(
                f"AESO Power Flow Study — Summary\n{self.scenario}",
                fontsize=12, weight="bold", pad=15,
            )
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

            # PDF metadata
            d = pdf.infodict()
            d["Title"]        = f"AESO Power Flow Study — {self.scenario}"
            d["Author"]       = "AESO Automation Tool"
            d["CreationDate"] = datetime.now()

    # ── Private: DataFrame builders ───────────────────────────────────────────

    def _summary_to_df(self) -> pd.DataFrame:
        r = self.results
        if r is None:
            return pd.DataFrame()
        return pd.DataFrame([
            {"Parameter": "Scenario",                  "Value": r.scenario},
            {"Parameter": "Converged",                 "Value": str(r.converged)},
            {"Parameter": "Total Generation (MW)",     "Value": r.total_generation_mw},
            {"Parameter": "Total Load (MW)",           "Value": r.total_load_mw},
            {"Parameter": "Total Losses (MW)",         "Value": r.total_losses_mw},
            {"Parameter": "Bus Violations (Cat A)",    "Value": len(r.bus_violations)},
            {"Parameter": "Branch Violations",         "Value": len(r.branch_violations)},
            {"Parameter": "N-1 Contingencies Run",     "Value": len(r.contingencies)},
            {"Parameter": "N-1 With Violations",       "Value": sum(
                1 for c in r.contingencies
                if c.bus_violations or c.branch_violations
            )},
            {"Parameter": "Cat A V_min (pu)",          "Value": self.v_min},
            {"Parameter": "Cat A V_max (pu)",          "Value": self.v_max},
            {"Parameter": "Cat B V_min (pu)",          "Value": self.v_min_cont},
            {"Parameter": "Cat B V_max (pu)",          "Value": self.v_max_cont},
            {"Parameter": "Thermal Limit (%)",         "Value": self.thermal_limit},
        ])

    def _buses_to_df(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "Bus Number":    b.bus_number,
            "Bus Name":      b.bus_name,
            "Base kV":       b.base_kv,
            "Bus Type":      b.bus_type,
            "Voltage (pu)":  b.voltage_pu,
            "Voltage (kV)":  b.voltage_kv,
            "Angle (deg)":   b.angle_deg,
            "Violation":     b.violation,
            "Violation Type":b.violation_type,
        } for b in (self.results.buses if self.results else [])])

    def _branches_to_df(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "From Bus":          b.from_bus,
            "To Bus":            b.to_bus,
            "Circuit ID":        b.circuit_id,
            "MW Flow":           b.mw_flow,
            "MVAR Flow":         b.mvar_flow,
            "MVA Flow":          b.mva_flow,
            "Rating MVA":        b.rating_mva,
            "Loading %":         b.loading_pct,
            "Thermal Violation": b.violation,
        } for b in (self.results.branches if self.results else [])])

    def _violations_to_df(self) -> pd.DataFrame:
        rows = []
        if self.results:
            for b in self.results.bus_violations:
                rows.append({
                    "Type":    "Voltage (Cat A)",
                    "Element": f"Bus {b.bus_number} ({b.bus_name})",
                    "Value":   b.voltage_pu,
                    "Detail":  b.violation_type,
                })
            for b in self.results.branch_violations:
                rows.append({
                    "Type":    "Thermal",
                    "Element": f"Branch {b.from_bus}-{b.to_bus} ckt {b.circuit_id}",
                    "Value":   b.loading_pct,
                    "Detail":  f"{b.loading_pct:.1f}% > {self.thermal_limit}%",
                })
        return pd.DataFrame(rows)

    def _contingency_summary_to_df(self) -> pd.DataFrame:
        rows = [{
            "Contingency":        c.contingency_name,
            "Converged":          c.converged,
            "Min Voltage (pu)":   c.min_voltage_pu,
            "Max Voltage (pu)":   c.max_voltage_pu,
            "Bus Violations":     len(c.bus_violations),
            "Branch Violations":  len(c.branch_violations),
            "AESO Status":        "VIOLATION" if (c.bus_violations or c.branch_violations) else "PASS",
        } for c in (self.results.contingencies if self.results else [])]
        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values("Bus Violations", ascending=False)
        return df

    # ── Mock data ─────────────────────────────────────────────────────────────

    @staticmethod
    def _mock_bus_results() -> List[BusResult]:
        return [
            BusResult(101, "BUS_CALGARY",      240.0, 1.020, -5.2,  244.8, 2),
            BusResult(102, "BUS_EDMONTON",     240.0, 0.998, -8.1,  239.5, 1),
            BusResult(103, "BUS_LETHBRIDGE",   138.0, 0.943, -12.3, 130.1, 1,
                      True, "LOW  (0.943 pu < 0.95 pu)"),
            BusResult(104, "BUS_RED_DEER",      69.0, 1.058,  2.1,   73.0, 1,
                      True, "HIGH (1.058 pu > 1.05 pu)"),
            BusResult(105, "BUS_MEDICINE_HAT", 138.0, 1.001, -7.8,  138.1, 1),
        ]

    @staticmethod
    def _mock_branch_results() -> List[BranchResult]:
        return [
            BranchResult(101, 102, "1", 450.2,  85.3, 458.2, 500.0,  91.6),
            BranchResult(102, 103, "1", 210.5,  40.1, 214.3, 200.0, 107.2, True),
            BranchResult(103, 104, "1",  90.8,  18.5,  92.7, 150.0,  61.8),
            BranchResult(104, 105, "1", 175.3,  32.2, 178.2, 200.0,  89.1),
        ]

    @staticmethod
    def _mock_contingency_results() -> List[ContingencyResult]:
        return [
            ContingencyResult(
                contingency_name  = "Branch_101_102_1",
                element_removed   = "Branch_101_102_1",
                converged         = True,
                bus_violations    = [
                    BusResult(103, "BUS_LETHBRIDGE", 138.0, 0.885, -14.0,
                              122.1, 1, True, "LOW  (0.885 pu < 0.90 pu Cat-B)")
                ],
                branch_violations = [],
                max_voltage_pu    = 1.052,
                min_voltage_pu    = 0.885,
            ),
        ]
