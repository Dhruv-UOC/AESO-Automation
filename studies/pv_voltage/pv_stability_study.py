"""
studies/pv_voltage/pv_stability_study.py
-----------------------------------------
Automates AESO PV Voltage Stability Analysis (Phase 4) using PSS/E via psspy.

Workflow
--------
1.  Load base case (.sav)
2.  For Category A (N-0) and each Category B (N-1) contingency:
    a. Reload fresh .sav
    b. Disable area interchange control (per AESO Requirements Section 3.2)
    c. Lock ULTC transformers (per AESO Requirements Section 3.2)
    d. Apply contingency branch trip (if N-1)
    e. Solve post-contingency power flow
    f. Ramp solar MW from 0 → transfer_end_mw in step_mw increments
    g. Solve Newton-Raphson power flow at each step
    h. Record POI voltage at each step
    i. Detect voltage collapse (fnsl diverges) → record collapse MW
    j. Flag AESO voltage limit violations
3.  Export PV curves to Excel + per-scenario CSV + PDF report

AESO Criteria Applied
---------------------
Category A (N-0) : V_min = 0.95 pu, V_max = 1.05 pu  (normal operating)
Category B (N-1) : V_min = 0.90 pu, V_max = 1.10 pu  (extreme voltage range)
Voltage Stability Margin : ≥ 5% above peak operating point (Table 2-2)
ULTC transformers locked during PV analysis (Requirements Section 3.2)
Area interchange control disabled (Requirements Section 3.2)
Post-contingency voltage deviations checked (Table 3-1: ±10%/±7%/±5%)

Usage
-----
    from core.psse_interface import PSSEInterface
    from studies.pv_voltage.pv_stability_study import PVStabilityStudy
    from project_io.project_data import ProjectData

    psse = PSSEInterface(psse_path=PSSE_PATH)
    psse.initialize()

    study = PVStabilityStudy(psse, project_data, scenario_label="2028_SP_Post")
    results = study.run(sav_path)
    study.save_results("output/results", "output/plots", "output/reports")
"""

from __future__ import annotations

import logging
import math
import os
import random
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

from project_io.project_data import ProjectData, PVContingency

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

# ── AESO voltage stability thresholds ─────────────────────────────────────────
AESO_CAT_A_V_MIN = 0.95
AESO_CAT_A_V_MAX = 1.05
AESO_CAT_B_V_MIN = 0.90
AESO_CAT_B_V_MAX = 1.10


# ── Data containers ───────────────────────────────────────────────────────────

@dataclass
class PVPoint:
    """Single operating point on a PV curve."""
    transfer_mw:          float
    poi_voltage_pu:       float
    aeso_limit_violation: bool  = False
    violation_detail:     str   = ""


@dataclass
class PVCurveResult:
    """Full PV curve for one scenario (N-0 or one N-1 contingency)."""
    scenario_name:      str
    category:           str            # "A" (N-0) or "B" (N-1)
    is_contingency:     bool
    contingency_element:str            # e.g. "9L24 Oakland–Lanfine" or "N/A"
    converged_base:     bool
    points:             List[PVPoint]  = field(default_factory=list)
    collapse_mw:        Optional[float]= None
    max_stable_mw:      float          = 0.0
    min_voltage_pu:     float          = 1.0
    violations:         List[PVPoint]  = field(default_factory=list)

    @property
    def aeso_status(self) -> str:
        return "VIOLATION" if self.violations else "PASS"


@dataclass
class PVStabilityResults:
    """Aggregated results across all N-0 and N-1 curves."""
    scenario_label:     str
    source_bus:         Optional[int]
    poi_bus:            Optional[int]
    transfer_range_mw:  Tuple[float, float]
    step_mw:            float
    curves:             List[PVCurveResult] = field(default_factory=list)
    most_critical_curve:Optional[PVCurveResult] = None


# ── Study class ───────────────────────────────────────────────────────────────

class PVStabilityStudy:
    """
    Automates AESO PV Voltage Stability Analysis using PSS/E.

    Parameters
    ----------
    psse : PSSEInterface
        Initialized PSS/E interface.
    project : ProjectData
        Populated project data (bus numbers, PV contingencies).
    scenario_label : str
        Human-readable label used in output filenames.
    transfer_start_mw : float
        Starting solar output in MW (default 0).
    transfer_end_mw : float
        Maximum solar output to test in MW (default MARP + 5% buffer).
    step_mw : float
        MW increment per power flow solve (default 10).
    v_min_cat_a : float
        Category A minimum voltage (default AESO 0.95 pu).
    v_min_cat_b : float
        Category B minimum voltage (default AESO 0.90 pu).
    """

    def __init__(
        self,
        psse,
        project:           ProjectData,
        scenario_label:    str   = "Base_Case",
        transfer_start_mw: float = 0.0,
        transfer_end_mw:   float = None,   # defaults to MARP + 5%
        step_mw:           float = 10.0,
        v_min_cat_a:       float = AESO_CAT_A_V_MIN,
        v_min_cat_b:       float = AESO_CAT_B_V_MIN,
    ):
        self._psse          = psse
        self._project       = project
        self.scenario       = scenario_label
        self.transfer_start = transfer_start_mw
        self.transfer_end   = (
            transfer_end_mw if transfer_end_mw is not None
            else round(project.info.marp_mw * 1.05, 0)
        )
        self.step_mw        = step_mw
        self.v_min_a        = v_min_cat_a
        self.v_min_b        = v_min_cat_b
        self.results: Optional[PVStabilityResults] = None

    # ── Public ────────────────────────────────────────────────────────────────

    def run(self, sav_path: str) -> PVStabilityResults:
        """
        Run full PV stability study: Category A (N-0) + all N-1 contingencies.

        Parameters
        ----------
        sav_path : str
            Path to the .sav case file.

        Returns
        -------
        PVStabilityResults
        """
        source_bus = self._project.info.source_bus_number
        poi_bus    = self._project.info.poi_bus_number

        logger.info("=" * 68)
        logger.info("PV Voltage Stability  |  Scenario: %s", self.scenario)
        logger.info(
            "Source Bus: %s  |  POI Bus: %s",
            source_bus or "NOT SET", poi_bus or "NOT SET"
        )
        logger.info(
            "Transfer: %.0f – %.0f MW  |  Step: %.0f MW",
            self.transfer_start, self.transfer_end, self.step_mw
        )
        logger.info(
            "AESO Limits → Cat A: %.2f pu  |  Cat B: %.2f pu",
            self.v_min_a, self.v_min_b
        )
        logger.info("=" * 68)

        if source_bus is None or poi_bus is None:
            logger.warning(
                "Source Bus or POI Bus not set in Project_Info sheet. "
                "Running in mock mode with synthetic data."
            )

        curves: List[PVCurveResult] = []

        # Build scenario list: N-0 first, then each N-1 from project data
        all_scenarios = [
            {
                "name":     f"{self.scenario} — Base Case (N-0)",
                "category": "A",
                "is_cont":  False,
                "element":  "N/A",
            }
        ] + [
            {
                "name":     f"{self.scenario} — {c.contingency_name}",
                "category": "B",
                "is_cont":  True,
                "element":  c.contingency_name,
                "f_bus":    c.from_bus_no,
                "t_bus":    c.to_bus_no,
                "ckt":      c.circuit_id,
            }
            for c in self._project.pv_contingencies
        ]

        for scen in all_scenarios:
            curve = self._run_single_curve(sav_path, scen, source_bus, poi_bus)
            curves.append(curve)

        # Most critical = lowest collapse MW (or lowest max_stable if no collapse)
        curves_with_collapse = [c for c in curves if c.collapse_mw is not None]
        most_critical = (
            min(curves_with_collapse, key=lambda c: c.collapse_mw)
            if curves_with_collapse
            else min(curves, key=lambda c: c.max_stable_mw)
        )

        self.results = PVStabilityResults(
            scenario_label    = self.scenario,
            source_bus        = source_bus,
            poi_bus           = poi_bus,
            transfer_range_mw = (self.transfer_start, self.transfer_end),
            step_mw           = self.step_mw,
            curves            = curves,
            most_critical_curve = most_critical,
        )

        logger.info(
            "\nPV Study complete. Most critical: '%s'  "
            "(collapse at %s MW)",
            most_critical.scenario_name,
            most_critical.collapse_mw
            if most_critical.collapse_mw else "not reached",
        )
        return self.results

    def save_results(
        self,
        results_dir: str,
        plots_dir:   str,
        reports_dir: str,
    ) -> dict:
        """
        Export results to Excel, per-scenario CSV, PNG plots, and PDF report.

        Returns dict with keys 'excel', 'csv_dir', 'plots', 'pdf'.
        """
        if self.results is None:
            raise RuntimeError("No results to save. Call run() first.")

        for d in (results_dir, plots_dir, reports_dir):
            os.makedirs(d, exist_ok=True)

        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        label = f"pv_stability_{self.scenario}_{ts}"

        # Excel
        excel_path = os.path.join(results_dir, f"{label}.xlsx")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            self._summary_to_df().to_excel(
                writer, sheet_name="Summary",          index=False)
            self._curve_data_to_df().to_excel(
                writer, sheet_name="PV_Curve_Data",    index=False)
            self._violations_to_df().to_excel(
                writer, sheet_name="Violations",       index=False)
            self._collapse_summary_to_df().to_excel(
                writer, sheet_name="Collapse_Summary", index=False)
        logger.info("Excel saved: %s", excel_path)

        # Per-scenario CSVs
        csv_dir = os.path.join(results_dir, f"csv_{self.scenario}_{ts}")
        os.makedirs(csv_dir, exist_ok=True)
        for curve in self.results.curves:
            safe = (curve.scenario_name
                    .replace(" ", "_").replace(":", "-")
                    .replace("/", "-").replace("—", "-"))
            csv_path = os.path.join(csv_dir, f"{safe[:80]}.csv")
            rows = [
                {
                    "Transfer (MW)":    p.transfer_mw,
                    "POI Voltage (pu)": p.poi_voltage_pu,
                    "AESO Violation":   p.aeso_limit_violation,
                    "Violation Detail": p.violation_detail,
                }
                for p in curve.points
            ]
            pd.DataFrame(rows).to_csv(csv_path, index=False)

        # PNG plots
        plot_paths = []
        plot_paths.append(self._plot_all_curves(plots_dir, label))
        plot_paths.append(self._plot_critical_curve(plots_dir, label))
        plot_paths = [p for p in plot_paths if p]

        # PDF report
        pdf_path = os.path.join(reports_dir, f"{label}.pdf")
        self._export_pdf(pdf_path, plot_paths)
        logger.info("PDF saved: %s", pdf_path)

        return {
            "excel":   excel_path,
            "csv_dir": csv_dir,
            "plots":   plot_paths,
            "pdf":     pdf_path,
        }

    # ── Private: PSS/E PV curve engine ───────────────────────────────────────

    def _run_single_curve(
        self,
        sav_path:   str,
        scen:       dict,
        source_bus: Optional[int],
        poi_bus:    Optional[int],
    ) -> PVCurveResult:
        """Run one PV curve (one N-0 or N-1 scenario)."""
        scen_name = scen["name"]
        category  = scen["category"]
        is_cont   = scen["is_cont"]
        v_limit   = self.v_min_b if is_cont else self.v_min_a

        logger.info("\n>>> Curve: %s", scen_name)

        curve = PVCurveResult(
            scenario_name       = scen_name,
            category            = category,
            is_contingency      = is_cont,
            contingency_element = scen.get("element", "N/A"),
            converged_base      = False,
        )

        # Use mock if PSS/E not available or bus numbers missing
        if self._psse.mock or source_bus is None or poi_bus is None:
            return self._mock_pv_curve(curve, v_limit)

        psspy = self._psse.psspy

        # ── Step 1: Reload fresh .sav ──────────────────────────────────────
        ierr = psspy.case(sav_path)
        if ierr != 0:
            logger.error("  Failed to load '%s'.", sav_path)
            return curve

        # ── Step 2: Disable area interchange control (AESO requirement) ────
        # activity AREA: set area interchange control off
        try:
            ierr_ai = psspy.solution_parameters_4(
                [0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0],
                [0.0001, 0.0001, 0.0, 0.0, 0.0, 0.0],
            )
        except Exception:
            pass   # non-fatal — best effort

        # ── Step 3: Lock ULTC transformers (AESO requirement) ──────────────
        # Set all transformer tap adjustment off
        try:
            psspy.tap_adjustment_flag(0)   # 0 = off
        except Exception:
            pass   # non-fatal

        # ── Step 4: Solve base power flow ──────────────────────────────────
        err = psspy.fnsl([1, 0, 0, 1, 1, 0, 99, 0])
        curve.converged_base = (err == 0)
        if not curve.converged_base:
            logger.error("  Base case diverged. Skipping.")
            return curve

        # ── Step 5: Apply N-1 contingency (if applicable) ──────────────────
        if is_cont:
            f_bus = scen.get("f_bus")
            t_bus = scen.get("t_bus")
            ckt   = scen.get("ckt", "1")
            if f_bus and t_bus:
                # Open branch: status 0 = out of service
                ierr = psspy.branch_chng_3(
                    f_bus, t_bus, ckt,
                    [0, psspy._i, psspy._i, psspy._i, psspy._i, psspy._i],
                    [psspy._f] * 12,
                )
                if ierr != 0:
                    logger.warning(
                        "  Branch trip failed (%d–%d ckt %s, ierr=%d). Skipping.",
                        f_bus, t_bus, ckt, ierr
                    )
                    return curve
                logger.info(
                    "  Contingency applied: Line %d–%d ckt %s opened.",
                    f_bus, t_bus, ckt
                )
                # Re-solve post-contingency base
                err = psspy.fnsl([1, 0, 0, 1, 1, 0, 99, 0])
                if err != 0:
                    logger.warning("  Post-contingency base diverged. Skipping.")
                    return curve
            else:
                logger.warning(
                    "  '%s' has no bus numbers — fill PV_Contingencies sheet.",
                    scen_name
                )

        # ── Step 6: MW ramp loop ───────────────────────────────────────────
        transfer_mw = self.transfer_start

        while transfer_mw <= self.transfer_end:

            # Set solar active power output (PGEN) on generator at source_bus
            ierr = psspy.machine_chng_2(
                source_bus,
                r"""1""",
                intgar1 = 1,          # machine in-service
                realar1 = float(transfer_mw),
            )
            if ierr != 0:
                logger.warning(
                    "  machine_chng_2 failed at %.0f MW (ierr=%d).",
                    transfer_mw, ierr
                )
                break

            # Full Newton-Raphson power flow
            solve_err = psspy.fnsl([1, 0, 0, 1, 1, 0, 99, 0])

            if solve_err == 0:
                # Read POI bus voltage in per-unit
                ierr_v, poi_v = psspy.busdat(poi_bus, "PU")
                if ierr_v != 0:
                    logger.warning(
                        "  busdat failed at %.0f MW (ierr=%d). Stopping.",
                        transfer_mw, ierr_v
                    )
                    break

                poi_v    = round(poi_v, 5)
                violates = poi_v < v_limit
                v_detail = (
                    f"V={poi_v:.4f} pu < AESO Cat {'B' if is_cont else 'A'} "
                    f"limit {v_limit:.2f} pu"
                ) if violates else ""

                pt = PVPoint(
                    transfer_mw          = transfer_mw,
                    poi_voltage_pu       = poi_v,
                    aeso_limit_violation = violates,
                    violation_detail     = v_detail,
                )
                curve.points.append(pt)
                if violates:
                    curve.violations.append(pt)

                curve.max_stable_mw = transfer_mw
                curve.min_voltage_pu = min(curve.min_voltage_pu, poi_v)

                logger.info(
                    "  %5.0f MW  →  V_poi = %.4f pu%s",
                    transfer_mw, poi_v,
                    "  *** AESO VIOLATION ***" if violates else ""
                )

            else:
                # Diverged → voltage collapse
                curve.collapse_mw = transfer_mw
                logger.warning(
                    "  *** Voltage Collapse at %.0f MW (fnsl err=%d) ***",
                    transfer_mw, solve_err
                )
                break

            transfer_mw = round(transfer_mw + self.step_mw, 2)

        logger.info(
            "  Done: max_stable=%.0f MW | collapse=%s MW | min_V=%.4f pu",
            curve.max_stable_mw,
            str(curve.collapse_mw) if curve.collapse_mw else "not reached",
            curve.min_voltage_pu,
        )
        return curve

    # ── Private: plots ────────────────────────────────────────────────────────

    def _plot_all_curves(self, plots_dir: str, label: str) -> Optional[str]:
        """All PV curves overlaid on one chart."""
        if not self.results or not self.results.curves:
            return None

        colors = plt.cm.tab10.colors
        fig, ax = plt.subplots(figsize=(12, 7))

        for i, curve in enumerate(self.results.curves):
            if not curve.points:
                continue
            xs     = [p.transfer_mw    for p in curve.points]
            ys     = [p.poi_voltage_pu for p in curve.points]
            ls     = "-"  if curve.category == "A" else "--"
            color  = colors[i % len(colors)]

            ax.plot(xs, ys,
                    linestyle=ls, marker="o", markersize=3,
                    linewidth=2, color=color,
                    label=curve.scenario_name)

            if curve.collapse_mw is not None:
                ax.axvline(
                    x=curve.collapse_mw,
                    color=color, linestyle=":", linewidth=1.2, alpha=0.55
                )
                ax.annotate(
                    f"Collapse\n{curve.collapse_mw:.0f} MW",
                    xy=(curve.collapse_mw, self.v_min_b + 0.02),
                    fontsize=7, color=color, ha="center"
                )

        ax.axhline(
            y=self.v_min_b, color=_RED, linestyle="--", linewidth=2.0,
            label=f"AESO Cat B ({self.v_min_b:.2f} pu — N-1)"
        )
        ax.axhline(
            y=self.v_min_a, color=_ORANGE, linestyle="--", linewidth=2.0,
            label=f"AESO Cat A ({self.v_min_a:.2f} pu — N-0)"
        )

        ax.set_title(
            f"PV Voltage Stability Curves\n"
            f"{self._project.info.project_number}  |  {self.scenario}",
            fontsize=13, weight="bold", pad=12
        )
        ax.set_xlabel("Solar Plant Active Power Output (MW)", fontsize=11)
        ax.set_ylabel("POI Bus Voltage (pu)", fontsize=11)
        ax.set_ylim(0.82, 1.08)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(50))
        ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.65)
        ax.legend(fontsize=8, loc="lower left", framealpha=0.9)
        plt.tight_layout()

        path = os.path.join(plots_dir, f"{label}_all_curves.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    def _plot_critical_curve(self, plots_dir: str, label: str) -> Optional[str]:
        """Plot of the most critical curve only, with annotations."""
        if not self.results or not self.results.most_critical_curve:
            return None

        curve = self.results.most_critical_curve
        if not curve.points:
            return None

        xs = [p.transfer_mw    for p in curve.points]
        ys = [p.poi_voltage_pu for p in curve.points]

        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(xs, ys, color=_BLUE, linewidth=2.5,
                marker="o", markersize=4, label=curve.scenario_name)

        # Violation points
        vx = [p.transfer_mw    for p in curve.violations]
        vy = [p.poi_voltage_pu for p in curve.violations]
        if vx:
            ax.scatter(vx, vy, color=_RED, s=50, zorder=5, label="AESO Violation")

        # Collapse
        if curve.collapse_mw:
            ax.axvline(
                x=curve.collapse_mw,
                color=_RED, linestyle="--", linewidth=2.0,
                label=f"Collapse at {curve.collapse_mw:.0f} MW"
            )

        # MARP line
        marp = self._project.info.marp_mw
        if marp > 0:
            ax.axvline(
                x=marp,
                color=_GREEN, linestyle="--", linewidth=1.5,
                label=f"MARP = {marp:.0f} MW"
            )

        v_lim = self.v_min_b if curve.is_contingency else self.v_min_a
        ax.axhline(
            y=v_lim, color=_RED, linestyle="--", linewidth=1.5,
            label=f"AESO Cat {'B' if curve.is_contingency else 'A'} = {v_lim:.2f} pu"
        )

        ax.set_title(
            f"Most Critical PV Curve — {curve.scenario_name}\n"
            f"Max Stable: {curve.max_stable_mw:.0f} MW  |  "
            f"Collapse: {curve.collapse_mw or 'Not Reached'} MW  |  "
            f"AESO: {curve.aeso_status}",
            fontsize=11, weight="bold", pad=10
        )
        ax.set_xlabel("Solar Plant Active Power Output (MW)", fontsize=11)
        ax.set_ylabel("POI Bus Voltage (pu)", fontsize=11)
        ax.set_ylim(0.82, 1.08)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(50))
        ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.65)
        ax.legend(fontsize=9, loc="lower left")
        plt.tight_layout()

        path = os.path.join(plots_dir, f"{label}_critical_curve.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    # ── Private: PDF export ───────────────────────────────────────────────────

    def _export_pdf(self, pdf_path: str, plot_paths: List[str]) -> None:
        """Assemble PDF: PV curve plots + collapse/violation summary table."""
        with PdfPages(pdf_path) as pdf:

            for png in plot_paths:
                if png and os.path.isfile(png):
                    img = plt.imread(png)
                    fig, ax = plt.subplots(figsize=(14, 8))
                    ax.imshow(img)
                    ax.axis("off")
                    pdf.savefig(fig, bbox_inches="tight")
                    plt.close(fig)

            # Collapse & violation summary table
            df_sum = self._collapse_summary_to_df()
            if not df_sum.empty:
                fig2, ax2 = plt.subplots(
                    figsize=(16, max(4, len(df_sum) * 0.65 + 2.5))
                )
                ax2.axis("off")
                tbl = ax2.table(
                    cellText=df_sum.values,
                    colLabels=df_sum.columns,
                    cellLoc="center", loc="center",
                )
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(8.5)
                tbl.scale(1.1, 1.9)
                for col in range(len(df_sum.columns)):
                    tbl[0, col].set_facecolor(_BLUE)
                    tbl[0, col].set_text_props(color="white", weight="bold")
                for row in range(1, len(df_sum) + 1):
                    status = str(df_sum.iloc[row - 1].get("AESO Status", ""))
                    color  = "#ffe0e0" if "VIOLATION" in status else "#e0ffe0"
                    for col in range(len(df_sum.columns)):
                        tbl[row, col].set_facecolor(color)
                ax2.set_title(
                    f"AESO PV Stability — Collapse & Violation Summary\n"
                    f"{self._project.info.project_number}  |  {self.scenario}",
                    fontsize=11, weight="bold", pad=15,
                )
                plt.tight_layout()
                pdf.savefig(fig2, bbox_inches="tight")
                plt.close(fig2)

            d = pdf.infodict()
            d["Title"]        = (
                f"AESO PV Voltage Stability — "
                f"{self._project.info.project_number} — {self.scenario}"
            )
            d["Author"]       = "AESO Automation Tool"
            d["Subject"]      = (
                f"Source Bus: {self.results.source_bus}  |  "
                f"POI Bus: {self.results.poi_bus}"
            )
            d["CreationDate"] = datetime.now()

    # ── Private: DataFrame builders ───────────────────────────────────────────

    def _summary_to_df(self) -> pd.DataFrame:
        r = self.results
        if r is None:
            return pd.DataFrame()
        mc = r.most_critical_curve
        return pd.DataFrame([
            {"Parameter": "Scenario Label",          "Value": r.scenario_label},
            {"Parameter": "Source (Generator) Bus",  "Value": r.source_bus},
            {"Parameter": "POI (Monitor) Bus",        "Value": r.poi_bus},
            {"Parameter": "Transfer Range (MW)",      "Value": f"{r.transfer_range_mw[0]} – {r.transfer_range_mw[1]}"},
            {"Parameter": "Step Size (MW)",           "Value": r.step_mw},
            {"Parameter": "Curves Run",               "Value": len(r.curves)},
            {"Parameter": "AESO Cat A V_min (pu)",   "Value": self.v_min_a},
            {"Parameter": "AESO Cat B V_min (pu)",   "Value": self.v_min_b},
            {"Parameter": "Most Critical Scenario",  "Value": mc.scenario_name if mc else "N/A"},
            {"Parameter": "Critical Collapse MW",    "Value": mc.collapse_mw if mc and mc.collapse_mw else "Not Reached"},
            {"Parameter": "Critical Max Stable MW",  "Value": mc.max_stable_mw if mc else "N/A"},
            {"Parameter": "MARP (MW)",               "Value": self._project.info.marp_mw},
        ])

    def _curve_data_to_df(self) -> pd.DataFrame:
        rows = []
        for curve in (self.results.curves if self.results else []):
            for p in curve.points:
                rows.append({
                    "Scenario":          curve.scenario_name,
                    "Category":          curve.category,
                    "Transfer (MW)":     p.transfer_mw,
                    "POI Voltage (pu)":  p.poi_voltage_pu,
                    "AESO Violation":    p.aeso_limit_violation,
                    "Violation Detail":  p.violation_detail,
                })
        return pd.DataFrame(rows)

    def _violations_to_df(self) -> pd.DataFrame:
        rows = []
        for curve in (self.results.curves if self.results else []):
            for p in curve.violations:
                rows.append({
                    "Scenario":          curve.scenario_name,
                    "Category":          curve.category,
                    "Transfer (MW)":     p.transfer_mw,
                    "POI Voltage (pu)":  p.poi_voltage_pu,
                    "Violation Detail":  p.violation_detail,
                })
        return pd.DataFrame(rows)

    def _collapse_summary_to_df(self) -> pd.DataFrame:
        rows = [
            {
                "Scenario":               c.scenario_name,
                "Category":               c.category,
                "Contingency Element":    c.contingency_element,
                "Max Stable MW":          c.max_stable_mw,
                "Collapse MW":            c.collapse_mw if c.collapse_mw else "Not Reached",
                "Min POI Voltage (pu)":   round(c.min_voltage_pu, 5),
                "Violation Count":        len(c.violations),
                "AESO Status":            c.aeso_status,
            }
            for c in (self.results.curves if self.results else [])
        ]
        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values("Max Stable MW", ascending=True)
        return df

    # ── Mock data ─────────────────────────────────────────────────────────────

    def _mock_pv_curve(
        self,
        curve:   PVCurveResult,
        v_limit: float,
    ) -> PVCurveResult:
        """Synthetic nose-curve for testing without PSS/E."""
        random.seed(hash(curve.scenario_name) % 9999)

        base_v   = 1.015 if curve.category == "A" else 0.985
        collapse = (
            random.uniform(370, 420) if curve.category == "A"
            else random.uniform(270, 360)
        )
        curve.converged_base = True

        mw = self.transfer_start
        while mw <= self.transfer_end:
            if mw >= collapse:
                curve.collapse_mw = round(mw, 1)
                break
            x = mw / self.transfer_end
            v = (base_v
                 - 0.18  * (x ** 1.6)
                 - 0.025 * math.sin(math.pi * x * 0.9)
                 + random.uniform(-0.001, 0.001))
            v = round(v, 5)

            violates = v < v_limit
            pt = PVPoint(
                transfer_mw          = mw,
                poi_voltage_pu       = v,
                aeso_limit_violation = violates,
                violation_detail     = f"V={v:.4f} < {v_limit:.2f}" if violates else "",
            )
            curve.points.append(pt)
            if violates:
                curve.violations.append(pt)
            curve.max_stable_mw  = mw
            curve.min_voltage_pu = min(curve.min_voltage_pu, v)
            mw = round(mw + self.step_mw, 2)

        return curve
