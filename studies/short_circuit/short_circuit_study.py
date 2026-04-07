"""
studies/short_circuit/short_circuit_study.py
---------------------------------------------
Automates AESO Short Circuit Analysis (Phase 2) using PSS/E via psspy.

Workflow
--------
1.  Load power flow case (.sav)
2.  Run fault simulations: 3PH, LG, LL, LLG at every bus
3.  Extract fault currents (kA) and fault MVA
4.  Rank faults by severity (highest current first)
5.  Flag buses exceeding equipment withstand limit (63 kA default)
6.  Compare pre-project vs post-project delta (optional)
7.  Export results to Excel + PNG plots + PDF report

AESO Criteria Applied
---------------------
- Worst-case scenario assumed (all area generators online)
- Three-phase faults and single line-to-ground faults required as minimum
  (Requirements doc Section 3.3) — all four types run here
- Results reported in polar coordinates and physical values (kA, MVA)
- Equipment withstand limit: 63 kA

Usage
-----
    from core.psse_interface import PSSEInterface
    from studies.short_circuit.short_circuit_study import ShortCircuitStudy
    from config.settings import PSSE_PATH, PSSE_VERSION, DEFAULT_SAV, AESO

    psse = PSSEInterface(psse_path=PSSE_PATH, psse_version=PSSE_VERSION)
    psse.initialize()

    study = ShortCircuitStudy(psse, scenario_label="Sample_Case")
    study.run(DEFAULT_SAV)
    study.save_results("output/results", "output/plots", "output/reports")
"""

import os
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages

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


# ── Data containers ───────────────────────────────────────────────────────────

@dataclass
class FaultResult:
    bus_number:           int
    bus_name:             str
    base_kv:              float
    fault_type:           str    # "3PH" | "LG" | "LL" | "LLG"
    fault_current_ka:     float  # Magnitude in kA
    fault_current_ang:    float  # Angle in degrees (polar form)
    fault_mva:            float  # Three-phase fault MVA
    pre_fault_voltage_pu: float
    severity_rank:        int  = 0
    violation:            bool = False   # True if fault_current_ka > limit


@dataclass
class ShortCircuitResults:
    scenario:             str
    fault_types_run:      List[str]
    faults:               List[FaultResult]  = field(default_factory=list)
    violations:           List[FaultResult]  = field(default_factory=list)
    max_fault_current_ka: float = 0.0
    max_fault_bus:        str   = ""
    max_fault_type:       str   = ""


# ── Study class ───────────────────────────────────────────────────────────────

class ShortCircuitStudy:
    """
    Automates PSS/E short circuit analysis for one operating scenario.

    All four fault types are run: 3PH, LG, LL, LLG.
    Results are reported in physical values (kA, MVA) and polar coordinates.

    Parameters
    ----------
    psse : PSSEInterface
        Initialized PSS/E interface.
    scenario_label : str
        Label used in output file names.
    max_fault_current_ka : float
        Equipment withstand limit (kA). Buses exceeding this are flagged.
        AESO default: 63 kA.
    bus_filter : list of int, optional
        Study only these bus numbers. None = all buses.
    """

    FAULT_TYPE_MAP = {
        "3PH": "Three-Phase Balanced",
        "LG":  "Single Line-to-Ground",
        "LL":  "Line-to-Line",
        "LLG": "Double Line-to-Ground",
    }

    # All four fault types — order matches AESO reporting convention
    FAULT_TYPES = ["3PH", "LG", "LL", "LLG"]

    def __init__(
        self,
        psse,
        scenario_label:       str        = "Base_Case",
        max_fault_current_ka: float      = 63.0,
        bus_filter:           List[int]  = None,
    ):
        self._psse         = psse
        self.scenario      = scenario_label
        self.max_fault_ka  = max_fault_current_ka
        self.bus_filter    = bus_filter
        self.results: Optional[ShortCircuitResults] = None

    # ── Public ────────────────────────────────────────────────────────────────

    def run(self, sav_path: str) -> ShortCircuitResults:
        """
        Execute short circuit study for all four fault types.

        Parameters
        ----------
        sav_path : str
            Path to the .sav case file.

        Returns
        -------
        ShortCircuitResults
        """
        logger.info("=" * 68)
        logger.info("Short Circuit Study  |  Scenario: %s", self.scenario)
        logger.info("Fault types: %s", ", ".join(self.FAULT_TYPES))
        logger.info("Equipment limit: %.1f kA", self.max_fault_ka)
        logger.info("=" * 68)

        self._psse.load_case(sav_path)

        all_faults: List[FaultResult] = []

        for fault_type in self.FAULT_TYPES:
            logger.info(
                "  Running %s (%s) ...",
                fault_type, self.FAULT_TYPE_MAP[fault_type]
            )
            faults = self._run_fault_type(fault_type)
            all_faults.extend(faults)
            logger.info("  → %d buses faulted", len(faults))

        # Sort by fault current descending and assign severity rank
        all_faults.sort(key=lambda f: f.fault_current_ka, reverse=True)
        for rank, f in enumerate(all_faults, start=1):
            f.severity_rank = rank

        violations = [f for f in all_faults if f.violation]
        max_fault  = all_faults[0] if all_faults else None

        self.results = ShortCircuitResults(
            scenario             = self.scenario,
            fault_types_run      = self.FAULT_TYPES,
            faults               = all_faults,
            violations           = violations,
            max_fault_current_ka = max_fault.fault_current_ka if max_fault else 0.0,
            max_fault_bus        = max_fault.bus_name         if max_fault else "",
            max_fault_type       = max_fault.fault_type       if max_fault else "",
        )

        logger.info(
            "Short circuit complete: %d faults  %d violations  "
            "max=%.3f kA at %s (%s)",
            len(all_faults), len(violations),
            self.results.max_fault_current_ka,
            self.results.max_fault_bus,
            self.results.max_fault_type,
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
        label = f"short_circuit_{self.scenario}_{ts}"

        # Excel
        excel_path = os.path.join(results_dir, f"{label}.xlsx")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            self._summary_to_df().to_excel(
                writer, sheet_name="Summary",          index=False)
            self._faults_to_df().to_excel(
                writer, sheet_name="All_Faults",       index=False)
            self._violations_to_df().to_excel(
                writer, sheet_name="Violations",       index=False)
            self._severity_ranking_to_df().to_excel(
                writer, sheet_name="Top20_Severity",   index=False)
        logger.info("Excel saved: %s", excel_path)

        # PNG plots
        plot_paths = []
        plot_paths.append(self._plot_fault_currents(plots_dir, label))
        plot_paths.append(self._plot_fault_by_type(plots_dir, label))
        if self.results.violations:
            plot_paths.append(self._plot_violations(plots_dir, label))

        # PDF report
        pdf_path = os.path.join(reports_dir, f"{label}.pdf")
        self._export_pdf(pdf_path, plot_paths)
        logger.info("PDF saved: %s", pdf_path)

        return {"excel": excel_path, "plots": plot_paths, "pdf": pdf_path}

    @staticmethod
    def compare(
        pre:  "ShortCircuitResults",
        post: "ShortCircuitResults",
    ) -> pd.DataFrame:
        """
        Generate a pre-project vs. post-project comparison DataFrame.

        Parameters
        ----------
        pre, post : ShortCircuitResults
            Results from two separate run() calls on different .sav files.

        Returns
        -------
        pd.DataFrame sorted by delta (highest increase first)
        """
        pre_dict  = {(f.bus_number, f.fault_type): f for f in pre.faults}
        post_dict = {(f.bus_number, f.fault_type): f for f in post.faults}

        rows = []
        for key, pf in pre_dict.items():
            postf = post_dict.get(key)
            if postf:
                delta    = postf.fault_current_ka - pf.fault_current_ka
                delta_pct = (delta / pf.fault_current_ka * 100.0
                             if pf.fault_current_ka else 0.0)
                rows.append({
                    "Bus Number":        pf.bus_number,
                    "Bus Name":          pf.bus_name,
                    "Base kV":           pf.base_kv,
                    "Fault Type":        pf.fault_type,
                    "Pre-Project (kA)":  round(pf.fault_current_ka,   4),
                    "Post-Project (kA)": round(postf.fault_current_ka, 4),
                    "Delta (kA)":        round(delta,     4),
                    "Delta (%)":         round(delta_pct, 2),
                    "Post Violation":    postf.violation,
                })

        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values("Delta (kA)", ascending=False)
        return df

    # ── Private: PSS/E fault calls ────────────────────────────────────────────

    def _run_fault_type(self, fault_type: str) -> List[FaultResult]:
        """Run one fault type at all buses (or bus_filter subset)."""
        if self._psse.mock:
            return self._mock_fault_results(fault_type)

        psspy = self._psse.psspy
        results = []

        # Collect bus data
        ierr, (bus_nums,)  = psspy.abusint(-1,  1, ["NUMBER"])
        ierr, (bus_names,) = psspy.abuschar(-1, 1, ["NAME"])
        ierr, (base_kvs,)  = psspy.abusreal(-1, 1, ["BASE"])
        ierr, (voltages,)  = psspy.abusreal(-1, 1, ["PU"])

        for i, bnum in enumerate(bus_nums):
            if self.bus_filter and bnum not in self.bus_filter:
                continue

            fault_ka, fault_ang = self._apply_fault(psspy, bnum, fault_type)
            fault_mva           = fault_ka * base_kvs[i] * 1.7321   # √3 × kV × kA

            results.append(FaultResult(
                bus_number           = bnum,
                bus_name             = bus_names[i].strip(),
                base_kv              = base_kvs[i],
                fault_type           = fault_type,
                fault_current_ka     = round(fault_ka,  4),
                fault_current_ang    = round(fault_ang, 2),
                fault_mva            = round(fault_mva, 2),
                pre_fault_voltage_pu = round(voltages[i], 5),
                violation            = fault_ka > self.max_fault_ka,
            ))

        return results

    @staticmethod
    def _apply_fault(
        psspy,
        bus_num:    int,
        fault_type: str,
    ) -> tuple:
        """
        Apply a fault at a bus and return (magnitude_kA, angle_deg).

        PSS/E short circuit API returns fault currents as complex numbers
        in per-unit on the system MVA base.  We take the magnitude and
        convert to kA using the bus base kV.

        Note: PSS/E returns currents in kA directly from sc3ph / sc1ph /
        sc2ph / sc2ph1 when called after psseinit — no per-unit conversion
        needed; the value returned IS in kA.

        Returns (0.0, 0.0) if the fault calculation fails.
        """
        import cmath

        # Common fault options arrays (all defaults)
        # intgar: [flt_type, bus_tie, phase, seq, output, ...]
        int_opts   = [0, 0, 0, 0, 0]
        float_opts = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

        try:
            if fault_type == "3PH":
                # Three-phase balanced fault
                # sc3ph returns (ierr, (Ia, Ib, Ic)) — complex kA values
                ierr, (ia, ib, ic) = psspy.sc3ph(
                    bus_num, 0, int_opts, float_opts
                )
                if ierr == 0 and ia is not None:
                    mag = abs(ia)
                    ang = cmath.phase(ia) * 180.0 / cmath.pi
                    return (mag, ang)

            elif fault_type == "LG":
                # Single line-to-ground fault (phase A to ground)
                # sc1ph returns (ierr, (Ia,)) — complex kA
                ierr, (ia,) = psspy.sc1ph(
                    bus_num, 0, 1, int_opts, float_opts
                )
                if ierr == 0 and ia is not None:
                    mag = abs(ia)
                    ang = cmath.phase(ia) * 180.0 / cmath.pi
                    return (mag, ang)

            elif fault_type == "LL":
                # Line-to-line fault (phases A and B)
                # sc2ph returns (ierr, (Ia, Ib)) — complex kA
                ierr, (ia, ib) = psspy.sc2ph(
                    bus_num, 0, [1, 2], int_opts, float_opts
                )
                if ierr == 0 and ia is not None:
                    mag = abs(ia)
                    ang = cmath.phase(ia) * 180.0 / cmath.pi
                    return (mag, ang)

            elif fault_type == "LLG":
                # Double line-to-ground fault (phases A and B to ground)
                # sc2ph1 returns (ierr, (Ia, Ib, Ic)) — complex kA
                ierr, (ia, ib, ic) = psspy.sc2ph1(
                    bus_num, 0, [1, 2], int_opts, float_opts
                )
                if ierr == 0 and ia is not None:
                    mag = abs(ia)
                    ang = cmath.phase(ia) * 180.0 / cmath.pi
                    return (mag, ang)

        except Exception as exc:
            logger.debug(
                "Fault calculation failed: bus=%d type=%s  reason: %s",
                bus_num, fault_type, exc
            )

        return (0.0, 0.0)

    # ── Private: plots ────────────────────────────────────────────────────────

    def _plot_fault_currents(
        self,
        plots_dir: str,
        label:     str,
        top_n:     int = 20,
    ) -> str:
        """Grouped bar chart: fault current by bus and fault type (top N)."""
        df = self._faults_to_df()
        if df.empty:
            return ""

        top_buses = (
            df.groupby("Bus Name")["Fault Current (kA)"]
            .max()
            .nlargest(top_n)
            .index.tolist()
        )
        df_top = df[df["Bus Name"].isin(top_buses)]

        fault_types = self.FAULT_TYPES
        x           = np.arange(len(top_buses))
        width       = 0.8 / len(fault_types)
        colors      = [_BLUE, _RED, _ORANGE, _GREEN]

        fig, ax = plt.subplots(figsize=(14, 6))
        for j, ft in enumerate(fault_types):
            sub  = df_top[df_top["Fault Type"] == ft].set_index("Bus Name")
            vals = [
                sub.loc[b, "Fault Current (kA)"] if b in sub.index else 0.0
                for b in top_buses
            ]
            ax.bar(
                x + j * width, vals,
                width=width * 0.9,
                label=f"{ft} — {self.FAULT_TYPE_MAP[ft]}",
                color=colors[j % len(colors)],
                zorder=3,
            )

        ax.axhline(
            self.max_fault_ka,
            color=_RED, linestyle="--", linewidth=1.8,
            label=f"Equipment Limit = {self.max_fault_ka} kA",
        )
        ax.set_xticks(x + width * (len(fault_types) - 1) / 2)
        ax.set_xticklabels(top_buses, rotation=45, ha="right", fontsize=8)
        ax.set_ylabel("Fault Current (kA)")
        ax.set_title(
            f"Short Circuit Fault Currents (Top {top_n} Buses) — {self.scenario}"
        )
        ax.grid(axis="y", zorder=0)
        ax.legend(fontsize=8)
        fig.tight_layout()

        path = os.path.join(plots_dir, f"{label}_fault_currents.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    def _plot_fault_by_type(self, plots_dir: str, label: str) -> str:
        """Box plot: fault current distribution per fault type."""
        df = self._faults_to_df()
        if df.empty:
            return ""

        fig, ax = plt.subplots(figsize=(9, 6))
        data   = [
            df[df["Fault Type"] == ft]["Fault Current (kA)"].dropna().values
            for ft in self.FAULT_TYPES
        ]
        bp = ax.boxplot(
            data,
            labels=[f"{ft}\n({self.FAULT_TYPE_MAP[ft]})" for ft in self.FAULT_TYPES],
            patch_artist=True,
            medianprops=dict(color="white", linewidth=2),
        )
        colors = [_BLUE, _RED, _ORANGE, _GREEN]
        for patch, color in zip(bp["boxes"], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.75)

        ax.axhline(
            self.max_fault_ka,
            color=_RED, linestyle="--", linewidth=1.8,
            label=f"Equipment Limit = {self.max_fault_ka} kA",
        )
        ax.set_ylabel("Fault Current (kA)")
        ax.set_title(f"Fault Current Distribution by Type — {self.scenario}")
        ax.grid(axis="y", zorder=0)
        ax.legend(fontsize=9)
        fig.tight_layout()

        path = os.path.join(plots_dir, f"{label}_fault_by_type.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    def _plot_violations(self, plots_dir: str, label: str) -> str:
        """Horizontal bar chart of buses that exceed the equipment limit."""
        df = self._violations_to_df()
        if df.empty:
            return ""

        fig, ax = plt.subplots(figsize=(10, max(4, len(df) * 0.4)))
        bars = ax.barh(
            range(len(df)),
            df["Fault Current (kA)"],
            color=_RED, height=0.6, zorder=3,
        )
        ax.axvline(
            self.max_fault_ka,
            color=_ORANGE, linestyle="--", linewidth=1.8,
            label=f"Limit = {self.max_fault_ka} kA",
        )
        ylabels = df.apply(
            lambda r: f"Bus {r['Bus Number']} {r['Bus Name']} ({r['Fault Type']})",
            axis=1,
        )
        ax.set_yticks(range(len(df)))
        ax.set_yticklabels(ylabels, fontsize=8)
        ax.invert_yaxis()
        ax.set_xlabel("Fault Current (kA)")
        ax.set_title(f"Equipment Limit Violations — {self.scenario}")
        ax.grid(axis="x", zorder=0)
        ax.legend(fontsize=9)
        fig.tight_layout()

        path = os.path.join(plots_dir, f"{label}_violations.png")
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        logger.info("Plot saved: %s", path)
        return path

    # ── Private: PDF export ───────────────────────────────────────────────────

    def _export_pdf(self, pdf_path: str, plot_paths: List[str]) -> None:
        """Assemble PDF report: plots + summary table + violations table."""
        with PdfPages(pdf_path) as pdf:

            # Embed each PNG plot as a full page
            for png in plot_paths:
                if png and os.path.isfile(png):
                    img = plt.imread(png)
                    fig, ax = plt.subplots(figsize=(14, 8))
                    ax.imshow(img)
                    ax.axis("off")
                    pdf.savefig(fig, bbox_inches="tight")
                    plt.close(fig)

            # Summary table page
            df_sum = self._summary_to_df()
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
                f"AESO Short Circuit Study — Summary\n{self.scenario}",
                fontsize=12, weight="bold", pad=15,
            )
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

            # Top 20 severity ranking table page
            df_top = self._severity_ranking_to_df()
            if not df_top.empty:
                fig2, ax2 = plt.subplots(
                    figsize=(16, max(4, len(df_top) * 0.45 + 2))
                )
                ax2.axis("off")
                tbl2 = ax2.table(
                    cellText=df_top.values,
                    colLabels=df_top.columns,
                    cellLoc="center", loc="center",
                )
                tbl2.auto_set_font_size(False)
                tbl2.set_fontsize(8)
                tbl2.scale(1.1, 1.8)
                for col in range(len(df_top.columns)):
                    tbl2[0, col].set_facecolor(_BLUE)
                    tbl2[0, col].set_text_props(color="white", weight="bold")
                for row in range(1, len(df_top) + 1):
                    if df_top.iloc[row - 1].get("Violation", False):
                        for col in range(len(df_top.columns)):
                            tbl2[row, col].set_facecolor("#ffe0e0")
                ax2.set_title(
                    "Top 20 Highest Fault Currents — Severity Ranking",
                    fontsize=11, weight="bold", pad=15,
                )
                plt.tight_layout()
                pdf.savefig(fig2, bbox_inches="tight")
                plt.close(fig2)

            # PDF metadata
            d = pdf.infodict()
            d["Title"]        = f"AESO Short Circuit Study — {self.scenario}"
            d["Author"]       = "AESO Automation Tool"
            d["CreationDate"] = datetime.now()

    # ── Private: DataFrame builders ───────────────────────────────────────────

    def _summary_to_df(self) -> pd.DataFrame:
        r = self.results
        if r is None:
            return pd.DataFrame()
        return pd.DataFrame([
            {"Parameter": "Scenario",
             "Value": r.scenario},
            {"Parameter": "Fault Types Run",
             "Value": ", ".join(r.fault_types_run)},
            {"Parameter": "Total Buses Faulted",
             "Value": len(set(f.bus_number for f in r.faults))},
            {"Parameter": "Total Fault Simulations",
             "Value": len(r.faults)},
            {"Parameter": "Equipment Violations",
             "Value": len(r.violations)},
            {"Parameter": "Max Fault Current (kA)",
             "Value": r.max_fault_current_ka},
            {"Parameter": "Critical Bus",
             "Value": r.max_fault_bus},
            {"Parameter": "Critical Fault Type",
             "Value": r.max_fault_type},
            {"Parameter": "Equipment Limit (kA)",
             "Value": self.max_fault_ka},
        ])

    def _faults_to_df(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "Severity Rank":          f.severity_rank,
            "Bus Number":             f.bus_number,
            "Bus Name":               f.bus_name,
            "Base kV":                f.base_kv,
            "Fault Type":             f.fault_type,
            "Fault Type (Full)":      self.FAULT_TYPE_MAP.get(f.fault_type, f.fault_type),
            "Fault Current (kA)":     f.fault_current_ka,
            "Fault Current Angle (°)":f.fault_current_ang,
            "Fault MVA":              f.fault_mva,
            "Pre-Fault Voltage (pu)": f.pre_fault_voltage_pu,
            "Violation":              f.violation,
        } for f in (self.results.faults if self.results else [])])

    def _violations_to_df(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "Bus Number":         f.bus_number,
            "Bus Name":           f.bus_name,
            "Base kV":            f.base_kv,
            "Fault Type":         f.fault_type,
            "Fault Current (kA)": f.fault_current_ka,
            "Limit (kA)":         self.max_fault_ka,
            "Excess (kA)":        round(f.fault_current_ka - self.max_fault_ka, 4),
        } for f in (self.results.violations if self.results else [])])

    def _severity_ranking_to_df(self) -> pd.DataFrame:
        """Top 20 faults by severity rank."""
        df = self._faults_to_df()
        if df.empty:
            return df
        return (
            df.nsmallest(20, "Severity Rank")
            .reset_index(drop=True)
        )

    # ── Mock data ─────────────────────────────────────────────────────────────

    def _mock_fault_results(self, fault_type: str) -> List[FaultResult]:
        """Synthetic fault results for all four types — used in unit tests."""
        import random
        random.seed(42 + hash(fault_type) % 100)

        mock_buses = [
            (101, "BUS_CALGARY",      240.0),
            (102, "BUS_EDMONTON",     240.0),
            (103, "BUS_LETHBRIDGE",   138.0),
            (104, "BUS_RED_DEER",      69.0),
            (105, "BUS_MEDICINE_HAT", 138.0),
        ]
        # Relative scaling of fault current by fault type
        # 3PH is the largest; LG ≈ 87%; LL ≈ 75%; LLG ≈ 90%
        scale = {"3PH": 1.00, "LG": 0.87, "LL": 0.75, "LLG": 0.90}.get(
            fault_type, 1.0
        )
        results = []
        for bnum, bname, bkv in mock_buses:
            ka  = round(random.uniform(5.0, 70.0) * scale, 4)
            ang = round(random.uniform(-85.0, -75.0), 2)
            results.append(FaultResult(
                bus_number           = bnum,
                bus_name             = bname,
                base_kv              = bkv,
                fault_type           = fault_type,
                fault_current_ka     = ka,
                fault_current_ang    = ang,
                fault_mva            = round(ka * bkv * 1.7321, 2),
                pre_fault_voltage_pu = round(random.uniform(0.98, 1.02), 5),
                violation            = ka > self.max_fault_ka,
            ))
        return results
