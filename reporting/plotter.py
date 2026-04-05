"""
reporting/plotter.py
---------------------
Standalone plot generation functions for all four AESO study types.

All functions accept DataFrames and return the saved file path.
They are called by main.py and the GUI after studies complete.
Each study module also generates its own plots internally —
these functions are for additional or combined plots requested
from outside the study modules.

Plots provided
--------------
Power Flow:
  plot_voltage_profile()      — Bus voltage bar chart with AESO limit lines
  plot_thermal_loading()      — Branch thermal loading horizontal bar chart
  plot_contingency_summary()  — N-1 scatter: min voltage vs violations

Short Circuit:
  plot_fault_currents()       — Grouped bar chart by bus and fault type
  plot_pre_post_comparison()  — Pre vs post-project fault current comparison

Transient Stability:
  plot_transient_response()   — Multi-panel time-domain response

PV Voltage Stability:
  plot_pv_curves()            — P-V curves for one or more scenarios
"""

import logging
import os
from typing import Dict, List, Optional

import matplotlib
matplotlib.use("Agg")   # non-interactive backend — safe for automated scripts
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# ── AESO colour palette ───────────────────────────────────────────────────────
AESO_BLUE   = "#003865"
AESO_RED    = "#C8102E"
AESO_ORANGE = "#E87722"
AESO_GREY   = "#6C6F70"
AESO_GREEN  = "#00853E"
PASS_COLOR  = AESO_GREEN
FAIL_COLOR  = AESO_RED

plt.rcParams.update({
    "figure.dpi":        150,
    "figure.facecolor":  "white",
    "axes.facecolor":    "#F7F7F7",
    "axes.edgecolor":    AESO_GREY,
    "axes.labelcolor":   AESO_BLUE,
    "axes.titleweight":  "bold",
    "axes.titlecolor":   AESO_BLUE,
    "xtick.color":       AESO_GREY,
    "ytick.color":       AESO_GREY,
    "grid.color":        "white",
    "grid.linewidth":    1.0,
    "font.family":       "DejaVu Sans",
})


def _save(fig: plt.Figure, path: str) -> str:
    """Save figure and close it. Returns the saved path."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    logger.info("Plot saved: %s", path)
    return path


# ── Power Flow Plots ──────────────────────────────────────────────────────────

def plot_voltage_profile(
    bus_df:     pd.DataFrame,
    scenario:   str,
    output_dir: str,
    v_min:      float = 0.95,
    v_max:      float = 1.05,
    filename:   Optional[str] = None,
) -> str:
    """
    Bar chart of bus voltage magnitudes with AESO limit lines.

    Parameters
    ----------
    bus_df : DataFrame
        Required columns: ["Bus Name", "Voltage (pu)", "Violation"]
    scenario : str
        Scenario label used in the chart title.
    output_dir : str
        Directory where the PNG is saved.
    v_min, v_max : float
        AESO voltage limits (pu) shown as horizontal reference lines.
    filename : str, optional
        Override the default filename.

    Returns
    -------
    str : Path to saved PNG file.
    """
    os.makedirs(output_dir, exist_ok=True)

    df     = bus_df.sort_values("Voltage (pu)")
    colors = [FAIL_COLOR if v else PASS_COLOR for v in df["Violation"]]

    fig, ax = plt.subplots(figsize=(max(10, len(df) * 0.35), 6))
    ax.bar(range(len(df)), df["Voltage (pu)"], color=colors, width=0.7, zorder=3)

    ax.axhline(v_min, color=AESO_RED,    linestyle="--", linewidth=1.5,
               label=f"Min = {v_min} pu")
    ax.axhline(v_max, color=AESO_ORANGE, linestyle="--", linewidth=1.5,
               label=f"Max = {v_max} pu")
    ax.axhline(1.00,  color=AESO_BLUE,   linestyle=":",  linewidth=1.0,
               alpha=0.5, label="Nominal (1.00 pu)")

    ax.set_xticks(range(len(df)))
    ax.set_xticklabels(df["Bus Name"], rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Voltage (pu)")
    ax.set_title(f"Bus Voltage Profile — {scenario}")
    ax.set_ylim(
        min(0.85, df["Voltage (pu)"].min() - 0.02),
        max(1.15, df["Voltage (pu)"].max() + 0.02),
    )
    ax.grid(axis="y", zorder=0)
    ax.legend(handles=[
        mpatches.Patch(color=PASS_COLOR,  label="Within Limits"),
        mpatches.Patch(color=FAIL_COLOR,  label="Violation"),
        mpatches.Patch(color=AESO_RED,    label=f"Min {v_min} pu"),
        mpatches.Patch(color=AESO_ORANGE, label=f"Max {v_max} pu"),
    ], loc="lower right", fontsize=8)
    fig.tight_layout()

    fname = filename or f"voltage_profile_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


def plot_thermal_loading(
    branch_df:  pd.DataFrame,
    scenario:   str,
    output_dir: str,
    limit_pct:  float = 100.0,
    warn_pct:   float = 95.0,
    top_n:      int   = 30,
    filename:   Optional[str] = None,
) -> str:
    """
    Horizontal bar chart of branch thermal loading, top N most loaded.

    Parameters
    ----------
    branch_df : DataFrame
        Required columns: ["From Bus", "To Bus", "Circuit ID",
                           "Loading %", "Thermal Violation"]
    warn_pct : float
        Branches at or above this % are highlighted in orange (warning).
    """
    os.makedirs(output_dir, exist_ok=True)

    df     = branch_df.sort_values("Loading %", ascending=False).head(top_n)
    labels = df.apply(
        lambda r: f"{r['From Bus']}-{r['To Bus']} [{r['Circuit ID']}]", axis=1
    )

    def _bar_color(row):
        if row["Thermal Violation"]:
            return FAIL_COLOR
        if row["Loading %"] >= warn_pct:
            return AESO_ORANGE
        return PASS_COLOR

    colors = [_bar_color(row) for _, row in df.iterrows()]

    fig, ax = plt.subplots(figsize=(10, max(6, len(df) * 0.35)))
    ax.barh(range(len(df)), df["Loading %"], color=colors, height=0.7, zorder=3)

    ax.axvline(limit_pct, color=AESO_RED, linestyle="--", linewidth=1.5,
               label=f"Limit = {limit_pct:.0f}%")
    ax.axvline(warn_pct, color=AESO_ORANGE, linestyle=":", linewidth=1.2,
               label=f"Warning = {warn_pct:.0f}%")

    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("Thermal Loading (%)")
    ax.set_title(f"Branch Thermal Loading (Top {top_n}) — {scenario}")
    ax.grid(axis="x", zorder=0)
    ax.legend(handles=[
        mpatches.Patch(color=FAIL_COLOR,  label=f"Violation (>{limit_pct:.0f}%)"),
        mpatches.Patch(color=AESO_ORANGE, label=f"Warning (>{warn_pct:.0f}%)"),
        mpatches.Patch(color=PASS_COLOR,  label="Normal"),
    ] + [plt.Line2D([0],[0], color=AESO_RED, linestyle="--",
                    label=f"Limit {limit_pct:.0f}%")],
        fontsize=8)
    fig.tight_layout()

    fname = filename or f"thermal_loading_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


def plot_contingency_summary(
    contingency_df: pd.DataFrame,
    scenario:       str,
    output_dir:     str,
    v_min_cont:     float = 0.90,
    filename:       Optional[str] = None,
) -> str:
    """
    Scatter plot: N-1 post-contingency min voltage vs. total violations.

    Parameters
    ----------
    contingency_df : DataFrame
        Required columns: ["Min Voltage (pu)", "Bus Violations",
                           "Branch Violations"]
    v_min_cont : float
        Category B voltage limit — shown as vertical reference line.
    """
    os.makedirs(output_dir, exist_ok=True)

    if contingency_df.empty:
        logger.warning("No contingency data to plot for %s.", scenario)
        return ""

    total_viol = (contingency_df["Bus Violations"]
                  + contingency_df["Branch Violations"])

    fig, ax = plt.subplots(figsize=(10, 6))
    scatter = ax.scatter(
        contingency_df["Min Voltage (pu)"],
        total_viol,
        c=total_viol,
        cmap="RdYlGn_r",
        s=60, alpha=0.8, zorder=3,
    )
    plt.colorbar(scatter, ax=ax, label="Total Violations")

    ax.axvline(v_min_cont, color=AESO_RED, linestyle="--", linewidth=1.5,
               label=f"Cat B Min V = {v_min_cont} pu")
    ax.set_xlabel("Post-Contingency Min Bus Voltage (pu)")
    ax.set_ylabel("Total Violations (Bus + Branch)")
    ax.set_title(f"N-1 Contingency Overview — {scenario}")
    ax.grid(zorder=0)
    ax.legend(fontsize=9)
    fig.tight_layout()

    fname = filename or f"n1_contingency_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


# ── Short Circuit Plots ───────────────────────────────────────────────────────

def plot_fault_currents(
    fault_df:   pd.DataFrame,
    scenario:   str,
    output_dir: str,
    limit_ka:   float = 63.0,
    top_n:      int   = 20,
    filename:   Optional[str] = None,
) -> str:
    """
    Grouped bar chart of fault current by bus and fault type.

    Parameters
    ----------
    fault_df : DataFrame
        Required columns: ["Severity Rank", "Bus Name", "Fault Type",
                           "Fault Current (kA)", "Violation"]
    """
    os.makedirs(output_dir, exist_ok=True)

    df = fault_df.nsmallest(top_n, "Severity Rank")
    if df.empty:
        return ""

    fault_types = ["3PH", "LG", "LL", "LLG"]
    fault_types = [ft for ft in fault_types if ft in df["Fault Type"].unique()]
    top_buses   = (
        df.groupby("Bus Name")["Fault Current (kA)"]
        .max()
        .nlargest(top_n)
        .index.tolist()
    )

    x      = np.arange(len(top_buses))
    width  = 0.8 / max(len(fault_types), 1)
    colors = [AESO_BLUE, AESO_RED, AESO_ORANGE, AESO_GREEN]

    fig, ax = plt.subplots(figsize=(14, 6))
    for j, ft in enumerate(fault_types):
        sub  = df[df["Fault Type"] == ft].set_index("Bus Name")
        vals = [
            float(sub.loc[b, "Fault Current (kA)"])
            if b in sub.index else 0.0
            for b in top_buses
        ]
        ax.bar(
            x + j * width, vals,
            width=width * 0.9,
            label=ft,
            color=colors[j % len(colors)],
            zorder=3,
        )

    ax.axhline(limit_ka, color=AESO_RED, linestyle="--", linewidth=1.8,
               label=f"Equipment Limit = {limit_ka} kA")
    ax.set_xticks(x + width * (len(fault_types) - 1) / 2)
    ax.set_xticklabels(top_buses, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Fault Current (kA)")
    ax.set_title(f"Short Circuit Fault Currents (Top {top_n} Buses) — {scenario}")
    ax.grid(axis="y", zorder=0)
    ax.legend(fontsize=9)
    fig.tight_layout()

    fname = filename or f"fault_currents_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


def plot_pre_post_comparison(
    comparison_df: pd.DataFrame,
    scenario:      str,
    output_dir:    str,
    filename:      Optional[str] = None,
) -> str:
    """
    Pre-project vs. post-project fault current comparison bar chart.

    Parameters
    ----------
    comparison_df : DataFrame
        Output of ShortCircuitStudy.compare() — requires columns:
        ["Bus Name", "Fault Type", "Pre-Project (kA)", "Post-Project (kA)"]
    """
    os.makedirs(output_dir, exist_ok=True)
    if comparison_df.empty:
        return ""

    top = comparison_df.head(20)
    x   = np.arange(len(top))
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(x - 0.2, top["Pre-Project (kA)"],  0.4,
           label="Pre-Project",  color=AESO_BLUE,   zorder=3)
    ax.bar(x + 0.2, top["Post-Project (kA)"], 0.4,
           label="Post-Project", color=AESO_ORANGE,  zorder=3)

    labels = top.apply(
        lambda r: f"{r['Bus Name']} ({r['Fault Type']})", axis=1
    )
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Fault Current (kA)")
    ax.set_title(f"Pre vs Post-Project Fault Current — {scenario}")
    ax.grid(axis="y", zorder=0)
    ax.legend(fontsize=9)
    fig.tight_layout()

    fname = filename or f"pre_post_comparison_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


# ── Transient Stability Plots ─────────────────────────────────────────────────

def plot_transient_response(
    time:          List[float],
    channels:      Dict[str, List[float]],
    scenario:      str,
    output_dir:    str,
    fault_time:    float = 1.0,
    clear_time:    float = 1.1,
    v_recovery_pu: float = 0.90,
    angle_limit:   float = 180.0,
    filename:      Optional[str] = None,
) -> str:
    """
    Multi-panel time-domain response plot.

    Parameters
    ----------
    time : list of float
        Time axis in seconds.
    channels : dict
        Mapping channel_name → list of float values.
        Expected key patterns: "voltage", "angle"/"rotor", "power"
    fault_time, clear_time : float
        Fault application and clearing times (shaded region).
    v_recovery_pu : float
        AESO voltage recovery limit — horizontal reference line on voltage panel.
    angle_limit : float
        AESO rotor angle limit — horizontal reference line on angle panel.
    """
    os.makedirs(output_dir, exist_ok=True)
    if not channels:
        return ""

    # Sort channels into panels
    volt_chs  = {n: v for n, v in channels.items()
                 if "voltage" in n.lower() or "volt" in n.lower()}
    angle_chs = {n: v for n, v in channels.items()
                 if "angle" in n.lower() or "rotor" in n.lower()}
    other_chs = {n: v for n, v in channels.items()
                 if n not in volt_chs and n not in angle_chs}

    panel_data = []
    if volt_chs:  panel_data.append(("Bus Voltage (pu)",  volt_chs,  v_recovery_pu))
    if angle_chs: panel_data.append(("Rotor Angle (deg)", angle_chs, angle_limit))
    if other_chs: panel_data.append(("Power",             other_chs, None))
    if not panel_data:
        return ""

    n_panels = len(panel_data)
    fig, axes = plt.subplots(
        n_panels, 1,
        figsize=(12, 3 * n_panels),
        sharex=True,
    )
    if n_panels == 1:
        axes = [axes]

    colors = [AESO_BLUE, AESO_RED, AESO_ORANGE, AESO_GREEN, AESO_GREY]

    for ax, (ylabel, chs, ref_val) in zip(axes, panel_data):
        for i, (name, values) in enumerate(list(chs.items())[:5]):
            ax.plot(time, values,
                    color=colors[i % len(colors)],
                    linewidth=1.5,
                    label=name[:40])
        # Shade fault period
        ax.axvspan(fault_time, clear_time,
                   alpha=0.15, color=AESO_RED, label="Fault period")
        # Reference limit line
        if ref_val is not None:
            ax.axhline(ref_val,
                       color=AESO_RED, linestyle="--", linewidth=1.2,
                       label=f"Limit = {ref_val}")
        ax.set_ylabel(ylabel, fontsize=9)
        ax.grid(True, zorder=0)
        ax.legend(fontsize=7, loc="upper right", ncol=2)

    axes[-1].set_xlabel("Time (s)", fontsize=10)
    axes[0].set_title(f"Transient Stability Response — {scenario}", fontsize=11)
    fig.tight_layout()

    fname = filename or f"transient_response_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))


# ── PV Voltage Stability Plots ────────────────────────────────────────────────

def plot_pv_curves(
    pv_data:    Dict[str, Dict],
    scenario:   str,
    output_dir: str,
    v_min_a:    float = 0.95,
    v_min_b:    float = 0.90,
    marp_mw:    float = 0.0,
    nose_points:Optional[Dict[str, tuple]] = None,
    filename:   Optional[str] = None,
) -> str:
    """
    P-V curves for one or more scenarios.

    Parameters
    ----------
    pv_data : dict
        Mapping scenario_name → {"P_mw": [...], "V_pu": [...],
                                  "category": "A" or "B"}
    nose_points : dict, optional
        Mapping scenario_name → (P_collapse_mw, V_collapse_pu)
    marp_mw : float
        MARP in MW — shown as vertical reference line if > 0.
    """
    os.makedirs(output_dir, exist_ok=True)
    if not pv_data:
        return ""

    colors = [AESO_BLUE, AESO_RED, AESO_ORANGE, AESO_GREEN, AESO_GREY]
    fig, ax = plt.subplots(figsize=(11, 7))

    for i, (name, data) in enumerate(pv_data.items()):
        color    = colors[i % len(colors)]
        cat      = data.get("category", "A")
        ls       = "-" if cat == "A" else "--"
        ax.plot(
            data["P_mw"], data["V_pu"],
            color=color, linewidth=2, linestyle=ls,
            marker="o", markersize=3,
            label=name,
        )
        if nose_points and name in nose_points:
            p_nose, v_nose = nose_points[name]
            ax.scatter([p_nose], [v_nose],
                       color=color, marker="*", s=200, zorder=5)
            ax.annotate(
                f"Collapse\n{p_nose:.0f} MW",
                (p_nose, v_nose),
                textcoords="offset points", xytext=(10, -15),
                fontsize=8, color=color,
            )

    ax.axhline(v_min_b, color=AESO_RED, linestyle="--", linewidth=1.5,
               label=f"AESO Cat B V_min = {v_min_b:.2f} pu")
    ax.axhline(v_min_a, color=AESO_ORANGE, linestyle="--", linewidth=1.5,
               label=f"AESO Cat A V_min = {v_min_a:.2f} pu")

    if marp_mw > 0:
        ax.axvline(marp_mw, color=AESO_GREEN, linestyle="--", linewidth=1.5,
                   label=f"MARP = {marp_mw:.0f} MW")

    ax.set_xlabel("Active Power (MW)", fontsize=11)
    ax.set_ylabel("POI Bus Voltage (pu)", fontsize=11)
    ax.set_title(f"P-V Voltage Stability Curves — {scenario}", fontsize=12)
    ax.set_ylim(0.82, 1.08)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(50))
    ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.65)
    ax.legend(fontsize=8, loc="lower left", framealpha=0.9)
    fig.tight_layout()

    fname = filename or f"pv_curves_{scenario}.png"
    return _save(fig, os.path.join(output_dir, fname))
