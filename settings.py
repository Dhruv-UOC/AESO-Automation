"""
config/settings.py
------------------
Central configuration for the AESO Interconnection Study Automation Tool.

╔══════════════════════════════════════════════════════════════════╗
║         USER CONFIGURATION — EDIT THIS SECTION ONLY             ║
║  Change these two values when switching machines or pen drives   ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os

# ── USER CONFIGURATION ────────────────────────────────────────────────────────
# 1. Root folder on the pen drive — contains projects/ and templates/ folders
BASE_DIR = r"D:\Final_Project\files"

# 2. PSS/E installation PSSBIN directory on the current machine
PSSE_PATH = r"C:\Program Files\PTI\PSSE35\35.0\PSSBIN"
# ─────────────────────────────────────────────────────────────────────────────
# DO NOT EDIT BELOW THIS LINE UNLESS CHANGING AESO THRESHOLDS
# ─────────────────────────────────────────────────────────────────────────────

# ── PSS/E Version ─────────────────────────────────────────────────────────────
PSSE_VERSION = 35
# Only change if you upgrade PSS/E to a different major version.
# Supported: 33, 34, 35

# ── Shared Directories ────────────────────────────────────────────────────────
# These are shared across all projects and do not change per project.
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
PROJECTS_DIR  = os.path.join(BASE_DIR, "projects")

# ── Default SAV and Results (used by utils/bus_listing.py fallback only) ──────
# Once you create a project folder, set the SAV path in
# the Project_Info sheet of study_scope_data.xlsx instead.
DEFAULT_SAV = os.path.join(BASE_DIR, "cases", "sample.sav")
RESULTS_DIR = os.path.join(BASE_DIR, "output", "results")

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_LEVEL = "INFO"    # DEBUG | INFO | WARNING | ERROR
LOG_FILE  = os.path.join(BASE_DIR, "output", "automation.log")

# ── AESO Compliance Thresholds ────────────────────────────────────────────────
# These are fixed AESO criteria that apply to ALL projects.
# All project-specific data (bus numbers, scenarios, dispatch, contingencies)
# lives in each project's study_scope_data.xlsx — NOT here.
AESO = {
    # ── Power Flow — Category A (N-0) normal operation ──────────────────────
    "voltage_min_pu":             0.95,   # Normal minimum voltage (pu)
    "voltage_max_pu":             1.05,   # Normal maximum voltage (pu)

    # ── Power Flow — Category B (N-1) post-contingency ──────────────────────
    "voltage_min_contingency":    0.90,   # Post-contingency minimum voltage (pu)
    "voltage_max_contingency":    1.10,   # Post-contingency maximum voltage (pu)

    # ── Post-contingency voltage deviation limits (Study Scope Table 3-1) ───
    "post_transient_dev":         0.10,   # ±10%  within first 30 seconds
    "post_auto_dev":              0.07,   # ±7%   30 sec – 5 min (after auto controls)
    "post_manual_dev":            0.05,   # ±5%   steady state (after manual control)

    # ── Thermal loading ──────────────────────────────────────────────────────
    "thermal_limit_pct":          100.0,  # % of rated MVA — violation if exceeded
    "thermal_warning_pct":        95.0,   # % of rated MVA — warning level
                                          # Study Scope Section 5.2: report separately

    # ── Short Circuit ────────────────────────────────────────────────────────
    "max_fault_current_ka":       63.0,   # Equipment withstand limit (kA)

    # ── Transient Stability ──────────────────────────────────────────────────
    "rotor_angle_limit_deg":      180.0,  # First-swing rotor angle stability limit
    "voltage_recovery_time_s":    1.0,    # Post-fault window to check recovery (s)
    "voltage_recovery_pu":        0.90,   # Minimum voltage after recovery window

    # Per-contingency fault clearing times come from TS_Contingencies sheet.
    # The values below are fallback defaults per AESO Table 2-3
    # (used only when a contingency has no FCT specified).
    "ts_default_near_end_cycles": 6.0,   # 138/144 kV with telecom
    "ts_default_far_end_cycles":  8.0,
    "ts_sim_duration_s":          10.0,  # Total simulation window (s)
    "ts_fault_apply_s":           1.0,   # Time at which fault is applied (s)
    "ts_time_step_s":             0.00833,  # 1/2 cycle at 60 Hz

    # ── PV Voltage Stability ─────────────────────────────────────────────────
    "pv_cat_a_v_min":             0.95,  # Category A (N-0) minimum POI voltage (pu)
    "pv_cat_a_v_max":             1.05,  # Category A (N-0) maximum POI voltage (pu)
    "pv_cat_b_v_min":             0.90,  # Category B (N-1) minimum POI voltage (pu)
    "pv_cat_b_v_max":             1.10,  # Category B (N-1) maximum POI voltage (pu)
    "pv_step_mw":                 10.0,  # MW increment per power flow solve
    # transfer_end_mw is NOT set here.
    # It defaults to MARP × 1.05, read from Project_Info sheet per project.
}
