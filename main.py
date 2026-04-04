"""
main.py
-------
Entry point for the AESO Interconnection Study Automation Tool.

Supports two modes:
  1. GUI mode  (default / --gui flag)
  2. CLI mode  (--project flag)

GUI mode
--------
    python main.py
    python main.py --gui

    Opens the tkinter GUI. All study configuration is done interactively.

CLI mode
--------
    python main.py --project D:\\Final_Project\\files\\projects\\P2611

    Reads study_scope_data.xlsx from the project folder,
    runs all studies marked in Study_Matrix for all scenarios,
    and writes output to the project's output/ subfolder.

    Optional CLI flags:
        --sav       PATH     Override the SAV file path
        --study     TYPES    Comma-separated: power_flow,short_circuit,
                             transient,voltage_stability (default: all)
        --scenario  NAMES    Comma-separated scenario names to run
                             (default: all scenarios in Study_Matrix)
        --mock               Use mock PSS/E (no licence required, for testing)
        --validate-only      Validate project data and exit without running

Examples
--------
    # Open GUI
    python main.py

    # Run all studies for P2611 from CLI
    python main.py --project D:\\Final_Project\\files\\projects\\P2611

    # Run only power flow and short circuit
    python main.py --project D:\\...\\P2611 --study power_flow,short_circuit

    # Run specific scenarios
    python main.py --project D:\\...\\P2611 --scenario "2028 SP Post-Project"

    # Test without PSS/E licence
    python main.py --project D:\\...\\P2611 --mock

    # Validate Excel data only
    python main.py --project D:\\...\\P2611 --validate-only
"""

import argparse
import logging
import os
import sys

# ── Ensure project root is on sys.path ────────────────────────────────────────
_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


# ── Logging setup ─────────────────────────────────────────────────────────────

def setup_logging(log_level: str, log_file: str) -> None:
    """Configure root logger with console and file handlers."""
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    logging.basicConfig(
        level=getattr(logging, log_level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, mode="a", encoding="utf-8"),
        ],
    )


logger = logging.getLogger(__name__)


# ── CLI argument parser ───────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="AESO Interconnection Study Automation Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the tkinter GUI (default if no --project given).",
    )
    parser.add_argument(
        "--project",
        metavar="DIR",
        help="Path to project folder containing study_scope_data.xlsx.",
    )
    parser.add_argument(
        "--sav",
        metavar="PATH",
        help="Override SAV file path (otherwise read from Project_Info sheet).",
    )
    parser.add_argument(
        "--study",
        metavar="TYPES",
        default="all",
        help=(
            "Comma-separated study types to run. "
            "Options: power_flow, short_circuit, transient, voltage_stability, all. "
            "Default: all."
        ),
    )
    parser.add_argument(
        "--scenario",
        metavar="NAMES",
        help=(
            "Comma-separated scenario names to run. "
            "Default: all scenarios required by Study_Matrix."
        ),
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        help="Run with mock PSS/E (no licence required). For testing only.",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Validate project data and print warnings. Exit without running.",
    )
    return parser


# ── Startup validation ────────────────────────────────────────────────────────

def validate_environment(psse_path: str) -> list:
    """
    Check that PSS/E is installed and output directories are writable.
    Returns list of warning strings. Does NOT raise — caller decides.
    """
    warnings = []
    if not os.path.isdir(psse_path):
        warnings.append(
            f"[CRITICAL] PSS/E PSSBIN directory not found: {psse_path}\n"
            f"  Update PSSE_PATH in config/settings.py."
        )
    else:
        psspy37 = os.path.abspath(os.path.join(psse_path, "..", "PSSPY37"))
        if not os.path.isdir(psspy37):
            warnings.append(
                f"[WARNING] PSSPY37 directory not found: {psspy37}\n"
                f"  psspy import may fail."
            )
    return warnings


def validate_project_folder(project_dir: str) -> list:
    """Check project folder structure."""
    warnings = []
    if not os.path.isdir(project_dir):
        warnings.append(f"[CRITICAL] Project folder not found: {project_dir}")
        return warnings

    excel_path = os.path.join(project_dir, "study_scope_data.xlsx")
    if not os.path.isfile(excel_path):
        warnings.append(
            f"[CRITICAL] study_scope_data.xlsx not found in: {project_dir}\n"
            f"  Copy templates/study_scope_template.xlsx to {excel_path} "
            f"and fill it from your Study Scope PDF."
        )
    return warnings


# ── Study runner ──────────────────────────────────────────────────────────────

def run_cli(
    project_dir:       str,
    sav_override:      str,
    study_types:       list,
    scenario_filter:   list,
    mock:              bool,
) -> int:
    """
    Run studies from CLI. Returns exit code (0=success, 1=failure).
    """
    from config.settings import PSSE_PATH, PSSE_VERSION, AESO
    from core.psse_interface import PSSEInterface
    from project_io.excel_reader import ExcelReader, ExcelReaderError
    from project_io.project_data import ProjectData

    # ── Load project data ─────────────────────────────────────────────────────
    excel_path = os.path.join(project_dir, "study_scope_data.xlsx")
    logger.info("Loading project data: %s", excel_path)

    try:
        reader  = ExcelReader(excel_path)
        project = reader.read()
    except ExcelReaderError as exc:
        logger.error("Failed to load project data: %s", exc)
        return 1

    # Set paths
    project.project_dir = project_dir
    project.output_dir  = os.path.join(project_dir, "output")

    # Resolve SAV file
    sav_path = (
        sav_override
        or project.info.sav_file_path
        or _find_sav(project_dir)
    )
    if sav_path:
        project.info.sav_file_path = sav_path
    if not sav_path or not os.path.isfile(sav_path):
        logger.error(
            "SAV file not found: '%s'\n"
            "  Set SAV File Path in Project_Info sheet or use --sav flag.",
            sav_path or "(not set)"
        )
        return 1

    # ── Validate project data ─────────────────────────────────────────────────
    env_warnings  = validate_environment(PSSE_PATH)
    data_warnings = project.validate()
    all_warnings  = env_warnings + data_warnings

    if all_warnings:
        logger.info("Validation warnings (%d):", len(all_warnings))
        for w in all_warnings:
            if "[CRITICAL]" in w:
                logger.error("  %s", w)
            else:
                logger.warning("  %s", w)

    critical = [w for w in all_warnings if "[CRITICAL]" in w]
    if critical and not mock:
        logger.error(
            "%d critical issue(s) found. Resolve before running. "
            "Use --mock to run anyway with synthetic data.",
            len(critical)
        )
        return 1

    # ── Determine scenarios to run ────────────────────────────────────────────
    if scenario_filter:
        scenarios = [
            sc for sc in project.scenarios
            if sc.scenario_name in scenario_filter
        ]
        if not scenarios:
            logger.error(
                "No scenarios matched filter: %s\n"
                "Available: %s",
                scenario_filter,
                [sc.scenario_name for sc in project.scenarios],
            )
            return 1
    else:
        # Run all scenarios that have at least one study required
        scenarios = [
            sc for sc in project.scenarios
            if _scenario_needs_any_study(project, sc.scenario_name, study_types)
        ]

    if not scenarios:
        logger.warning(
            "No scenarios require the selected studies. "
            "Check Study_Matrix sheet."
        )
        return 0

    logger.info(
        "Running %d scenario(s): %s",
        len(scenarios),
        ", ".join(sc.scenario_name for sc in scenarios)
    )

    # ── Initialise PSS/E ──────────────────────────────────────────────────────
    logger.info("Initialising PSS/E (version=%d, mock=%s)…", PSSE_VERSION, mock)
    psse = PSSEInterface(
        psse_path=PSSE_PATH,
        psse_version=PSSE_VERSION,
        mock=mock,
    )
    try:
        psse.initialize()
    except Exception as exc:
        logger.error("PSS/E initialisation failed: %s", exc)
        if not mock:
            logger.info(
                "Tip: Use --mock flag to run with synthetic data "
                "while PSS/E licence issues are resolved."
            )
            return 1

    # ── Output directories ────────────────────────────────────────────────────
    results_dir = os.path.join(project.output_dir, "results")
    plots_dir   = os.path.join(project.output_dir, "plots")
    reports_dir = os.path.join(project.output_dir, "reports")
    for d in (results_dir, plots_dir, reports_dir):
        os.makedirs(d, exist_ok=True)

    # ── Run studies per scenario ──────────────────────────────────────────────
    exit_code = 0
    for sc in scenarios:
        logger.info("=" * 60)
        logger.info("Scenario: %s  (%s)", sc.scenario_name, sc.pre_post)
        logger.info("=" * 60)

        matrix = project.get_study_matrix(sc.scenario_name)

        # Power Flow
        if "power_flow" in study_types or "all" in study_types:
            try:
                exit_code |= _run_power_flow(
                    psse, project, sc, sav_path,
                    results_dir, plots_dir, reports_dir, AESO
                )
            except Exception as exc:
                logger.error("Power Flow failed for %s: %s", sc.scenario_name, exc)
                exit_code = 1

        # Short Circuit
        if "short_circuit" in study_types or "all" in study_types:
            try:
                exit_code |= _run_short_circuit(
                    psse, project, sc, sav_path,
                    results_dir, plots_dir, reports_dir, AESO
                )
            except Exception as exc:
                logger.error("Short Circuit failed for %s: %s", sc.scenario_name, exc)
                exit_code = 1

        # Transient Stability
        if "transient" in study_types or "all" in study_types:
            needs_ts = matrix and (
                matrix.transient_cat_a
                or matrix.transient_cat_b
                or matrix.transient_conditional
            )
            if needs_ts:
                try:
                    exit_code |= _run_transient(
                        psse, project, sc, sav_path,
                        results_dir, plots_dir, reports_dir, AESO
                    )
                except Exception as exc:
                    logger.error(
                        "Transient Stability failed for %s: %s",
                        sc.scenario_name, exc
                    )
                    exit_code = 1
            else:
                logger.info(
                    "Transient Stability not required for %s (Study_Matrix).",
                    sc.scenario_name
                )

        # PV Voltage Stability
        if "voltage_stability" in study_types or "all" in study_types:
            needs_vs = matrix and (
                matrix.volt_stability_cat_a or matrix.volt_stability_cat_b
            )
            if needs_vs:
                try:
                    exit_code |= _run_pv_stability(
                        psse, project, sc, sav_path,
                        results_dir, plots_dir, reports_dir, AESO
                    )
                except Exception as exc:
                    logger.error(
                        "PV Voltage Stability failed for %s: %s",
                        sc.scenario_name, exc
                    )
                    exit_code = 1
            else:
                logger.info(
                    "PV Voltage Stability not required for %s (Study_Matrix).",
                    sc.scenario_name
                )

    logger.info("=" * 60)
    logger.info(
        "All studies complete. Status: %s",
        "SUCCESS" if exit_code == 0 else "COMPLETED WITH ERRORS"
    )
    logger.info("Output directory: %s", project.output_dir)
    return exit_code


# ── Individual study runners ──────────────────────────────────────────────────

def _run_power_flow(
    psse, project, sc, sav_path,
    results_dir, plots_dir, reports_dir, aeso_cfg,
) -> int:
    from studies.power_flow.power_flow_study import PowerFlowStudy
    logger.info("--- Power Flow ---")
    study = PowerFlowStudy(
        psse,
        scenario_label          = sc.scenario_name,
        project                 = project,
        season_label            = f"{sc.year} {sc.season}",
        voltage_min             = aeso_cfg["voltage_min_pu"],
        voltage_max             = aeso_cfg["voltage_max_pu"],
        voltage_min_contingency = aeso_cfg["voltage_min_contingency"],
        voltage_max_contingency = aeso_cfg["voltage_max_contingency"],
        thermal_limit_pct       = aeso_cfg["thermal_limit_pct"],
    )
    results = study.run(sav_path)
    files   = study.save_results(results_dir, plots_dir, reports_dir)
    _print_power_flow_summary(sc.scenario_name, results)
    logger.info("Power Flow outputs: %s", files)
    return 0


def _run_short_circuit(
    psse, project, sc, sav_path,
    results_dir, plots_dir, reports_dir, aeso_cfg,
) -> int:
    from studies.short_circuit.short_circuit_study import ShortCircuitStudy
    logger.info("--- Short Circuit ---")
    # Use SC target substations bus filter if available
    sc_filter = [
        s.bus_no for s in project.sc_substations if s.bus_no is not None
    ] or None
    study = ShortCircuitStudy(
        psse,
        scenario_label       = sc.scenario_name,
        max_fault_current_ka = aeso_cfg["max_fault_current_ka"],
        bus_filter           = sc_filter,
    )
    results = study.run(sav_path)
    files   = study.save_results(results_dir, plots_dir, reports_dir)
    _print_short_circuit_summary(sc.scenario_name, results)
    logger.info("Short Circuit outputs: %s", files)
    return 0


def _run_transient(
    psse, project, sc, sav_path,
    results_dir, plots_dir, reports_dir, aeso_cfg,
) -> int:
    from studies.transient_stability.transient_stability_study import (
        TransientStabilityStudy,
    )
    logger.info("--- Transient Stability ---")
    study = TransientStabilityStudy(
        psse, project,
        scenario_label            = sc.scenario_name,
        rotor_angle_limit_deg     = aeso_cfg["rotor_angle_limit_deg"],
        voltage_recovery_pu       = aeso_cfg["voltage_recovery_pu"],
        voltage_recovery_window_s = aeso_cfg["voltage_recovery_time_s"],
    )
    results = study.run(sav_path)
    files   = study.save_results(results_dir, plots_dir, reports_dir)
    _print_transient_summary(sc.scenario_name, results)
    logger.info("Transient Stability outputs: %s", files)
    return 0


def _run_pv_stability(
    psse, project, sc, sav_path,
    results_dir, plots_dir, reports_dir, aeso_cfg,
) -> int:
    from studies.pv_voltage.pv_stability_study import PVStabilityStudy
    logger.info("--- PV Voltage Stability ---")

    if project.info.source_bus_number is None or project.info.poi_bus_number is None:
        logger.warning(
            "PV Voltage Stability skipped for %s — "
            "Source Bus or POI Bus not set. "
            "Fill Bus_Numbers sheet and Project_Info after running "
            "utils/bus_listing.py.",
            sc.scenario_name
        )
        return 0

    study = PVStabilityStudy(
        psse, project,
        scenario_label = sc.scenario_name,
        v_min_cat_a    = aeso_cfg["pv_cat_a_v_min"],
        v_min_cat_b    = aeso_cfg["pv_cat_b_v_min"],
    )
    results = study.run(sav_path)
    files   = study.save_results(results_dir, plots_dir, reports_dir)
    _print_pv_summary(sc.scenario_name, results)
    logger.info("PV Voltage Stability outputs: %s", files)
    return 0


# ── Console summaries ─────────────────────────────────────────────────────────

def _print_power_flow_summary(scenario: str, results) -> None:
    sep = "-" * 56
    print(f"\n{sep}")
    print(f"  Power Flow | {scenario}")
    print(sep)
    print(f"  Converged          : {results.converged}")
    print(f"  Total Gen  (MW)    : {results.total_generation_mw:.1f}")
    print(f"  Total Load (MW)    : {results.total_load_mw:.1f}")
    print(f"  Losses     (MW)    : {results.total_losses_mw:.1f}")
    print(f"  Bus Violations     : {len(results.bus_violations)}")
    print(f"  Branch Violations  : {len(results.branch_violations)}")
    print(f"  N-1 Contingencies  : {len(results.contingencies)}")
    for b in results.bus_violations[:5]:
        print(f"    [!] Bus {b.bus_number} {b.bus_name}: {b.violation_type}")
    print(sep)


def _print_short_circuit_summary(scenario: str, results) -> None:
    sep = "-" * 56
    print(f"\n{sep}")
    print(f"  Short Circuit | {scenario}")
    print(sep)
    print(f"  Fault types        : {', '.join(results.fault_types_run)}")
    print(f"  Total faults       : {len(results.faults)}")
    print(f"  Violations         : {len(results.violations)}")
    print(f"  Max fault (kA)     : {results.max_fault_current_ka:.3f}")
    print(f"  Critical bus       : {results.max_fault_bus} ({results.max_fault_type})")
    print(sep)


def _print_transient_summary(scenario: str, results) -> None:
    sep = "-" * 56
    print(f"\n{sep}")
    print(f"  Transient Stability | {scenario}")
    print(sep)
    print(f"  Contingencies run  : {len(results.contingencies)}")
    print(f"  AESO PASS          : {results.total_pass}")
    print(f"  AESO FAIL          : {results.total_fail}")
    for c in results.contingencies:
        status = "PASS" if c.aeso_pass else "FAIL"
        print(
            f"    [{status}] {c.contingency_name[:40]:<40} "
            f"angle={c.max_rotor_angle_deg:.1f}°  "
            f"V_poi={c.min_poi_voltage_pu:.4f} pu"
        )
    print(sep)


def _print_pv_summary(scenario: str, results) -> None:
    sep = "-" * 56
    print(f"\n{sep}")
    print(f"  PV Voltage Stability | {scenario}")
    print(sep)
    print(f"  Curves run         : {len(results.curves)}")
    mc = results.most_critical_curve
    if mc:
        print(f"  Most critical      : {mc.scenario_name}")
        print(f"  Collapse at        : {mc.collapse_mw or 'Not reached'} MW")
        print(f"  Max stable         : {mc.max_stable_mw:.0f} MW")
    for c in results.curves:
        print(f"    [{c.aeso_status}] {c.scenario_name[:45]:<45} "
              f"collapse={c.collapse_mw or '—':>6}  "
              f"min_V={c.min_voltage_pu:.4f} pu")
    print(sep)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _find_sav(project_dir: str) -> str:
    """
    Look for a .sav file in project_dir or its cases/ subdirectory.
    Returns the first found path, or empty string.
    """
    for search_dir in (project_dir, os.path.join(project_dir, "cases")):
        if os.path.isdir(search_dir):
            for f in os.listdir(search_dir):
                if f.endswith(".sav"):
                    return os.path.join(search_dir, f)
    return ""


def _scenario_needs_any_study(project, scenario_name: str, study_types: list) -> bool:
    """Return True if the scenario requires at least one of the given study types."""
    matrix = project.get_study_matrix(scenario_name)
    if matrix is None:
        return False
    if "all" in study_types:
        return (
            matrix.power_flow_cat_a
            or matrix.power_flow_cat_b
            or matrix.transient_cat_a
            or matrix.transient_cat_b
            or matrix.transient_conditional
            or matrix.volt_stability_cat_a
            or matrix.volt_stability_cat_b
            or matrix.short_circuit_cat_a
        )
    checks = {
        "power_flow":        matrix.power_flow_cat_a or matrix.power_flow_cat_b,
        "transient":         matrix.transient_cat_a  or matrix.transient_cat_b
                             or matrix.transient_conditional,
        "voltage_stability": matrix.volt_stability_cat_a or matrix.volt_stability_cat_b,
        "short_circuit":     matrix.short_circuit_cat_a,
    }
    return any(checks.get(st, False) for st in study_types)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = build_parser()
    args   = parser.parse_args()

    # Decide mode
    use_gui = args.gui or (args.project is None)

    # ── GUI mode ──────────────────────────────────────────────────────────────
    if use_gui:
        try:
            from gui.main_window import launch
            launch()
            return 0
        except ImportError as exc:
            print(f"Could not launch GUI: {exc}")
            print("Ensure tkinter is available (Python built with Tk support).")
            return 1

    # ── CLI mode ──────────────────────────────────────────────────────────────

    # Import settings now (after path is set)
    from config.settings import PSSE_PATH, LOG_LEVEL, LOG_FILE, AESO

    # Setup logging
    log_file = LOG_FILE
    if args.project:
        log_file = os.path.join(args.project, "output", "automation.log")
    setup_logging(LOG_LEVEL, log_file)

    logger.info("=" * 60)
    logger.info("  AESO Interconnection Study Automation Tool")
    logger.info("  University of Calgary — ECE Capstone")
    logger.info("=" * 60)
    logger.info("Mode: CLI")
    logger.info("Project: %s", args.project)
    logger.info("Mock: %s", args.mock)

    # Validate project folder
    folder_warnings = validate_project_folder(args.project)
    if folder_warnings:
        for w in folder_warnings:
            logger.error(w)
        if any("[CRITICAL]" in w for w in folder_warnings):
            return 1

    # Validate environment (PSS/E paths)
    if not args.mock:
        env_warnings = validate_environment(PSSE_PATH)
        for w in env_warnings:
            logger.warning(w)
        if any("[CRITICAL]" in w for w in env_warnings):
            logger.error(
                "PSS/E environment check failed. "
                "Use --mock to run without PSS/E."
            )
            return 1

    # Parse study types
    if args.study.lower() == "all":
        study_types = ["all"]
    else:
        study_types = [s.strip() for s in args.study.split(",")]
        valid = {"power_flow", "short_circuit", "transient", "voltage_stability"}
        invalid = [s for s in study_types if s not in valid]
        if invalid:
            logger.error(
                "Unknown study type(s): %s\n"
                "Valid options: %s",
                invalid, sorted(valid)
            )
            return 1

    # Parse scenario filter
    scenario_filter = []
    if args.scenario:
        scenario_filter = [s.strip() for s in args.scenario.split(",")]

    # Validate-only mode
    if args.validate_only:
        from project_io.excel_reader import ExcelReader, ExcelReaderError
        excel_path = os.path.join(args.project, "study_scope_data.xlsx")
        try:
            project = ExcelReader(excel_path).read()
        except ExcelReaderError as exc:
            logger.error("Failed to load: %s", exc)
            return 1
        warnings = project.validate()
        if warnings:
            print(f"\nValidation warnings ({len(warnings)}):")
            for w in warnings:
                prefix = "  [CRITICAL]" if "[CRITICAL]" in w else "  [WARNING] "
                print(f"{prefix} {w.replace('[CRITICAL]','').replace('[WARNING]','').strip()}")
        else:
            print("\n✔ No issues found.")
        critical = [w for w in warnings if "[CRITICAL]" in w]
        return 1 if critical else 0

    # Run studies
    return run_cli(
        project_dir     = args.project,
        sav_override    = args.sav or "",
        study_types     = study_types,
        scenario_filter = scenario_filter,
        mock            = args.mock,
    )


if __name__ == "__main__":
    sys.exit(main())
