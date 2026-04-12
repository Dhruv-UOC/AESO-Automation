"""
validate_single_bus.py
----------------------
Run ALL four AESO studies (Power Flow, Short Circuit, Transient Stability,
PV Voltage Stability) for a SINGLE bus number and print a consolidated
pass/fail report to the console.

Purpose
-------
Use this script to validate that the automation results match your PSS/E
manual run for one specific bus before committing to a full batch run.

Bus Number Source
-----------------
The bus number is resolved in the following priority order so the script
stays integrated with the GUI workflow:

  1. Command-line argument:   python validate_single_bus.py 959
  2. GUI state file:          gui/.last_manual_bus  (written by the GUI
                               whenever the user types a bus number and
                               clicks "Run Manual Study")
  3. Interactive prompt:      if neither above is available, the script
                               asks for input.

All study parameters are imported directly from config/settings.py so
this file stays in sync automatically when thresholds change.

Usage
-----
  # With a bus number on the command line (no GUI needed):
  python validate_single_bus.py 959

  # Let the script read whatever bus number the GUI field last held:
  python validate_single_bus.py

  # With mock PSS/E (no licence required — for CI / offline testing):
  python validate_single_bus.py 959 --mock

Output
------
  • Colour-coded pass/fail table printed to stdout
  • Full traceback written to the Study Log (shared logging setup)
  • Exit code 0 if all studies pass, 1 if any fail or an error occurs
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from typing import List, Optional

# ── Ensure project root is on sys.path ────────────────────────────────────────
_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# ── Logging (mirrors gui/main_window.py setup) ────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── ANSI colours for console output (graceful fallback on Windows) ────────────
try:
    import colorama  # type: ignore
    colorama.init(autoreset=True)
    _GREEN  = colorama.Fore.GREEN  + colorama.Style.BRIGHT
    _RED    = colorama.Fore.RED    + colorama.Style.BRIGHT
    _CYAN   = colorama.Fore.CYAN   + colorama.Style.BRIGHT
    _RESET  = colorama.Style.RESET_ALL
    _MUTED  = colorama.Fore.WHITE  + colorama.Style.DIM
except ImportError:
    _GREEN = _RED = _CYAN = _RESET = _MUTED = ""


# ─────────────────────────────────────────────────────────────────────────────
# Bus-number resolution
# ─────────────────────────────────────────────────────────────────────────────

#: Path where gui/main_window.py saves the last manually entered bus number.
#: The GUI writes this file automatically when the user clicks "Run Manual Study".
_GUI_STATE_FILE = os.path.join(_ROOT, "gui", ".last_manual_bus")


def _resolve_bus_number(cli_arg: Optional[str]) -> int:
    """
    Return the bus number to validate, using the priority chain described in
    the module docstring.

    Parameters
    ----------
    cli_arg:
        Raw string from the command-line argument, or None if not supplied.

    Returns
    -------
    int
        Validated bus number.

    Raises
    ------
    SystemExit
        If the user provides an invalid (non-integer) value.
    """
    # 1. Command-line argument
    if cli_arg is not None:
        try:
            return int(cli_arg.strip())
        except ValueError:
            print(f"[ERROR] '{cli_arg}' is not a valid integer bus number.", file=sys.stderr)
            sys.exit(1)

    # 2. GUI state file (written by _run_manual_study in gui/main_window.py)
    if os.path.isfile(_GUI_STATE_FILE):
        try:
            with open(_GUI_STATE_FILE, "r", encoding="utf-8") as fh:
                raw = fh.read().strip()
            if raw:
                bus = int(raw)
                logger.info("Bus number read from GUI state file: %d", bus)
                return bus
        except (ValueError, OSError):
            pass  # fall through to interactive prompt

    # 3. Interactive prompt
    raw = input("Enter bus number to validate: ").strip()
    try:
        return int(raw)
    except ValueError:
        print(f"[ERROR] '{raw}' is not a valid integer bus number.", file=sys.stderr)
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# Printing helpers (mirror the colour tags in gui/main_window.py)
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(text: str) -> None:
    print(f"\n{_CYAN}{'─' * 50}{_RESET}")
    print(f"{_CYAN}  {text}{_RESET}")
    print(f"{_CYAN}{'─' * 50}{_RESET}")


def _row(line: str, tag: str = "plain") -> None:
    """Print a single result row with colour matching the GUI tags."""
    if tag == "pass":
        print(f"{_GREEN}{line}{_RESET}")
    elif tag == "fail":
        print(f"{_RED}{line}{_RESET}")
    elif tag == "muted":
        print(f"{_MUTED}{line}{_RESET}")
    else:
        print(line)


def _lines(pairs: List[tuple]) -> None:
    """Flush a list of (tag, text) pairs — same structure as the GUI thread."""
    for tag, text in pairs:
        _row(text.rstrip("\n"), tag)


# ─────────────────────────────────────────────────────────────────────────────
# Study runners — identical parameter sets to gui/main_window.py _run_manual_study
# ─────────────────────────────────────────────────────────────────────────────

def _run_short_circuit(psse, bus_no: int, sav: str, aeso: dict, out: List[tuple]) -> bool:
    """Run short-circuit study and append result rows to *out*. Returns True if pass."""
    from studies.short_circuit.short_circuit_study import ShortCircuitStudy

    study = ShortCircuitStudy(
        psse,
        scenario_label       = f"Validate  Bus {bus_no}",
        max_fault_current_ka = aeso["max_fault_current_ka"],
        bus_filter           = [bus_no],
    )
    results = study.run(sav)

    # Attribute resolution identical to gui/main_window.py Bug-4 fix
    bus_faults = [
        f for f in results.faults
        if getattr(f, "bus_number", getattr(f, "bus_no", None)) == bus_no
    ]

    out.append(("header", f"\n{'─' * 50}"))
    out.append(("header", f"  SHORT CIRCUIT  |  Bus {bus_no}"))
    out.append(("header", f"{'─' * 50}"))

    passed = True
    if not bus_faults:
        out.append(("muted", "  No fault data returned for this bus."))
        return True  # no data → not a failure

    out.append(("plain", f"  {'Fault':<6}  {'I_fault (kA)':>12}  {'Limit (kA)':>10}  Result"))
    out.append(("muted", f"  {'─' * 6}  {'─' * 12}  {'─' * 10}  {'─' * 6}"))
    for f in bus_faults:
        ok = f.fault_current_ka <= aeso["max_fault_current_ka"]
        if not ok:
            passed = False
        tag = "pass" if ok else "fail"
        status = "  PASS" if ok else "  FAIL"
        out.append(("plain",
            f"  {f.fault_type:<6}  "
            f"{f.fault_current_ka:>12.3f}  "
            f"{aeso['max_fault_current_ka']:>10.1f}"))
        out.append((tag, status))

    return passed


def _run_power_flow(psse, bus_no: int, sav: str, aeso: dict, project, out: List[tuple]) -> bool:
    """Run power-flow study and append result rows to *out*. Returns True if pass."""
    from studies.power_flow.power_flow_study import PowerFlowStudy

    out.append(("muted",
        f"\n  [Info] Running full-system power flow to extract Bus {bus_no} result…"))

    # All parameters identical to gui/main_window.py _run_manual_study
    study = PowerFlowStudy(
        psse,
        scenario_label          = f"Validate  Bus {bus_no}",
        project                 = project,
        season_label            = "Validate",
        voltage_min             = aeso["voltage_min_pu"],
        voltage_max             = aeso["voltage_max_pu"],
        voltage_min_contingency = aeso["voltage_min_contingency"],
        voltage_max_contingency = aeso["voltage_max_contingency"],
        thermal_limit_pct       = aeso["thermal_limit_pct"],
    )
    results = study.run(sav)

    out.append(("header", f"\n{'─' * 50}"))
    out.append(("header", f"  POWER FLOW  |  Bus {bus_no}"))
    out.append(("header", f"{'─' * 50}"))
    out.append(("plain", f"  Converged: {results.converged}"))

    # Bug-7 fix from gui/main_window.py: use 'buses', NOT 'bus_results'
    bus_results      = [b for b in results.buses      if b.bus_number == bus_no]
    violation_buses  = {v.bus_number for v in results.bus_violations}

    if not bus_results:
        out.append(("muted", f"  Bus {bus_no} not found in power flow results."))
        out.append(("muted", f"  (Tip: verify bus {bus_no} exists in the loaded .sav file.)"))
        return True  # bus absent is not a violation

    out.append(("plain",
        f"  {'V (pu)':>8}  {'Angle (°)':>10}  "
        f"{'V_min':>6}  {'V_max':>6}  Result"))
    out.append(("muted",
        f"  {'─' * 8}  {'─' * 10}  {'─' * 6}  {'─' * 6}  {'─' * 6}"))

    passed = True
    for b in bus_results:
        viol = b.bus_number in violation_buses
        if viol:
            passed = False
        tag = "fail" if viol else "pass"
        status = "  FAIL" if viol else "  PASS"
        out.append(("plain",
            f"  {b.voltage_pu:>8.4f}  "
            f"{b.angle_deg:>10.2f}  "
            f"{aeso['voltage_min_pu']:>6.3f}  "
            f"{aeso['voltage_max_pu']:>6.3f}"))
        out.append((tag, status))

    return passed


def _run_transient(psse, bus_no: int, sav: str, aeso: dict, project, out: List[tuple]) -> bool:
    """Run transient-stability study and append result rows to *out*. Returns True if pass."""
    from studies.transient_stability.transient_stability_study import TransientStabilityStudy

    study = TransientStabilityStudy(
        psse, project,
        scenario_label            = f"Validate  Bus {bus_no}",
        # Identical parameter names and sources as gui/main_window.py _run_manual_study
        rotor_angle_limit_deg     = aeso["rotor_angle_limit_deg"],
        voltage_recovery_pu       = aeso["voltage_recovery_pu"],
        voltage_recovery_window_s = aeso["voltage_recovery_time_s"],
    )
    results = study.run(sav)

    out.append(("header", f"\n{'─' * 50}"))
    out.append(("header", f"  TRANSIENT STABILITY  |  Bus {bus_no}"))
    out.append(("header", f"{'─' * 50}"))

    if not results.contingencies:
        out.append(("muted", "  No contingencies run."))
        return True

    out.append(("plain",
        f"  {'Contingency':<36}  {'Angle':>6}  {'V_poi':>6}  Result"))
    out.append(("muted",
        f"  {'─' * 36}  {'─' * 6}  {'─' * 6}  {'─' * 6}"))

    passed = True
    for c in results.contingencies:
        if not c.aeso_pass:
            passed = False
        tag = "pass" if c.aeso_pass else "fail"
        status = "  PASS" if c.aeso_pass else "  FAIL"
        out.append(("plain",
            f"  {c.contingency_name[:36]:<36}  "
            f"{c.max_rotor_angle_deg:>6.1f}  "
            f"{c.min_poi_voltage_pu:>6.4f}"))
        out.append((tag, status))

    return passed


def _run_pv_stability(psse, bus_no: int, sav: str, aeso: dict, project, out: List[tuple]) -> bool:
    """Run PV voltage-stability study and append result rows to *out*. Returns True if pass."""
    from studies.pv_voltage.pv_stability_study import PVStabilityStudy

    study = PVStabilityStudy(
        psse, project,
        scenario_label = f"Validate  Bus {bus_no}",
        # Identical parameter names and sources as gui/main_window.py _run_manual_study
        v_min_cat_a    = aeso["pv_cat_a_v_min"],
        v_min_cat_b    = aeso["pv_cat_b_v_min"],
    )
    results = study.run(sav)

    out.append(("header", f"\n{'─' * 50}"))
    out.append(("header", f"  PV VOLTAGE STABILITY  |  Bus {bus_no}"))
    out.append(("header", f"{'─' * 50}"))

    if not results.curves:
        out.append(("muted", "  No PV curves generated."))
        return True

    out.append(("plain",
        f"  {'Contingency':<36}  {'Collapse':>8}  {'Min V':>6}  Result"))
    out.append(("muted",
        f"  {'─' * 36}  {'─' * 8}  {'─' * 6}  {'─' * 6}"))

    passed = True
    for c in results.curves:
        if c.aeso_status != "PASS":
            passed = False
        tag = "pass" if c.aeso_status == "PASS" else "fail"
        collapse = f"{c.collapse_mw} MW" if c.collapse_mw else "    —"
        out.append(("plain",
            f"  {c.scenario_name[:36]:<36}  "
            f"{collapse:>8}  "
            f"{c.min_voltage_pu:>6.4f}"))
        out.append((tag, f"  {c.aeso_status}"))

    return passed


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> int:
    """
    Entry point.  Returns 0 on full pass, 1 on any failure or error.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Run all four AESO studies for a single bus and compare results "
            "against PSS/E manual output."
        )
    )
    parser.add_argument(
        "bus",
        nargs="?",
        default=None,
        metavar="BUS_NUMBER",
        help="Integer bus number to validate (e.g. 959). "
             "If omitted, the script reads from the GUI state file or prompts.",
    )
    parser.add_argument(
        "--mock",
        action="store_true",
        default=False,
        help="Run in Mock PSS/E mode (no licence required — for offline testing). "
             "Mirrors the 'Mock PSS/E mode' checkbox in the GUI.",
    )
    parser.add_argument(
        "--sav",
        default=None,
        metavar="PATH",
        help="Override the SAV file path.  Defaults to the path stored in the "
             "project's study_scope_data.xlsx (project_info.sav_file_path), "
             "falling back to config.settings.DEFAULT_SAV.",
    )
    parser.add_argument(
        "--excel",
        default=None,
        metavar="PATH",
        help="Path to study_scope_data.xlsx.  When omitted, the script looks "
             "for study_scope_data.xlsx next to the SAV file.",
    )
    args = parser.parse_args()

    # ── 1. Resolve bus number ─────────────────────────────────────────────────
    bus_no: int = _resolve_bus_number(args.bus)
    logger.info("Validating bus %d  (mock=%s)", bus_no, args.mock)

    # ── 2. Import config — identical to both GUI code paths ───────────────────
    from config.settings import PSSE_PATH, PSSE_VERSION, AESO, DEFAULT_SAV  # noqa: E402

    # ── 3. Resolve SAV file ───────────────────────────────────────────────────
    sav_path: str = args.sav or DEFAULT_SAV
    if not os.path.isfile(sav_path):
        print(
            f"[ERROR] SAV file not found: {sav_path}\n"
            "  Pass --sav <path> or check config/settings.py DEFAULT_SAV.",
            file=sys.stderr,
        )
        return 1

    # ── 4. Load project (Excel) — needed by PowerFlow, Transient, PV ─────────
    project = None
    excel_path = args.excel
    if excel_path is None:
        # Auto-discover next to the SAV file
        candidate = os.path.join(os.path.dirname(sav_path), "study_scope_data.xlsx")
        if os.path.isfile(candidate):
            excel_path = candidate
        else:
            # Walk one level up (SAV may live in cases/ sub-folder)
            candidate2 = os.path.join(
                os.path.dirname(os.path.dirname(sav_path)), "study_scope_data.xlsx"
            )
            if os.path.isfile(candidate2):
                excel_path = candidate2

    if excel_path and os.path.isfile(excel_path):
        try:
            from project_io.excel_reader import ExcelReader
            reader  = ExcelReader(excel_path)
            project = reader.read()
            project.project_dir = os.path.dirname(excel_path)
            project.output_dir  = os.path.join(project.project_dir, "output")
            logger.info("Loaded project: %s", excel_path)
        except Exception as exc:
            logger.warning("Could not load project Excel: %s — continuing without it.", exc)
    else:
        logger.warning(
            "study_scope_data.xlsx not found.  "
            "Power Flow, Transient, and PV studies may fail without project data."
        )

    # ── 5. Initialise PSS/E ───────────────────────────────────────────────────
    from core.psse_interface import PSSEInterface  # noqa: E402

    psse = PSSEInterface(
        psse_path    = PSSE_PATH,
        psse_version = PSSE_VERSION,
        mock         = args.mock,   # mirrors Bug-1 fix in gui/main_window.py
    )
    try:
        psse.initialize()
    except Exception as exc:
        print(
            f"\n[PSS/E INIT ERROR]  {exc}\n"
            "Check PSSE_PATH in config/settings.py, or pass --mock for offline testing.",
            file=sys.stderr,
        )
        return 1

    # ── 6. Run all four studies ───────────────────────────────────────────────
    overall_pass = True
    results_out: List[tuple] = []

    study_runners = [
        ("Short Circuit",       _run_short_circuit,  [psse, bus_no, sav_path, AESO]),
        ("Power Flow",          _run_power_flow,     [psse, bus_no, sav_path, AESO, project]),
        ("Transient Stability", _run_transient,      [psse, bus_no, sav_path, AESO, project]),
        ("PV Voltage Stability",_run_pv_stability,   [psse, bus_no, sav_path, AESO, project]),
    ]

    for study_name, runner_fn, runner_args in study_runners:
        logger.info("Running %s for bus %d …", study_name, bus_no)
        study_lines: List[tuple] = []
        try:
            passed = runner_fn(*runner_args, out=study_lines)
            if not passed:
                overall_pass = False
        except Exception as exc:
            logger.error("%s failed: %s", study_name, exc, exc_info=True)
            study_lines.append(("fail",  f"\n[{study_name.upper()} ERROR]  {exc}"))
            study_lines.append(("muted", "See console log for full traceback."))
            overall_pass = False

        results_out.extend(study_lines)

    # ── 7. Print consolidated results ─────────────────────────────────────────
    print()
    print(f"{_CYAN}{'=' * 50}{_RESET}")
    print(f"{_CYAN}  AESO VALIDATION REPORT — Bus {bus_no}{_RESET}")
    print(f"{_CYAN}  SAV: {os.path.basename(sav_path)}{_RESET}")
    print(f"{_CYAN}{'=' * 50}{_RESET}")

    _lines(results_out)

    print()
    print(f"{_CYAN}{'─' * 50}{_RESET}")
    if overall_pass:
        print(f"{_GREEN}  ✔  ALL STUDIES PASSED  —  Bus {bus_no}{_RESET}")
    else:
        print(f"{_RED}  ✘  ONE OR MORE STUDIES FAILED  —  Bus {bus_no}{_RESET}")
    print(f"{_CYAN}{'─' * 50}{_RESET}\n")

    return 0 if overall_pass else 1


if __name__ == "__main__":
    sys.exit(main())
