"""
tests/extract_golden.py
------------------------
One-shot script that runs every study on your real .sav file and writes
the results directly into tests/golden_results.json.

Run this ONCE after the automation code is confirmed correct (or after a
manual PSS/E cross-check), then commit the JSON as your regression baseline.

Usage
-----
    python tests/extract_golden.py

The script reads SAV_PATH from the environment variable AESO_SAV_PATH, or
falls back to the default hard-coded path below.  Set AESO_SAV_PATH before
running if your .sav lives elsewhere:

    # Windows PowerShell
    $env:AESO_SAV_PATH = "D:\\Models\\AESO_Project.sav"
    python tests/extract_golden.py

    # Windows CMD
    set AESO_SAV_PATH=D:\Models\AESO_Project.sav
    python tests/extract_golden.py

Requirements
------------
  - PSS/E installed and on sys.path
  - Project Excel workbook accessible (path from config/settings.py)
  - .sav file accessible
"""

import json
import os
import sys

# ── Configuration ──────────────────────────────────────────────────────────────
SAV_PATH = os.environ.get(
    "AESO_SAV_PATH",
    r"C:\PSS_E_Models\AESO_Project.sav",  # <-- update if AESO_SAV_PATH not set
)

OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "golden_results.json")

# ── Bootstrap path ─────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def _check_psse():
    try:
        import psspy  # noqa: F401
    except ImportError:
        print("[ERROR] psspy not importable. PSS/E must be installed and on PATH.")
        sys.exit(1)


def _load_project():
    from project_io.project_data import ProjectData
    from config.settings import DEFAULT_EXCEL
    print(f"[INFO] Loading project data from: {DEFAULT_EXCEL}")
    return ProjectData.from_excel(DEFAULT_EXCEL)


def _init_psse():
    from core.psse_interface import PSSEInterface
    from config.settings import PSSE_PATH, PSSE_VERSION
    print(f"[INFO] Initializing PSS/E at: {PSSE_PATH}")
    psse = PSSEInterface(psse_path=PSSE_PATH, psse_version=PSSE_VERSION, mock=False)
    psse.initialize()
    return psse


# ── Extractors ─────────────────────────────────────────────────────────────────

def extract_power_flow(psse, project) -> dict:
    from studies.power_flow.power_flow_study import PowerFlowStudy
    print("\n[1/4] Running Power Flow ...")
    study   = PowerFlowStudy(psse, project, scenario_label="GoldenExtract")
    results = study.run(SAV_PATH)

    if not results.converged:
        print("[WARN] Power flow did not converge — golden PF values will be null.")
        return {
            "poi_bus_number":    project.info.poi_bus_number,
            "poi_voltage_pu":    None,
            "source_bus_number": project.info.source_bus_number,
            "source_voltage_pu": None,
            "key_buses":         [],
        }

    bus_map = {b.bus_number: b for b in results.buses}

    poi_bus    = project.info.poi_bus_number
    source_bus = project.info.source_bus_number

    poi_v    = round(bus_map[poi_bus].voltage_pu,    5) if poi_bus    in bus_map else None
    source_v = round(bus_map[source_bus].voltage_pu, 5) if source_bus in bus_map else None

    # Record every bus that exists in the model (capped at 20 for readability)
    key_buses = [
        {
            "bus_number": b.bus_number,
            "bus_name":   b.bus_name,
            "voltage_pu": round(b.voltage_pu, 5),
        }
        for b in sorted(results.buses, key=lambda x: x.bus_number)[:20]
    ]

    print(f"  POI bus {poi_bus}: {poi_v} pu  |  Source bus {source_bus}: {source_v} pu")
    print(f"  {len(key_buses)} key buses recorded")

    return {
        "poi_bus_number":    poi_bus,
        "poi_voltage_pu":    poi_v,
        "source_bus_number": source_bus,
        "source_voltage_pu": source_v,
        "key_buses":         key_buses,
    }


def extract_short_circuit(psse, project) -> dict:
    from studies.short_circuit.short_circuit_study import ShortCircuitStudy
    print("\n[2/4] Running Short Circuit ...")
    study   = ShortCircuitStudy(psse, scenario_label="GoldenExtract")
    results = study.run(SAV_PATH)

    # Build per-bus dict keyed by bus_number
    bus_data: dict = {}
    for f in results.faults:
        bno = f.bus_number
        if bno not in bus_data:
            bus_data[bno] = {
                "bus_number": bno,
                "bus_name":   f.bus_name,
                "base_kv":    f.base_kv,
                "3ph_ka":     None,
                "lg_ka":      None,
            }
        if f.fault_type == "3PH":
            bus_data[bno]["3ph_ka"] = round(f.fault_current_ka, 4)
        elif f.fault_type == "LG":
            bus_data[bno]["lg_ka"]  = round(f.fault_current_ka, 4)

    # Identify the bus with the highest 3PH current
    three_ph = [f for f in results.faults if f.fault_type == "3PH"]
    max_bus  = max(three_ph, key=lambda f: f.fault_current_ka).bus_number if three_ph else None

    print(f"  {len(bus_data)} buses  |  {len(results.violations)} violations  "
          f"|  max 3PH bus: {max_bus}")

    return {
        "violation_count":      len(results.violations),
        "max_fault_bus_number": max_bus,
        "buses":                list(bus_data.values()),
    }


def extract_transient_stability(psse, project) -> dict:
    from studies.transient_stability.transient_stability_study import (
        TransientStabilityStudy,
    )
    print("\n[3/4] Running Transient Stability ...")
    study   = TransientStabilityStudy(
        psse, project, scenario_label="GoldenExtract"
    )
    results = study.run(SAV_PATH)

    contingencies = []
    for c in results.contingencies:
        contingencies.append({
            "contingency_name":    c.contingency_name,
            "aeso_pass":           c.aeso_pass,
            "max_rotor_angle_deg": round(c.max_rotor_angle_deg, 2),
            "min_poi_voltage_pu":  round(c.min_poi_voltage_pu,  5),
            "recovery_time_s":     round(c.recovery_time_s,     4),
        })
        status = "PASS" if c.aeso_pass else "FAIL"
        print(f"  [{status}] {c.contingency_name}  "
              f"max_angle={c.max_rotor_angle_deg:.1f}°  "
              f"min_V={c.min_poi_voltage_pu:.4f} pu")

    return {"contingencies": contingencies}


def extract_pv_stability(psse, project) -> dict:
    from studies.pv_voltage.pv_stability_study import PVStabilityStudy
    print("\n[4/4] Running PV Stability ...")
    study   = PVStabilityStudy(
        psse, project, scenario_label="GoldenExtract"
    )
    results = study.run(SAV_PATH)

    n0 = next((c for c in results.curves if not c.is_contingency), None)
    n0_collapse = round(n0.collapse_mw, 1) if (n0 and n0.collapse_mw) else None

    n1_entries = []
    for c in results.curves:
        if not c.is_contingency:
            continue
        n1_entries.append({
            "contingency_element": c.contingency_element,
            "collapse_mw":         round(c.collapse_mw, 1) if c.collapse_mw else None,
            "min_voltage_pu":      round(c.min_voltage_pu, 5),
        })
        print(f"  N-1 '{c.contingency_element}'  "
              f"collapse={c.collapse_mw or 'NR'} MW  "
              f"min_V={c.min_voltage_pu:.4f} pu")

    return {
        "marp_mw":          project.info.marp_mw,
        "n0_collapse_mw":   n0_collapse,
        "n1_contingencies": n1_entries,
    }


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 68)
    print(" AESO Automation — Golden Results Extractor")
    print(f" SAV:    {SAV_PATH}")
    print(f" Output: {OUTPUT_FILE}")
    print("=" * 68)

    _check_psse()

    if not os.path.isfile(SAV_PATH):
        print(f"[ERROR] .sav file not found: {SAV_PATH}")
        print("Set the AESO_SAV_PATH environment variable to the correct path.")
        sys.exit(1)

    psse    = _init_psse()
    project = _load_project()

    golden = {
        "_instructions": [
            "Auto-generated by tests/extract_golden.py — do NOT edit manually.",
            "Re-run  python tests/extract_golden.py  after any model change.",
            f"Source .sav: {SAV_PATH}"
        ],
        "power_flow":          extract_power_flow(psse, project),
        "short_circuit":       extract_short_circuit(psse, project),
        "transient_stability": extract_transient_stability(psse, project),
        "pv_stability":        extract_pv_stability(psse, project),
    }

    with open(OUTPUT_FILE, "w") as fh:
        json.dump(golden, fh, indent=2, default=str)

    print(f"\n{'=' * 68}")
    print(f" Golden results written to: {OUTPUT_FILE}")
    print(" Verify the values against your manual PSS/E output, then commit.")
    print("=" * 68)


if __name__ == "__main__":
    main()
