"""
tests/test_psse_live.py
-----------------------
Live PSS/E validation tests — compare automation output against manually
recorded golden reference values from a real PSS/E run on the project .sav.

These tests REQUIRE:
  1. A licensed PSS/E installation (psspy importable)
  2. The project .sav file at the path defined in SAV_PATH below
  3. A populated tests/golden_results.json
     Run  `python tests/extract_golden.py`  once to auto-generate it,
     then verify the values match your manual PSS/E output.

Skip automatically when PSS/E is not available:
    pytest tests/test_psse_live.py            # runs if psse importable
    pytest tests/test_mock.py                 # always runs (CI safe)

Usage
-----
    python -m pytest tests/test_psse_live.py -v
"""

import json
import os
import sys
import pytest

# ── Configuration ──────────────────────────────────────────────────────────────
# Adjust SAV_PATH to the absolute path of your .sav file.
SAV_PATH = os.environ.get(
    "AESO_SAV_PATH",
    r"C:\PSS_E_Models\AESO_Project.sav",  # <-- update this
)

GOLDEN_FILE = os.path.join(os.path.dirname(__file__), "golden_results.json")

# ── Tolerances ─────────────────────────────────────────────────────────────────
TOL_VOLTAGE_PU    = 0.005   # ±0.5 % on per-unit bus voltage
TOL_FAULT_KA      = 0.05    # ±50 A on fault current magnitude
TOL_ANGLE_DEG     = 1.0     # ±1 ° on fault current angle
TOL_ROTOR_DEG     = 2.0     # ±2 ° on peak rotor angle
TOL_RECOVERY_TIME = 0.001   # ±1 ms on voltage recovery time
TOL_COLLAPSE_MW   = 5.0     # ±5 MW on PV curve collapse point
TOL_MISMATCH_PCT  = 0.5     # max MW mismatch % for power flow energy balance

# ── pytest skip marker ─────────────────────────────────────────────────────────
def _psse_available() -> bool:
    """Return True only when psspy can actually be imported."""
    try:
        import psspy  # noqa: F401
        return True
    except ImportError:
        return False


psse_required = pytest.mark.skipif(
    not _psse_available(),
    reason="PSS/E not installed or not on PATH — skipping live tests",
)


# ── Fixtures ───────────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def golden():
    """Load golden reference values from JSON."""
    if not os.path.isfile(GOLDEN_FILE):
        pytest.skip(
            f"Golden results file not found: {GOLDEN_FILE}\n"
            "Run  python tests/extract_golden.py  to generate it."
        )
    with open(GOLDEN_FILE, "r") as fh:
        return json.load(fh)


@pytest.fixture(scope="session")
def live_psse():
    """Initialised PSS/E interface (real, not mock)."""
    from core.psse_interface import PSSEInterface
    from config.settings import PSSE_PATH, PSSE_VERSION
    psse = PSSEInterface(
        psse_path=PSSE_PATH,
        psse_version=PSSE_VERSION,
        mock=False,
    )
    psse.initialize()
    return psse


@pytest.fixture(scope="session")
def project_data():
    """Load ProjectData from the .sav-adjacent Excel workbook."""
    from project_io.project_data import ProjectData
    from config.settings import DEFAULT_EXCEL
    return ProjectData.from_excel(DEFAULT_EXCEL)


# ══════════════════════════════════════════════════════════════════════════════
# 1.  POWER FLOW
# ══════════════════════════════════════════════════════════════════════════════

@psse_required
class TestPowerFlow:
    """Validate power flow automation output against golden values."""

    @pytest.fixture(scope="class")
    def pf_results(self, live_psse, project_data):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        study = PowerFlowStudy(live_psse, project_data, scenario_label="Validation")
        return study.run(SAV_PATH)

    def test_converged(self, pf_results):
        """Power flow must converge."""
        assert pf_results.converged, "Power flow did not converge on .sav file"

    def test_energy_balance(self, pf_results):
        """Total MW generation must match total load + losses within 0.5 %."""
        gen  = pf_results.total_generation_mw
        load = pf_results.total_load_mw
        if gen > 0:
            mismatch_pct = abs(gen - load) / gen * 100.0
            assert mismatch_pct < TOL_MISMATCH_PCT, (
                f"MW energy balance mismatch too high: {mismatch_pct:.3f} %  "
                f"(gen={gen:.1f} MW, load={load:.1f} MW)"
            )

    def test_poi_bus_voltage(self, pf_results, golden):
        """POI bus voltage must be within ±0.005 pu of the golden value."""
        expected = golden["power_flow"]["poi_voltage_pu"]
        poi_bus_no = golden["power_flow"]["poi_bus_number"]
        actual = next(
            (b.voltage_pu for b in pf_results.buses if b.bus_number == poi_bus_no),
            None,
        )
        assert actual is not None, f"POI bus {poi_bus_no} not found in results"
        assert abs(actual - expected) <= TOL_VOLTAGE_PU, (
            f"POI voltage mismatch: got {actual:.5f} pu, expected {expected:.5f} pu "
            f"(tol ±{TOL_VOLTAGE_PU})"
        )

    def test_source_bus_voltage(self, pf_results, golden):
        """Source (generator) bus voltage must be within ±0.005 pu."""
        expected   = golden["power_flow"]["source_voltage_pu"]
        source_bus = golden["power_flow"]["source_bus_number"]
        actual = next(
            (b.voltage_pu for b in pf_results.buses if b.bus_number == source_bus),
            None,
        )
        assert actual is not None, f"Source bus {source_bus} not found in results"
        assert abs(actual - expected) <= TOL_VOLTAGE_PU, (
            f"Source bus voltage mismatch: got {actual:.5f} pu, "
            f"expected {expected:.5f} pu (tol ±{TOL_VOLTAGE_PU})"
        )

    def test_key_bus_voltages(self, pf_results, golden):
        """Spot-check up to 10 key buses against the golden table."""
        key_buses = golden["power_flow"].get("key_buses", [])
        bus_map = {b.bus_number: b.voltage_pu for b in pf_results.buses}
        for entry in key_buses:
            bno      = entry["bus_number"]
            expected = entry["voltage_pu"]
            actual   = bus_map.get(bno)
            assert actual is not None, f"Bus {bno} not found in power flow results"
            assert abs(actual - expected) <= TOL_VOLTAGE_PU, (
                f"Bus {bno} voltage: got {actual:.5f}, expected {expected:.5f} "
                f"(tol ±{TOL_VOLTAGE_PU})"
            )


# ══════════════════════════════════════════════════════════════════════════════
# 2.  SHORT CIRCUIT
# ══════════════════════════════════════════════════════════════════════════════

@psse_required
class TestShortCircuit:
    """Validate short circuit automation output against golden values."""

    @pytest.fixture(scope="class")
    def sc_results(self, live_psse, golden):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        # Only fault the buses listed in the golden file to keep runtime short
        bus_filter = [
            e["bus_number"]
            for e in golden["short_circuit"]["buses"]
        ]
        study = ShortCircuitStudy(
            live_psse,
            scenario_label="Validation",
            bus_filter=bus_filter if bus_filter else None,
        )
        return study.run(SAV_PATH)

    def test_fault_count(self, sc_results, golden):
        """Number of faulted buses must match the golden reference count."""
        expected_buses = len(golden["short_circuit"]["buses"])
        actual_buses   = len(set(f.bus_number for f in sc_results.faults))
        # 4 fault types × N buses
        assert actual_buses == expected_buses, (
            f"Expected {expected_buses} faulted buses, got {actual_buses}"
        )

    def test_3ph_fault_currents(self, sc_results, golden):
        """3-phase fault current at each golden bus within ±0.05 kA."""
        fault_map = {
            (f.bus_number, f.fault_type): f
            for f in sc_results.faults
        }
        for entry in golden["short_circuit"]["buses"]:
            bno      = entry["bus_number"]
            expected = entry["3ph_ka"]
            key      = (bno, "3PH")
            assert key in fault_map, (
                f"3PH fault result missing for bus {bno}"
            )
            actual = fault_map[key].fault_current_ka
            assert abs(actual - expected) <= TOL_FAULT_KA, (
                f"Bus {bno} 3PH fault current: got {actual:.4f} kA, "
                f"expected {expected:.4f} kA (tol ±{TOL_FAULT_KA})"
            )

    def test_lg_fault_currents(self, sc_results, golden):
        """Line-to-ground fault current at each golden bus within ±0.05 kA."""
        fault_map = {
            (f.bus_number, f.fault_type): f
            for f in sc_results.faults
        }
        for entry in golden["short_circuit"]["buses"]:
            bno      = entry["bus_number"]
            expected = entry.get("lg_ka")
            if expected is None:
                continue
            key    = (bno, "LG")
            assert key in fault_map, f"LG fault result missing for bus {bno}"
            actual = fault_map[key].fault_current_ka
            assert abs(actual - expected) <= TOL_FAULT_KA, (
                f"Bus {bno} LG fault current: got {actual:.4f} kA, "
                f"expected {expected:.4f} kA (tol ±{TOL_FAULT_KA})"
            )

    def test_no_new_violations(self, sc_results, golden):
        """Equipment withstand violations must not exceed golden count."""
        expected_count = golden["short_circuit"].get("violation_count", 0)
        actual_count   = len(sc_results.violations)
        assert actual_count <= expected_count, (
            f"More violations than expected: got {actual_count}, "
            f"golden baseline has {expected_count}"
        )

    def test_max_fault_bus(self, sc_results, golden):
        """The bus with the highest 3PH fault current must match the golden bus."""
        expected_bus = golden["short_circuit"].get("max_fault_bus_number")
        if expected_bus is None:
            pytest.skip("max_fault_bus_number not set in golden_results.json")
        three_ph = [
            f for f in sc_results.faults if f.fault_type == "3PH"
        ]
        if not three_ph:
            pytest.skip("No 3PH results in sc_results")
        actual_bus = max(three_ph, key=lambda f: f.fault_current_ka).bus_number
        assert actual_bus == expected_bus, (
            f"Highest 3PH fault bus: got {actual_bus}, expected {expected_bus}"
        )


# ══════════════════════════════════════════════════════════════════════════════
# 3.  TRANSIENT STABILITY
# ══════════════════════════════════════════════════════════════════════════════

@psse_required
class TestTransientStability:
    """Validate transient stability automation output against golden values."""

    @pytest.fixture(scope="class")
    def ts_results(self, live_psse, project_data):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        study = TransientStabilityStudy(
            live_psse,
            project_data,
            scenario_label="Validation",
        )
        return study.run(SAV_PATH)

    def test_all_contingencies_ran(self, ts_results, golden):
        """Every contingency listed in the golden file must appear in results."""
        expected_names = {
            c["contingency_name"]
            for c in golden["transient_stability"]["contingencies"]
        }
        actual_names = {c.contingency_name for c in ts_results.contingencies}
        missing = expected_names - actual_names
        assert not missing, (
            f"Missing contingency results: {missing}"
        )

    def test_aeso_pass_status(self, ts_results, golden):
        """Each contingency AESO pass/fail status must match the golden value."""
        result_map = {
            c.contingency_name: c for c in ts_results.contingencies
        }
        for entry in golden["transient_stability"]["contingencies"]:
            name     = entry["contingency_name"]
            expected = entry["aeso_pass"]
            cont     = result_map.get(name)
            if cont is None:
                continue
            assert cont.aeso_pass == expected, (
                f"Contingency '{name}' AESO status mismatch: "
                f"got {cont.aeso_pass}, expected {expected}"
            )

    def test_max_rotor_angles(self, ts_results, golden):
        """Peak rotor angle must be within ±2 ° of the golden value."""
        result_map = {
            c.contingency_name: c for c in ts_results.contingencies
        }
        for entry in golden["transient_stability"]["contingencies"]:
            name     = entry["contingency_name"]
            expected = entry.get("max_rotor_angle_deg")
            if expected is None:
                continue
            cont = result_map.get(name)
            if cont is None:
                continue
            assert abs(cont.max_rotor_angle_deg - expected) <= TOL_ROTOR_DEG, (
                f"Contingency '{name}' rotor angle: "
                f"got {cont.max_rotor_angle_deg:.2f} °, "
                f"expected {expected:.2f} ° (tol ±{TOL_ROTOR_DEG})"
            )

    def test_min_poi_voltages(self, ts_results, golden):
        """Minimum POI voltage during fault must be within ±0.005 pu."""
        result_map = {
            c.contingency_name: c for c in ts_results.contingencies
        }
        for entry in golden["transient_stability"]["contingencies"]:
            name     = entry["contingency_name"]
            expected = entry.get("min_poi_voltage_pu")
            if expected is None:
                continue
            cont = result_map.get(name)
            if cont is None:
                continue
            assert abs(cont.min_poi_voltage_pu - expected) <= TOL_VOLTAGE_PU, (
                f"Contingency '{name}' min POI voltage: "
                f"got {cont.min_poi_voltage_pu:.4f} pu, "
                f"expected {expected:.4f} pu (tol ±{TOL_VOLTAGE_PU})"
            )

    def test_voltage_recovery_aeso(self, ts_results):
        """
        AESO criterion: POI voltage must recover to ≥ 0.90 pu within 1.0 s
        post-fault for every contingency.
        """
        for cont in ts_results.contingencies:
            assert cont.voltage_recovered, (
                f"Contingency '{cont.contingency_name}': voltage did NOT recover "
                f"to ≥ 0.90 pu within 1.0 s (AESO Cat A criterion)"
            )

    def test_rotor_angle_stability_aeso(self, ts_results):
        """AESO criterion: no pole slip — rotor angle must stay below 180 °."""
        for cont in ts_results.contingencies:
            assert cont.rotor_angle_stable, (
                f"Contingency '{cont.contingency_name}': rotor angle instability detected "
                f"(max angle = {cont.max_rotor_angle_deg:.1f} °, limit = 180 °)"
            )


# ══════════════════════════════════════════════════════════════════════════════
# 4.  PV VOLTAGE STABILITY
# ══════════════════════════════════════════════════════════════════════════════

@psse_required
class TestPVStability:
    """Validate PV voltage stability output against golden reference values."""

    @pytest.fixture(scope="class")
    def pv_results(self, live_psse, project_data):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        study = PVStabilityStudy(
            live_psse,
            project_data,
            scenario_label="Validation",
        )
        return study.run(SAV_PATH)

    def test_n0_curve_converged(self, pv_results):
        """N-0 (Cat A) base case curve must converge."""
        n0_curve = next(
            (c for c in pv_results.curves if not c.is_contingency), None
        )
        assert n0_curve is not None, "N-0 PV curve not found in results"
        assert n0_curve.converged_base, (
            "N-0 PV base power flow did not converge"
        )

    def test_n0_collapse_mw(self, pv_results, golden):
        """N-0 collapse MW must be within ±5 MW of the golden value."""
        expected = golden["pv_stability"].get("n0_collapse_mw")
        if expected is None:
            pytest.skip("n0_collapse_mw not set in golden_results.json")
        n0_curve = next(
            (c for c in pv_results.curves if not c.is_contingency), None
        )
        assert n0_curve is not None
        actual = n0_curve.collapse_mw
        if actual is None:
            pytest.skip("N-0 collapse point not reached in this run")
        assert abs(actual - expected) <= TOL_COLLAPSE_MW, (
            f"N-0 collapse MW: got {actual:.1f}, expected {expected:.1f} "
            f"(tol ±{TOL_COLLAPSE_MW})"
        )

    def test_n0_no_aeso_violations(self, pv_results, golden):
        """N-0 curve must have zero AESO Cat A voltage violations at MARP."""
        marp = golden["pv_stability"].get("marp_mw", 0)
        n0_curve = next(
            (c for c in pv_results.curves if not c.is_contingency), None
        )
        if n0_curve is None or marp == 0:
            pytest.skip("N-0 curve or MARP not available")
        violations_at_marp = [
            p for p in n0_curve.violations if p.transfer_mw <= marp
        ]
        assert not violations_at_marp, (
            f"AESO Cat A violations at or below MARP ({marp} MW): "
            f"{[p.transfer_mw for p in violations_at_marp]} MW points violated"
        )

    def test_n1_collapse_mws(self, pv_results, golden):
        """Each N-1 contingency collapse MW within ±5 MW of golden values."""
        curve_map = {
            c.contingency_element: c
            for c in pv_results.curves
            if c.is_contingency
        }
        for entry in golden["pv_stability"].get("n1_contingencies", []):
            element  = entry["contingency_element"]
            expected = entry.get("collapse_mw")
            if expected is None:
                continue
            curve = curve_map.get(element)
            if curve is None:
                continue
            actual = curve.collapse_mw
            if actual is None:
                continue
            assert abs(actual - expected) <= TOL_COLLAPSE_MW, (
                f"N-1 '{element}' collapse MW: got {actual:.1f}, "
                f"expected {expected:.1f} (tol ±{TOL_COLLAPSE_MW})"
            )

    def test_n1_min_voltages(self, pv_results, golden):
        """N-1 minimum POI voltage at MARP must be within ±0.005 pu."""
        curve_map = {
            c.contingency_element: c
            for c in pv_results.curves
            if c.is_contingency
        }
        marp = golden["pv_stability"].get("marp_mw", 0)
        for entry in golden["pv_stability"].get("n1_contingencies", []):
            element  = entry["contingency_element"]
            expected = entry.get("min_voltage_pu")
            if expected is None or marp == 0:
                continue
            curve = curve_map.get(element)
            if curve is None:
                continue
            # Voltage at the point closest to MARP
            pts_at_marp = [
                p for p in curve.points
                if abs(p.transfer_mw - marp) <= 5
            ]
            if not pts_at_marp:
                continue
            actual = min(p.poi_voltage_pu for p in pts_at_marp)
            assert abs(actual - expected) <= TOL_VOLTAGE_PU, (
                f"N-1 '{element}' min voltage at MARP: "
                f"got {actual:.4f} pu, expected {expected:.4f} pu "
                f"(tol ±{TOL_VOLTAGE_PU})"
            )

    def test_most_critical_curve_identified(self, pv_results):
        """A most-critical curve must always be identified."""
        assert pv_results.most_critical_curve is not None, (
            "PV study did not identify a most-critical curve"
        )
