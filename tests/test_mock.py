"""
tests/test_mock.py
-------------------
Unit tests for all study modules using mock PSS/E mode.

Runs without a PSS/E licence — all PSS/E calls are replaced by
synthetic data in mock mode.

Run from the project root:
    python -m pytest tests/test_mock.py -v
    python -m pytest tests/test_mock.py -v --tb=short

Coverage
--------
  TestPSSEInterface          — initialize, load_case, mock guards
  TestProjectData            — validate(), convenience accessors
  TestExcelReaderWriter      — round-trip write/read of ProjectData
  TestPowerFlowStudy         — run, violations, save_results
  TestShortCircuitStudy      — all four fault types, violations, compare
  TestTransientStabilityStudy— run, mock channel data, AESO criteria
  TestPVStabilityStudy       — N-0 and N-1 curves, collapse detection
  TestPlotter                — smoke tests for all plot functions
"""

import os
import sys
import tempfile

import pytest

# ── Ensure project root is on path ───────────────────────────────────────────
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.psse_interface import PSSEInterface
from project_io.project_data import (
    BusNumberEntry,
    GeneratorDispatch,
    IntertieFow,
    ProjectData,
    ProjectInfo,
    PVContingency,
    RenewableDispatch,
    SCSubstation,
    Scenario,
    StudyMatrixEntry,
    TSContingency,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────

@pytest.fixture
def mock_psse():
    """Initialized mock PSS/E interface."""
    psse = PSSEInterface(mock=True)
    psse.initialize()
    return psse


@pytest.fixture
def dummy_sav(tmp_path):
    """A dummy .sav file (content irrelevant in mock mode)."""
    sav = tmp_path / "test_case.sav"
    sav.write_bytes(b"PSS/E dummy SAV")
    return str(sav)


@pytest.fixture
def minimal_project(tmp_path):
    """A minimal ProjectData with enough fields to run all studies."""
    p = ProjectData()

    p.info = ProjectInfo(
        project_number       = "TEST001",
        project_name         = "Unit Test Project",
        marp_mw              = 400.0,
        max_capability_mw    = 400.0,
        connection_voltage_kv= 240.0,
        poc_substation_name  = "Test Substation",
        sav_file_path        = str(tmp_path / "test.sav"),
        source_bus_number    = 3011,
        poi_bus_number       = 153,
        ts_fault_bus_number  = 153,
        machine_id           = "1",
    )

    p.scenarios = [
        Scenario(1, 2028, "SP", "HG", "2028 SP Pre-Project",  "Pre",  0, 0),
        Scenario(4, 2028, "SP", "HG", "2028 SP Post-Project", "Post", 0, 400),
    ]

    p.study_matrix = [
        StudyMatrixEntry(
            scenario_name        = "2028 SP Pre-Project",
            power_flow_cat_a     = True,
            power_flow_cat_b     = True,
            transient_cat_a      = True,
            transient_cat_b      = True,
            short_circuit_cat_a  = True,
        ),
        StudyMatrixEntry(
            scenario_name        = "2028 SP Post-Project",
            power_flow_cat_a     = True,
            power_flow_cat_b     = True,
            volt_stability_cat_a = True,
            volt_stability_cat_b = True,
            transient_cat_a      = True,
            transient_cat_b      = True,
            short_circuit_cat_a  = True,
        ),
    ]

    p.conv_gen = [
        GeneratorDispatch(
            facility_name = "Battle River #4",
            unit_no       = "4",
            bus_no        = 1496,
            mc_mw         = 155.0,
            area_no       = 36,
            dispatch_mw   = {"2028 SP": 31.0, "2028 SL": 31.0},
        ),
    ]

    p.renewables = [
        RenewableDispatch(
            facility_name = "Garden Plain Wind",
            gen_type      = "Wind",
            bus_no        = 565002,
            mc_mw         = 130.0,
            area_no       = 42,
            dispatch_mw   = {"2028 SP": 130.0, "2028 SL": 85.3},
        ),
    ]

    p.intertie_flows = [
        IntertieFow(
            scenario_name = "2028 SP Post-Project",
            flows         = {"AB-BC": 851.0, "AB-SK": 150.0, "MATL": 186.0},
        ),
    ]

    p.ts_contingencies = [
        TSContingency(
            contingency_name  = "9L24 Oakland-Lanfine",
            from_bus_name     = "Oakland 946S",
            to_bus_name       = "Lanfine 959S",
            from_bus_no       = 151,
            to_bus_no         = 152,
            circuit_id        = "1",
            fault_location    = "Oakland 946S",
            near_end_cycles   = 5.0,
            far_end_cycles    = 6.0,
        ),
    ]

    p.sc_substations = [
        SCSubstation("Lanfine 959S",  bus_no=152, notes="POI"),
        SCSubstation("Oakland 946S",  bus_no=151, notes=""),
        SCSubstation("Oyen 767S",     bus_no=None, notes="Bus No TBD"),
    ]

    p.pv_contingencies = [
        PVContingency(
            contingency_name = "N-1: 9L24 Oakland-Lanfine",
            from_bus_name    = "Oakland 946S",
            to_bus_name      = "Lanfine 959S",
            from_bus_no      = 151,
            to_bus_no        = 152,
            circuit_id       = "1",
            category         = "B",
        ),
    ]

    p.bus_numbers = [
        BusNumberEntry("Lanfine 959S", 152, 240.0, 1, "POI bus"),
        BusNumberEntry("Oakland 946S", 151, 240.0, 1, ""),
    ]

    p.project_dir = str(tmp_path)
    p.output_dir  = str(tmp_path / "output")
    return p


# ═════════════════════════════════════════════════════════════════════════════
# PSSEInterface tests
# ═════════════════════════════════════════════════════════════════════════════

class TestPSSEInterface:

    def test_initialize_mock(self, mock_psse):
        assert mock_psse._initialized is True
        assert mock_psse.mock is True

    def test_load_case_mock_no_file_check(self, mock_psse, tmp_path):
        """Mock mode must NOT raise FileNotFoundError on non-existent paths."""
        mock_psse.load_case("nonexistent_path/case.sav")   # must not raise

    def test_load_case_real_file_mock(self, mock_psse, dummy_sav):
        """Mock mode accepts a real file path without error."""
        mock_psse.load_case(dummy_sav)   # must not raise

    def test_psspy_raises_in_mock(self, mock_psse):
        with pytest.raises(RuntimeError, match="mock mode"):
            _ = mock_psse.psspy

    def test_check_initialized_raises_before_init(self):
        psse = PSSEInterface(mock=True)
        with pytest.raises(RuntimeError, match="initialize"):
            psse.load_case("any.sav")

    def test_save_case_mock(self, mock_psse, tmp_path):
        mock_psse.save_case(str(tmp_path / "output.sav"))   # must not raise

    def test_load_dynamics_mock(self, mock_psse, tmp_path):
        mock_psse.load_dynamics("nonexistent.dyr")   # mock: no-op

    def test_supported_versions(self):
        assert 35 in PSSEInterface.SUPPORTED_VERSIONS


# ═════════════════════════════════════════════════════════════════════════════
# ProjectData tests
# ═════════════════════════════════════════════════════════════════════════════

class TestProjectData:

    def test_get_scenario_found(self, minimal_project):
        sc = minimal_project.get_scenario("2028 SP Post-Project")
        assert sc is not None
        assert sc.scenario_no == 4
        assert sc.project_gen_mw == 400.0

    def test_get_scenario_not_found(self, minimal_project):
        sc = minimal_project.get_scenario("Nonexistent Scenario")
        assert sc is None

    def test_get_study_matrix_found(self, minimal_project):
        m = minimal_project.get_study_matrix("2028 SP Post-Project")
        assert m is not None
        assert m.volt_stability_cat_a is True
        assert m.power_flow_cat_a is True

    def test_scenarios_requiring_power_flow(self, minimal_project):
        scenarios = minimal_project.scenarios_requiring_study("power_flow")
        assert len(scenarios) == 2

    def test_scenarios_requiring_voltage_stability(self, minimal_project):
        scenarios = minimal_project.scenarios_requiring_study("voltage_stability")
        names = [s.scenario_name for s in scenarios]
        assert "2028 SP Post-Project" in names
        assert "2028 SP Pre-Project"  not in names

    def test_scenarios_requiring_transient(self, minimal_project):
        scenarios = minimal_project.scenarios_requiring_study("transient")
        assert len(scenarios) == 2

    def test_scenarios_requiring_short_circuit(self, minimal_project):
        scenarios = minimal_project.scenarios_requiring_study("short_circuit")
        assert len(scenarios) == 2

    def test_get_dispatch_for_season(self, minimal_project):
        conv, ren = minimal_project.get_dispatch_for_season("2028 SP")
        assert len(conv) == 1
        assert len(ren)  == 1
        assert conv[0].dispatch_mw["2028 SP"] == 31.0
        assert ren[0].dispatch_mw["2028 SP"]  == 130.0

    def test_get_dispatch_missing_season(self, minimal_project):
        conv, ren = minimal_project.get_dispatch_for_season("2099 XY")
        assert conv == []
        assert ren  == []

    def test_get_intertie_flows(self, minimal_project):
        flows = minimal_project.get_intertie_flows("2028 SP Post-Project")
        assert flows["AB-BC"] == 851.0
        assert flows["MATL"]  == 186.0

    def test_get_intertie_flows_missing(self, minimal_project):
        flows = minimal_project.get_intertie_flows("Nonexistent")
        assert flows == {}

    def test_get_bus_number_found(self, minimal_project):
        bnum = minimal_project.get_bus_number("Lanfine 959S")
        assert bnum == 152

    def test_get_bus_number_case_insensitive(self, minimal_project):
        bnum = minimal_project.get_bus_number("lanfine 959s")
        assert bnum == 152

    def test_get_bus_number_not_found(self, minimal_project):
        bnum = minimal_project.get_bus_number("Nonexistent 999S")
        assert bnum is None

    def test_season_labels(self, minimal_project):
        labels = minimal_project.season_labels()
        assert "2028 SP" in labels
        assert "2028 SL" in labels

    def test_ts_contingency_seconds_conversion(self, minimal_project):
        cont = minimal_project.ts_contingencies[0]
        assert cont.near_end_cycles   == 5.0
        assert cont.near_end_seconds  == round(5.0 / 60.0, 6)
        assert cont.far_end_seconds   == round(6.0 / 60.0, 6)

    def test_validate_no_critical(self, minimal_project, tmp_path):
        # Create the dummy sav file so validate passes the file check
        sav = tmp_path / "test.sav"
        sav.write_bytes(b"dummy")
        minimal_project.info.sav_file_path = str(sav)
        warnings = minimal_project.validate()
        critical = [w for w in warnings if "[CRITICAL]" in w]
        assert critical == [], f"Unexpected critical warnings: {critical}"

    def test_validate_missing_sav(self, minimal_project):
        minimal_project.info.sav_file_path = "/nonexistent/path/case.sav"
        warnings = minimal_project.validate()
        critical = [w for w in warnings if "[CRITICAL]" in w]
        assert any("SAV file not found" in w for w in critical)

    def test_validate_missing_poi_bus(self, minimal_project, tmp_path):
        sav = tmp_path / "test.sav"
        sav.write_bytes(b"dummy")
        minimal_project.info.sav_file_path = str(sav)
        minimal_project.info.poi_bus_number = None
        warnings = minimal_project.validate()
        assert any("POI Bus Number" in w for w in warnings)

    def test_str_representation(self, minimal_project):
        s = str(minimal_project)
        assert "TEST001" in s
        assert "scenarios=2" in s


# ═════════════════════════════════════════════════════════════════════════════
# Excel Reader / Writer round-trip
# ═════════════════════════════════════════════════════════════════════════════

class TestExcelReaderWriter:

    def test_new_project_file_created(self, tmp_path):
        """new_project_file() copies the template."""
        from project_io.excel_writer import ExcelWriter

        template = os.path.join(_ROOT, "templates", "study_scope_template.xlsx")
        if not os.path.isfile(template):
            pytest.skip("Template file not found — run from project root.")

        dest = str(tmp_path / "study_scope_data.xlsx")
        ExcelWriter.new_project_file(template, dest)
        assert os.path.isfile(dest)

    def test_new_project_file_no_overwrite(self, tmp_path):
        """new_project_file() raises if destination exists and overwrite=False."""
        from project_io.excel_writer import ExcelWriter

        template = os.path.join(_ROOT, "templates", "study_scope_template.xlsx")
        if not os.path.isfile(template):
            pytest.skip("Template file not found.")

        dest = str(tmp_path / "study_scope_data.xlsx")
        ExcelWriter.new_project_file(template, dest)
        with pytest.raises(FileExistsError):
            ExcelWriter.new_project_file(template, dest, overwrite=False)

    def test_write_read_roundtrip(self, tmp_path, minimal_project):
        """Write ProjectData to Excel, read it back, check key fields."""
        from project_io.excel_writer import ExcelWriter
        from project_io.excel_reader import ExcelReader

        template = os.path.join(_ROOT, "templates", "study_scope_template.xlsx")
        if not os.path.isfile(template):
            pytest.skip("Template file not found.")

        dest = str(tmp_path / "study_scope_data.xlsx")
        ExcelWriter.new_project_file(template, dest)
        ExcelWriter(dest).write(minimal_project)

        read_back = ExcelReader(dest).read()

        assert read_back.info.project_number == "TEST001"
        assert read_back.info.marp_mw        == 400.0
        assert len(read_back.scenarios)       == 2
        assert len(read_back.study_matrix)    == 2
        assert len(read_back.conv_gen)        == 1
        assert len(read_back.renewables)      == 1
        assert len(read_back.ts_contingencies)== 1
        assert len(read_back.sc_substations)  == 3
        assert len(read_back.pv_contingencies) == 1
        assert len(read_back.bus_numbers)     == 2

    def test_reader_missing_file(self, tmp_path):
        from project_io.excel_reader import ExcelReader, ExcelReaderError
        with pytest.raises(ExcelReaderError, match="not found"):
            ExcelReader(str(tmp_path / "nonexistent.xlsx")).read()


# ═════════════════════════════════════════════════════════════════════════════
# PowerFlowStudy tests
# ═════════════════════════════════════════════════════════════════════════════

class TestPowerFlowStudy:

    def test_run_returns_results(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        study   = PowerFlowStudy(mock_psse, scenario_label="Test")
        results = study.run(dummy_sav)
        assert results is not None
        assert results.scenario   == "Test"
        assert results.converged  is True

    def test_mock_bus_voltages_reasonable(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        for bus in results.buses:
            assert 0.5 < bus.voltage_pu < 1.5, (
                f"Unreasonable voltage at bus {bus.bus_number}: {bus.voltage_pu}"
            )

    def test_violations_flagged_correctly(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        study   = PowerFlowStudy(mock_psse, voltage_min=0.95, voltage_max=1.05)
        results = study.run(dummy_sav)
        for bus in results.bus_violations:
            assert bus.violation is True
            assert bus.voltage_pu < 0.95 or bus.voltage_pu > 1.05

    def test_branch_violations_flagged(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse, thermal_limit_pct=100.0).run(dummy_sav)
        for br in results.branch_violations:
            assert br.loading_pct > 100.0
            assert br.violation is True

    def test_contingencies_returned(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        assert len(results.contingencies) > 0

    def test_system_totals_positive(self, mock_psse, dummy_sav):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        assert results.total_generation_mw > 0
        assert results.total_load_mw       > 0

    def test_save_results_creates_files(self, mock_psse, dummy_sav, tmp_path):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        study = PowerFlowStudy(mock_psse, scenario_label="SaveTest")
        study.run(dummy_sav)
        files = study.save_results(
            str(tmp_path / "results"),
            str(tmp_path / "plots"),
            str(tmp_path / "reports"),
        )
        assert os.path.isfile(files["excel"]), "Excel not created"
        assert os.path.isfile(files["pdf"]),   "PDF not created"
        assert len(files["plots"]) > 0

    def test_save_without_run_raises(self, mock_psse, tmp_path):
        from studies.power_flow.power_flow_study import PowerFlowStudy
        study = PowerFlowStudy(mock_psse)
        with pytest.raises(RuntimeError, match="run()"):
            study.save_results(
                str(tmp_path), str(tmp_path), str(tmp_path)
            )


# ═════════════════════════════════════════════════════════════════════════════
# ShortCircuitStudy tests
# ═════════════════════════════════════════════════════════════════════════════

class TestShortCircuitStudy:

    def test_run_all_four_fault_types(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        results = ShortCircuitStudy(mock_psse, scenario_label="SC_Test").run(dummy_sav)
        types_found = {f.fault_type for f in results.faults}
        assert "3PH" in types_found
        assert "LG"  in types_found
        assert "LL"  in types_found
        assert "LLG" in types_found

    def test_faults_sorted_by_severity(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        results = ShortCircuitStudy(mock_psse).run(dummy_sav)
        currents = [f.fault_current_ka for f in results.faults]
        assert currents == sorted(currents, reverse=True), (
            "Faults not sorted by severity (descending current)"
        )

    def test_violations_above_limit(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        study   = ShortCircuitStudy(mock_psse, max_fault_current_ka=50.0)
        results = study.run(dummy_sav)
        for f in results.violations:
            assert f.fault_current_ka > 50.0
            assert f.violation is True

    def test_max_fault_tracked(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        results = ShortCircuitStudy(mock_psse).run(dummy_sav)
        assert results.max_fault_current_ka == results.faults[0].fault_current_ka
        assert results.max_fault_bus        == results.faults[0].bus_name

    def test_fault_current_angle_present(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        results = ShortCircuitStudy(mock_psse).run(dummy_sav)
        for f in results.faults:
            assert isinstance(f.fault_current_ang, float)

    def test_compare_pre_post(self, mock_psse, dummy_sav):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        pre  = ShortCircuitStudy(mock_psse, scenario_label="Pre").run(dummy_sav)
        post = ShortCircuitStudy(mock_psse, scenario_label="Post").run(dummy_sav)
        df   = ShortCircuitStudy.compare(pre, post)
        assert not df.empty
        assert "Delta (kA)" in df.columns
        assert "Bus Number"  in df.columns

    def test_save_results_creates_files(self, mock_psse, dummy_sav, tmp_path):
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        study = ShortCircuitStudy(mock_psse, scenario_label="SC_Save")
        study.run(dummy_sav)
        files = study.save_results(
            str(tmp_path / "results"),
            str(tmp_path / "plots"),
            str(tmp_path / "reports"),
        )
        assert os.path.isfile(files["excel"])
        assert os.path.isfile(files["pdf"])


# ═════════════════════════════════════════════════════════════════════════════
# TransientStabilityStudy tests
# ═════════════════════════════════════════════════════════════════════════════

class TestTransientStabilityStudy:

    def test_run_returns_results(self, mock_psse, dummy_sav, minimal_project):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        study   = TransientStabilityStudy(mock_psse, minimal_project, "TS_Test")
        results = study.run(dummy_sav)
        assert results is not None
        assert results.scenario_label == "TS_Test"
        assert len(results.contingencies) == len(minimal_project.ts_contingencies)

    def test_contingency_has_channels(self, mock_psse, dummy_sav, minimal_project):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        results = TransientStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        for cont in results.contingencies:
            assert len(cont.channels) > 0, "No channels in contingency result"
            for name, ch in cont.channels.items():
                assert len(ch.time_s) == len(ch.values), (
                    f"Time and value arrays differ for channel '{name}'"
                )

    def test_fault_clearing_time_from_contingency(
        self, mock_psse, dummy_sav, minimal_project
    ):
        """Fault clearing time must use per-contingency cycles, not a global value."""
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        study   = TransientStabilityStudy(
            mock_psse, minimal_project,
            fault_apply_time_s=1.0
        )
        results = study.run(dummy_sav)
        cont    = results.contingencies[0]
        expected_clear_time = round(
            1.0 + minimal_project.ts_contingencies[0].near_end_seconds, 6
        )
        assert abs(cont.fault_clear_time_s - expected_clear_time) < 1e-5

    def test_rotor_angle_checked(self, mock_psse, dummy_sav, minimal_project):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        results = TransientStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        for cont in results.contingencies:
            assert cont.max_rotor_angle_deg >= 0.0

    def test_mock_voltage_dip_during_fault(
        self, mock_psse, dummy_sav, minimal_project
    ):
        """Mock voltage must dip below 0.5 pu during the fault window."""
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        study   = TransientStabilityStudy(
            mock_psse, minimal_project,
            fault_apply_time_s=1.0,
        )
        results = study.run(dummy_sav)
        cont    = results.contingencies[0]
        v_ch    = cont.channels.get("Bus Voltage POI")
        if v_ch:
            fa  = cont.fault_apply_time_s
            fc  = cont.fault_clear_time_s
            fault_voltages = [
                v for t, v in zip(v_ch.time_s, v_ch.values)
                if fa <= t <= fc
            ]
            if fault_voltages:
                assert min(fault_voltages) < 0.5, (
                    "Voltage should dip during fault in mock mode"
                )

    def test_save_results_creates_files(
        self, mock_psse, dummy_sav, minimal_project, tmp_path
    ):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        study = TransientStabilityStudy(mock_psse, minimal_project, "TS_Save")
        study.run(dummy_sav)
        files = study.save_results(
            str(tmp_path / "results"),
            str(tmp_path / "plots"),
            str(tmp_path / "reports"),
        )
        assert os.path.isfile(files["excel"])
        assert os.path.isfile(files["pdf"])

    def test_pass_fail_counts(self, mock_psse, dummy_sav, minimal_project):
        from studies.transient_stability.transient_stability_study import (
            TransientStabilityStudy,
        )
        results = TransientStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        total = results.total_pass + results.total_fail
        assert total == len(results.contingencies)


# ═════════════════════════════════════════════════════════════════════════════
# PVStabilityStudy tests
# ═════════════════════════════════════════════════════════════════════════════

class TestPVStabilityStudy:

    def test_run_returns_results(self, mock_psse, dummy_sav, minimal_project):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        study   = PVStabilityStudy(mock_psse, minimal_project, "PV_Test")
        results = study.run(dummy_sav)
        assert results is not None
        assert results.scenario_label == "PV_Test"

    def test_n0_and_n1_curves_generated(
        self, mock_psse, dummy_sav, minimal_project
    ):
        """N-0 plus one N-1 contingency = 2 curves."""
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        n_curves = len(results.curves)
        expected = 1 + len(minimal_project.pv_contingencies)
        assert n_curves == expected, (
            f"Expected {expected} curves, got {n_curves}"
        )

    def test_n0_curve_category_a(self, mock_psse, dummy_sav, minimal_project):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        n0      = results.curves[0]
        assert n0.category       == "A"
        assert n0.is_contingency is False

    def test_n1_curve_category_b(self, mock_psse, dummy_sav, minimal_project):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        n1      = results.curves[1]
        assert n1.category       == "B"
        assert n1.is_contingency is True

    def test_curves_have_points(self, mock_psse, dummy_sav, minimal_project):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        for curve in results.curves:
            assert len(curve.points) > 0, (
                f"No points in curve: {curve.scenario_name}"
            )

    def test_voltage_decreases_with_transfer(
        self, mock_psse, dummy_sav, minimal_project
    ):
        """POI voltage should generally decrease as transfer increases."""
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        for curve in results.curves:
            if len(curve.points) > 5:
                first_v = curve.points[0].poi_voltage_pu
                last_v  = curve.points[-1].poi_voltage_pu
                assert first_v > last_v, (
                    f"Voltage should decrease along PV curve: {curve.scenario_name}"
                )

    def test_most_critical_is_n1(self, mock_psse, dummy_sav, minimal_project):
        """N-1 contingency should produce lower collapse MW than N-0."""
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        results = PVStabilityStudy(mock_psse, minimal_project).run(dummy_sav)
        mc      = results.most_critical_curve
        assert mc is not None
        # Most critical should be the N-1 curve (lower collapse MW)
        if mc.collapse_mw:
            n0_collapse = results.curves[0].collapse_mw
            if n0_collapse:
                assert mc.collapse_mw <= n0_collapse

    def test_transfer_end_from_marp(self, mock_psse, dummy_sav, minimal_project):
        """transfer_end defaults to MARP * 1.05."""
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        study = PVStabilityStudy(mock_psse, minimal_project)
        expected = round(minimal_project.info.marp_mw * 1.05, 0)
        assert study.transfer_end == expected

    def test_save_results_creates_files(
        self, mock_psse, dummy_sav, minimal_project, tmp_path
    ):
        from studies.pv_voltage.pv_stability_study import PVStabilityStudy
        study = PVStabilityStudy(mock_psse, minimal_project, "PV_Save")
        study.run(dummy_sav)
        files = study.save_results(
            str(tmp_path / "results"),
            str(tmp_path / "plots"),
            str(tmp_path / "reports"),
        )
        assert os.path.isfile(files["excel"])
        assert os.path.isfile(files["pdf"])
        assert os.path.isdir(files["csv_dir"])


# ═════════════════════════════════════════════════════════════════════════════
# Plotter smoke tests
# ═════════════════════════════════════════════════════════════════════════════

class TestPlotter:
    """Smoke tests — verify plots are created without errors."""

    def _voltage_df(self, mock_psse, dummy_sav):
        import pandas as pd
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        return pd.DataFrame([{
            "Bus Name":     b.bus_name,
            "Voltage (pu)": b.voltage_pu,
            "Violation":    b.violation,
        } for b in results.buses])

    def _branch_df(self, mock_psse, dummy_sav):
        import pandas as pd
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        return pd.DataFrame([{
            "From Bus":          b.from_bus,
            "To Bus":            b.to_bus,
            "Circuit ID":        b.circuit_id,
            "Loading %":         b.loading_pct,
            "Thermal Violation": b.violation,
        } for b in results.branches])

    def test_voltage_profile(self, mock_psse, dummy_sav, tmp_path):
        from reporting.plotter import plot_voltage_profile
        df   = self._voltage_df(mock_psse, dummy_sav)
        path = plot_voltage_profile(df, "Test", str(tmp_path))
        assert os.path.isfile(path)

    def test_thermal_loading(self, mock_psse, dummy_sav, tmp_path):
        from reporting.plotter import plot_thermal_loading
        df   = self._branch_df(mock_psse, dummy_sav)
        path = plot_thermal_loading(df, "Test", str(tmp_path))
        assert os.path.isfile(path)

    def test_contingency_summary(self, mock_psse, dummy_sav, tmp_path):
        import pandas as pd
        from reporting.plotter import plot_contingency_summary
        from studies.power_flow.power_flow_study import PowerFlowStudy
        results = PowerFlowStudy(mock_psse).run(dummy_sav)
        df = pd.DataFrame([{
            "Min Voltage (pu)":   c.min_voltage_pu,
            "Bus Violations":     len(c.bus_violations),
            "Branch Violations":  len(c.branch_violations),
        } for c in results.contingencies])
        if not df.empty:
            path = plot_contingency_summary(df, "Test", str(tmp_path))
            assert path == "" or os.path.isfile(path)

    def test_fault_currents(self, mock_psse, dummy_sav, tmp_path):
        import pandas as pd
        from reporting.plotter import plot_fault_currents
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        results = ShortCircuitStudy(mock_psse).run(dummy_sav)
        df = pd.DataFrame([{
            "Severity Rank":      f.severity_rank,
            "Bus Name":           f.bus_name,
            "Fault Type":         f.fault_type,
            "Fault Current (kA)": f.fault_current_ka,
            "Violation":          f.violation,
        } for f in results.faults])
        path = plot_fault_currents(df, "Test", str(tmp_path))
        assert os.path.isfile(path)

    def test_transient_response(self, tmp_path):
        from reporting.plotter import plot_transient_response
        import math
        t      = [round(i * 0.01, 2) for i in range(200)]
        values = [1.0 - 0.5 * math.exp(-t_i) for t_i in t]
        path   = plot_transient_response(
            time       = t,
            channels   = {"voltage_pu": values, "rotor_angle_deg": values},
            scenario   = "Test",
            output_dir = str(tmp_path),
            fault_time = 1.0,
            clear_time = 1.1,
        )
        assert os.path.isfile(path)

    def test_pv_curves(self, tmp_path):
        from reporting.plotter import plot_pv_curves
        import numpy as np
        x = list(np.linspace(0, 400, 41))
        y = [1.02 - 0.0005 * xi for xi in x]
        path = plot_pv_curves(
            pv_data    = {
                "N-0 Base Case": {"P_mw": x, "V_pu": y, "category": "A"},
                "N-1 Cont 1":   {"P_mw": x[:30], "V_pu": y[:30], "category": "B"},
            },
            scenario   = "Test",
            output_dir = str(tmp_path),
            marp_mw    = 400.0,
        )
        assert os.path.isfile(path)

    def test_pre_post_comparison(self, mock_psse, dummy_sav, tmp_path):
        from reporting.plotter import plot_pre_post_comparison
        from studies.short_circuit.short_circuit_study import ShortCircuitStudy
        pre  = ShortCircuitStudy(mock_psse, scenario_label="Pre").run(dummy_sav)
        post = ShortCircuitStudy(mock_psse, scenario_label="Post").run(dummy_sav)
        df   = ShortCircuitStudy.compare(pre, post)
        if not df.empty:
            path = plot_pre_post_comparison(df, "Test", str(tmp_path))
            assert os.path.isfile(path)
