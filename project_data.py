"""
project_io/project_data.py
---------------------------
Central data model for all project-specific study inputs.

ProjectData is populated by excel_reader.py from study_scope_data.xlsx
and consumed by every study module. It is the single source of truth —
no study module reads the Excel file directly.

Design principles
-----------------
- All fields use Python built-in types (str, int, float, bool, list, dict)
  so they can be serialised, displayed in GUI tables, and validated easily.
- Optional fields that require bus numbers are None until filled by the user
  after running utils/bus_listing.py.
- No PSS/E imports here — this module must be importable without PSS/E.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


# ─────────────────────────────────────────────────────────────────────────────
# Sub-dataclasses (one per Excel sheet)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ProjectInfo:
    """
    Sheet: Project_Info
    Maps to Study Scope Section 1 (Table 1-1 and Table 1-2).
    """
    project_number:       str   = ""
    project_name:         str   = ""
    market_participant:   str   = ""
    studies_consultant:   str   = ""
    in_service_date:      str   = ""      # ISO string "YYYY-MM-DD"
    generation_type:      str   = "Solar"
    marp_mw:              float = 0.0
    max_capability_mw:    float = 0.0
    requested_sts_mw:     float = 0.0
    connection_voltage_kv:float = 240.0
    poc_substation_name:  str   = ""
    study_area_regions:   str   = ""      # comma-separated planning area numbers
    sav_file_path:        str   = ""      # absolute path to .sav file
    source_bus_number:    Optional[int]   = None   # solar plant generator bus
    poi_bus_number:       Optional[int]   = None   # point of interconnection bus
    ts_fault_bus_number:  Optional[int]   = None   # bus for TS fault application
    machine_id:           str   = "1"


@dataclass
class Scenario:
    """
    One row from Sheet: Scenarios (Table 4-1).
    """
    scenario_no:     int    = 0
    year:            int    = 0
    season:          str    = ""     # e.g. "SP", "SL", "WP"
    dispatch_cond:   str    = ""     # e.g. "HG", "HW"
    scenario_name:   str    = ""     # full label used in outputs
    pre_post:        str    = ""     # "Pre" or "Post"
    project_load_mw: float  = 0.0
    project_gen_mw:  float  = 0.0


@dataclass
class StudyMatrixEntry:
    """
    One row from Sheet: Study_Matrix (Table 5-1).
    Flags which studies to run for each scenario.
    """
    scenario_name:        str  = ""
    power_flow_cat_a:     bool = False
    power_flow_cat_b:     bool = False
    volt_stability_cat_a: bool = False
    volt_stability_cat_b: bool = False
    transient_cat_a:      bool = False
    transient_cat_b:      bool = False
    transient_conditional:bool = False   # True if marked X* (only if post shows issues)
    motor_starting_cat_a: bool = False
    motor_starting_cat_b: bool = False
    short_circuit_cat_a:  bool = False


@dataclass
class GeneratorDispatch:
    """
    One row from Sheet: Conv_Generation (Table 4-4).
    Represents a conventional generating unit (not wind/solar).
    dispatch_mw: dict mapping season_name → MW output
    e.g. {"2028 SP": 31.0, "2028 SL": 31.0, "2028 WP": 31.0}
    """
    facility_name:  str              = ""
    unit_no:        str              = ""
    bus_no:         Optional[int]    = None
    mc_mw:          float            = 0.0
    area_no:        int              = 0
    dispatch_mw:    Dict[str, float] = field(default_factory=dict)


@dataclass
class RenewableDispatch:
    """
    One row from Sheet: Renewable_Dispatch (Tables 4-5 and 4-6 combined).
    gen_type: "Wind" or "Solar"
    dispatch_mw: dict mapping season_name → MW output
    """
    facility_name:  str              = ""
    gen_type:       str              = ""    # "Wind" | "Solar"
    bus_no:         Optional[int]    = None
    mc_mw:          float            = 0.0
    area_no:        int              = 0
    dispatch_mw:    Dict[str, float] = field(default_factory=dict)


@dataclass
class IntertieFow:
    """
    One row from Sheet: Intertie_Flows (Table 4-7 or 4-9).
    flows: dict mapping intertie_name → MW (positive = export, negative = import)
    e.g. {"AB-BC": +851, "AB-SK": +150, "MATL": +186}
    """
    scenario_name: str              = ""
    flows:         Dict[str, float] = field(default_factory=dict)


@dataclass
class TSContingency:
    """
    One row from Sheet: TS_Contingencies (Table 4-8 or 4-14).
    Each physical line has two rows (near end fault, far end fault).
    """
    contingency_name:    str           = ""
    from_bus_name:       str           = ""
    to_bus_name:         str           = ""
    from_bus_no:         Optional[int] = None   # filled after bus listing
    to_bus_no:           Optional[int] = None   # filled after bus listing
    circuit_id:          str           = "1"
    fault_location:      str           = ""     # substation name where fault applied
    near_end_cycles:     float         = 6.0    # cycles at 60 Hz
    far_end_cycles:      float         = 8.0
    # Derived: fault clearing time in seconds (cycles / 60)
    @property
    def near_end_seconds(self) -> float:
        return round(self.near_end_cycles / 60.0, 6)

    @property
    def far_end_seconds(self) -> float:
        return round(self.far_end_cycles / 60.0, 6)


@dataclass
class SCSubstation:
    """
    One row from Sheet: SC_Substations (Section 5.4).
    """
    substation_name: str           = ""
    bus_no:          Optional[int] = None   # filled after bus listing
    notes:           str           = ""


@dataclass
class PVContingency:
    """
    One row from Sheet: PV_Contingencies.
    Subset of TS contingencies used for PV voltage stability N-1 analysis.
    """
    contingency_name: str           = ""
    from_bus_name:    str           = ""
    to_bus_name:      str           = ""
    from_bus_no:      Optional[int] = None
    to_bus_no:        Optional[int] = None
    circuit_id:       str           = "1"
    category:         str           = "B"   # Always "B" for N-1


@dataclass
class BusNumberEntry:
    """
    One row from Sheet: Bus_Numbers.
    Maps substation names from the Study Scope to PSS/E bus numbers.
    Filled by user after running utils/bus_listing.py.
    """
    substation_name: str           = ""
    bus_number:      Optional[int] = None
    base_kv:         float         = 0.0
    bus_type:        int           = 1      # 1=load, 2=gen, 3=slack
    notes:           str           = ""


# ─────────────────────────────────────────────────────────────────────────────
# Master ProjectData container
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ProjectData:
    """
    Complete data model for one AESO interconnection study project.

    Populated by project_io.excel_reader.ExcelReader.read()
    Consumed by all study modules and the GUI.

    Accessing convenience methods
    -----------------------------
    project.get_scenario("2028 SP Post-Project")
    project.get_study_matrix("2028 SP Post-Project")
    project.get_dispatch_for_season("2028 SP")
    project.scenarios_requiring_study("power_flow")
    project.validate()  → list of warning strings
    """

    # ── Core info ─────────────────────────────────────────────────────────────
    info:           ProjectInfo                  = field(default_factory=ProjectInfo)

    # ── Table data (one list per sheet) ───────────────────────────────────────
    scenarios:      List[Scenario]               = field(default_factory=list)
    study_matrix:   List[StudyMatrixEntry]        = field(default_factory=list)
    conv_gen:       List[GeneratorDispatch]       = field(default_factory=list)
    renewables:     List[RenewableDispatch]       = field(default_factory=list)
    intertie_flows: List[IntertieFow]             = field(default_factory=list)
    ts_contingencies: List[TSContingency]         = field(default_factory=list)
    sc_substations: List[SCSubstation]            = field(default_factory=list)
    pv_contingencies: List[PVContingency]         = field(default_factory=list)
    bus_numbers:    List[BusNumberEntry]          = field(default_factory=list)

    # ── Runtime paths (set by GUI or main.py, not stored in Excel) ────────────
    project_dir:    str = ""    # e.g. D:\Final_Project\files\projects\P2611
    output_dir:     str = ""    # e.g. D:\Final_Project\files\projects\P2611\output

    # ─────────────────────────────────────────────────────────────────────────
    # Convenience accessors
    # ─────────────────────────────────────────────────────────────────────────

    def get_scenario(self, scenario_name: str) -> Optional[Scenario]:
        """Return the Scenario matching scenario_name, or None."""
        for s in self.scenarios:
            if s.scenario_name == scenario_name:
                return s
        return None

    def get_study_matrix(self, scenario_name: str) -> Optional[StudyMatrixEntry]:
        """Return StudyMatrixEntry for a scenario, or None."""
        for m in self.study_matrix:
            if m.scenario_name == scenario_name:
                return m
        return None

    def scenarios_requiring_study(self, study_type: str) -> List[Scenario]:
        """
        Return list of scenarios where the given study is required.

        study_type must be one of:
            "power_flow", "voltage_stability", "transient",
            "motor_starting", "short_circuit"
        """
        result = []
        for entry in self.study_matrix:
            required = False
            if study_type == "power_flow":
                required = entry.power_flow_cat_a or entry.power_flow_cat_b
            elif study_type == "voltage_stability":
                required = entry.volt_stability_cat_a or entry.volt_stability_cat_b
            elif study_type == "transient":
                required = (entry.transient_cat_a or entry.transient_cat_b
                            or entry.transient_conditional)
            elif study_type == "motor_starting":
                required = entry.motor_starting_cat_a or entry.motor_starting_cat_b
            elif study_type == "short_circuit":
                required = entry.short_circuit_cat_a

            if required:
                sc = self.get_scenario(entry.scenario_name)
                if sc:
                    result.append(sc)
        return result

    def get_dispatch_for_season(
        self,
        season_label: str,
    ) -> Tuple[List[GeneratorDispatch], List[RenewableDispatch]]:
        """
        Return (conv_gen_list, renewable_list) filtered to season_label.
        Only units with a non-zero MW value for that season are returned
        (units at 0 MW are included since PSS/E needs them set explicitly).
        season_label must match a key in dispatch_mw dicts,
        e.g. "2028 SP", "2025 SL HW".
        """
        conv     = [g for g in self.conv_gen
                    if season_label in g.dispatch_mw]
        renewals = [r for r in self.renewables
                    if season_label in r.dispatch_mw]
        return conv, renewals

    def get_intertie_flows(self, scenario_name: str) -> Dict[str, float]:
        """
        Return intertie flow dict for a scenario name.
        e.g. {"AB-BC": +851, "AB-SK": +150, "MATL": +186}
        Returns empty dict if not found.
        """
        for entry in self.intertie_flows:
            if entry.scenario_name == scenario_name:
                return entry.flows
        return {}

    def get_bus_number(self, substation_name: str) -> Optional[int]:
        """
        Look up a bus number by substation name from the Bus_Numbers sheet.
        Returns None if not found or not yet filled.
        """
        for entry in self.bus_numbers:
            if entry.substation_name.strip().lower() == substation_name.strip().lower():
                return entry.bus_number
        return None

    def season_labels(self) -> List[str]:
        """
        Return the unique season labels present in the conv_gen dispatch data.
        e.g. ["2028 SP", "2028 SL", "2028 WP"]
        Used by readers to discover what seasons exist without hardcoding.
        """
        labels = set()
        for g in self.conv_gen:
            labels.update(g.dispatch_mw.keys())
        for r in self.renewables:
            labels.update(r.dispatch_mw.keys())
        return sorted(labels)

    # ─────────────────────────────────────────────────────────────────────────
    # Validation
    # ─────────────────────────────────────────────────────────────────────────

    def validate(self) -> List[str]:
        """
        Check ProjectData for missing or inconsistent values.

        Returns a list of warning strings. Empty list = no issues.
        Warnings do NOT block study execution — the GUI shows them
        and lets the user decide whether to proceed.

        Critical warnings (prefixed [CRITICAL]) will cause a specific
        study to be skipped if the required bus number is missing.
        """
        warnings: List[str] = []

        # ── Project info ──────────────────────────────────────────────────────
        if not self.info.project_number:
            warnings.append("[CRITICAL] Project number is empty.")
        if not self.info.sav_file_path:
            warnings.append("[CRITICAL] SAV file path is empty. No studies can run.")
        elif not os.path.isfile(self.info.sav_file_path):
            warnings.append(
                f"[CRITICAL] SAV file not found: {self.info.sav_file_path}"
            )
        if self.info.marp_mw <= 0:
            warnings.append("[WARNING] MARP is 0 or not set.")

        # ── Scenarios ─────────────────────────────────────────────────────────
        if not self.scenarios:
            warnings.append("[CRITICAL] No scenarios defined. Check Scenarios sheet.")
        if not self.study_matrix:
            warnings.append("[CRITICAL] Study matrix is empty. Check Study_Matrix sheet.")

        # ── Bus numbers needed per study type ─────────────────────────────────
        needs_pv = any(
            m.volt_stability_cat_a or m.volt_stability_cat_b
            for m in self.study_matrix
        )
        needs_ts = any(
            m.transient_cat_a or m.transient_cat_b or m.transient_conditional
            for m in self.study_matrix
        )

        if needs_pv:
            if self.info.source_bus_number is None:
                warnings.append(
                    "[CRITICAL] PV Voltage Stability requires Source Bus Number "
                    "(solar plant generator bus). Fill in Project_Info sheet "
                    "after running utils/bus_listing.py."
                )
            if self.info.poi_bus_number is None:
                warnings.append(
                    "[CRITICAL] PV Voltage Stability requires POI Bus Number. "
                    "Fill in Project_Info sheet after running utils/bus_listing.py."
                )

        if needs_ts:
            if self.info.ts_fault_bus_number is None:
                warnings.append(
                    "[WARNING] Transient Stability: TS Fault Bus Number is not set. "
                    "Study will use first TS contingency bus if available."
                )
            ts_missing_buses = [
                c.contingency_name for c in self.ts_contingencies
                if c.from_bus_no is None or c.to_bus_no is None
            ]
            if ts_missing_buses:
                warnings.append(
                    f"[WARNING] {len(ts_missing_buses)} TS contingencies have no "
                    f"bus numbers. Fill From/To Bus No columns in TS_Contingencies "
                    f"sheet. Affected: {', '.join(ts_missing_buses[:3])}"
                    + (" ..." if len(ts_missing_buses) > 3 else "")
                )

        # ── SC substations ────────────────────────────────────────────────────
        if not self.sc_substations:
            warnings.append(
                "[WARNING] No SC target substations defined. Short circuit study "
                "will run on all buses in the model."
            )
        else:
            sc_missing = [
                s.substation_name for s in self.sc_substations
                if s.bus_no is None
            ]
            if sc_missing:
                warnings.append(
                    f"[WARNING] {len(sc_missing)} SC target substations have no "
                    f"bus numbers. Short circuit will run on all buses. "
                    f"Fill Bus_No in SC_Substations sheet. "
                    f"Affected: {', '.join(sc_missing[:3])}"
                    + (" ..." if len(sc_missing) > 3 else "")
                )

        # ── Generation dispatch ───────────────────────────────────────────────
        gen_missing_buses = [
            g.facility_name for g in self.conv_gen if g.bus_no is None
        ]
        if gen_missing_buses:
            warnings.append(
                f"[WARNING] {len(gen_missing_buses)} conventional generators have "
                f"no bus number. Dispatch setting will be skipped for these units. "
                f"Affected: {', '.join(gen_missing_buses[:3])}"
                + (" ..." if len(gen_missing_buses) > 3 else "")
            )

        ren_missing_buses = [
            r.facility_name for r in self.renewables if r.bus_no is None
        ]
        if ren_missing_buses:
            warnings.append(
                f"[WARNING] {len(ren_missing_buses)} renewable generators have "
                f"no bus number. Dispatch setting will be skipped for these units. "
                f"Affected: {', '.join(ren_missing_buses[:3])}"
                + (" ..." if len(ren_missing_buses) > 3 else "")
            )

        # ── Intertie flows ────────────────────────────────────────────────────
        scenario_names = {s.scenario_name for s in self.scenarios}
        intertie_names = {i.scenario_name for i in self.intertie_flows}
        missing_intertie = scenario_names - intertie_names
        if missing_intertie:
            warnings.append(
                f"[WARNING] Intertie flows missing for {len(missing_intertie)} "
                f"scenarios: {', '.join(sorted(missing_intertie)[:3])}"
                + (" ..." if len(missing_intertie) > 3 else "")
            )

        return warnings

    # ─────────────────────────────────────────────────────────────────────────
    # String representation
    # ─────────────────────────────────────────────────────────────────────────

    def __str__(self) -> str:
        return (
            f"ProjectData("
            f"project={self.info.project_number} "
            f"'{self.info.project_name}', "
            f"scenarios={len(self.scenarios)}, "
            f"conv_gen={len(self.conv_gen)}, "
            f"renewables={len(self.renewables)}, "
            f"ts_contingencies={len(self.ts_contingencies)}, "
            f"sc_substations={len(self.sc_substations)}"
            f")"
        )
