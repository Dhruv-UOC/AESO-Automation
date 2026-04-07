# AESO Interconnection Study Automation Tool

A Python desktop tool that automates the four interconnection-capacity studies required for an AESO ICA application — Power Flow, Short Circuit, Transient Stability, and P-V / Voltage Stability — by driving PSS/E 35 programmatically via its `psspy` API.  All study parameters are read from a single Excel workbook (`study_scope_data.xlsx`), results are written back to Excel, and summary plots are generated automatically.

---

## Prerequisites

| Requirement | Detail |
|---|---|
| **Operating System** | Windows (required by PSS/E) |
| **Python** | 3.7 or later (64-bit) |
| **PSS/E** | Version 35, licensed installation |
| **PSS/E Python path** | `PSSE35\PSSBIN` added to `PYTHONPATH` or configured in `config/settings.py` |

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Configuration

Open `config/settings.py` and set the two paths for your machine:

```python
BASE_DIR  = r"D:\Final_Project\files"   # root folder that contains the projects\ subfolder
PSSE_PATH = r"C:\Program Files\PTI\PSSE35\PSSBIN"
```

All other settings (log level, output sub-folder names, AESO constants) are also in that file.

---

## Project Setup

1. Copy the template workbook into your project folder:

   ```
   templates\study_scope_template.xlsx  →  projects\<project_number>\study_scope_data.xlsx
   ```

2. Fill in every tab from the AESO Study Scope PDF supplied by AESO.

3. Look up bus numbers from your base-case `.sav` file if needed:

   ```bash
   python utils/bus_listing.py --sav "D:\models\base_case.sav"
   ```

---

## Running

### GUI mode (default)

```bash
python main.py
python main.py --gui
```

Opens a Tkinter window.  Use the interface to select the project folder, choose studies, and click **Run**.

### CLI mode

```bash
python main.py --project "D:\Final_Project\files\projects\P2611"
```

Add `--mock` to run without a PSS/E licence (uses stub results — useful for testing the pipeline):

```bash
python main.py --project "D:\Final_Project\files\projects\P2611" --mock
```

---

## Testing

The test suite uses mock PSS/E mode and requires no licence:

```bash
python -m pytest tests/test_mock.py -v
```

---

## Output Structure

After a run, the project folder contains:

```
projects/<project_number>/
├── study_scope_data.xlsx        ← input (you provide)
└── output/
    ├── results/                 ← per-study Excel result files
    ├── plots/                   ← PNG / PDF plots
    └── reports/                 ← consolidated summary reports
```

---

## Repository Layout

```
AESO-Automation/
├── main.py                          ← entry point (GUI + CLI)
├── requirements.txt
├── templates/
│   └── study_scope_template.xlsx
├── config/
│   └── settings.py
├── core/
│   └── psse_interface.py
├── project_io/
│   ├── excel_reader.py
│   ├── excel_writer.py
│   └── project_data.py
├── studies/
│   ├── power_flow/
│   │   └── power_flow_study.py
│   ├── short_circuit/
│   │   └── short_circuit_study.py
│   ├── transient_stability/
│   │   └── transient_stability_study.py
│   └── pv_voltage/
│       └── pv_stability_study.py
├── reporting/
│   └── plotter.py
├── gui/
│   └── main_window.py
├── utils/
│   └── bus_listing.py
└── tests/
    └── test_mock.py
```
