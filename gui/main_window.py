"""
gui/main_window.py
------------------
Tkinter GUI for the AESO Interconnection Study Automation Tool.

Layout
------
  ┌─────────────────────────────────────────────────────────────┐
  │  Header: AESO logo + project title                          │
  ├──────────────────┬──────────────────────────────────────────┤
  │  LEFT PANEL      │  RIGHT PANEL                             │
  │                  │                                          │
  │  [Project Setup] │  [Data Tabs]                             │
  │  • Project dir   │  Project_Info / Scenarios / Study_Matrix │
  │  • SAV file      │  Conv_Gen / Renewables / Intertie /      │
  │  • Excel file    │  TS_Cont / SC_Subs / PV_Cont / Bus_Nos  │
  │                  │                                          │
  │  [Validation]    │                                          │
  │  • Warnings list │                                          │
  │                  │                                          │
  │  [Run Studies]   │                                          │
  │  • Checkboxes    │                                          │
  │  • RUN button    │                                          │
  ├──────────────────┴──────────────────────────────────────────┤
  │  LOG WINDOW (live scrolling output)                         │
  └─────────────────────────────────────────────────────────────┘

Usage
-----
    python gui/main_window.py           # launch GUI directly
    python main.py --gui                # launch via main.py
"""

import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

# ── Ensure project root is on path ───────────────────────────────────────────
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from project_io.excel_reader import ExcelReader, ExcelReaderError
from project_io.excel_writer import ExcelWriter
from project_io.project_data import ProjectData

logger = logging.getLogger(__name__)

# ── AESO Colour Palette ───────────────────────────────────────────────────────
C_DARK_BLUE  = "#003865"
C_MID_BLUE   = "#1F5FA6"
C_LIGHT_BLUE = "#D6E4F0"
C_ORANGE     = "#E87722"
C_GREEN      = "#00853E"
C_RED        = "#C8102E"
C_WHITE      = "#FFFFFF"
C_LIGHT_GREY = "#F0F2F5"
C_MID_GREY   = "#D0D3D8"
C_DARK_GREY  = "#4A4E54"
C_YELLOW     = "#FFF2CC"

# Sheet display names mapped to ProjectData field names
SHEET_TABS = [
    ("Project Info",   "info"),
    ("Scenarios",      "scenarios"),
    ("Study Matrix",   "study_matrix"),
    ("Conv Gen",       "conv_gen"),
    ("Renewables",     "renewables"),
    ("Intertie Flows", "intertie_flows"),
    ("TS Contingencies","ts_contingencies"),
    ("SC Substations", "sc_substations"),
    ("PV Contingencies","pv_contingencies"),
    ("Bus Numbers",    "bus_numbers"),
]


class QueueHandler(logging.Handler):
    """Routes log records into a thread-safe queue for GUI display."""
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord):
        self.log_queue.put(self.format(record))


class AESOStudyGUI:
    """Main application window."""

    def __init__(self, root: tk.Tk):
        self.root       = root
        self.project:   Optional[ProjectData] = None
        self.log_queue  = queue.Queue()
        self._run_thread: Optional[threading.Thread] = None

        # StringVars for path fields
        self.var_project_dir = tk.StringVar()
        self.var_sav_path    = tk.StringVar()
        self.var_excel_path  = tk.StringVar()

        # Study selection checkboxes
        self.study_vars: Dict[str, tk.BooleanVar] = {
            "power_flow":         tk.BooleanVar(value=True),
            "short_circuit":      tk.BooleanVar(value=True),
            "transient":          tk.BooleanVar(value=True),
            "voltage_stability":  tk.BooleanVar(value=True),
        }
        STUDY_LABELS = {
            "power_flow":        "Power Flow",
            "short_circuit":     "Short Circuit",
            "transient":         "Transient Stability",
            "voltage_stability": "PV Voltage Stability",
        }
        self.study_labels = STUDY_LABELS

        self._setup_root()
        self._build_ui()
        self._setup_logging()
        self._poll_log_queue()

    # ── Window setup ──────────────────────────────────────────────────────────

    def _setup_root(self):
        self.root.title("AESO Interconnection Study Automation Tool")
        self.root.geometry("1280x860")
        self.root.minsize(1100, 720)
        self.root.configure(bg=C_LIGHT_GREY)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame",       background=C_LIGHT_GREY)
        style.configure("TLabel",       background=C_LIGHT_GREY,
                        font=("Segoe UI", 10), foreground=C_DARK_GREY)
        style.configure("TButton",      font=("Segoe UI", 10, "bold"),
                        background=C_MID_BLUE, foreground=C_WHITE,
                        borderwidth=0, focusthickness=0, padding=(10, 5))
        style.map("TButton",
                  background=[("active", C_DARK_BLUE), ("disabled", C_MID_GREY)],
                  foreground=[("disabled", C_DARK_GREY)])
        style.configure("Run.TButton",  font=("Segoe UI", 12, "bold"),
                        background=C_GREEN, foreground=C_WHITE, padding=(16, 8))
        style.map("Run.TButton",
                  background=[("active", "#005f2e"), ("disabled", C_MID_GREY)])
        style.configure("Warn.TButton", background=C_ORANGE, foreground=C_WHITE)
        style.configure("TNotebook",    background=C_LIGHT_GREY, borderwidth=0)
        style.configure("TNotebook.Tab",
                        font=("Segoe UI", 9), padding=(8, 4),
                        background=C_MID_GREY, foreground=C_DARK_GREY)
        style.map("TNotebook.Tab",
                  background=[("selected", C_DARK_BLUE)],
                  foreground=[("selected", C_WHITE)])
        style.configure("Treeview",
                        font=("Segoe UI", 9), rowheight=22,
                        background=C_WHITE, fieldbackground=C_WHITE,
                        foreground=C_DARK_GREY)
        style.configure("Treeview.Heading",
                        font=("Segoe UI", 9, "bold"),
                        background=C_DARK_BLUE, foreground=C_WHITE)
        style.map("Treeview", background=[("selected", C_LIGHT_BLUE)])
        style.configure("TCheckbutton",
                        background=C_LIGHT_GREY, font=("Segoe UI", 10),
                        foreground=C_DARK_GREY)
        style.configure("TEntry",       font=("Segoe UI", 9), padding=(4, 3))
        style.configure("TLabelframe",
                        background=C_LIGHT_GREY,
                        font=("Segoe UI", 10, "bold"),
                        foreground=C_DARK_BLUE)
        style.configure("TLabelframe.Label",
                        background=C_LIGHT_GREY,
                        font=("Segoe UI", 10, "bold"),
                        foreground=C_DARK_BLUE)

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=C_DARK_BLUE, height=56)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)

        tk.Label(
            hdr,
            text="  AESO  ",
            bg=C_ORANGE, fg=C_WHITE,
            font=("Segoe UI", 14, "bold"),
            padx=6, pady=4,
        ).pack(side="left", padx=(16, 0), pady=8)

        tk.Label(
            hdr,
            text="Interconnection Study Automation Tool",
            bg=C_DARK_BLUE, fg=C_WHITE,
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left", padx=12)

        self.lbl_project_title = tk.Label(
            hdr,
            text="No project loaded",
            bg=C_DARK_BLUE, fg=C_LIGHT_BLUE,
            font=("Segoe UI", 10),
        )
        self.lbl_project_title.pack(side="right", padx=16)

        # ── Main content area ─────────────────────────────────────────────────
        main = ttk.Frame(self.root)
        main.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(1, weight=1)

        # ── Left panel ────────────────────────────────────────────────────────
        left = ttk.Frame(main, width=310)
        left.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=8)
        left.grid_propagate(False)
        left.grid_columnconfigure(0, weight=1)

        self._build_project_setup(left)
        self._build_validation_panel(left)
        self._build_run_panel(left)

        # ── Right panel (data tabs) ───────────────────────────────────────────
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self._build_data_tabs(right)

        # ── Log window ────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="  Study Log")
        log_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=10, wrap="word",
            bg="#1a1e2e", fg="#c8d8e8",
            font=("Consolas", 9),
            borderwidth=0, insertbackground=C_WHITE,
            state="disabled",
        )
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=0, column=0, sticky="ew", padx=(4, 0), pady=4)
        log_scroll.grid(row=0, column=1, sticky="ns", pady=4)

        # Tag colours for log levels
        self.log_text.tag_configure("INFO",    foreground="#c8d8e8")
        self.log_text.tag_configure("WARNING", foreground=C_ORANGE)
        self.log_text.tag_configure("ERROR",   foreground=C_RED)
        self.log_text.tag_configure("CRITICAL",foreground=C_RED)

        # Clear log button
        tk.Button(
            log_frame,
            text="Clear Log",
            bg=C_MID_GREY, fg=C_DARK_GREY,
            font=("Segoe UI", 8),
            relief="flat",
            command=self._clear_log,
        ).grid(row=1, column=0, sticky="e", padx=4, pady=(0, 4))

    def _build_project_setup(self, parent):
        """Project folder, SAV file, and Excel file pickers."""
        frame = ttk.LabelFrame(parent, text="  Project Setup")
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(1, weight=1)

        # Project folder
        ttk.Label(frame, text="Project Folder:").grid(
            row=0, column=0, sticky="w", padx=(8, 4), pady=(8, 2))
        ttk.Entry(frame, textvariable=self.var_project_dir,
                  state="readonly").grid(
            row=0, column=1, sticky="ew", padx=4, pady=(8, 2))
        ttk.Button(frame, text="Browse…",
                   command=self._browse_project_dir).grid(
            row=0, column=2, padx=(4, 8), pady=(8, 2))

        # SAV file
        ttk.Label(frame, text="SAV File:").grid(
            row=1, column=0, sticky="w", padx=(8, 4), pady=2)
        ttk.Entry(frame, textvariable=self.var_sav_path,
                  state="readonly").grid(
            row=1, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(frame, text="Browse…",
                   command=self._browse_sav).grid(
            row=1, column=2, padx=(4, 8), pady=2)

        # Excel file
        ttk.Label(frame, text="Study Scope Excel:").grid(
            row=2, column=0, sticky="w", padx=(8, 4), pady=2)
        ttk.Entry(frame, textvariable=self.var_excel_path,
                  state="readonly").grid(
            row=2, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(frame, text="Browse…",
                   command=self._browse_excel).grid(
            row=2, column=2, padx=(4, 8), pady=2)

        # Action buttons row
        btn_row = ttk.Frame(frame)
        btn_row.grid(row=3, column=0, columnspan=3,
                     sticky="ew", padx=8, pady=(4, 8))
        btn_row.grid_columnconfigure((0, 1, 2), weight=1)

        ttk.Button(btn_row, text="New Project",
                   command=self._new_project).grid(
            row=0, column=0, sticky="ew", padx=(0, 3))
        ttk.Button(btn_row, text="Load Excel",
                   command=self._load_excel).grid(
            row=0, column=1, sticky="ew", padx=3)
        ttk.Button(btn_row, text="Save Excel",
                   command=self._save_excel).grid(
            row=0, column=2, sticky="ew", padx=(3, 0))

    def _build_validation_panel(self, parent):
        """Validation warnings list."""
        frame = ttk.LabelFrame(parent, text="  Validation Warnings")
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(0, weight=1)

        self.warn_text = tk.Text(
            frame,
            height=8, wrap="word",
            bg=C_YELLOW, fg=C_DARK_GREY,
            font=("Segoe UI", 9),
            borderwidth=0, state="disabled",
        )
        warn_scroll = ttk.Scrollbar(frame, command=self.warn_text.yview)
        self.warn_text.configure(yscrollcommand=warn_scroll.set)
        self.warn_text.grid(row=0, column=0, sticky="ew", padx=(4, 0), pady=4)
        warn_scroll.grid(row=0, column=1, sticky="ns", pady=4)
        self.warn_text.tag_configure("critical", foreground=C_RED,
                                     font=("Segoe UI", 9, "bold"))
        self.warn_text.tag_configure("warning",  foreground="#7a5000")

        ttk.Button(frame, text="Validate Now",
                   command=self._run_validation).grid(
            row=1, column=0, columnspan=2,
            sticky="ew", padx=8, pady=(0, 6))

    def _build_run_panel(self, parent):
        """Study selection checkboxes and RUN button."""
        frame = ttk.LabelFrame(parent, text="  Run Studies")
        frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(0, weight=1)

        for i, (key, label) in enumerate(self.study_labels.items()):
            ttk.Checkbutton(
                frame,
                text=label,
                variable=self.study_vars[key],
            ).grid(row=i, column=0, sticky="w", padx=12, pady=2)

        ttk.Separator(frame).grid(
            row=len(self.study_labels), column=0,
            sticky="ew", padx=8, pady=6)

        # Scenario filter
        ttk.Label(frame, text="Scenarios:").grid(
            row=len(self.study_labels) + 1, column=0,
            sticky="w", padx=12)
        self.scenario_listbox = tk.Listbox(
            frame,
            selectmode="multiple",
            height=6,
            font=("Segoe UI", 9),
            bg=C_WHITE,
            fg=C_DARK_GREY,
            selectbackground=C_LIGHT_BLUE,
            selectforeground=C_DARK_BLUE,
            borderwidth=1,
            relief="solid",
        )
        self.scenario_listbox.grid(
            row=len(self.study_labels) + 2, column=0,
            sticky="ew", padx=12, pady=(2, 4))

        ttk.Button(
            frame,
            text="Select All Scenarios",
            command=self._select_all_scenarios,
        ).grid(row=len(self.study_labels) + 3, column=0,
               sticky="ew", padx=12, pady=(0, 4))

        # RUN button
        self.btn_run = ttk.Button(
            frame,
            text="▶  RUN STUDIES",
            style="Run.TButton",
            command=self._run_studies,
            state="disabled",
        )
        self.btn_run.grid(
            row=len(self.study_labels) + 4, column=0,
            sticky="ew", padx=8, pady=(4, 8))

        # Open output folder button
        self.btn_open_output = ttk.Button(
            frame,
            text="📂  Open Output Folder",
            command=self._open_output_folder,
            state="disabled",
        )
        self.btn_open_output.grid(
            row=len(self.study_labels) + 5, column=0,
            sticky="ew", padx=8, pady=(0, 8))

    def _build_data_tabs(self, parent):
        """Notebook with one editable tab per data sheet."""
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.tab_trees: Dict[str, ttk.Treeview] = {}

        for display_name, field_name in SHEET_TABS:
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=f" {display_name} ")
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)

            tree = self._make_treeview(tab)
            self.tab_trees[field_name] = tree

        # Project_Info tab uses a special key-value layout
        # (all other tabs use generic table display)
        self._show_placeholder_tabs()

    @staticmethod
    def _make_treeview(parent) -> ttk.Treeview:
        """Create a scrollable Treeview inside a frame."""
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        vsb  = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
        hsb  = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid( row=0, column=1, sticky="ns")
        hsb.grid( row=1, column=0, sticky="ew")

        return tree

    def _show_placeholder_tabs(self):
        """Show 'No data loaded' message in all tabs."""
        for field_name, tree in self.tab_trees.items():
            tree["columns"] = ("message",)
            tree.heading("message", text="Status")
            tree.column( "message", width=400, anchor="w")
            tree.delete(*tree.get_children())
            tree.insert("", "end", values=("No project loaded. "
                                            "Browse to a Study Scope Excel file "
                                            "and click Load Excel.",))

    # ── File pickers ──────────────────────────────────────────────────────────

    def _browse_project_dir(self):
        path = filedialog.askdirectory(title="Select Project Folder")
        if path:
            self.var_project_dir.set(path)
            # Auto-detect SAV and Excel files
            for f in os.listdir(path):
                if f.endswith(".sav") and not self.var_sav_path.get():
                    self.var_sav_path.set(os.path.join(path, f))
                if ("study_scope" in f.lower() or "scope_data" in f.lower()) \
                        and f.endswith(".xlsx"):
                    self.var_excel_path.set(os.path.join(path, f))
            # Check cases subfolder for .sav
            cases_dir = os.path.join(path, "cases")
            if os.path.isdir(cases_dir) and not self.var_sav_path.get():
                for f in os.listdir(cases_dir):
                    if f.endswith(".sav"):
                        self.var_sav_path.set(os.path.join(cases_dir, f))
                        break

    def _browse_sav(self):
        path = filedialog.askopenfilename(
            title="Select PSS/E Case File",
            filetypes=[("PSS/E Case", "*.sav"), ("All Files", "*.*")],
        )
        if path:
            self.var_sav_path.set(path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Study Scope Excel File",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
        )
        if path:
            self.var_excel_path.set(path)

    # ── Project operations ────────────────────────────────────────────────────

    def _new_project(self):
        """Create a new project folder with a blank study_scope_data.xlsx."""
        # Ask user for parent directory and project number
        parent_dir = filedialog.askdirectory(
            title="Select Parent Folder for New Project"
        )
        if not parent_dir:
            return

        project_num = tk.simpledialog.askstring(
            "New Project",
            "Enter project number (e.g. P2611):",
            parent=self.root,
        )
        if not project_num:
            return

        project_dir = os.path.join(parent_dir, project_num)
        excel_path  = os.path.join(project_dir, "study_scope_data.xlsx")
        cases_dir   = os.path.join(project_dir, "cases")

        # Find template
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "templates", "study_scope_template.xlsx"
        )
        if not os.path.isfile(template_path):
            messagebox.showerror(
                "Template Not Found",
                f"Could not find:\n{template_path}\n\n"
                "Ensure templates/study_scope_template.xlsx exists."
            )
            return

        try:
            os.makedirs(cases_dir, exist_ok=True)
            os.makedirs(os.path.join(project_dir, "output"), exist_ok=True)
            ExcelWriter.new_project_file(template_path, excel_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        self.var_project_dir.set(project_dir)
        self.var_excel_path.set(excel_path)
        self._log(f"New project created: {project_dir}")
        messagebox.showinfo(
            "Project Created",
            f"Project folder created:\n{project_dir}\n\n"
            f"Fill study_scope_data.xlsx from your Study Scope PDF, "
            f"place your .sav file in the cases/ subfolder, "
            f"then click Load Excel."
        )

    def _load_excel(self):
        """Read study_scope_data.xlsx and populate all data tabs."""
        excel_path = self.var_excel_path.get()
        if not excel_path:
            excel_path = filedialog.askopenfilename(
                title="Select Study Scope Excel",
                filetypes=[("Excel Workbook", "*.xlsx")],
            )
            if not excel_path:
                return
            self.var_excel_path.set(excel_path)

        try:
            reader = ExcelReader(excel_path)
            self.project = reader.read()
        except ExcelReaderError as exc:
            messagebox.showerror("Load Error", str(exc))
            return

        # Update SAV path from project info if set
        if self.project.info.sav_file_path:
            self.var_sav_path.set(self.project.info.sav_file_path)

        # Set project dir
        if not self.var_project_dir.get():
            self.var_project_dir.set(os.path.dirname(excel_path))

        self.project.project_dir = self.var_project_dir.get()
        self.project.output_dir  = os.path.join(
            self.project.project_dir, "output"
        )

        self._refresh_all_tabs()
        self._refresh_scenario_list()
        self._run_validation()

        # Update header title
        title = (
            f"{self.project.info.project_number} — "
            f"{self.project.info.project_name}"
        )
        self.lbl_project_title.configure(text=title)
        self.btn_run.configure(state="normal")
        self._log(f"Loaded: {self.project}")

    def _save_excel(self):
        """Save current tab data back to the Excel file."""
        if self.project is None:
            messagebox.showwarning("No Data", "Load a project first.")
            return
        excel_path = self.var_excel_path.get()
        if not excel_path:
            messagebox.showwarning("No File", "No Excel file path set.")
            return
        # Update SAV path from GUI field
        self.project.info.sav_file_path = self.var_sav_path.get()
        try:
            writer = ExcelWriter(excel_path)
            writer.write(self.project)
            self._log(f"Saved: {excel_path}")
            messagebox.showinfo("Saved", f"Study scope data saved to:\n{excel_path}")
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    # ── Tab population ────────────────────────────────────────────────────────

    def _refresh_all_tabs(self):
        """Populate all data tabs from self.project."""
        if self.project is None:
            return

        self._populate_info_tab()
        self._populate_generic_tab("scenarios",        self._scenarios_rows())
        self._populate_generic_tab("study_matrix",     self._matrix_rows())
        self._populate_generic_tab("conv_gen",         self._conv_gen_rows())
        self._populate_generic_tab("renewables",       self._renewables_rows())
        self._populate_generic_tab("intertie_flows",   self._intertie_rows())
        self._populate_generic_tab("ts_contingencies", self._ts_cont_rows())
        self._populate_generic_tab("sc_substations",   self._sc_subs_rows())
        self._populate_generic_tab("pv_contingencies", self._pv_cont_rows())
        self._populate_generic_tab("bus_numbers",      self._bus_num_rows())

    def _populate_info_tab(self):
        """Project_Info tab: two columns — Field | Value."""
        tree = self.tab_trees["info"]
        tree["columns"] = ("field", "value")
        tree.heading("field", text="Field")
        tree.heading("value", text="Value")
        tree.column("field", width=220, anchor="w")
        tree.column("value", width=380, anchor="w")
        tree.delete(*tree.get_children())

        info = self.project.info
        rows = [
            ("Project Number",          info.project_number),
            ("Project Name",            info.project_name),
            ("Market Participant",       info.market_participant),
            ("Studies Consultant",       info.studies_consultant),
            ("In-Service Date",          info.in_service_date),
            ("Generation Type",          info.generation_type),
            ("MARP (MW)",               info.marp_mw),
            ("Max Capability (MW)",      info.max_capability_mw),
            ("Connection Voltage (kV)",  info.connection_voltage_kv),
            ("POC Substation",           info.poc_substation_name),
            ("Study Area Regions",       info.study_area_regions),
            ("SAV File Path",            info.sav_file_path),
            ("Source Bus Number",        info.source_bus_number or "⚠ Not set"),
            ("POI Bus Number",           info.poi_bus_number    or "⚠ Not set"),
            ("TS Fault Bus Number",      info.ts_fault_bus_number or "⚠ Not set"),
        ]
        for field, value in rows:
            tag = "missing" if str(value).startswith("⚠") else ""
            tree.insert("", "end", values=(field, value), tags=(tag,))
        tree.tag_configure("missing", foreground=C_ORANGE)

    def _populate_generic_tab(self, field_name: str, rows_data: tuple):
        """Generic tab: columns + rows."""
        tree = self.tab_trees[field_name]
        columns, rows = rows_data

        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=max(80, len(col) * 9), anchor="w",
                        minwidth=60)
        tree.delete(*tree.get_children())

        alt = False
        for row in rows:
            tag = "alt" if alt else ""
            tree.insert("", "end", values=row, tags=(tag,))
            alt = not alt
        tree.tag_configure("alt", background=C_LIGHT_BLUE)

    # ── Row builders ──────────────────────────────────────────────────────────

    def _scenarios_rows(self):
        cols = ("No", "Year", "Season", "Dispatch", "Scenario Name",
                "Pre/Post", "Load (MW)", "Gen (MW)")
        rows = [
            (s.scenario_no, s.year, s.season, s.dispatch_cond,
             s.scenario_name, s.pre_post,
             s.project_load_mw, s.project_gen_mw)
            for s in self.project.scenarios
        ]
        return cols, rows

    def _matrix_rows(self):
        cols = ("Scenario", "PF-A", "PF-B", "VS-A", "VS-B",
                "TS-A", "TS-B", "TS-Cond", "MS-A", "MS-B", "SC-A")
        def _x(b): return "X" if b else ""
        rows = [
            (m.scenario_name,
             _x(m.power_flow_cat_a), _x(m.power_flow_cat_b),
             _x(m.volt_stability_cat_a), _x(m.volt_stability_cat_b),
             _x(m.transient_cat_a), _x(m.transient_cat_b),
             "X*" if m.transient_conditional else "",
             _x(m.motor_starting_cat_a), _x(m.motor_starting_cat_b),
             _x(m.short_circuit_cat_a))
            for m in self.project.study_matrix
        ]
        return cols, rows

    def _conv_gen_rows(self):
        season_labels = sorted({
            k for g in self.project.conv_gen for k in g.dispatch_mw
        })
        cols = ("Facility", "Unit", "Bus No", "MC (MW)", "Area") + tuple(season_labels)
        rows = []
        for g in self.project.conv_gen:
            base = (g.facility_name, g.unit_no,
                    g.bus_no or "—", g.mc_mw, g.area_no)
            mws  = tuple(g.dispatch_mw.get(sl, "") for sl in season_labels)
            rows.append(base + mws)
        return cols, rows

    def _renewables_rows(self):
        season_labels = sorted({
            k for r in self.project.renewables for k in r.dispatch_mw
        })
        cols = ("Facility", "Type", "Bus No", "MC (MW)", "Area") + tuple(season_labels)
        rows = []
        for r in self.project.renewables:
            base = (r.facility_name, r.gen_type,
                    r.bus_no or "—", r.mc_mw, r.area_no)
            mws  = tuple(r.dispatch_mw.get(sl, "") for sl in season_labels)
            rows.append(base + mws)
        return cols, rows

    def _intertie_rows(self):
        if not self.project.intertie_flows:
            return ("Scenario",), []
        all_keys = sorted({
            k for f in self.project.intertie_flows for k in f.flows
        })
        cols = ("Scenario",) + tuple(all_keys)
        rows = [
            (f.scenario_name,) + tuple(f.flows.get(k, "") for k in all_keys)
            for f in self.project.intertie_flows
        ]
        return cols, rows

    def _ts_cont_rows(self):
        cols = ("Contingency", "From Bus Name", "To Bus Name",
                "From Bus No", "To Bus No", "Ckt",
                "Fault Location", "Near End (cyc)", "Far End (cyc)")
        rows = [
            (c.contingency_name, c.from_bus_name, c.to_bus_name,
             c.from_bus_no or "—", c.to_bus_no or "—", c.circuit_id,
             c.fault_location, c.near_end_cycles, c.far_end_cycles)
            for c in self.project.ts_contingencies
        ]
        return cols, rows

    def _sc_subs_rows(self):
        cols = ("Substation Name", "Bus No", "Notes")
        rows = [
            (s.substation_name, s.bus_no or "—", s.notes)
            for s in self.project.sc_substations
        ]
        return cols, rows

    def _pv_cont_rows(self):
        cols = ("Contingency", "From Bus", "To Bus",
                "From No", "To No", "Ckt", "Category")
        rows = [
            (c.contingency_name, c.from_bus_name, c.to_bus_name,
             c.from_bus_no or "—", c.to_bus_no or "—",
             c.circuit_id, c.category)
            for c in self.project.pv_contingencies
        ]
        return cols, rows

    def _bus_num_rows(self):
        cols = ("Substation Name", "Bus Number", "Base kV", "Bus Type", "Notes")
        rows = [
            (b.substation_name, b.bus_number or "—",
             b.base_kv, b.bus_type, b.notes)
            for b in self.project.bus_numbers
        ]
        return cols, rows

    # ── Validation ────────────────────────────────────────────────────────────

    def _run_validation(self):
        """Run ProjectData.validate() and show warnings."""
        if self.project is None:
            return

        # Sync SAV path from GUI field to project
        self.project.info.sav_file_path = self.var_sav_path.get()

        warnings = self.project.validate()

        self.warn_text.configure(state="normal")
        self.warn_text.delete("1.0", "end")
        if not warnings:
            self.warn_text.insert("end", "✔ No issues found. Ready to run.", "")
        else:
            for w in warnings:
                tag = "critical" if "[CRITICAL]" in w else "warning"
                self.warn_text.insert("end", w + "\n", tag)
        self.warn_text.configure(state="disabled")

        # Enable/disable run based on critical warnings
        critical = [w for w in warnings if "[CRITICAL]" in w]
        sav_ok   = bool(self.var_sav_path.get())
        self.btn_run.configure(
            state="normal" if (self.project and sav_ok) else "disabled"
        )

    # ── Scenario list ─────────────────────────────────────────────────────────

    def _refresh_scenario_list(self):
        self.scenario_listbox.delete(0, "end")
        if self.project:
            for sc in self.project.scenarios:
                self.scenario_listbox.insert("end", sc.scenario_name)
            self.scenario_listbox.select_set(0, "end")

    def _select_all_scenarios(self):
        self.scenario_listbox.select_set(0, "end")

    def _get_selected_scenarios(self) -> List[str]:
        indices = self.scenario_listbox.curselection()
        return [self.scenario_listbox.get(i) for i in indices]

    # ── Run studies ───────────────────────────────────────────────────────────

    def _run_studies(self):
        """Launch studies in a background thread."""
        if self.project is None:
            messagebox.showwarning("No Project", "Load a project first.")
            return

        sav_path = self.var_sav_path.get()
        if not sav_path or not os.path.isfile(sav_path):
            messagebox.showerror(
                "SAV Not Found",
                f"SAV file not found:\n{sav_path}\n\n"
                "Browse to a valid .sav file."
            )
            return

        selected_scenarios = self._get_selected_scenarios()
        if not selected_scenarios:
            messagebox.showwarning(
                "No Scenarios Selected",
                "Select at least one scenario in the Run Studies panel."
            )
            return

        selected_studies = [k for k, v in self.study_vars.items() if v.get()]
        if not selected_studies:
            messagebox.showwarning(
                "No Studies Selected",
                "Check at least one study type to run."
            )
            return

        # Disable run button during execution
        self.btn_run.configure(state="disabled", text="⏳  Running…")
        self.root.update()

        # Update SAV path in project
        self.project.info.sav_file_path = sav_path
        self.project.project_dir = self.var_project_dir.get()
        self.project.output_dir  = os.path.join(
            self.project.project_dir, "output"
        )

        # Launch in thread so GUI stays responsive
        self._run_thread = threading.Thread(
            target=self._run_studies_thread,
            args=(sav_path, selected_scenarios, selected_studies),
            daemon=True,
        )
        self._run_thread.start()

    def _run_studies_thread(
        self,
        sav_path:           str,
        selected_scenarios: List[str],
        selected_studies:   List[str],
    ):
        """Background thread: imports study modules and runs them."""
        try:
            from config.settings import PSSE_PATH, PSSE_VERSION
            from core.psse_interface import PSSEInterface
            from studies.power_flow.power_flow_study import PowerFlowStudy
            from studies.short_circuit.short_circuit_study import ShortCircuitStudy
            from studies.transient_stability.transient_stability_study import (
                TransientStabilityStudy,
            )
            from studies.pv_voltage.pv_stability_study import PVStabilityStudy
            from config.settings import AESO

            # Initialise PSS/E
            logger.info("Initialising PSS/E (version %d)…", PSSE_VERSION)
            psse = PSSEInterface(
                psse_path=PSSE_PATH,
                psse_version=PSSE_VERSION,
                mock=False,
            )
            psse.initialize()
            logger.info("PSS/E initialised.")

            output_dir   = self.project.output_dir
            results_dir  = os.path.join(output_dir, "results")
            plots_dir    = os.path.join(output_dir, "plots")
            reports_dir  = os.path.join(output_dir, "reports")

            # Filter project scenarios to selected only
            scenarios_to_run = [
                sc for sc in self.project.scenarios
                if sc.scenario_name in selected_scenarios
            ]

            for sc in scenarios_to_run:
                logger.info("=" * 60)
                logger.info("Scenario: %s", sc.scenario_name)
                logger.info("=" * 60)

                if "power_flow" in selected_studies:
                    logger.info("Running Power Flow…")
                    study = PowerFlowStudy(
                        psse,
                        scenario_label          = sc.scenario_name,
                        project                 = self.project,
                        season_label            = f"{sc.year} {sc.season}",
                        voltage_min             = AESO["voltage_min_pu"],
                        voltage_max             = AESO["voltage_max_pu"],
                        voltage_min_contingency = AESO["voltage_min_contingency"],
                        voltage_max_contingency = AESO["voltage_max_contingency"],
                        thermal_limit_pct       = AESO["thermal_limit_pct"],
                    )
                    study.run(sav_path)
                    study.save_results(results_dir, plots_dir, reports_dir)
                    logger.info("Power Flow complete.")

                if "short_circuit" in selected_studies:
                    logger.info("Running Short Circuit…")
                    # Filter to SC target substations if bus numbers are filled
                    sc_filter = [
                        s.bus_no for s in self.project.sc_substations
                        if s.bus_no is not None
                    ] or None
                    study = ShortCircuitStudy(
                        psse,
                        scenario_label       = sc.scenario_name,
                        max_fault_current_ka = AESO["max_fault_current_ka"],
                        bus_filter           = sc_filter,
                    )
                    study.run(sav_path)
                    study.save_results(results_dir, plots_dir, reports_dir)
                    logger.info("Short Circuit complete.")

                if "transient" in selected_studies:
                    matrix = self.project.get_study_matrix(sc.scenario_name)
                    needs_ts = (matrix and (
                        matrix.transient_cat_a
                        or matrix.transient_cat_b
                        or matrix.transient_conditional
                    ))
                    if needs_ts:
                        logger.info("Running Transient Stability…")
                        study = TransientStabilityStudy(
                            psse, self.project,
                            scenario_label           = sc.scenario_name,
                            sim_duration_s           = AESO.get("ts_sim_duration_s", 10.0),
                            fault_apply_time_s       = AESO.get("ts_fault_apply_s", 1.0),
                            rotor_angle_limit_deg    = AESO["rotor_angle_limit_deg"],
                            voltage_recovery_pu      = AESO["voltage_recovery_pu"],
                            voltage_recovery_window_s= AESO["voltage_recovery_time_s"],
                        )
                        study.run(sav_path)
                        study.save_results(results_dir, plots_dir, reports_dir)
                        logger.info("Transient Stability complete.")
                    else:
                        logger.info(
                            "Transient Stability not required for %s "
                            "(Study Matrix).", sc.scenario_name
                        )

                if "voltage_stability" in selected_studies:
                    matrix = self.project.get_study_matrix(sc.scenario_name)
                    needs_vs = (matrix and (
                        matrix.volt_stability_cat_a
                        or matrix.volt_stability_cat_b
                    ))
                    if needs_vs:
                        if (self.project.info.source_bus_number is None
                                or self.project.info.poi_bus_number is None):
                            logger.warning(
                                "PV Stability skipped for %s — "
                                "Source Bus or POI Bus not set. "
                                "Fill Bus_Numbers sheet and Project_Info.",
                                sc.scenario_name
                            )
                        else:
                            logger.info("Running PV Voltage Stability…")
                            study = PVStabilityStudy(
                                psse, self.project,
                                scenario_label = sc.scenario_name,
                                v_min_cat_a    = AESO["pv_cat_a_v_min"],
                                v_min_cat_b    = AESO["pv_cat_b_v_min"],
                            )
                            study.run(sav_path)
                            study.save_results(results_dir, plots_dir, reports_dir)
                            logger.info("PV Voltage Stability complete.")
                    else:
                        logger.info(
                            "PV Stability not required for %s "
                            "(Study Matrix).", sc.scenario_name
                        )

            logger.info("=" * 60)
            logger.info("All selected studies complete.")
            logger.info("Output: %s", output_dir)

            # Re-enable run button and open output folder button
            self.root.after(0, self._on_studies_complete, output_dir)

        except Exception as exc:
            logger.error("Study run failed: %s", exc, exc_info=True)
            self.root.after(0, self._on_studies_failed, str(exc))

    def _on_studies_complete(self, output_dir: str):
        self.btn_run.configure(state="normal", text="▶  RUN STUDIES")
        self.btn_open_output.configure(state="normal")
        self.project.output_dir = output_dir
        messagebox.showinfo(
            "Studies Complete",
            f"All selected studies finished successfully.\n\n"
            f"Output folder:\n{output_dir}"
        )

    def _on_studies_failed(self, error_msg: str):
        self.btn_run.configure(state="normal", text="▶  RUN STUDIES")
        messagebox.showerror(
            "Study Run Failed",
            f"An error occurred during the study run:\n\n{error_msg}\n\n"
            "Check the log window for details."
        )

    def _open_output_folder(self):
        """Open the output folder in Windows Explorer."""
        output_dir = (self.project.output_dir
                      if self.project else self.var_project_dir.get())
        if output_dir and os.path.isdir(output_dir):
            subprocess.Popen(f'explorer "{output_dir}"')
        else:
            messagebox.showwarning(
                "Not Found",
                f"Output folder does not exist yet:\n{output_dir}"
            )

    # ── Logging ───────────────────────────────────────────────────────────────

    def _setup_logging(self):
        handler = QueueHandler(self.log_queue)
        handler.setFormatter(
            logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                               datefmt="%H:%M:%S")
        )
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.INFO)

    def _log(self, message: str):
        """Direct log message (bypasses logging system)."""
        self.log_queue.put(message)

    def _poll_log_queue(self):
        """Periodically drain the log queue into the log Text widget."""
        try:
            while True:
                record = self.log_queue.get_nowait()
                self._append_log(record)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _append_log(self, message: str):
        self.log_text.configure(state="normal")
        # Choose tag based on level keyword in message
        tag = "INFO"
        for level in ("CRITICAL", "ERROR", "WARNING"):
            if level in message:
                tag = level
                break
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


# ── Entry point ───────────────────────────────────────────────────────────────

def launch():
    """Launch the AESO Study Automation GUI."""
    try:
        import tkinter.simpledialog  # ensure available
    except ImportError:
        print("tkinter not available. Install Python with Tk support.")
        sys.exit(1)

    root = tk.Tk()
    app  = AESOStudyGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch()
