"""
gui/main_window.py (v2 — improved)
-----------------------------------
Improvements over v1:
  - Resizable left / right panes  (ttk.PanedWindow)
  - Left panel scrolls vertically so nothing is ever clipped
  - Responsive window: log panel, notebook and panels all grow with window
  - Status bar at the bottom showing current state
  - Progress bar while studies are running (indeterminate)
  - Auto-scroll toggle on the log panel
  - Results / Plots / Reports  split into 3 separate Open-buttons
    (was a single "Open Output Folder" that opened the wrong level)
  - New "Output Files" tab in the notebook that auto-refreshes after a run,
    shows files in results / plots / reports with size + date,
    and lets you double-click to open any file directly
  - Cleaner card-style section layout on the left panel
  - Better typography hierarchy and spacing
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

# ── Ensure project root is on path ────────────────────────────────────────────
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from project_io.excel_reader import ExcelReader, ExcelReaderError
from project_io.excel_writer import ExcelWriter
from project_io.project_data import ProjectData

logger = logging.getLogger(__name__)

# ── AESO Colour Palette ───────────────────────────────────────────────────────
C_DARK_BLUE   = "#003865"
C_MID_BLUE    = "#1F5FA6"
C_LIGHT_BLUE  = "#D6E4F0"
C_ORANGE      = "#E87722"
C_GREEN       = "#00853E"
C_GREEN_DARK  = "#005f2e"
C_RED         = "#C8102E"
C_WHITE       = "#FFFFFF"
C_LIGHT_GREY  = "#F0F2F5"
C_MID_GREY    = "#D0D3D8"
C_DARK_GREY   = "#4A4E54"
C_YELLOW      = "#FFFBE6"
C_BG          = "#F4F6F9"
C_SURFACE     = "#FFFFFF"
C_BORDER      = "#DDE1E7"
C_TEXT        = "#1C2B3A"
C_TEXT_MUTED  = "#6B7280"

# Sheet display names mapped to ProjectData field names
SHEET_TABS = [
    ("Project Info",      "info"),
    ("Scenarios",         "scenarios"),
    ("Study Matrix",      "study_matrix"),
    ("Conv Gen",          "conv_gen"),
    ("Renewables",        "renewables"),
    ("Intertie Flows",    "intertie_flows"),
    ("TS Contingencies",  "ts_contingencies"),
    ("SC Substations",    "sc_substations"),
    ("PV Contingencies",  "pv_contingencies"),
    ("Bus Numbers",       "bus_numbers"),
    ("Output Files",      "_output"),          # ← NEW: file browser tab
]


class QueueHandler(logging.Handler):
    """Routes log records into a thread-safe queue for GUI display."""
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord):
        self.log_queue.put(self.format(record))


class AESOStudyGUI:
    """Main application window — responsive layout v2."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.project: Optional[ProjectData] = None
        self.log_queue = queue.Queue()
        self._run_thread: Optional[threading.Thread] = None
        self._auto_scroll = tk.BooleanVar(value=True)

        # StringVars
        self.var_project_dir = tk.StringVar()
        self.var_sav_path    = tk.StringVar()
        self.var_excel_path  = tk.StringVar()
        self.var_status      = tk.StringVar(value="Ready — no project loaded")

        # Study checkboxes
        self.study_vars: Dict[str, tk.BooleanVar] = {
            "power_flow":        tk.BooleanVar(value=True),
            "short_circuit":     tk.BooleanVar(value=True),
            "transient":         tk.BooleanVar(value=True),
            "voltage_stability": tk.BooleanVar(value=True),
        }
        self.study_labels = {
            "power_flow":        "Power Flow",
            "short_circuit":     "Short Circuit",
            "transient":         "Transient Stability",
            "voltage_stability": "PV Voltage Stability",
        }

        self._setup_root()
        self._build_ui()
        self._setup_logging()
        self._poll_log_queue()

    # ── Window & style setup ──────────────────────────────────────────────────
    def _setup_root(self):
        self.root.title("AESO Interconnection Study Automation Tool")
        self.root.geometry("1400x900")
        self.root.minsize(1100, 720)
        self.root.configure(bg=C_BG)
        # Row weights: header=0, body=1 (grows), log=0, status=0
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_rowconfigure(3, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame",      background=C_BG)
        style.configure("Card.TFrame", background=C_SURFACE,
                        relief="flat")

        style.configure("TLabel",
            background=C_BG, font=("Segoe UI", 10), foreground=C_TEXT)
        style.configure("Card.TLabel",
            background=C_SURFACE, font=("Segoe UI", 10), foreground=C_TEXT)
        style.configure("Muted.TLabel",
            background=C_BG, font=("Segoe UI", 9), foreground=C_TEXT_MUTED)

        style.configure("TButton",
            font=("Segoe UI", 9, "bold"),
            background=C_MID_BLUE, foreground=C_WHITE,
            borderwidth=0, focusthickness=0, padding=(10, 6), relief="flat")
        style.map("TButton",
            background=[("active", C_DARK_BLUE), ("disabled", C_MID_GREY)],
            foreground=[("disabled", C_DARK_GREY)])

        style.configure("Run.TButton",
            font=("Segoe UI", 12, "bold"),
            background=C_GREEN, foreground=C_WHITE,
            padding=(16, 10), relief="flat")
        style.map("Run.TButton",
            background=[("active", C_GREEN_DARK), ("disabled", C_MID_GREY)])

        style.configure("OutDir.TButton",
            font=("Segoe UI", 9),
            background="#374151", foreground=C_WHITE,
            padding=(6, 5), relief="flat")
        style.map("OutDir.TButton",
            background=[("active", C_DARK_GREY), ("disabled", C_MID_GREY)])

        style.configure("Ghost.TButton",
            font=("Segoe UI", 9),
            background=C_BG, foreground=C_MID_BLUE,
            borderwidth=1, relief="solid", padding=(8, 5))
        style.map("Ghost.TButton",
            background=[("active", C_LIGHT_BLUE)])

        style.configure("TNotebook",     background=C_BG, borderwidth=0)
        style.configure("TNotebook.Tab",
            font=("Segoe UI", 9), padding=(10, 5),
            background=C_MID_GREY, foreground=C_DARK_GREY)
        style.map("TNotebook.Tab",
            background=[("selected", C_DARK_BLUE)],
            foreground=[("selected", C_WHITE)])

        style.configure("Treeview",
            font=("Segoe UI", 9), rowheight=24,
            background=C_WHITE, fieldbackground=C_WHITE,
            foreground=C_TEXT)
        style.configure("Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background=C_DARK_BLUE, foreground=C_WHITE)
        style.map("Treeview",
            background=[("selected", C_LIGHT_BLUE)])

        style.configure("TLabelframe",
            background=C_SURFACE, relief="flat",
            borderwidth=1, bordercolor=C_BORDER)
        style.configure("TLabelframe.Label",
            background=C_SURFACE,
            font=("Segoe UI", 10, "bold"),
            foreground=C_DARK_BLUE)

        style.configure("TCheckbutton",
            background=C_SURFACE, font=("Segoe UI", 10),
            foreground=C_TEXT)
        style.configure("TEntry",
            font=("Segoe UI", 9), padding=(6, 4))
        style.configure("TSeparator", background=C_BORDER)
        style.configure("Accent.Horizontal.TProgressbar",
            troughcolor=C_LIGHT_GREY,
            background=C_MID_BLUE,
            borderwidth=0, thickness=5)

    # ── Main UI construction ──────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_body()
        self._build_log_panel()
        self._build_status_bar()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C_DARK_BLUE)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_columnconfigure(1, weight=1)

        logo = tk.Frame(hdr, bg=C_ORANGE)
        logo.grid(row=0, column=0, padx=(16, 0), pady=10)
        tk.Label(logo, text="  AESO  ",
            bg=C_ORANGE, fg=C_WHITE,
            font=("Segoe UI", 15, "bold"),
            padx=8, pady=6).pack()

        tk.Label(hdr,
            text="Interconnection Study Automation Tool",
            bg=C_DARK_BLUE, fg=C_WHITE,
            font=("Segoe UI", 13, "bold")).grid(
                row=0, column=1, sticky="w", padx=14)

        right_hdr = tk.Frame(hdr, bg=C_DARK_BLUE)
        right_hdr.grid(row=0, column=2, sticky="e", padx=16)

        self.lbl_project_title = tk.Label(
            right_hdr, text="No project loaded",
            bg=C_DARK_BLUE, fg=C_LIGHT_BLUE,
            font=("Segoe UI", 10))
        self.lbl_project_title.pack(side="right", padx=(8, 0))

        tk.Label(right_hdr, text="v2.0",
            bg=C_DARK_BLUE, fg="#5580a0",
            font=("Segoe UI", 9)).pack(side="right")

    # ── Body: resizable paned window ──────────────────────────────────────────
    def _build_body(self):
        body = ttk.Frame(self.root)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)

        self.paned = ttk.PanedWindow(body, orient="horizontal")
        self.paned.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        # ── Left side: scrollable canvas ──────────────────────────────────────
        left_outer = tk.Frame(self.paned, bg=C_BG, width=340)
        left_outer.grid_propagate(False)
        left_outer.grid_rowconfigure(0, weight=1)
        left_outer.grid_columnconfigure(0, weight=1)

        left_canvas = tk.Canvas(left_outer, bg=C_BG, highlightthickness=0)
        left_scroll  = ttk.Scrollbar(left_outer, orient="vertical",
                                     command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_canvas.grid(row=0, column=0, sticky="nsew")
        left_scroll.grid(row=0, column=1, sticky="ns")

        self.left_frame = ttk.Frame(left_canvas)
        self.left_frame.grid_columnconfigure(0, weight=1)
        _win = left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")

        def _on_frame_configure(e):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _on_canvas_configure(e):
            left_canvas.itemconfig(_win, width=e.width)

        self.left_frame.bind("<Configure>", _on_frame_configure)
        left_canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse-wheel scroll (Windows)
        def _on_mousewheel(e):
            left_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        left_canvas.bind("<Enter>",  lambda e: left_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        left_canvas.bind("<Leave>",  lambda e: left_canvas.unbind_all("<MouseWheel>"))

        self.paned.add(left_outer, weight=0)

        # ── Right side: notebook ──────────────────────────────────────────────
        right = ttk.Frame(self.paned)
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)
        self.paned.add(right, weight=1)

        # Build sections
        self._build_project_setup(self.left_frame)
        self._build_validation_panel(self.left_frame)
        self._build_run_panel(self.left_frame)
        self._build_data_tabs(right)

    # ── Left panel: Project Setup ─────────────────────────────────────────────
    def _build_project_setup(self, parent):
        card = self._card(parent, " Project Setup", row=0)
        card.grid_columnconfigure(1, weight=1)

        fields = [
            ("Project Folder:",    self.var_project_dir, self._browse_project_dir),
            ("SAV File:",          self.var_sav_path,    self._browse_sav),
            ("Study Scope Excel:", self.var_excel_path,  self._browse_excel),
        ]
        for i, (lbl, var, cmd) in enumerate(fields):
            pad_top = 10 if i == 0 else 4
            ttk.Label(card, text=lbl, style="Card.TLabel").grid(
                row=i, column=0, sticky="w", padx=(10, 6), pady=(pad_top, 4))
            ttk.Entry(card, textvariable=var, state="readonly").grid(
                row=i, column=1, sticky="ew", padx=4, pady=(pad_top, 4))
            ttk.Button(card, text="Browse…", command=cmd, width=8).grid(
                row=i, column=2, padx=(4, 10), pady=(pad_top, 4))

        btn_row = ttk.Frame(card, style="Card.TFrame")
        btn_row.grid(row=3, column=0, columnspan=3, sticky="ew", padx=10, pady=(4, 10))
        btn_row.grid_columnconfigure((0, 1, 2), weight=1)
        for col, (text, cmd) in enumerate([
            ("New Project", self._new_project),
            ("Load Excel",  self._load_excel),
            ("Save Excel",  self._save_excel),
        ]):
            ttk.Button(btn_row, text=text, command=cmd).grid(
                row=0, column=col, sticky="ew",
                padx=(0 if col == 0 else 3, 3 if col < 2 else 0))

    # ── Left panel: Validation ────────────────────────────────────────────────
    def _build_validation_panel(self, parent):
        card = self._card(parent, " Validation", row=1)
        card.grid_columnconfigure(0, weight=1)

        self.warn_text = tk.Text(
            card, height=7, wrap="word",
            bg=C_YELLOW, fg=C_TEXT,
            font=("Segoe UI", 9),
            borderwidth=0, state="disabled", relief="flat", padx=6, pady=4)
        warn_scroll = ttk.Scrollbar(card, command=self.warn_text.yview)
        self.warn_text.configure(yscrollcommand=warn_scroll.set)
        self.warn_text.grid(row=0, column=0, sticky="ew", padx=(10, 0), pady=(8, 4))
        warn_scroll.grid(row=0, column=1, sticky="ns", pady=(8, 4), padx=(0, 10))

        self.warn_text.tag_configure("ok",       foreground=C_GREEN, font=("Segoe UI", 9, "bold"))
        self.warn_text.tag_configure("critical",  foreground=C_RED,   font=("Segoe UI", 9, "bold"))
        self.warn_text.tag_configure("warning",   foreground="#7a5000")

        ttk.Button(card, text="Validate Now", command=self._run_validation).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))

    # ── Left panel: Run Studies ───────────────────────────────────────────────
    def _build_run_panel(self, parent):
        card = self._card(parent, " Run Studies", row=2)
        card.grid_columnconfigure(0, weight=1)

        # Study type checkboxes
        for i, (key, label) in enumerate(self.study_labels.items()):
            ttk.Checkbutton(card, text=label, variable=self.study_vars[key]).grid(
                row=i, column=0, sticky="w",
                padx=14, pady=(8 if i == 0 else 2, 2))

        sep_row = len(self.study_labels)
        ttk.Separator(card).grid(row=sep_row, column=0, sticky="ew", padx=10, pady=6)

        # Scenario filter
        ttk.Label(card, text="Scenarios to run:", style="Card.TLabel").grid(
            row=sep_row + 1, column=0, sticky="w", padx=14, pady=(0, 3))

        sc_frame = ttk.Frame(card)
        sc_frame.grid(row=sep_row + 2, column=0, sticky="ew", padx=10, pady=(0, 4))
        sc_frame.grid_columnconfigure(0, weight=1)

        self.scenario_listbox = tk.Listbox(
            sc_frame, selectmode="multiple", height=5,
            font=("Segoe UI", 9),
            bg=C_WHITE, fg=C_TEXT,
            selectbackground=C_LIGHT_BLUE, selectforeground=C_DARK_BLUE,
            borderwidth=1, relief="solid", activestyle="none")
        sc_scroll = ttk.Scrollbar(sc_frame, command=self.scenario_listbox.yview)
        self.scenario_listbox.configure(yscrollcommand=sc_scroll.set)
        self.scenario_listbox.grid(row=0, column=0, sticky="ew")
        sc_scroll.grid(row=0, column=1, sticky="ns")

        ttk.Button(card, text="Select All Scenarios",
            command=self._select_all_scenarios,
            style="Ghost.TButton").grid(
                row=sep_row + 3, column=0, sticky="ew", padx=10, pady=(0, 6))

        # Progress bar (hidden until run)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            card, variable=self.progress_var,
            mode="indeterminate",
            style="Accent.Horizontal.TProgressbar")
        self.progress_bar.grid(row=sep_row + 4, column=0, sticky="ew", padx=10, pady=(0, 4))
        self.progress_bar.grid_remove()

        # RUN button
        self.btn_run = ttk.Button(
            card, text="▶  RUN STUDIES",
            style="Run.TButton",
            command=self._run_studies,
            state="disabled")
        self.btn_run.grid(row=sep_row + 5, column=0, sticky="ew", padx=10, pady=(4, 8))

        ttk.Separator(card).grid(row=sep_row + 6, column=0, sticky="ew", padx=10, pady=(0, 6))

        # Output sub-folder buttons (3-up)
        ttk.Label(card, text="Open output folder:", style="Card.TLabel").grid(
            row=sep_row + 7, column=0, sticky="w", padx=14, pady=(0, 4))

        out_row = ttk.Frame(card)
        out_row.grid(row=sep_row + 8, column=0, sticky="ew", padx=10, pady=(0, 10))
        out_row.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_open_results = ttk.Button(
            out_row, text="Results",
            style="OutDir.TButton",
            command=lambda: self._open_subfolder("results"),
            state="disabled")
        self.btn_open_results.grid(row=0, column=0, sticky="ew", padx=(0, 2))

        self.btn_open_plots = ttk.Button(
            out_row, text="Plots",
            style="OutDir.TButton",
            command=lambda: self._open_subfolder("plots"),
            state="disabled")
        self.btn_open_plots.grid(row=0, column=1, sticky="ew", padx=2)

        self.btn_open_reports = ttk.Button(
            out_row, text="Reports",
            style="OutDir.TButton",
            command=lambda: self._open_subfolder("reports"),
            state="disabled")
        self.btn_open_reports.grid(row=0, column=2, sticky="ew", padx=(2, 0))

    # ── Right panel: Notebook ─────────────────────────────────────────────────
    def _build_data_tabs(self, parent):
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self.tab_trees: Dict[str, ttk.Treeview] = {}

        for display_name, field_name in SHEET_TABS:
            tab = ttk.Frame(self.notebook)
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            self.notebook.add(tab, text=f"  {display_name}  ")

            if field_name == "_output":
                self._build_output_files_tab(tab)
            else:
                tree = self._make_treeview(tab)
                self.tab_trees[field_name] = tree

        self._show_placeholder_tabs()

    def _build_output_files_tab(self, parent):
        """Output files browser — shows results/plots/reports tree after run."""
        outer = ttk.Frame(parent)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.grid_rowconfigure(1, weight=1)
        outer.grid_columnconfigure(0, weight=1)

        # Toolbar
        toolbar = ttk.Frame(outer)
        toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        ttk.Label(toolbar, text="Generated output files — double-click a file to open it:").pack(side="left")
        ttk.Button(toolbar, text="↻  Refresh",
            command=self._refresh_output_tab,
            style="Ghost.TButton").pack(side="right")
        ttk.Button(toolbar, text="Open Folder",
            command=lambda: self._open_subfolder(""),
            style="Ghost.TButton").pack(side="right", padx=(0, 6))

        tree_frame = ttk.Frame(outer)
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.output_tree = ttk.Treeview(
            tree_frame, show="tree headings",
            columns=("size", "modified"), selectmode="browse")
        self.output_tree.heading("#0",       text="File / Folder")
        self.output_tree.heading("size",     text="Size")
        self.output_tree.heading("modified", text="Modified")
        self.output_tree.column("#0",       width=380, minwidth=200)
        self.output_tree.column("size",     width=80,  minwidth=60, anchor="e")
        self.output_tree.column("modified", width=160, minwidth=100)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.output_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.output_tree.xview)
        self.output_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.output_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Double-click opens the file
        self.output_tree.bind("<Double-1>", self._on_output_file_open)

        # Placeholder message shown before any run
        self._output_placeholder = tk.Label(
            outer,
            text="No output files yet.\nRun studies first — files will appear here automatically.",
            bg=C_BG, fg=C_TEXT_MUTED,
            font=("Segoe UI", 10),
            justify="center")
        self._output_placeholder.place(relx=0.5, rely=0.5, anchor="center")

    def _refresh_output_tab(self):
        """Populate the output files tree from the project output directory."""
        self.output_tree.delete(*self.output_tree.get_children())

        if not self.project:
            return
        output_dir = getattr(self.project, "output_dir", None)
        if not output_dir or not os.path.isdir(output_dir):
            return

        self._output_placeholder.place_forget()

        import datetime
        has_files = False
        for subfolder in ("results", "plots", "reports"):
            path = os.path.join(output_dir, subfolder)
            if not os.path.isdir(path):
                continue
            files = sorted(os.listdir(path))
            if not files:
                continue
            has_files = True
            folder_node = self.output_tree.insert(
                "", "end",
                text=f"  {subfolder}/",
                values=("", ""),
                open=True)
            for fname in files:
                fpath = os.path.join(path, fname)
                if not os.path.isfile(fpath):
                    continue
                stat  = os.stat(fpath)
                size  = self._fmt_size(stat.st_size)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d  %H:%M")
                if fname.endswith(".xlsx"):
                    icon = "  "
                elif fname.endswith((".png", ".pdf")):
                    icon = "  "
                else:
                    icon = "  "
                # Store the filepath in tags for double-click open
                self.output_tree.insert(
                    folder_node, "end",
                    text=f"    {icon}{fname}",
                    values=(size, mtime),
                    tags=(fpath,))

        if not has_files:
            self._output_placeholder.place(relx=0.5, rely=0.5, anchor="center")

    @staticmethod
    def _fmt_size(size: int) -> str:
        if size < 1024:    return f"{size} B"
        if size < 1048576: return f"{size / 1024:.1f} KB"
        return f"{size / 1048576:.1f} MB"

    def _on_output_file_open(self, _event):
        """Open a file from the output tree on double-click."""
        item = self.output_tree.focus()
        if not item:
            return
        tags = self.output_tree.item(item, "tags")
        if tags:
            fpath = tags[0]
            if os.path.isfile(fpath):
                try:
                    os.startfile(fpath)
                except Exception:
                    subprocess.Popen(["explorer", "/select,", fpath])

    @staticmethod
    def _make_treeview(parent) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        vsb  = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
        hsb  = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        return tree

    def _show_placeholder_tabs(self):
        for field_name, tree in self.tab_trees.items():
            tree["columns"] = ("message",)
            tree.heading("message", text="Status")
            tree.column("message", width=500, anchor="w")
            tree.delete(*tree.get_children())
            tree.insert("", "end", values=(
                "No project loaded — browse to a Study Scope Excel file and click Load Excel.",))

    # ── Log panel ─────────────────────────────────────────────────────────────
    def _build_log_panel(self):
        log_outer = ttk.LabelFrame(self.root, text=" Study Log")
        log_outer.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 0))
        log_outer.grid_columnconfigure(0, weight=1)

        # Toolbar
        toolbar = tk.Frame(log_outer, bg=C_LIGHT_GREY)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 2))

        tk.Checkbutton(toolbar, text="Auto-scroll",
            variable=self._auto_scroll,
            bg=C_LIGHT_GREY, fg=C_DARK_GREY,
            font=("Segoe UI", 8),
            activebackground=C_LIGHT_GREY,
            relief="flat", bd=0).pack(side="left")
        tk.Button(toolbar, text="Clear",
            bg=C_MID_GREY, fg=C_DARK_GREY,
            font=("Segoe UI", 8), relief="flat",
            padx=6, pady=2,
            command=self._clear_log).pack(side="right")

        self.log_text = tk.Text(
            log_outer, height=9, wrap="word",
            bg="#14181f", fg="#c8d8e8",
            font=("Consolas", 9),
            borderwidth=0, insertbackground=C_WHITE,
            state="disabled", padx=8, pady=4)
        log_scroll = ttk.Scrollbar(log_outer, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=1, column=0, sticky="ew", padx=(6, 0), pady=(0, 6))
        log_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 6), padx=(0, 6))

        self.log_text.tag_configure("INFO",     foreground="#c8d8e8")
        self.log_text.tag_configure("WARNING",  foreground=C_ORANGE)
        self.log_text.tag_configure("ERROR",    foreground=C_RED)
        self.log_text.tag_configure("CRITICAL", foreground=C_RED)

    # ── Status bar ────────────────────────────────────────────────────────────
    def _build_status_bar(self):
        bar = tk.Frame(self.root, bg=C_DARK_BLUE, height=26)
        bar.grid(row=3, column=0, sticky="ew")
        bar.grid_propagate(False)
        bar.grid_columnconfigure(0, weight=1)

        self.status_label = tk.Label(
            bar, textvariable=self.var_status,
            bg=C_DARK_BLUE, fg=C_LIGHT_BLUE,
            font=("Segoe UI", 9), anchor="w", padx=12)
        self.status_label.grid(row=0, column=0, sticky="ew")

        tk.Label(bar, text="AESO Automation  |  University of Calgary",
            bg=C_DARK_BLUE, fg="#4a6070",
            font=("Segoe UI", 8), anchor="e", padx=12).grid(row=0, column=1, sticky="e")

    # ── Card helper ───────────────────────────────────────────────────────────
    def _card(self, parent, title: str, row: int) -> ttk.Frame:
        lf = ttk.LabelFrame(parent, text=title)
        lf.grid(row=row, column=0, sticky="ew", padx=4, pady=(0, 8))
        lf.grid_columnconfigure(0, weight=1)
        return lf

    # ── File pickers ──────────────────────────────────────────────────────────
    def _browse_project_dir(self):
        path = filedialog.askdirectory(title="Select Project Folder")
        if not path:
            return
        self.var_project_dir.set(path)
        for f in os.listdir(path):
            if f.endswith(".sav") and not self.var_sav_path.get():
                self.var_sav_path.set(os.path.join(path, f))
            if ("study_scope" in f.lower() or "scope_data" in f.lower()) and f.endswith(".xlsx"):
                self.var_excel_path.set(os.path.join(path, f))
        cases_dir = os.path.join(path, "cases")
        if os.path.isdir(cases_dir) and not self.var_sav_path.get():
            for f in os.listdir(cases_dir):
                if f.endswith(".sav"):
                    self.var_sav_path.set(os.path.join(cases_dir, f))
                    break

    def _browse_sav(self):
        path = filedialog.askopenfilename(
            title="Select PSS/E Case File",
            filetypes=[("PSS/E Case", "*.sav"), ("All Files", "*.*")])
        if path:
            self.var_sav_path.set(path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Study Scope Excel File",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")])
        if path:
            self.var_excel_path.set(path)

    # ── Project operations ────────────────────────────────────────────────────
    def _new_project(self):
        parent_dir = filedialog.askdirectory(title="Select Parent Folder for New Project")
        if not parent_dir:
            return
        project_num = tk.simpledialog.askstring(
            "New Project", "Enter project number (e.g. P2611):", parent=self.root)
        if not project_num:
            return
        project_dir   = os.path.join(parent_dir, project_num)
        excel_path    = os.path.join(project_dir, "study_scope_data.xlsx")
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "templates", "study_scope_template.xlsx")
        if not os.path.isfile(template_path):
            messagebox.showerror("Template Not Found", f"Could not find:\n{template_path}")
            return
        try:
            os.makedirs(os.path.join(project_dir, "cases"),  exist_ok=True)
            os.makedirs(os.path.join(project_dir, "output"), exist_ok=True)
            ExcelWriter.new_project_file(template_path, excel_path)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self.var_project_dir.set(project_dir)
        self.var_excel_path.set(excel_path)
        self._log(f"New project created: {project_dir}")
        self._set_status(f"New project: {project_num}")
        messagebox.showinfo("Project Created",
            f"Project folder created:\n{project_dir}\n\n"
            "Fill study_scope_data.xlsx from your Study Scope PDF,\n"
            "place your .sav file in the cases/ subfolder,\n"
            "then click Load Excel.")

    def _load_excel(self):
        excel_path = self.var_excel_path.get()
        if not excel_path:
            excel_path = filedialog.askopenfilename(
                title="Select Study Scope Excel",
                filetypes=[("Excel Workbook", "*.xlsx")])
        if not excel_path:
            return
        self.var_excel_path.set(excel_path)
        try:
            reader = ExcelReader(excel_path)
            self.project = reader.read()
        except ExcelReaderError as exc:
            messagebox.showerror("Load Error", str(exc))
            return
        if self.project.info.sav_file_path:
            self.var_sav_path.set(self.project.info.sav_file_path)
        if not self.var_project_dir.get():
            self.var_project_dir.set(os.path.dirname(excel_path))
        self.project.project_dir = self.var_project_dir.get()
        self.project.output_dir  = os.path.join(self.project.project_dir, "output")

        self._refresh_all_tabs()
        self._refresh_scenario_list()
        self._run_validation()
        self._refresh_output_tab()

        title = f"{self.project.info.project_number} — {self.project.info.project_name}"
        self.lbl_project_title.configure(text=title)
        self.btn_run.configure(state="normal")
        self._set_status(f"Loaded: {self.project.info.project_number}")
        self._log(f"Loaded: {self.project}")

    def _save_excel(self):
        if self.project is None:
            messagebox.showwarning("No Data", "Load a project first.")
            return
        excel_path = self.var_excel_path.get()
        if not excel_path:
            messagebox.showwarning("No File", "No Excel file path set.")
            return
        self.project.info.sav_file_path = self.var_sav_path.get()
        try:
            ExcelWriter(excel_path).write(self.project)
            self._log(f"Saved: {excel_path}")
            self._set_status("Excel saved successfully")
            messagebox.showinfo("Saved", f"Study scope data saved to:\n{excel_path}")
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    # ── Tab population ────────────────────────────────────────────────────────
    def _refresh_all_tabs(self):
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
            ("In-Service Date",         info.in_service_date),
            ("Generation Type",         info.generation_type),
            ("MARP (MW)",               info.marp_mw),
            ("Max Capability (MW)",     info.max_capability_mw),
            ("Connection Voltage (kV)", info.connection_voltage_kv),
            ("POC Substation",          info.poc_substation_name),
            ("Study Area Regions",      info.study_area_regions),
            ("SAV File Path",           info.sav_file_path),
            ("Source Bus Number",       info.source_bus_number or "⚠ Not set"),
            ("POI Bus Number",          info.poi_bus_number    or "⚠ Not set"),
            ("TS Fault Bus Number",     info.ts_fault_bus_number or "⚠ Not set"),
        ]
        for field, value in rows:
            tag = "missing" if str(value).startswith("⚠") else ""
            tree.insert("", "end", values=(field, value), tags=(tag,))
        tree.tag_configure("missing", foreground=C_ORANGE)

    def _populate_generic_tab(self, field_name: str, rows_data: tuple):
        tree = self.tab_trees[field_name]
        columns, rows = rows_data
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=max(80, len(col) * 9), anchor="w", minwidth=60)
        tree.delete(*tree.get_children())
        alt = False
        for row in rows:
            tree.insert("", "end", values=row, tags=("alt" if alt else "",))
            alt = not alt
        tree.tag_configure("alt", background=C_LIGHT_BLUE)

    # ── Row builders (identical to v1) ────────────────────────────────────────
    def _scenarios_rows(self):
        cols = ("No", "Year", "Season", "Dispatch", "Scenario Name",
                "Pre/Post", "Load (MW)", "Gen (MW)")
        rows = [(s.scenario_no, s.year, s.season, s.dispatch_cond,
                 s.scenario_name, s.pre_post, s.project_load_mw, s.project_gen_mw)
                for s in self.project.scenarios]
        return cols, rows

    def _matrix_rows(self):
        cols = ("Scenario", "PF-A", "PF-B", "VS-A", "VS-B",
                "TS-A", "TS-B", "TS-Cond", "MS-A", "MS-B", "SC-A")
        def _x(b): return "X" if b else ""
        rows = [(m.scenario_name,
                 _x(m.power_flow_cat_a), _x(m.power_flow_cat_b),
                 _x(m.volt_stability_cat_a), _x(m.volt_stability_cat_b),
                 _x(m.transient_cat_a), _x(m.transient_cat_b),
                 "X*" if m.transient_conditional else "",
                 _x(m.motor_starting_cat_a), _x(m.motor_starting_cat_b),
                 _x(m.short_circuit_cat_a))
                for m in self.project.study_matrix]
        return cols, rows

    def _conv_gen_rows(self):
        sl = sorted({k for g in self.project.conv_gen for k in g.dispatch_mw})
        cols = ("Facility", "Unit", "Bus No", "MC (MW)", "Area") + tuple(sl)
        rows = [(g.facility_name, g.unit_no, g.bus_no or "—", g.mc_mw, g.area_no)
                + tuple(g.dispatch_mw.get(s, "") for s in sl)
                for g in self.project.conv_gen]
        return cols, rows

    def _renewables_rows(self):
        sl = sorted({k for r in self.project.renewables for k in r.dispatch_mw})
        cols = ("Facility", "Type", "Bus No", "MC (MW)", "Area") + tuple(sl)
        rows = [(r.facility_name, r.gen_type, r.bus_no or "—", r.mc_mw, r.area_no)
                + tuple(r.dispatch_mw.get(s, "") for s in sl)
                for r in self.project.renewables]
        return cols, rows

    def _intertie_rows(self):
        if not self.project.intertie_flows:
            return ("Scenario",), []
        keys = sorted({k for f in self.project.intertie_flows for k in f.flows})
        cols = ("Scenario",) + tuple(keys)
        rows = [(f.scenario_name,) + tuple(f.flows.get(k, "") for k in keys)
                for f in self.project.intertie_flows]
        return cols, rows

    def _ts_cont_rows(self):
        cols = ("Contingency", "From Bus Name", "To Bus Name", "From Bus No", "To Bus No",
                "Ckt", "Fault Location", "Near End (cyc)", "Far End (cyc)")
        rows = [(c.contingency_name, c.from_bus_name, c.to_bus_name,
                 c.from_bus_no or "—", c.to_bus_no or "—",
                 c.circuit_id, c.fault_location, c.near_end_cycles, c.far_end_cycles)
                for c in self.project.ts_contingencies]
        return cols, rows

    def _sc_subs_rows(self):
        cols = ("Substation Name", "Bus No", "Notes")
        rows = [(s.substation_name, s.bus_no or "—", s.notes)
                for s in self.project.sc_substations]
        return cols, rows

    def _pv_cont_rows(self):
        cols = ("Contingency", "From Bus", "To Bus", "From No", "To No", "Ckt", "Category")
        rows = [(c.contingency_name, c.from_bus_name, c.to_bus_name,
                 c.from_bus_no or "—", c.to_bus_no or "—", c.circuit_id, c.category)
                for c in self.project.pv_contingencies]
        return cols, rows

    def _bus_num_rows(self):
        cols = ("Substation Name", "Bus Number", "Base kV", "Bus Type", "Notes")
        rows = [(b.substation_name, b.bus_number or "—", b.base_kv, b.bus_type, b.notes)
                for b in self.project.bus_numbers]
        return cols, rows

    # ── Validation ────────────────────────────────────────────────────────────
    def _run_validation(self):
        if self.project is None:
            return
        self.project.info.sav_file_path = self.var_sav_path.get()
        warnings = self.project.validate()
        self.warn_text.configure(state="normal")
        self.warn_text.delete("1.0", "end")
        if not warnings:
            self.warn_text.configure(bg="#f0fdf4")
            self.warn_text.insert("end", "✔  No issues found — ready to run.", "ok")
        else:
            self.warn_text.configure(bg=C_YELLOW)
            for w in warnings:
                tag = "critical" if "[CRITICAL]" in w else "warning"
                self.warn_text.insert("end", w + "\n", tag)
        self.warn_text.configure(state="disabled")
        sav_ok = bool(self.var_sav_path.get())
        self.btn_run.configure(state="normal" if (self.project and sav_ok) else "disabled")

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
        return [self.scenario_listbox.get(i) for i in self.scenario_listbox.curselection()]

    # ── Run studies ───────────────────────────────────────────────────────────
    def _run_studies(self):
        if self.project is None:
            messagebox.showwarning("No Project", "Load a project first.")
            return
        sav_path = self.var_sav_path.get()
        if not sav_path or not os.path.isfile(sav_path):
            messagebox.showerror("SAV Not Found",
                f"SAV file not found:\n{sav_path}\n\nBrowse to a valid .sav file.")
            return
        selected_scenarios = self._get_selected_scenarios()
        if not selected_scenarios:
            messagebox.showwarning("No Scenarios Selected",
                "Select at least one scenario in the Run Studies panel.")
            return
        selected_studies = [k for k, v in self.study_vars.items() if v.get()]
        if not selected_studies:
            messagebox.showwarning("No Studies Selected",
                "Check at least one study type to run.")
            return

        self.btn_run.configure(state="disabled", text="⏳  Running…")
        self.progress_bar.grid()
        self.progress_bar.start(12)
        self._set_status("Running studies… please wait")
        self.root.update()

        self.project.info.sav_file_path = sav_path
        self.project.project_dir = self.var_project_dir.get()
        self.project.output_dir  = os.path.join(self.project.project_dir, "output")

        self._run_thread = threading.Thread(
            target=self._run_studies_thread,
            args=(sav_path, selected_scenarios, selected_studies),
            daemon=True)
        self._run_thread.start()

    def _run_studies_thread(self, sav_path, selected_scenarios, selected_studies):
        try:
            from config.settings import PSSE_PATH, PSSE_VERSION, AESO
            from core.psse_interface import PSSEInterface
            from studies.power_flow.power_flow_study import PowerFlowStudy
            from studies.short_circuit.short_circuit_study import ShortCircuitStudy
            from studies.transient_stability.transient_stability_study import TransientStabilityStudy
            from studies.pv_voltage.pv_stability_study import PVStabilityStudy

            logger.info("Initialising PSS/E (version %d)…", PSSE_VERSION)
            psse = PSSEInterface(psse_path=PSSE_PATH, psse_version=PSSE_VERSION, mock=False)
            psse.initialize()
            logger.info("PSS/E initialised.")

            output_dir  = self.project.output_dir
            results_dir = os.path.join(output_dir, "results")
            plots_dir   = os.path.join(output_dir, "plots")
            reports_dir = os.path.join(output_dir, "reports")

            scenarios_to_run = [sc for sc in self.project.scenarios
                                if sc.scenario_name in selected_scenarios]

            for sc in scenarios_to_run:
                logger.info("=" * 60)
                logger.info("Scenario: %s", sc.scenario_name)
                logger.info("=" * 60)

                if "power_flow" in selected_studies:
                    logger.info("Running Power Flow…")
                    study = PowerFlowStudy(
                        psse, scenario_label=sc.scenario_name, project=self.project,
                        season_label=f"{sc.year} {sc.season}",
                        voltage_min=AESO["voltage_min_pu"], voltage_max=AESO["voltage_max_pu"],
                        voltage_min_contingency=AESO["voltage_min_contingency"],
                        voltage_max_contingency=AESO["voltage_max_contingency"],
                        thermal_limit_pct=AESO["thermal_limit_pct"])
                    study.run(sav_path)
                    study.save_results(results_dir, plots_dir, reports_dir)
                    logger.info("Power Flow complete.")

                if "short_circuit" in selected_studies:
                    logger.info("Running Short Circuit…")
                    sc_filter = [s.bus_no for s in self.project.sc_substations
                                 if s.bus_no is not None] or None
                    study = ShortCircuitStudy(
                        psse, scenario_label=sc.scenario_name,
                        max_fault_current_ka=AESO["max_fault_current_ka"],
                        bus_filter=sc_filter)
                    study.run(sav_path)
                    study.save_results(results_dir, plots_dir, reports_dir)
                    logger.info("Short Circuit complete.")

                if "transient" in selected_studies:
                    matrix = self.project.get_study_matrix(sc.scenario_name)
                    if matrix and (matrix.transient_cat_a or matrix.transient_cat_b
                                   or matrix.transient_conditional):
                        logger.info("Running Transient Stability…")
                        study = TransientStabilityStudy(
                            psse, self.project, scenario_label=sc.scenario_name,
                            sim_duration_s=AESO.get("ts_sim_duration_s", 10.0),
                            fault_apply_time_s=AESO.get("ts_fault_apply_s", 1.0),
                            rotor_angle_limit_deg=AESO["rotor_angle_limit_deg"],
                            voltage_recovery_pu=AESO["voltage_recovery_pu"],
                            voltage_recovery_window_s=AESO["voltage_recovery_time_s"])
                        study.run(sav_path)
                        study.save_results(results_dir, plots_dir, reports_dir)
                        logger.info("Transient Stability complete.")
                    else:
                        logger.info("Transient Stability not required for %s.", sc.scenario_name)

                if "voltage_stability" in selected_studies:
                    matrix = self.project.get_study_matrix(sc.scenario_name)
                    if matrix and (matrix.volt_stability_cat_a or matrix.volt_stability_cat_b):
                        if (self.project.info.source_bus_number
                                and self.project.info.poi_bus_number):
                            logger.info("Running PV Voltage Stability…")
                            study = PVStabilityStudy(
                                psse, self.project, scenario_label=sc.scenario_name,
                                v_min_cat_a=AESO["pv_cat_a_v_min"],
                                v_min_cat_b=AESO["pv_cat_b_v_min"])
                            study.run(sav_path)
                            study.save_results(results_dir, plots_dir, reports_dir)
                            logger.info("PV Voltage Stability complete.")
                        else:
                            logger.warning("PV Stability skipped for %s — "
                                           "Source Bus or POI Bus not set.", sc.scenario_name)
                    else:
                        logger.info("PV Stability not required for %s.", sc.scenario_name)

            logger.info("=" * 60)
            logger.info("All selected studies complete.  Output: %s", output_dir)
            self.root.after(0, self._on_studies_complete, output_dir)

        except Exception as exc:
            logger.error("Study run failed: %s", exc, exc_info=True)
            self.root.after(0, self._on_studies_failed, str(exc))

    def _on_studies_complete(self, output_dir: str):
        self.btn_run.configure(state="normal", text="▶  RUN STUDIES")
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self.project.output_dir = output_dir
        for btn in (self.btn_open_results, self.btn_open_plots, self.btn_open_reports):
            btn.configure(state="normal")
        self._set_status("✔  Studies completed successfully")
        self._refresh_output_tab()
        # Auto-switch to the Output Files tab
        self.notebook.select(len(SHEET_TABS) - 1)
        messagebox.showinfo("Studies Complete",
            f"All selected studies finished successfully.\n\nOutput folder:\n{output_dir}")

    def _on_studies_failed(self, error_msg: str):
        self.btn_run.configure(state="normal", text="▶  RUN STUDIES")
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        self._set_status("✘  Study run failed — see log for details")
        messagebox.showerror("Study Run Failed",
            f"An error occurred during the study run:\n\n{error_msg}\n\n"
            "Check the log window for details.")

    # ── Output folder helpers ─────────────────────────────────────────────────
    def _open_subfolder(self, subfolder: str):
        """Open a specific output sub-folder (results/plots/reports) in Explorer."""
        output_dir = (self.project.output_dir if self.project
                      else self.var_project_dir.get())
        if not output_dir:
            messagebox.showwarning("Not Found", "No output directory configured.")
            return
        target = os.path.join(output_dir, subfolder) if subfolder else output_dir
        if os.path.isdir(target):
            subprocess.Popen(f'explorer "{target}"')
        else:
            ans = messagebox.askyesno(
                "Folder Not Found",
                f"The folder does not exist yet:\n{target}\n\n"
                "Run studies first to generate output.\n\n"
                "Open the parent output folder instead?")
            if ans:
                os.makedirs(output_dir, exist_ok=True)
                subprocess.Popen(f'explorer "{output_dir}"')

    # ── Status bar helper ─────────────────────────────────────────────────────
    def _set_status(self, message: str):
        self.var_status.set(f"  {message}")

    # ── Logging ───────────────────────────────────────────────────────────────
    def _setup_logging(self):
        handler = QueueHandler(self.log_queue)
        handler.setFormatter(logging.Formatter(
            "%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%H:%M:%S"))
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.INFO)

    def _log(self, message: str):
        self.log_queue.put(message)

    def _poll_log_queue(self):
        try:
            while True:
                record = self.log_queue.get_nowait()
                self._append_log(record)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _append_log(self, message: str):
        self.log_text.configure(state="normal")
        tag = "INFO"
        for level in ("CRITICAL", "ERROR", "WARNING"):
            if level in message:
                tag = level
                break
        self.log_text.insert("end", message + "\n", tag)
        if self._auto_scroll.get():
            self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


# ── Entry point ────────────────────────────────────────────────────────────────
def launch():
    try:
        import tkinter.simpledialog  # noqa: F401
    except ImportError:
        print("tkinter not available. Install Python with Tk support.")
        sys.exit(1)
    root = tk.Tk()
    AESOStudyGUI(root)
    root.mainloop()


if __name__ == "__main__":
    launch()
