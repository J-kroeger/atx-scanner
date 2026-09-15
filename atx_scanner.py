#!/usr/bin/env python3
"""
ATX-Scanner - ANTEX File Summary Browser
==========================================
Desktop application to browse, filter, search, and export ANTEX antenna
calibration metadata. Built with ttkbootstrap (darkly theme).

Features:
  - Browse all parsed ANTEX records in a sortable table
  - Filter by antenna type, radome, method, institution, frequencies
  - Live search across all fields
  - Detail panel for selected records
  - Export filtered results to Excel or CSV
  - Load built-in database or scan external ATX folder
  - Standalone .exe packaging support (PyInstaller)

Author: Amr Fawzy, Institut für Erdmessung (IfE), Leibniz Universität Hannover
Date: 2026-04-01
"""

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import sqlite3
import os
import sys
import csv
import threading
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

# Shared opt-in dialog for the IfE software mailing list. Never let a missing
# copy stop the program from starting; the canonical file is in shared/.
try:
    import mailing_list
except Exception:
    mailing_list = None

# Version and release date, from this tool's own tool_version.py. The same pair
# sits in the README tag that the PCC-Suite launcher reads, so the launcher no
# longer has to guess a release from a file timestamp.
try:
    import tool_version
    import version_info
    VERSION_TEXT = version_info.about_line(tool_version)
except Exception:
    VERSION_TEXT = "Version: unknown"

try:
    from PIL import Image, ImageTk, ImageEnhance
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ── Resource path helper (works for dev and PyInstaller) ──────────────────────
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ── Column definitions ────────────────────────────────────────────────────────
COLUMNS = [
    ("antex_version",   "ANTEX Ver.",          70),
    ("antenna_type",    "Antenna Type",       200),
    ("serial_number",   "Serial No.",         100),
    ("radome",          "Radome",              70),
    ("method",          "Method",              80),
    ("institution",     "Institution",        160),
    ("num_calibrations","# Cal.",              50),
    ("calibration_date","Cal. Date",          100),
    ("dazi",            "DAZI",                50),
    ("zen1",            "ZEN1",                50),
    ("zen2",            "ZEN2",                50),
    ("dzen",            "DZEN",                50),
    ("num_frequencies", "# Freq.",             55),
    ("frequencies",     "Frequencies",        200),
    ("file_name",       "File Name",          220),
    ("relative_path",   "Relative Path",      300),
]

DB_COLUMNS = [c[0] for c in COLUMNS]
DISPLAY_NAMES = {c[0]: c[1] for c in COLUMNS}


# ── Main Application ──────────────────────────────────────────────────────────
class ATXExplorer(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        self.title("ATX-Scanner")
        self.geometry("1400x850")
        self.minsize(1000, 600)

        # State
        self.all_records = []
        self.filtered_records = []
        self.db_path = None
        self.sort_column = None
        self.sort_reverse = False
        self.logo_img = None
        self.luh_logo_img = None
        self._scan_cancel = False

        # Set window icon
        ico_path = resource_path(os.path.join("assets", "ife_logo.ico"))
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass

        self._build_ui()

        # Close confirmation
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_top_bar()
        self._build_toolbar()
        self._build_db_indicator()
        self._build_main_area()
        self._build_status_bar()

    def _build_top_bar(self):
        top = ttk.Frame(self, padding=(15, 8))
        top.pack(fill="x")

        # Left: IfE logo + title
        left = ttk.Frame(top)
        left.pack(side="left")

        logo_path = resource_path(os.path.join("assets", "ife_logo.png"))
        if HAS_PIL and os.path.exists(logo_path):
            try:
                pil_img = Image.open(logo_path)
                ar = pil_img.width / pil_img.height
                h = 45
                resized = pil_img.resize((int(h * ar), h), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(resized)
                ttk.Label(left, image=self.logo_img).pack(side="left", padx=(0, 12))
            except Exception:
                pass

        text_frame = ttk.Frame(left)
        text_frame.pack(side="left")
        ttk.Label(text_frame, text="ATX-Scanner", font=("Helvetica", 20, "bold")).pack(anchor="w")
        ttk.Label(text_frame, text="Institut f\u00fcr Erdmessung (IfE)", font=("Helvetica", 10), bootstyle="light").pack(anchor="w")

        # Right: LUH logo
        right = ttk.Frame(top)
        right.pack(side="right")

        luh_path = resource_path(os.path.join("assets", "luh_logo.png"))
        if HAS_PIL and os.path.exists(luh_path):
            try:
                pil_luh = Image.open(luh_path)
                h2 = 55
                ar2 = pil_luh.width / pil_luh.height
                resized_luh = pil_luh.resize((int(h2 * ar2), h2), Image.Resampling.LANCZOS)
                enhancer = ImageEnhance.Sharpness(resized_luh)
                final_luh = enhancer.enhance(1.5)
                self.luh_logo_img = ImageTk.PhotoImage(final_luh)
                ttk.Label(right, image=self.luh_logo_img).pack(side="right")
            except Exception:
                pass

        # Middle-right: info buttons
        btn_frame = ttk.Frame(top)
        btn_frame.pack(side="right", padx=20)
        ttk.Button(btn_frame, text="Contact", command=self._show_contact, bootstyle="outline-info").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="Information", command=self._show_about, bootstyle="outline-secondary").pack(side="left", padx=3)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

    def _build_toolbar(self):
        toolbar = ttk.Frame(self, padding=(10, 6))
        toolbar.pack(fill="x")

        # Data source buttons
        ttk.Button(toolbar, text="Load Database", command=self._load_database, bootstyle="info-outline").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Scan ATX Folder", command=self._scan_folder, bootstyle="success-outline").pack(side="left", padx=3)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)

        # Export buttons
        if HAS_OPENPYXL:
            ttk.Button(toolbar, text="Export to Excel", command=self._export_excel, bootstyle="warning-outline").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Export to CSV", command=self._export_csv, bootstyle="warning-outline").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Export Database", command=self._export_database, bootstyle="warning-outline").pack(side="left", padx=3)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)

        ttk.Button(toolbar, text="Reset Filters", command=self._reset_filters, bootstyle="secondary-outline").pack(side="left", padx=3)

        # Right side: quit
        ttk.Button(toolbar, text="Quit", command=self._on_closing, bootstyle="danger-outline").pack(side="right", padx=3)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

    def _build_db_indicator(self):
        """Build the database source indicator bar."""
        db_bar = ttk.Frame(self, padding=(10, 4))
        db_bar.pack(fill="x")

        ttk.Label(db_bar, text="Database:", font=("Helvetica", 9, "bold")).pack(side="left", padx=(0, 5))
        self.db_indicator_var = tk.StringVar(value="No database loaded")
        self.db_indicator_label = ttk.Label(db_bar, textvariable=self.db_indicator_var,
                                             font=("Consolas", 9), foreground="#4da6ff")
        self.db_indicator_label.pack(side="left", fill="x", expand=True)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

    def _build_main_area(self):
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=0, pady=0)

        # Left filter panel
        filter_panel = ttk.Frame(main, width=320)
        filter_panel.pack(side="left", fill="y", padx=(5, 0), pady=5)
        filter_panel.pack_propagate(False)
        self._build_filter_panel(filter_panel)

        # Vertical separator
        ttk.Separator(main, orient="vertical").pack(side="left", fill="y", padx=2)

        # Right: table + detail
        right_area = ttk.Frame(main)
        right_area.pack(side="left", fill="both", expand=True, padx=(0, 5), pady=5)

        # Paned window: table on top, detail on bottom
        paned = ttk.PanedWindow(right_area, orient="vertical")
        paned.pack(fill="both", expand=True)

        # Table frame
        table_frame = ttk.Frame(paned)
        paned.add(table_frame, weight=4)
        self._build_table(table_frame)

        # Detail frame
        detail_frame = ttk.LabelFrame(paned, text="Record Details", padding=10)
        paned.add(detail_frame, weight=1)
        self._build_detail_panel(detail_frame)

    def _build_filter_panel(self, parent):
        # Scrollable filter panel
        canvas = tk.Canvas(parent, highlightthickness=0, bg="#222222")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

        fp = scroll_frame  # shorthand

        # ── Search ──
        search_frame = ttk.LabelFrame(fp, text="Search", padding=8)
        search_frame.pack(fill="x", pady=(5, 5), padx=5)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._apply_filters())
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(fill="x")
        ttk.Label(search_frame, text="Search across all fields", font=("Helvetica", 8), foreground="gray").pack(anchor="w")

        # ── Antenna Name (without radome) ──
        type_frame = ttk.LabelFrame(fp, text="Antenna Name", padding=8)
        type_frame.pack(fill="x", pady=5, padx=5)

        self.type_var = tk.StringVar(value="All")
        self.type_combo = ttk.Combobox(type_frame, textvariable=self.type_var, state="readonly", width=28)
        self.type_combo.pack(fill="x")
        self.type_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # ── Radome ──
        radome_frame = ttk.LabelFrame(fp, text="Radome Code", padding=8)
        radome_frame.pack(fill="x", pady=5, padx=5)

        self.radome_var = tk.StringVar(value="All")
        self.radome_combo = ttk.Combobox(radome_frame, textvariable=self.radome_var, state="readonly", width=28)
        self.radome_combo.pack(fill="x")
        self.radome_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # ── Method ──
        method_frame = ttk.LabelFrame(fp, text="Calibration Method", padding=8)
        method_frame.pack(fill="x", pady=5, padx=5)

        self.method_var = tk.StringVar(value="All")
        self.method_combo = ttk.Combobox(method_frame, textvariable=self.method_var, state="readonly", width=28)
        self.method_combo.pack(fill="x")
        self.method_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # ── Institution ──
        inst_frame = ttk.LabelFrame(fp, text="Institution", padding=8)
        inst_frame.pack(fill="x", pady=5, padx=5)

        self.inst_var = tk.StringVar(value="All")
        self.inst_combo = ttk.Combobox(inst_frame, textvariable=self.inst_var, state="readonly", width=28)
        self.inst_combo.pack(fill="x")
        self.inst_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        # ── ANTEX Version Filter (multi-select checkboxes) ──
        version_frame = ttk.LabelFrame(fp, text="ANTEX Version", padding=8)
        version_frame.pack(fill="x", pady=5, padx=5)

        self.version_checkbox_vars = {}  # {version: BooleanVar}
        self.version_checkbox_frame = ttk.Frame(version_frame)
        self.version_checkbox_frame.pack(fill="x")
        ttk.Label(version_frame, text="none ticked = all versions",
                  foreground="gray", font=("Helvetica", 8)).pack(anchor="w", pady=(4, 0))

        # ── Frequency Filter (multi-select checkboxes) ──
        freq_frame = ttk.LabelFrame(fp, text="Contained Frequencies", padding=8)
        freq_frame.pack(fill="x", pady=5, padx=5)

        self.freq_checkbox_vars = {}  # {freq_code: BooleanVar}
        self.freq_checkbox_frame = ttk.Frame(freq_frame)
        self.freq_checkbox_frame.pack(fill="x")
        ttk.Label(freq_frame, text="Tick one or more to filter", font=("Helvetica", 8), foreground="gray").pack(anchor="w")

        # ── Statistics ──
        stats_frame = ttk.LabelFrame(fp, text="Database Statistics", padding=8)
        stats_frame.pack(fill="x", pady=5, padx=5)

        self.stats_text = tk.StringVar(value="No database loaded.")
        ttk.Label(stats_frame, textvariable=self.stats_text, wraplength=240, justify="left", font=("Helvetica", 9)).pack(fill="x")

    def _build_table(self, parent):
        # Table with scrollbars
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(container, orient="vertical")
        vsb.pack(side="right", fill="y")

        # Horizontal scrollbar
        hsb = ttk.Scrollbar(container, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        col_ids = [c[0] for c in COLUMNS]
        self.tree = ttk.Treeview(
            container,
            columns=col_ids,
            show="headings",
            selectmode="browse",
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
        )

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Configure columns
        for col_id, display_name, width in COLUMNS:
            self.tree.heading(col_id, text=display_name, command=lambda c=col_id: self._sort_by_column(c))
            self.tree.column(col_id, width=width, minwidth=40)

        self.tree.pack(fill="both", expand=True)

        # Bind selection
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

        # Tag for alternating rows
        self.tree.tag_configure("oddrow", background="#2a2a2a")
        self.tree.tag_configure("evenrow", background="#1e1e1e")

    def _build_detail_panel(self, parent):
        self.detail_text = tk.Text(parent, height=6, wrap="word", bg="#1a1a2e", fg="#e0e0e0",
                                    font=("Consolas", 10), relief="flat", padx=10, pady=8)
        self.detail_text.pack(fill="both", expand=True)
        self.detail_text.insert("1.0", "Select a record from the table above to view details.")
        self.detail_text.config(state="disabled")

        # Button bar for detail actions
        detail_btn_frame = ttk.Frame(parent)
        detail_btn_frame.pack(fill="x", pady=(5, 0))

        self.show_explorer_btn = ttk.Button(detail_btn_frame, text="Show in Explorer",
            command=self._show_in_explorer, bootstyle="info-outline", state="disabled")
        self.show_explorer_btn.pack(side="left", padx=3)

        self.copy_path_btn = ttk.Button(detail_btn_frame, text="Copy Path",
            command=self._copy_path_to_clipboard, bootstyle="secondary-outline", state="disabled")
        self.copy_path_btn.pack(side="left", padx=3)

    def _build_status_bar(self):
        status = ttk.Frame(self, padding=(10, 3))
        status.pack(fill="x", side="bottom")

        self.status_var = tk.StringVar(value="Ready. Load a database or scan an ATX folder to begin.")
        ttk.Label(status, textvariable=self.status_var, font=("Helvetica", 9)).pack(side="left")

        self.count_var = tk.StringVar(value="")
        ttk.Label(status, textvariable=self.count_var, font=("Helvetica", 9, "bold"), bootstyle="info").pack(side="right")

    # ── Data Loading ──────────────────────────────────────────────────────────

    def _try_load_default_db(self):
        """Try to load the built-in database on startup."""
        default = resource_path("igs20_database.db")
        if os.path.exists(default):
            self._load_db_file(default)

    def _load_database(self):
        """Open file dialog to load an external SQLite database."""
        path = filedialog.askopenfilename(
            title="Select ANTEX Database",
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")],
            initialdir=os.path.dirname(os.path.abspath(__file__)),
        )
        if path:
            self._load_db_file(path)

    def _scan_folder(self):
        """Scan an external folder (including all subfolders) of ATX files with cancel support."""
        folder = filedialog.askdirectory(title="Select folder containing .atx files (all subfolders will be scanned)")
        if not folder:
            return

        # Import the parser
        try:
            from parse_antex import parse_antex_file
        except ImportError:
            messagebox.showerror("Import Error",
                "Could not import parse_antex module.\n"
                "Ensure parse_antex.py is in the same directory as atx_scanner.py.")
            return

        # Progress dialog with cancel button
        self._scan_cancel = False
        progress_win = tk.Toplevel(self)
        progress_win.title("Scanning...")
        progress_win.geometry("460x200")
        progress_win.resizable(False, False)
        progress_win.transient(self)
        progress_win.grab_set()

        pframe = ttk.Frame(progress_win, padding=20)
        pframe.pack(fill="both", expand=True)

        scan_status = tk.StringVar(value=f"Searching for .atx files in:\n{folder}")
        ttk.Label(pframe, textvariable=scan_status, wraplength=400, justify="left").pack(fill="x")

        progress_bar = ttk.Progressbar(pframe, mode="determinate")
        progress_bar.pack(fill="x", pady=(10, 5))

        def cancel_scan():
            self._scan_cancel = True
            cancel_btn.config(state="disabled", text="Cancelling...")

        cancel_btn = ttk.Button(pframe, text="Cancel", command=cancel_scan, bootstyle="danger")
        cancel_btn.pack()

        self.update_idletasks()

        # Find files
        scan_status.set(f"Searching for .atx files in:\n{folder}")
        self.update_idletasks()

        atx_files = sorted(Path(folder).rglob("*.atx"))
        if not atx_files:
            progress_win.destroy()
            messagebox.showinfo("No Files", f"No .atx files found in:\n{folder}")
            return

        total = len(atx_files)
        progress_bar["maximum"] = total
        all_records = []

        for i, fpath in enumerate(atx_files):
            if self._scan_cancel:
                progress_win.destroy()
                self.status_var.set("Scan cancelled by user.")
                return

            scan_status.set(f"Processing file {i + 1}/{total}:\n{fpath.name}")
            progress_bar["value"] = i + 1
            self.update_idletasks()

            records = parse_antex_file(fpath)
            for r in records:
                if "frequencies_str" not in r:
                    r["frequencies_str"] = ", ".join(r.get("frequencies", []))
                if "num_frequencies_detected" not in r:
                    r["num_frequencies_detected"] = str(len(r.get("frequencies", [])))
            all_records.extend(records)

        progress_win.destroy()

        if not all_records:
            messagebox.showinfo("No Data", "No antenna records found in the scanned files.")
            return

        # Convert to list-of-dicts with DB column names
        self.all_records = []
        for r in all_records:
            row = {
                # This remap lists the fields to keep, so anything omitted is
                # silently dropped - which is exactly how the ANTEX version went
                # missing: the parser filled it in and the table had a column for
                # it, but the value never survived this step.
                "antex_version": r.get("antex_version", ""),
                "antenna_type": r.get("antenna_type", ""),
                "serial_number": r.get("serial_number", ""),
                "radome": r.get("radome", ""),
                "method": r.get("method", ""),
                "institution": r.get("institution", ""),
                "num_calibrations": r.get("num_calibrations", ""),
                "calibration_date": r.get("calibration_date", ""),
                "dazi": r.get("dazi", ""),
                "zen1": r.get("zen1", ""),
                "zen2": r.get("zen2", ""),
                "dzen": r.get("dzen", ""),
                "num_frequencies": r.get("num_frequencies_detected", ""),
                "frequencies": r.get("frequencies_str", ""),
                "file_name": r.get("file_name", ""),
                "relative_path": r.get("relative_path", ""),
            }
            self.all_records.append(row)

        self.db_path = f"[Scanned: {folder}]"
        self._populate_filters()
        self._apply_filters()
        self._update_stats()
        self._update_db_indicator()
        self.status_var.set(f"Scanned {len(atx_files)} files, loaded {len(self.all_records)} records from: {folder}")

    def _load_db_file(self, path):
        """Load records from a SQLite database file."""
        try:
            conn = sqlite3.connect(path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()

            # Select only the columns this database actually has. A database
            # exported before a column was introduced (e.g. antex_version) would
            # otherwise fail outright with "no such column"; the missing fields
            # are filled with blanks instead.
            existing = {row[1] for row in cursor.execute("PRAGMA table_info(antennas)")}
            available = [c for c in DB_COLUMNS if c in existing]
            missing = [c for c in DB_COLUMNS if c not in existing]
            cursor.execute(f"SELECT {', '.join(available)} FROM antennas "
                           "ORDER BY antenna_type, serial_number")
            rows = cursor.fetchall()
            conn.close()

            self.all_records = [dict(row, **{c: "" for c in missing}) for row in rows]
            if missing:
                print(f"  Note: this database predates {', '.join(missing)} - "
                      "re-scan the folder to populate those columns.")
            self.db_path = path

            self._populate_filters()
            self._apply_filters()
            self._update_stats()
            self._update_db_indicator()

            self.status_var.set(f"Loaded {len(self.all_records)} records from: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to load database:\n{e}")

    # ── Filter Logic ──────────────────────────────────────────────────────────

    def _get_antenna_name(self, antenna_type):
        """Extract antenna name without radome code."""
        # antenna_type is typically: 'LEIAR25 LEIT' or 'AOAD/M_T_RFI_GL NONE'
        # Split on last space to separate name from radome
        parts = antenna_type.rsplit(" ", 1)
        return parts[0] if parts else antenna_type

    def _populate_filters(self):
        """Populate filter dropdowns from current data."""
        def unique_sorted(key):
            vals = sorted(set(r.get(key, "") for r in self.all_records))
            return ["All"] + [v for v in vals if v]

        # Antenna names without radome code
        antenna_names = sorted(set(
            self._get_antenna_name(r.get("antenna_type", ""))
            for r in self.all_records if r.get("antenna_type")
        ))
        self.type_combo["values"] = ["All"] + antenna_names
        self.radome_combo["values"] = unique_sorted("radome")
        self.method_combo["values"] = unique_sorted("method")
        self.inst_combo["values"] = unique_sorted("institution")

        # Frequencies: extract individual freq codes and build checkboxes
        # Filter out unofficial "C"-variant frequencies (EC1, EC5, GC1, GC5)
        EXCLUDED_FREQS = {"EC1", "EC5", "GC1", "GC5"}
        all_freqs = set()
        for r in self.all_records:
            freqs = r.get("frequencies", "")
            if freqs:
                for f in freqs.split(", "):
                    f = f.strip()
                    if f and f not in EXCLUDED_FREQS:
                        all_freqs.add(f)

        # ANTEX version checkboxes - shown alongside the frequency
        # boxes, so a scan can be narrowed to v1.4 or v2.0 files.
        if hasattr(self, 'version_checkbox_frame'):
            for widget in self.version_checkbox_frame.winfo_children():
                widget.destroy()
            self.version_checkbox_vars.clear()
            versions = sorted({r.get("antex_version", "") for r in self.all_records
                               if r.get("antex_version")})
            for i, ver in enumerate(versions):
                var = tk.BooleanVar(value=False)
                var.trace_add("write", lambda *a: self._apply_filters())
                self.version_checkbox_vars[ver] = var
                ttk.Checkbutton(self.version_checkbox_frame, text=f"v{ver}",
                                variable=var).grid(row=0, column=i, sticky="w",
                                                   padx=4, pady=1)

        # Clear old checkboxes
        for widget in self.freq_checkbox_frame.winfo_children():
            widget.destroy()
        self.freq_checkbox_vars.clear()

        # Create checkboxes in a 4-column grid
        for i, freq in enumerate(sorted(all_freqs)):
            var = tk.BooleanVar(value=False)
            var.trace_add("write", lambda *a: self._apply_filters())
            self.freq_checkbox_vars[freq] = var
            cb = ttk.Checkbutton(self.freq_checkbox_frame, text=freq, variable=var)
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=2, pady=1)

        # Reset selections
        self.type_var.set("All")
        self.radome_var.set("All")
        self.method_var.set("All")
        self.inst_var.set("All")
        self.search_var.set("")

    def _apply_filters(self, *args):
        """Apply all active filters and update the table."""
        search = self.search_var.get().lower().strip()
        type_f = self.type_var.get()
        radome_f = self.radome_var.get()
        method_f = self.method_var.get()
        inst_f = self.inst_var.get()

        # Get selected frequency checkboxes
        selected_freqs = [freq for freq, var in self.freq_checkbox_vars.items() if var.get()]

        # Selected ANTEX versions. None ticked = show every version.
        selected_versions = [v for v, var in getattr(self, 'version_checkbox_vars', {}).items()
                             if var.get()]

        filtered = []
        for r in self.all_records:
            if selected_versions and r.get("antex_version", "") not in selected_versions:
                continue
            # Antenna name filter (without radome)
            if type_f != "All":
                ant_name = self._get_antenna_name(r.get("antenna_type", ""))
                if ant_name != type_f:
                    continue
            if radome_f != "All" and r.get("radome", "") != radome_f:
                continue
            if method_f != "All" and r.get("method", "") != method_f:
                continue
            if inst_f != "All" and r.get("institution", "") != inst_f:
                continue

            # Frequency checkbox filter: record must contain ALL selected frequencies
            if selected_freqs:
                record_freqs = r.get("frequencies", "")
                if not all(sf in record_freqs for sf in selected_freqs):
                    continue

            # Text search
            if search:
                haystack = " ".join(str(v) for v in r.values()).lower()
                if search not in haystack:
                    continue

            filtered.append(r)

        self.filtered_records = filtered
        self._refresh_table()
        self.count_var.set(f"Showing {len(filtered)} of {len(self.all_records)} records")

    def _reset_filters(self):
        """Reset all filters to default."""
        self.type_var.set("All")
        self.radome_var.set("All")
        self.method_var.set("All")
        self.inst_var.set("All")
        # Uncheck all frequency checkboxes
        for var in self.freq_checkbox_vars.values():
            var.set(False)
        # Uncheck all ANTEX version checkboxes
        for var in getattr(self, 'version_checkbox_vars', {}).values():
            var.set(False)
        self.search_var.set("")
        self._apply_filters()

    # ── Table Operations ──────────────────────────────────────────────────────

    def _refresh_table(self):
        """Clear and repopulate the treeview with filtered records."""
        self.tree.delete(*self.tree.get_children())

        for i, record in enumerate(self.filtered_records):
            values = [record.get(c, "") for c in DB_COLUMNS]
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert("", "end", iid=str(i), values=values, tags=(tag,))

    def _sort_by_column(self, col):
        """Sort the table by clicking a column header."""
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False

        self.filtered_records.sort(
            key=lambda r: str(r.get(col, "")).lower(),
            reverse=self.sort_reverse
        )

        # Update heading to show sort indicator
        for c_id, c_name, _ in COLUMNS:
            indicator = ""
            if c_id == col:
                indicator = " v" if self.sort_reverse else " ^"
            self.tree.heading(c_id, text=f"{DISPLAY_NAMES[c_id]}{indicator}")

        self._refresh_table()

    def _on_row_selected(self, event):
        """Show details for the selected record."""
        sel = self.tree.selection()
        if not sel:
            return

        idx = int(sel[0])
        if idx >= len(self.filtered_records):
            return

        record = self.filtered_records[idx]

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")

        lines = []
        lines.append(f"{'Antenna Type:':<22s} {record.get('antenna_type', '')}")
        lines.append(f"{'Serial Number:':<22s} {record.get('serial_number', '')}")
        lines.append(f"{'Radome Code:':<22s} {record.get('radome', '')}")
        lines.append(f"{'Cal. Method:':<22s} {record.get('method', '')}")
        lines.append(f"{'Institution:':<22s} {record.get('institution', '')}")
        lines.append(f"{'# Calibrations:':<22s} {record.get('num_calibrations', '')}")
        lines.append(f"{'Cal. Date:':<22s} {record.get('calibration_date', '')}")
        lines.append(f"{'DAZI:':<22s} {record.get('dazi', '')} deg")
        lines.append(f"{'ZEN1 / ZEN2 / DZEN:':<22s} {record.get('zen1', '')} / {record.get('zen2', '')} / {record.get('dzen', '')} deg")
        lines.append(f"{'# Frequencies:':<22s} {record.get('num_frequencies', '')}")
        lines.append(f"{'Frequencies:':<22s} {record.get('frequencies', '')}")
        lines.append(f"{'File Name:':<22s} {record.get('file_name', '')}")
        lines.append(f"{'Relative Path:':<22s} {record.get('relative_path', '')}")

        self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.config(state="disabled")

        # Enable detail buttons
        self._selected_record = record
        self.show_explorer_btn.config(state="normal")
        self.copy_path_btn.config(state="normal")

    def _show_in_explorer(self):
        """Open the folder containing the selected record's ATX file in Windows Explorer."""
        record = getattr(self, '_selected_record', None)
        if not record:
            return
        file_path = record.get('relative_path', '') or record.get('file_name', '')
        if not file_path or not os.path.exists(file_path):
            # Try to find via db_path context
            messagebox.showinfo("Path Not Available",
                f"The file path could not be resolved on this system.\n\n"
                f"Stored path: {file_path}\n\n"
                "Use 'Copy Path' to copy the stored path to your clipboard.")
            return
        # Open explorer with file selected
        subprocess.Popen(f'explorer /select,"{os.path.abspath(file_path)}"')

    def _copy_path_to_clipboard(self):
        """Copy the selected record's file path to the clipboard."""
        record = getattr(self, '_selected_record', None)
        if not record:
            return
        file_path = record.get('relative_path', '') or record.get('file_name', '')
        self.clipboard_clear()
        self.clipboard_append(file_path)
        self.status_var.set(f"Path copied to clipboard: {file_path}")

    # ── Export ────────────────────────────────────────────────────────────────

    def _export_database(self):
        """Export current data as a SQLite database to a user-chosen path."""
        if not self.all_records:
            messagebox.showinfo("No Data", "No records to export.")
            return

        path = filedialog.asksaveasfilename(
            title="Export Database",
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")],
            initialfile=f"atx_export_{datetime.now():%Y%m%d_%H%M%S}.db",
        )
        if not path:
            return

        try:
            conn = sqlite3.connect(path)
            cursor = conn.cursor()
            cursor.execute('''CREATE TABLE IF NOT EXISTS antennas (
                antenna_type TEXT, serial_number TEXT, radome TEXT,
                method TEXT, institution TEXT, num_calibrations TEXT,
                calibration_date TEXT, dazi TEXT, zen1 TEXT, zen2 TEXT, dzen TEXT,
                num_frequencies TEXT, frequencies TEXT, file_name TEXT, relative_path TEXT
            )''')
            cursor.execute('DELETE FROM antennas')
            for r in self.filtered_records:
                cursor.execute(
                    'INSERT INTO antennas VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    tuple(r.get(c, "") for c in DB_COLUMNS)
                )
            conn.commit()
            conn.close()
            self.status_var.set(f"Exported {len(self.filtered_records)} records to: {os.path.basename(path)}")
            messagebox.showinfo("Export Complete",
                f"Exported {len(self.filtered_records)} records to:\n{path}\n\n"
                "This .db file can be opened with:\n"
                "  - DB Browser for SQLite (free)\n"
                "  - DBeaver (free)\n"
                "  - Python (sqlite3 module)\n"
                "  - Any SQLite-compatible tool")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export database:\n{e}")

    def _export_excel(self):
        """Export filtered records to Excel."""
        if not self.filtered_records:
            messagebox.showinfo("No Data", "No records to export.")
            return

        path = filedialog.asksaveasfilename(
            title="Export to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"atx_export_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        )
        if not path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "ATX Export"

            header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin_border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
            alt_fill = PatternFill(start_color="E8EDF5", end_color="E8EDF5", fill_type="solid")

            # Header
            for ci, (col_id, display_name, _) in enumerate(COLUMNS, 1):
                cell = ws.cell(row=1, column=ci, value=display_name)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border

            # Data
            for ri, record in enumerate(self.filtered_records, 2):
                for ci, (col_id, _, _) in enumerate(COLUMNS, 1):
                    cell = ws.cell(row=ri, column=ci, value=record.get(col_id, ""))
                    cell.border = thin_border
                    if ri % 2 == 0:
                        cell.fill = alt_fill

            # Auto-width
            for ci, (col_id, display_name, _) in enumerate(COLUMNS, 1):
                max_len = len(display_name)
                for r in self.filtered_records[:100]:
                    max_len = max(max_len, len(str(r.get(col_id, ""))))
                ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 50)

            ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(self.filtered_records) + 1}"
            ws.freeze_panes = "A2"

            wb.save(path)
            self.status_var.set(f"Exported {len(self.filtered_records)} records to: {os.path.basename(path)}")
            messagebox.showinfo("Export Complete", f"Exported {len(self.filtered_records)} records to:\n{path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{e}")

    def _export_csv(self):
        """Export filtered records to CSV."""
        if not self.filtered_records:
            messagebox.showinfo("No Data", "No records to export.")
            return

        path = filedialog.asksaveasfilename(
            title="Export to CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=f"atx_export_{datetime.now():%Y%m%d_%H%M%S}.csv",
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=DB_COLUMNS, extrasaction="ignore")
                writer.writerow({c: DISPLAY_NAMES[c] for c in DB_COLUMNS})
                writer.writerows(self.filtered_records)

            self.status_var.set(f"Exported {len(self.filtered_records)} records to: {os.path.basename(path)}")
            messagebox.showinfo("Export Complete", f"Exported {len(self.filtered_records)} records to:\n{path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{e}")

    # ── Statistics ────────────────────────────────────────────────────────────

    def _update_stats(self):
        """Update the statistics panel."""
        n = len(self.all_records)
        types = len(set(r.get("antenna_type", "") for r in self.all_records))
        serials = len(set(r.get("serial_number", "") for r in self.all_records if r.get("serial_number")))
        methods = set(r.get("method", "") for r in self.all_records if r.get("method"))
        insts = len(set(r.get("institution", "") for r in self.all_records if r.get("institution")))

        text = (
            f"Total records:      {n}\n"
            f"Antenna types:      {types}\n"
            f"Serial numbers:     {serials}\n"
            f"Methods:            {', '.join(sorted(methods))}\n"
            f"Institutions:       {insts}"
        )
        self.stats_text.set(text)

    def _update_db_indicator(self):
        """Update the database source indicator in the UI."""
        if self.db_path:
            self.db_indicator_var.set(str(self.db_path))
        else:
            self.db_indicator_var.set("No database loaded")

    # ── Dialogs ───────────────────────────────────────────────────────────────

    def _show_contact(self):
        import webbrowser
        dialog = tk.Toplevel(self)
        dialog.title("Contact Information")
        dialog.geometry("460x300")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Institut f\u00fcr Erdmessung (IfE)", font=("Helvetica", 12, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Leibniz Universität Hannover").pack(anchor="w")
        ttk.Label(frame, text="Schneiderberg 50").pack(anchor="w")
        ttk.Label(frame, text="D-30167 Hannover").pack(anchor="w", pady=(0, 15))

        ttk.Label(frame, text="Contact: Dr.-Ing. Johannes Kr\u00f6ger", font=("Helvetica", 10, "bold")).pack(anchor="w")

        email_frame = ttk.Frame(frame)
        email_frame.pack(anchor="w", pady=2)
        ttk.Label(email_frame, text="Email: ").pack(side="left")
        email_link = ttk.Label(email_frame, text="kroeger@ife.uni-hannover.de", foreground="#4da6ff", cursor="hand2")
        email_link.pack(side="left")
        email_link.bind("<Button-1>", lambda e: webbrowser.open("mailto:kroeger@ife.uni-hannover.de"))

        web_frame = ttk.Frame(frame)
        web_frame.pack(anchor="w", pady=2)
        ttk.Label(web_frame, text="Web: ").pack(side="left")
        web_link = ttk.Label(web_frame, text="www.ife.uni-hannover.de", foreground="#4da6ff", cursor="hand2")
        web_link.pack(side="left")
        web_link.bind("<Button-1>", lambda e: webbrowser.open("https://www.ife.uni-hannover.de"))

        def _open_mailing_list():
            # Close first. Two stacked modal dialogs each take the grab, and the
            # inner one releasing it would leave this window unresponsive.
            dialog.destroy()
            self._show_mailing_list()

        buttons = ttk.Frame(frame)
        buttons.pack(pady=(20, 0))
        ttk.Button(buttons, text="Mailing list", command=_open_mailing_list,
                   bootstyle="outline-info").pack(side="left", padx=4)
        ttk.Button(buttons, text="Close", command=dialog.destroy,
                   bootstyle="secondary").pack(side="left", padx=4)

        # Size to the content, with 460x300 as the minimum: a fixed size would cut
        # the buttons off on systems with larger fonts or display scaling.
        dialog.update_idletasks()
        dialog.geometry("%dx%d" % (max(460, dialog.winfo_reqwidth()),
                                   max(300, dialog.winfo_reqheight())))

    def _show_about(self):
        messagebox.showinfo("About ATX-Scanner",
            "ATX-Scanner\n\n"
            "ANTEX File Summary Browser\n"
            + VERSION_TEXT + "\n\n"
            "Desktop application for browsing, filtering,\n"
            "searching, and exporting GNSS antenna calibration\n"
            "metadata from ANTEX files.\n\n"
            "Features:\n"
            "  - Recursive scanning of all subfolders\n"
            "  - Export to Excel, CSV, and SQLite database\n"
            "  - Filter by antenna, radome, method, institution, frequency\n\n"
            "Developed at:\n"
            "Institut f\u00fcr Erdmessung (IfE)\n"
            "Leibniz Universität Hannover\n\n"
            "License: GNU General Public License v3 (GPLv3)\n\n"
            "Developers:\n"
            "  Amr Fawzy, M.Sc.\n"
            "  Dr.-Ing. Johannes Kr\u00f6ger"
        )

    def _show_mailing_list(self):
        """The mailing list dialog, on request, whatever was answered before."""
        if mailing_list is None:
            messagebox.showwarning(
                "IfE software mailing list",
                "This installation is missing mailing_list.py, so the "
                "subscription dialog cannot be shown. Please report it to "
                "the IfE, the Contact button has the address.")
            return
        mailing_list.show_on_request(self, "ATX-Scanner")

    def _on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to close ATX-Scanner?"):
            self.destroy()


# ── Entry Point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ATXExplorer()
    if mailing_list is not None:
        mailing_list.maybe_show(app, "ATX-Scanner")
    app.mainloop()
