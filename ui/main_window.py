"""Main GUI window for Deflection Check Tool."""
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict

from utils.sap_connector import SapConnector
from utils.deflection_calc import DeflectionCalculator
from utils.pmember_detector import PmemberDetector, FrameData
from data.models import BeamGroup, BeamCheckSummary
from output.exporters import ExcelExporter, TxtExporter


class MainWindow:
    EXPIRY_DATE = (2027, 4, 1)   # year, month, day

    # ── Color palette ─────────────────────────────────────────────────────────
    _NAVY   = "#1A3A6C"   # primary — header, headings, run button
    _BLUE   = "#2E6DB4"   # medium blue — connect button, hover
    _STEEL  = "#4A90D9"   # light blue — active/hover
    _TEAL   = "#00838F"   # teal — create actions
    _TEAL_H = "#006064"   # teal hover
    _GREEN  = "#2E7D32"   # green — export / success
    _GREEN_H= "#1B5E20"
    _RED    = "#C62828"   # red — danger / remove
    _RED_H  = "#8B0000"
    _ORANGE = "#D84315"   # orange — cantilever
    _ORANGE_H="#BF360C"
    _BG     = "#EEF2F7"   # main background (light blue-gray)
    _PANEL  = "#F5F7FA"   # section panel background
    _WHITE  = "#FFFFFF"
    _GRAY   = "#546E7A"   # text gray
    _LGRAY  = "#90A4AE"   # light gray
    _LOG_BG = "#1E2433"   # dark terminal background
    _LOG_FG = "#A8C7FA"   # light blue terminal text
    # ─────────────────────────────────────────────────────────────────────────

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Deflection Check Tool - SAP2000")
        self.root.geometry("1650x1050")
        self.root.minsize(1250, 750)
        self.root.configure(bg=self._BG)
        self._setup_style()
        self._check_expiry()

        self.sap = SapConnector()
        self.calculator = DeflectionCalculator()
        self.beam_groups: List[BeamGroup] = []
        self.results: List[BeamCheckSummary] = []
        self.all_groups: List[str] = []
        self.all_load_cases: List[str] = []
        self.all_combos: List[str] = []

        self._build_ui()

    def _check_expiry(self):
        from datetime import date
        expiry = date(*self.EXPIRY_DATE)
        if date.today() > expiry:
            messagebox.showwarning(
                "License Expired",
                f"This application license expired on {expiry.strftime('%d/%m/%Y')}.\n\n"
                "Please contact Roberto to continue using this application."
            )
            self.root.destroy()

    def _setup_style(self):
        s = ttk.Style()
        s.theme_use("clam")

        BG = self._BG

        # ── Base ──────────────────────────────────────────────────────────────
        s.configure(".",
                    background=BG,
                    font=("Segoe UI", 10),
                    troughcolor=BG)
        s.configure("TFrame",      background=BG)
        s.configure("TLabel",      background=BG, font=("Segoe UI", 10))
        s.configure("TEntry",      fieldbackground=self._WHITE, font=("Segoe UI", 10))
        s.configure("TRadiobutton",background=BG, font=("Segoe UI", 10))
        s.configure("TCheckbutton",background=BG, font=("Segoe UI", 10))
        s.configure("TScrollbar",  background="#B0BEC5", troughcolor="#DDE4EC",
                    arrowcolor=self._GRAY)
        s.configure("TProgressbar",background=self._BLUE, troughcolor="#D0D8E4")

        # ── LabelFrame ────────────────────────────────────────────────────────
        s.configure("TLabelframe",
                    background=self._PANEL,
                    bordercolor="#C5CDD6",
                    relief="solid",
                    borderwidth=1)
        s.configure("TLabelframe.Label",
                    font=("Segoe UI", 10, "bold"),
                    foreground=self._NAVY,
                    background=BG)

        # ── Buttons — default (neutral) ───────────────────────────────────────
        s.configure("TButton",
                    font=("Segoe UI", 10),
                    padding=(8, 4),
                    background="#DDE3EC",
                    foreground="#2C3E50",
                    bordercolor="#B0BEC5",
                    relief="flat")
        s.map("TButton",
              background=[("active", "#BCC9D8"), ("disabled", "#D5D5D5")],
              foreground=[("disabled", "#999999")],
              relief=[("active", "flat")])

        # ── RUN button — navy ─────────────────────────────────────────────────
        s.configure("Run.TButton",
                    font=("Segoe UI", 11, "bold"),
                    padding=(14, 6),
                    background=self._NAVY,
                    foreground="white",
                    relief="flat")
        s.map("Run.TButton",
              background=[("active", self._BLUE), ("disabled", "#888888")],
              foreground=[("disabled", "#cccccc")])

        # ── Create/Action — teal ──────────────────────────────────────────────
        s.configure("Create.TButton",
                    font=("Segoe UI", 10),
                    padding=(8, 4),
                    background=self._TEAL,
                    foreground="white",
                    relief="flat")
        s.map("Create.TButton",
              background=[("active", self._TEAL_H), ("disabled", "#888888")],
              foreground=[("disabled", "#cccccc")])

        # ── Export — green ────────────────────────────────────────────────────
        s.configure("Export.TButton",
                    font=("Segoe UI", 10),
                    padding=(8, 4),
                    background=self._GREEN,
                    foreground="white",
                    relief="flat")
        s.map("Export.TButton",
              background=[("active", self._GREEN_H), ("disabled", "#888888")],
              foreground=[("disabled", "#cccccc")])

        # ── Danger — red ──────────────────────────────────────────────────────
        s.configure("Danger.TButton",
                    font=("Segoe UI", 10),
                    padding=(8, 4),
                    background=self._RED,
                    foreground="white",
                    relief="flat")
        s.map("Danger.TButton",
              background=[("active", self._RED_H), ("disabled", "#888888")],
              foreground=[("disabled", "#cccccc")])

        # ── Cantilever — orange ───────────────────────────────────────────────
        s.configure("Warning.TButton",
                    font=("Segoe UI", 10),
                    padding=(8, 4),
                    background=self._ORANGE,
                    foreground="white",
                    relief="flat")
        s.map("Warning.TButton",
              background=[("active", self._ORANGE_H), ("disabled", "#888888")],
              foreground=[("disabled", "#cccccc")])

        # ── Notebook ─────────────────────────────────────────────────────────
        s.configure("TNotebook",
                    background=BG,
                    bordercolor="#C5CDD6")
        s.configure("TNotebook.Tab",
                    font=("Segoe UI", 10),
                    padding=(12, 5),
                    background="#D0D8E4",
                    foreground="#2C3E50")
        s.map("TNotebook.Tab",
              background=[("selected", self._NAVY)],
              foreground=[("selected", "white")])

        # ── Treeview ─────────────────────────────────────────────────────────
        s.configure("Treeview",
                    font=("Segoe UI", 10),
                    rowheight=24,
                    background=self._WHITE,
                    fieldbackground=self._WHITE,
                    bordercolor="#C5CDD6")
        s.configure("Treeview.Heading",
                    font=("Segoe UI", 10, "bold"),
                    background=self._NAVY,
                    foreground="white",
                    relief="flat",
                    padding=(4, 6))
        s.map("Treeview.Heading",
              background=[("active", self._BLUE)])
        s.map("Treeview",
              background=[("selected", "#BDD7EE")],
              foreground=[("selected", "#1A1A1A")])

    # =========================================================================
    # Build UI
    # =========================================================================
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=0)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Header bar ────────────────────────────────────────────────────────
        header = tk.Frame(main, bg=self._NAVY, padx=16, pady=10)
        header.pack(fill=tk.X)

        hdr_left = tk.Frame(header, bg=self._NAVY)
        hdr_left.pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(hdr_left, text="⬡  Deflection Check Tool",
                 font=("Segoe UI", 16, "bold"),
                 bg=self._NAVY, fg="#FFFFFF").pack(anchor=tk.W)
        tk.Label(hdr_left,
                 text=("Civil & Structural Engineering Analysis"
                       "  │  Roberto Inhouse  │  Version V2.2"),
                 font=("Segoe UI", 9),
                 bg=self._NAVY, fg="#A8C4E0").pack(anchor=tk.W)

        hdr_right = tk.Frame(header, bg=self._NAVY)
        hdr_right.pack(side=tk.RIGHT, fill=tk.Y, pady=2)

        tk.Button(hdr_right, text="🔗  Connect SAP2000",
                  font=("Segoe UI", 10, "bold"),
                  bg=self._BLUE, fg="white",
                  activebackground=self._STEEL, activeforeground="white",
                  relief=tk.FLAT, padx=12, pady=6,
                  cursor="hand2",
                  bd=0,
                  command=self._on_connect).pack(side=tk.LEFT, padx=(0, 14))

        status_pill = tk.Frame(hdr_right, bg="#2A2A3E",
                               padx=10, pady=5,
                               highlightbackground="#4A4A6A",
                               highlightthickness=1)
        status_pill.pack(side=tk.LEFT)
        self.lbl_status = tk.Label(status_pill, text="● Not connected",
                                   font=("Segoe UI", 10, "bold"),
                                   bg="#2A2A3E", fg="#FF6B6B")
        self.lbl_status.pack()
        self._status_pill = status_pill

        # ── Content area ─────────────────────────────────────────────────────
        content = ttk.Frame(main, padding=(10, 8, 10, 4))
        content.pack(fill=tk.BOTH, expand=True)

        paned = ttk.PanedWindow(content, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(paned, padding=5, width=420)
        paned.add(left, weight=2)
        right = ttk.Frame(paned, padding=5)
        paned.add(right, weight=5)

        self._build_group_section(left)
        self._build_loadcase_section(left)
        self._build_settings_section(left)
        self._build_action_buttons(left)
        self._build_results_section(right)

        # ── Log ───────────────────────────────────────────────────────────────
        log_f = ttk.LabelFrame(content, text="Log", padding=0)
        log_f.pack(fill=tk.X, pady=(6, 0))
        self.log_text = tk.Text(log_f, height=4, wrap=tk.WORD,
                                 font=("Consolas", 10),
                                 bg=self._LOG_BG, fg=self._LOG_FG,
                                 insertbackground="white",
                                 selectbackground="#3A4A6B",
                                 selectforeground="white",
                                 relief=tk.FLAT, padx=8, pady=6,
                                 borderwidth=0)
        self.log_text.pack(fill=tk.X)
        self.log_text.config(state=tk.DISABLED)

    # =========================================================================
    # Group Selection
    # =========================================================================
    def _build_group_section(self, parent):
        frame = ttk.LabelFrame(parent,
                               text="① Beam Groups (numeric names = beams)",
                               padding=6)
        frame.pack(fill=tk.X, pady=(0, 5))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="🔄 Load Groups",
                   command=self._load_groups).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Select All",
                   command=self._select_all_groups).pack(side=tk.LEFT, padx=2)
        self.btn_select_beams = ttk.Button(btn_row, text="Select Beams",
                                           command=self._select_beam_groups)
        self.btn_select_beams.pack(side=tk.LEFT, padx=2)
        self.btn_from_sap = ttk.Button(btn_row, text="📌 From SAP",
                                       command=self._on_select_from_sap)
        self.btn_from_sap.pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Clear",
                   command=self._clear_groups).pack(side=tk.LEFT, padx=2)

        btn_row2 = ttk.Frame(frame)
        btn_row2.pack(fill=tk.X, pady=(3, 0))
        self.btn_auto_groups = ttk.Button(btn_row2, text="⚙ Auto Create",
                                          style="Create.TButton",
                                          command=self._on_auto_groups)
        self.btn_auto_groups.pack(side=tk.LEFT, padx=2)
        self.btn_remove_groups = ttk.Button(btn_row2, text="🗑 Remove",
                                            style="Danger.TButton",
                                            command=self._on_remove_groups)
        self.btn_remove_groups.pack(side=tk.LEFT, padx=2)
        self.btn_manual_group = ttk.Button(btn_row2, text="✏ Manual Create",
                                           style="Create.TButton",
                                           command=self._on_manual_create_group)
        self.btn_manual_group.pack(side=tk.LEFT, padx=2)
        self.btn_cantilever_group = ttk.Button(btn_row2, text="🔰 Cantilever",
                                               style="Warning.TButton",
                                               command=self._on_cantilever_create_group)
        self.btn_cantilever_group.pack(side=tk.LEFT, padx=2)

        self.lbl_grp_count = ttk.Label(frame, text="Groups: 0",
                                        font=("Segoe UI", 9, "italic"),
                                        foreground=self._GRAY)
        self.lbl_grp_count.pack(anchor=tk.W, pady=(3, 0))

        self.beam_detect_frame = ttk.Frame(frame)
        self.beam_detect_frame.pack(fill=tk.X, pady=(2, 0))
        self.lbl_detect_progress = ttk.Label(self.beam_detect_frame, text="",
                                              font=("Segoe UI", 9),
                                              foreground=self._GRAY)
        self.lbl_detect_progress.pack(anchor=tk.W)
        self.beam_detect_progress = ttk.Progressbar(self.beam_detect_frame,
                                                     orient=tk.HORIZONTAL,
                                                     mode="determinate",
                                                     length=200)
        self.beam_detect_progress.pack(fill=tk.X)
        self.beam_detect_frame.pack_forget()

        self.auto_grp_progress_frame = ttk.Frame(frame)
        self.auto_grp_progress_frame.pack(fill=tk.X, pady=(2, 0))
        self.lbl_auto_grp_progress = ttk.Label(self.auto_grp_progress_frame,
                                                text="",
                                                font=("Segoe UI", 9),
                                                foreground=self._GRAY)
        self.lbl_auto_grp_progress.pack(anchor=tk.W)
        self.auto_grp_progress = ttk.Progressbar(self.auto_grp_progress_frame,
                                                  orient=tk.HORIZONTAL,
                                                  mode="determinate")
        self.auto_grp_progress.pack(fill=tk.X)
        self.auto_grp_progress_frame.pack_forget()

        list_f = ttk.Frame(frame)
        list_f.pack(fill=tk.X, pady=(4, 0))
        self.grp_listbox = tk.Listbox(list_f, height=9, selectmode=tk.EXTENDED,
                                       font=("Consolas", 10),
                                       exportselection=False,
                                       bg=self._WHITE,
                                       selectbackground=self._BLUE,
                                       selectforeground="white",
                                       activestyle="none",
                                       relief=tk.FLAT,
                                       highlightthickness=1,
                                       highlightbackground="#C5CDD6",
                                       highlightcolor=self._BLUE)
        sb = ttk.Scrollbar(list_f, orient=tk.VERTICAL,
                           command=self.grp_listbox.yview)
        self.grp_listbox.configure(yscrollcommand=sb.set)
        self.grp_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self._sap_highlight_after_id = None
        self.grp_listbox.bind("<<ListboxSelect>>", self._on_grp_listbox_select)

    # =========================================================================
    # Load Case / Combo
    # =========================================================================
    def _build_loadcase_section(self, parent):
        frame = ttk.LabelFrame(parent,
                               text="② Load Cases / Combinations",
                               padding=6)
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        ttk.Button(frame, text="🔄 Refresh",
                   command=self._refresh_lc).pack(anchor=tk.W)

        pf = ttk.Frame(frame)
        pf.pack(fill=tk.X, pady=4)
        ttk.Label(pf, text="Quick prefix:").pack(side=tk.LEFT)
        self.ent_prefix = ttk.Entry(pf, width=10)
        self.ent_prefix.pack(side=tk.LEFT, padx=3)
        ttk.Button(pf, text="Select",
                   command=self._select_by_prefix).pack(side=tk.LEFT)
        ttk.Button(pf, text="Clear",
                   command=self._clear_lc).pack(side=tk.LEFT, padx=2)

        nb = ttk.Notebook(frame)
        nb.pack(fill=tk.BOTH, expand=True, pady=(3, 0))

        def _styled_listbox(parent_frame):
            lb = tk.Listbox(parent_frame, selectmode=tk.EXTENDED,
                            font=("Consolas", 10),
                            exportselection=False,
                            bg=self._WHITE,
                            selectbackground=self._BLUE,
                            selectforeground="white",
                            activestyle="none",
                            relief=tk.FLAT,
                            highlightthickness=1,
                            highlightbackground="#C5CDD6",
                            highlightcolor=self._BLUE)
            sb = ttk.Scrollbar(parent_frame, orient=tk.VERTICAL,
                               command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            return lb

        lcb_f = ttk.Frame(nb)
        nb.add(lcb_f, text="Load Combinations")
        self.lcb_listbox = _styled_listbox(lcb_f)

        lc_f = ttk.Frame(nb)
        nb.add(lc_f, text="Load Cases")
        self.lc_listbox = _styled_listbox(lc_f)

        self.lc_notebook = nb

    # =========================================================================
    # Settings
    # =========================================================================
    def _build_settings_section(self, parent):
        frame = ttk.LabelFrame(parent, text="③ Settings", padding=6)
        frame.pack(fill=tk.X, pady=(0, 5))

        self.var_check_mode = tk.StringVar(value="")   # "rel" or "abs"

        r1 = ttk.Frame(frame)
        r1.pack(fill=tk.X)
        ttk.Radiobutton(r1, text="Relative check  L /",
                        variable=self.var_check_mode,
                        value="rel",
                        command=self._on_check_toggle).pack(side=tk.LEFT)
        self.var_ratio = tk.StringVar(value="360")
        self.ent_ratio = ttk.Entry(r1, textvariable=self.var_ratio, width=6)
        self.ent_ratio.pack(side=tk.LEFT, padx=3)

        r2 = ttk.Frame(frame)
        r2.pack(fill=tk.X, pady=(5, 0))
        ttk.Radiobutton(r2, text="Absolute check  Limit:",
                        variable=self.var_check_mode,
                        value="abs",
                        command=self._on_check_toggle).pack(side=tk.LEFT)
        self.var_abs_limit = tk.StringVar(value="25")
        self.ent_abs_limit = ttk.Entry(r2, textvariable=self.var_abs_limit,
                                        width=6)
        self.ent_abs_limit.pack(side=tk.LEFT, padx=3)
        ttk.Label(r2, text="mm").pack(side=tk.LEFT)

        self.ent_ratio.config(state=tk.DISABLED)
        self.ent_abs_limit.config(state=tk.DISABLED)

    def _on_check_toggle(self):
        mode = self.var_check_mode.get()
        self.ent_ratio.config(state=tk.NORMAL if mode == "rel" else tk.DISABLED)
        self.ent_abs_limit.config(state=tk.NORMAL if mode == "abs" else tk.DISABLED)

    # =========================================================================
    # Action Buttons
    # =========================================================================
    def _build_action_buttons(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=6)

        self.btn_run = ttk.Button(frame, text="▶  RUN CHECK",
                                   style="Run.TButton",
                                   command=self._on_run)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_export_xl = ttk.Button(frame, text="📊 Export Excel",
                                         style="Export.TButton",
                                         command=self._on_export_xl,
                                         state=tk.DISABLED)
        self.btn_export_xl.pack(side=tk.LEFT, padx=3)

        self.btn_export_txt = ttk.Button(frame, text="📄 Export TXT",
                                          style="Export.TButton",
                                          command=self._on_export_txt,
                                          state=tk.DISABLED)
        self.btn_export_txt.pack(side=tk.LEFT, padx=3)

        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=120)
        self.progress.pack(side=tk.RIGHT, padx=5)

    # =========================================================================
    # Results
    # =========================================================================
    def _build_results_section(self, parent):
        frame = ttk.LabelFrame(parent, text="④ Results", padding=6)
        frame.pack(fill=tk.BOTH, expand=True)

        top_bar = ttk.Frame(frame)
        top_bar.pack(fill=tk.X, pady=(0, 6))
        self.lbl_overall = ttk.Label(top_bar, text="No results yet",
                                      font=("Segoe UI", 10, "italic"),
                                      foreground=self._GRAY)
        self.lbl_overall.pack(side=tk.LEFT)
        self.var_ng_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(top_bar, text="Show NG only",
                        variable=self.var_ng_only,
                        command=self._display_results).pack(side=tk.RIGHT)

        cols = ("grp", "section", "lc", "span", "defl_ratio", "max_defl",
                "criteria", "result", "abs_max", "abs_result",
                "ctrl_node", "elements")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=20)

        hdg = {
            "grp":        ("Beam",        70),
            "section":    ("Section",    150),
            "lc":         ("Ctrl LC",     90),
            "span":       ("Span(mm)",    80),
            "defl_ratio": ("Deflection", 100),
            "max_defl":   ("Max(mm)",     75),
            "criteria":   ("Criteria",    70),
            "result":     ("Rel",         50),
            "abs_max":    ("Abs(mm)",     75),
            "abs_result": ("Abs",         50),
            "ctrl_node":  ("Ctrl Node",   90),
            "elements":   ("Elements",   150),
        }
        for cid, (txt, w) in hdg.items():
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor=tk.CENTER)

        self._all_tree_cols = cols

        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("ok", background="#C8EFC8", foreground="#1A4D1A")
        self.tree.tag_configure("ng", background="#FADADD", foreground="#7B0000")

    # =========================================================================
    # Handlers
    # =========================================================================
    def _on_connect(self):
        self._log("Connecting to SAP2000...")
        ok = self.sap.connect()
        if ok:
            m = os.path.basename(self.sap.model_path)
            self.lbl_status.config(text=f"● Connected: {m}",
                                   fg="#4CAF50", bg="#1E3A1E")
            self._status_pill.config(bg="#1E3A1E")
            self._log(f"Connected: {self.sap.model_path}")
            self._load_groups()
            self._refresh_lc()
        else:
            self.lbl_status.config(text="● Connection failed",
                                   fg="#FF6B6B", bg="#3A1E1E")
            self._status_pill.config(bg="#3A1E1E")
            messagebox.showerror("Error", "Cannot connect to SAP2000.")

    def _load_groups(self):
        if not self.sap.is_connected:
            return
        all_g = self.sap.get_group_names()
        beam_groups = []
        for g in all_g:
            stripped = g.strip()
            if stripped.isdigit() or stripped.replace('.', '').isdigit():
                beam_groups.append(stripped)
            else:
                frames = self.sap.get_group_frames(stripped)
                if frames and stripped not in ("ALL", "All"):
                    beam_groups.append(stripped)

        import re
        def _natural_key(s):
            return [int(p) if p.isdigit() else p.lower()
                    for p in re.split(r'(\d+)', s)]

        self.all_groups = sorted(beam_groups, key=_natural_key)
        self.grp_listbox.delete(0, tk.END)
        for g in self.all_groups:
            self.grp_listbox.insert(tk.END, g)
        self.lbl_grp_count.config(text=f"Groups: {len(self.all_groups)}")
        self._log(f"Found {len(self.all_groups)} beam groups")

    def _select_all_groups(self):
        self.grp_listbox.select_set(0, tk.END)

    def _select_beam_groups(self):
        """Select only groups whose frames are beams (near-horizontal)."""
        if not self.sap.is_connected or not self.all_groups:
            return
        self.grp_listbox.selection_clear(0, tk.END)
        self._log("Detecting beams...")

        total = len(self.all_groups)
        self.beam_detect_progress["maximum"] = total
        self.beam_detect_progress["value"] = 0
        self.lbl_detect_progress.config(text=f"Scanning 0 / {total} groups...")
        self.beam_detect_frame.pack(fill=tk.X, pady=(2, 0))
        self.btn_select_beams.config(state=tk.DISABLED)

        def _detect():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                beam_indices = []
                for i, gname in enumerate(self.all_groups):
                    name = gname.strip()
                    if not name.isdigit() and not name.startswith('G-') and \
                            not name.startswith('GC-'):
                        continue
                    frames = self.sap.get_group_frames(gname)
                    if frames and self.sap.is_frame_beam(frames[0]):
                        beam_indices.append(i)

                    done = i + 1
                    self.root.after(0, lambda d=done: (
                        self.beam_detect_progress.__setitem__("value", d),
                        self.lbl_detect_progress.config(
                            text=f"Scanning {d} / {total} groups"
                                 f"  ({int(d / total * 100)}%)")
                    ))

                def _apply():
                    for idx in beam_indices:
                        self.grp_listbox.selection_set(idx)
                    self.beam_detect_frame.pack_forget()
                    self.btn_select_beams.config(state=tk.NORMAL)
                    self._log(f"Selected {len(beam_indices)} beam groups"
                              f" (of {total} total)")

                self.root.after(0, _apply)
            except Exception as e:
                self.root.after(0, lambda: self._log(f"Beam detect error: {e}"))
                self.root.after(0, lambda: self.beam_detect_frame.pack_forget())
                self.root.after(0, lambda: self.btn_select_beams.config(
                    state=tk.NORMAL))
            finally:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_detect, daemon=True).start()

    def _clear_groups(self):
        self.grp_listbox.selection_clear(0, tk.END)

    def _on_grp_listbox_select(self, event=None):
        if not self.sap.is_connected:
            return
        if self._sap_highlight_after_id is not None:
            self.root.after_cancel(self._sap_highlight_after_id)
        self._sap_highlight_after_id = self.root.after(
            250, self._sap_highlight_selected_groups)

    def _sap_highlight_selected_groups(self):
        self._sap_highlight_after_id = None
        sel_indices = self.grp_listbox.curselection()
        if not sel_indices:
            return
        group_names = [self.grp_listbox.get(i) for i in sel_indices]

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                m = self.sap.sap_model
                m.SelectObj.ClearSelection()
                for gname in group_names:
                    m.SelectObj.Group(gname, False)
                m.View.RefreshView(0, False)
            except Exception as e:
                self.root.after(0, lambda: self._log(f"SAP highlight error: {e}"))
            finally:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    def _on_select_from_sap(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return
        if not self.all_groups:
            messagebox.showwarning("Warning", "Load groups first!")
            return

        self.btn_from_sap.config(state=tk.DISABLED)

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                sel_frames = self.sap.get_selected_frames()
                if not sel_frames:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "From SAP Selection",
                        "No frame elements selected in SAP2000.\n"
                        "Please select frame(s) in SAP2000 first."))
                    return

                sel_set = set(sel_frames)
                frame_to_groups: dict = {}
                for gname in self.all_groups:
                    frames = self.sap.get_group_frames(gname)
                    for fname in frames:
                        if fname not in frame_to_groups:
                            frame_to_groups[fname] = []
                        frame_to_groups[fname].append(gname)

                matched_groups = set()
                for fname in sel_set:
                    for gname in frame_to_groups.get(fname, []):
                        if gname.startswith('G-'):
                            matched_groups.add(gname)

                def _apply():
                    self.grp_listbox.selection_clear(0, tk.END)
                    count = 0
                    for i, gname in enumerate(self.all_groups):
                        if gname in matched_groups:
                            self.grp_listbox.selection_set(i)
                            count += 1
                    if count:
                        for i, gname in enumerate(self.all_groups):
                            if gname in matched_groups:
                                self.grp_listbox.see(i)
                                break
                    self._log(f"From SAP: {len(sel_frames)} frames selected → "
                              f"{count} group(s) highlighted")
                    if count == 0:
                        messagebox.showinfo(
                            "From SAP Selection",
                            f"{len(sel_frames)} frame(s) selected in SAP2000,\n"
                            "but none belong to any group in the current list.\n\n"
                            "Try loading groups first.")

                self.root.after(0, _apply)

            except Exception as e:
                self.root.after(0, lambda: self._log(f"From SAP error: {e}"))
            finally:
                self.root.after(0, lambda: self.btn_from_sap.config(
                    state=tk.NORMAL))
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    def _refresh_lc(self):
        if not self.sap.is_connected:
            return
        import re
        def _nk(s):
            return [int(p) if p.isdigit() else p.lower()
                    for p in re.split(r'(\d+)', s)]

        self.all_load_cases = sorted(self.sap.get_load_case_names(), key=_nk)
        self.all_combos     = sorted(self.sap.get_combo_names(),     key=_nk)
        self.lc_listbox.delete(0, tk.END)
        for lc in self.all_load_cases:
            self.lc_listbox.insert(tk.END, lc)
        self.lcb_listbox.delete(0, tk.END)
        for lcb in self.all_combos:
            self.lcb_listbox.insert(tk.END, lcb)
        self._log(f"Load Cases: {len(self.all_load_cases)},"
                  f" Combos: {len(self.all_combos)}")

    @staticmethod
    def _extract_numeric_part(name: str) -> str:
        upper = name.upper()
        for pfx in ("COMBO", "COMB", "LC", "CB"):
            if upper.startswith(pfx):
                return name[len(pfx):]
        return name

    def _select_by_prefix(self):
        prefix = self.ent_prefix.get().strip()
        if not prefix:
            return
        active = self.lc_notebook.index(self.lc_notebook.select())
        lb = self.lcb_listbox if active == 0 else self.lc_listbox
        lb.selection_clear(0, tk.END)
        cnt = 0
        for i in range(lb.size()):
            item = lb.get(i)
            if self._extract_numeric_part(item).startswith(prefix):
                lb.selection_set(i)
                cnt += 1
        self._log(f"Selected {cnt} items with prefix '{prefix}'")

    def _clear_lc(self):
        self.lc_listbox.selection_clear(0, tk.END)
        self.lcb_listbox.selection_clear(0, tk.END)

    def _on_run(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return

        sel_groups = [self.grp_listbox.get(i)
                      for i in self.grp_listbox.curselection()]
        if not sel_groups:
            messagebox.showwarning("Warning", "Select at least one beam group!")
            return

        sel_lc  = [self.lc_listbox.get(i)
                   for i in self.lc_listbox.curselection()]
        sel_lcb = [self.lcb_listbox.get(i)
                   for i in self.lcb_listbox.curselection()]
        if not sel_lc and not sel_lcb:
            messagebox.showwarning("Warning",
                                   "Select at least one Load Case/Combo!")
            return

        mode    = self.var_check_mode.get()
        use_rel = mode == "rel"
        use_abs = mode == "abs"
        if not use_rel and not use_abs:
            messagebox.showwarning(
                "Warning",
                "Select a check type in Settings (Relative or Absolute).")
            return

        allow_ratio = 360.0
        if use_rel:
            try:
                allow_ratio = float(self.var_ratio.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid deflection ratio value.")
                return

        abs_limit = 0.0
        if use_abs:
            try:
                abs_limit = float(self.var_abs_limit.get())
                if abs_limit <= 0:
                    messagebox.showerror("Error",
                                         "Absolute limit must be > 0.")
                    return
            except ValueError:
                messagebox.showerror("Error", "Invalid absolute limit value.")
                return

        self._use_rel = use_rel
        self._use_abs = use_abs
        self.calculator = DeflectionCalculator(allowable_ratio=allow_ratio,
                                               abs_limit_mm=abs_limit)
        self.btn_run.config(state=tk.DISABLED)
        self.progress.start(10)
        self._log(f"Running... Groups={len(sel_groups)},"
                  f" LCs={len(sel_lc)}, LCBs={len(sel_lcb)}")

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass

            try:
                beams = []
                all_node_names = set()
                all_elm_joint_map = {}

                for gname in sel_groups:
                    frames = self.sap.get_group_frames(gname)
                    if not frames:
                        self._log(f"  Group '{gname}': no frames, skipped")
                        continue

                    ordered_nodes = []
                    frame_elements = []
                    for fname in frames:
                        fe = self.sap.get_frame_info(fname)
                        if fe:
                            frame_elements.append(fe)
                        mesh_pts, elm_map = \
                            self.sap.get_frame_obj_mesh_points(fname)
                        all_elm_joint_map.update(elm_map)
                        if mesh_pts:
                            if not ordered_nodes:
                                ordered_nodes.extend(mesh_pts)
                            else:
                                if mesh_pts[0] == ordered_nodes[-1]:
                                    ordered_nodes.extend(mesh_pts[1:])
                                else:
                                    ordered_nodes.extend(mesh_pts)

                    if len(ordered_nodes) < 2 and frame_elements:
                        for fe in frame_elements:
                            if not ordered_nodes:
                                ordered_nodes.append(fe.joint_i)
                            ordered_nodes.append(fe.joint_j)

                    all_node_names.update(ordered_nodes)
                    coords = self.sap.get_node_coordinates(ordered_nodes)

                    if gname.startswith('GC-'):
                        core = gname[3:]
                        if core.endswith('_S') or core.endswith('_E'):
                            core = core[:-2]
                        display_name = core
                    elif gname.startswith('G-'):
                        display_name = gname[2:]
                    else:
                        display_name = gname

                    bg = self.calculator.build_beam_group(
                        group_name=display_name,
                        frame_elements=frame_elements,
                        all_nodes_ordered=ordered_nodes,
                        nodes_coords=coords
                    )

                    if gname.startswith('GC-'):
                        bg.is_cantilever = True
                        if gname.endswith('_S'):
                            bg.free_end_node = bg.start_node
                        elif gname.endswith('_E'):
                            bg.free_end_node = bg.end_node

                    cantilever_tag = " [CANTILEVER]" if bg.is_cantilever else ""
                    beams.append(bg)
                    self._log(f"  Group '{gname}': {len(frame_elements)} elements,"
                              f" {len(ordered_nodes)} nodes,"
                              f" L={bg.total_length:.0f}mm{cantilever_tag}")

                if not beams:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Warning", "No valid beams found."))
                    return

                all_coords = self.sap.get_node_coordinates(
                    list(all_node_names))
                self._log(f"Total nodes: {len(all_coords)}")

                disps = self.sap.get_joint_displacements(
                    node_names=list(all_node_names),
                    load_cases=sel_lc  if sel_lc  else None,
                    load_combos=sel_lcb if sel_lcb else None,
                    elm_joint_map=all_elm_joint_map
                                  if all_elm_joint_map else None,
                )
                self._log(f"Displacement records: {len(disps)}")

                if len(disps) == 0:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "No Results",
                        "0 displacement records. Check analysis and LC/LCB."))

                self.results = self.calculator.run_full_check(
                    beams, all_coords, disps)
                self.root.after(0, self._display_results)

            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass
                self.root.after(0, self._run_done)

        threading.Thread(target=_run, daemon=True).start()

    def _run_done(self):
        self.btn_run.config(state=tk.NORMAL)
        self.progress.stop()
        has = bool(self.results)
        self.btn_export_xl.config(state=tk.NORMAL if has else tk.DISABLED)
        self.btn_export_txt.config(state=tk.NORMAL if has else tk.DISABLED)

    def _display_results(self):
        self.tree.delete(*self.tree.get_children())
        if not self.results:
            return

        use_rel = getattr(self, '_use_rel', False)
        use_abs = getattr(self, '_use_abs', False)
        ng_only = self.var_ng_only.get()

        rel_cols  = ("grp", "section", "lc", "span",
                     "defl_ratio", "max_defl", "criteria", "result")
        abs_cols  = ("abs_max", "abs_result")
        tail_cols = ("ctrl_node", "elements")
        if use_rel and use_abs:
            vis = rel_cols + abs_cols + tail_cols
        elif use_rel:
            vis = rel_cols + tail_cols
        else:
            vis = ("grp", "section", "lc", "span") + abs_cols + tail_cols
        self.tree["displaycolumns"] = vis

        def _is_ok(r):
            if use_rel and use_abs:
                return r.rel_is_ok and r.abs_is_ok
            if use_rel:
                return r.rel_is_ok
            return r.abs_is_ok

        total = len(self.results)
        fail  = sum(1 for r in self.results if not _is_ok(r))

        if fail == 0:
            self.lbl_overall.config(
                text=f"✓  ALL PASS — {total} beams checked",
                foreground=self._GREEN,
                font=("Segoe UI", 10, "bold"))
        else:
            suffix = "  [showing NG only]" if ng_only else ""
            self.lbl_overall.config(
                text=f"✗  {fail}/{total} FAILED{suffix}",
                foreground=self._RED,
                font=("Segoe UI", 10, "bold"))

        def _sort_key(x):
            n = x.group_name
            return (0, int(n)) if n.isdigit() else (1, n)

        for r in sorted(self.results, key=_sort_key):
            ok = _is_ok(r)
            if ng_only and ok:
                continue
            tag = "ok" if ok else "ng"
            el  = ",".join(r.element_list)
            self.tree.insert("", tk.END, values=(
                r.group_name, r.section, r.controlling_lc,
                f"{r.span_mm:.0f}", r.ratio_str,
                f"{r.max_deflection_mm:.3f}",
                f"L/{int(r.allowable_ratio)}",
                "OK" if r.rel_is_ok else "NG",
                f"{r.max_abs_deflection_mm:.3f}",
                "OK" if r.abs_is_ok else "NG",
                r.critical_node,
                f"({el})"
            ), tags=(tag,))

        shown = fail if ng_only else total
        self._log(f"Done: {total} beams, {total - fail} OK, {fail} NG"
                  + (f"  — showing {shown} NG" if ng_only else ""))

    def _on_export_xl(self):
        if not self.results:
            return
        d = self._output_dir()
        try:
            use_rel   = getattr(self, '_use_rel', False)
            use_abs   = getattr(self, '_use_abs', False)
            ratio     = float(self.var_ratio.get()) if use_rel else 360.0
            abs_limit = float(self.var_abs_limit.get()) if use_abs else 0.0
            proj = (os.path.basename(self.sap.model_path)
                    if self.sap.model_path else "")
            p = ExcelExporter().export(self.results, d, ratio, proj, abs_limit,
                                       use_rel=use_rel, use_abs=use_abs)
            self._log(f"Excel: {p}")
            messagebox.showinfo("Done", f"Saved:\n{p}")
            os.startfile(p)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _on_export_txt(self):
        if not self.results:
            return
        d = self._output_dir()
        try:
            use_rel   = getattr(self, '_use_rel', False)
            use_abs   = getattr(self, '_use_abs', False)
            ratio     = float(self.var_ratio.get()) if use_rel else 360.0
            abs_limit = float(self.var_abs_limit.get()) if use_abs else 0.0
            p = TxtExporter().export(self.results, d, ratio, abs_limit,
                                     use_rel=use_rel, use_abs=use_abs)
            self._log(f"TXT: {p}")
            messagebox.showinfo("Done", f"Saved:\n{p}")
            os.startfile(os.path.dirname(p))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _output_dir(self) -> str:
        if self.sap.model_directory:
            d = os.path.join(self.sap.model_directory, "Deflection Check")
        else:
            d = os.path.join(os.path.expanduser("~"), "Deflection Check")
        os.makedirs(d, exist_ok=True)
        return d

    # =========================================================================
    # Manual Create Group
    # =========================================================================
    def _on_manual_create_group(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return

        self.btn_manual_group.config(state=tk.DISABLED)

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                sel_frames = self.sap.get_selected_frames()

                if not sel_frames:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Manual Create Group",
                        "No frame elements selected in SAP2000.\n"
                        "Please select frame(s) in SAP2000 first, "
                        "then click this button."))
                    return

                if len(sel_frames) > 1:
                    frame_nodes = {}
                    for fname in sel_frames:
                        try:
                            pts = self.sap.sap_model.FrameObj.GetPoints(fname)
                            frame_nodes[fname] = (str(pts[0]), str(pts[1]))
                        except Exception as e:
                            self.root.after(0, lambda err=e: messagebox.showerror(
                                "Manual Create Group",
                                f"Cannot read node info: {err}"))
                            return

                    if not self.sap._is_chain(sel_frames, frame_nodes):
                        names = ", ".join(sorted(sel_frames))
                        self.root.after(0, lambda: messagebox.showerror(
                            "Manual Create Group",
                            f"Selected frames are NOT contiguous:\n{names}\n\n"
                            "They must form a single connected chain."))
                        return

                ok, msg = self.sap.manual_create_group(sel_frames)

                def _finish():
                    if ok:
                        self._log(f"Manual Group: {msg}")
                    else:
                        messagebox.showerror("Manual Create Group", msg)
                        self._log(f"Manual Group FAILED: {msg}")
                    self._load_groups()

                self.root.after(0, _finish)

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Manual Create Group", f"Unexpected error:\n{e}"))
                self.root.after(0, lambda: self._log(
                    f"Manual Group error: {e}"))
            finally:
                self.root.after(0, lambda: self.btn_manual_group.config(
                    state=tk.NORMAL))
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    # =========================================================================
    # Cantilever Create Group
    # =========================================================================
    def _on_cantilever_create_group(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return

        self.btn_cantilever_group.config(state=tk.DISABLED)

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                sel_frames, sel_points = \
                    self.sap.get_selected_frames_and_points()

                if not sel_frames:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Cantilever Create",
                        "No frame elements selected in SAP2000.\n\n"
                        "Please select the cantilever beam frame(s) AND\n"
                        "the free end node in SAP2000, then click this button."))
                    return

                if not sel_points:
                    names = ", ".join(sorted(sel_frames))
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Cantilever Create",
                        f"Beam(s) selected: {names}\n\n"
                        "⚠ Free End Node not selected!\n\n"
                        "Please also select the FREE END NODE (point object)\n"
                        "of the cantilever in SAP2000."))
                    return

                if len(sel_points) > 1:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Cantilever Create",
                        f"{len(sel_points)} nodes selected:"
                        f" {', '.join(sel_points)}\n\n"
                        "Please select exactly 1 Free End Node."))
                    return

                free_node = sel_points[0]

                frame_nodes = {}
                if len(sel_frames) > 1:
                    for fname in sel_frames:
                        try:
                            pts = self.sap.sap_model.FrameObj.GetPoints(fname)
                            frame_nodes[fname] = (str(pts[0]), str(pts[1]))
                        except Exception as e:
                            self.root.after(0, lambda err=e: messagebox.showerror(
                                "Cantilever Create",
                                f"Cannot read node info: {err}"))
                            return

                    if not self.sap._is_chain(sel_frames, frame_nodes):
                        names = ", ".join(sorted(sel_frames))
                        self.root.after(0, lambda: messagebox.showerror(
                            "Cantilever Create",
                            f"Selected frames are NOT contiguous:\n{names}\n\n"
                            "They must form a single connected chain."))
                        return
                else:
                    try:
                        pts = self.sap.sap_model.FrameObj.GetPoints(
                            sel_frames[0])
                        frame_nodes[sel_frames[0]] = (str(pts[0]), str(pts[1]))
                    except Exception as e:
                        self.root.after(0, lambda err=e: messagebox.showerror(
                            "Cantilever Create",
                            f"Cannot read node info: {err}"))
                        return

                ok, msg = self.sap.cantilever_create_group(
                    sel_frames, free_node)

                def _finish():
                    if ok:
                        self._log(f"Cantilever Group: {msg}")
                    else:
                        messagebox.showerror("Cantilever Create", msg)
                        self._log(f"Cantilever Group FAILED: {msg}")
                    self._load_groups()

                self.root.after(0, _finish)

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Cantilever Create", f"Unexpected error:\n{e}"))
                self.root.after(0, lambda: self._log(
                    f"Cantilever Group error: {e}"))
            finally:
                self.root.after(0, lambda: self.btn_cantilever_group.config(
                    state=tk.NORMAL))
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    # =========================================================================
    # Remove Selected Groups
    # =========================================================================
    def _on_remove_groups(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return
        sel_indices = self.grp_listbox.curselection()
        if not sel_indices:
            messagebox.showwarning("Warning",
                                   "Select at least one group to remove!")
            return
        sel_groups = [self.grp_listbox.get(i) for i in sel_indices]
        if not messagebox.askyesno(
            "Confirm Remove",
            f"Delete {len(sel_groups)} group(s) from SAP2000?\n\n"
            + "\n".join(sel_groups[:10])
            + ("\n..." if len(sel_groups) > 10 else "")
        ):
            return

        self.btn_remove_groups.config(state=tk.DISABLED)

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                n_deleted = 0
                errors = []
                for gname in sel_groups:
                    try:
                        ret = self.sap.sap_model.GroupDef.Delete(gname)
                        if ret == 0:
                            n_deleted += 1
                        else:
                            errors.append(f"'{gname}' ret={ret}")
                    except Exception as e:
                        errors.append(f"'{gname}': {e}")

                def _finish():
                    self._log(f"Remove Groups: {n_deleted}/"
                              f"{len(sel_groups)} deleted.")
                    if errors:
                        self._log(
                            f"  Errors: {'; '.join(errors[:5])}"
                            + (" ..." if len(errors) > 5 else ""))
                    self._load_groups()
                    self.btn_remove_groups.config(state=tk.NORMAL)

                self.root.after(0, _finish)
            except Exception as e:
                self.root.after(0, lambda: self._log(
                    f"Remove Groups error: {e}"))
                self.root.after(0, lambda: self.btn_remove_groups.config(
                    state=tk.NORMAL))
            finally:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    # =========================================================================
    # Auto Create Groups
    # =========================================================================
    def _on_auto_groups(self):
        if not self.sap.is_connected:
            messagebox.showwarning("Warning", "Connect to SAP2000 first!")
            return

        mode = self._ask_mesh_mode()
        if mode is None:
            return

        self.btn_auto_groups.config(state=tk.DISABLED)
        self._log(f"Auto Groups [{mode}]: reading model data...")

        def _set_progress(text, value, maximum=None):
            def _ui():
                self.auto_grp_progress_frame.pack(fill=tk.X, pady=(2, 0))
                self.lbl_auto_grp_progress.config(text=text)
                if maximum is not None:
                    self.auto_grp_progress.config(maximum=maximum)
                self.auto_grp_progress["value"] = value
            self.root.after(0, _ui)

        def _run():
            try:
                import comtypes
                comtypes.CoInitialize()
            except Exception:
                pass
            try:
                _set_progress("Reading frames from SAP2000...", 0,
                              maximum=100)
                raw_frames = self.sap.get_all_frames_raw()
                n = len(raw_frames)
                _set_progress(f"{n} frames read. Detecting groups...", 30,
                              maximum=100)
                self.root.after(0, lambda: self._log(
                    f"Auto Groups: {n} frames read"))

                if mode == 'auto_mesh':
                    pmembers = [
                        {'group_name': 'G-' + r['name'],
                         'frames': [r['name']]}
                        for r in raw_frames
                    ]
                    self.root.after(0, lambda: self._log(
                        f"Auto Groups: {len(pmembers)} groups (1 per frame)"))
                else:
                    frames = [
                        FrameData(
                            name=r['name'],
                            node_i=r['node_i'], node_j=r['node_j'],
                            xi=r['xi'], yi=r['yi'], zi=r['zi'],
                            xj=r['xj'], yj=r['yj'], zj=r['zj'],
                            section=r.get('section', ''),
                        )
                        for r in raw_frames
                    ]
                    releases    = self.sap.get_all_releases()
                    post_frames = self.sap.get_post_frame_names()
                    detector    = PmemberDetector()
                    pmembers    = detector.detect(frames, releases, post_frames)
                    self.root.after(0, lambda: self._log(
                        f"Auto Groups: {len(pmembers)} PMember groups detected"))

                total_pm = len(pmembers)
                _set_progress(f"Writing 0 / {total_pm} groups to SAP2000...",
                              0, maximum=total_pm)

                n_created, n_deleted, errors = self.sap.create_pmember_groups(
                    pmembers,
                    progress_cb=lambda done: _set_progress(
                        f"Writing {done} / {total_pm} groups to SAP2000..."
                        f"  ({int(done / total_pm * 100)}%)",
                        done,
                    )
                )

                def _finish():
                    self.auto_grp_progress_frame.pack_forget()
                    self._log(f"Auto Groups done: "
                              f"{n_deleted} old G-... groups deleted, "
                              f"{n_created} new groups created.")
                    if errors:
                        self._log(f"  Warnings ({len(errors)}): "
                                  f"{'; '.join(errors[:5])}"
                                  + (" ..." if len(errors) > 5 else ""))
                    self._load_groups()
                    self.btn_auto_groups.config(state=tk.NORMAL)

                self.root.after(0, _finish)

            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: self.auto_grp_progress_frame.pack_forget())
                self.root.after(0, lambda: self._log(f"Auto Groups error: {e}"))
                self.root.after(0, lambda: self.btn_auto_groups.config(
                    state=tk.NORMAL))
            finally:
                try:
                    comtypes.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    def _ask_mesh_mode(self):
        result  = [None]
        BG      = "#F5F7FA"
        C1_BG   = "#E8F5E9"; C1_FG = "#1B5E20"; C1_HOV = "#C8E6C9"
        C2_BG   = "#E3F2FD"; C2_FG = "#0D47A1"; C2_HOV = "#BBDEFB"
        HDR_BG  = self._NAVY

        dlg = tk.Toplevel(self.root)
        dlg.title("Auto Create Groups")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.configure(bg=BG)

        # Header strip
        tk.Frame(dlg, bg=HDR_BG, height=6).pack(fill=tk.X)
        tk.Label(dlg, text="Select meshing mode",
                 font=("Segoe UI", 12, "bold"),
                 bg=BG, fg=self._NAVY).pack(pady=(16, 12))

        def _choose(mode):
            result[0] = mode
            dlg.destroy()

        card_frame = tk.Frame(dlg, bg=BG)
        card_frame.pack(padx=22, pady=(0, 14))

        def _make_card(parent, col, bg, hov, title_fg, title, body, mode):
            card = tk.Frame(parent, bg=bg, relief=tk.FLAT, bd=0,
                            highlightbackground="#CCCCCC",
                            highlightthickness=1,
                            cursor="hand2", width=175, height=115)
            card.grid(row=0, column=col, padx=8, sticky=tk.NSEW)
            card.grid_propagate(False)

            tk.Label(card, text=title,
                     font=("Segoe UI", 10, "bold"),
                     bg=bg, fg=title_fg,
                     wraplength=155, justify=tk.CENTER
                     ).place(relx=0.5, rely=0.32, anchor=tk.CENTER)
            tk.Label(card, text=body,
                     font=("Segoe UI", 8),
                     bg=bg, fg="#444444",
                     wraplength=155, justify=tk.CENTER
                     ).place(relx=0.5, rely=0.70, anchor=tk.CENTER)

            def _on_enter(e):
                card.configure(bg=hov)
                for w in card.winfo_children():
                    w.configure(bg=hov)
            def _on_leave(e):
                card.configure(bg=bg)
                for w in card.winfo_children():
                    w.configure(bg=bg)

            for w in (card, *card.winfo_children()):
                w.bind("<Button-1>", lambda e, m=mode: _choose(m))
                w.bind("<Enter>", _on_enter)
                w.bind("<Leave>", _on_leave)

        _make_card(card_frame, 0, C1_BG, C1_HOV, C1_FG,
                   "No Auto Meshing",
                   "Frames manually split.\n1 PMember = 1 Group",
                   'no_mesh')
        _make_card(card_frame, 1, C2_BG, C2_HOV, C2_FG,
                   "Auto Mesh Frame Object",
                   "SAP2000 Auto Mesh.\n1 Frame = 1 Group",
                   'auto_mesh')

        tk.Frame(dlg, bg="#DDDDDD", height=1).pack(fill=tk.X)
        btn_row = tk.Frame(dlg, bg=BG)
        btn_row.pack(fill=tk.X, padx=16, pady=10)
        ttk.Button(btn_row, text="Cancel",
                   command=lambda: _choose(None)).pack(side=tk.RIGHT)

        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width()  - 410) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 220) // 2
        dlg.geometry(f"410x220+{x}+{y}")

        dlg.wait_window()
        return result[0]

    def _log(self, msg: str):
        from datetime import datetime
        self.log_text.config(state=tk.NORMAL)
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{ts}]  {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
