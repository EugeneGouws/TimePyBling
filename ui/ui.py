"""
ui.py — TimePyBling main interface

Tabs
----
  Timetable   — browse Block → SubBlock → Class → Students, with live search
  Verification — clash report, teacher qualification check, cost breakdown
  Exams        — exam slot scheduling by grade with exclusion management
  Export       — write optimised ST1.xlsx

Top bar
-------
  [Load Timetable]  filename
  [Load Teachers]   filename
  [Load Students]   filename  (placeholder — not yet wired)

Usage
-----
    python ui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd

from core.timetable_tree  import build_timetable_tree_from_file
from reader.exam_tree       import build_exam_tree_from_timetable_tree
from reader.exam_clash      import build_clash_graph, dsatur_colouring, is_excluded
from reader.verify_timetable import _find_clashes
from optimiser.cost_function   import evaluate, load_teacher_prefs_from_xlsx, CostConfig
from core.timetable_converter import timetable_tree_to_block_tree
from core.timetable_converter  import timetable_tree_to_block_tree
from optimiser.optimiser        import SARunner, SAConfig

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────

DEFAULT_EXCLUSIONS = {"ST", "LIB", "PE", "RDI"}

# Columns in teachers.xlsx that list subject codes
TEACHER_SUBJECT_COLS = ["sua", "sub", "suc"]

# Colour scheme
CLR_HEADER   = "#2c3e50"
CLR_GREEN    = "#27ae60"
CLR_BLUE     = "#2980b9"
CLR_RED      = "#c0392b"
CLR_LIGHT    = "#ecf0f1"
CLR_MID      = "#bdc3c7"
CLR_WHITE    = "white"
CLR_BG       = "#f5f5f5"


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def _scrolled_text(parent, **kw) -> tk.Text:
    """Text widget + vertical scrollbar packed into parent."""
    frame = tk.Frame(parent, bg=CLR_WHITE)
    frame.pack(fill=tk.BOTH, expand=True)
    sb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    t = tk.Text(
        frame,
        font=("Courier", 9),
        relief=tk.FLAT,
        bg="#f8f8f8",
        state=tk.DISABLED,
        wrap=tk.NONE,
        yscrollcommand=sb.set,
        **kw
    )
    t.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.config(command=t.yview)
    return t


def _write(widget: tk.Text, text: str, tag: str = ""):
    """Append text to a disabled Text widget, optionally with a colour tag."""
    widget.config(state=tk.NORMAL)
    if tag:
        widget.insert(tk.END, text, tag)
    else:
        widget.insert(tk.END, text)
    widget.config(state=tk.DISABLED)


def _clear(widget: tk.Text):
    widget.config(state=tk.NORMAL)
    widget.delete("1.0", tk.END)
    widget.config(state=tk.DISABLED)


def _load_teacher_subject_map(path: Path) -> dict[str, set[str]]:
    """
    Parse teachers.xlsx and return {teacher_code: {subject_codes}}.
    Columns: Teacher Code, sua, sub, suc
    """
    df = pd.read_excel(path)
    result = {}
    for _, row in df.iterrows():
        code = str(row.get("Teacher Code", "")).strip()
        if not code:
            continue
        subjects = set()
        for col in TEACHER_SUBJECT_COLS:
            val = row.get(col)
            if val and not pd.isna(val):
                subjects.add(str(val).strip())
        result[code] = subjects
    return result


def _extract_teacher_subjects_from_tree(tree) -> dict[str, set[str]]:
    """
    Walk a TimetableTree and return {teacher_code: {subject_codes}}
    based on what each teacher is actually assigned to teach.
    """
    actual: dict[str, set[str]] = {}
    for block in tree.blocks.values():
        for subblock in block.subblocks.values():
            for label in subblock.class_lists:
                parts = label.split("_")
                subject = parts[0]
                teacher = "_".join(parts[1:-1])
                if teacher:
                    actual.setdefault(teacher, set()).add(subject)
    return actual


# ─────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────

class TimePyBlingApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("TimePyBling")
        self.geometry("1200x800")
        self.configure(bg=CLR_BG)

        # ── state ──
        self.timetable_tree    = None
        self.block_tree        = None
        self.exam_tree         = None
        self.st1_path          = None
        self.teachers_path     = None
        self.sa_runner         = None
        self.teacher_subj_map  = {}      # {teacher_code: {subject_codes}}
        self.exclusions        = set(DEFAULT_EXCLUSIONS)

        self._build_ui()

    # ─────────────────────────────────────────────────────────
    # BUILD UI
    # ─────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_topbar()
        self._build_notebook()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=CLR_HEADER, pady=8, padx=10)
        bar.pack(fill=tk.X)

        tk.Label(
            bar, text="TimePyBling",
            font=("Helvetica", 14, "bold"),
            bg=CLR_HEADER, fg=CLR_WHITE
        ).pack(side=tk.LEFT, padx=(0, 24))

        # — Load Timetable —
        tk.Button(
            bar, text="Load Timetable", command=self._load_st1,
            bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 9, "bold"), padx=10, pady=3
        ).pack(side=tk.LEFT)

        self.st1_label = tk.Label(
            bar, text="No timetable loaded",
            bg=CLR_HEADER, fg=CLR_MID, font=("Helvetica", 9)
        )
        self.st1_label.pack(side=tk.LEFT, padx=(6, 20))

        # — Load Teachers —
        tk.Button(
            bar, text="Load Teachers", command=self._load_teachers,
            bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 9, "bold"), padx=10, pady=3
        ).pack(side=tk.LEFT)

        self.teachers_label = tk.Label(
            bar, text="No teachers loaded",
            bg=CLR_HEADER, fg=CLR_MID, font=("Helvetica", 9)
        )
        self.teachers_label.pack(side=tk.LEFT, padx=(6, 20))

        # — Load Students (future) —
        tk.Button(
            bar, text="Load Students",
            state=tk.DISABLED,
            bg="#555", fg=CLR_MID, relief=tk.FLAT,
            font=("Helvetica", 9, "bold"), padx=10, pady=3
        ).pack(side=tk.LEFT)

        tk.Label(
            bar, text="Coming soon",
            bg=CLR_HEADER, fg="#666", font=("Helvetica", 9, "italic")
        ).pack(side=tk.LEFT, padx=(6, 0))

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_timetable    = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_verification = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_exams        = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_export       = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_optimiser = tk.Frame(self.notebook, bg=CLR_WHITE)

        self.notebook.add(self.tab_timetable,    text="  Timetable  ")
        self.notebook.add(self.tab_verification, text="  Verification  ")
        self.notebook.add(self.tab_exams,        text="  Exams  ")
        self.notebook.add(self.tab_export,       text="  Export  ")
        self.notebook.add(self.tab_optimiser, text="  Optimiser  ")

        self._build_timetable_tab()
        self._build_verification_tab()
        self._build_exam_tab()
        self._build_export_tab()
        self._build_optimiser_tab()

    # ─────────────────────────────────────────────────────────
    # TAB 1 — TIMETABLE
    # ─────────────────────────────────────────────────────────

    def _build_timetable_tab(self):
        search_bar = tk.Frame(self.tab_timetable, bg=CLR_WHITE, pady=6, padx=8)
        search_bar.pack(fill=tk.X)

        tk.Label(search_bar, text="Search:", bg=CLR_WHITE,
                 font=("Helvetica", 10)).pack(side=tk.LEFT)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)

        tk.Entry(
            search_bar, textvariable=self.search_var,
            font=("Helvetica", 10), width=30, relief=tk.SOLID, bd=1
        ).pack(side=tk.LEFT, padx=6)

        tk.Label(
            search_bar,
            text="student ID · subject code · teacher name",
            bg=CLR_WHITE, fg="#888", font=("Helvetica", 9)
        ).pack(side=tk.LEFT)

        tk.Button(
            search_bar, text="Clear",
            command=lambda: self.search_var.set(""),
            relief=tk.FLAT, bg=CLR_LIGHT, font=("Helvetica", 9), padx=8
        ).pack(side=tk.LEFT, padx=4)

        tree_frame = tk.Frame(self.tab_timetable, bg=CLR_WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.tt_tree = ttk.Treeview(tree_frame, show="tree")
        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                           command=self.tt_tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tt_tree.configure(yscrollcommand=sb.set)
        self.tt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._style_tree(self.tt_tree)

    # ─────────────────────────────────────────────────────────
    # TAB 2 — OPTIMISER
    # ─────────────────────────────────────────────────────────
    def _build_optimiser_tab(self):
        content = tk.Frame(self.tab_optimiser, bg=CLR_WHITE, padx=12, pady=12)
        content.pack(fill=tk.BOTH, expand=True)

        # ── Parameter row ──
        params = tk.Frame(content, bg=CLR_WHITE)
        params.pack(fill=tk.X, pady=(0, 10))

        def param(label, default, width=8):
            tk.Label(params, text=label, bg=CLR_WHITE,
                     font=("Helvetica", 9)).pack(side=tk.LEFT, padx=(0, 2))
            v = tk.StringVar(value=default)
            tk.Entry(params, textvariable=v, width=width,
                     font=("Helvetica", 9), relief=tk.SOLID, bd=1).pack(
                side=tk.LEFT, padx=(0, 12))
            return v

        self.sa_t_start = param("T start", "1000.0")
        self.sa_t_min = param("T min", "0.1")
        self.sa_cooling = param("Cooling rate", "0.9999", width=10)
        self.sa_max_iter = param("Max iter", "500000", width=10)

        # ── Button row ──
        btn_row = tk.Frame(content, bg=CLR_WHITE)
        btn_row.pack(fill=tk.X, pady=(0, 10))

        self.sa_run_btn = tk.Button(
            btn_row, text="Run SA",
            command=self._sa_run,
            bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 10, "bold"), padx=14, pady=6
        )
        self.sa_run_btn.pack(side=tk.LEFT)

        self.sa_stop_btn = tk.Button(
            btn_row, text="Stop",
            command=self._sa_stop,
            bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 10, "bold"), padx=14, pady=6,
            state=tk.DISABLED
        )
        self.sa_stop_btn.pack(side=tk.LEFT, padx=8)

        self.sa_status = tk.Label(
            btn_row, text="Ready",
            bg=CLR_WHITE, fg="#888", font=("Helvetica", 9)
        )
        self.sa_status.pack(side=tk.LEFT, padx=8)

        # ── Log ──
        tk.Label(content, text="Progress log",
                 bg=CLR_WHITE, font=("Helvetica", 10, "bold")).pack(anchor=tk.W)

        log_frame = tk.Frame(content, bg=CLR_WHITE)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.sa_log = _scrolled_text(log_frame)
        self.sa_log.tag_config("good", foreground=CLR_GREEN)
        self.sa_log.tag_config("bad", foreground=CLR_RED)
        self.sa_log.tag_config("info", foreground=CLR_BLUE)
        self.sa_log.tag_config("dim", foreground="#888")

    # ─────────────────────────────────────────────────────────
    # TAB 3 — VERIFICATION
    # ─────────────────────────────────────────────────────────

    def _build_verification_tab(self):
        pane = tk.PanedWindow(
            self.tab_verification, orient=tk.HORIZONTAL,
            bg="#ccc", sashwidth=5
        )
        pane.pack(fill=tk.BOTH, expand=True)

        # ── Left: clash report ──
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=520)

        clash_header = tk.Frame(left, bg=CLR_WHITE)
        clash_header.pack(fill=tk.X, padx=8, pady=(8, 2))

        tk.Label(
            clash_header, text="Clash Report",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold")
        ).pack(side=tk.LEFT)

        tk.Button(
            clash_header, text="Re-run",
            command=self._run_verification,
            bg=CLR_LIGHT, font=("Helvetica", 8), relief=tk.FLAT, padx=8
        ).pack(side=tk.RIGHT)

        clash_frame = tk.Frame(left, bg=CLR_WHITE)
        clash_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.clash_report = _scrolled_text(clash_frame)
        # Colour tags
        self.clash_report.tag_config("pass",    foreground=CLR_GREEN)
        self.clash_report.tag_config("fail",    foreground=CLR_RED)
        self.clash_report.tag_config("heading", foreground="#333",
                                     font=("Courier", 9, "bold"))
        self.clash_report.tag_config("warn",    foreground="#e67e22")
        self.clash_report.tag_config("dim",     foreground="#888")

        # ── Right: cost + teacher checks ──
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=300)

        # Cost breakdown
        cost_lf = tk.LabelFrame(
            right, text="Cost Function  E(T)",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
            padx=6, pady=6
        )
        cost_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))
        self.cost_text = _scrolled_text(cost_lf)
        self.cost_text.tag_config("good", foreground=CLR_GREEN)
        self.cost_text.tag_config("bad",  foreground=CLR_RED)

        # Teacher qualification
        qual_lf = tk.LabelFrame(
            right, text="Teacher Qualifications",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
            padx=6, pady=6
        )
        qual_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))
        self.qual_text = _scrolled_text(qual_lf)
        self.qual_text.tag_config("ok",   foreground=CLR_GREEN)
        self.qual_text.tag_config("warn", foreground=CLR_RED)
        self.qual_text.tag_config("dim",  foreground="#888")

    # ─────────────────────────────────────────────────────────
    # TAB 4 — EXAMS
    # ─────────────────────────────────────────────────────────

    def _build_exam_tab(self):
        pane = tk.PanedWindow(
            self.tab_exams, orient=tk.HORIZONTAL,
            bg="#ddd", sashwidth=5
        )
        pane.pack(fill=tk.BOTH, expand=True)

        # Left: exam tree
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=500)

        tk.Label(
            left, text="Exam Tree", bg=CLR_WHITE,
            font=("Helvetica", 10, "bold"), pady=6
        ).pack()

        tree_frame = tk.Frame(left, bg=CLR_WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        self.ex_tree = ttk.Treeview(tree_frame, show="tree")
        ex_sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                               command=self.ex_tree.yview)
        ex_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.ex_tree.configure(yscrollcommand=ex_sb.set)
        self.ex_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._style_tree(self.ex_tree)

        # Right: exclusions + slot summary
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=260)

        excl_frame = tk.LabelFrame(
            right, text="Exam Exclusions",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
            padx=8, pady=8
        )
        excl_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(
            excl_frame,
            text="Subject codes excluded from exam scheduling:",
            bg=CLR_WHITE, fg="#555", font=("Helvetica", 8),
            wraplength=220, justify=tk.LEFT
        ).pack(anchor=tk.W)

        self.excl_listbox = tk.Listbox(
            excl_frame, height=6,
            font=("Courier", 10), relief=tk.SOLID, bd=1,
            selectmode=tk.SINGLE
        )
        self.excl_listbox.pack(fill=tk.X, pady=6)
        self._refresh_exclusion_listbox()

        add_row = tk.Frame(excl_frame, bg=CLR_WHITE)
        add_row.pack(fill=tk.X)

        self.excl_entry = tk.Entry(
            add_row, font=("Helvetica", 10),
            relief=tk.SOLID, bd=1, width=10
        )
        self.excl_entry.pack(side=tk.LEFT)
        self.excl_entry.bind("<Return>", lambda e: self._add_exclusion())

        tk.Button(
            add_row, text="Add", command=self._add_exclusion,
            bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 9, "bold"), padx=8
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            excl_frame, text="Remove Selected",
            command=self._remove_exclusion,
            bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 9), padx=8, pady=2
        ).pack(anchor=tk.W, pady=(4, 0))

        tk.Button(
            right, text="Rebuild Exam Tree",
            command=self._rebuild_exam,
            bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 10, "bold"), padx=12, pady=6
        ).pack(padx=10, pady=8, fill=tk.X)

        slot_lf = tk.LabelFrame(
            right, text="Slot Summary",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
            padx=8, pady=8
        )
        slot_lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.slot_text = _scrolled_text(slot_lf)

    # ─────────────────────────────────────────────────────────
    # TAB 5 — EXPORT
    # ─────────────────────────────────────────────────────────

    def _build_export_tab(self):
        content = tk.Frame(self.tab_export, bg=CLR_WHITE, padx=16, pady=16)
        content.pack(fill=tk.BOTH, expand=True)

        # ── Output path row ──
        path_frame = tk.Frame(content, bg=CLR_WHITE)
        path_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(
            path_frame, text="Output path:",
            bg=CLR_WHITE, font=("Helvetica", 10)
        ).pack(side=tk.LEFT)

        self.export_path_var = tk.StringVar(value="output/ST1_optimised.xlsx")
        tk.Entry(
            path_frame, textvariable=self.export_path_var,
            font=("Helvetica", 10), relief=tk.SOLID, bd=1, width=50
        ).pack(side=tk.LEFT, padx=8)

        tk.Button(
            path_frame, text="Browse…",
            command=self._browse_export_path,
            bg=CLR_LIGHT, font=("Helvetica", 9), relief=tk.FLAT, padx=8
        ).pack(side=tk.LEFT)

        # ── Export button ──
        tk.Button(
            content, text="Export ST1.xlsx",
            command=self._export,
            bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Helvetica", 11, "bold"), padx=16, pady=8
        ).pack(anchor=tk.W, pady=(0, 12))

        tk.Label(
            content, text="Export log",
            bg=CLR_WHITE, font=("Helvetica", 10, "bold")
        ).pack(anchor=tk.W)

        log_frame = tk.Frame(content, bg=CLR_WHITE)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.export_log = _scrolled_text(log_frame)
        self.export_log.tag_config("ok",   foreground=CLR_GREEN)
        self.export_log.tag_config("err",  foreground=CLR_RED)
        self.export_log.tag_config("info", foreground="#2980b9")

    # ─────────────────────────────────────────────────────────
    # LOAD — ST1
    # ─────────────────────────────────────────────────────────

    def _load_st1(self):
        path = filedialog.askopenfilename(
            title="Select student timetable (data/ST1.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")]
        )
        if not path:
            return

        self.st1_path = Path(path)
        self.st1_label.config(text=self.st1_path.name, fg=CLR_WHITE)

        try:
            self.timetable_tree = build_timetable_tree_from_file(self.st1_path)
            self.block_tree = timetable_tree_to_block_tree(self.timetable_tree)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return

        # Suggest export path alongside source file
        suggested = self.st1_path.parent / "ST1_optimised.xlsx"
        self.export_path_var.set(str(suggested))

        self._populate_timetable_tree()
        self._rebuild_exam()
        self._run_verification()
        self.notebook.select(self.tab_verification)

    # ─────────────────────────────────────────────────────────
    # LOAD — TEACHERS
    # ─────────────────────────────────────────────────────────

    def _load_teachers(self):
        path = filedialog.askopenfilename(
            title="Select teachers file (teachers.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")]
        )
        if not path:
            return

        self.teachers_path = Path(path)
        self.teachers_label.config(text=self.teachers_path.name, fg=CLR_WHITE)

        try:
            self.teacher_subj_map = _load_teacher_subject_map(self.teachers_path)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return

        # Re-run verification so teacher qualification panel updates
        if self.timetable_tree:
            self._run_verification()

    # ─────────────────────────────────────────────────────────
    # SA
    # ─────────────────────────────────────────────────────────

    def _sa_run(self):
        if not self.block_tree:
            messagebox.showwarning("No timetable", "Load a timetable first.")
            return
        if self.sa_runner and self.sa_runner.is_running():
            return

        _clear(self.sa_log)

        try:
            config = SAConfig(
                T_start=float(self.sa_t_start.get()),
                T_min=float(self.sa_t_min.get()),
                cooling_rate=float(self.sa_cooling.get()),
                max_iter=int(self.sa_max_iter.get()),
            )
        except ValueError as e:
            messagebox.showerror("Invalid parameter", str(e))
            return

        self.sa_run_btn.config(state=tk.DISABLED)
        self.sa_stop_btn.config(state=tk.NORMAL)
        self.sa_status.config(text="Running…", fg=CLR_GREEN)

        self.sa_runner = SARunner(
            bt=self.block_tree,
            config=config,
            progress_cb=self._sa_log_msg,
            done_cb=self._sa_done,
        )
        self.sa_runner.start()

    def _sa_stop(self):
        if self.sa_runner:
            self.sa_runner.stop()
        self.sa_status.config(text="Stopping…", fg="#e67e22")

    def _sa_log_msg(self, msg: str):
        # Called from background thread — must schedule on main thread
        self.after(0, lambda: _write(self.sa_log, msg + "\n"))

    def _sa_done(self, result):
        def _finish():
            self.sa_run_btn.config(state=tk.NORMAL)
            self.sa_stop_btn.config(state=tk.DISABLED)
            tag = "good" if result.improved else "bad"
            status = f"Done — cost {result.initial_cost} → {result.best_cost}"
            self.sa_status.config(
                text=status,
                fg=CLR_GREEN if result.improved else CLR_RED
            )
            _write(self.sa_log, "\n" + result.summary() + "\n", tag)
            # Refresh verification panel with new BlockTree state
            self._run_verification()

        self.after(0, _finish)

    # ─────────────────────────────────────────────────────────
    # VERIFICATION
    # ─────────────────────────────────────────────────────────

    def _run_verification(self):
        if not self.timetable_tree:
            return
        self._update_clash_report()
        self._update_cost_panel()
        self._update_qualification_panel()

    def _update_clash_report(self):
        w = self.clash_report
        _clear(w)

        student_clashes, teacher_clashes = _find_clashes(self.timetable_tree)
        is_legal = not student_clashes and not teacher_clashes

        if is_legal:
            _write(w, "PASS ✓  —  no student or teacher clashes found\n", "pass")
        else:
            n_s = len(student_clashes)
            n_t = len(teacher_clashes)
            _write(w, f"FAIL ✗  —  {n_s} student clash(es), {n_t} teacher clash(es)\n", "fail")

        _write(w, "─" * 60 + "\n", "dim")

        if student_clashes:
            _write(w, "\nSTUDENT DOUBLE-BOOKINGS\n", "heading")
            # Group by subblock
            by_sb: dict[str, list] = {}
            for c in student_clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, f"\n  Subblock {sb}\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["student"]):
                    classes = "  vs  ".join(entry["classes"])
                    _write(w, f"    Student {entry['student']:>6}:  {classes}\n", "fail")

        if teacher_clashes:
            _write(w, "\nTEACHER DOUBLE-BOOKINGS\n", "heading")
            by_sb = {}
            for c in teacher_clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, f"\n  Subblock {sb}\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["teacher"]):
                    classes = "  vs  ".join(entry["classes"])
                    _write(w, f"    {entry['teacher']:<20}:  {classes}\n", "fail")

        if is_legal:
            _write(w, "\n  No violations detected.\n", "dim")

        total = len(student_clashes) + len(teacher_clashes)
        _write(w, "\n" + "─" * 60 + "\n", "dim")
        _write(w, f"Total violations: {total}\n", "pass" if total == 0 else "fail")

    def _update_cost_panel(self):
        w = self.cost_text
        _clear(w)

        if not self.block_tree:
            return

        config = CostConfig()
        if self.teachers_path and self.teachers_path.exists():
            try:
                config.teacher_prefs = load_teacher_prefs_from_xlsx(
                    str(self.teachers_path)
                )
            except Exception:
                pass

        result = evaluate(self.block_tree, config)

        def row(label, value, is_stub=False):
            suffix = "  *stub*" if is_stub else ""
            tag = "bad" if value > 0 and not is_stub else "good"
            _write(w, f"  {label:<28} {value:>6}{suffix}\n", tag)

        _write(w, "E(T) cost breakdown\n", "good" if result.is_feasible() else "bad")
        _write(w, "─" * 42 + "\n")
        row("C_s   student clashes",     result.C_s)
        row("C_t   teacher clashes",     result.C_t)
        row("P_g12 Gr 12 teacher pref",  result.P_g12,  is_stub=True)
        row("P_tg  teacher grade pref",  result.P_tg,   is_stub=True)
        row("P_f   teacher free day",    result.P_f,    is_stub=True)
        row("P_stg sparse staggering",   result.P_stg)
        row("P_alloc allocation",        result.P_alloc, is_stub=True)
        _write(w, "─" * 42 + "\n")
        tag = "good" if result.total == 0 else "bad"
        _write(w, f"  {'E(T)  TOTAL':<28} {result.total:>6}\n", tag)
        _write(w, "\n")
        feasible_text = "Yes ✓" if result.is_feasible() else "No ✗"
        feasible_tag  = "good" if result.is_feasible() else "bad"
        _write(w, f"  Feasible (no hard clashes): ")
        _write(w, feasible_text + "\n", feasible_tag)

    def _update_qualification_panel(self):
        w = self.qual_text
        _clear(w)

        if not self.timetable_tree:
            return

        if not self.teacher_subj_map:
            _write(w, "Load teachers.xlsx to check qualifications.\n", "dim")
            return

        actual = _extract_teacher_subjects_from_tree(self.timetable_tree)
        issues = []
        ok_count = 0

        for teacher, subjects_taught in sorted(actual.items()):
            pool = self.teacher_subj_map.get(teacher)
            if pool is None:
                issues.append(
                    f"  {teacher:<16} not found in teachers.xlsx\n"
                )
                continue
            unqualified = subjects_taught - pool
            if unqualified:
                for subj in sorted(unqualified):
                    issues.append(
                        f"  {teacher:<16} teaching {subj} "
                        f"(pool: {', '.join(sorted(pool))})\n"
                    )
            else:
                ok_count += 1

        if issues:
            _write(w, f"{len(issues)} qualification issue(s) found:\n", "warn")
            _write(w, "─" * 44 + "\n", "dim")
            for line in issues:
                _write(w, line, "warn")
            _write(w, "─" * 44 + "\n", "dim")
            _write(w, f"\n  {ok_count} teacher(s) fully qualified.\n", "ok")
        else:
            _write(w, "PASS ✓  —  all teachers qualified for assigned subjects\n", "ok")
            _write(w, f"\n  {ok_count} teacher(s) checked.\n", "dim")

    # ─────────────────────────────────────────────────────────
    # POPULATE TIMETABLE TREE
    # ─────────────────────────────────────────────────────────

    def _populate_timetable_tree(self, filter_text=""):
        self.tt_tree.delete(*self.tt_tree.get_children())

        if not self.timetable_tree:
            return

        ft = filter_text.strip().lower()

        for block_name in sorted(self.timetable_tree.blocks.keys()):
            block = self.timetable_tree.blocks[block_name]
            block_node = self.tt_tree.insert(
                "", tk.END, text=f"Block {block_name}", open=bool(ft)
            )

            for sb_name in sorted(block.subblocks, key=lambda n: int(n[1:])):
                subblock = block.subblocks[sb_name]
                sb_node  = None

                for class_label in sorted(subblock.class_lists):
                    cl = subblock.class_lists[class_label]

                    if ft:
                        match = (
                            ft in class_label.lower() or
                            ft in str(cl.student_list.get_sorted())
                        )
                        if not match:
                            continue

                    if sb_node is None:
                        sb_node = self.tt_tree.insert(
                            block_node, tk.END, text=sb_name, open=bool(ft)
                        )

                    count   = len(cl.student_list)
                    cl_node = self.tt_tree.insert(
                        sb_node, tk.END,
                        text=f"{class_label}  ({count} students)"
                    )

                    students = cl.student_list.get_sorted()
                    for i in range(0, len(students), 20):
                        self.tt_tree.insert(
                            cl_node, tk.END, text=str(students[i:i + 20])
                        )

    # ─────────────────────────────────────────────────────────
    # POPULATE EXAM TREE
    # ─────────────────────────────────────────────────────────

    def _populate_exam_tree(self):
        self.ex_tree.delete(*self.ex_tree.get_children())

        if not self.exam_tree:
            return

        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node = self.exam_tree.grades[grade_label]
            grade_ui   = self.ex_tree.insert("", tk.END,
                                             text=grade_label, open=False)

            for subj_label in sorted(grade_node.exam_subjects.keys()):
                subject  = grade_node.exam_subjects[subj_label]
                subj_ui  = self.ex_tree.insert(grade_ui, tk.END,
                                               text=subj_label, open=False)

                for class_label in sorted(subject.class_lists.keys()):
                    cl    = subject.class_lists[class_label]
                    count = len(cl.student_list)
                    cl_ui = self.ex_tree.insert(
                        subj_ui, tk.END,
                        text=f"{class_label}  ({count} students)"
                    )
                    students = cl.student_list.get_sorted()
                    for i in range(0, len(students), 20):
                        self.ex_tree.insert(
                            cl_ui, tk.END, text=str(students[i:i + 20])
                        )

    # ─────────────────────────────────────────────────────────
    # REBUILD EXAM TREE
    # ─────────────────────────────────────────────────────────

    def _rebuild_exam(self):
        if not self.timetable_tree:
            return
        self.exam_tree = build_exam_tree_from_timetable_tree(self.timetable_tree)
        self._populate_exam_tree()
        self._update_slot_summary()

    def _update_slot_summary(self):
        if not self.exam_tree:
            return

        lines = []
        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node   = self.exam_tree.grades[grade_label]
            student_sets = {
                label: subject.all_students()
                for label, subject in grade_node.exam_subjects.items()
                if not is_excluded(label, self.exclusions)
            }
            if not student_sets:
                continue

            graph      = build_clash_graph(student_sets)
            assignment = dsatur_colouring(graph)
            num_slots  = max(assignment.values()) + 1

            slots: dict[int, list] = {}
            for subj, slot in assignment.items():
                slots.setdefault(slot, []).append(subj)

            lines.append(f"{grade_label}  —  {num_slots} slot(s)")
            for slot_num in sorted(slots):
                group         = sorted(slots[slot_num])
                slot_students = set()
                for s in group:
                    slot_students |= student_sets[s]
                lines.append(
                    f"  Slot {slot_num + 1:>2} "
                    f"({len(slot_students):>3} students): "
                    f"{', '.join(group)}"
                )
            lines.append("")

        _clear(self.slot_text)
        self.slot_text.config(state=tk.NORMAL)
        self.slot_text.insert(tk.END, "\n".join(lines))
        self.slot_text.config(state=tk.DISABLED)

    # ─────────────────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────────────────

    def _browse_export_path(self):
        path = filedialog.asksaveasfilename(
            title="Save optimised timetable as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.export_path_var.set(path)

    def _export(self):
        w = self.export_log
        _clear(w)

        if not self.timetable_tree:
            _write(w, "No timetable loaded — nothing to export.\n", "err")
            return

        out_path = self.export_path_var.get().strip()
        if not out_path:
            _write(w, "No output path set.\n", "err")
            return

        _write(w, "Verifying timetable before export…\n", "info")
        student_clashes, teacher_clashes = _find_clashes(self.timetable_tree)

        if student_clashes or teacher_clashes:
            _write(
                w,
                f"WARNING: {len(student_clashes)} student clash(es), "
                f"{len(teacher_clashes)} teacher clash(es) — "
                f"exporting anyway.\n",
                "err"
            )
        else:
            _write(w, "Verification passed ✓\n", "ok")

        # timetable_tree_to_block_tree() not yet written.
        # This stub logs the intention and will be wired once the
        # converter exists.
        _write(w, "\nExport not yet available.\n", "err")
        _write(
            w,
            "Waiting on:  timetable_converter.timetable_tree_to_block_tree()\n",
            "info"
        )
        _write(
            w,
            "Once that function exists, wire it here and call:\n"
            "  block_exporter.export_to_xlsx(block_tree, out_path)\n",
            "info"
        )

    # ─────────────────────────────────────────────────────────
    # SEARCH
    # ─────────────────────────────────────────────────────────

    def _on_search_change(self, *args):
        self._populate_timetable_tree(filter_text=self.search_var.get())

    # ─────────────────────────────────────────────────────────
    # EXCLUSION LIST
    # ─────────────────────────────────────────────────────────

    def _refresh_exclusion_listbox(self):
        self.excl_listbox.delete(0, tk.END)
        for code in sorted(self.exclusions):
            self.excl_listbox.insert(tk.END, code)

    def _add_exclusion(self):
        code = self.excl_entry.get().strip().upper()
        if not code:
            return
        self.exclusions.add(code)
        self.excl_entry.delete(0, tk.END)
        self._refresh_exclusion_listbox()
        self._update_slot_summary()

    def _remove_exclusion(self):
        sel = self.excl_listbox.curselection()
        if not sel:
            return
        code = self.excl_listbox.get(sel[0])
        self.exclusions.discard(code)
        self._refresh_exclusion_listbox()
        self._update_slot_summary()

    # ─────────────────────────────────────────────────────────
    # STYLING
    # ─────────────────────────────────────────────────────────

    def _style_tree(self, tree: ttk.Treeview):
        style = ttk.Style()
        style.configure(
            "Treeview",
            font=("Courier", 9), rowheight=22,
            background=CLR_WHITE, fieldbackground=CLR_WHITE
        )
        style.configure("Treeview.Heading",
                        font=("Helvetica", 10, "bold"))


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = TimePyBlingApp()
    app.mainloop()