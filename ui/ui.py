"""
ui.py — TimePyBling main interface

Tabs
----
  Timetable    — browse Block → SubBlock → Class → Students, with live search
  Verification — clash report, teacher qualification check, cost breakdown
  Exams        — exam slot scheduling + dated exam timetable per grade
  Export       — write optimised ST1.xlsx
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import date, timedelta
import pandas as pd

from core.timetable_tree     import build_timetable_tree_from_file
from reader.exam_tree        import build_exam_tree_from_timetable_tree
from reader.verify_timetable import _find_clashes
from reader.exam_paper       import ExamPaper, ExamPaperRegistry
from reader.exam_scheduler   import (build_schedule, DEFAULT_START_DATE,
                                     DEFAULT_TOTAL_DAYS, SESSIONS, EXAM_WEEKDAYS,
                                     _exam_dates)

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────

DEFAULT_EXCLUSIONS   = {"ST", "LIB", "PE", "RDI"}
TEACHER_SUBJECT_COLS = ["sua", "sub", "suc"]

CLR_HEADER   = "#2c3e50"
CLR_GREEN    = "#27ae60"
CLR_BLUE     = "#2980b9"
CLR_RED      = "#c0392b"
CLR_LIGHT    = "#ecf0f1"
CLR_MID      = "#bdc3c7"
CLR_WHITE    = "white"
CLR_BG       = "#f5f5f5"
CLR_MORNING  = "#e8f5e9"   # light green tint for morning rows
CLR_AFTERNOON= "#fff8e1"   # light amber tint for afternoon rows


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def _scrolled_text(parent, **kw) -> tk.Text:
    frame = tk.Frame(parent, bg=CLR_WHITE)
    frame.pack(fill=tk.BOTH, expand=True)
    sb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    t = tk.Text(frame, font=("Courier", 9), relief=tk.FLAT,
                bg="#f8f8f8", state=tk.DISABLED, wrap=tk.NONE,
                yscrollcommand=sb.set, **kw)
    t.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.config(command=t.yview)
    return t


def _write(widget: tk.Text, text: str, tag: str = ""):
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
    actual: dict[str, set[str]] = {}
    for block in tree.blocks.values():
        for subblock in block.subblocks.values():
            for label in subblock.class_lists:
                parts   = label.split("_")
                subject = parts[0]
                teacher = "_".join(parts[1:-1])
                if teacher:
                    actual.setdefault(teacher, set()).add(subject)
    return actual


def _strip_grade(label: str) -> str:
    """
    Strip the grade suffix for compact display in the schedule.
    "EN_12" -> "EN",  "CAT_12" -> "CAT"
    """
    return label.split("_")[0]


# ─────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────

class TimePyBlingApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("TimePyBling")
        self.geometry("1200x800")
        self.configure(bg=CLR_BG)

        self.timetable_tree   = None
        self.exam_tree        = None
        self.paper_registry   = None   # ExamPaperRegistry
        self.schedule_result  = None   # ScheduleResult
        self.st1_path         = None
        self.teachers_path    = None
        self.teacher_subj_map = {}
        self.exclusions       = set(DEFAULT_EXCLUSIONS)
        self._selected_paper_label = None   # label of paper selected in paper panel
        self._sessions        = None        # None = use start/end entries; set = custom list
        self.session_count_label = None     # set in _build_exam_tab

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

        tk.Label(bar, text="TimePyBling",
                 font=("Helvetica", 14, "bold"),
                 bg=CLR_HEADER, fg=CLR_WHITE).pack(side=tk.LEFT, padx=(0, 24))

        tk.Button(bar, text="Load Timetable", command=self._load_st1,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3).pack(side=tk.LEFT)

        self.st1_label = tk.Label(bar, text="No timetable loaded",
                                   bg=CLR_HEADER, fg=CLR_MID,
                                   font=("Helvetica", 9))
        self.st1_label.pack(side=tk.LEFT, padx=(6, 20))

        tk.Button(bar, text="Load Teachers", command=self._load_teachers,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3).pack(side=tk.LEFT)

        self.teachers_label = tk.Label(bar, text="No teachers loaded",
                                        bg=CLR_HEADER, fg=CLR_MID,
                                        font=("Helvetica", 9))
        self.teachers_label.pack(side=tk.LEFT, padx=(6, 20))

        tk.Button(bar, text="Load Students", state=tk.DISABLED,
                  bg="#555", fg=CLR_MID, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3).pack(side=tk.LEFT)

        tk.Label(bar, text="Coming soon",
                 bg=CLR_HEADER, fg="#666",
                 font=("Helvetica", 9, "italic")).pack(side=tk.LEFT, padx=(6, 0))

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_timetable    = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_verification = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_exams        = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_export       = tk.Frame(self.notebook, bg=CLR_WHITE)

        self.notebook.add(self.tab_timetable,    text="  Timetable  ")
        self.notebook.add(self.tab_verification, text="  Verification  ")
        self.notebook.add(self.tab_exams,        text="  Exams  ")
        self.notebook.add(self.tab_export,       text="  Export  ")

        self._build_timetable_tab()
        self._build_verification_tab()
        self._build_exam_tab()
        self._build_export_tab()

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
        tk.Entry(search_bar, textvariable=self.search_var,
                 font=("Helvetica", 10), width=30,
                 relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=6)

        tk.Label(search_bar,
                 text="student ID · subject code · teacher name",
                 bg=CLR_WHITE, fg="#888",
                 font=("Helvetica", 9)).pack(side=tk.LEFT)

        tk.Button(search_bar, text="Clear",
                  command=lambda: self.search_var.set(""),
                  relief=tk.FLAT, bg=CLR_LIGHT,
                  font=("Helvetica", 9), padx=8).pack(side=tk.LEFT, padx=4)

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
    # TAB 2 — VERIFICATION
    # ─────────────────────────────────────────────────────────

    def _build_verification_tab(self):
        pane = tk.PanedWindow(self.tab_verification, orient=tk.HORIZONTAL,
                              bg="#ccc", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)

        # Left: clash report
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=520)

        clash_hdr = tk.Frame(left, bg=CLR_WHITE)
        clash_hdr.pack(fill=tk.X, padx=8, pady=(8, 2))
        tk.Label(clash_hdr, text="Clash Report", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(clash_hdr, text="Re-run", command=self._run_verification,
                  bg=CLR_LIGHT, font=("Helvetica", 8),
                  relief=tk.FLAT, padx=8).pack(side=tk.RIGHT)

        clash_frame = tk.Frame(left, bg=CLR_WHITE)
        clash_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.clash_report = _scrolled_text(clash_frame)
        self.clash_report.tag_config("pass",    foreground=CLR_GREEN)
        self.clash_report.tag_config("fail",    foreground=CLR_RED)
        self.clash_report.tag_config("heading", foreground="#333",
                                     font=("Courier", 9, "bold"))
        self.clash_report.tag_config("warn",    foreground="#e67e22")
        self.clash_report.tag_config("dim",     foreground="#888")

        # Right: cost + teacher checks
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=300)

        cost_lf = tk.LabelFrame(right, text="Cost Function  E(T)",
                                 bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
                                 padx=6, pady=6)
        cost_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))
        self.cost_text = _scrolled_text(cost_lf)
        self.cost_text.tag_config("good", foreground=CLR_GREEN)
        self.cost_text.tag_config("bad",  foreground=CLR_RED)

        qual_lf = tk.LabelFrame(right, text="Teacher Qualifications",
                                 bg=CLR_WHITE, font=("Helvetica", 10, "bold"),
                                 padx=6, pady=6)
        qual_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))
        self.qual_text = _scrolled_text(qual_lf)
        self.qual_text.tag_config("ok",   foreground=CLR_GREEN)
        self.qual_text.tag_config("warn", foreground=CLR_RED)
        self.qual_text.tag_config("dim",  foreground="#888")

    # ─────────────────────────────────────────────────────────
    # TAB 3 — EXAMS
    # ─────────────────────────────────────────────────────────

    def _build_exam_tab(self):
        pane = tk.PanedWindow(self.tab_exams, orient=tk.HORIZONTAL,
                              bg="#ddd", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)

        # ── Left: exam tree + paper panel ──
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=440)

        left_top = tk.Frame(left, bg=CLR_WHITE)
        left_top.pack(fill=tk.X, padx=6, pady=(6, 0))
        tk.Label(left_top, text="Exam Tree", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(left_top, text="Rebuild", command=self._rebuild_exam,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8, "bold"), padx=8
                  ).pack(side=tk.RIGHT)

        tree_frame = tk.Frame(left, bg=CLR_WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(2, 0))

        self.ex_tree = ttk.Treeview(tree_frame, show="tree", selectmode="extended")
        ex_sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                               command=self.ex_tree.yview)
        ex_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.ex_tree.configure(yscrollcommand=ex_sb.set)
        self.ex_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._style_tree(self.ex_tree)
        self.ex_tree.bind("<<TreeviewSelect>>", self._on_exam_tree_select)

        # ── Paper panel (below exam tree) ──
        self.paper_lf = tk.LabelFrame(left, text="Papers for selected subject",
                                       bg=CLR_WHITE, font=("Helvetica", 9, "bold"),
                                       padx=6, pady=6)
        self.paper_lf.pack(fill=tk.X, padx=6, pady=6)

        self.paper_listbox = tk.Listbox(self.paper_lf, height=4,
                                         font=("Courier", 9),
                                         relief=tk.SOLID, bd=1,
                                         selectmode=tk.SINGLE)
        self.paper_listbox.pack(fill=tk.X, pady=(0, 4))
        self.paper_listbox.bind("<<ListboxSelect>>", self._on_paper_select)

        paper_btn_row = tk.Frame(self.paper_lf, bg=CLR_WHITE)
        paper_btn_row.pack(fill=tk.X, pady=(0, 4))
        tk.Button(paper_btn_row, text="+ Add Paper",
                  command=self._add_paper,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8, "bold"), padx=6
                  ).pack(side=tk.LEFT)
        tk.Button(paper_btn_row, text="− Remove",
                  command=self._remove_paper,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(paper_btn_row, text="📌 Pin…",
                  command=self._pin_paper,
                  bg="#8e44ad", fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6
                  ).pack(side=tk.LEFT)
        tk.Button(paper_btn_row, text="Unpin",
                  command=self._unpin_paper,
                  bg="#888", fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6
                  ).pack(side=tk.LEFT, padx=4)

        constr_row = tk.Frame(self.paper_lf, bg=CLR_WHITE)
        constr_row.pack(fill=tk.X)
        tk.Label(constr_row, text="Constraint:", bg=CLR_WHITE,
                 font=("Helvetica", 8)).pack(side=tk.LEFT)
        self.constr_entry = tk.Entry(constr_row, font=("Helvetica", 9),
                                      relief=tk.SOLID, bd=1, width=6)
        self.constr_entry.pack(side=tk.LEFT, padx=2)
        self.constr_entry.bind("<Return>", lambda e: self._add_constraint())
        self._constr_add_btn = tk.Button(constr_row, text="Add",
                                          command=self._add_constraint,
                                          bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                                          font=("Helvetica", 8), padx=4)
        self._constr_add_btn.pack(side=tk.LEFT)
        self._constr_remove_btn = tk.Button(constr_row, text="Remove",
                                             command=self._remove_constraint,
                                             bg="#888", fg=CLR_WHITE, relief=tk.FLAT,
                                             font=("Helvetica", 8), padx=4)
        self._constr_remove_btn.pack(side=tk.LEFT, padx=2)

        self.constr_listbox = tk.Listbox(self.paper_lf, height=2,
                                          font=("Courier", 9),
                                          relief=tk.SOLID, bd=1,
                                          selectmode=tk.SINGLE)
        self.constr_listbox.pack(fill=tk.X, pady=(4, 0))

        # ── Right: exclusions + scheduler controls + output ──
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=340)

        # Exclusions
        excl_frame = tk.LabelFrame(right, text="Exam Exclusions",
                                    bg=CLR_WHITE, font=("Helvetica", 9, "bold"),
                                    padx=8, pady=6)
        excl_frame.pack(fill=tk.X, padx=8, pady=(8, 4))

        self.excl_listbox = tk.Listbox(excl_frame, height=3,
                                        font=("Courier", 9),
                                        relief=tk.SOLID, bd=1,
                                        selectmode=tk.SINGLE)
        self.excl_listbox.pack(fill=tk.X, pady=(0, 4))
        self._refresh_exclusion_listbox()

        excl_row = tk.Frame(excl_frame, bg=CLR_WHITE)
        excl_row.pack(fill=tk.X)
        self.excl_entry = tk.Entry(excl_row, font=("Helvetica", 9),
                                    relief=tk.SOLID, bd=1, width=8)
        self.excl_entry.pack(side=tk.LEFT)
        self.excl_entry.bind("<Return>", lambda e: self._add_exclusion())
        tk.Button(excl_row, text="Add", command=self._add_exclusion,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8, "bold"), padx=6).pack(side=tk.LEFT, padx=3)
        tk.Button(excl_row, text="Remove", command=self._remove_exclusion,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6).pack(side=tk.LEFT)

        # Scheduler controls
        sched_lf = tk.LabelFrame(right, text="Exam Schedule",
                                   bg=CLR_WHITE, font=("Helvetica", 9, "bold"),
                                   padx=8, pady=6)
        sched_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        ctrl = tk.Frame(sched_lf, bg=CLR_WHITE)
        ctrl.pack(fill=tk.X, pady=(0, 4))

        # Row 0: Start / End date entries + live slot count
        tk.Label(ctrl, text="Start:", bg=CLR_WHITE,
                 font=("Helvetica", 9)).grid(row=0, column=0, sticky=tk.W)
        self.sched_start_var = tk.StringVar(
            value=DEFAULT_START_DATE.strftime("%Y-%m-%d"))
        self._start_entry = tk.Entry(ctrl, textvariable=self.sched_start_var,
                                      font=("Helvetica", 9), relief=tk.SOLID,
                                      bd=1, width=11)
        self._start_entry.grid(row=0, column=1, sticky=tk.W, padx=(2, 8))
        self._start_entry.bind("<FocusOut>", lambda e: self._update_session_count_label())
        self._start_entry.bind("<Return>",   lambda e: self._update_session_count_label())

        _default_end = _exam_dates(DEFAULT_START_DATE, DEFAULT_TOTAL_DAYS)[-1]
        tk.Label(ctrl, text="End:", bg=CLR_WHITE,
                 font=("Helvetica", 9)).grid(row=0, column=2, sticky=tk.W)
        self.sched_end_var = tk.StringVar(
            value=_default_end.strftime("%Y-%m-%d"))
        self._end_entry = tk.Entry(ctrl, textvariable=self.sched_end_var,
                                    font=("Helvetica", 9), relief=tk.SOLID,
                                    bd=1, width=11)
        self._end_entry.grid(row=0, column=3, sticky=tk.W, padx=(2, 8))
        self._end_entry.bind("<FocusOut>", lambda e: self._on_session_dates_changed())
        self._end_entry.bind("<Return>",   lambda e: self._on_session_dates_changed())

        self.session_count_label = tk.Label(ctrl, text="", bg=CLR_WHITE,
                                             fg=CLR_BLUE,
                                             font=("Helvetica", 9, "bold"))
        self.session_count_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 4))

        # Row 1: Configure sessions button
        tk.Button(ctrl, text="Configure sessions…",
                  command=self._open_session_calendar,
                  bg=CLR_LIGHT, font=("Helvetica", 8),
                  relief=tk.FLAT, padx=8
                  ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))

        # Row 2: Grade filter + Generate
        tk.Label(ctrl, text="Grade:", bg=CLR_WHITE,
                 font=("Helvetica", 9)).grid(row=2, column=0, sticky=tk.W,
                                              pady=(4, 0))
        self.sched_grade_var = tk.StringVar()
        self.sched_grade_cb  = ttk.Combobox(
            ctrl, textvariable=self.sched_grade_var,
            state="readonly", width=10, font=("Helvetica", 9))
        self.sched_grade_cb.grid(row=2, column=1, columnspan=2, sticky=tk.W,
                                  padx=(4, 12), pady=(4, 0))
        self.sched_grade_cb.bind("<<ComboboxSelected>>",
                                  lambda e: self._render_schedule())

        tk.Button(ctrl, text="Generate Schedule",
                  command=self._generate_exam_schedule,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3
                  ).grid(row=2, column=3, sticky=tk.W, padx=(4, 0), pady=(4, 0))

        tk.Button(ctrl, text="Export / View All",
                  command=self._open_schedule_popout,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3
                  ).grid(row=3, column=3, sticky=tk.W, padx=(4, 0), pady=(4, 0))

        # Initialise the slot count label
        self._update_session_count_label()

        # Schedule output
        self.sched_text = _scrolled_text(sched_lf)
        self.sched_text.tag_config("header",
                                    font=("Courier", 9, "bold"),
                                    foreground="#2c3e50")
        self.sched_text.tag_config("am",  background=CLR_MORNING)
        self.sched_text.tag_config("pm",  background=CLR_AFTERNOON)
        self.sched_text.tag_config("dim",     foreground="#888")
        self.sched_text.tag_config("warn",    foreground="#e67e22")
        self.sched_text.tag_config("day_sep", foreground="#bdc3c7")

    # ─────────────────────────────────────────────────────────
    # TAB 4 — EXPORT
    # ─────────────────────────────────────────────────────────

    def _build_export_tab(self):
        content = tk.Frame(self.tab_export, bg=CLR_WHITE, padx=16, pady=16)
        content.pack(fill=tk.BOTH, expand=True)

        path_frame = tk.Frame(content, bg=CLR_WHITE)
        path_frame.pack(fill=tk.X, pady=(0, 12))
        tk.Label(path_frame, text="Output path:", bg=CLR_WHITE,
                 font=("Helvetica", 10)).pack(side=tk.LEFT)
        self.export_path_var = tk.StringVar(value="output/ST1_optimised.xlsx")
        tk.Entry(path_frame, textvariable=self.export_path_var,
                 font=("Helvetica", 10), relief=tk.SOLID, bd=1,
                 width=50).pack(side=tk.LEFT, padx=8)
        tk.Button(path_frame, text="Browse…",
                  command=self._browse_export_path,
                  bg=CLR_LIGHT, font=("Helvetica", 9),
                  relief=tk.FLAT, padx=8).pack(side=tk.LEFT)

        tk.Button(content, text="Export ST1.xlsx", command=self._export,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 11, "bold"), padx=16, pady=8
                  ).pack(anchor=tk.W, pady=(0, 12))

        tk.Label(content, text="Export log", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(anchor=tk.W)
        log_frame = tk.Frame(content, bg=CLR_WHITE)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        self.export_log = _scrolled_text(log_frame)
        self.export_log.tag_config("ok",   foreground=CLR_GREEN)
        self.export_log.tag_config("err",  foreground=CLR_RED)
        self.export_log.tag_config("info", foreground=CLR_BLUE)

    # ─────────────────────────────────────────────────────────
    # LOAD
    # ─────────────────────────────────────────────────────────

    def _load_st1(self):
        path = filedialog.askopenfilename(
            title="Select student timetable (ST1.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")])
        if not path:
            return
        self.st1_path = Path(path)
        self.st1_label.config(text=self.st1_path.name, fg=CLR_WHITE)
        try:
            self.timetable_tree = build_timetable_tree_from_file(self.st1_path)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        suggested = self.st1_path.parent / "ST1_optimised.xlsx"
        self.export_path_var.set(str(suggested))
        self._populate_timetable_tree()
        self._rebuild_exam()
        self._run_verification()
        self.notebook.select(self.tab_exams)

    def _load_teachers(self):
        path = filedialog.askopenfilename(
            title="Select teachers file (teachers.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")])
        if not path:
            return
        self.teachers_path = Path(path)
        self.teachers_label.config(text=self.teachers_path.name, fg=CLR_WHITE)
        try:
            self.teacher_subj_map = _load_teacher_subject_map(self.teachers_path)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        if self.timetable_tree:
            self._run_verification()

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
            _write(w, f"FAIL ✗  —  {len(student_clashes)} student clash(es), "
                      f"{len(teacher_clashes)} teacher clash(es)\n", "fail")
        _write(w, "─" * 60 + "\n", "dim")

        if student_clashes:
            _write(w, "\nSTUDENT DOUBLE-BOOKINGS\n", "heading")
            by_sb: dict[str, list] = {}
            for c in student_clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, f"\n  Subblock {sb}\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["student"]):
                    _write(w, f"    Student {entry['student']:>6}:  "
                              f"{'  vs  '.join(entry['classes'])}\n", "fail")

        if teacher_clashes:
            _write(w, "\nTEACHER DOUBLE-BOOKINGS\n", "heading")
            by_sb = {}
            for c in teacher_clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, f"\n  Subblock {sb}\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["teacher"]):
                    _write(w, f"    {entry['teacher']:<20}:  "
                              f"{'  vs  '.join(entry['classes'])}\n", "fail")

        total = len(student_clashes) + len(teacher_clashes)
        _write(w, "\n" + "─" * 60 + "\n", "dim")
        _write(w, f"Total violations: {total}\n",
               "pass" if total == 0 else "fail")

    def _update_cost_panel(self):
        w = self.cost_text
        _clear(w)
        _write(w, "Cost function (optimiser module) not yet available.\n", "bad")

    def _update_qualification_panel(self):
        w = self.qual_text
        _clear(w)
        if not self.timetable_tree:
            return
        if not self.teacher_subj_map:
            _write(w, "Load teachers.xlsx to check qualifications.\n", "dim")
            return
        actual = _extract_teacher_subjects_from_tree(self.timetable_tree)
        issues, ok_count = [], 0
        for teacher, subjects_taught in sorted(actual.items()):
            pool = self.teacher_subj_map.get(teacher)
            if pool is None:
                issues.append(f"  {teacher:<16} not found in teachers.xlsx\n")
                continue
            unqualified = subjects_taught - pool
            if unqualified:
                for subj in sorted(unqualified):
                    issues.append(f"  {teacher:<16} teaching {subj} "
                                  f"(pool: {', '.join(sorted(pool))})\n")
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
            _write(w, "PASS ✓  —  all teachers qualified\n", "ok")
            _write(w, f"\n  {ok_count} teacher(s) checked.\n", "dim")

    # ─────────────────────────────────────────────────────────
    # TIMETABLE TREE
    # ─────────────────────────────────────────────────────────

    def _populate_timetable_tree(self, filter_text=""):
        self.tt_tree.delete(*self.tt_tree.get_children())
        if not self.timetable_tree:
            return
        ft = filter_text.strip().lower()
        for block_name in sorted(self.timetable_tree.blocks.keys()):
            block      = self.timetable_tree.blocks[block_name]
            block_node = self.tt_tree.insert("", tk.END,
                                             text=f"Block {block_name}",
                                             open=bool(ft))
            for sb_name in sorted(block.subblocks, key=lambda n: int(n[1:])):
                subblock = block.subblocks[sb_name]
                sb_node  = None
                for class_label in sorted(subblock.class_lists):
                    cl = subblock.class_lists[class_label]
                    if ft and not (ft in class_label.lower() or
                                   ft in str(cl.student_list.get_sorted())):
                        continue
                    if sb_node is None:
                        sb_node = self.tt_tree.insert(block_node, tk.END,
                                                       text=sb_name,
                                                       open=bool(ft))
                    count   = len(cl.student_list)
                    cl_node = self.tt_tree.insert(sb_node, tk.END,
                                                   text=f"{class_label}  ({count} students)")
                    students = cl.student_list.get_sorted()
                    for i in range(0, len(students), 20):
                        self.tt_tree.insert(cl_node, tk.END,
                                             text=str(students[i:i + 20]))

    def _on_search_change(self, *args):
        self._populate_timetable_tree(filter_text=self.search_var.get())

    # ─────────────────────────────────────────────────────────
    # EXAM TREE
    # ─────────────────────────────────────────────────────────

    def _populate_exam_tree(self):
        self.ex_tree.delete(*self.ex_tree.get_children())
        if not self.paper_registry:
            return
        for grade in self.paper_registry.grades():
            grade_ui = self.ex_tree.insert("", tk.END, text=grade, open=False)
            for subj in self.paper_registry.subjects_for_grade(grade):
                papers = self.paper_registry.papers_for_subject_grade(subj, grade)
                count  = max(p.student_count() for p in papers)
                self.ex_tree.insert(grade_ui, tk.END,
                                    text=f"{subj}  ({count} students)",
                                    tags=("subject",),
                                    values=(subj, grade))

    def _exam_tree_get_state(self) -> tuple[set[str], list[tuple[str, str]]]:
        """Return (expanded_grade_texts, selected_(subject,grade)_pairs)."""
        expanded: set[str] = set()
        for item in self.ex_tree.get_children():
            if self.ex_tree.item(item, "open"):
                expanded.add(self.ex_tree.item(item, "text"))
        selected_pairs: list[tuple[str, str]] = []
        for item in self.ex_tree.selection():
            tags = self.ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self.ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    selected_pairs.append((vals[0], vals[1]))
        return expanded, selected_pairs

    def _exam_tree_restore_state(self, expanded: set[str],
                                  selected_pairs: list[tuple[str, str]]):
        """Re-expand grade nodes and re-select subject nodes after a rebuild."""
        for item in self.ex_tree.get_children():
            if self.ex_tree.item(item, "text") in expanded:
                self.ex_tree.item(item, open=True)
        if not selected_pairs:
            return
        pair_set = set(selected_pairs)
        to_select = []
        for grade_item in self.ex_tree.get_children():
            for subj_item in self.ex_tree.get_children(grade_item):
                tags = self.ex_tree.item(subj_item, "tags")
                if "subject" not in (tags or []):
                    continue
                vals = self.ex_tree.item(subj_item, "values")
                if vals and len(vals) >= 2 and (vals[0], vals[1]) in pair_set:
                    to_select.append(subj_item)
        if to_select:
            self.ex_tree.selection_set(to_select)
            self.ex_tree.see(to_select[0])

    def _on_exam_tree_select(self, event=None):
        sel = self.ex_tree.selection()
        if not sel:
            self._refresh_paper_panel(None, None)
            return
        subject_pairs = []
        for item in sel:
            tags = self.ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self.ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if not subject_pairs:
            self._refresh_paper_panel(None, None)
        elif len(subject_pairs) == 1:
            self._refresh_paper_panel(subject_pairs[0][0], subject_pairs[0][1])
        else:
            self._refresh_paper_panel_multi(subject_pairs)

    def _refresh_paper_panel(self, subject: str | None, grade: str | None):
        self.paper_listbox.delete(0, tk.END)
        self.constr_listbox.delete(0, tk.END)
        self._selected_paper_label = None
        self.paper_lf.config(text="Papers for selected subject")
        self._set_constraint_ui_enabled(True)
        if not subject or not grade or not self.paper_registry:
            return
        papers     = self.paper_registry.papers_for_subject_grade(subject, grade)
        pin_clashes = (self.schedule_result.pin_clash_warnings
                       if self.schedule_result else {})
        for p in papers:
            constr_str    = (", ".join(sorted(p.constraints))
                             if p.constraints else "no constraints")
            pin_indicator = "📌 " if p.pinned_slot is not None else "   "
            clash_flag    = "  ⚠ pin clash" if p.label in pin_clashes else ""
            self.paper_listbox.insert(
                tk.END,
                f"{pin_indicator}{p.subject} P{p.paper_number}"
                f" — {p.student_count()} students — {constr_str}{clash_flag}"
            )
        if papers:
            self.paper_listbox.selection_set(0)
            self._selected_paper_label = papers[0].label
            self._refresh_constraint_list(papers[0])

    def _refresh_paper_panel_multi(self, subject_grade_pairs: list[tuple[str, str]]):
        """Summary view when multiple subjects are selected."""
        self.paper_listbox.delete(0, tk.END)
        self.constr_listbox.delete(0, tk.END)
        self._selected_paper_label = None
        n = len(subject_grade_pairs)
        self.paper_lf.config(text=f"{n} subjects selected")
        self.paper_listbox.insert(tk.END, f"{n} subjects selected")
        if self.paper_registry:
            addable = sum(
                1 for subj, grade in subject_grade_pairs
                if self.paper_registry.papers_for_subject_grade(subj, grade)
                and max(p.paper_number
                        for p in self.paper_registry.papers_for_subject_grade(subj, grade)
                        ) < 3
            )
            if addable:
                self.paper_listbox.insert(
                    tk.END, f"  {addable} eligible for [+ Add Paper]")
        self._set_constraint_ui_enabled(False)
        self.constr_listbox.insert(
            tk.END, "Select a single paper to edit constraints")

    def _set_constraint_ui_enabled(self, enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        self.constr_entry.config(state=state)
        self._constr_add_btn.config(state=state)
        self._constr_remove_btn.config(state=state)

    def _on_paper_select(self, event=None):
        sel = self.paper_listbox.curselection()
        if not sel:
            return
        tree_sel = self.ex_tree.selection()
        if not tree_sel:
            return
        # Only act if exactly one subject is selected in the tree
        subject_pairs = []
        for item in tree_sel:
            tags = self.ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self.ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if len(subject_pairs) != 1:
            return
        subject, grade = subject_pairs[0]
        papers = self.paper_registry.papers_for_subject_grade(subject, grade)
        idx = sel[0]
        if idx < len(papers):
            self._selected_paper_label = papers[idx].label
            self._refresh_constraint_list(papers[idx])

    def _refresh_constraint_list(self, paper):
        self.constr_listbox.delete(0, tk.END)
        for code in sorted(paper.constraints):
            self.constr_listbox.insert(tk.END, code)

    def _rebuild_exam(self):
        if not self.timetable_tree:
            return
        self.exam_tree      = build_exam_tree_from_timetable_tree(self.timetable_tree)
        self.paper_registry = ExamPaperRegistry.from_exam_tree(
            self.exam_tree, exclusions=self.exclusions)
        self.schedule_result = None
        self._selected_paper_label = None
        self._populate_exam_tree()
        self._update_sched_grade_list()

    def _update_sched_grade_list(self):
        if not self.paper_registry:
            return
        grades = self.paper_registry.grades()
        options = ["All Grades"] + grades
        self.sched_grade_cb["values"] = options
        if grades:
            self.sched_grade_var.set(grades[-1])

    # ─────────────────────────────────────────────────────────
    # PAPER PANEL ACTIONS
    # ─────────────────────────────────────────────────────────

    def _current_subject_grade(self) -> tuple[str, str] | tuple[None, None]:
        tree_sel = self.ex_tree.selection()
        if not tree_sel:
            return None, None
        vals = self.ex_tree.item(tree_sel[0], "values")
        if not vals or len(vals) < 2:
            return None, None
        return vals[0], vals[1]

    def _add_paper(self):
        if not self.paper_registry:
            return
        # Collect all selected subject nodes
        sel = self.ex_tree.selection()
        subject_pairs = []
        for item in sel:
            tags = self.ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self.ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if not subject_pairs:
            return

        expanded, selected_pairs = self._exam_tree_get_state()

        if len(subject_pairs) == 1:
            subject, grade = subject_pairs[0]
            paper = self.paper_registry.add_paper(subject, grade)
            if paper is None:
                messagebox.showinfo("Cannot add",
                                    "Maximum 3 papers per subject per grade.")
                return
        else:
            added = sum(
                1 for subj, grade in subject_pairs
                if self.paper_registry.add_paper(subj, grade) is not None
            )
            if added == 0:
                messagebox.showinfo("Cannot add",
                                    "All selected subjects already have 3 papers.")
                return

        self._populate_exam_tree()
        self._exam_tree_restore_state(expanded, selected_pairs)
        # Refresh paper panel to match restored selection
        self._on_exam_tree_select()

    def _remove_paper(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
        # Parse subject+grade from the label (format: SUBJ_PN_GRADE)
        parts = self._selected_paper_label.split("_")
        subject = parts[0] if parts else None
        grade   = parts[-1] if len(parts) >= 3 else None

        expanded, selected_pairs = self._exam_tree_get_state()

        removed = self.paper_registry.remove_paper(self._selected_paper_label)
        if not removed:
            messagebox.showinfo("Cannot remove",
                                "Cannot remove the only paper (P1) for a subject.")
            return

        self._populate_exam_tree()
        self._exam_tree_restore_state(expanded, selected_pairs)
        self._refresh_paper_panel(subject, grade)

    def _add_constraint(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
        code = self.constr_entry.get().strip().upper()
        if not code:
            return
        self.paper_registry.add_constraint(self._selected_paper_label, code)
        self.constr_entry.delete(0, tk.END)
        paper = self.paper_registry.get(self._selected_paper_label)
        if paper:
            self._refresh_constraint_list(paper)
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

    def _remove_constraint(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
        sel = self.constr_listbox.curselection()
        if not sel:
            return
        code = self.constr_listbox.get(sel[0])
        self.paper_registry.remove_constraint(self._selected_paper_label, code)
        paper = self.paper_registry.get(self._selected_paper_label)
        if paper:
            self._refresh_constraint_list(paper)
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

    # ─────────────────────────────────────────────────────────
    # EXAM SCHEDULE GENERATOR
    # ─────────────────────────────────────────────────────────

    def _generate_exam_schedule(self):
        if not self.paper_registry:
            messagebox.showinfo("No data", "Load a timetable first.")
            return
        sessions = self._get_effective_sessions()
        if sessions is None:
            messagebox.showerror("Invalid dates",
                                  "Enter valid start and end dates (YYYY-MM-DD).")
            return
        if not sessions:
            messagebox.showerror("No sessions",
                                  "No exam sessions in the selected date range.")
            return
        self.schedule_result = build_schedule(
            registry  = self.paper_registry,
            sessions  = sessions,
            exam_tree = self.exam_tree,
        )
        self._render_schedule()

    def _render_schedule(self):
        w = self.sched_text
        _clear(w)
        if not self.schedule_result:
            return

        result      = self.schedule_result
        grade_filter = self.sched_grade_var.get().strip()

        if grade_filter and grade_filter != "All Grades":
            items = [sp for sp in result.scheduled
                     if sp.paper.grade == grade_filter]
        else:
            items = list(result.scheduled)

        if not items:
            _write(w, f"No papers scheduled for {grade_filter}.\n", "warn")
            return

        items.sort(key=lambda sp: (sp.slot_index, sp.paper.grade))

        _write(w, f"Exam Schedule"
                  f"  —  {result.total_days} day(s)"
                  f"  —  {result.total_slots} slot(s)"
                  f"  —  {len(result.scheduled)} papers total\n", "header")
        if grade_filter and grade_filter != "All Grades":
            _write(w, f"Filtered: {grade_filter}\n", "dim")
        _write(w, "─" * 68 + "\n", "dim")
        _write(w, f"  {'Date':<14} {'Sess':<5} {'Paper':<22} "
                  f"{'Students':>8}  Warnings\n", "header")
        _write(w, "─" * 68 + "\n", "dim")

        prev_date = None
        for sp in items:
            if prev_date is not None and sp.date != prev_date:
                _write(w, "\n", "day_sep")
            prev_date = sp.date

            date_str  = sp.date.strftime("%a %d %b")
            pin_mark  = " 📌" if sp.pinned else ""
            warn_flag = "  ⚠" if sp.warnings else ""
            line = (f"  {date_str:<14} {sp.session:<5} "
                    f"{sp.paper.label:<22} "
                    f"{sp.paper.student_count():>8}{pin_mark}{warn_flag}\n")
            tag = "am" if sp.session == "AM" else "pm"
            _write(w, line, tag)

        _write(w, "─" * 68 + "\n", "dim")

        # Warnings section
        all_warnings = [
            w_msg
            for sp in items
            for w_msg in sp.warnings
        ]
        seen: set[str] = set()
        unique_warnings = [x for x in all_warnings
                           if not (x in seen or seen.add(x))]
        if unique_warnings:
            _write(w, "\nWarnings:\n", "warn")
            for msg in unique_warnings:
                _write(w, f"  ⚠  {msg}\n", "warn")

        if self.schedule_result and self.schedule_result.teacher_warnings:
            _write(w, "\nTeacher marking load conflicts:\n", "warn")
            for msg in self.schedule_result.teacher_warnings:
                _write(w, f"  ⚠  {msg}\n", "warn")

    # ─────────────────────────────────────────────────────────
    # SESSION CALENDAR
    # ─────────────────────────────────────────────────────────

    def _get_effective_sessions(self) -> list[tuple[date, str]] | None:
        """
        Return the active session list.
        If self._sessions is set (from popup), return it directly.
        Otherwise compute from start/end date entries.
        Returns None if dates are invalid.
        """
        if self._sessions is not None:
            return self._sessions
        try:
            start = date.fromisoformat(self.sched_start_var.get().strip())
            end   = date.fromisoformat(self.sched_end_var.get().strip())
        except ValueError:
            return None
        if end < start:
            return None
        days: list[date] = []
        d = start
        while d <= end:
            if d.weekday() in EXAM_WEEKDAYS:
                days.append(d)
            d += timedelta(days=1)
        return [(day, sess) for day in days for sess in SESSIONS]

    def _update_session_count_label(self):
        if self.session_count_label is None:
            return
        sessions = self._get_effective_sessions()
        if sessions is None:
            self.session_count_label.config(text="invalid dates", fg=CLR_RED)
            return
        n_days  = len({d for d, _ in sessions})
        n_slots = len(sessions)
        custom  = self._sessions is not None
        txt = f"{n_days} days  →  {n_slots} slots"
        if custom:
            txt += "  (custom)"
        self.session_count_label.config(
            text=txt,
            fg="#8e44ad" if custom else CLR_BLUE
        )

    def _on_session_dates_changed(self):
        """Called when start or end date entry loses focus — reset custom sessions."""
        self._sessions = None
        self._update_session_count_label()

    def _open_session_calendar(self):
        """Open a popup to toggle individual AM/PM sessions on or off."""
        sessions = self._get_effective_sessions()
        if not sessions:
            messagebox.showerror("Invalid dates",
                                  "Enter valid start and end dates first.")
            return

        # Build day → {AM: bool, PM: bool} map from current effective sessions
        day_state: dict[date, dict[str, bool]] = {}
        for d, s in sessions:
            day_state.setdefault(d, {"AM": False, "PM": False})
            day_state[d][s] = True
        # If we have custom sessions, mark unchecked slots too
        if self._sessions is not None:
            for d in day_state:
                for s in SESSIONS:
                    if (d, s) not in {(ds, ss) for ds, ss in self._sessions}:
                        day_state[d][s] = False

        top = tk.Toplevel(self)
        top.title("Configure Exam Sessions")
        top.geometry("360x480")
        top.configure(bg=CLR_WHITE)

        tk.Label(top, text="Toggle AM/PM sessions on or off:",
                 bg=CLR_WHITE, font=("Helvetica", 10, "bold")
                 ).pack(pady=(10, 4), padx=10, anchor=tk.W)

        cf = tk.Frame(top, bg=CLR_WHITE)
        cf.pack(fill=tk.BOTH, expand=True, padx=10)
        canvas = tk.Canvas(cf, bg=CLR_WHITE, highlightthickness=0)
        v_sb   = ttk.Scrollbar(cf, orient=tk.VERTICAL, command=canvas.yview)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(yscrollcommand=v_sb.set)
        inner = tk.Frame(canvas, bg=CLR_WHITE)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        session_vars: dict[tuple[date, str], tk.BooleanVar] = {}
        for i, (d, state) in enumerate(sorted(day_state.items())):
            bg  = CLR_MORNING if i % 2 == 0 else CLR_WHITE
            row = tk.Frame(inner, bg=bg)
            row.pack(fill=tk.X, pady=1)
            tk.Label(row, text=d.strftime("%a %d %b"), bg=bg,
                     font=("Courier", 9), width=14, anchor=tk.W
                     ).pack(side=tk.LEFT, padx=(4, 8))
            for sess in SESSIONS:
                v = tk.BooleanVar(value=state.get(sess, True))
                session_vars[(d, sess)] = v
                tk.Checkbutton(row, text=sess, variable=v,
                               bg=bg, font=("Helvetica", 9)
                               ).pack(side=tk.LEFT, padx=2)

        inner.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        btn_frame = tk.Frame(top, bg=CLR_WHITE, pady=8)
        btn_frame.pack(fill=tk.X, padx=10)

        def _apply():
            new_sessions = [
                (d, s)
                for (d, s), v in sorted(session_vars.items())
                if v.get()
            ]
            if not new_sessions:
                messagebox.showwarning("No sessions",
                                       "At least one session must be enabled.",
                                       parent=top)
                return
            self._sessions = new_sessions
            self._update_session_count_label()
            top.destroy()

        def _reset_all():
            for v in session_vars.values():
                v.set(True)

        tk.Button(btn_frame, text="Apply", command=_apply,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=16, pady=4
                  ).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=top.destroy,
                  bg=CLR_LIGHT, font=("Helvetica", 9),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Reset to all on", command=_reset_all,
                  bg=CLR_LIGHT, font=("Helvetica", 9),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT)

    # ─────────────────────────────────────────────────────────
    # SLOT PINNING
    # ─────────────────────────────────────────────────────────

    def _pin_paper(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
        paper = self.paper_registry.get(self._selected_paper_label)
        if not paper:
            return
        sessions = self._get_effective_sessions()
        if not sessions:
            messagebox.showinfo("No sessions",
                                "Configure exam sessions (start/end dates) first.")
            return
        self._open_slot_picker(paper, sessions)

    def _unpin_paper(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
        paper = self.paper_registry.get(self._selected_paper_label)
        if paper and paper.pinned_slot is not None:
            paper.pinned_slot = None
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

    def _open_slot_picker(self, paper: ExamPaper,
                           sessions: list[tuple[date, str]]):
        top = tk.Toplevel(self)
        top.title(f"Pin {paper.label} to slot")
        top.geometry("310x420")
        top.configure(bg=CLR_WHITE)

        tk.Label(top, text=f"Select a slot for  {paper.label}:",
                 bg=CLR_WHITE, font=("Helvetica", 10, "bold")
                 ).pack(pady=(10, 4), padx=10, anchor=tk.W)

        lb_frame = tk.Frame(top, bg=CLR_WHITE)
        lb_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        lb_sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL)
        lb_sb.pack(side=tk.RIGHT, fill=tk.Y)
        lb = tk.Listbox(lb_frame, font=("Courier", 9), relief=tk.SOLID, bd=1,
                        selectmode=tk.SINGLE, yscrollcommand=lb_sb.set)
        lb_sb.config(command=lb.yview)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, (d, sess) in enumerate(sessions):
            marker = "  📌" if paper.pinned_slot == i else ""
            bg     = CLR_MORNING if sess == "AM" else CLR_AFTERNOON
            lb.insert(tk.END,
                      f"Slot {i + 1:>3}  {d.strftime('%a %d %b')}  {sess}{marker}")
            lb.itemconfig(i, background=bg)

        if paper.pinned_slot is not None and paper.pinned_slot < len(sessions):
            lb.selection_set(paper.pinned_slot)
            lb.see(paper.pinned_slot)

        btn_frame = tk.Frame(top, bg=CLR_WHITE, pady=8)
        btn_frame.pack(fill=tk.X, padx=10)

        def _pin():
            sel = lb.curselection()
            if not sel:
                return
            paper.pinned_slot = sel[0]
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)
            top.destroy()

        tk.Button(btn_frame, text="Pin to slot", command=_pin,
                  bg="#8e44ad", fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=16, pady=4
                  ).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=top.destroy,
                  bg=CLR_LIGHT, font=("Helvetica", 9),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT, padx=8)

    # ─────────────────────────────────────────────────────────
    # SCHEDULE POPOUT  (Task 3)
    # ─────────────────────────────────────────────────────────

    def _open_schedule_popout(self):
        if not self.schedule_result:
            messagebox.showinfo("No schedule", "Generate a schedule first.")
            return

        from collections import defaultdict

        result = self.schedule_result
        grades = sorted(
            {sp.paper.grade for sp in result.scheduled},
            key=lambda g: int(g[2:]) if g[2:].isdigit() else 0
        )

        # grid[slot_index][grade] = [subject_codes]
        grid: dict[int, dict[str, list[str]]] = defaultdict(
            lambda: {g: [] for g in grades})
        slot_meta: dict[int, tuple] = {}

        for sp in result.scheduled:
            grid[sp.slot_index][sp.paper.grade].append(sp.paper.subject)
            slot_meta[sp.slot_index] = (sp.date, sp.session)

        all_slots = sorted(grid.keys())

        top = tk.Toplevel(self)
        top.title("Full Exam Schedule — All Grades")
        top.geometry("980x620")
        top.configure(bg=CLR_WHITE)

        bar = tk.Frame(top, bg=CLR_LIGHT, pady=6, padx=10)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="Full Exam Schedule — All Grades",
                 font=("Helvetica", 11, "bold"), bg=CLR_LIGHT).pack(side=tk.LEFT)
        tk.Button(bar, text="Save as PDF / TXT",
                  command=lambda: self._save_schedule_export(
                      top, grades, grid, slot_meta, all_slots),
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10
                  ).pack(side=tk.RIGHT)

        # Scrollable canvas
        cf = tk.Frame(top, bg=CLR_WHITE)
        cf.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(cf, bg=CLR_WHITE, highlightthickness=0)
        h_sb = ttk.Scrollbar(cf, orient=tk.HORIZONTAL, command=canvas.xview)
        v_sb = ttk.Scrollbar(cf, orient=tk.VERTICAL,   command=canvas.yview)
        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        v_sb.pack(side=tk.RIGHT,  fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(xscrollcommand=h_sb.set, yscrollcommand=v_sb.set)

        gf = tk.Frame(canvas, bg=CLR_WHITE)
        canvas.create_window((0, 0), window=gf, anchor="nw")

        COL_W = 10   # character width for grade columns
        ROW_W = 22   # character width for row-label column
        HDR_BG = "#d5d8dc"

        # Header row
        tk.Label(gf, text="Slot / Date / Sess", font=("Helvetica", 8, "bold"),
                 bg=HDR_BG, width=ROW_W, anchor="w",
                 relief=tk.RIDGE, bd=1).grid(row=0, column=0, sticky="nsew",
                                              padx=1, pady=1)
        for ci, grade in enumerate(grades):
            tk.Label(gf, text=grade, font=("Helvetica", 8, "bold"),
                     bg=HDR_BG, width=COL_W, anchor="center",
                     relief=tk.RIDGE, bd=1).grid(row=0, column=ci + 1,
                                                  sticky="nsew", padx=1, pady=1)

        # Data rows
        for ri, slot_idx in enumerate(all_slots):
            d, session = slot_meta[slot_idx]
            row_bg = CLR_MORNING if session == "AM" else CLR_AFTERNOON
            lbl = f"Slot {slot_idx+1}  {d.strftime('%a %d %b')}  {session}"
            tk.Label(gf, text=lbl, font=("Courier", 8),
                     bg=row_bg, width=ROW_W, anchor="w",
                     relief=tk.RIDGE, bd=1).grid(row=ri + 1, column=0,
                                                  sticky="nsew", padx=1, pady=1)
            for ci, grade in enumerate(grades):
                subjects = sorted(grid[slot_idx][grade])
                cell = ", ".join(subjects) if subjects else ""
                tk.Label(gf, text=cell, font=("Courier", 8),
                         bg=row_bg, width=COL_W, anchor="center",
                         relief=tk.RIDGE, bd=1).grid(row=ri + 1, column=ci + 1,
                                                      sticky="nsew", padx=1, pady=1)

        gf.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _save_schedule_export(self, top, grades, grid, slot_meta, all_slots):
        try:
            import reportlab  # noqa: F401
            has_reportlab = True
        except ImportError:
            has_reportlab = False

        if has_reportlab:
            path = filedialog.asksaveasfilename(
                parent=top, title="Save exam schedule as PDF",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
            if not path:
                return
            self._save_schedule_pdf(path, grades, grid, slot_meta, all_slots)
            messagebox.showinfo("Saved", f"Schedule saved as PDF:\n{path}", parent=top)
        else:
            path = filedialog.asksaveasfilename(
                parent=top,
                title="Save exam schedule as text  (install reportlab for PDF)",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
            if not path:
                return
            self._save_schedule_txt(path, grades, grid, slot_meta, all_slots)
            messagebox.showinfo(
                "Saved",
                f"Schedule saved as text:\n{path}\n\n"
                "Tip: pip install reportlab for PDF export.",
                parent=top)

    def _save_schedule_pdf(self, path, grades, grid, slot_meta, all_slots):
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib import colors

        doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                                leftMargin=18, rightMargin=18,
                                topMargin=18, bottomMargin=18)

        header_row = ["Slot / Date / Session"] + grades
        rows = [header_row]
        for slot_idx in all_slots:
            d, session = slot_meta[slot_idx]
            label = f"Slot {slot_idx+1}  {d.strftime('%a %d %b')}  {session}"
            row = [label]
            for grade in grades:
                subjects = sorted(grid[slot_idx][grade])
                row.append(", ".join(subjects) if subjects else "")
            rows.append(row)

        col_widths = [130] + [50] * len(grades)
        table = Table(rows, colWidths=col_widths, repeatRows=1)

        am_bg = colors.Color(0.91, 0.96, 0.91)
        pm_bg = colors.Color(1.0,  0.97, 0.88)

        style_cmds = [
            ("FONTNAME",   (0, 0), (-1, 0),  "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1, -1), 7),
            ("BACKGROUND", (0, 0), (-1, 0),  colors.Color(0.84, 0.86, 0.86)),
            ("GRID",       (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN",      (1, 1), (-1, -1), "CENTER"),
        ]
        for ri, slot_idx in enumerate(all_slots, start=1):
            _, session = slot_meta[slot_idx]
            bg = am_bg if session == "AM" else pm_bg
            style_cmds.append(("BACKGROUND", (0, ri), (-1, ri), bg))

        table.setStyle(TableStyle(style_cmds))
        doc.build([table])

    def _save_schedule_txt(self, path, grades, grid, slot_meta, all_slots):
        col_w = 14
        header = f"{'Slot / Date / Sess':<26}" + "".join(
            f"{g:<{col_w}}" for g in grades)
        lines = [header, "-" * len(header)]
        for slot_idx in all_slots:
            d, session = slot_meta[slot_idx]
            label = f"Slot {slot_idx+1}  {d.strftime('%a %d %b')}  {session}"
            row = f"{label:<26}"
            for grade in grades:
                subjects = sorted(grid[slot_idx][grade])
                row += f"{', '.join(subjects) if subjects else '':<{col_w}}"
            lines.append(row)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))

    # ─────────────────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────────────────

    def _browse_export_path(self):
        path = filedialog.asksaveasfilename(
            title="Save optimised timetable as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
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
            _write(w, f"WARNING: {len(student_clashes)} student clash(es), "
                      f"{len(teacher_clashes)} teacher clash(es) — "
                      f"exporting anyway.\n", "err")
        else:
            _write(w, "Verification passed ✓\n", "ok")
        _write(w, "\nExport not yet available.\n", "err")
        _write(w, "Waiting on:  timetable_converter.timetable_tree_to_block_tree()\n",
               "info")

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

    def _remove_exclusion(self):
        sel = self.excl_listbox.curselection()
        if not sel:
            return
        code = self.excl_listbox.get(sel[0])
        self.exclusions.discard(code)
        self._refresh_exclusion_listbox()

    # ─────────────────────────────────────────────────────────
    # STYLING
    # ─────────────────────────────────────────────────────────

    def _style_tree(self, tree: ttk.Treeview):
        style = ttk.Style()
        style.configure("Treeview",
                         font=("Courier", 9), rowheight=22,
                         background=CLR_WHITE, fieldbackground=CLR_WHITE)
        style.configure("Treeview.Heading",
                         font=("Helvetica", 10, "bold"))


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = TimePyBlingApp()
    app.mainloop()