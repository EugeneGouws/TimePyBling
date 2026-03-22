"""
ui.py — TimePyBling main interface

Tabs
----
  Timetable    — 8×7 rotation grid; click cell for class detail; view entity timetable popup
  Verification — clash report, data integrity, schedulable pairs
  Exams        — exam slot scheduling + dated exam timetable per grade
"""

import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import date, timedelta
from collections import defaultdict
import pandas as pd

from core.timetable_tree     import (build_timetable_tree_from_file,
                                      timetable_tree_to_dict,
                                      timetable_tree_from_dict)
from core.conflict_matrix    import ConflictMatrix
from reader.exam_tree        import build_exam_tree_from_timetable_tree
from reader.verify_timetable import _find_clashes
from reader.exam_paper       import ExamPaper, ExamPaperRegistry
from reader.exam_scheduler   import (build_schedule, SESSIONS, EXAM_WEEKDAYS)
from reader.exam_clash       import (build_clash_graph, dsatur_colouring,
                                     is_excluded as _is_subject_excluded)

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────

DEFAULT_EXCLUSIONS   = {"ST", "LIB", "PE", "RDI"}
TEACHER_SUBJECT_COLS = ["sua", "sub", "suc"]

# Rotation timetable grid: 8 days × 7 periods.
# Each entry is the subblock that falls in that (day, period) slot.
# Column index == period number; block letter rotates with H as anchor.
TIMETABLE_GRID = [
    ["A1","B2","C3","D4","E5","F6","G7"],  # Day 1
    ["G1","H2","B3","C4","D5","E6","F7"],  # Day 2
    ["F1","A2","H3","B4","C5","D6","E7"],  # Day 3
    ["E1","F2","A3","H4","B5","C6","D7"],  # Day 4
    ["D1","E2","F3","A4","H5","B6","C7"],  # Day 5
    ["C1","D2","E3","F4","A5","H6","B7"],  # Day 6
    ["B1","C2","D3","E4","F5","A6","H7"],  # Day 7
    ["H1","B2","C3","D4","E5","F6","A7"],  # Day 8
]

CLR_GRID_CELL   = "#f0f4f8"
CLR_GRID_HEADER = "#dde3ea"
CLR_GRID_ACTIVE = "#d0e8ff"

DEFAULT_EXAM_START = date(2026, 6, 1)
DEFAULT_EXAM_END   = date(2026, 6, 23)

CLR_HEADER   = "#2c3e50"
CLR_GREEN    = "#27ae60"
CLR_BLUE     = "#2980b9"
CLR_RED      = "#c0392b"
CLR_LIGHT    = "#ecf0f1"
CLR_MID      = "#bdc3c7"
CLR_WHITE    = "white"
CLR_BG       = "#f5f5f5"
CLR_MORNING  = "#e8f5e9"
CLR_AFTERNOON= "#fff8e1"


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
    return label.split("_")[0]


def _data_integrity_issues(tree) -> list[dict]:
    """Find class labels with fewer than 5 students across all subblocks."""
    class_info: dict[str, dict] = defaultdict(
        lambda: {"students": set(), "subblocks": set()})
    for block in tree.blocks.values():
        for sb_name, subblock in block.subblocks.items():
            for label, cl in subblock.class_lists.items():
                class_info[label]["students"] |= cl.student_list.students
                class_info[label]["subblocks"].add(sb_name)
    issues = []
    for label, info in sorted(class_info.items()):
        if len(info["students"]) < 5:
            issues.append({
                "label":    label,
                "count":    len(info["students"]),
                "subblocks": sorted(info["subblocks"],
                                    key=lambda x: (x[0], int(x[1:]))),
                "students": sorted(info["students"]),
            })
    return issues


# ─────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────

class TimePyBlingApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("TimePyBling")
        self.configure(bg=CLR_BG)
        self.minsize(900, 600)
        self.after(0, lambda: self.state("zoomed"))

        self.timetable_tree        = None
        self.exam_tree             = None
        self.paper_registry        = None
        self.schedule_result       = None
        self.st1_path              = None
        self.teachers_path         = None
        self.teacher_subj_map      = {}
        self.exclusions            = set(DEFAULT_EXCLUSIONS)
        self._selected_paper_label = None
        self._sessions             = None
        self.session_count_label   = None
        self._subblock_popup       = None

        # AM / PM session toggles (BUG 6)
        self._am_var = tk.BooleanVar(value=True)
        self._pm_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.after(300, self._auto_load)   # BUG 1
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ─────────────────────────────────────────────────────────
    # AUTO-LOAD  (BUG 1)
    # ─────────────────────────────────────────────────────────

    def _auto_load(self):
        # Prefer a pre-exported state file so the .xlsx is not needed at runtime
        state_candidates = [
            "data/timetable_state.json",
        ]
        for c in state_candidates:
            p = Path(c)
            if p.exists():
                self._load_state_json(p)
                return
        xlsx_candidates = [
            "data/ST1 2026.xlsx",
            "data/ST12026.xlsx",
            "data/ST1_2026.xlsx",
            "data/ST1.xlsx",
        ]
        for c in xlsx_candidates:
            p = Path(c)
            if p.exists():
                self._load_st1_path(p)
                return

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


    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_timetable    = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_verification = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_exams        = tk.Frame(self.notebook, bg=CLR_WHITE)

        self.notebook.add(self.tab_timetable,    text="  Timetable  ")
        self.notebook.add(self.tab_verification, text="  Verification  ")
        self.notebook.add(self.tab_exams,        text="  Exams  ")

        self._build_timetable_tab()
        self._build_verification_tab()
        self._build_exam_tab()

    # ─────────────────────────────────────────────────────────
    # TAB 1 — TIMETABLE
    # ─────────────────────────────────────────────────────────

    def _build_timetable_tab(self):
        pane = tk.PanedWindow(self.tab_timetable, orient=tk.HORIZONTAL,
                              bg="#ccc", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)

        # ══════════════════════════════════════════════════════════
        # LEFT — main 8×7 rotation grid
        # ══════════════════════════════════════════════════════════
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=280)

        canvas = tk.Canvas(left, bg=CLR_WHITE, highlightthickness=0)
        vsb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=canvas.yview)
        hsb = ttk.Scrollbar(left, orient=tk.HORIZONTAL, command=canvas.xview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._grid_frame = tk.Frame(canvas, bg=CLR_WHITE)
        canvas.create_window((0, 0), window=self._grid_frame, anchor="nw")
        self._grid_frame.bind("<Configure>",
                              lambda e: canvas.configure(
                                  scrollregion=canvas.bbox("all")))

        # Column headers
        tk.Label(self._grid_frame, text="", bg=CLR_GRID_HEADER,
                 width=3, relief=tk.RIDGE, bd=1
                 ).grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        for col in range(7):
            tk.Label(self._grid_frame, text=f"P{col+1}",
                     bg=CLR_GRID_HEADER, font=("Helvetica", 8, "bold"),
                     width=4, relief=tk.RIDGE, bd=1
                     ).grid(row=0, column=col+1, sticky="nsew", padx=0, pady=0)

        # Row headers + subblock buttons
        self._grid_cells = []
        for day in range(8):
            tk.Label(self._grid_frame, text=f"D{day+1}",
                     bg=CLR_GRID_HEADER, font=("Helvetica", 8, "bold"),
                     width=3, relief=tk.RIDGE, bd=1
                     ).grid(row=day+1, column=0, sticky="nsew", padx=0, pady=0)
            row_cells = []
            for col in range(7):
                subblock = TIMETABLE_GRID[day][col]
                btn = tk.Button(
                    self._grid_frame,
                    text=subblock,
                    font=("Helvetica", 8, "bold"),
                    bg=CLR_GRID_CELL,
                    relief=tk.RIDGE, bd=1,
                    width=4, pady=2, padx=0,
                    command=lambda sb=subblock: self._show_subblock_detail(sb),
                )
                btn.grid(row=day+1, column=col+1, sticky="nsew", padx=0, pady=0)
                row_cells.append(btn)
            self._grid_cells.append(row_cells)

        # ══════════════════════════════════════════════════════════
        # RIGHT — entity timetable viewer (embedded, not a popup)
        # ══════════════════════════════════════════════════════════
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=420)

        # ── selector bar ──────────────────────────────────────────
        sel = tk.Frame(right, bg=CLR_WHITE, padx=8, pady=6)
        sel.pack(fill=tk.X)

        tk.Label(sel, text="View:", bg=CLR_WHITE,
                 font=("Helvetica", 10)).pack(side=tk.LEFT)

        self._entity_type_var = tk.StringVar(value="Student")
        entity_type_cb = ttk.Combobox(sel, textvariable=self._entity_type_var,
                                      values=["Student", "Teacher", "Subject"],
                                      state="readonly", width=10,
                                      font=("Helvetica", 10))
        entity_type_cb.pack(side=tk.LEFT, padx=(4, 8))
        entity_type_cb.bind("<<ComboboxSelected>>", self._on_entity_type_change)

        self._entity_value_var = tk.StringVar()
        self._entity_search_entry = tk.Entry(sel,
                                             textvariable=self._entity_value_var,
                                             width=20, font=("Helvetica", 10),
                                             relief=tk.SOLID, bd=1)
        self._entity_search_entry.pack(side=tk.LEFT, padx=(0, 6))
        self._entity_value_var.trace_add("write", self._on_entity_search_change)

        tk.Button(sel, text="View Timetable",
                  command=self._on_view_timetable,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10
                  ).pack(side=tk.LEFT)

        # ── options list (updates while typing) ───────────────────
        list_frame = tk.Frame(right, bg=CLR_WHITE)
        list_frame.pack(fill=tk.X, padx=8, pady=(0, 4))

        self._entity_listbox = tk.Listbox(list_frame, height=5,
                                           font=("Courier", 9),
                                           relief=tk.SOLID, bd=1,
                                           selectmode=tk.SINGLE)
        list_sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                command=self._entity_listbox.yview)
        list_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._entity_listbox.config(yscrollcommand=list_sb.set)
        self._entity_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._entity_listbox.bind("<<ListboxSelect>>",
                                  self._on_entity_listbox_select)

        # ── entity heading ────────────────────────────────────────
        self._entity_heading_var = tk.StringVar()
        tk.Label(right, textvariable=self._entity_heading_var,
                 bg=CLR_WHITE, font=("Helvetica", 11, "bold")
                 ).pack(padx=8, pady=(4, 2))

        # ── entity timetable grid (embedded) ─────────────────────
        entity_canvas = tk.Canvas(right, bg=CLR_WHITE, highlightthickness=0)
        evsb = ttk.Scrollbar(right, orient=tk.VERTICAL,
                              command=entity_canvas.yview)
        evsb.pack(side=tk.RIGHT, fill=tk.Y)
        entity_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                           padx=8, pady=(0, 8))
        entity_canvas.configure(yscrollcommand=evsb.set)

        self._entity_grid_frame = tk.Frame(entity_canvas, bg=CLR_WHITE)
        entity_canvas.create_window((0, 0), window=self._entity_grid_frame,
                                    anchor="nw")
        self._entity_grid_frame.bind(
            "<Configure>",
            lambda e: entity_canvas.configure(
                scrollregion=entity_canvas.bbox("all")))

    # ─────────────────────────────────────────────────────────
    # TAB 2 — VERIFICATION  (BUG 2 + BUG 3)
    # ─────────────────────────────────────────────────────────

    def _build_verification_tab(self):
        pane = tk.PanedWindow(self.tab_verification, orient=tk.HORIZONTAL,
                              bg="#ccc", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)

        # Left: Clashes + Schedulable Pairs
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=520)

        hdr = tk.Frame(left, bg=CLR_WHITE)
        hdr.pack(fill=tk.X, padx=8, pady=(8, 2))
        tk.Label(hdr, text="Clashes  &  Schedulable Pairs", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(hdr, text="Re-run", command=self._run_verification,
                  bg=CLR_LIGHT, font=("Helvetica", 8),
                  relief=tk.FLAT, padx=8).pack(side=tk.RIGHT)

        report_frame = tk.Frame(left, bg=CLR_WHITE)
        report_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.clash_report = _scrolled_text(report_frame)
        self.clash_report.tag_config("pass",    foreground=CLR_GREEN)
        self.clash_report.tag_config("fail",    foreground=CLR_RED)
        self.clash_report.tag_config("heading", foreground="#333",
                                     font=("Courier", 9, "bold"))
        self.clash_report.tag_config("warn",    foreground="#e67e22")
        self.clash_report.tag_config("dim",     foreground="#888")

        # Right: Data Integrity
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=300)

        di_hdr = tk.Frame(right, bg=CLR_WHITE)
        di_hdr.pack(fill=tk.X, padx=8, pady=(8, 2))
        tk.Label(di_hdr, text="Data Integrity", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)

        di_frame = tk.Frame(right, bg=CLR_WHITE)
        di_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.integrity_report = _scrolled_text(di_frame)
        self.integrity_report.tag_config("pass",    foreground=CLR_GREEN)
        self.integrity_report.tag_config("heading", foreground="#333",
                                          font=("Courier", 9, "bold"))
        self.integrity_report.tag_config("warn",    foreground="#e67e22")
        self.integrity_report.tag_config("dim",     foreground="#888")

    # ─────────────────────────────────────────────────────────
    # TAB 3 — EXAMS  (BUG 4 + 5 + 6)
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

        tk.Button(left_top, text="Save State…", command=self._export_exam_state,
                  bg="#666", fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6).pack(side=tk.RIGHT, padx=(2, 0))
        tk.Button(left_top, text="Load State…", command=self._import_exam_state,
                  bg="#666", fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8), padx=6).pack(side=tk.RIGHT, padx=2)
        tk.Button(left_top, text="Rebuild", command=self._rebuild_exam,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8, "bold"), padx=8
                  ).pack(side=tk.RIGHT, padx=2)

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

        # ── Right: exclusions + scheduler controls + cost function ──
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
        sched_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        ctrl = tk.Frame(sched_lf, bg=CLR_WHITE)
        ctrl.pack(fill=tk.X, pady=(0, 4))

        # Row 0: Start / End date entries + live slot count
        tk.Label(ctrl, text="Start:", bg=CLR_WHITE,
                 font=("Helvetica", 9)).grid(row=0, column=0, sticky=tk.W)
        self.sched_start_var = tk.StringVar(
            value=DEFAULT_EXAM_START.strftime("%Y-%m-%d"))
        self._start_entry = tk.Entry(ctrl, textvariable=self.sched_start_var,
                                      font=("Helvetica", 9), relief=tk.SOLID,
                                      bd=1, width=11)
        self._start_entry.grid(row=0, column=1, sticky=tk.W, padx=(2, 8))
        self._start_entry.bind("<FocusOut>", lambda e: self._on_session_param_changed())
        self._start_entry.bind("<Return>",   lambda e: self._on_session_param_changed())

        tk.Label(ctrl, text="End:", bg=CLR_WHITE,
                 font=("Helvetica", 9)).grid(row=0, column=2, sticky=tk.W)
        self.sched_end_var = tk.StringVar(
            value=DEFAULT_EXAM_END.strftime("%Y-%m-%d"))
        self._end_entry = tk.Entry(ctrl, textvariable=self.sched_end_var,
                                    font=("Helvetica", 9), relief=tk.SOLID,
                                    bd=1, width=11)
        self._end_entry.grid(row=0, column=3, sticky=tk.W, padx=(2, 8))
        self._end_entry.bind("<FocusOut>", lambda e: self._on_session_param_changed())
        self._end_entry.bind("<Return>",   lambda e: self._on_session_param_changed())

        self.session_count_label = tk.Label(ctrl, text="", bg=CLR_WHITE,
                                             fg=CLR_BLUE,
                                             font=("Helvetica", 9, "bold"))
        self.session_count_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 4))

        # Row 1: AM / PM checkboxes + Configure sessions button  (BUG 6)
        tk.Checkbutton(ctrl, text="AM", variable=self._am_var,
                       bg=CLR_WHITE, font=("Helvetica", 9),
                       command=self._on_session_param_changed
                       ).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        tk.Checkbutton(ctrl, text="PM", variable=self._pm_var,
                       bg=CLR_WHITE, font=("Helvetica", 9),
                       command=self._on_session_param_changed
                       ).grid(row=1, column=1, sticky=tk.W, pady=(4, 0))
        tk.Button(ctrl, text="Configure sessions…",
                  command=self._open_session_calendar,
                  bg=CLR_LIGHT, font=("Helvetica", 8),
                  relief=tk.FLAT, padx=8
                  ).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=(4, 0))

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

        # Per-grade slot summary (BUG 6)
        self.slot_summary_text = tk.Text(sched_lf, height=5,
                                          font=("Courier", 8),
                                          state=tk.DISABLED, relief=tk.FLAT,
                                          bg=CLR_LIGHT, wrap=tk.NONE)
        self.slot_summary_text.pack(fill=tk.X, padx=2, pady=(0, 4))
        self.slot_summary_text.tag_config("ok",    foreground=CLR_GREEN)
        self.slot_summary_text.tag_config("short", foreground=CLR_RED)
        self.slot_summary_text.tag_config("dim",   foreground="#888")

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

        # ── Exam Cost Function  (BUG 3 — exam tab)  ──
        cost_lf = tk.LabelFrame(right, text="Exam Cost Function",
                                 bg=CLR_WHITE, font=("Helvetica", 9, "bold"),
                                 padx=8, pady=6)
        cost_lf.pack(fill=tk.X, padx=8, pady=(0, 8))

        wf = tk.Frame(cost_lf, bg=CLR_WHITE)
        wf.pack(fill=tk.X)
        weight_defs = [("W_day", 5), ("W_week", 1),
                       ("W_consec", 50), ("W_marking", 20)]
        self._cost_weight_vars: list[tk.StringVar] = []
        for i, (lbl, dflt) in enumerate(weight_defs):
            r, c = i // 2, (i % 2) * 3
            tk.Label(wf, text=f"{lbl}:", bg=CLR_WHITE,
                     font=("Helvetica", 8)).grid(row=r, column=c, sticky=tk.W,
                                                  padx=(0, 2), pady=2)
            v = tk.StringVar(value=str(dflt))
            self._cost_weight_vars.append(v)
            tk.Entry(wf, textvariable=v, font=("Helvetica", 8),
                     relief=tk.SOLID, bd=1, width=5
                     ).grid(row=r, column=c + 1, sticky=tk.W, padx=(0, 12), pady=2)

        calc_row = tk.Frame(cost_lf, bg=CLR_WHITE)
        calc_row.pack(fill=tk.X, pady=(4, 0))
        tk.Button(calc_row, text="Calculate",
                  command=self._calculate_exam_cost,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 8, "bold"), padx=8
                  ).pack(side=tk.LEFT)
        self.exam_cost_result_label = tk.Label(calc_row, text="",
                                                bg=CLR_WHITE,
                                                font=("Courier", 8),
                                                justify=tk.LEFT)
        self.exam_cost_result_label.pack(side=tk.LEFT, padx=6)

    # ─────────────────────────────────────────────────────────
    # LOAD
    # ─────────────────────────────────────────────────────────

    def _load_st1_path(self, path: Path):
        """Internal load — used by auto-load and the Load button."""
        self.st1_path = path
        self.st1_label.config(text=path.name, fg=CLR_WHITE)
        try:
            self.timetable_tree = build_timetable_tree_from_file(path)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        self._populate_timetable_grid()
        self._rebuild_exam()
        self._run_verification()
        self.notebook.select(self.tab_verification)

    def _load_st1(self):
        path = filedialog.askopenfilename(
            title="Select student timetable (ST1.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")])
        if not path:
            return
        self._load_st1_path(Path(path))

    def _load_teachers(self):
        path = filedialog.askopenfilename(
            title="Select teachers file (teachers.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files",   "*.*")])
        if not path:
            return
        self.teachers_path = Path(path)
        try:
            self.teacher_subj_map = _load_teacher_subject_map(self.teachers_path)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        if self.timetable_tree:
            self._run_verification()

    # ─────────────────────────────────────────────────────────
    # VERIFICATION  (BUG 2)
    # ─────────────────────────────────────────────────────────

    def _run_verification(self):
        if not self.timetable_tree:
            return
        self._update_clash_report()
        self._update_integrity_panel()

    def _update_clash_report(self):
        """Left panel: Clashes + Schedulable Pairs."""
        w = self.clash_report
        _clear(w)

        # ── CLASHES ──
        _write(w, "━" * 60 + "\n", "dim")
        _write(w, "CLASHES\n", "heading")
        _write(w, "━" * 60 + "\n", "dim")

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

        # ── SCHEDULABLE PAIRS ──
        _write(w, "\n" + "━" * 60 + "\n", "dim")
        _write(w, "SCHEDULABLE PAIRS  (share no students — can sit same slot)\n",
               "heading")
        _write(w, "━" * 60 + "\n", "dim")

        if not self.exam_tree:
            _write(w, "(Load timetable to see schedulable pairs)\n", "dim")
        else:
            any_pairs = False
            for grade_label in sorted(self.exam_tree.grades.keys()):
                grade_node = self.exam_tree.grades[grade_label]
                groups = {
                    subj_label: subject.all_students()
                    for subj_label, subject in grade_node.exam_subjects.items()
                    if not _is_subject_excluded(subj_label, self.exclusions)
                }
                if not groups:
                    continue
                cm = ConflictMatrix(grade_label, groups)
                pairs = cm.free_pairs()
                if pairs:
                    any_pairs = True
                    _write(w, f"\n  {grade_label}  ({len(pairs)} pair(s)):\n",
                           "heading")
                    for a, b in pairs:
                        _write(w, f"    {a}  +  {b}\n", "pass")
            if not any_pairs:
                _write(w, "  No free pairs — all subjects share students.\n", "dim")

    def _update_integrity_panel(self):
        """Right panel: Data Integrity — classes with fewer than 5 students."""
        w = self.integrity_report
        _clear(w)
        _write(w, "Classes with fewer than 5 students:\n", "dim")
        _write(w, "─" * 44 + "\n", "dim")
        issues = _data_integrity_issues(self.timetable_tree)
        if not issues:
            _write(w, "PASS ✓  —  all classes have 5 or more students\n", "pass")
        else:
            _write(w, f"WARN ⚠  —  {len(issues)} class(es) flagged:\n", "warn")
            for info in issues:
                _write(w, f"\n  {info['label']}\n", "heading")
                _write(w, f"    Count:     {info['count']}\n", "warn")
                _write(w, f"    Subblocks: {info['subblocks']}\n", "dim")
                _write(w, f"    Students:  {info['students']}\n", "dim")

    # ─────────────────────────────────────────────────────────
    # TIMETABLE GRID
    # ─────────────────────────────────────────────────────────

    def _populate_timetable_grid(self):
        """Refresh grid cell colours; no content change — buttons always show subblock name."""
        for day in range(8):
            for col in range(7):
                self._grid_cells[day][col].config(bg=CLR_GRID_CELL)
        self._refresh_entity_listbox()

    # ─────────────────────────────────────────────────────────
    # TIMETABLE GRID — DATA HELPERS & ENTITY SELECTOR
    # ─────────────────────────────────────────────────────────

    def _all_students(self) -> list:
        students = set()
        for block in self.timetable_tree.blocks.values():
            for subblock in block.subblocks.values():
                for cl in subblock.class_lists.values():
                    students |= cl.student_list.students
        return sorted(students)

    def _all_teachers(self) -> list:
        teachers = set()
        for block in self.timetable_tree.blocks.values():
            for subblock in block.subblocks.values():
                for label in subblock.class_lists:
                    parts = label.split("_")
                    if len(parts) >= 3:
                        teachers.add("_".join(parts[1:-1]))
        return sorted(teachers)

    def _all_subjects(self) -> list:
        subjects = set()
        for block in self.timetable_tree.blocks.values():
            for subblock in block.subblocks.values():
                for label in subblock.class_lists:
                    subjects.add(label.split("_")[0])
        return sorted(subjects)

    def _on_entity_type_change(self, *_):
        self._entity_value_var.set("")
        self._refresh_entity_listbox()

    def _on_entity_search_change(self, *_):
        self._refresh_entity_listbox()

    def _entity_full_list(self) -> list:
        if not self.timetable_tree:
            return []
        etype = self._entity_type_var.get()
        if etype == "Student":
            return [str(s) for s in self._all_students()]
        elif etype == "Teacher":
            return self._all_teachers()
        return self._all_subjects()

    def _refresh_entity_listbox(self):
        """Update the options listbox based on typed text; focus stays in entry."""
        typed = self._entity_value_var.get().strip().lower()
        full  = self._entity_full_list()
        filtered = [v for v in full if typed in v.lower()] if typed else full
        self._entity_listbox.delete(0, tk.END)
        for v in filtered:
            self._entity_listbox.insert(tk.END, v)

    def _on_entity_listbox_select(self, *_):
        sel = self._entity_listbox.curselection()
        if sel:
            value = self._entity_listbox.get(sel[0])
            self._entity_value_var.set(value)
            self._refresh_entity_grid(self._entity_type_var.get(), value)

    def _on_view_timetable(self):
        etype = self._entity_type_var.get()
        value = self._entity_value_var.get().strip()
        if not value or not self.timetable_tree:
            return
        self._refresh_entity_grid(etype, value)

    def _show_subblock_detail(self, subblock_name: str):
        """Open a small popup listing all classes in a subblock (one at a time)."""
        if not self.timetable_tree:
            return
        if self._subblock_popup and self._subblock_popup.winfo_exists():
            self._subblock_popup.destroy()

        block_letter = subblock_name[0]
        block = self.timetable_tree.blocks.get(block_letter)
        if not block:
            return
        sb = block.subblocks.get(subblock_name)

        win = tk.Toplevel(self)
        self._subblock_popup = win
        win.title(f"Subblock {subblock_name}")
        win.configure(bg=CLR_WHITE)
        win.resizable(True, True)

        tk.Label(win, text=f"Classes in {subblock_name}",
                 font=("Helvetica", 11, "bold"),
                 bg=CLR_WHITE).pack(padx=14, pady=(10, 4))

        frame = tk.Frame(win, bg=CLR_WHITE)
        frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 4))

        if not sb or not sb.class_lists:
            tk.Label(frame, text="(no classes)", fg="#888",
                     bg=CLR_WHITE, font=("Helvetica", 9)).pack()
        else:
            for label in sorted(sb.class_lists):
                cl = sb.class_lists[label]
                count = len(cl.student_list)
                tk.Label(frame, text=f"  {label}   ({count} students)",
                         bg=CLR_WHITE, font=("Courier", 9),
                         anchor="w").pack(fill=tk.X)

        tk.Button(win, text="Close", command=win.destroy,
                  bg=CLR_LIGHT, relief=tk.FLAT,
                  font=("Helvetica", 9), padx=12
                  ).pack(pady=(0, 10))

    def _refresh_entity_grid(self, entity_type: str, value: str):
        """Rebuild the embedded entity timetable grid on the right panel."""
        for w in self._entity_grid_frame.winfo_children():
            w.destroy()

        self._entity_heading_var.set(f"{entity_type}: {value}")

        gf = self._entity_grid_frame

        # Column headers
        tk.Label(gf, text="", bg=CLR_GRID_HEADER,
                 width=8, relief=tk.RIDGE, bd=1
                 ).grid(row=0, column=0, padx=1, pady=1, sticky="nsew")
        for col in range(7):
            tk.Label(gf, text=f"P{col+1}",
                     bg=CLR_GRID_HEADER, font=("Helvetica", 9, "bold"),
                     width=18, relief=tk.RIDGE, bd=1
                     ).grid(row=0, column=col+1, padx=1, pady=1, sticky="nsew")

        for day in range(8):
            tk.Label(gf, text=f"Day {day+1}",
                     bg=CLR_GRID_HEADER, font=("Helvetica", 9, "bold"),
                     width=8, relief=tk.RIDGE, bd=1
                     ).grid(row=day+1, column=0, padx=1, pady=1, sticky="nsew")
            for col in range(7):
                subblock_name = TIMETABLE_GRID[day][col]
                block_letter  = subblock_name[0]
                block = self.timetable_tree.blocks.get(block_letter)
                matching = []
                if block:
                    sb = block.subblocks.get(subblock_name)
                    if sb:
                        for label, cl in sb.class_lists.items():
                            parts = label.split("_")
                            if entity_type == "Student":
                                try:
                                    sid = int(value)
                                except ValueError:
                                    sid = None
                                if sid is not None and sid in cl.student_list.students:
                                    matching.append(label)
                            elif entity_type == "Teacher":
                                teacher = "_".join(parts[1:-1]) if len(parts) >= 3 else ""
                                if teacher == value:
                                    matching.append(label)
                            else:  # Subject
                                if parts[0] == value:
                                    matching.append(label)

                def _fmt(lbl):
                    p = lbl.split("_")
                    subj  = p[0]
                    grade = p[-1]
                    tchr  = "_".join(p[1:-1]) if len(p) >= 3 else ""
                    if entity_type == "Teacher":
                        return f"{subj}  Gr{grade}"
                    elif entity_type == "Student":
                        return f"{subj}  {tchr}"
                    else:  # Subject
                        return f"{subj}  {tchr}  Gr{grade}"

                cell_text = "\n".join(_fmt(lbl) for lbl in matching) if matching else ""
                cell_bg   = CLR_GRID_ACTIVE if matching else CLR_GRID_CELL
                tk.Label(gf,
                         text=cell_text,
                         bg=cell_bg,
                         font=("Courier", 8),
                         width=18, wraplength=140,
                         justify=tk.CENTER,
                         relief=tk.RIDGE, bd=1, pady=4
                         ).grid(row=day+1, column=col+1, padx=1, pady=1, sticky="nsew")

    # ─────────────────────────────────────────────────────────
    # EXAM TREE  (BUG 4)
    # ─────────────────────────────────────────────────────────

    def _populate_exam_tree(self):
        """Populate from exam_tree; excluded subjects shown with [excl] suffix."""
        self.ex_tree.delete(*self.ex_tree.get_children())
        if not self.exam_tree:
            return
        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node = self.exam_tree.grades[grade_label]
            grade_ui = self.ex_tree.insert("", tk.END, text=grade_label, open=False)
            for subj_label in sorted(grade_node.exam_subjects.keys()):
                exam_subject = grade_node.exam_subjects[subj_label]
                excl = _is_subject_excluded(subj_label, self.exclusions)
                if excl:
                    self.ex_tree.insert(grade_ui, tk.END,
                                        text=f"{subj_label}  [excl]",
                                        tags=("excluded",),
                                        values=())
                else:
                    subj_code = subj_label.split("_")[0]
                    if self.paper_registry:
                        papers = self.paper_registry.papers_for_subject_grade(
                            subj_code, grade_label)
                        count = (max(p.student_count() for p in papers)
                                 if papers else len(exam_subject.all_students()))
                    else:
                        count = len(exam_subject.all_students())
                    self.ex_tree.insert(grade_ui, tk.END,
                                        text=f"{subj_code}  ({count} students)",
                                        tags=("subject",),
                                        values=(subj_code, grade_label))

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
            raw = self.ex_tree.item(item, "text")
            if raw in expanded:
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

    def _navigate_to_exam_subject(self, subject: str, grade: str, popout=None):
        if popout:
            popout.destroy()
        self.notebook.select(self.tab_exams)
        for grade_item in self.ex_tree.get_children():
            if self.ex_tree.item(grade_item, "text") != grade:
                continue
            self.ex_tree.item(grade_item, open=True)
            for subj_item in self.ex_tree.get_children(grade_item):
                tags = self.ex_tree.item(subj_item, "tags")
                if "subject" not in (tags or []):
                    continue
                vals = self.ex_tree.item(subj_item, "values")
                if vals and vals[0] == subject and vals[1] == grade:
                    self.ex_tree.selection_set(subj_item)
                    self.ex_tree.see(subj_item)
                    self._on_exam_tree_select()
                    return

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
        if not self.paper_registry:
            return
        tree_sel = self.ex_tree.selection()
        if not tree_sel:
            return
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
        self._update_session_count_label()

    def _update_sched_grade_list(self):
        if not self.paper_registry:
            return
        grades = self.paper_registry.grades()
        options = ["All Grades"] + grades
        self.sched_grade_cb["values"] = options
        if grades:
            self.sched_grade_var.set(grades[-1])

    # ─────────────────────────────────────────────────────────
    # EXCLUSION LIST  (BUG 4 — no _rebuild_exam on change)
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
        self._on_exclusion_change()

    def _remove_exclusion(self):
        sel = self.excl_listbox.curselection()
        if not sel:
            return
        code = self.excl_listbox.get(sel[0])
        self.exclusions.discard(code)
        self._on_exclusion_change()

    def _on_exclusion_change(self):
        expanded, selected_pairs = self._exam_tree_get_state()
        self._refresh_exclusion_listbox()
        self._rebuild_exam()
        self._populate_exam_tree()
        self._exam_tree_restore_state(expanded, selected_pairs)
        self._update_session_count_label()

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
        self._on_exam_tree_select()

    def _remove_paper(self):
        if not self.paper_registry or not self._selected_paper_label:
            return
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
    # IMPORT / EXPORT EXAM STATE  (BUG 5)
    # ─────────────────────────────────────────────────────────

    def _build_state_dict(self) -> dict:
        papers_data: dict[str, list[str]] = {}
        if self.paper_registry:
            for grade in self.paper_registry.grades():
                grade_num = grade.replace("Gr", "")
                for subj in self.paper_registry.subjects_for_grade(grade):
                    ps = self.paper_registry.papers_for_subject_grade(subj, grade)
                    papers_data[f"{subj}_{grade_num}"] = [
                        f"P{p.paper_number}" for p in ps]
        return {
            "timetable_tree": (timetable_tree_to_dict(self.timetable_tree)
                               if self.timetable_tree else None),
            "exclusions": sorted(self.exclusions),
            "papers":     papers_data,
            "session": {
                "start": self.sched_start_var.get(),
                "end":   self.sched_end_var.get(),
                "am":    self._am_var.get(),
                "pm":    self._pm_var.get(),
            },
        }

    def _save_state_json(self, path) -> bool:
        """Write current state to path. Returns True on success."""
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(self._build_state_dict(), fh, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Save Error", str(e))
            return False

    def _export_exam_state(self):
        path = filedialog.asksaveasfilename(
            title="Save timetable state",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        if self._save_state_json(path):
            messagebox.showinfo("Saved", f"State saved to:\n{path}")

    def _on_close(self):
        """Prompt to save before exit; always saves to the default auto-load path."""
        if not self.timetable_tree:
            self.destroy()
            return

        DEFAULT_PATH = Path("data/timetable_state.json")

        win = tk.Toplevel(self)
        win.title("Save before exit?")
        win.configure(bg=CLR_WHITE)
        win.resizable(False, False)
        win.grab_set()
        win.focus_force()

        tk.Label(win,
                 text="Save current timetable and exam state\nbefore closing?",
                 bg=CLR_WHITE, font=("Helvetica", 10),
                 justify=tk.CENTER).pack(padx=24, pady=(18, 12))

        btn_row = tk.Frame(win, bg=CLR_WHITE)
        btn_row.pack(padx=24, pady=(0, 18))

        def save_and_exit():
            win.destroy()
            DEFAULT_PATH.parent.mkdir(parents=True, exist_ok=True)
            if self._save_state_json(DEFAULT_PATH):
                self.destroy()

        def exit_no_save():
            win.destroy()
            self.destroy()

        def cancel():
            win.destroy()

        tk.Button(btn_row, text="Save & Exit",
                  command=save_and_exit,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=12
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_row, text="Exit without saving",
                  command=exit_no_save,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9), padx=12
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_row, text="Cancel",
                  command=cancel,
                  bg=CLR_LIGHT, relief=tk.FLAT,
                  font=("Helvetica", 9), padx=12
                  ).pack(side=tk.LEFT, padx=4)

    def _import_exam_state(self):
        path = filedialog.askopenfilename(
            title="Load timetable state",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        self._apply_state(data, source_label=Path(path).name)
        messagebox.showinfo("Loaded", f"State loaded from:\n{path}")

    def _load_state_json(self, path: Path):
        """Load a saved state JSON without showing a dialog (used by auto-load)."""
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as e:
            messagebox.showerror("Auto-load Error", str(e))
            return
        self._apply_state(data, source_label=path.name)

    def _apply_state(self, data: dict, source_label: str = "state"):
        """Restore full app state from a previously saved dict."""
        # ── Timetable tree ──────────────────────────────────────
        if data.get("timetable_tree"):
            try:
                self.timetable_tree = timetable_tree_from_dict(
                    data["timetable_tree"])
                self.st1_path = None
                self.st1_label.config(text=source_label, fg=CLR_WHITE)
                self._populate_timetable_grid()
            except Exception as e:
                messagebox.showerror("Load Error",
                                     f"Could not restore timetable:\n{e}")
                return

        # ── Exclusions ──────────────────────────────────────────
        if "exclusions" in data:
            self.exclusions = set(data["exclusions"])
            self._refresh_exclusion_listbox()

        # ── Session settings ────────────────────────────────────
        if "session" in data:
            s = data["session"]
            if "start" in s:
                self.sched_start_var.set(s["start"])
            if "end" in s:
                self.sched_end_var.set(s["end"])
            if "am" in s:
                self._am_var.set(bool(s["am"]))
            if "pm" in s:
                self._pm_var.set(bool(s["pm"]))
            self._sessions = None

        # ── Rebuild paper registry from (restored) timetable ────
        if self.timetable_tree:
            self._rebuild_exam()

        # ── Restore extra papers (P2, P3) ───────────────────────
        if "papers" in data and self.paper_registry:
            for subj_grade, paper_nums in data["papers"].items():
                parts = subj_grade.split("_")
                if len(parts) != 2:
                    continue
                subj, grade_num = parts
                grade = f"Gr{grade_num}"
                max_num = max(
                    (int(p.replace("P", "")) for p in paper_nums
                     if p.startswith("P") and p[1:].isdigit()),
                    default=1,
                )
                for _ in range(max_num - 1):
                    self.paper_registry.add_paper(subj, grade)

        self._populate_exam_tree()
        self._run_verification()
        self._update_session_count_label()

    # ─────────────────────────────────────────────────────────
    # EXAM COST FUNCTION  (BUG 3 — exam tab)
    # ─────────────────────────────────────────────────────────

    def _calculate_exam_cost(self):
        if not self.schedule_result:
            self.exam_cost_result_label.config(text="Generate schedule first.")
            return
        try:
            w_day     = float(self._cost_weight_vars[0].get())
            w_week    = float(self._cost_weight_vars[1].get())
            w_consec  = float(self._cost_weight_vars[2].get())
            w_marking = float(self._cost_weight_vars[3].get())
        except ValueError:
            self.exam_cost_result_label.config(text="Invalid weight values.")
            return

        # Student clustering cost already computed by scheduler
        t_student = self.schedule_result.student_cost

        # Consecutive same-subject papers (adjacent slot indices)
        consec_count = 0
        by_sg: dict[tuple[str, str], list] = defaultdict(list)
        for sp in self.schedule_result.scheduled:
            by_sg[(sp.paper.subject, sp.paper.grade)].append(sp.slot_index)
        for slots_list in by_sg.values():
            sorted_s = sorted(slots_list)
            for i in range(len(sorted_s) - 1):
                if sorted_s[i + 1] - sorted_s[i] == 1:
                    consec_count += 1
        t_consec = w_consec * consec_count

        marking_count = len(self.schedule_result.teacher_warnings)
        t_marking = w_marking * marking_count

        total = t_student + t_consec + t_marking
        self.exam_cost_result_label.config(
            text=(f"E = {total:.0f}\n"
                  f"  student (W_day={w_day:.0f}, W_week={w_week:.0f}) = {t_student}\n"
                  f"  consec {w_consec:.0f}×{consec_count} = {t_consec:.0f}\n"
                  f"  marking {w_marking:.0f}×{marking_count} = {t_marking:.0f}")
        )

    # ─────────────────────────────────────────────────────────
    # EXAM SCHEDULE GENERATOR
    # ─────────────────────────────────────────────────────────

    def _generate_exam_schedule(self):
        if not self.paper_registry:
            messagebox.showinfo("No data", "Load a timetable first.")
            return
        sessions = self._get_effective_sessions()
        if sessions is None:
            try:
                start = date.fromisoformat(self.sched_start_var.get().strip())
                end   = date.fromisoformat(self.sched_end_var.get().strip())
                msg = "End date must be on or after start date."
            except ValueError:
                msg = "Enter valid start and end dates (YYYY-MM-DD)."
            messagebox.showerror("Invalid dates", msg)
            return
        if not sessions:
            messagebox.showerror("No sessions",
                                  "No exam sessions in the selected date range.\n"
                                  "Check AM/PM checkboxes and date range.")
            return
        try:
            w_day  = int(float(self._cost_weight_vars[0].get()))
            w_week = int(float(self._cost_weight_vars[1].get()))
        except (ValueError, AttributeError):
            w_day, w_week = 5, 1
        self.schedule_result = build_schedule(
            registry  = self.paper_registry,
            sessions  = sessions,
            exam_tree = self.exam_tree,
            w_day     = w_day,
            w_week    = w_week,
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

        all_warnings = [w_msg for sp in items for w_msg in sp.warnings]
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
    # SESSION DATES + SLOT COUNT  (BUG 6)
    # ─────────────────────────────────────────────────────────

    def _get_effective_sessions(self) -> list[tuple[date, str]] | None:
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
        am = self._am_var.get()
        pm = self._pm_var.get()
        sessions: list[tuple[date, str]] = []
        for day in days:
            if am:
                sessions.append((day, "AM"))
            if pm:
                sessions.append((day, "PM"))
        return sessions

    def _on_session_param_changed(self):
        """Date entry or AM/PM checkbox changed — reset any custom session list."""
        self._sessions = None
        self._update_session_count_label()

    def _update_session_count_label(self):
        if self.session_count_label is None:
            return
        sessions = self._get_effective_sessions()
        if sessions is None:
            self.session_count_label.config(text="invalid dates", fg=CLR_RED)
            self._update_slot_summary(None)
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
        self._update_slot_summary(n_slots)

    def _needed_slots_per_grade(self) -> dict[str, int]:
        """Run DSatur per grade on the current paper_registry."""
        if not self.paper_registry:
            return {}
        result: dict[str, int] = {}
        for grade in self.paper_registry.grades():
            papers = self.paper_registry.papers_for_grade(grade)
            if not papers:
                continue
            student_sets = {p.label: p.student_ids for p in papers}
            graph = build_clash_graph(student_sets)
            coloring = dsatur_colouring(graph)
            result[grade] = (max(coloring.values()) + 1) if coloring else 0
        return result

    def _update_slot_summary(self, available: int | None):
        if not hasattr(self, "slot_summary_text"):
            return
        w = self.slot_summary_text
        w.config(state=tk.NORMAL)
        w.delete("1.0", tk.END)
        if available is None or not self.paper_registry:
            w.config(state=tk.DISABLED)
            return
        needed = self._needed_slots_per_grade()
        if not needed:
            w.insert(tk.END, "  (no papers — load timetable first)\n", "dim")
        for grade in sorted(needed.keys()):
            n     = needed[grade]
            spare = available - n
            if spare >= 0:
                line = f"  {grade}: {n} needed / {available} available  ✓ {spare} spare\n"
                tag  = "ok"
            else:
                line = f"  {grade}: {n} needed / {available} available  ✗ SHORT by {-spare}\n"
                tag  = "short"
            w.insert(tk.END, line, tag)
        w.config(state=tk.DISABLED)

    def _open_session_calendar(self):
        sessions = self._get_effective_sessions()
        if not sessions:
            messagebox.showerror("Invalid dates",
                                  "Enter valid start and end dates first.")
            return

        day_state: dict[date, dict[str, bool]] = {}
        for d, s in sessions:
            day_state.setdefault(d, {"AM": False, "PM": False})
            day_state[d][s] = True
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
    # SCHEDULE POPOUT
    # ─────────────────────────────────────────────────────────

    def _open_schedule_popout(self):
        if not self.schedule_result:
            messagebox.showinfo("No schedule", "Generate a schedule first.")
            return

        result = self.schedule_result
        grades = sorted(
            {sp.paper.grade for sp in result.scheduled},
            key=lambda g: int(g[2:]) if g[2:].isdigit() else 0
        )

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

        COL_W = 10
        ROW_W = 22
        HDR_BG = "#d5d8dc"

        tk.Label(gf, text="Slot / Date / Sess", font=("Helvetica", 8, "bold"),
                 bg=HDR_BG, width=ROW_W, anchor="w",
                 relief=tk.RIDGE, bd=1).grid(row=0, column=0, sticky="nsew",
                                              padx=1, pady=1)
        for ci, grade in enumerate(grades):
            tk.Label(gf, text=grade, font=("Helvetica", 8, "bold"),
                     bg=HDR_BG, width=COL_W, anchor="center",
                     relief=tk.RIDGE, bd=1).grid(row=0, column=ci + 1,
                                                  sticky="nsew", padx=1, pady=1)

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
                if not subjects:
                    tk.Label(gf, text="", font=("Courier", 8),
                             bg=row_bg, width=COL_W, anchor="center",
                             relief=tk.RIDGE, bd=1).grid(
                                 row=ri + 1, column=ci + 1,
                                 sticky="nsew", padx=1, pady=1)
                else:
                    cell_frame = tk.Frame(gf, bg=row_bg, relief=tk.RIDGE, bd=1)
                    cell_frame.grid(row=ri + 1, column=ci + 1,
                                    sticky="nsew", padx=1, pady=1)
                    for subj in subjects:
                        lbl = tk.Label(cell_frame, text=subj, font=("Courier", 8),
                                       bg=row_bg, width=COL_W, anchor="center",
                                       cursor="hand2")
                        lbl.pack(fill=tk.X)
                        lbl.bind("<Button-1>",
                                 lambda e, s=subj, g=grade, p=top:
                                 self._navigate_to_exam_subject(s, g, p))

        gf.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _save_schedule_export(self, top, grades, grid, slot_meta, all_slots):
        try:
            import reportlab  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "PDF export unavailable",
                "reportlab is not installed.\n\nRun:  pip install reportlab",
                parent=top)
            return
        path = filedialog.asksaveasfilename(
            parent=top, title="Save exam schedule as PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path:
            return
        self._save_schedule_pdf(path, grades, grid, slot_meta, all_slots)
        messagebox.showinfo("Saved", f"Schedule saved as PDF:\n{path}", parent=top)

    def _save_schedule_pdf(self, path, grades, grid, slot_meta, all_slots):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
            from reportlab.lib import colors
        except ImportError as e:
            messagebox.showerror("PDF Error", f"reportlab not installed:\n{e}")
            return

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
        try:
            doc.build([table])
        except Exception as e:
            messagebox.showerror("PDF Error", f"Failed to write PDF:\n{e}")


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
