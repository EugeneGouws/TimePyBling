from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from datetime import date

from app.controller import (
    AppController,
    EVT_REGISTRY_BUILT,
    EVT_PAPERS_CHANGED,
    EVT_SCHEDULE_GENERATED,
    EVT_STATE_LOADED,
    EVT_TIMETABLE_LOADED,
)
from ui.constants import (
    SESSIONS,
    CLR_WHITE, CLR_LIGHT, CLR_BG,
    CLR_HEADER, CLR_BLUE, CLR_ORANGE, CLR_PINK,
    CLR_GREEN, CLR_RED, CLR_MID,
    CLR_MORNING, CLR_AFTERNOON,
    CLR_GRID_HEADER,
    CLR_DIFF_RED, CLR_DIFF_YELLOW, CLR_DIFF_GREEN, CLR_DIFF_GRAY,
    CLR_HOVER, CLR_ST_ROW,
    DEFAULT_EXAM_START, DEFAULT_EXAM_END,
)
from ui.helpers import _scrolled_text, _write, _clear

# Matrix cell dimensions
CELL_W    = 90
CELL_H    = 36
HDR_COL_W = 120
HDR_ROW_H = 30

DIFF_COLORS = {"red": CLR_DIFF_RED, "yellow": CLR_DIFF_YELLOW, "green": CLR_DIFF_GREEN}


class ExamTab(tk.Frame):

    def __init__(self, parent, controller: AppController, bus) -> None:
        super().__init__(parent, bg=CLR_WHITE)
        self._controller = controller
        self._bus = bus

        # Custom session list (None = derive from date range + AM/PM)
        self._sessions: list[tuple] | None = None

        # Slot summary widget — created in _build_scheduler, needed by
        # _update_session_count_label which may be called during _build_left
        self._slot_summary_text: tk.Text | None = None

        # Matrix state
        self._cell_ids: dict[tuple[str, str], int] = {}
        self._cell_text_ids: dict[tuple[str, str], int] = {}
        self._matrix_subjects: list[str] = []
        self._matrix_grades: list[str] = []
        self._hover_key: tuple[str, str] | None = None

        self._build()

        bus.subscribe(EVT_REGISTRY_BUILT,     self._on_exam_rebuilt)
        bus.subscribe(EVT_TIMETABLE_LOADED,   self._on_exam_rebuilt)
        bus.subscribe(EVT_PAPERS_CHANGED,     self._on_papers_changed)
        bus.subscribe(EVT_STATE_LOADED,       self._on_state_loaded)
        bus.subscribe(EVT_SCHEDULE_GENERATED, self._on_schedule_generated)

    # ------------------------------------------------------------------
    # Build
    # ------------------------------------------------------------------

    def _build(self) -> None:
        pane = tk.PanedWindow(self, orient=tk.HORIZONTAL,
                              bg="#ddd", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)
        self._build_left(pane)
        self._build_right(pane)

    def _build_left(self, pane: tk.PanedWindow) -> None:
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=520)

        # ── Header bar ─────────────────────────────────────────────
        hdr = tk.Frame(left, bg=CLR_WHITE)
        hdr.pack(fill=tk.X, padx=6, pady=(6, 0))

        # Row 0: Start / End dates + session count
        tk.Label(hdr, text="Start:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=0, column=0, sticky=tk.W)
        self._sched_start_var = tk.StringVar(
            value=DEFAULT_EXAM_START.strftime("%Y-%m-%d"))
        self._start_entry = tk.Entry(
            hdr, textvariable=self._sched_start_var,
            font=("Calibri", 10), relief=tk.SOLID, bd=1, width=11)
        self._start_entry.grid(row=0, column=1, sticky=tk.W, padx=(2, 8))
        self._start_entry.bind("<FocusOut>",
                               lambda e: self._on_session_param_changed())
        self._start_entry.bind("<Return>",
                               lambda e: self._on_session_param_changed())

        tk.Label(hdr, text="End:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=0, column=2, sticky=tk.W)
        self._sched_end_var = tk.StringVar(
            value=DEFAULT_EXAM_END.strftime("%Y-%m-%d"))
        self._end_entry = tk.Entry(
            hdr, textvariable=self._sched_end_var,
            font=("Calibri", 10), relief=tk.SOLID, bd=1, width=11)
        self._end_entry.grid(row=0, column=3, sticky=tk.W, padx=(2, 8))
        self._end_entry.bind("<FocusOut>",
                             lambda e: self._on_session_param_changed())
        self._end_entry.bind("<Return>",
                             lambda e: self._on_session_param_changed())

        self._session_count_label = tk.Label(
            hdr, text="", bg=CLR_WHITE, fg=CLR_ORANGE,
            font=("Calibri", 10, "bold"))
        self._session_count_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 4))

        # Row 1: AM / PM + Configure sessions
        self._am_var = tk.BooleanVar(value=True)
        self._pm_var = tk.BooleanVar(value=True)
        tk.Checkbutton(hdr, text="AM", variable=self._am_var,
                       bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                       selectcolor=CLR_ORANGE,
                       command=self._on_session_param_changed
                       ).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        tk.Checkbutton(hdr, text="PM", variable=self._pm_var,
                       bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                       selectcolor=CLR_PINK,
                       command=self._on_session_param_changed
                       ).grid(row=1, column=1, sticky=tk.W, pady=(4, 0))
        tk.Button(hdr, text="Configure sessions...",
                  command=self._open_session_calendar,
                  bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                  relief=tk.FLAT, padx=10
                  ).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=(4, 0))

        # Row 2: Rebuild / Save / Load
        tk.Button(hdr, text="Rebuild",
                  command=self._rebuild_exam,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=10
                  ).grid(row=2, column=0, sticky=tk.W, pady=(6, 0))
        tk.Button(hdr, text="Save State...",
                  command=self._export_exam_state,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8
                  ).grid(row=2, column=1, columnspan=2, sticky=tk.W,
                         padx=(4, 0), pady=(6, 0))
        tk.Button(hdr, text="Load State...",
                  command=self._import_exam_state,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8
                  ).grid(row=2, column=3, sticky=tk.W, pady=(6, 0))

        self._update_session_count_label()

        # ── Canvas matrix ──────────────────────────────────────────
        self._matrix_frame = tk.Frame(left, bg=CLR_WHITE)
        self._matrix_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(4, 6))

        self._matrix_canvas = tk.Canvas(
            self._matrix_frame, bg=CLR_WHITE, highlightthickness=0)
        h_sb = ttk.Scrollbar(self._matrix_frame, orient=tk.HORIZONTAL,
                             command=self._matrix_canvas.xview)
        v_sb = ttk.Scrollbar(self._matrix_frame, orient=tk.VERTICAL,
                             command=self._matrix_canvas.yview)
        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._matrix_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._matrix_canvas.configure(
            xscrollcommand=h_sb.set, yscrollcommand=v_sb.set)

        self._matrix_canvas.bind("<Motion>", self._on_matrix_motion)
        self._matrix_canvas.bind("<Leave>", self._on_matrix_leave)
        self._matrix_canvas.bind("<Button-1>", self._on_matrix_click)

    def _build_right(self, pane: tk.PanedWindow) -> None:
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=340)
        self._build_exclusions(right)
        self._build_scheduler(right)
        self._build_cost_function(right)

    def _build_exclusions(self, parent: tk.Frame) -> None:
        excl_frame = tk.LabelFrame(parent, text="Exam Exclusions",
                                   bg=CLR_WHITE, font=("Calibri", 10, "bold"),
                                   fg=CLR_BLUE, padx=8, pady=6)
        excl_frame.pack(fill=tk.X, padx=8, pady=(8, 4))

        self._excl_listbox = tk.Listbox(
            excl_frame, height=3, font=("Calibri", 10),
            relief=tk.SOLID, bd=1, bg=CLR_WHITE,
            selectbackground=CLR_ORANGE, selectforeground=CLR_WHITE,
            selectmode=tk.SINGLE)
        self._excl_listbox.pack(fill=tk.X, pady=(0, 4))
        self._refresh_exclusion_listbox()

        excl_row = tk.Frame(excl_frame, bg=CLR_WHITE)
        excl_row.pack(fill=tk.X)
        self._excl_entry = tk.Entry(excl_row, font=("Calibri", 10),
                                    relief=tk.SOLID, bd=1, width=8)
        self._excl_entry.pack(side=tk.LEFT)
        self._excl_entry.bind("<Return>", lambda e: self._add_exclusion())
        tk.Button(excl_row, text="Add", command=self._add_exclusion,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=8).pack(side=tk.LEFT, padx=3)
        tk.Button(excl_row, text="Remove", command=self._remove_exclusion,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.LEFT)

    def _build_scheduler(self, parent: tk.Frame) -> None:
        sched_lf = tk.LabelFrame(parent, text="Exam Schedule",
                                  bg=CLR_WHITE, font=("Calibri", 10, "bold"),
                                  fg=CLR_BLUE, padx=8, pady=6)
        sched_lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        ctrl = tk.Frame(sched_lf, bg=CLR_WHITE)
        ctrl.pack(fill=tk.X, pady=(0, 4))

        # Row 0: Grade filter + Generate
        tk.Label(ctrl, text="Grade:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=0, column=0, sticky=tk.W)
        self._sched_grade_var = tk.StringVar()
        self._sched_grade_cb = ttk.Combobox(
            ctrl, textvariable=self._sched_grade_var,
            state="readonly", width=10, font=("Calibri", 10))
        self._sched_grade_cb.grid(row=0, column=1, sticky=tk.W,
                                   padx=(4, 12))
        self._sched_grade_cb.bind("<<ComboboxSelected>>",
                                   lambda e: self._render_schedule())
        tk.Button(ctrl, text="Generate Schedule",
                  command=self._generate_exam_schedule,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12, pady=4
                  ).grid(row=0, column=2, sticky=tk.W, padx=(4, 0))

        # Row 1: Export / View All
        tk.Button(ctrl, text="Export / View All",
                  command=self._open_schedule_popout,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12, pady=4
                  ).grid(row=1, column=2, sticky=tk.W, padx=(4, 0), pady=(4, 0))

        # Slot summary
        self._slot_summary_text = tk.Text(
            sched_lf, height=5, font=("Calibri", 10),
            state=tk.DISABLED, relief=tk.FLAT, bg=CLR_LIGHT, wrap=tk.WORD)
        self._slot_summary_text.pack(fill=tk.X, padx=2, pady=(0, 4))
        self._slot_summary_text.tag_config("ok",    foreground=CLR_GREEN,
                                            font=("Calibri", 10))
        self._slot_summary_text.tag_config("short", foreground=CLR_RED,
                                            font=("Calibri", 10))
        self._slot_summary_text.tag_config("dim",   foreground="#6B7280",
                                            font=("Calibri", 10))

        # Schedule output text
        self._sched_text = _scrolled_text(sched_lf)
        self._sched_text.tag_config("header",
                                    font=("Calibri", 11, "bold"),
                                    foreground=CLR_HEADER)
        self._sched_text.tag_config("am",  background=CLR_MORNING,
                                    foreground="#7C2D12")
        self._sched_text.tag_config("pm",  background=CLR_AFTERNOON,
                                    foreground="#831843")
        self._sched_text.tag_config("dim",     foreground="#6B7280",
                                    font=("Calibri", 10))
        self._sched_text.tag_config("warn",    foreground=CLR_ORANGE,
                                    font=("Calibri", 10))
        self._sched_text.tag_config("day_sep", foreground=CLR_MID,
                                    font=("Calibri", 10))

    def _build_cost_function(self, parent: tk.Frame) -> None:
        cost_lf = tk.LabelFrame(parent, text="Cost Configuration",
                                 bg=CLR_WHITE, font=("Calibri", 10, "bold"),
                                 fg=CLR_BLUE, padx=8, pady=6)
        cost_lf.pack(fill=tk.X, padx=8, pady=(0, 8))

        wf = tk.Frame(cost_lf, bg=CLR_WHITE)
        wf.pack(fill=tk.X)

        cost_fields = [
            ("day_density_factor", 5),
            ("week_density_base", 6),
            ("same_week_penalty", 1),
            ("teacher_load_penalty", 1),
        ]
        self._cost_vars: dict[str, tk.StringVar] = {}
        for i, (field_name, default) in enumerate(cost_fields):
            r, c = i // 2, (i % 2) * 3
            tk.Label(wf, text=f"{field_name}:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10)).grid(row=r, column=c, sticky=tk.W,
                                                 padx=(0, 2), pady=2)
            v = tk.StringVar(value=str(default))
            self._cost_vars[field_name] = v
            tk.Entry(wf, textvariable=v, font=("Calibri", 10),
                     relief=tk.SOLID, bd=1, width=5
                     ).grid(row=r, column=c + 1, sticky=tk.W,
                             padx=(0, 12), pady=2)

        # Hard constraint display-only checkboxes
        chk_frame = tk.Frame(cost_lf, bg=CLR_WHITE)
        chk_frame.pack(fill=tk.X, pady=(4, 0))
        self._enforce_clash_var = tk.BooleanVar(value=True)
        tk.Checkbutton(chk_frame, text="enforce_student_clash",
                       variable=self._enforce_clash_var, state=tk.DISABLED,
                       bg=CLR_WHITE, fg="#6B7280", font=("Calibri", 9)
                       ).pack(side=tk.LEFT)
        self._enforce_constr_var = tk.BooleanVar(value=True)
        tk.Checkbutton(chk_frame, text="enforce_constraint_code",
                       variable=self._enforce_constr_var, state=tk.DISABLED,
                       bg=CLR_WHITE, fg="#6B7280", font=("Calibri", 9)
                       ).pack(side=tk.LEFT, padx=(8, 0))

        # Action buttons
        btn_frame = tk.Frame(cost_lf, bg=CLR_WHITE)
        btn_frame.pack(fill=tk.X, pady=(4, 0))
        tk.Button(btn_frame, text="Rebuild Schedule",
                  command=self._generate_exam_schedule,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=10).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Breakdown",
                  command=self._open_penalty_breakdown,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=10).pack(side=tk.LEFT, padx=4)

        self._exam_cost_result_label = tk.Label(
            cost_lf, text="", bg=CLR_WHITE, fg="#1E293B",
            font=("Calibri", 10), justify=tk.LEFT, anchor=tk.W)
        self._exam_cost_result_label.pack(fill=tk.X, pady=(4, 0))

    # ------------------------------------------------------------------
    # Event callbacks
    # ------------------------------------------------------------------

    def _on_exam_rebuilt(self, **_) -> None:
        self._populate_matrix()
        self._update_sched_grade_list()
        self._update_session_count_label()

    def _on_papers_changed(self, **_) -> None:
        self._populate_matrix()

    def _on_state_loaded(self, state, **_) -> None:
        cfg = state.session_config
        if cfg is not None:
            self._sched_start_var.set(cfg.start.strftime("%Y-%m-%d"))
            self._sched_end_var.set(cfg.end.strftime("%Y-%m-%d"))
            self._am_var.set(cfg.am)
            self._pm_var.set(cfg.pm)
            self._sessions = None
        # Restore cost config values
        cost = state.cost_config
        if cost is not None:
            self._cost_vars["day_density_factor"].set(str(cost.day_density_factor))
            self._cost_vars["week_density_base"].set(str(cost.week_density_base))
            self._cost_vars["same_week_penalty"].set(str(cost.same_week_penalty))
            self._cost_vars["teacher_load_penalty"].set(str(cost.teacher_load_penalty))
        self._refresh_exclusion_listbox()
        self._populate_matrix()
        self._update_sched_grade_list()
        self._update_session_count_label()

    def _on_schedule_generated(self, **_) -> None:
        self._render_schedule()
        self._update_cost_display()

    # ------------------------------------------------------------------
    # Canvas matrix
    # ------------------------------------------------------------------

    def _populate_matrix(self) -> None:
        canvas = self._matrix_canvas
        canvas.delete("all")
        self._cell_ids.clear()
        self._cell_text_ids.clear()
        self._hover_key = None

        registry = self._controller.state.paper_registry
        if not registry:
            canvas.create_text(
                10, 10, text="Load a timetable to see the exam matrix.",
                anchor="nw", font=("Calibri", 11), fill="#6B7280")
            return

        exclusions = self._controller.state.exclusions
        grades = registry.grades()
        all_subjects: set[str] = set()
        for grade in grades:
            for subj in registry.subjects_for_grade(grade):
                if subj != "ST" and subj not in exclusions:
                    all_subjects.add(subj)
        subjects = sorted(all_subjects)
        self._matrix_subjects = subjects
        self._matrix_grades = grades

        # ── Header row (grades) ──
        for ci, grade in enumerate(grades):
            x1 = HDR_COL_W + ci * CELL_W
            x2 = x1 + CELL_W
            canvas.create_rectangle(x1, 0, x2, HDR_ROW_H,
                                    fill=CLR_GRID_HEADER, outline="#CBD5E1")
            canvas.create_text((x1 + x2) // 2, HDR_ROW_H // 2,
                               text=grade, font=("Calibri", 10, "bold"),
                               fill="#1E293B")

        # ── Header column (subjects) ──
        for ri, subj in enumerate(subjects):
            y1 = HDR_ROW_H + ri * CELL_H
            y2 = y1 + CELL_H
            canvas.create_rectangle(0, y1, HDR_COL_W, y2,
                                    fill=CLR_WHITE, outline="#CBD5E1")
            canvas.create_text(8, (y1 + y2) // 2, text=subj,
                               anchor="w", font=("Calibri", 10, "bold"),
                               fill="#1E293B")

        # ── Body cells ──
        for ri, subj in enumerate(subjects):
            for ci, grade in enumerate(grades):
                x1 = HDR_COL_W + ci * CELL_W
                y1 = HDR_ROW_H + ri * CELL_H
                x2 = x1 + CELL_W
                y2 = y1 + CELL_H

                papers = registry.papers_for_subject_grade(subj, grade)
                if not papers:
                    fill = CLR_DIFF_GRAY
                    text = ""
                else:
                    diff = registry.get_difficulty(subj, grade)
                    fill = DIFF_COLORS.get(diff, CLR_DIFF_GREEN)
                    parts = [f"P{p.paper_number}" for p in papers]
                    pinned = any(p.pinned_slot is not None for p in papers)
                    text = " ".join(parts)
                    if pinned:
                        text += " \U0001f4cc"  # pin emoji

                rect_id = canvas.create_rectangle(
                    x1, y1, x2, y2, fill=fill, outline="#CBD5E1")
                text_id = canvas.create_text(
                    (x1 + x2) // 2, (y1 + y2) // 2,
                    text=text, font=("Calibri", 9), fill="#1E293B")
                self._cell_ids[(subj, grade)] = rect_id
                self._cell_text_ids[(subj, grade)] = text_id

        # ── ST row ──
        st_y1 = HDR_ROW_H + len(subjects) * CELL_H + 4
        st_y2 = st_y1 + CELL_H
        canvas.create_rectangle(0, st_y1, HDR_COL_W, st_y2,
                                fill=CLR_ST_ROW, outline="#CBD5E1")
        canvas.create_text(8, (st_y1 + st_y2) // 2, text="ST (Study)",
                           anchor="w", font=("Calibri", 10, "bold"),
                           fill="#1E293B")
        for ci, grade in enumerate(grades):
            x1 = HDR_COL_W + ci * CELL_W
            x2 = x1 + CELL_W
            st_papers = registry.papers_for_subject_grade("ST", grade)
            text = str(len(st_papers)) if st_papers else "+"
            rect_id = canvas.create_rectangle(
                x1, st_y1, x2, st_y2, fill=CLR_ST_ROW, outline="#CBD5E1")
            text_id = canvas.create_text(
                (x1 + x2) // 2, (st_y1 + st_y2) // 2,
                text=text, font=("Calibri", 10, "bold"),
                fill="#1E293B" if st_papers else CLR_GREEN)
            self._cell_ids[("ST", grade)] = rect_id
            self._cell_text_ids[("ST", grade)] = text_id

        canvas.configure(scrollregion=canvas.bbox("all"))

    def _cell_at(self, cx: float, cy: float) -> tuple[str, str] | None:
        """Return (subject, grade) for canvas coords, or None."""
        if cx < HDR_COL_W or cy < HDR_ROW_H:
            return None
        ci = int((cx - HDR_COL_W) // CELL_W)
        if ci < 0 or ci >= len(self._matrix_grades):
            return None
        grade = self._matrix_grades[ci]

        # Check ST row
        st_y1 = HDR_ROW_H + len(self._matrix_subjects) * CELL_H + 4
        if cy >= st_y1:
            return ("ST", grade)

        ri = int((cy - HDR_ROW_H) // CELL_H)
        if ri < 0 or ri >= len(self._matrix_subjects):
            return None
        return (self._matrix_subjects[ri], grade)

    def _on_matrix_motion(self, event) -> None:
        cx = self._matrix_canvas.canvasx(event.x)
        cy = self._matrix_canvas.canvasy(event.y)
        key = self._cell_at(cx, cy)

        if key == self._hover_key:
            return
        # Restore previous hover
        if self._hover_key and self._hover_key in self._cell_ids:
            self._restore_cell_fill(self._hover_key)
        self._hover_key = key
        # Highlight new
        if key and key in self._cell_ids:
            self._matrix_canvas.itemconfig(self._cell_ids[key], fill=CLR_HOVER)

    def _on_matrix_leave(self, event) -> None:
        if self._hover_key and self._hover_key in self._cell_ids:
            self._restore_cell_fill(self._hover_key)
        self._hover_key = None

    def _on_matrix_click(self, event) -> None:
        cx = self._matrix_canvas.canvasx(event.x)
        cy = self._matrix_canvas.canvasy(event.y)
        key = self._cell_at(cx, cy)
        if not key:
            return
        subj, grade = key
        if subj == "ST":
            self._open_st_cell_popout(grade)
        else:
            # Only open popout for non-empty cells
            registry = self._controller.state.paper_registry
            if registry and registry.papers_for_subject_grade(subj, grade):
                self._open_cell_popout(subj, grade)

    def _restore_cell_fill(self, key: tuple[str, str]) -> None:
        subj, grade = key
        registry = self._controller.state.paper_registry
        if not registry:
            return
        if subj == "ST":
            fill = CLR_ST_ROW
        else:
            papers = registry.papers_for_subject_grade(subj, grade)
            if not papers:
                fill = CLR_DIFF_GRAY
            else:
                diff = registry.get_difficulty(subj, grade)
                fill = DIFF_COLORS.get(diff, CLR_DIFF_GREEN)
        if key in self._cell_ids:
            self._matrix_canvas.itemconfig(self._cell_ids[key], fill=fill)

    # ------------------------------------------------------------------
    # Cell popout (subject + grade)
    # ------------------------------------------------------------------

    def _open_cell_popout(self, subject: str, grade: str) -> None:
        registry = self._controller.state.paper_registry
        if not registry:
            return
        papers = registry.papers_for_subject_grade(subject, grade)
        if not papers:
            return

        top = tk.Toplevel(self)
        top.title(f"{subject} \u2014 {grade}")
        top.geometry("420x560")
        top.configure(bg=CLR_WHITE)
        top.grab_set()

        tk.Label(top, text=f"{subject} \u2014 {grade}",
                 bg=CLR_WHITE, fg=CLR_HEADER,
                 font=("Calibri", 13, "bold")).pack(
                     pady=(10, 6), padx=10, anchor=tk.W)

        # Scrollable content
        content_canvas = tk.Canvas(top, bg=CLR_WHITE, highlightthickness=0)
        csb = ttk.Scrollbar(top, orient=tk.VERTICAL, command=content_canvas.yview)
        csb.pack(side=tk.RIGHT, fill=tk.Y)
        content_canvas.pack(fill=tk.BOTH, expand=True, padx=10)
        content_canvas.configure(yscrollcommand=csb.set)
        inner = tk.Frame(content_canvas, bg=CLR_WHITE)
        content_canvas.create_window((0, 0), window=inner, anchor="nw")

        def _refresh_popout():
            for w in inner.winfo_children():
                w.destroy()
            _build_popout_content()
            inner.update_idletasks()
            content_canvas.configure(scrollregion=content_canvas.bbox("all"))

        def _build_popout_content():
            nonlocal papers
            papers = registry.papers_for_subject_grade(subject, grade)

            # ── Paper list ──
            tk.Label(inner, text="Papers:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10, "bold")).pack(anchor=tk.W, pady=(4, 2))
            for p in papers:
                pf = tk.Frame(inner, bg=CLR_LIGHT)
                pf.pack(fill=tk.X, pady=1)
                constr_str = (", ".join(sorted(p.constraints))
                              if p.constraints else "")
                pin_str = f" \U0001f4cc slot {p.pinned_slot + 1}" if p.pinned_slot is not None else ""
                link_str = f" \u2194 {', '.join(sorted(p.links))}" if p.links else ""
                tk.Label(pf, text=f"P{p.paper_number}  ({p.student_count()} students)"
                              f"{pin_str}{link_str}",
                         bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                         anchor=tk.W).pack(side=tk.LEFT, padx=4)
                if constr_str:
                    tk.Label(pf, text=f"[{constr_str}]",
                             bg=CLR_LIGHT, fg=CLR_PINK, font=("Calibri", 9),
                             anchor=tk.W).pack(side=tk.LEFT, padx=2)

                # Per-paper buttons
                btn_row = tk.Frame(pf, bg=CLR_LIGHT)
                btn_row.pack(side=tk.RIGHT)
                lbl = p.label
                tk.Button(btn_row, text="Pin",
                          command=lambda l=lbl: _pin_paper(l),
                          bg=CLR_PINK, fg=CLR_WHITE, relief=tk.FLAT,
                          font=("Calibri", 9), padx=4).pack(side=tk.LEFT, padx=1)
                if p.pinned_slot is not None:
                    tk.Button(btn_row, text="Unpin",
                              command=lambda l=lbl: _unpin_paper(l),
                              bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                              font=("Calibri", 9), padx=4).pack(side=tk.LEFT, padx=1)

            # ── Add / Remove paper ──
            paper_btn = tk.Frame(inner, bg=CLR_WHITE)
            paper_btn.pack(fill=tk.X, pady=(4, 0))
            tk.Button(paper_btn, text="+ Add Paper",
                      command=lambda: _add_paper(),
                      bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                      font=("Calibri", 10, "bold"), padx=8).pack(side=tk.LEFT)
            if len(papers) > 1:
                tk.Button(paper_btn, text="- Remove Last",
                          command=lambda: _remove_last_paper(),
                          bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                          font=("Calibri", 10), padx=8).pack(side=tk.LEFT, padx=4)

            # ── Difficulty ──
            tk.Label(inner, text="Difficulty:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10, "bold")).pack(anchor=tk.W, pady=(8, 2))
            diff_frame = tk.Frame(inner, bg=CLR_WHITE)
            diff_frame.pack(fill=tk.X)
            diff_var = tk.StringVar(value=registry.get_difficulty(subject, grade))
            for diff_val, diff_label, diff_clr in [
                ("green", "Green", CLR_DIFF_GREEN),
                ("yellow", "Yellow", CLR_DIFF_YELLOW),
                ("red", "Red", CLR_DIFF_RED),
            ]:
                tk.Radiobutton(diff_frame, text=diff_label, variable=diff_var,
                               value=diff_val,
                               command=lambda v=diff_var: _set_difficulty(v.get()),
                               bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                               selectcolor=diff_clr
                               ).pack(side=tk.LEFT, padx=4)

            # ── Links ──
            tk.Label(inner, text="Links:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10, "bold")).pack(anchor=tk.W, pady=(8, 2))
            current_links: set[str] = set()
            for p in papers:
                current_links |= p.links
            if current_links:
                for link_label in sorted(current_links):
                    lf = tk.Frame(inner, bg=CLR_WHITE)
                    lf.pack(fill=tk.X, pady=1)
                    tk.Label(lf, text=link_label, bg=CLR_WHITE, fg="#1E293B",
                             font=("Calibri", 10)).pack(side=tk.LEFT, padx=4)
                    tk.Button(lf, text="x",
                              command=lambda ll=link_label: _remove_link(ll),
                              bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                              font=("Calibri", 9), padx=4).pack(side=tk.LEFT)
            else:
                tk.Label(inner, text="  (none)", bg=CLR_WHITE, fg="#6B7280",
                         font=("Calibri", 10)).pack(anchor=tk.W)

            link_row = tk.Frame(inner, bg=CLR_WHITE)
            link_row.pack(fill=tk.X, pady=(2, 0))
            link_entry = tk.Entry(link_row, font=("Calibri", 10),
                                  relief=tk.SOLID, bd=1, width=18)
            link_entry.pack(side=tk.LEFT, padx=(0, 4))
            tk.Button(link_row, text="Add Link",
                      command=lambda: _add_link(link_entry.get()),
                      bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                      font=("Calibri", 10), padx=6).pack(side=tk.LEFT)

            # ── Constraints ──
            tk.Label(inner, text="Constraints:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10, "bold")).pack(anchor=tk.W, pady=(8, 2))
            for p in papers:
                cf = tk.Frame(inner, bg=CLR_WHITE)
                cf.pack(fill=tk.X, pady=1)
                tk.Label(cf, text=f"P{p.paper_number}:",
                         bg=CLR_WHITE, fg="#1E293B",
                         font=("Calibri", 10)).pack(side=tk.LEFT, padx=4)
                for code in sorted(p.constraints):
                    tk.Label(cf, text=code, bg=CLR_PINK, fg=CLR_WHITE,
                             font=("Calibri", 9), padx=3).pack(side=tk.LEFT, padx=1)
                    tk.Button(cf, text="x",
                              command=lambda l=p.label, c=code: _remove_constraint(l, c),
                              bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                              font=("Calibri", 8), padx=2).pack(side=tk.LEFT)
            constr_row = tk.Frame(inner, bg=CLR_WHITE)
            constr_row.pack(fill=tk.X, pady=(2, 0))
            constr_entry = tk.Entry(constr_row, font=("Calibri", 10),
                                    relief=tk.SOLID, bd=1, width=8)
            constr_entry.pack(side=tk.LEFT, padx=(0, 4))
            # Apply to which paper?
            paper_var = tk.StringVar(value=papers[0].label if papers else "")
            if len(papers) > 1:
                tk.Label(constr_row, text="for:", bg=CLR_WHITE,
                         font=("Calibri", 10)).pack(side=tk.LEFT)
                paper_cb = ttk.Combobox(constr_row, textvariable=paper_var,
                                        state="readonly", width=16,
                                        values=[p.label for p in papers])
                paper_cb.pack(side=tk.LEFT, padx=2)
            tk.Button(constr_row, text="Add",
                      command=lambda: _add_constraint(paper_var.get(),
                                                       constr_entry.get()),
                      bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                      font=("Calibri", 10), padx=6).pack(side=tk.LEFT)

            # ── Close ──
            tk.Button(inner, text="Close", command=top.destroy,
                      bg=CLR_LIGHT, fg="#1E293B", relief=tk.FLAT,
                      font=("Calibri", 10), padx=16, pady=4
                      ).pack(pady=(12, 8))

        def _add_paper():
            self._controller.add_paper(subject, grade)
            _refresh_popout()

        def _remove_last_paper():
            ps = registry.papers_for_subject_grade(subject, grade)
            if len(ps) <= 1:
                return
            last = max(ps, key=lambda p: p.paper_number)
            self._controller.remove_paper(last.label)
            _refresh_popout()

        def _set_difficulty(value: str):
            self._controller.set_difficulty(subject, grade, value)
            _refresh_popout()

        def _add_link(target_label: str):
            target_label = target_label.strip()
            if not target_label:
                return
            # Link from the first paper
            try:
                self._controller.add_link(papers[0].label, target_label)
            except ValueError as e:
                messagebox.showwarning("Invalid link", str(e), parent=top)
                return
            _refresh_popout()

        def _remove_link(target_label: str):
            for p in papers:
                if target_label in p.links:
                    self._controller.remove_link(p.label, target_label)
            _refresh_popout()

        def _add_constraint(paper_label: str, code: str):
            code = code.strip().upper()
            if not code or not paper_label:
                return
            self._controller.add_constraint(paper_label, code)
            _refresh_popout()

        def _remove_constraint(paper_label: str, code: str):
            self._controller.remove_constraint(paper_label, code)
            _refresh_popout()

        def _pin_paper(paper_label: str):
            paper = registry.get(paper_label)
            if not paper:
                return
            sessions = self._get_effective_sessions()
            if not sessions:
                messagebox.showinfo(
                    "No sessions",
                    "Configure exam sessions (start/end dates) first.",
                    parent=top)
                return
            self._open_slot_picker(paper, sessions, on_done=_refresh_popout)

        def _unpin_paper(paper_label: str):
            paper = registry.get(paper_label)
            if paper:
                paper.pinned_slot = None
                self._bus.publish(EVT_PAPERS_CHANGED, state=self._controller.state)
            _refresh_popout()

        _build_popout_content()
        inner.update_idletasks()
        content_canvas.configure(scrollregion=content_canvas.bbox("all"))

    # ------------------------------------------------------------------
    # ST cell popout
    # ------------------------------------------------------------------

    def _open_st_cell_popout(self, grade: str) -> None:
        registry = self._controller.state.paper_registry
        if not registry:
            return

        top = tk.Toplevel(self)
        top.title(f"Study Papers \u2014 {grade}")
        top.geometry("360x400")
        top.configure(bg=CLR_WHITE)
        top.grab_set()

        tk.Label(top, text=f"Study Papers \u2014 {grade}",
                 bg=CLR_WHITE, fg=CLR_HEADER,
                 font=("Calibri", 13, "bold")).pack(
                     pady=(10, 6), padx=10, anchor=tk.W)

        list_frame = tk.Frame(top, bg=CLR_WHITE)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        def _refresh_st():
            for w in list_frame.winfo_children():
                w.destroy()
            st_papers = registry.papers_for_subject_grade("ST", grade)
            if not st_papers:
                tk.Label(list_frame, text="No study papers for this grade.",
                         bg=CLR_WHITE, fg="#6B7280",
                         font=("Calibri", 10)).pack(anchor=tk.W, pady=4)
            for p in st_papers:
                pf = tk.Frame(list_frame, bg=CLR_LIGHT)
                pf.pack(fill=tk.X, pady=2)
                pin_str = f"  \U0001f4cc slot {p.pinned_slot + 1}" if p.pinned_slot is not None else ""
                tk.Label(pf, text=f"{p.label}{pin_str}",
                         bg=CLR_LIGHT, fg="#1E293B",
                         font=("Calibri", 10)).pack(side=tk.LEFT, padx=4)
                tk.Button(pf, text="Remove",
                          command=lambda l=p.label: _remove_st(l),
                          bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                          font=("Calibri", 9), padx=4).pack(side=tk.RIGHT, padx=2)
                if p.pinned_slot is not None:
                    tk.Button(pf, text="Unpin",
                              command=lambda l=p.label: _unpin_st(l),
                              bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                              font=("Calibri", 9), padx=4).pack(side=tk.RIGHT, padx=1)
                tk.Button(pf, text="Pin",
                          command=lambda l=p.label: _pin_st(l),
                          bg=CLR_PINK, fg=CLR_WHITE, relief=tk.FLAT,
                          font=("Calibri", 9), padx=4).pack(side=tk.RIGHT, padx=1)

        def _remove_st(label: str):
            self._controller.remove_study_paper(label)
            _refresh_st()

        def _unpin_st(label: str):
            paper = registry.get(label)
            if paper:
                paper.pinned_slot = None
                self._bus.publish(EVT_PAPERS_CHANGED, state=self._controller.state)
            _refresh_st()

        def _pin_st(label: str):
            paper = registry.get(label)
            if not paper:
                return
            sessions = self._get_effective_sessions()
            if not sessions:
                messagebox.showinfo("No sessions",
                                    "Configure exam sessions first.",
                                    parent=top)
                return
            self._open_slot_picker(paper, sessions, on_done=_refresh_st)

        _refresh_st()

        btn_frame = tk.Frame(top, bg=CLR_WHITE)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)
        tk.Button(btn_frame, text="+ Add Study Paper",
                  command=lambda: (self._controller.add_study_paper(grade),
                                   _refresh_st()),
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=10).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Close", command=top.destroy,
                  bg=CLR_LIGHT, fg="#1E293B", relief=tk.FLAT,
                  font=("Calibri", 10), padx=10).pack(side=tk.LEFT, padx=8)

    # ------------------------------------------------------------------
    # Navigate to cell (from schedule popout)
    # ------------------------------------------------------------------

    def _navigate_to_exam_subject(self, subject: str, grade: str,
                                   popout: tk.Toplevel | None = None) -> None:
        if popout:
            popout.destroy()
        key = (subject, grade)
        if key not in self._cell_ids:
            return
        rect_id = self._cell_ids[key]
        coords = self._matrix_canvas.coords(rect_id)
        if coords:
            bbox = self._matrix_canvas.cget("scrollregion")
            if bbox:
                parts = bbox.split()
                total_h = float(parts[3]) if len(parts) >= 4 else 1
                self._matrix_canvas.yview_moveto(coords[1] / total_h)
        # Flash highlight
        original_fill = self._matrix_canvas.itemcget(rect_id, "fill")
        self._matrix_canvas.itemconfig(rect_id, fill=CLR_HOVER)
        self._matrix_canvas.after(
            600, lambda: self._matrix_canvas.itemconfig(rect_id, fill=original_fill))

    # ------------------------------------------------------------------
    # Rebuild
    # ------------------------------------------------------------------

    def _rebuild_exam(self) -> None:
        if self._controller.state.timetable_tree is None:
            return
        self._controller.build_exam_tree()
        self._controller.build_registry()
        self._controller.state.schedule_result = None

    # ------------------------------------------------------------------
    # Grade list + slot summary
    # ------------------------------------------------------------------

    def _update_sched_grade_list(self) -> None:
        registry = self._controller.state.paper_registry
        if not registry:
            return
        grades = registry.grades()
        options = ["All Grades"] + grades
        self._sched_grade_cb["values"] = options
        if grades:
            self._sched_grade_var.set(grades[-1])

    # ------------------------------------------------------------------
    # Exclusions
    # ------------------------------------------------------------------

    def _refresh_exclusion_listbox(self) -> None:
        self._excl_listbox.delete(0, tk.END)
        for code in sorted(self._controller.state.exclusions):
            self._excl_listbox.insert(tk.END, code)

    def _add_exclusion(self) -> None:
        code = self._excl_entry.get().strip().upper()
        if not code:
            return
        new_excl = set(self._controller.state.exclusions)
        new_excl.add(code)
        self._excl_entry.delete(0, tk.END)
        self._controller.set_exclusions(new_excl)

    def _remove_exclusion(self) -> None:
        sel = self._excl_listbox.curselection()
        if not sel:
            return
        code = self._excl_listbox.get(sel[0])
        new_excl = set(self._controller.state.exclusions)
        new_excl.discard(code)
        self._controller.set_exclusions(new_excl)

    # ------------------------------------------------------------------
    # Import / Export state
    # ------------------------------------------------------------------

    def _export_exam_state(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save timetable state",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            self._controller.save_to_json(path)
            messagebox.showinfo("Saved", f"State saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def _import_exam_state(self) -> None:
        path = filedialog.askopenfilename(
            title="Load timetable state",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            self._controller.load_from_json(path)
            messagebox.showinfo("Loaded", f"State loaded from:\n{path}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    # ------------------------------------------------------------------
    # Schedule generation
    # ------------------------------------------------------------------

    def _generate_exam_schedule(self) -> None:
        if not self._controller.state.paper_registry:
            messagebox.showinfo("No data", "Load a timetable first.")
            return
        sessions = self._get_effective_sessions()
        if sessions is None:
            try:
                date.fromisoformat(self._sched_start_var.get().strip())
                date.fromisoformat(self._sched_end_var.get().strip())
                msg = "End date must be on or after start date."
            except ValueError:
                msg = "Enter valid start and end dates (YYYY-MM-DD)."
            messagebox.showerror("Invalid dates", msg)
            return
        if not sessions:
            messagebox.showerror(
                "No sessions",
                "No exam sessions in the selected date range.\n"
                "Check AM/PM checkboxes and date range.")
            return

        # Build CostConfig from UI entries
        from app.cost_config import CostConfig  # noqa: PLC0415
        try:
            config = CostConfig(
                day_density_factor=int(float(
                    self._cost_vars["day_density_factor"].get())),
                week_density_base=int(float(
                    self._cost_vars["week_density_base"].get())),
                same_week_penalty=int(float(
                    self._cost_vars["same_week_penalty"].get())),
                teacher_load_penalty=int(float(
                    self._cost_vars["teacher_load_penalty"].get())),
            )
        except (ValueError, KeyError):
            config = CostConfig()

        self._controller.state.cost_config = config
        self._controller.generate_schedule(sessions=sessions)

    def _render_schedule(self) -> None:
        w = self._sched_text
        _clear(w)
        result = self._controller.state.schedule_result
        if not result:
            return

        grade_filter = self._sched_grade_var.get().strip()
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
                  f"  \u2014  {result.total_days} day(s)"
                  f"  \u2014  {result.total_slots} slot(s)"
                  f"  \u2014  {len(result.scheduled)} papers total\n", "header")
        if grade_filter and grade_filter != "All Grades":
            _write(w, f"Filtered: {grade_filter}\n", "dim")
        _write(w, "\u2500" * 68 + "\n", "dim")
        _write(w, f"  {'Date':<14} {'Sess':<5} {'Paper':<22} "
                  f"{'Students':>8}  Warnings\n", "header")
        _write(w, "\u2500" * 68 + "\n", "dim")

        prev_date = None
        for sp in items:
            if prev_date is not None and sp.date != prev_date:
                _write(w, "\n", "day_sep")
            prev_date = sp.date
            date_str  = sp.date.strftime("%a %d %b")
            pin_mark  = " \U0001f4cc" if sp.pinned else ""
            warn_flag = "  \u26a0" if sp.warnings else ""
            line = (f"  {date_str:<14} {sp.session:<5} "
                    f"{sp.paper.label:<22} "
                    f"{sp.paper.student_count():>8}{pin_mark}{warn_flag}\n")
            tag = "am" if sp.session == "AM" else "pm"
            _write(w, line, tag)

        _write(w, "\u2500" * 68 + "\n", "dim")

        all_warnings = [w_msg for sp in items for w_msg in sp.warnings]
        seen: set[str] = set()
        unique_warnings = [x for x in all_warnings
                           if not (x in seen or seen.add(x))]
        if unique_warnings:
            _write(w, "\nWarnings:\n", "warn")
            for msg in unique_warnings:
                _write(w, f"  \u26a0  {msg}\n", "warn")

        if result.teacher_warnings:
            _write(w, "\nTeacher marking load conflicts:\n", "warn")
            for msg in result.teacher_warnings:
                _write(w, f"  \u26a0  {msg}\n", "warn")

    # ------------------------------------------------------------------
    # Cost display
    # ------------------------------------------------------------------

    def _update_cost_display(self) -> None:
        result = self._controller.state.schedule_result
        if not result:
            self._exam_cost_result_label.config(text="")
            return

        if hasattr(result, "penalty_log") and result.penalty_log:
            by_type: dict[str, int] = defaultdict(int)
            for entry in result.penalty_log:
                by_type[entry.constraint] += entry.value
            total = sum(by_type.values())
            lines = [f"Total cost: {total}"]
            for constraint, subtotal in sorted(by_type.items()):
                lines.append(f"  {constraint}: {subtotal}")
            self._exam_cost_result_label.config(text="\n".join(lines))
        else:
            self._exam_cost_result_label.config(
                text=f"Student cost: {result.student_cost}")

    # ------------------------------------------------------------------
    # Penalty breakdown popout
    # ------------------------------------------------------------------

    def _open_penalty_breakdown(self) -> None:
        result = self._controller.state.schedule_result
        if not result:
            messagebox.showinfo("No schedule", "Generate a schedule first.")
            return
        if not hasattr(result, "penalty_log") or not result.penalty_log:
            messagebox.showinfo("No data",
                                "No penalty entries recorded.\n"
                                "Generate a schedule first.")
            return

        top = tk.Toplevel(self)
        top.title("Penalty Breakdown")
        top.geometry("780x520")
        top.configure(bg=CLR_WHITE)

        # Treeview
        cols = ("constraint", "papers", "entity", "value")
        tree = ttk.Treeview(top, columns=cols, show="headings", height=20)

        sort_state: dict[str, bool] = {}

        def _sort_column(col: str):
            reverse = sort_state.get(col, False)
            sort_state[col] = not reverse
            items = [(tree.set(iid, col), iid) for iid in tree.get_children()]
            if col == "value":
                items.sort(key=lambda x: int(x[0]) if x[0].isdigit() else 0,
                           reverse=reverse)
            else:
                items.sort(key=lambda x: x[0], reverse=reverse)
            for idx, (_, iid) in enumerate(items):
                tree.move(iid, "", idx)

        for col, width in [("constraint", 120), ("papers", 300),
                           ("entity", 120), ("value", 80)]:
            tree.heading(col, text=col.title(),
                         command=lambda c=col: _sort_column(c))
            tree.column(col, width=width)

        sb = ttk.Scrollbar(top, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)

        tree_frame = tk.Frame(top, bg=CLR_WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 0))
        tree.pack(in_=tree_frame, side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(in_=tree_frame, side=tk.RIGHT, fill=tk.Y)

        # Populate (default: sort by value descending)
        entries = sorted(result.penalty_log, key=lambda e: -e.value)
        for entry in entries:
            tree.insert("", tk.END, values=(
                entry.constraint,
                ", ".join(entry.papers[:6]) + ("..." if len(entry.papers) > 6 else ""),
                entry.entity,
                entry.value,
            ))

        # Summary
        by_type: dict[str, int] = defaultdict(int)
        for entry in result.penalty_log:
            by_type[entry.constraint] += entry.value
        summary = "  |  ".join(f"{k}: {v}" for k, v in sorted(by_type.items()))
        tk.Label(top, text=f"Entries: {len(result.penalty_log)}  |  {summary}",
                 bg=CLR_WHITE, fg="#6B7280", font=("Calibri", 9),
                 anchor=tk.W).pack(fill=tk.X, padx=8)

        # Buttons
        btn_frame = tk.Frame(top, bg=CLR_WHITE)
        btn_frame.pack(fill=tk.X, padx=8, pady=8)

        def _export_text():
            path = filedialog.asksaveasfilename(
                parent=top, title="Export penalty log",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt")])
            if not path:
                return
            with open(path, "w", encoding="utf-8") as f:
                f.write("constraint\tpapers\tentity\tvalue\n")
                for entry in sorted(result.penalty_log, key=lambda e: -e.value):
                    f.write(f"{entry.constraint}\t"
                            f"{', '.join(entry.papers)}\t"
                            f"{entry.entity}\t{entry.value}\n")
            messagebox.showinfo("Exported", f"Saved to:\n{path}", parent=top)

        tk.Button(btn_frame, text="Export as text", command=_export_text,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=10).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Close", command=top.destroy,
                  bg=CLR_LIGHT, fg="#1E293B", relief=tk.FLAT,
                  font=("Calibri", 10), padx=10).pack(side=tk.LEFT, padx=8)

    # ------------------------------------------------------------------
    # Session dates + slot count
    # ------------------------------------------------------------------

    def _get_effective_sessions(self) -> list[tuple] | None:
        if self._sessions is not None:
            return self._sessions
        return self._controller.effective_sessions(
            self._sched_start_var.get(),
            self._sched_end_var.get(),
            self._am_var.get(),
            self._pm_var.get(),
        )

    def _on_session_param_changed(self) -> None:
        self._sessions = None
        self._update_session_count_label()

    def _update_session_count_label(self) -> None:
        sessions = self._get_effective_sessions()
        if sessions is None:
            self._session_count_label.config(text="invalid dates", fg=CLR_RED)
            self._update_slot_summary(None)
            return
        n_days  = len({d for d, _ in sessions})
        n_slots = len(sessions)
        custom  = self._sessions is not None
        txt = f"{n_days} days  \u2192  {n_slots} slots"
        if custom:
            txt += "  (custom)"
        self._session_count_label.config(
            text=txt,
            fg=CLR_PINK if custom else CLR_ORANGE,
        )
        self._update_slot_summary(n_slots)

    def _update_slot_summary(self, available: int | None) -> None:
        w = self._slot_summary_text
        if w is None:
            return
        w.config(state=tk.NORMAL)
        w.delete("1.0", tk.END)
        if available is None or not self._controller.state.paper_registry:
            w.config(state=tk.DISABLED)
            return
        needed = self._controller.needed_slots_per_grade()
        if not needed:
            w.insert(tk.END, "  (no papers \u2014 load timetable first)\n", "dim")
        for grade in sorted(needed.keys()):
            n     = needed[grade]
            spare = available - n
            if spare >= 0:
                line = (f"  {grade}: {n} needed / {available} available"
                        f"  \u2713 {spare} spare\n")
                tag = "ok"
            else:
                line = (f"  {grade}: {n} needed / {available} available"
                        f"  \u2717 SHORT by {-spare}\n")
                tag = "short"
            w.insert(tk.END, line, tag)
        w.config(state=tk.DISABLED)

    def _open_session_calendar(self) -> None:
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
            custom_set = set(self._sessions)
            for d in day_state:
                for s in SESSIONS:
                    if (d, s) not in custom_set:
                        day_state[d][s] = False

        top = tk.Toplevel(self)
        top.title("Configure Exam Sessions")
        top.geometry("360x480")
        top.configure(bg=CLR_WHITE)

        tk.Label(top, text="Toggle AM/PM sessions on or off:",
                 bg=CLR_WHITE, fg=CLR_HEADER, font=("Calibri", 11, "bold")
                 ).pack(pady=(10, 4), padx=10, anchor=tk.W)

        cf = tk.Frame(top, bg=CLR_WHITE)
        cf.pack(fill=tk.BOTH, expand=True, padx=10)
        canvas = tk.Canvas(cf, bg=CLR_WHITE, highlightthickness=0)
        v_sb = ttk.Scrollbar(cf, orient=tk.VERTICAL, command=canvas.yview)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(yscrollcommand=v_sb.set)
        inner = tk.Frame(canvas, bg=CLR_WHITE)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        session_vars: dict[tuple[date, str], tk.BooleanVar] = {}
        for i, (d, state) in enumerate(sorted(day_state.items())):
            bg = CLR_MORNING if i % 2 == 0 else CLR_WHITE
            row = tk.Frame(inner, bg=bg)
            row.pack(fill=tk.X, pady=1)
            tk.Label(row, text=d.strftime("%a %d %b"), bg=bg,
                     fg="#1E293B", font=("Calibri", 10), width=14,
                     anchor=tk.W).pack(side=tk.LEFT, padx=(4, 8))
            for sess in SESSIONS:
                v = tk.BooleanVar(value=state.get(sess, True))
                session_vars[(d, sess)] = v
                sc = CLR_ORANGE if sess == "AM" else CLR_PINK
                tk.Checkbutton(row, text=sess, variable=v,
                               bg=bg, fg="#1E293B", font=("Calibri", 10),
                               selectcolor=sc).pack(side=tk.LEFT, padx=2)

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
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=16, pady=4
                  ).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=top.destroy,
                  bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Reset to all on", command=_reset_all,
                  bg=CLR_BLUE, fg=CLR_WHITE, font=("Calibri", 10),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT)

    # ------------------------------------------------------------------
    # Slot pinning
    # ------------------------------------------------------------------

    def _open_slot_picker(self, paper, sessions: list[tuple],
                          on_done=None) -> None:
        top = tk.Toplevel(self)
        top.title(f"Pin {paper.label} to slot")
        top.geometry("310x420")
        top.configure(bg=CLR_WHITE)

        tk.Label(top, text=f"Select a slot for  {paper.label}:",
                 bg=CLR_WHITE, fg=CLR_HEADER, font=("Calibri", 11, "bold")
                 ).pack(pady=(10, 4), padx=10, anchor=tk.W)

        lb_frame = tk.Frame(top, bg=CLR_WHITE)
        lb_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        lb_sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL)
        lb_sb.pack(side=tk.RIGHT, fill=tk.Y)
        lb = tk.Listbox(lb_frame, font=("Calibri", 10), relief=tk.SOLID, bd=1,
                        bg=CLR_WHITE,
                        selectbackground=CLR_PINK, selectforeground=CLR_WHITE,
                        selectmode=tk.SINGLE, yscrollcommand=lb_sb.set)
        lb_sb.config(command=lb.yview)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, (d, sess) in enumerate(sessions):
            marker = "  \U0001f4cc" if paper.pinned_slot == i else ""
            bg = CLR_MORNING if sess == "AM" else CLR_AFTERNOON
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
            self._bus.publish(EVT_PAPERS_CHANGED, state=self._controller.state)
            if on_done:
                on_done()
            top.destroy()

        tk.Button(btn_frame, text="Pin to slot", command=_pin,
                  bg=CLR_PINK, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=16, pady=4
                  ).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=top.destroy,
                  bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                  relief=tk.FLAT, padx=10, pady=4
                  ).pack(side=tk.LEFT, padx=8)

    # ------------------------------------------------------------------
    # Schedule popout + export
    # ------------------------------------------------------------------

    def _open_schedule_popout(self) -> None:
        result = self._controller.state.schedule_result
        if not result:
            messagebox.showinfo("No schedule", "Generate a schedule first.")
            return

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
        top.title("Full Exam Schedule \u2014 All Grades")
        top.geometry("980x620")
        top.configure(bg=CLR_WHITE)

        bar = tk.Frame(top, bg=CLR_HEADER, pady=8, padx=10)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="Full Exam Schedule \u2014 All Grades",
                 font=("Calibri", 13, "bold"),
                 bg=CLR_HEADER, fg=CLR_WHITE).pack(side=tk.LEFT)
        tk.Button(bar, text="Save as PDF / TXT",
                  command=lambda: self._save_schedule_export(
                      top, grades, grid, slot_meta, all_slots),
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12
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

        HDR_BG = CLR_GRID_HEADER
        gf.columnconfigure(0, weight=2)
        for ci in range(len(grades)):
            gf.columnconfigure(ci + 1, weight=1)

        tk.Label(gf, text="Slot / Date / Session",
                 font=("Calibri", 10, "bold"),
                 bg=HDR_BG, fg="#1E293B", anchor="w",
                 padx=8, pady=5, relief=tk.RIDGE, bd=1
                 ).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        for ci, grade in enumerate(grades):
            tk.Label(gf, text=grade, font=("Calibri", 10, "bold"),
                     bg=HDR_BG, fg="#1E293B", anchor="center",
                     padx=6, pady=5, relief=tk.RIDGE, bd=1
                     ).grid(row=0, column=ci + 1,
                             sticky="nsew", padx=1, pady=1)

        for ri, slot_idx in enumerate(all_slots):
            d, session = slot_meta[slot_idx]
            row_bg = CLR_MORNING   if session == "AM" else CLR_AFTERNOON
            row_fg = "#7C2D12"     if session == "AM" else "#831843"
            lbl_txt = f"Slot {slot_idx+1}  {d.strftime('%a %d %b')}  {session}"
            tk.Label(gf, text=lbl_txt, font=("Calibri", 10),
                     bg=row_bg, fg=row_fg, anchor="w",
                     padx=8, pady=4, relief=tk.RIDGE, bd=1
                     ).grid(row=ri + 1, column=0,
                             sticky="nsew", padx=1, pady=1)
            for ci, grade in enumerate(grades):
                subjects = sorted(grid[slot_idx][grade])
                if not subjects:
                    tk.Label(gf, text="", font=("Calibri", 10),
                             bg=row_bg, anchor="center",
                             padx=4, pady=4, relief=tk.RIDGE, bd=1
                             ).grid(row=ri + 1, column=ci + 1,
                                    sticky="nsew", padx=1, pady=1)
                else:
                    cell_frame = tk.Frame(gf, bg=row_bg,
                                          relief=tk.RIDGE, bd=1)
                    cell_frame.grid(row=ri + 1, column=ci + 1,
                                    sticky="nsew", padx=1, pady=1)
                    for subj in subjects:
                        lbl = tk.Label(cell_frame, text=subj,
                                       font=("Calibri", 10, "bold"),
                                       bg=row_bg, fg=row_fg, anchor="center",
                                       padx=4, pady=3, cursor="hand2")
                        lbl.pack(fill=tk.X)
                        lbl.bind("<Button-1>",
                                 lambda e, s=subj, g=grade, p=top:
                                 self._navigate_to_exam_subject(s, g, p))

        gf.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _save_schedule_export(self, top, grades, grid, slot_meta,
                               all_slots) -> None:
        path = filedialog.asksaveasfilename(
            parent=top, title="Save exam schedule",
            defaultextension=".pdf",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ])
        if not path:
            return
        try:
            if path.lower().endswith(".txt"):
                self._controller.export_txt(
                    path, grades, grid, slot_meta, all_slots)
            else:
                self._controller.export_pdf(
                    path, grades, grid, slot_meta, all_slots)
            messagebox.showinfo("Saved", f"Schedule saved to:\n{path}",
                                parent=top)
        except ImportError:
            messagebox.showerror(
                "PDF export unavailable",
                "reportlab is not installed.\n\nRun:  pip install reportlab",
                parent=top)
        except Exception as e:
            messagebox.showerror("Export Error", str(e), parent=top)
