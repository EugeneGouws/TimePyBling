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
    DEFAULT_EXAM_START, DEFAULT_EXAM_END,
)
from ui.helpers import _scrolled_text, _write, _clear


class ExamTab(tk.Frame):

    def __init__(self, parent, controller: AppController, bus) -> None:
        super().__init__(parent, bg=CLR_WHITE)
        self._controller = controller
        self._bus = bus

        # Custom session list (None = derive from date range + AM/PM)
        self._sessions: list[tuple] | None = None
        self._selected_paper_label: str | None = None

        self._build()

        bus.subscribe(EVT_REGISTRY_BUILT,    self._on_exam_rebuilt)
        bus.subscribe(EVT_TIMETABLE_LOADED,  self._on_exam_rebuilt)
        bus.subscribe(EVT_PAPERS_CHANGED,    self._on_papers_changed)
        bus.subscribe(EVT_STATE_LOADED,      self._on_state_loaded)
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
        pane.add(left, minsize=440)

        # ── header row ──────────────────────────────────────────────
        left_top = tk.Frame(left, bg=CLR_WHITE)
        left_top.pack(fill=tk.X, padx=6, pady=(6, 0))
        tk.Label(left_top, text="Exam Tree", bg=CLR_WHITE,
                 font=("Calibri", 12, "bold"), fg="#1E293B").pack(side=tk.LEFT)
        tk.Button(left_top, text="Save State…",
                  command=self._export_exam_state,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.RIGHT, padx=(2, 0))
        tk.Button(left_top, text="Load State…",
                  command=self._import_exam_state,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.RIGHT, padx=2)
        tk.Button(left_top, text="Rebuild",
                  command=self._rebuild_exam,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=10
                  ).pack(side=tk.RIGHT, padx=2)

        # ── exam tree ───────────────────────────────────────────────
        tree_frame = tk.Frame(left, bg=CLR_WHITE)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(2, 0))
        self._ex_tree = ttk.Treeview(tree_frame, show="tree",
                                     selectmode="extended")
        ex_sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                               command=self._ex_tree.yview)
        ex_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._ex_tree.configure(yscrollcommand=ex_sb.set)
        self._ex_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._ex_tree.tag_configure("excluded", foreground="#9CA3AF")
        self._ex_tree.tag_configure("subject",  foreground="#1E293B")
        self._ex_tree.bind("<<TreeviewSelect>>", self._on_exam_tree_select)

        # ── paper panel ─────────────────────────────────────────────
        self._paper_lf = tk.LabelFrame(left, text="Papers for selected subject",
                                       bg=CLR_WHITE, font=("Calibri", 10, "bold"),
                                       fg=CLR_BLUE, padx=6, pady=6)
        self._paper_lf.pack(fill=tk.X, padx=6, pady=6)

        self._paper_listbox = tk.Listbox(
            self._paper_lf, height=4, font=("Calibri", 10),
            relief=tk.SOLID, bd=1, bg=CLR_WHITE,
            selectbackground=CLR_ORANGE, selectforeground=CLR_WHITE,
            selectmode=tk.SINGLE)
        self._paper_listbox.pack(fill=tk.X, pady=(0, 4))
        self._paper_listbox.bind("<<ListboxSelect>>", self._on_paper_select)

        paper_btn_row = tk.Frame(self._paper_lf, bg=CLR_WHITE)
        paper_btn_row.pack(fill=tk.X, pady=(0, 4))
        tk.Button(paper_btn_row, text="+ Add Paper",
                  command=self._add_paper,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=8).pack(side=tk.LEFT)
        tk.Button(paper_btn_row, text="− Remove",
                  command=self._remove_paper,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.LEFT, padx=4)
        tk.Button(paper_btn_row, text="📌 Pin…",
                  command=self._pin_paper,
                  bg=CLR_PINK, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.LEFT)
        tk.Button(paper_btn_row, text="Unpin",
                  command=self._unpin_paper,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10), padx=8).pack(side=tk.LEFT, padx=4)

        constr_row = tk.Frame(self._paper_lf, bg=CLR_WHITE)
        constr_row.pack(fill=tk.X)
        tk.Label(constr_row, text="Constraint:", bg=CLR_WHITE,
                 fg="#1E293B", font=("Calibri", 10)).pack(side=tk.LEFT)
        self._constr_entry = tk.Entry(constr_row, font=("Calibri", 10),
                                      relief=tk.SOLID, bd=1, width=6)
        self._constr_entry.pack(side=tk.LEFT, padx=2)
        self._constr_entry.bind("<Return>", lambda e: self._add_constraint())
        self._constr_add_btn = tk.Button(
            constr_row, text="Add", command=self._add_constraint,
            bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Calibri", 10), padx=6)
        self._constr_add_btn.pack(side=tk.LEFT)
        self._constr_remove_btn = tk.Button(
            constr_row, text="Remove", command=self._remove_constraint,
            bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
            font=("Calibri", 10), padx=6)
        self._constr_remove_btn.pack(side=tk.LEFT, padx=2)

        self._constr_listbox = tk.Listbox(
            self._paper_lf, height=2, font=("Calibri", 10),
            relief=tk.SOLID, bd=1, bg=CLR_WHITE,
            selectbackground=CLR_PINK, selectforeground=CLR_WHITE,
            selectmode=tk.SINGLE)
        self._constr_listbox.pack(fill=tk.X, pady=(4, 0))

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

        # Row 0: Start / End + live slot count
        tk.Label(ctrl, text="Start:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=0, column=0, sticky=tk.W)
        self._sched_start_var = tk.StringVar(
            value=DEFAULT_EXAM_START.strftime("%Y-%m-%d"))
        self._start_entry = tk.Entry(
            ctrl, textvariable=self._sched_start_var,
            font=("Calibri", 10), relief=tk.SOLID, bd=1, width=11)
        self._start_entry.grid(row=0, column=1, sticky=tk.W, padx=(2, 8))
        self._start_entry.bind("<FocusOut>",
                               lambda e: self._on_session_param_changed())
        self._start_entry.bind("<Return>",
                               lambda e: self._on_session_param_changed())

        tk.Label(ctrl, text="End:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=0, column=2, sticky=tk.W)
        self._sched_end_var = tk.StringVar(
            value=DEFAULT_EXAM_END.strftime("%Y-%m-%d"))
        self._end_entry = tk.Entry(
            ctrl, textvariable=self._sched_end_var,
            font=("Calibri", 10), relief=tk.SOLID, bd=1, width=11)
        self._end_entry.grid(row=0, column=3, sticky=tk.W, padx=(2, 8))
        self._end_entry.bind("<FocusOut>",
                             lambda e: self._on_session_param_changed())
        self._end_entry.bind("<Return>",
                             lambda e: self._on_session_param_changed())

        self._session_count_label = tk.Label(
            ctrl, text="", bg=CLR_WHITE, fg=CLR_ORANGE,
            font=("Calibri", 10, "bold"))
        self._session_count_label.grid(row=0, column=4, sticky=tk.W, padx=(0, 4))

        # Row 1: AM / PM checkboxes + Configure sessions
        self._am_var = tk.BooleanVar(value=True)
        self._pm_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="AM", variable=self._am_var,
                       bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                       selectcolor=CLR_ORANGE,
                       command=self._on_session_param_changed
                       ).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        tk.Checkbutton(ctrl, text="PM", variable=self._pm_var,
                       bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                       selectcolor=CLR_PINK,
                       command=self._on_session_param_changed
                       ).grid(row=1, column=1, sticky=tk.W, pady=(4, 0))
        tk.Button(ctrl, text="Configure sessions…",
                  command=self._open_session_calendar,
                  bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                  relief=tk.FLAT, padx=10
                  ).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=(4, 0))

        # Row 2: Grade filter + Generate
        tk.Label(ctrl, text="Grade:", bg=CLR_WHITE, fg="#1E293B",
                 font=("Calibri", 10)).grid(row=2, column=0, sticky=tk.W,
                                             pady=(4, 0))
        self._sched_grade_var = tk.StringVar()
        self._sched_grade_cb = ttk.Combobox(
            ctrl, textvariable=self._sched_grade_var,
            state="readonly", width=10, font=("Calibri", 10))
        self._sched_grade_cb.grid(row=2, column=1, columnspan=2, sticky=tk.W,
                                   padx=(4, 12), pady=(4, 0))
        self._sched_grade_cb.bind("<<ComboboxSelected>>",
                                   lambda e: self._render_schedule())
        tk.Button(ctrl, text="Generate Schedule",
                  command=self._generate_exam_schedule,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12, pady=4
                  ).grid(row=2, column=3, sticky=tk.W, padx=(4, 0), pady=(4, 0))
        tk.Button(ctrl, text="Export / View All",
                  command=self._open_schedule_popout,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12, pady=4
                  ).grid(row=3, column=3, sticky=tk.W, padx=(4, 0), pady=(4, 0))

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
        self._update_session_count_label()

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
        cost_lf = tk.LabelFrame(parent, text="Exam Cost Function",
                                 bg=CLR_WHITE, font=("Calibri", 10, "bold"),
                                 fg=CLR_BLUE, padx=8, pady=6)
        cost_lf.pack(fill=tk.X, padx=8, pady=(0, 8))

        wf = tk.Frame(cost_lf, bg=CLR_WHITE)
        wf.pack(fill=tk.X)
        weight_defs = [("W_day", 5), ("W_week", 1),
                       ("W_consec", 50), ("W_marking", 20)]
        self._cost_weight_vars: list[tk.StringVar] = []
        for i, (lbl, dflt) in enumerate(weight_defs):
            r, c = i // 2, (i % 2) * 3
            tk.Label(wf, text=f"{lbl}:", bg=CLR_WHITE, fg="#1E293B",
                     font=("Calibri", 10)).grid(row=r, column=c, sticky=tk.W,
                                                 padx=(0, 2), pady=2)
            v = tk.StringVar(value=str(dflt))
            self._cost_weight_vars.append(v)
            tk.Entry(wf, textvariable=v, font=("Calibri", 10),
                     relief=tk.SOLID, bd=1, width=5
                     ).grid(row=r, column=c + 1, sticky=tk.W,
                             padx=(0, 12), pady=2)

        calc_row = tk.Frame(cost_lf, bg=CLR_WHITE)
        calc_row.pack(fill=tk.X, pady=(4, 0))
        tk.Button(calc_row, text="Calculate",
                  command=self._calculate_exam_cost,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=10).pack(side=tk.LEFT)
        self._exam_cost_result_label = tk.Label(
            calc_row, text="", bg=CLR_WHITE, fg="#1E293B",
            font=("Calibri", 10), justify=tk.LEFT)
        self._exam_cost_result_label.pack(side=tk.LEFT, padx=6)

    # ------------------------------------------------------------------
    # Event callbacks
    # ------------------------------------------------------------------

    def _on_exam_rebuilt(self, **_) -> None:
        self._populate_exam_tree()
        self._update_sched_grade_list()
        self._update_session_count_label()

    def _on_papers_changed(self, **_) -> None:
        expanded, selected_pairs = self._exam_tree_get_state()
        self._populate_exam_tree()
        self._exam_tree_restore_state(expanded, selected_pairs)
        subject, grade = self._current_subject_grade()
        self._refresh_paper_panel(subject, grade)

    def _on_state_loaded(self, state, **_) -> None:
        # Restore session config vars from state
        cfg = state.session_config
        if cfg is not None:
            self._sched_start_var.set(cfg.start.strftime("%Y-%m-%d"))
            self._sched_end_var.set(cfg.end.strftime("%Y-%m-%d"))
            self._am_var.set(cfg.am)
            self._pm_var.set(cfg.pm)
            self._sessions = None
        self._refresh_exclusion_listbox()
        self._populate_exam_tree()
        self._update_sched_grade_list()
        self._update_session_count_label()

    def _on_schedule_generated(self, **_) -> None:
        self._render_schedule()

    # ------------------------------------------------------------------
    # Exam tree
    # ------------------------------------------------------------------

    def _populate_exam_tree(self) -> None:
        """Populate from exam_tree; excluded subjects shown with [excl] suffix."""
        self._ex_tree.delete(*self._ex_tree.get_children())
        exam_tree = self._controller.state.exam_tree
        if not exam_tree:
            return
        registry = self._controller.state.paper_registry
        exclusions = self._controller.state.exclusions
        from reader.exam_clash import is_excluded as _excl  # noqa: PLC0415
        for grade_label in sorted(exam_tree.grades.keys()):
            grade_node = exam_tree.grades[grade_label]
            grade_ui = self._ex_tree.insert("", tk.END,
                                            text=grade_label, open=False)
            for subj_label in sorted(grade_node.exam_subjects.keys()):
                exam_subject = grade_node.exam_subjects[subj_label]
                if _excl(subj_label, exclusions):
                    self._ex_tree.insert(grade_ui, tk.END,
                                         text=f"{subj_label}  [excl]",
                                         tags=("excluded",), values=())
                else:
                    subj_code = subj_label.split("_")[0]
                    if registry:
                        papers = registry.papers_for_subject_grade(
                            subj_code, grade_label)
                        count = (max(p.student_count() for p in papers)
                                 if papers
                                 else len(exam_subject.all_students()))
                    else:
                        count = len(exam_subject.all_students())
                    self._ex_tree.insert(grade_ui, tk.END,
                                         text=f"{subj_code}  ({count} students)",
                                         tags=("subject",),
                                         values=(subj_code, grade_label))

    def _exam_tree_get_state(self) -> tuple[set[str], list[tuple[str, str]]]:
        """Return (expanded_grade_texts, selected_(subject,grade)_pairs)."""
        expanded: set[str] = set()
        for item in self._ex_tree.get_children():
            if self._ex_tree.item(item, "open"):
                expanded.add(self._ex_tree.item(item, "text"))
        selected_pairs: list[tuple[str, str]] = []
        for item in self._ex_tree.selection():
            tags = self._ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self._ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    selected_pairs.append((vals[0], vals[1]))
        return expanded, selected_pairs

    def _exam_tree_restore_state(self, expanded: set[str],
                                  selected_pairs: list[tuple[str, str]]) -> None:
        """Re-expand grade nodes and re-select subject nodes after a rebuild."""
        for item in self._ex_tree.get_children():
            raw = self._ex_tree.item(item, "text")
            if raw in expanded:
                self._ex_tree.item(item, open=True)
        if not selected_pairs:
            return
        pair_set = set(selected_pairs)
        to_select = []
        for grade_item in self._ex_tree.get_children():
            for subj_item in self._ex_tree.get_children(grade_item):
                tags = self._ex_tree.item(subj_item, "tags")
                if "subject" not in (tags or []):
                    continue
                vals = self._ex_tree.item(subj_item, "values")
                if vals and len(vals) >= 2 and (vals[0], vals[1]) in pair_set:
                    to_select.append(subj_item)
        if to_select:
            self._ex_tree.selection_set(to_select)
            self._ex_tree.see(to_select[0])

    def _navigate_to_exam_subject(self, subject: str, grade: str,
                                   popout: tk.Toplevel | None = None) -> None:
        # Note: notebook.select() is intentionally omitted — this is always
        # called from a popout opened while the exam tab is already visible.
        if popout:
            popout.destroy()
        for grade_item in self._ex_tree.get_children():
            if self._ex_tree.item(grade_item, "text") != grade:
                continue
            self._ex_tree.item(grade_item, open=True)
            for subj_item in self._ex_tree.get_children(grade_item):
                tags = self._ex_tree.item(subj_item, "tags")
                if "subject" not in (tags or []):
                    continue
                vals = self._ex_tree.item(subj_item, "values")
                if vals and vals[0] == subject and vals[1] == grade:
                    self._ex_tree.selection_set(subj_item)
                    self._ex_tree.see(subj_item)
                    self._on_exam_tree_select()
                    return

    def _on_exam_tree_select(self, event=None) -> None:
        sel = self._ex_tree.selection()
        if not sel:
            self._refresh_paper_panel(None, None)
            return
        subject_pairs = []
        for item in sel:
            tags = self._ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self._ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if not subject_pairs:
            self._refresh_paper_panel(None, None)
        elif len(subject_pairs) == 1:
            self._refresh_paper_panel(subject_pairs[0][0], subject_pairs[0][1])
        else:
            self._refresh_paper_panel_multi(subject_pairs)

    # ------------------------------------------------------------------
    # Paper panel
    # ------------------------------------------------------------------

    def _refresh_paper_panel(self, subject: str | None,
                              grade: str | None) -> None:
        self._paper_listbox.delete(0, tk.END)
        self._constr_listbox.delete(0, tk.END)
        self._selected_paper_label = None
        self._paper_lf.config(text="Papers for selected subject")
        self._set_constraint_ui_enabled(True)
        registry = self._controller.state.paper_registry
        if not subject or not grade or not registry:
            return
        papers = registry.papers_for_subject_grade(subject, grade)
        result = self._controller.state.schedule_result
        pin_clashes = result.pin_clash_warnings if result else {}
        for p in papers:
            constr_str = (", ".join(sorted(p.constraints))
                          if p.constraints else "no constraints")
            pin_indicator = "📌 " if p.pinned_slot is not None else "   "
            clash_flag = "  ⚠ pin clash" if p.label in pin_clashes else ""
            self._paper_listbox.insert(
                tk.END,
                f"{pin_indicator}{p.subject} P{p.paper_number}"
                f" — {p.student_count()} students — {constr_str}{clash_flag}"
            )
        if papers:
            self._paper_listbox.selection_set(0)
            self._selected_paper_label = papers[0].label
            self._refresh_constraint_list(papers[0])

    def _refresh_paper_panel_multi(
            self, subject_grade_pairs: list[tuple[str, str]]) -> None:
        self._paper_listbox.delete(0, tk.END)
        self._constr_listbox.delete(0, tk.END)
        self._selected_paper_label = None
        n = len(subject_grade_pairs)
        self._paper_lf.config(text=f"{n} subjects selected")
        self._paper_listbox.insert(tk.END, f"{n} subjects selected")
        registry = self._controller.state.paper_registry
        if registry:
            addable = sum(
                1 for subj, grade in subject_grade_pairs
                if registry.papers_for_subject_grade(subj, grade)
                and max(p.paper_number
                        for p in registry.papers_for_subject_grade(subj, grade)
                        ) < 3
            )
            if addable:
                self._paper_listbox.insert(
                    tk.END, f"  {addable} eligible for [+ Add Paper]")
        self._set_constraint_ui_enabled(False)
        self._constr_listbox.insert(
            tk.END, "Select a single paper to edit constraints")

    def _set_constraint_ui_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self._constr_entry.config(state=state)
        self._constr_add_btn.config(state=state)
        self._constr_remove_btn.config(state=state)

    def _on_paper_select(self, event=None) -> None:
        sel = self._paper_listbox.curselection()
        if not sel:
            return
        registry = self._controller.state.paper_registry
        if not registry:
            return
        tree_sel = self._ex_tree.selection()
        if not tree_sel:
            return
        subject_pairs = []
        for item in tree_sel:
            tags = self._ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self._ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if len(subject_pairs) != 1:
            return
        subject, grade = subject_pairs[0]
        papers = registry.papers_for_subject_grade(subject, grade)
        idx = sel[0]
        if idx < len(papers):
            self._selected_paper_label = papers[idx].label
            self._refresh_constraint_list(papers[idx])

    def _refresh_constraint_list(self, paper) -> None:
        self._constr_listbox.delete(0, tk.END)
        for code in sorted(paper.constraints):
            self._constr_listbox.insert(tk.END, code)

    # ------------------------------------------------------------------
    # Rebuild
    # ------------------------------------------------------------------

    def _rebuild_exam(self) -> None:
        if self._controller.state.timetable_tree is None:
            return
        self._controller.build_exam_tree()
        self._controller.build_registry()
        self._controller.state.schedule_result = None
        self._selected_paper_label = None

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
        self._on_exclusion_change(new_excl)

    def _remove_exclusion(self) -> None:
        sel = self._excl_listbox.curselection()
        if not sel:
            return
        code = self._excl_listbox.get(sel[0])
        new_excl = set(self._controller.state.exclusions)
        new_excl.discard(code)
        self._on_exclusion_change(new_excl)

    def _on_exclusion_change(self, new_excl: set[str]) -> None:
        expanded, selected_pairs = self._exam_tree_get_state()
        self._controller.set_exclusions(new_excl)
        # set_exclusions triggers EVT_EXCLUSIONS_CHANGED → EVT_REGISTRY_BUILT
        # which fires _on_exam_rebuilt; tree restore must happen after that.
        self._exam_tree_restore_state(expanded, selected_pairs)

    # ------------------------------------------------------------------
    # Paper / constraint actions
    # ------------------------------------------------------------------

    def _current_subject_grade(self) -> tuple[str, str] | tuple[None, None]:
        tree_sel = self._ex_tree.selection()
        if not tree_sel:
            return None, None
        vals = self._ex_tree.item(tree_sel[0], "values")
        if not vals or len(vals) < 2:
            return None, None
        return vals[0], vals[1]

    def _add_paper(self) -> None:
        if not self._controller.state.paper_registry:
            return
        sel = self._ex_tree.selection()
        subject_pairs = []
        for item in sel:
            tags = self._ex_tree.item(item, "tags")
            if "subject" in (tags or []):
                vals = self._ex_tree.item(item, "values")
                if vals and len(vals) >= 2:
                    subject_pairs.append((vals[0], vals[1]))
        if not subject_pairs:
            return
        expanded, selected_pairs = self._exam_tree_get_state()
        if len(subject_pairs) == 1:
            subject, grade = subject_pairs[0]
            result = self._controller.add_paper(subject, grade)
            if result is None:
                messagebox.showinfo("Cannot add",
                                    "Maximum 3 papers per subject per grade.")
                return
        else:
            registry = self._controller.state.paper_registry
            added = sum(
                1 for subj, grade in subject_pairs
                if registry.papers_for_subject_grade(subj, grade) is not None
                and self._controller.add_paper(subj, grade) is not None
            )
            if added == 0:
                messagebox.showinfo("Cannot add",
                                    "All selected subjects already have 3 papers.")
                return
        self._exam_tree_restore_state(expanded, selected_pairs)
        self._on_exam_tree_select()

    def _remove_paper(self) -> None:
        if not self._controller.state.paper_registry or \
                not self._selected_paper_label:
            return
        parts = self._selected_paper_label.split("_")
        subject = parts[0] if parts else None
        grade   = parts[-1] if len(parts) >= 3 else None
        expanded, selected_pairs = self._exam_tree_get_state()
        removed = self._controller.remove_paper(self._selected_paper_label)
        if not removed:
            messagebox.showinfo("Cannot remove",
                                "Cannot remove the only paper (P1) for a subject.")
            return
        self._exam_tree_restore_state(expanded, selected_pairs)
        self._refresh_paper_panel(subject, grade)

    def _add_constraint(self) -> None:
        if not self._controller.state.paper_registry or \
                not self._selected_paper_label:
            return
        code = self._constr_entry.get().strip().upper()
        if not code:
            return
        self._controller.add_constraint(self._selected_paper_label, code)
        self._constr_entry.delete(0, tk.END)
        paper = self._controller.state.paper_registry.get(
            self._selected_paper_label)
        if paper:
            self._refresh_constraint_list(paper)
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

    def _remove_constraint(self) -> None:
        if not self._controller.state.paper_registry or \
                not self._selected_paper_label:
            return
        sel = self._constr_listbox.curselection()
        if not sel:
            return
        code = self._constr_listbox.get(sel[0])
        self._controller.remove_constraint(self._selected_paper_label, code)
        paper = self._controller.state.paper_registry.get(
            self._selected_paper_label)
        if paper:
            self._refresh_constraint_list(paper)
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

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
                start = date.fromisoformat(
                    self._sched_start_var.get().strip())
                end = date.fromisoformat(self._sched_end_var.get().strip())
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
        try:
            w_day  = int(float(self._cost_weight_vars[0].get()))
            w_week = int(float(self._cost_weight_vars[1].get()))
        except (ValueError, AttributeError):
            w_day, w_week = 5, 1
        self._controller.generate_schedule(
            sessions=sessions,
        )
        # EVT_SCHEDULE_GENERATED → _on_schedule_generated → _render_schedule

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

        if result.teacher_warnings:
            _write(w, "\nTeacher marking load conflicts:\n", "warn")
            for msg in result.teacher_warnings:
                _write(w, f"  ⚠  {msg}\n", "warn")

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
        """Date entry or AM/PM checkbox changed — reset any custom session list."""
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
        txt = f"{n_days} days  →  {n_slots} slots"
        if custom:
            txt += "  (custom)"
        self._session_count_label.config(
            text=txt,
            fg=CLR_PINK if custom else CLR_ORANGE,
        )
        self._update_slot_summary(n_slots)

    def _update_slot_summary(self, available: int | None) -> None:
        w = self._slot_summary_text
        w.config(state=tk.NORMAL)
        w.delete("1.0", tk.END)
        if available is None or not self._controller.state.paper_registry:
            w.config(state=tk.DISABLED)
            return
        needed = self._controller.needed_slots_per_grade()
        if not needed:
            w.insert(tk.END, "  (no papers — load timetable first)\n", "dim")
        for grade in sorted(needed.keys()):
            n     = needed[grade]
            spare = available - n
            if spare >= 0:
                line = (f"  {grade}: {n} needed / {available} available"
                        f"  ✓ {spare} spare\n")
                tag = "ok"
            else:
                line = (f"  {grade}: {n} needed / {available} available"
                        f"  ✗ SHORT by {-spare}\n")
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

    def _pin_paper(self) -> None:
        registry = self._controller.state.paper_registry
        if not registry or not self._selected_paper_label:
            return
        paper = registry.get(self._selected_paper_label)
        if not paper:
            return
        sessions = self._get_effective_sessions()
        if not sessions:
            messagebox.showinfo(
                "No sessions",
                "Configure exam sessions (start/end dates) first.")
            return
        self._open_slot_picker(paper, sessions)

    def _unpin_paper(self) -> None:
        registry = self._controller.state.paper_registry
        if not registry or not self._selected_paper_label:
            return
        paper = registry.get(self._selected_paper_label)
        # pin_slot is mutated directly on the ExamPaper domain object;
        # pin state is paper-local and has no registry-level operation.
        if paper and paper.pinned_slot is not None:
            paper.pinned_slot = None
            subject, grade = self._current_subject_grade()
            self._refresh_paper_panel(subject, grade)

    def _open_slot_picker(self, paper, sessions: list[tuple]) -> None:
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
        top.title("Full Exam Schedule — All Grades")
        top.geometry("980x620")
        top.configure(bg=CLR_WHITE)

        bar = tk.Frame(top, bg=CLR_HEADER, pady=8, padx=10)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="Full Exam Schedule — All Grades",
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

    # ------------------------------------------------------------------
    # Exam cost function
    # ------------------------------------------------------------------

    def _calculate_exam_cost(self) -> None:
        result = self._controller.state.schedule_result
        if not result:
            self._exam_cost_result_label.config(
                text="Generate schedule first.")
            return
        try:
            w_day     = float(self._cost_weight_vars[0].get())
            w_week    = float(self._cost_weight_vars[1].get())
            w_consec  = float(self._cost_weight_vars[2].get())
            w_marking = float(self._cost_weight_vars[3].get())
        except ValueError:
            self._exam_cost_result_label.config(text="Invalid weight values.")
            return

        t_student = result.student_cost

        consec_count = 0
        by_sg: dict[tuple[str, str], list] = defaultdict(list)
        for sp in result.scheduled:
            by_sg[(sp.paper.subject, sp.paper.grade)].append(sp.slot_index)
        for slots_list in by_sg.values():
            sorted_s = sorted(slots_list)
            for i in range(len(sorted_s) - 1):
                if sorted_s[i + 1] - sorted_s[i] == 1:
                    consec_count += 1
        t_consec = w_consec * consec_count

        marking_count = len(result.teacher_warnings)
        t_marking = w_marking * marking_count

        total = t_student + t_consec + t_marking
        self._exam_cost_result_label.config(
            text=(f"E = {total:.0f}\n"
                  f"  student (W_day={w_day:.0f}, W_week={w_week:.0f})"
                  f" = {t_student}\n"
                  f"  consec {w_consec:.0f}×{consec_count}"
                  f" = {t_consec:.0f}\n"
                  f"  marking {w_marking:.0f}×{marking_count}"
                  f" = {t_marking:.0f}")
        )
