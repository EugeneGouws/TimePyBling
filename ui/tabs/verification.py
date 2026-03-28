from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from app.controller import (AppController, EVT_TIMETABLE_LOADED,
                             EVT_STATE_LOADED)
from ui.constants import (CLR_WHITE, CLR_LIGHT, CLR_BLUE, CLR_GREEN,
                           CLR_RED, CLR_ORANGE)
from ui.helpers import _scrolled_text, _write, _clear, student_display


class VerificationTab(tk.Frame):

    def __init__(self, parent, controller: AppController, bus) -> None:
        super().__init__(parent, bg=CLR_WHITE)
        self._controller = controller
        self._bus = bus

        self._build()

        bus.subscribe(EVT_TIMETABLE_LOADED, self._on_data_loaded)
        bus.subscribe(EVT_STATE_LOADED,     self._on_data_loaded)

    # ------------------------------------------------------------------
    # Build
    # ------------------------------------------------------------------

    def _build(self):
        pane = tk.PanedWindow(self, orient=tk.HORIZONTAL,
                              bg="#ccc", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)

        # ── Left: Clashes + Schedulable Pairs ──────────────────────────
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=520)

        hdr = tk.Frame(left, bg=CLR_WHITE)
        hdr.pack(fill=tk.X, padx=8, pady=(8, 2))
        tk.Label(hdr, text="Clashes  &  Schedulable Pairs", bg=CLR_WHITE,
                 font=("Calibri", 12, "bold"), fg="#1E293B").pack(side=tk.LEFT)
        tk.Button(hdr, text="Re-run", command=self._run_verification,
                  bg=CLR_LIGHT, fg="#1E293B", font=("Calibri", 10),
                  relief=tk.FLAT, padx=10).pack(side=tk.RIGHT)

        report_frame = tk.Frame(left, bg=CLR_WHITE)
        report_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self._clash_report = _scrolled_text(report_frame)
        self._clash_report.tag_config("pass",    foreground=CLR_GREEN,
                                      font=("Calibri", 10, "bold"))
        self._clash_report.tag_config("fail",    foreground=CLR_RED,
                                      font=("Calibri", 10))
        self._clash_report.tag_config("heading", foreground=CLR_BLUE,
                                      font=("Calibri", 11, "bold"))
        self._clash_report.tag_config("warn",    foreground=CLR_ORANGE,
                                      font=("Calibri", 10))
        self._clash_report.tag_config("dim",     foreground="#6B7280",
                                      font=("Calibri", 10))

        # ── Right: Data Integrity ───────────────────────────────────────
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=300)

        di_hdr = tk.Frame(right, bg=CLR_WHITE)
        di_hdr.pack(fill=tk.X, padx=8, pady=(8, 2))
        tk.Label(di_hdr, text="Data Integrity", bg=CLR_WHITE,
                 font=("Calibri", 12, "bold"), fg="#1E293B").pack(side=tk.LEFT)

        di_frame = tk.Frame(right, bg=CLR_WHITE)
        di_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self._integrity_report = _scrolled_text(di_frame)
        self._integrity_report.tag_config("pass",    foreground=CLR_GREEN,
                                           font=("Calibri", 10, "bold"))
        self._integrity_report.tag_config("heading", foreground=CLR_BLUE,
                                           font=("Calibri", 11, "bold"))
        self._integrity_report.tag_config("warn",    foreground=CLR_ORANGE,
                                           font=("Calibri", 10))
        self._integrity_report.tag_config("dim",     foreground="#6B7280",
                                           font=("Calibri", 10))

    # ------------------------------------------------------------------
    # Event callbacks
    # ------------------------------------------------------------------

    def _on_data_loaded(self, **_) -> None:
        self._run_verification()

    # ------------------------------------------------------------------
    # Verification logic
    # ------------------------------------------------------------------

    def _run_verification(self) -> None:
        if self._controller.state.timetable_tree is None:
            return
        self._update_clash_report()
        self._update_integrity_panel()

    def _update_clash_report(self) -> None:
        """Left panel: Clashes + Schedulable Pairs."""
        w = self._clash_report
        _clear(w)

        # ── CLASHES ──
        _write(w, "CLASHES\n", "heading")

        student_clashes, teacher_clashes = self._controller.find_clashes()
        is_legal = not student_clashes and not teacher_clashes

        if is_legal:
            _write(w, "PASS ✓  —  no student or teacher clashes found\n", "pass")
        else:
            _write(w, f"FAIL ✗  —  {len(student_clashes)} student clash(es), "
                      f"{len(teacher_clashes)} teacher clash(es)\n", "fail")

        if student_clashes:
            _write(w, "\nSTUDENT DOUBLE-BOOKINGS\n", "heading")
            by_sb: dict[str, list] = {}
            for c in student_clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, f"\n  Subblock {sb}\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["student"]):
                    _write(w, f"    {self._student_display(entry['student']):<30}:  "
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
        _write(w, f"\nTotal violations: {total}\n",
               "pass" if total == 0 else "fail")

        # ── SCHEDULABLE PAIRS ──
        _write(w, "\nSCHEDULABLE PAIRS  (share no students — can sit same slot)\n",
               "heading")

        if self._controller.state.exam_tree is None:
            _write(w, "(Load timetable to see schedulable pairs)\n", "dim")
        else:
            pairs_by_grade = self._controller.schedulable_pairs()
            if not pairs_by_grade:
                _write(w, "  No free pairs — all subjects share students.\n", "dim")
            else:
                for grade_label, pairs in pairs_by_grade.items():
                    count = len(pairs)
                    label = "pair" if count == 1 else "pairs"
                    _write(w, f"\n  {grade_label}  {count} {label}:\n", "heading")
                    for a, b in pairs:
                        _write(w, f"    {a}  +  {b}\n", "pass")

    def _update_integrity_panel(self) -> None:
        """Right panel: Data Integrity — classes with fewer than 5 students."""
        w = self._integrity_report
        _clear(w)
        _write(w, "Classes with fewer than 5 students:\n", "dim")
        issues = self._controller.data_integrity_issues()
        if not issues:
            _write(w, "PASS ✓  —  all classes have 5 or more students\n", "pass")
        else:
            _write(w, f"WARN ⚠  —  {len(issues)} class(es) flagged:\n", "warn")
            for info in issues:
                _write(w, f"\n  {info['label']}\n", "heading")
                _write(w, f"    Count:     {info['count']}\n", "warn")
                _write(w, f"    Subblocks: {info['subblocks']}\n", "dim")
                student_displays = [self._student_display(s) for s in info["students"]]
                _write(w, f"    Students:  {student_displays}\n", "dim")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _student_display(self, student_id: int) -> str:
        return student_display(self._controller.state.timetable_tree, student_id)
