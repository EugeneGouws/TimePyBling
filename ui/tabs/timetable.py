from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from app.controller import (AppController, EVT_TIMETABLE_LOADED,
                             EVT_STATE_LOADED)
from ui.constants import (TIMETABLE_GRID, CLR_WHITE, CLR_GRID_CELL,
                           CLR_GRID_HEADER, CLR_GRID_ACTIVE,
                           CLR_BLUE, CLR_ORANGE)
from ui.helpers import student_display as _student_display_fn


class TimetableTab(tk.Frame):

    def __init__(self, parent, controller: AppController, bus) -> None:
        super().__init__(parent, bg=CLR_WHITE)
        self._controller = controller
        self._bus = bus
        self._subblock_popup: tk.Toplevel | None = None

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

        # ── LEFT — main 8×7 rotation grid ─────────────────────────────
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
        self._grid_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        for c in range(8):
            self._grid_frame.columnconfigure(c, weight=1)

        tk.Label(self._grid_frame, text="", bg=CLR_GRID_HEADER,
                 relief=tk.RIDGE, bd=1, padx=8, pady=4
                 ).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        for col in range(7):
            tk.Label(self._grid_frame, text=f"P{col+1}",
                     bg=CLR_GRID_HEADER, font=("Calibri", 10, "bold"),
                     fg="#1E293B", relief=tk.RIDGE, bd=1, padx=6, pady=4
                     ).grid(row=0, column=col+1, sticky="nsew", padx=1, pady=1)

        self._grid_cells = []
        for day in range(8):
            tk.Label(self._grid_frame, text=f"D{day+1}",
                     bg=CLR_GRID_HEADER, font=("Calibri", 10, "bold"),
                     fg="#1E293B", relief=tk.RIDGE, bd=1, padx=6, pady=4
                     ).grid(row=day+1, column=0, sticky="nsew", padx=1, pady=1)
            row_cells = []
            for col in range(7):
                subblock = TIMETABLE_GRID[day][col]
                btn = tk.Button(
                    self._grid_frame,
                    text=subblock,
                    font=("Calibri", 10, "bold"),
                    bg=CLR_GRID_CELL, fg="#1E293B",
                    relief=tk.RIDGE, bd=1, padx=10, pady=6,
                    command=lambda sb=subblock: self._show_subblock_detail(sb),
                )
                btn.grid(row=day+1, column=col+1, sticky="nsew", padx=1, pady=1)
                row_cells.append(btn)
            self._grid_cells.append(row_cells)

        # ── RIGHT — entity timetable viewer ───────────────────────────
        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=420)

        sel = tk.Frame(right, bg=CLR_WHITE, padx=8, pady=6)
        sel.pack(fill=tk.X)

        tk.Label(sel, text="View:", bg=CLR_WHITE,
                 font=("Calibri", 11), fg="#1E293B").pack(side=tk.LEFT)

        self._entity_type_var = tk.StringVar(value="Student")
        entity_type_cb = ttk.Combobox(sel, textvariable=self._entity_type_var,
                                      values=["Student", "Teacher", "Subject"],
                                      state="readonly", width=10,
                                      font=("Calibri", 11))
        entity_type_cb.pack(side=tk.LEFT, padx=(4, 8))
        entity_type_cb.bind("<<ComboboxSelected>>", self._on_entity_type_change)

        self._entity_value_var = tk.StringVar()
        self._entity_search_entry = tk.Entry(sel,
                                             textvariable=self._entity_value_var,
                                             width=22, font=("Calibri", 11),
                                             relief=tk.SOLID, bd=1)
        self._entity_search_entry.pack(side=tk.LEFT, padx=(0, 6))
        self._entity_value_var.trace_add("write", self._on_entity_search_change)

        tk.Button(sel, text="View Timetable",
                  command=self._on_view_timetable,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 11, "bold"), padx=12, pady=3
                  ).pack(side=tk.LEFT)

        list_frame = tk.Frame(right, bg=CLR_WHITE)
        list_frame.pack(fill=tk.X, padx=8, pady=(0, 4))

        self._entity_listbox = tk.Listbox(list_frame, height=5,
                                           font=("Calibri", 11),
                                           relief=tk.SOLID, bd=1,
                                           bg=CLR_WHITE,
                                           selectbackground=CLR_ORANGE,
                                           selectforeground=CLR_WHITE,
                                           selectmode=tk.SINGLE)
        list_sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                command=self._entity_listbox.yview)
        list_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._entity_listbox.config(yscrollcommand=list_sb.set)
        self._entity_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._entity_listbox.bind("<<ListboxSelect>>",
                                  self._on_entity_listbox_select)

        self._entity_heading_var = tk.StringVar()
        tk.Label(right, textvariable=self._entity_heading_var,
                 bg=CLR_WHITE, font=("Calibri", 12, "bold"),
                 fg=CLR_BLUE).pack(padx=8, pady=(4, 2))

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

    # ------------------------------------------------------------------
    # Event subscriptions
    # ------------------------------------------------------------------

    def _on_data_loaded(self, **_):
        self._populate_grid()
        self._refresh_entity_listbox()

    # ------------------------------------------------------------------
    # Grid population
    # ------------------------------------------------------------------

    def _populate_grid(self):
        for day in range(8):
            for col in range(7):
                self._grid_cells[day][col].config(bg=CLR_GRID_CELL)
        self._refresh_entity_listbox()

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------

    def _all_students(self) -> list:
        tree = self._controller.state.timetable_tree
        students = set()
        for block in tree.blocks.values():
            for subblock in block.subblocks.values():
                for cl in subblock.class_lists.values():
                    students |= cl.student_list.students
        return sorted(students)

    def _student_display(self, student_id: int) -> str:
        return _student_display_fn(self._controller.state.timetable_tree, student_id)

    @staticmethod
    def _parse_student_id(value: str):
        value = value.strip()
        if value.endswith(")") and "(" in value:
            try:
                return int(value[value.rfind("(") + 1:-1])
            except ValueError:
                pass
        try:
            return int(value)
        except ValueError:
            return None

    def _all_teachers(self) -> list:
        tree = self._controller.state.timetable_tree
        teachers = set()
        for block in tree.blocks.values():
            for subblock in block.subblocks.values():
                for label in subblock.class_lists:
                    parts = label.split("_")
                    if len(parts) >= 3:
                        teachers.add("_".join(parts[1:-1]))
        return sorted(teachers)

    def _all_subjects(self) -> list:
        tree = self._controller.state.timetable_tree
        subjects = set()
        for block in tree.blocks.values():
            for subblock in block.subblocks.values():
                for label in subblock.class_lists:
                    subjects.add(label.split("_")[0])
        return sorted(subjects)

    # ------------------------------------------------------------------
    # Entity selector
    # ------------------------------------------------------------------

    def _entity_full_list(self) -> list:
        if not self._controller.state.timetable_tree:
            return []
        etype = self._entity_type_var.get()
        if etype == "Student":
            return [self._student_display(s) for s in self._all_students()]
        elif etype == "Teacher":
            return self._all_teachers()
        return self._all_subjects()

    def _refresh_entity_listbox(self):
        typed  = self._entity_value_var.get().strip().lower()
        full   = self._entity_full_list()
        filtered = [v for v in full if typed in v.lower()] if typed else full
        self._entity_listbox.delete(0, tk.END)
        for v in filtered:
            self._entity_listbox.insert(tk.END, v)

    def _on_entity_type_change(self, *_):
        self._entity_value_var.set("")
        self._refresh_entity_listbox()

    def _on_entity_search_change(self, *_):
        self._refresh_entity_listbox()

    def _on_entity_listbox_select(self, *_):
        sel = self._entity_listbox.curselection()
        if sel:
            value = self._entity_listbox.get(sel[0])
            self._entity_value_var.set(value)
            self._refresh_entity_grid(self._entity_type_var.get(), value)

    def _on_view_timetable(self):
        etype = self._entity_type_var.get()
        value = self._entity_value_var.get().strip()
        if not value or not self._controller.state.timetable_tree:
            return
        self._refresh_entity_grid(etype, value)

    # ------------------------------------------------------------------
    # Subblock popup
    # ------------------------------------------------------------------

    def _show_subblock_detail(self, subblock_name: str):
        tree = self._controller.state.timetable_tree
        if not tree:
            return
        if self._subblock_popup and self._subblock_popup.winfo_exists():
            self._subblock_popup.destroy()

        block_letter = subblock_name[0]
        block = tree.blocks.get(block_letter)
        if not block:
            return
        sb = block.subblocks.get(subblock_name)

        win = tk.Toplevel(self)
        self._subblock_popup = win
        win.title(f"Subblock {subblock_name}")
        win.configure(bg=CLR_WHITE)
        win.resizable(True, True)

        tk.Label(win, text=f"Classes in {subblock_name}",
                 font=("Calibri", 12, "bold"),
                 fg=CLR_BLUE, bg=CLR_WHITE).pack(padx=14, pady=(10, 4))

        frame = tk.Frame(win, bg=CLR_WHITE)
        frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 4))

        if not sb or not sb.class_lists:
            tk.Label(frame, text="(no classes)", fg="#9CA3AF",
                     bg=CLR_WHITE, font=("Calibri", 10)).pack()
        else:
            for label in sorted(sb.class_lists):
                cl = sb.class_lists[label]
                count = len(cl.student_list)
                tk.Label(frame, text=f"  {label}   ({count} students)",
                         bg=CLR_WHITE, fg="#1E293B", font=("Calibri", 10),
                         anchor="w").pack(fill=tk.X)

        tk.Button(win, text="Close", command=win.destroy,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=14
                  ).pack(pady=(0, 10))

    # ------------------------------------------------------------------
    # Entity grid
    # ------------------------------------------------------------------

    def _refresh_entity_grid(self, entity_type: str, value: str):
        for w in self._entity_grid_frame.winfo_children():
            w.destroy()

        self._entity_heading_var.set(f"{entity_type}: {value}")

        gf = self._entity_grid_frame
        for c in range(8):
            gf.columnconfigure(c, weight=1)

        tk.Label(gf, text="", bg=CLR_GRID_HEADER,
                 relief=tk.RIDGE, bd=1, padx=6, pady=4
                 ).grid(row=0, column=0, padx=1, pady=1, sticky="nsew")
        for col in range(7):
            tk.Label(gf, text=f"P{col+1}",
                     bg=CLR_GRID_HEADER, font=("Calibri", 10, "bold"),
                     fg="#1E293B", relief=tk.RIDGE, bd=1, padx=6, pady=4
                     ).grid(row=0, column=col+1, padx=1, pady=1, sticky="nsew")

        tree = self._controller.state.timetable_tree
        for day in range(8):
            tk.Label(gf, text=f"Day {day+1}",
                     bg=CLR_GRID_HEADER, font=("Calibri", 10, "bold"),
                     fg="#1E293B", relief=tk.RIDGE, bd=1, padx=6, pady=4
                     ).grid(row=day+1, column=0, padx=1, pady=1, sticky="nsew")
            for col in range(7):
                subblock_name = TIMETABLE_GRID[day][col]
                block_letter  = subblock_name[0]
                block = tree.blocks.get(block_letter)
                matching = []
                if block:
                    sb = block.subblocks.get(subblock_name)
                    if sb:
                        for label, cl in sb.class_lists.items():
                            parts = label.split("_")
                            if entity_type == "Student":
                                sid = self._parse_student_id(value)
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
                    else:
                        return f"{subj}  {tchr}  Gr{grade}"

                cell_text = "\n".join(_fmt(lbl) for lbl in matching) if matching else ""
                cell_bg   = CLR_GRID_ACTIVE if matching else CLR_GRID_CELL
                cell_fg   = "#1E293B" if matching else "#9CA3AF"
                tk.Label(gf,
                         text=cell_text,
                         bg=cell_bg, fg=cell_fg,
                         font=("Calibri", 10),
                         wraplength=150,
                         justify=tk.CENTER,
                         relief=tk.RIDGE, bd=1, padx=4, pady=6
                         ).grid(row=day+1, column=col+1,
                                padx=1, pady=1, sticky="nsew")
