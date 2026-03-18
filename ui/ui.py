"""
ui.py - TimePyBling main interface
4 tabs: Timetable | Verification | Timetable View | Exams
"""

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import pandas as pd

from core.timetable_tree     import build_timetable_tree_from_file
from reader.exam_tree        import build_exam_tree_from_timetable_tree
from reader.exam_clash       import build_clash_graph, dsatur_colouring, is_excluded
from reader.verify_timetable import find_student_clashes

DEFAULT_EXCLUSIONS = {"ST", "LIB", "PE", "RDI"}
BLOCKS = list("ABCDEFGH")
DAYS   = list(range(1, 8))

BLOCK_COLOURS = {
    "A": "#EEF4FB", "B": "#E8F5E9", "C": "#FFF8E1",
    "D": "#FCE4EC", "E": "#F3E5F5", "F": "#E0F7FA",
    "G": "#FBE9E7", "H": "#F1F8E9",
}

CLR_HEADER   = "#2c3e50"
CLR_GREEN    = "#27ae60"
CLR_BLUE     = "#2980b9"
CLR_RED      = "#c0392b"
CLR_LIGHT    = "#ecf0f1"
CLR_MID      = "#bdc3c7"
CLR_WHITE    = "white"
CLR_BG       = "#f5f5f5"
CLR_GRID_HDR = "#34495e"
CLR_EMPTY    = "#fafafa"


def _scrolled_text(parent, **kw):
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


def _write(widget, text, tag=""):
    widget.config(state=tk.NORMAL)
    if tag:
        widget.insert(tk.END, text, tag)
    else:
        widget.insert(tk.END, text)
    widget.config(state=tk.DISABLED)


def _clear(widget):
    widget.config(state=tk.NORMAL)
    widget.delete("1.0", tk.END)
    widget.config(state=tk.DISABLED)


def _extract_teachers_from_tree(tree):
    teachers = set()
    for block in tree.blocks.values():
        for subblock in block.subblocks.values():
            for label in subblock.class_lists:
                parts = label.split("_")
                if len(parts) >= 3:
                    code = "_".join(parts[1:-1])
                    if code:
                        teachers.add(code)
    return sorted(teachers)


class SearchableList(tk.Frame):
    def __init__(self, parent, label, on_select=None, height=8, **kw):
        super().__init__(parent, bg=CLR_WHITE, **kw)
        self._all_items = []
        self._on_select = on_select
        self._selected  = None

        tk.Label(self, text=label, bg=CLR_WHITE,
                 font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(0,2))

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._on_search)
        tk.Entry(self, textvariable=self._search_var,
                 font=("Helvetica", 9), relief=tk.SOLID, bd=1
                 ).pack(fill=tk.X, pady=(0,2))

        lb_frame = tk.Frame(self, bg=CLR_WHITE)
        lb_frame.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._lb = tk.Listbox(lb_frame, height=height, font=("Courier", 9),
                               relief=tk.SOLID, bd=1, selectmode=tk.SINGLE,
                               yscrollcommand=sb.set, exportselection=False)
        self._lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=self._lb.yview)
        self._lb.bind("<<ListboxSelect>>", self._on_listbox_select)

        self._sel_label = tk.Label(self, text="None selected", bg=CLR_WHITE,
                                    fg="#888", font=("Helvetica", 8, "italic"),
                                    anchor=tk.W)
        self._sel_label.pack(fill=tk.X, pady=(2,0))

    def set_items(self, items):
        self._all_items = list(items)
        self._refresh_listbox(self._search_var.get())

    def get_selection(self):
        return self._selected

    def clear_selection(self):
        self._selected = None
        self._lb.selection_clear(0, tk.END)
        self._sel_label.config(text="None selected", fg="#888")

    def _on_search(self, *_):
        self._refresh_listbox(self._search_var.get())

    def _refresh_listbox(self, query):
        q        = query.strip().lower()
        filtered = [i for i in self._all_items if q in i.lower()] if q else self._all_items
        self._lb.delete(0, tk.END)
        for item in filtered:
            self._lb.insert(tk.END, item)
        if self._selected:
            for idx in range(self._lb.size()):
                if self._lb.get(idx) == self._selected:
                    self._lb.selection_set(idx)
                    self._lb.see(idx)
                    break

    def _on_listbox_select(self, _event):
        sel = self._lb.curselection()
        if not sel:
            return
        item = self._lb.get(sel[0])
        self._selected = item
        short = item if len(item) <= 40 else item[:37] + "..."
        self._sel_label.config(text="OK  " + short, fg=CLR_GREEN)
        if self._on_select:
            self._on_select(item)


class TimetableGrid(tk.Frame):
    """
    7-day x 8-block grid.
    Rows = cycle days 1-7, Columns = lesson blocks A-H.
    """
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=CLR_WHITE, **kw)
        self._cells = {}
        self._build_grid()

    def _build_grid(self):
        tk.Label(self, text="", bg=CLR_GRID_HDR,
                 relief=tk.FLAT, bd=0, width=6
                 ).grid(row=0, column=0, sticky="nsew", padx=(0,1), pady=(0,1))

        for col_idx, block in enumerate(BLOCKS, start=1):
            tk.Label(self, text="Block " + block,
                     bg=BLOCK_COLOURS[block], fg="#333",
                     font=("Helvetica", 9, "bold"), width=13,
                     relief=tk.FLAT, bd=0, anchor=tk.CENTER
                     ).grid(row=0, column=col_idx, sticky="nsew",
                            padx=(0,1), pady=(0,1), ipady=5)

        for row_idx, day in enumerate(DAYS, start=1):
            tk.Label(self, text="Day " + str(day),
                     bg=CLR_GRID_HDR, fg=CLR_WHITE,
                     font=("Helvetica", 9, "bold"), width=6,
                     relief=tk.FLAT, bd=0, anchor=tk.CENTER
                     ).grid(row=row_idx, column=0, sticky="nsew",
                            padx=(0,1), pady=(0,1), ipadx=4)

            for col_idx, block in enumerate(BLOCKS, start=1):
                lbl = tk.Label(self, text="", bg=CLR_EMPTY, fg="#444",
                               font=("Courier", 8), width=13,
                               wraplength=96, justify=tk.CENTER,
                               anchor=tk.CENTER, relief=tk.FLAT, bd=0)
                lbl.grid(row=row_idx, column=col_idx, sticky="nsew",
                         padx=(0,1), pady=(0,1), ipady=9)
                self._cells[(day, block)] = lbl

        for c in range(len(BLOCKS)+1):
            self.columnconfigure(c, weight=1)
        for r in range(len(DAYS)+1):
            self.rowconfigure(r, weight=1)

    def clear(self):
        for lbl in self._cells.values():
            lbl.config(text="", bg=CLR_EMPTY, fg="#444")

    def render_student(self, schedule):
        self.clear()
        for (day, block), lbl in self._cells.items():
            val = schedule.get(block + str(day), "")
            if val and val not in ("nan", "FREE"):
                parts   = val.strip().split()
                subj    = parts[0]
                teacher = " ".join(parts[1:]) if len(parts) > 1 else ""
                display = (subj + "\n" + teacher) if teacher else subj
                lbl.config(text=display,
                           bg=BLOCK_COLOURS.get(block, CLR_EMPTY),
                           fg="#1a1a1a")
            elif val == "FREE":
                lbl.config(text="FREE", bg="#f0f0f0", fg="#bbb")

    def render_teacher(self, schedule):
        self.clear()
        for (day, block), lbl in self._cells.items():
            labels = schedule.get(block + str(day), [])
            if not labels:
                continue
            lines = []
            for lab in labels:
                parts = lab.split("_")
                subj  = parts[0]
                grade = parts[-1] if len(parts) >= 3 else ""
                lines.append(subj + " Gr" + grade)
            fg = CLR_RED if len(labels) > 1 else "#1a1a1a"
            lbl.config(text="\n".join(lines),
                       bg=BLOCK_COLOURS.get(block, CLR_EMPTY),
                       fg=fg)


class TimePyBlingApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("TimePyBling")
        self.geometry("1280x860")
        self.configure(bg=CLR_BG)

        self.timetable_tree  = None
        self.exam_tree       = None
        self.st1_path        = None
        self.st1_df          = None
        self.student_roster  = None
        self.teacher_list    = []
        self.exclusions      = set(DEFAULT_EXCLUSIONS)

        self._build_ui()

    def _build_ui(self):
        self._build_topbar()
        self._build_notebook()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=CLR_HEADER, pady=8, padx=10)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="TimePyBling", font=("Helvetica", 14, "bold"),
                 bg=CLR_HEADER, fg=CLR_WHITE).pack(side=tk.LEFT, padx=(0,24))
        tk.Button(bar, text="Load Timetable", command=self._load_st1,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=10, pady=3).pack(side=tk.LEFT)
        self.st1_label = tk.Label(bar, text="No timetable loaded",
                                   bg=CLR_HEADER, fg=CLR_MID, font=("Helvetica", 9))
        self.st1_label.pack(side=tk.LEFT, padx=(6,20))

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.tab_timetable    = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_verification = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_view         = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.tab_exams        = tk.Frame(self.notebook, bg=CLR_WHITE)
        self.notebook.add(self.tab_timetable,    text="  Timetable  ")
        self.notebook.add(self.tab_verification, text="  Verification  ")
        self.notebook.add(self.tab_view,         text="  Timetable View  ")
        self.notebook.add(self.tab_exams,        text="  Exams  ")
        self._build_timetable_tab()
        self._build_verification_tab()
        self._build_view_tab()
        self._build_exam_tab()

    # ── TAB 1: TIMETABLE TREE ────────────────────────────────

    def _build_timetable_tab(self):
        bar = tk.Frame(self.tab_timetable, bg=CLR_WHITE, pady=6, padx=8)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="Search:", bg=CLR_WHITE, font=("Helvetica", 10)).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)
        tk.Entry(bar, textvariable=self.search_var, font=("Helvetica", 10),
                 width=30, relief=tk.SOLID, bd=1).pack(side=tk.LEFT, padx=6)
        tk.Label(bar, text="student ID  subject code  teacher name",
                 bg=CLR_WHITE, fg="#888", font=("Helvetica", 9)).pack(side=tk.LEFT)
        tk.Button(bar, text="Clear", command=lambda: self.search_var.set(""),
                  relief=tk.FLAT, bg=CLR_LIGHT, font=("Helvetica", 9), padx=8
                  ).pack(side=tk.LEFT, padx=4)
        f = tk.Frame(self.tab_timetable, bg=CLR_WHITE)
        f.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        self.tt_tree = ttk.Treeview(f, show="tree")
        sb = ttk.Scrollbar(f, orient=tk.VERTICAL, command=self.tt_tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tt_tree.configure(yscrollcommand=sb.set)
        self.tt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._style_tree(self.tt_tree)

    # ── TAB 2: VERIFICATION ──────────────────────────────────

    def _build_verification_tab(self):
        hf = tk.Frame(self.tab_verification, bg=CLR_WHITE)
        hf.pack(fill=tk.X, padx=8, pady=(8,2))
        tk.Label(hf, text="Student Clash Report", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(hf, text="Re-run", command=self._run_verification,
                  bg=CLR_LIGHT, font=("Helvetica", 8), relief=tk.FLAT, padx=8
                  ).pack(side=tk.RIGHT)
        rf = tk.Frame(self.tab_verification, bg=CLR_WHITE)
        rf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        self.clash_report = _scrolled_text(rf)
        self.clash_report.tag_config("pass",    foreground=CLR_GREEN)
        self.clash_report.tag_config("fail",    foreground=CLR_RED)
        self.clash_report.tag_config("heading", foreground="#333",
                                     font=("Courier", 9, "bold"))
        self.clash_report.tag_config("dim",     foreground="#888")

    # ── TAB 3: TIMETABLE VIEW ────────────────────────────────

    def _build_view_tab(self):
        outer = tk.PanedWindow(self.tab_view, orient=tk.HORIZONTAL,
                               bg="#ccc", sashwidth=5)
        outer.pack(fill=tk.BOTH, expand=True)

        # Left selector panel
        left = tk.Frame(outer, bg=CLR_WHITE, padx=10, pady=10)
        outer.add(left, minsize=280)

        tk.Label(left, text="Select Teacher or Student", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold")).pack(anchor=tk.W, pady=(0,4))
        tk.Label(left,
                 text="Choose one. Selecting a teacher clears the\n"
                      "student selection and vice versa.",
                 bg=CLR_WHITE, fg="#777", font=("Helvetica", 8),
                 justify=tk.LEFT).pack(anchor=tk.W, pady=(0,10))

        self._teacher_selector = SearchableList(
            left, label="Teacher",
            on_select=self._on_teacher_selected, height=8)
        self._teacher_selector.pack(fill=tk.X, pady=(0,8))

        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self._student_selector = SearchableList(
            left, label="Student  (name, surname, or ID)",
            on_select=self._on_student_selected, height=10)
        self._student_selector.pack(fill=tk.X, pady=(0,10))

        tk.Button(left, text="Generate Timetable",
                  command=self._generate_view,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 10, "bold"), pady=7
                  ).pack(fill=tk.X, pady=(4,0))

        self._view_info_label = tk.Label(
            left, text="Load a timetable first.",
            bg=CLR_WHITE, fg="#888", font=("Helvetica", 8, "italic"),
            wraplength=240, justify=tk.LEFT)
        self._view_info_label.pack(anchor=tk.W, pady=(6,0))

        # Right grid panel
        right = tk.Frame(outer, bg=CLR_WHITE)
        outer.add(right, minsize=700)

        self._grid_title = tk.Label(right, text="No timetable generated yet",
                                     bg=CLR_WHITE, fg="#999",
                                     font=("Helvetica", 10, "bold"),
                                     anchor=tk.W, pady=6, padx=10)
        self._grid_title.pack(fill=tk.X)
        tk.Label(right,
                 text="Rows = Cycle Days 1-7   |   Columns = Lesson Blocks A-H",
                 bg=CLR_WHITE, fg="#aaa", font=("Helvetica", 8, "italic"),
                 anchor=tk.W, padx=10).pack(fill=tk.X)
        ttk.Separator(right, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        canvas_frame = tk.Frame(right, bg=CLR_WHITE)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        canvas = tk.Canvas(canvas_frame, bg=CLR_WHITE, highlightthickness=0)
        v_sb = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        h_sb = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(xscrollcommand=h_sb.set, yscrollcommand=v_sb.set)

        self._grid_container = tk.Frame(canvas, bg=CLR_WHITE)
        cw = canvas.create_window((0,0), window=self._grid_container, anchor="nw")
        self._grid_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind(
            "<Configure>",
            lambda e, _cw=cw: canvas.itemconfig(_cw, width=e.width))

        self._tt_grid = TimetableGrid(self._grid_container)
        self._tt_grid.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

    # ── TAB 4: EXAMS ─────────────────────────────────────────

    def _build_exam_tab(self):
        pane = tk.PanedWindow(self.tab_exams, orient=tk.HORIZONTAL,
                              bg="#ddd", sashwidth=5)
        pane.pack(fill=tk.BOTH, expand=True)
        left = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(left, minsize=500)
        tk.Label(left, text="Exam Tree", bg=CLR_WHITE,
                 font=("Helvetica", 10, "bold"), pady=6).pack()
        tf = tk.Frame(left, bg=CLR_WHITE)
        tf.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0,6))
        self.ex_tree = ttk.Treeview(tf, show="tree")
        esb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.ex_tree.yview)
        esb.pack(side=tk.RIGHT, fill=tk.Y)
        self.ex_tree.configure(yscrollcommand=esb.set)
        self.ex_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._style_tree(self.ex_tree)

        right = tk.Frame(pane, bg=CLR_WHITE)
        pane.add(right, minsize=260)
        excl_frame = tk.LabelFrame(right, text="Exam Exclusions", bg=CLR_WHITE,
                                    font=("Helvetica", 10, "bold"), padx=8, pady=8)
        excl_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(excl_frame, text="Subject codes excluded from exam scheduling:",
                 bg=CLR_WHITE, fg="#555", font=("Helvetica", 8),
                 wraplength=220, justify=tk.LEFT).pack(anchor=tk.W)
        self.excl_listbox = tk.Listbox(excl_frame, height=6, font=("Courier", 10),
                                        relief=tk.SOLID, bd=1, selectmode=tk.SINGLE)
        self.excl_listbox.pack(fill=tk.X, pady=6)
        self._refresh_exclusion_listbox()
        add_row = tk.Frame(excl_frame, bg=CLR_WHITE)
        add_row.pack(fill=tk.X)
        self.excl_entry = tk.Entry(add_row, font=("Helvetica", 10),
                                    relief=tk.SOLID, bd=1, width=10)
        self.excl_entry.pack(side=tk.LEFT)
        self.excl_entry.bind("<Return>", lambda e: self._add_exclusion())
        tk.Button(add_row, text="Add", command=self._add_exclusion,
                  bg=CLR_BLUE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9, "bold"), padx=8).pack(side=tk.LEFT, padx=4)
        tk.Button(excl_frame, text="Remove Selected", command=self._remove_exclusion,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 9), padx=8, pady=2).pack(anchor=tk.W, pady=(4,0))
        tk.Button(right, text="Rebuild Exam Tree", command=self._rebuild_exam,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Helvetica", 10, "bold"), padx=12, pady=6
                  ).pack(padx=10, pady=8, fill=tk.X)
        slot_lf = tk.LabelFrame(right, text="Slot Summary", bg=CLR_WHITE,
                                 font=("Helvetica", 10, "bold"), padx=8, pady=8)
        slot_lf.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))
        self.slot_text = _scrolled_text(slot_lf)

    # ── LOAD ─────────────────────────────────────────────────

    def _load_st1(self):
        path = filedialog.askopenfilename(
            title="Select student timetable (ST1.xlsx)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not path:
            return
        self.st1_path = Path(path)
        self.st1_label.config(text=self.st1_path.name, fg=CLR_WHITE)
        try:
            self.timetable_tree = build_timetable_tree_from_file(self.st1_path)
            self.st1_df         = pd.read_excel(self.st1_path)
            self.teacher_list   = _extract_teachers_from_tree(self.timetable_tree)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return
        self._populate_timetable_tree()
        self._populate_view_selectors()
        self._rebuild_exam()
        self._run_verification()
        self.notebook.select(self.tab_verification)

    # ── VERIFICATION ─────────────────────────────────────────

    def _run_verification(self):
        if not self.timetable_tree:
            return
        w = self.clash_report
        _clear(w)
        clashes = find_student_clashes(self.timetable_tree)
        if not clashes:
            _write(w, "PASS  --  no student clashes found\n", "pass")
        else:
            _write(w, "FAIL  --  " + str(len(clashes)) + " student clash(es)\n", "fail")
        _write(w, "-" * 60 + "\n", "dim")
        if clashes:
            _write(w, "\nSTUDENT DOUBLE-BOOKINGS\n", "heading")
            by_sb = {}
            for c in clashes:
                by_sb.setdefault(c["subblock"], []).append(c)
            for sb in sorted(by_sb, key=lambda n: (n[0], int(n[1:]))):
                _write(w, "\n  Subblock " + sb + "\n", "heading")
                for entry in sorted(by_sb[sb], key=lambda e: e["student"]):
                    classes = "  vs  ".join(entry["classes"])
                    _write(w, "    Student " + str(entry["student"]).rjust(6) + ":  " + classes + "\n", "fail")
        _write(w, "\n" + "-" * 60 + "\n", "dim")
        _write(w, "Total violations: " + str(len(clashes)) + "\n",
               "pass" if not clashes else "fail")

    # ── TIMETABLE TREE BROWSER ───────────────────────────────

    def _populate_timetable_tree(self, filter_text=""):
        self.tt_tree.delete(*self.tt_tree.get_children())
        if not self.timetable_tree:
            return
        ft = filter_text.strip().lower()
        for block_name in sorted(self.timetable_tree.blocks.keys()):
            block      = self.timetable_tree.blocks[block_name]
            block_node = self.tt_tree.insert("", tk.END,
                                             text="Block " + block_name, open=bool(ft))
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
                                                       text=sb_name, open=bool(ft))
                    count   = len(cl.student_list)
                    cl_node = self.tt_tree.insert(sb_node, tk.END,
                                                   text=class_label + "  (" + str(count) + " students)")
                    students = cl.student_list.get_sorted()
                    for i in range(0, len(students), 20):
                        self.tt_tree.insert(cl_node, tk.END, text=str(students[i:i+20]))

    def _on_search_change(self, *args):
        self._populate_timetable_tree(filter_text=self.search_var.get())

    # ── VIEW SELECTORS ───────────────────────────────────────

    def _populate_view_selectors(self):
        self._teacher_selector.set_items(self.teacher_list)
        if self.student_roster:
            items = [r.display for r in self.student_roster.all_records()]
            self._student_selector.set_items(items)
            self._view_info_label.config(
                text=str(len(self.teacher_list)) + " teachers  |  " +
                     str(len(self.student_roster)) + " students loaded.",
                fg="#555")

    def _on_teacher_selected(self, _):
        self._student_selector.clear_selection()

    def _on_student_selected(self, _):
        self._teacher_selector.clear_selection()

    # ── GENERATE GRID ────────────────────────────────────────

    def _generate_view(self):
        if not self.timetable_tree:
            messagebox.showinfo("No data", "Load a timetable first.")
            return
        teacher = self._teacher_selector.get_selection()
        student = self._student_selector.get_selection()
        if not teacher and not student:
            messagebox.showinfo("Nothing selected",
                                "Select a teacher or a student first.")
            return
        if teacher:
            self._generate_teacher_view(teacher)
        else:
            self._generate_student_view(student)

    def _generate_teacher_view(self, teacher_code):
        schedule = {}
        for block in self.timetable_tree.blocks.values():
            for sb_name, subblock in block.subblocks.items():
                for label in subblock.class_lists:
                    parts  = label.split("_")
                    t_code = "_".join(parts[1:-1]) if len(parts) >= 3 else ""
                    if t_code == teacher_code:
                        schedule.setdefault(sb_name, []).append(label)
        self._grid_title.config(text="Teacher:  " + teacher_code, fg="#2c3e50")
        self._tt_grid.render_teacher(schedule)

    def _generate_student_view(self, display_str):
        if self.st1_df is None:
            return
        m = re.search(r'\[(\d+)\]', display_str)
        if not m:
            messagebox.showerror("Error", "Could not parse student ID from: " + display_str)
            return
        student_id = int(m.group(1))

        df = self.st1_df.copy()
        df["_sid"] = df["Studentid"].apply(
            lambda x: int(float(x)) if not pd.isna(x) else -1)
        rows = df[df["_sid"] == student_id]
        if rows.empty:
            messagebox.showerror("Not found", "Student ID " + str(student_id) + " not found.")
            return
        row = rows.iloc[0]

        timetable_cols = [c for c in self.st1_df.columns
                          if re.fullmatch(r"[A-H]\d+", str(c))]
        schedule = {}
        for col in timetable_cols:
            val = row[col]
            if not pd.isna(val):
                schedule[col] = str(val).strip()

        surname   = str(row.get("SSurname",   "")).strip()
        firstname = str(row.get("SFirstname", "")).strip()
        grade_val = row.get("Grade", "")
        grade     = int(float(grade_val)) if not pd.isna(grade_val) else "?"
        reg_cls   = str(row.get("Class", "")).strip()

        self._grid_title.config(
            text="Student " + str(student_id) + ":  " + surname + ", " + firstname +
                 "  --  Grade " + str(grade) + "  (" + reg_cls + ")",
            fg="#2c3e50")
        self._tt_grid.render_student(schedule)

    # ── EXAM HELPERS ─────────────────────────────────────────

    def _populate_exam_tree(self):
        self.ex_tree.delete(*self.ex_tree.get_children())
        if not self.exam_tree:
            return
        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node = self.exam_tree.grades[grade_label]
            grade_ui   = self.ex_tree.insert("", tk.END, text=grade_label, open=False)
            for subj_label in sorted(grade_node.exam_subjects.keys()):
                subject = grade_node.exam_subjects[subj_label]
                subj_ui = self.ex_tree.insert(grade_ui, tk.END, text=subj_label, open=False)
                for class_label in sorted(subject.class_lists.keys()):
                    cl    = subject.class_lists[class_label]
                    count = len(cl.student_list)
                    cl_ui = self.ex_tree.insert(subj_ui, tk.END,
                                                 text=class_label + "  (" + str(count) + " students)")
                    students = cl.student_list.get_sorted()
                    for i in range(0, len(students), 20):
                        self.ex_tree.insert(cl_ui, tk.END, text=str(students[i:i+20]))

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
            slots = {}
            for subj, slot in assignment.items():
                slots.setdefault(slot, []).append(subj)
            lines.append(grade_label + "  --  " + str(num_slots) + " slot(s)")
            for slot_num in sorted(slots):
                group         = sorted(slots[slot_num])
                slot_students = set()
                for s in group:
                    slot_students |= student_sets[s]
                lines.append("  Slot " + str(slot_num+1).rjust(2) +
                              " (" + str(len(slot_students)).rjust(3) + " students): " +
                              ", ".join(group))
            lines.append("")
        _clear(self.slot_text)
        self.slot_text.config(state=tk.NORMAL)
        self.slot_text.insert(tk.END, "\n".join(lines))
        self.slot_text.config(state=tk.DISABLED)

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

    def _style_tree(self, tree):
        style = ttk.Style()
        style.configure("Treeview", font=("Courier", 9), rowheight=22,
                         background=CLR_WHITE, fieldbackground=CLR_WHITE)
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))


if __name__ == "__main__":
    app = TimePyBlingApp()
    app.mainloop()