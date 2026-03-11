"""
ui.py

Purpose
-------
Main UI for TimePyBling.

Layout
------
  [Load File]  [file path shown here]

  [ Timetable Tab ]  [ Exam Tab ]

  Timetable Tab:
    Search bar (student ID / subject code / teacher name)
    Collapsible tree:  Block -> SubBlock -> Class -> Students

  Exam Tab:
    Left panel:  Collapsible exam tree by grade
    Right panel: Exclusion list
                 - Pre-loaded with defaults
                 - Add / Remove controls
                 - [Rebuild Exam Tree] button

Usage
-----
    python ui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from timetable_tree import build_timetable_tree_from_file
from exam_tree import build_exam_tree_from_timetable_tree
from exam_clash import build_clash_graph, dsatur_colouring, is_excluded

# ------------------------------------------------
# DEFAULT EXCLUSIONS
# ------------------------------------------------
DEFAULT_EXCLUSIONS = {"ST", "LIB", "PE", "RDI"}


# ------------------------------------------------
# MAIN APPLICATION
# ------------------------------------------------
class TimePyBlingApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("TimePyBling")
        self.geometry("1100x750")
        self.configure(bg="#f0f0f0")

        # State
        self.timetable_tree = None
        self.exam_tree      = None
        self.file_path      = None
        self.exclusions     = set(DEFAULT_EXCLUSIONS)



        self._build_ui()

    # ------------------------------------------------
    # BUILD UI
    # ------------------------------------------------
    def _build_ui(self):
        # ---- Top bar ----
        top = tk.Frame(self, bg="#2c3e50", pady=8, padx=10)
        top.pack(fill=tk.X)

        tk.Label(
            top, text="TimePyBling", font=("Helvetica", 14, "bold"),
            bg="#2c3e50", fg="white"
        ).pack(side=tk.LEFT, padx=(0, 20))

        tk.Button(
            top, text="Load File", command=self._load_file,
            bg="#27ae60", fg="white", relief=tk.FLAT,
            font=("Helvetica", 10, "bold"), padx=12, pady=4
        ).pack(side=tk.LEFT)

        self.file_label = tk.Label(
            top, text="No file loaded",
            bg="#2c3e50", fg="#bdc3c7",
            font=("Helvetica", 9)
        )
        self.file_label.pack(side=tk.LEFT, padx=12)

        # ---- Notebook ----
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab frames
        self.timetable_frame = tk.Frame(self.notebook, bg="white")
        self.exam_frame      = tk.Frame(self.notebook, bg="white")

        self.notebook.add(self.timetable_frame, text="  Timetable  ")
        self.notebook.add(self.exam_frame,      text="  Exams  ")

        self._build_timetable_tab()
        self._build_exam_tab()

    # ------------------------------------------------
    # TIMETABLE TAB
    # ------------------------------------------------
    def _build_timetable_tab(self):
        # Search bar
        search_bar = tk.Frame(self.timetable_frame, bg="white", pady=6, padx=8)
        search_bar.pack(fill=tk.X)

        tk.Label(
            search_bar, text="Search:", bg="white",
            font=("Helvetica", 10)
        ).pack(side=tk.LEFT)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)

        search_entry = tk.Entry(
            search_bar, textvariable=self.search_var,
            font=("Helvetica", 10), width=30, relief=tk.SOLID, bd=1
        )
        search_entry.pack(side=tk.LEFT, padx=6)

        tk.Label(
            search_bar,
            text="student ID, subject code or teacher name",
            bg="white", fg="#888", font=("Helvetica", 9)
        ).pack(side=tk.LEFT)

        tk.Button(
            search_bar, text="Clear",
            command=lambda: self.search_var.set(""),
            relief=tk.FLAT, bg="#ecf0f1", font=("Helvetica", 9), padx=8
        ).pack(side=tk.LEFT, padx=4)

        # Tree
        tree_frame = tk.Frame(self.timetable_frame, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.tt_tree = ttk.Treeview(tree_frame, show="tree")
        self.tt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tt_scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.tt_tree.yview
        )
        tt_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tt_tree.configure(yscrollcommand=tt_scroll.set)

        self._style_tree(self.tt_tree)

    # ------------------------------------------------
    # EXAM TAB
    # ------------------------------------------------
    def _build_exam_tab(self):
        # Split: left = tree, right = exclusions + clash
        pane = tk.PanedWindow(
            self.exam_frame, orient=tk.HORIZONTAL,
            bg="#ddd", sashwidth=5
        )
        pane.pack(fill=tk.BOTH, expand=True)

        # ---- Left: Exam tree ----
        left = tk.Frame(pane, bg="white")
        pane.add(left, minsize=500)

        tk.Label(
            left, text="Exam Tree", bg="white",
            font=("Helvetica", 10, "bold"), pady=6
        ).pack()

        tree_frame = tk.Frame(left, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        self.ex_tree = ttk.Treeview(tree_frame, show="tree")
        self.ex_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ex_scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.ex_tree.yview
        )
        ex_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.ex_tree.configure(yscrollcommand=ex_scroll.set)

        self._style_tree(self.ex_tree)

        # ---- Right: Controls ----
        right = tk.Frame(pane, bg="white")
        pane.add(right, minsize=260)

        # Exclusions panel
        excl_frame = tk.LabelFrame(
            right, text="Exam Exclusions",
            bg="white", font=("Helvetica", 10, "bold"),
            padx=8, pady=8
        )
        excl_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(
            excl_frame,
            text="Subject codes excluded from exam scheduling:",
            bg="white", fg="#555", font=("Helvetica", 8),
            wraplength=220, justify=tk.LEFT
        ).pack(anchor=tk.W)

        self.excl_listbox = tk.Listbox(
            excl_frame, height=8,
            font=("Courier", 10), relief=tk.SOLID, bd=1,
            selectmode=tk.SINGLE
        )
        self.excl_listbox.pack(fill=tk.X, pady=6)
        self._refresh_exclusion_listbox()

        # Add row
        add_row = tk.Frame(excl_frame, bg="white")
        add_row.pack(fill=tk.X)

        self.excl_entry = tk.Entry(
            add_row, font=("Helvetica", 10),
            relief=tk.SOLID, bd=1, width=10
        )
        self.excl_entry.pack(side=tk.LEFT)
        self.excl_entry.bind("<Return>", lambda e: self._add_exclusion())

        tk.Button(
            add_row, text="Add",
            command=self._add_exclusion,
            bg="#2980b9", fg="white", relief=tk.FLAT,
            font=("Helvetica", 9, "bold"), padx=8
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            excl_frame, text="Remove Selected",
            command=self._remove_exclusion,
            bg="#c0392b", fg="white", relief=tk.FLAT,
            font=("Helvetica", 9), padx=8, pady=2
        ).pack(anchor=tk.W, pady=(4, 0))

        tk.Button(
            right, text="Rebuild Exam Tree",
            command=self._rebuild_exam,
            bg="#27ae60", fg="white", relief=tk.FLAT,
            font=("Helvetica", 10, "bold"), padx=12, pady=6
        ).pack(padx=10, pady=8, fill=tk.X)

        # Clash / slot summary panel
        clash_frame = tk.LabelFrame(
            right, text="Slot Summary",
            bg="white", font=("Helvetica", 10, "bold"),
            padx=8, pady=8
        )
        clash_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.clash_text = tk.Text(
            clash_frame, font=("Courier", 9),
            relief=tk.FLAT, bg="#f8f8f8",
            state=tk.DISABLED, wrap=tk.NONE
        )
        clash_scroll = ttk.Scrollbar(
            clash_frame, orient=tk.VERTICAL,
            command=self.clash_text.yview
        )
        self.clash_text.configure(yscrollcommand=clash_scroll.set)
        clash_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.clash_text.pack(fill=tk.BOTH, expand=True)

    # ------------------------------------------------
    # LOAD FILE
    # ------------------------------------------------
    def _load_file(self):
        path = filedialog.askopenfilename(
            title="Select student timetable file",
            filetypes=[
                ("Timetable files", "*.csv *.xlsx *.xls"),
                ("CSV files",       "*.csv"),
                ("Excel files",     "*.xlsx *.xls"),
            ]
        )
        if not path:
            return

        self.file_path = Path(path)
        self.file_label.config(text=str(self.file_path.name))

        try:
            self.timetable_tree = build_timetable_tree_from_file(self.file_path)
            self._populate_timetable_tree()
            self._rebuild_exam()
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    # ------------------------------------------------
    # POPULATE TIMETABLE TREE VIEW
    # ------------------------------------------------
    def _populate_timetable_tree(self, filter_text=""):
        self.tt_tree.delete(*self.tt_tree.get_children())

        if not self.timetable_tree:
            return

        ft = filter_text.strip().lower()

        for block_name in sorted(self.timetable_tree.blocks.keys()):
            block = self.timetable_tree.blocks[block_name]

            block_node = self.tt_tree.insert(
                "", tk.END, text=f"Block {block_name}",
                open=bool(ft)
            )

            sorted_subblocks = sorted(
                block.subblocks.keys(),
                key=lambda n: int(n[1:])
            )

            for sb_name in sorted_subblocks:
                subblock   = block.subblocks[sb_name]
                sb_node    = None

                for class_label in sorted(subblock.class_lists.keys()):
                    cl = subblock.class_lists[class_label]

                    # Filter logic
                    if ft:
                        students_str = str(cl.student_list.get_sorted())
                        label_lower  = class_label.lower()
                        match = (
                            ft in label_lower or
                            ft in students_str
                        )
                        if not match:
                            continue

                    # Lazy-create subblock node only if it has matches
                    if sb_node is None:
                        sb_node = self.tt_tree.insert(
                            block_node, tk.END,
                            text=sb_name, open=bool(ft)
                        )

                    count    = len(cl.student_list)
                    cl_node  = self.tt_tree.insert(
                        sb_node, tk.END,
                        text=f"{class_label}  ({count} students)",
                        open=False
                    )

                    # Students as leaf nodes — show in chunks of 20
                    students = cl.student_list.get_sorted()
                    chunk_size = 20
                    for i in range(0, len(students), chunk_size):
                        chunk = students[i:i + chunk_size]
                        self.tt_tree.insert(
                            cl_node, tk.END,
                            text=str(chunk)
                        )

    # ------------------------------------------------
    # POPULATE EXAM TREE VIEW
    # ------------------------------------------------
    def _populate_exam_tree(self):
        self.ex_tree.delete(*self.ex_tree.get_children())

        if not self.exam_tree:
            return

        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node = self.exam_tree.grades[grade_label]

            grade_node_ui = self.ex_tree.insert(
                "", tk.END, text=grade_label, open=False
            )

            for subject_label in sorted(grade_node.exam_subjects.keys()):
                subject = grade_node.exam_subjects[subject_label]

                subj_node = self.ex_tree.insert(
                    grade_node_ui, tk.END,
                    text=subject_label, open=False
                )

                for class_label in sorted(subject.class_lists.keys()):
                    cl = subject.class_lists[class_label]
                    count = len(cl.student_list)

                    cl_node = self.ex_tree.insert(
                        subj_node, tk.END,
                        text=f"{class_label}  ({count} students)",
                        open=False
                    )

                    students = cl.student_list.get_sorted()
                    chunk_size = 20
                    for i in range(0, len(students), chunk_size):
                        chunk = students[i:i + chunk_size]
                        self.ex_tree.insert(
                            cl_node, tk.END, text=str(chunk)
                        )

    # ------------------------------------------------
    # REBUILD EXAM TREE (respects exclusions)
    # ------------------------------------------------
    def _rebuild_exam(self):
        if not self.timetable_tree:
            return

        self.exam_tree = build_exam_tree_from_timetable_tree(
            self.timetable_tree
        )
        self._populate_exam_tree()
        self._update_clash_summary()

    # ------------------------------------------------
    # SLOT SUMMARY
    # ------------------------------------------------
    def _update_clash_summary(self):
        if not self.exam_tree:
            return

        lines = []

        for grade_label in sorted(self.exam_tree.grades.keys()):
            grade_node = self.exam_tree.grades[grade_label]

            # Build student sets, skipping excluded subjects
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

            # Group by slot
            slots = {}
            for subj, slot in assignment.items():
                slots.setdefault(slot, []).append(subj)

            lines.append(f"{grade_label}  —  {num_slots} slots")

            for slot_num in sorted(slots.keys()):
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

        self.clash_text.config(state=tk.NORMAL)
        self.clash_text.delete("1.0", tk.END)
        self.clash_text.insert(tk.END, "\n".join(lines))
        self.clash_text.config(state=tk.DISABLED)

    # ------------------------------------------------
    # SEARCH
    # ------------------------------------------------
    def _on_search_change(self, *args):
        self._populate_timetable_tree(
            filter_text=self.search_var.get()
        )

    # ------------------------------------------------
    # EXCLUSION LIST
    # ------------------------------------------------
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
        self._update_clash_summary()

    def _remove_exclusion(self):
        selection = self.excl_listbox.curselection()
        if not selection:
            return
        code = self.excl_listbox.get(selection[0])
        self.exclusions.discard(code)
        self._refresh_exclusion_listbox()
        self._update_clash_summary()

    # ------------------------------------------------
    # TREE STYLING
    # ------------------------------------------------
    def _style_tree(self, tree: ttk.Treeview):
        style = ttk.Style()
        style.configure(
            "Treeview",
            font=("Courier", 9),
            rowheight=22,
            background="white",
            fieldbackground="white"
        )
        style.configure(
            "Treeview.Heading",
            font=("Helvetica", 10, "bold")
        )
        tree.tag_configure("oddrow",  background="#f9f9f9")
        tree.tag_configure("evenrow", background="white")


# ------------------------------------------------
# ENTRY POINT
# ------------------------------------------------
if __name__ == "__main__":
    app = TimePyBlingApp()
    app.mainloop()