from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from app.state import AppState
from app.events import EventBus
from app.controller import AppController, EVT_TIMETABLE_LOADED, EVT_STATE_LOADED
from ui.constants import (CLR_BG, CLR_HEADER, CLR_ORANGE, CLR_WHITE,
                           CLR_LIGHT, CLR_MID, CLR_BLUE, CLR_GREEN, CLR_RED,
                           CLR_GRID_HEADER)
from ui.tabs.timetable    import TimetableTab
from ui.tabs.verification import VerificationTab
from ui.tabs.exam         import ExamTab

DEFAULT_STATE_PATH = Path("data/timetable_state.json")


class TimePyBlingApp(tk.Tk):

    def __init__(self) -> None:
        super().__init__()
        self.title("TimePyBling")
        self.configure(bg=CLR_BG)
        self.minsize(900, 600)
        self.after(0, lambda: self.state("zoomed"))

        self._bus        = EventBus()
        self._state      = AppState()
        self._controller = AppController(self._state, self._bus)

        self._apply_theme()
        self._build_topbar()
        self._build_notebook()

        self._bus.subscribe(EVT_TIMETABLE_LOADED, self._on_data_loaded)
        self._bus.subscribe(EVT_STATE_LOADED,      self._on_data_loaded)

        self.after(300, self._auto_load)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _apply_theme(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background=CLR_BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=("Calibri", 11), padding=[14, 6],
                         background=CLR_LIGHT, foreground="#1E293B")
        style.map("TNotebook.Tab",
                  background=[("selected", CLR_ORANGE)],
                  foreground=[("selected", CLR_WHITE)])
        style.configure("Treeview", font=("Calibri", 10), rowheight=26,
                         background=CLR_WHITE, fieldbackground=CLR_WHITE,
                         foreground="#1E293B")
        style.configure("Treeview.Heading", font=("Calibri", 11, "bold"),
                         background=CLR_GRID_HEADER, foreground="#1E293B")
        style.map("Treeview",
                  background=[("selected", CLR_BLUE)],
                  foreground=[("selected", CLR_WHITE)])
        style.configure("TCombobox", font=("Calibri", 10),
                         selectbackground=CLR_BLUE, fieldbackground=CLR_WHITE)
        style.configure("TScrollbar", background=CLR_MID, troughcolor=CLR_LIGHT)

    def _build_topbar(self) -> None:
        bar = tk.Frame(self, bg=CLR_HEADER, pady=8, padx=10)
        bar.pack(fill=tk.X)
        tk.Label(bar, text="TimePyBling", font=("Calibri", 15, "bold"),
                 bg=CLR_HEADER, fg=CLR_WHITE).pack(side=tk.LEFT, padx=(0, 24))
        tk.Button(bar, text="Load Timetable", command=self._load_st1,
                  bg=CLR_ORANGE, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 10, "bold"), padx=12, pady=4).pack(side=tk.LEFT)
        self._st1_label = tk.Label(bar, text="No timetable loaded",
                                    bg=CLR_HEADER, fg=CLR_MID,
                                    font=("Calibri", 10))
        self._st1_label.pack(side=tk.LEFT, padx=(8, 20))

    def _build_notebook(self) -> None:
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        c, b = self._controller, self._bus
        self._timetable_tab    = TimetableTab(nb, c, b)
        self._verification_tab = VerificationTab(nb, c, b)
        self._exam_tab         = ExamTab(nb, c, b)
        nb.add(self._timetable_tab,    text="  Timetable  ")
        nb.add(self._verification_tab, text="  Verification  ")
        nb.add(self._exam_tab,         text="  Exams  ")

    def _auto_load(self) -> None:
        if DEFAULT_STATE_PATH.exists():
            try:
                self._controller.load_from_json(str(DEFAULT_STATE_PATH))
            except Exception as e:
                messagebox.showerror("Auto-load Error", str(e))
            return
        for candidate in ("data/ST1 2026.xlsx", "data/ST12026.xlsx",
                           "data/ST1_2026.xlsx", "data/ST1.xlsx"):
            p = Path(candidate)
            if p.exists():
                try:
                    self._controller.load_from_excel(str(p))
                except Exception as e:
                    messagebox.showerror("Auto-load Error", str(e))
                return

    def _load_st1(self) -> None:
        path = filedialog.askopenfilename(
            title="Open timetable spreadsheet",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not path:
            return
        try:
            self._controller.load_from_excel(path)
            self._st1_label.config(text=Path(path).name, fg=CLR_WHITE)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def _on_data_loaded(self, **_) -> None:
        self._st1_label.config(text="Timetable loaded", fg=CLR_WHITE)

    def _on_close(self) -> None:
        if self._state.timetable_tree is None:
            self.destroy()
            return
        win = tk.Toplevel(self)
        win.title("Save before exit?")
        win.configure(bg=CLR_WHITE)
        win.resizable(False, False)
        win.grab_set()
        win.focus_force()
        tk.Label(win, text="Save current timetable and exam state\nbefore closing?",
                 bg=CLR_WHITE, font=("Calibri", 10),
                 justify=tk.CENTER).pack(padx=24, pady=(18, 12))
        btn_row = tk.Frame(win, bg=CLR_WHITE)
        btn_row.pack(padx=24, pady=(0, 18))

        def save_and_exit():
            win.destroy()
            DEFAULT_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
            try:
                self._controller.save_to_json(str(DEFAULT_STATE_PATH))
                self.destroy()
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

        def exit_no_save():
            win.destroy()
            self.destroy()

        tk.Button(btn_row, text="Save & Exit", command=save_and_exit,
                  bg=CLR_GREEN, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 9, "bold"), padx=12).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_row, text="Exit without saving", command=exit_no_save,
                  bg=CLR_RED, fg=CLR_WHITE, relief=tk.FLAT,
                  font=("Calibri", 9), padx=12).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_row, text="Cancel", command=win.destroy,
                  bg=CLR_LIGHT, relief=tk.FLAT,
                  font=("Calibri", 9), padx=12).pack(side=tk.LEFT, padx=4)
