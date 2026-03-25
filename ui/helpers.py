import tkinter as tk
from tkinter import ttk

from ui.constants import CLR_WHITE


def _scrolled_text(parent, **kw) -> tk.Text:
    frame = tk.Frame(parent, bg=CLR_WHITE)
    frame.pack(fill=tk.BOTH, expand=True)
    sb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    t = tk.Text(frame, font=("Calibri", 10), relief=tk.FLAT,
                bg="#F8F9FF", state=tk.DISABLED, wrap=tk.WORD,
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


def student_display(tree, student_id: int) -> str:
    """Return 'Firstname Lastname (ID)' if name known, else just str(ID)."""
    names = getattr(tree, "student_names", {})
    name = names.get(student_id)
    if name:
        return f"{name} ({student_id})"
    return str(student_id)
