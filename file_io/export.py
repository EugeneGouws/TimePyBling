from __future__ import annotations

from datetime import date
from pathlib import Path

from reader.exam_scheduler import ScheduleResult


def to_pdf(
    path: Path,
    schedule_result: ScheduleResult,
    grades: list[str],
    grid: dict[int, dict[str, list[str]]],
    slot_meta: dict[int, tuple[date, str]],
    all_slots: list[int],
) -> None:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.lib import colors

    doc = SimpleDocTemplate(
        str(path), pagesize=landscape(A4),
        leftMargin=18, rightMargin=18,
        topMargin=18, bottomMargin=18,
    )
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
    am_bg  = colors.Color(1.0,  0.95, 0.88)
    pm_bg  = colors.Color(0.99, 0.89, 0.93)
    hdr_bg = colors.Color(0.86, 0.91, 0.99)
    style_cmds = [
        ("FONTNAME",   (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTNAME",   (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",   (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0),  hdr_bg),
        ("TEXTCOLOR",  (0, 0), (-1, 0),  colors.Color(0.12, 0.23, 0.37)),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.Color(0.7, 0.7, 0.8)),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",      (1, 1), (-1, -1), "CENTER"),
    ]
    for ri, slot_idx in enumerate(all_slots, start=1):
        _, session = slot_meta[slot_idx]
        bg = am_bg if session == "AM" else pm_bg
        style_cmds.append(("BACKGROUND", (0, ri), (-1, ri), bg))
    table.setStyle(TableStyle(style_cmds))
    doc.build([table])


def to_txt(
    path: Path,
    schedule_result: ScheduleResult,
    grades: list[str],
    grid: dict[int, dict[str, list[str]]],
    slot_meta: dict[int, tuple[date, str]],
    all_slots: list[int],
) -> None:
    col_w = 8
    header = f"{'Slot':<4}  {'Date':<14}  {'Ses':<3} | " + " | ".join(
        f"{g:^{col_w}}" for g in grades
    )
    separator = "-" * len(header)
    lines = [header, separator]
    for slot_idx in all_slots:
        d, session = slot_meta[slot_idx]
        label = f"{slot_idx+1:<4}  {d.strftime('%a %d %b %Y'):<14}  {session:<3}"
        cols = " | ".join(
            f"{', '.join(sorted(grid[slot_idx][g])):^{col_w}}" if grid[slot_idx][g] else " " * col_w
            for g in grades
        )
        lines.append(f"{label} | {cols}")

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")
