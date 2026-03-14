"""
xlsx_loader.py
--------------
Generic xlsx row parser reused by all ingestion scripts.

Provides a single function load_rows() that:
  - Opens any xlsx file
  - Optionally detects and skips a header row
  - Returns rows as plain lists, filtered and cleaned
  - Accepts a row_parser callable to convert raw rows into objects

Usage:
    from xlsx_loader import load_rows

    def parse_student(row):
        # row is a list of cleaned cell values
        ...

    students = load_rows("data/students.xlsx", row_parser=parse_student)
"""

import openpyxl
from typing import Callable, Any


def load_rows(
    filepath: str,
    row_parser: Callable[[list], Any],
    sheet_name: str | None = None,
    id_column: int = 0,
    skip_non_numeric_id: bool = True,
) -> list[Any]:
    """
    Generic xlsx row loader.

    Parameters
    ----------
    filepath             : path to the xlsx file
    row_parser           : callable that receives a cleaned row (list of values)
                           and returns an object, or None to skip the row
    sheet_name           : sheet to load, defaults to active sheet
    id_column            : column index used to detect header rows
    skip_non_numeric_id  : if True, rows where id_column is non-numeric are
                           skipped (handles header rows automatically)

    Returns
    -------
    List of objects returned by row_parser (None results are excluded)
    """
    wb = openpyxl.load_workbook(filepath)
    ws = wb[sheet_name] if sheet_name else wb.active

    results = []
    skipped = 0

    for raw_row in ws.iter_rows(values_only=True):
        # Skip completely empty rows
        if not any(raw_row):
            continue

        # Clean each cell: strip whitespace, normalise empty to None
        row = [
            str(cell).strip() if cell is not None else None
            for cell in raw_row
        ]

        # Detect and skip header rows via non-numeric id column
        if skip_non_numeric_id:
            try:
                int(row[id_column])
            except (TypeError, ValueError, IndexError):
                skipped += 1
                continue

        result = row_parser(row)
        if result is not None:
            results.append(result)
        else:
            skipped += 1

    if skipped:
        print(f"[xlsx_loader] {filepath}: skipped {skipped} rows.")

    return results


def collect_subjects(row: list, start_col: int) -> list[str]:
    """
    Helper: extract non-empty subject codes from a row
    starting at start_col, uppercased.
    Used by both student and teacher parsers.
    """
    return [
        cell.upper()
        for cell in row[start_col:]
        if cell is not None and cell.strip() != ""
    ]