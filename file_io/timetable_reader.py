from pathlib import Path

from core.timetable_tree import TimetableTree, build_timetable_tree_from_file


def read_timetable(path: Path) -> TimetableTree:
    return build_timetable_tree_from_file(path)
