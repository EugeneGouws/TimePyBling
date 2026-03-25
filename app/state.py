from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from core.timetable_tree import TimetableTree
from reader.exam_tree import ExamTree
from reader.exam_paper import ExamPaperRegistry
from reader.exam_scheduler import ScheduleResult

DEFAULT_EXCLUSIONS: frozenset[str] = frozenset({"ST", "LIB", "PE", "RDI"})


@dataclass
class SessionConfig:
    start: date
    end: date
    am: bool = True
    pm: bool = True


@dataclass
class AppState:
    timetable_tree: Optional[TimetableTree] = None
    exam_tree: Optional[ExamTree] = None
    paper_registry: Optional[ExamPaperRegistry] = None
    schedule_result: Optional[ScheduleResult] = None
    exclusions: set[str] = field(default_factory=lambda: set(DEFAULT_EXCLUSIONS))
    session_config: Optional[SessionConfig] = None
