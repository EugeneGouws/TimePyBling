from __future__ import annotations

from typing import Optional

from app.events import EventBus
from app.state import AppState, SessionConfig
from reader.exam_tree import build_exam_tree_from_timetable_tree
from reader.exam_paper import ExamPaperRegistry
from reader.exam_scheduler import build_schedule, ScheduleResult, PenaltyEntry

# Event name constants
EVT_TIMETABLE_LOADED    = "timetable_loaded"
EVT_EXAM_TREE_BUILT     = "exam_tree_built"
EVT_REGISTRY_BUILT      = "registry_built"
EVT_PAPERS_CHANGED      = "papers_changed"
EVT_SCHEDULE_GENERATED  = "schedule_generated"
EVT_STATE_SAVED         = "state_saved"
EVT_STATE_LOADED        = "state_loaded"
EVT_EXCLUSIONS_CHANGED  = "exclusions_changed"
EVT_SESSION_CFG_CHANGED = "session_config_changed"


class AppController:
    def __init__(self, state: AppState, bus: EventBus) -> None:
        self._state = state
        self._bus = bus

    @property
    def state(self) -> AppState:
        return self._state

    # ------------------------------------------------------------------
    # File I/O  (file_io/ imports are lazy)
    # ------------------------------------------------------------------

    def load_from_excel(self, path: str) -> None:
        from pathlib import Path
        from file_io.timetable_reader import read_timetable  # noqa: PLC0415
        self._state.timetable_tree = read_timetable(Path(path))
        self.build_exam_tree()
        self.build_registry()
        self._bus.publish(EVT_TIMETABLE_LOADED, state=self._state)

    def load_from_json(self, path: str) -> None:
        from pathlib import Path
        from file_io.state_repo import StateRepository  # noqa: PLC0415
        repo = StateRepository()
        repo.load(Path(path), self._state)
        if self._state.timetable_tree is not None:
            self.build_exam_tree()
            self.build_registry()
            repo.apply_pending_papers(self._state)
        self._bus.publish(EVT_STATE_LOADED, state=self._state)

    def save_to_json(self, path: str) -> None:
        from pathlib import Path
        from file_io.state_repo import StateRepository  # noqa: PLC0415
        StateRepository().save(Path(path), self._state)
        self._bus.publish(EVT_STATE_SAVED, path=path)

    # ------------------------------------------------------------------
    # Domain orchestration
    # ------------------------------------------------------------------

    def build_exam_tree(self) -> None:
        if self._state.timetable_tree is None:
            raise ValueError("timetable_tree is not loaded")
        self._state.exam_tree = build_exam_tree_from_timetable_tree(
            self._state.timetable_tree
        )
        self._bus.publish(EVT_EXAM_TREE_BUILT, state=self._state)

    def build_registry(self) -> None:
        if self._state.exam_tree is None:
            raise ValueError("exam_tree is not built")
        self._state.paper_registry = ExamPaperRegistry.from_exam_tree(
            self._state.exam_tree,
            exclusions=self._state.exclusions,
            prior=self._state.paper_registry,
        )
        self._bus.publish(EVT_REGISTRY_BUILT, state=self._state)

    # ------------------------------------------------------------------
    # Paper / constraint mutations
    # ------------------------------------------------------------------

    def add_paper(self, subject: str, grade: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.add_paper(subject, grade)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def remove_paper(self, label: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.remove_paper(label)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def add_constraint(self, label: str, code: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.add_constraint(label, code)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def remove_constraint(self, label: str, code: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.remove_constraint(label, code)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def set_difficulty(self, subject: str, grade: str, difficulty: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.set_difficulty(subject, grade, difficulty)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def add_link(self, label_a: str, label_b: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        paper_a = self._state.paper_registry.get(label_a)
        paper_b = self._state.paper_registry.get(label_b)
        if paper_a is None or paper_b is None:
            raise ValueError(f"Paper not found: {label_a} or {label_b}")
        paper_a.links.add(label_b)
        paper_b.links.add(label_a)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def remove_link(self, label_a: str, label_b: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        paper_a = self._state.paper_registry.get(label_a)
        paper_b = self._state.paper_registry.get(label_b)
        if paper_a:
            paper_a.links.discard(label_b)
        if paper_b:
            paper_b.links.discard(label_a)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def add_study_paper(self, grade: str, pinned_slot: int | None = None) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.add_study_paper(grade, pinned_slot)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    def remove_study_paper(self, label: str) -> None:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        self._state.paper_registry.remove_study_paper(label)
        self._bus.publish(EVT_PAPERS_CHANGED, state=self._state)

    # ------------------------------------------------------------------
    # Configuration
    # ------------------------------------------------------------------

    def set_exclusions(self, exclusions: set[str]) -> None:
        self._state.exclusions = exclusions
        self._bus.publish(EVT_EXCLUSIONS_CHANGED, state=self._state)
        if self._state.exam_tree is not None:
            self.build_registry()

    def set_session_config(self, config: SessionConfig) -> None:
        self._state.session_config = config
        self._bus.publish(EVT_SESSION_CFG_CHANGED, state=self._state)

    # ------------------------------------------------------------------
    # Verification queries  (read-only, no state mutation)
    # ------------------------------------------------------------------

    def find_clashes(self) -> tuple[list, list]:
        """Return (student_clashes, teacher_clashes). teacher list is always []."""
        if self._state.timetable_tree is None:
            return [], []
        from reader.verify_timetable import find_student_clashes  # noqa: PLC0415
        return find_student_clashes(self._state.timetable_tree), []

    def data_integrity_issues(self) -> list[dict]:
        """Return classes with fewer than 5 students."""
        if self._state.timetable_tree is None:
            return []
        from reader.verify_timetable import data_integrity_issues  # noqa: PLC0415
        return data_integrity_issues(self._state.timetable_tree)

    def schedulable_pairs(self) -> dict[str, list[tuple[str, str]]]:
        """
        Return per-grade dict of (subj_a, subj_b) pairs that share no students.
        Subjects in state.exclusions are ignored.
        """
        if self._state.exam_tree is None:
            return {}
        from core.conflict_matrix import ConflictMatrix          # noqa: PLC0415
        from reader.exam_clash import is_excluded as _excl       # noqa: PLC0415
        result: dict[str, list[tuple[str, str]]] = {}
        for grade_label, grade_node in sorted(self._state.exam_tree.grades.items()):
            groups = {
                subj_label: subject.all_students()
                for subj_label, subject in grade_node.exam_subjects.items()
                if not _excl(subj_label, self._state.exclusions)
            }
            if not groups:
                continue
            pairs = ConflictMatrix(grade_label, groups).free_pairs()
            if pairs:
                result[grade_label] = pairs
        return result

    # ------------------------------------------------------------------
    # Exam scheduling helpers  (read-only, no state mutation)
    # ------------------------------------------------------------------

    def needed_slots_per_grade(self) -> dict[str, int]:
        """Return minimum slots needed per grade via DSatur colouring."""
        if self._state.paper_registry is None:
            return {}
        from reader.exam_clash import build_clash_graph, dsatur_colouring  # noqa: PLC0415
        result: dict[str, int] = {}
        for grade in self._state.paper_registry.grades():
            papers = self._state.paper_registry.papers_for_grade(grade)
            if not papers:
                continue
            student_sets = {p.label: p.student_ids for p in papers}
            graph = build_clash_graph(student_sets)
            colouring = dsatur_colouring(graph)
            result[grade] = (max(colouring.values()) + 1) if colouring else 0
        return result

    def effective_sessions(
        self,
        start_str: str,
        end_str: str,
        am: bool,
        pm: bool,
    ) -> "list[tuple] | None":
        """
        Return list of (date, session) tuples for exam weekdays in [start, end].
        Returns None if dates are invalid or end < start.
        """
        from datetime import date, timedelta              # noqa: PLC0415
        from reader.exam_scheduler import EXAM_WEEKDAYS  # noqa: PLC0415
        try:
            start = date.fromisoformat(start_str.strip())
            end   = date.fromisoformat(end_str.strip())
        except ValueError:
            return None
        if end < start:
            return None
        days: list[date] = []
        d = start
        while d <= end:
            if d.weekday() in EXAM_WEEKDAYS:
                days.append(d)
            d += timedelta(days=1)
        sessions: list[tuple] = []
        for day in days:
            if am:
                sessions.append((day, "AM"))
            if pm:
                sessions.append((day, "PM"))
        return sessions

    def export_pdf(
        self,
        path: str,
        grades: list,
        grid: dict,
        slot_meta: dict,
        all_slots: list,
    ) -> None:
        """Write schedule as PDF via file_io.export. Raises ImportError if reportlab missing."""
        from pathlib import Path                          # noqa: PLC0415
        from file_io.export import to_pdf                # noqa: PLC0415
        to_pdf(Path(path), self._state.schedule_result,
               grades, grid, slot_meta, all_slots)

    def export_txt(
        self,
        path: str,
        grades: list,
        grid: dict,
        slot_meta: dict,
        all_slots: list,
    ) -> None:
        """Write schedule as pipe-delimited text via file_io.export."""
        from pathlib import Path                          # noqa: PLC0415
        from file_io.export import to_txt                # noqa: PLC0415
        to_txt(Path(path), self._state.schedule_result,
               grades, grid, slot_meta, all_slots)

    # ------------------------------------------------------------------
    # Schedule generation
    # ------------------------------------------------------------------

    def generate_schedule(
        self,
        sessions: Optional[list[tuple]] = None,
    ) -> ScheduleResult:
        if self._state.paper_registry is None:
            raise ValueError("paper_registry is not built")
        result = build_schedule(
            self._state.paper_registry,
            sessions=sessions,
            exam_tree=self._state.exam_tree,
            config=self._state.cost_config,
        )
        self._state.schedule_result = result
        self._bus.publish(EVT_SCHEDULE_GENERATED, state=self._state)
        return result
