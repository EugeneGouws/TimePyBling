from __future__ import annotations

import json
from datetime import date
from pathlib import Path

from core.timetable_tree import timetable_tree_to_dict, timetable_tree_from_dict
from app.state import AppState, SessionConfig


class StateRepository:
    def __init__(self) -> None:
        self.pending_papers: dict[str, list[str]] = {}

    def save(self, path: Path, state: AppState) -> None:
        papers: dict[str, list[str]] = {}
        if state.paper_registry:
            for grade in state.paper_registry.grades():
                grade_num = grade.replace("Gr", "")
                for subj in state.paper_registry.subjects_for_grade(grade):
                    ps = state.paper_registry.papers_for_subject_grade(subj, grade)
                    papers[f"{subj}_{grade_num}"] = [f"P{p.paper_number}" for p in ps]

        cfg = state.session_config
        payload = {
            "timetable_tree": (
                timetable_tree_to_dict(state.timetable_tree)
                if state.timetable_tree else None
            ),
            "exclusions": sorted(state.exclusions),
            "papers": papers,
            "session": {
                "start": cfg.start.isoformat() if cfg else None,
                "end":   cfg.end.isoformat()   if cfg else None,
                "am":    cfg.am                 if cfg else True,
                "pm":    cfg.pm                 if cfg else True,
            },
        }
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)

    def load(self, path: Path, state: AppState) -> None:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)

        if data.get("timetable_tree"):
            state.timetable_tree = timetable_tree_from_dict(data["timetable_tree"])

        if "exclusions" in data:
            state.exclusions = set(data["exclusions"])

        if "session" in data:
            s = data["session"]
            start = date.fromisoformat(s["start"]) if s.get("start") else date.today()
            end   = date.fromisoformat(s["end"])   if s.get("end")   else date.today()
            state.session_config = SessionConfig(
                start=start,
                end=end,
                am=bool(s.get("am", True)),
                pm=bool(s.get("pm", True)),
            )

        self.pending_papers = dict(data.get("papers", {}))

    def apply_pending_papers(self, state: AppState) -> None:
        """Apply P2/P3 papers stored from the last load() call to an existing registry."""
        if state.paper_registry is None or not self.pending_papers:
            return
        for subj_grade, paper_nums in self.pending_papers.items():
            parts = subj_grade.split("_")
            if len(parts) != 2:
                continue
            subj, grade_num = parts
            grade = f"Gr{grade_num}"
            max_num = max(
                (int(p[1:]) for p in paper_nums if p.startswith("P") and p[1:].isdigit()),
                default=1,
            )
            for _ in range(max_num - 1):
                state.paper_registry.add_paper(subj, grade)
