from __future__ import annotations

import json
from dataclasses import asdict
from datetime import date
from pathlib import Path

from core.timetable_tree import timetable_tree_to_dict, timetable_tree_from_dict
from app.state import AppState, SessionConfig
from app.cost_config import CostConfig


class StateRepository:
    def __init__(self) -> None:
        self.pending_papers: dict[str, dict | list] = {}

    def save(self, path: Path, state: AppState) -> None:
        papers: dict[str, dict] = {}
        if state.paper_registry:
            for grade in state.paper_registry.grades():
                grade_num = grade.replace("Gr", "")
                for subj in state.paper_registry.subjects_for_grade(grade):
                    ps = state.paper_registry.papers_for_subject_grade(subj, grade)
                    key = f"{subj}_{grade_num}"
                    # Collect constraints and links across all papers for this subject+grade
                    all_constraints: set[str] = set()
                    all_links: set[str] = set()
                    pinned: int | None = None
                    for p in ps:
                        all_constraints |= p.constraints
                        all_links |= p.links
                        if p.pinned_slot is not None:
                            pinned = p.pinned_slot
                    papers[key] = {
                        "papers": [f"P{p.paper_number}" for p in ps],
                        "constraints": sorted(all_constraints),
                        "difficulty": state.paper_registry.get_difficulty(subj, grade),
                        "links": sorted(all_links),
                        "pinned_slot": pinned,
                    }

        cfg = state.session_config
        cost = state.cost_config
        payload = {
            "timetable_tree": (
                timetable_tree_to_dict(state.timetable_tree)
                if state.timetable_tree else None
            ),
            "exclusions": sorted(state.exclusions),
            "cost_config": asdict(cost),
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

        if "cost_config" in data:
            state.cost_config = CostConfig(**data["cost_config"])

        self.pending_papers = dict(data.get("papers", {}))

    def apply_pending_papers(self, state: AppState) -> None:
        """
        Apply paper config stored from the last load() call to an existing registry.
        Handles both old format (list of paper names) and new format (dict with metadata).
        """
        if state.paper_registry is None or not self.pending_papers:
            return
        for subj_grade, entry in self.pending_papers.items():
            parts = subj_grade.split("_")
            if len(parts) != 2:
                continue
            subj, grade_num = parts
            grade = f"Gr{grade_num}"

            # Migration: old format is a plain list ["P1", "P2"]
            if isinstance(entry, list):
                paper_names = entry
                constraints: list[str] = []
                difficulty = "green"
                links: list[str] = []
                pinned_slot = None
            else:
                paper_names = entry.get("papers", ["P1"])
                constraints = entry.get("constraints", [])
                difficulty = entry.get("difficulty", "green")
                links = entry.get("links", [])
                pinned_slot = entry.get("pinned_slot")

            # Handle study papers
            if subj == "ST":
                state.paper_registry.add_study_paper(grade, pinned_slot=pinned_slot)
                continue

            # Add extra papers (P2, P3)
            max_num = max(
                (int(p[1:]) for p in paper_names if p.startswith("P") and p[1:].isdigit()),
                default=1,
            )
            for _ in range(max_num - 1):
                state.paper_registry.add_paper(subj, grade)

            # Apply constraints to all papers for this subject+grade
            for code in constraints:
                for paper in state.paper_registry.papers_for_subject_grade(subj, grade):
                    paper.constraints.add(code)

            # Apply difficulty
            if difficulty != "green":
                state.paper_registry.set_difficulty(subj, grade, difficulty)

            # Apply links
            for link_label in links:
                for paper in state.paper_registry.papers_for_subject_grade(subj, grade):
                    paper.links.add(link_label)

            # Apply pinned slot
            if pinned_slot is not None:
                papers = state.paper_registry.papers_for_subject_grade(subj, grade)
                if papers:
                    papers[0].pinned_slot = pinned_slot
