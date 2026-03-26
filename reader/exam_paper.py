"""
exam_paper.py
-------------
ExamPaper — one schedulable exam event (subject + paper number + grade).
ExamPaperRegistry — holds all papers, built from an ExamTree.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
from reader.exam_tree import ExamTree


@dataclass
class ExamPaper:
    grade        : str        # e.g. "Gr12"
    subject      : str        # e.g. "MA"
    paper_number : int        # 1, 2, or 3
    student_ids  : set[int] = field(default_factory=set)
    constraints  : set[str] = field(default_factory=set)
    links        : set[str] = field(default_factory=set)  # labels of linked partner papers
    pinned_slot  : Optional[int] = None   # slot index, None = free

    @property
    def label(self) -> str:
        return f"{self.subject}_P{self.paper_number}_{self.grade}"

    def student_count(self) -> int:
        return len(self.student_ids)


class ExamPaperRegistry:
    """
    Holds all ExamPaper objects keyed by label.
    Built from an ExamTree; every subject starts with one paper (P1).
    Papers can be added or removed per grade independently.
    """

    def __init__(self):
        # key = paper label (e.g. "MA_P1_Gr12"), value = ExamPaper
        self._papers: dict[str, ExamPaper] = {}
        # difficulty: "red" | "yellow" | "green", default "green"
        # key = f"{subject}_{grade}" e.g. "MA_Gr12"
        self._difficulty: dict[str, str] = {}

    # ── Difficulty ────────────────────────────────────────────────────────

    def get_difficulty(self, subject: str, grade: str) -> str:
        return self._difficulty.get(f"{subject}_{grade}", "green")

    def set_difficulty(self, subject: str, grade: str, value: str) -> None:
        if value not in ("red", "yellow", "green"):
            raise ValueError(f"Invalid difficulty: {value!r}")
        self._difficulty[f"{subject}_{grade}"] = value

    # ── Build ──────────────────────────────────────────────────────────────

    @classmethod
    def from_exam_tree(
        cls,
        tree: ExamTree,
        exclusions: set[str] | None = None,
        prior: "ExamPaperRegistry | None" = None,
    ) -> "ExamPaperRegistry":
        """
        Build a registry from an ExamTree.
        Each subject+grade becomes one P1 paper.
        Excluded subject codes (e.g. {"ST", "LIB", "PE", "RDI"}) are skipped.

        If *prior* is provided, copies constraints, difficulty, links,
        pinned slots, and extra papers (P2/P3) from the prior registry
        for any subject+grade that still exists in the new tree.
        """
        reg = cls()
        excl = exclusions or set()
        for grade_label, grade_node in tree.grades.items():
            for subj_label, exam_subject in grade_node.exam_subjects.items():
                subject_code = subj_label.split("_")[0]
                if subject_code in excl or subj_label in excl:
                    continue
                paper = ExamPaper(
                    grade        = grade_label,
                    subject      = subject_code,
                    paper_number = 1,
                    student_ids  = exam_subject.all_students(),
                )
                reg._papers[paper.label] = paper

        if prior:
            reg._copy_from_prior(prior)

        return reg

    def _copy_from_prior(self, prior: "ExamPaperRegistry") -> None:
        """Copy difficulty, links, constraints, pinned slots, and extra papers from a prior registry."""
        # Copy difficulty settings
        for key, value in prior._difficulty.items():
            # key = "SUBJ_Grnn" — only copy if subject+grade still exists
            parts = key.split("_", 1)
            if len(parts) == 2:
                subj, grade = parts
                if self.papers_for_subject_grade(subj, grade):
                    self._difficulty[key] = value

        # Copy per-paper state and add extra papers
        for label, old_paper in prior._papers.items():
            if old_paper.subject == "ST":
                # Carry over study papers directly
                self._papers[label] = ExamPaper(
                    grade        = old_paper.grade,
                    subject      = "ST",
                    paper_number = old_paper.paper_number,
                    student_ids  = set(),
                    constraints  = set(old_paper.constraints),
                    links        = set(old_paper.links),
                    pinned_slot  = old_paper.pinned_slot,
                )
                continue

            if label in self._papers:
                # P1 exists in new registry — copy constraints, links, pinned
                new_paper = self._papers[label]
                new_paper.constraints = set(old_paper.constraints)
                new_paper.links = set(old_paper.links)
                new_paper.pinned_slot = old_paper.pinned_slot
            elif old_paper.paper_number > 1:
                # Extra paper (P2/P3) — add if the subject+grade still exists
                siblings = self.papers_for_subject_grade(old_paper.subject, old_paper.grade)
                if siblings:
                    p1 = next((p for p in siblings if p.paper_number == 1), None)
                    if p1:
                        new_paper = ExamPaper(
                            grade        = old_paper.grade,
                            subject      = old_paper.subject,
                            paper_number = old_paper.paper_number,
                            student_ids  = set(p1.student_ids),
                            constraints  = set(old_paper.constraints),
                            links        = set(old_paper.links),
                            pinned_slot  = old_paper.pinned_slot,
                        )
                        self._papers[new_paper.label] = new_paper

        # Clean up dangling link references — remove links to papers that
        # no longer exist (ghost subjects that disappeared from the tree)
        valid_labels = set(self._papers.keys())
        for paper in self._papers.values():
            paper.links = paper.links & valid_labels

    # ── Queries ────────────────────────────────────────────────────────────

    def all_papers(self) -> list[ExamPaper]:
        return list(self._papers.values())

    def get(self, label: str) -> ExamPaper | None:
        return self._papers.get(label)

    def papers_for_grade(self, grade: str) -> list[ExamPaper]:
        return [p for p in self._papers.values() if p.grade == grade]

    def papers_for_subject_grade(self, subject: str, grade: str) -> list[ExamPaper]:
        return sorted(
            [p for p in self._papers.values()
             if p.subject == subject and p.grade == grade],
            key=lambda p: p.paper_number,
        )

    # ── Mutations ──────────────────────────────────────────────────────────

    def add_paper(self, subject: str, grade: str) -> ExamPaper | None:
        """
        Add the next paper number (P2 or P3) for subject+grade.
        Returns the new paper, or None if P3 already exists.
        Duplicates the student set from P1.
        """
        existing = self.papers_for_subject_grade(subject, grade)
        if not existing:
            return None
        next_num = max(p.paper_number for p in existing) + 1
        if next_num > 3:
            return None
        p1 = next(p for p in existing if p.paper_number == 1)
        new_paper = ExamPaper(
            grade        = grade,
            subject      = subject,
            paper_number = next_num,
            student_ids  = set(p1.student_ids),
        )
        self._papers[new_paper.label] = new_paper
        return new_paper

    def remove_paper(self, label: str) -> bool:
        """
        Remove a paper by label.
        Cannot remove P1 if it is the only paper for that subject+grade.
        Returns True if removed, False otherwise.
        """
        paper = self._papers.get(label)
        if paper is None:
            return False
        siblings = self.papers_for_subject_grade(paper.subject, paper.grade)
        if paper.paper_number == 1 and len(siblings) == 1:
            return False
        del self._papers[label]
        return True

    def add_constraint(self, label: str, code: str) -> bool:
        paper = self._papers.get(label)
        if paper is None:
            return False
        paper.constraints.add(code.strip().upper())
        return True

    def remove_constraint(self, label: str, code: str) -> bool:
        paper = self._papers.get(label)
        if paper is None:
            return False
        paper.constraints.discard(code.strip().upper())
        return True

    def add_study_paper(self, grade: str, pinned_slot: int | None = None) -> ExamPaper:
        """
        Add a Study (ST) paper for a grade.
        Study papers have no student IDs, are always pinned, and act as
        hard slot blockers for their grade. Label: ST_P{n}_Gr{grade}.
        """
        existing = self.papers_for_subject_grade("ST", grade)
        next_num = max((p.paper_number for p in existing), default=0) + 1
        paper = ExamPaper(
            grade        = grade,
            subject      = "ST",
            paper_number = next_num,
            student_ids  = set(),
            pinned_slot  = pinned_slot,
        )
        self._papers[paper.label] = paper
        return paper

    def remove_study_paper(self, label: str) -> bool:
        """Remove a study paper by label. Only ST papers can be removed this way."""
        paper = self._papers.get(label)
        if paper is None or paper.subject != "ST":
            return False
        del self._papers[label]
        return True

    # ── Grade / subject listing ────────────────────────────────────────────

    def grades(self) -> list[str]:
        return sorted({p.grade for p in self._papers.values()})

    def subjects_for_grade(self, grade: str) -> list[str]:
        return sorted({p.subject for p in self._papers.values()
                       if p.grade == grade})
