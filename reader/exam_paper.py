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

    # ── Build ──────────────────────────────────────────────────────────────

    @classmethod
    def from_exam_tree(cls, tree: ExamTree,
                       exclusions: set[str] | None = None) -> "ExamPaperRegistry":
        """
        Build a registry from an ExamTree.
        Each subject+grade becomes one P1 paper.
        Excluded subject codes (e.g. {"ST", "LIB", "PE", "RDI"}) are skipped.
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
        return reg

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

    # ── Grade / subject listing ────────────────────────────────────────────

    def grades(self) -> list[str]:
        return sorted({p.grade for p in self._papers.values()})

    def subjects_for_grade(self, grade: str) -> list[str]:
        return sorted({p.subject for p in self._papers.values()
                       if p.grade == grade})
