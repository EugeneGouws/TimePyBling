"""
conflict_matrix.py
------------------
Pure conflict logic — completely context-free.

A ConflictMatrix takes any mapping of:
    { group_name: set_of_member_ids }

and determines which pairs of groups share at least one member.
Shared members = conflict = groups cannot occupy the same slot.

This is reused identically by:
    - Timetable scheduling  (groups = subject enrolments per grade)
    - Exam scheduling       (groups = subject enrolments across grades)
    - Any future context    (sports, co-curricular, venue allocation)
"""

from itertools import combinations


class ConflictMatrix:
    """
    Context-free conflict matrix.

    Parameters
    ----------
    label  : human-readable label (e.g. grade name, exam session)
    groups : dict mapping group_name -> set of member IDs
             e.g. {"MA": {"480","481","482"}, "SC": {"480","483"}, ...}
    """

    def __init__(self, label: str, groups: dict[str, set[str]]):
        self.label    = label
        self.subjects = sorted(groups.keys())

        # conflict_map[a][b] = True  →  a and b share at least one member
        self.conflict_map: dict[str, dict[str, bool]] = {
            s: {t: False for t in self.subjects} for s in self.subjects
        }

        # shared[a][b] = set of member IDs in both groups
        self.shared: dict[str, dict[str, set]] = {
            s: {t: set() for t in self.subjects} for s in self.subjects
        }

        self._build(groups)

    def _build(self, groups: dict[str, set[str]]):
        for a, b in combinations(self.subjects, 2):
            shared = groups.get(a, set()) & groups.get(b, set())
            if shared:
                self.conflict_map[a][b] = True
                self.conflict_map[b][a] = True
                self.shared[a][b] = shared
                self.shared[b][a] = shared

    # ------------------------------------------------------------------
    # Queries
    # ------------------------------------------------------------------

    def conflicts(self, a: str, b: str) -> bool:
        """True if groups a and b share at least one member."""
        return self.conflict_map.get(a, {}).get(b, False)

    def conflicts_with(self, subject: str) -> list[str]:
        """All groups that conflict with the given group."""
        return [
            s for s in self.subjects
            if s != subject and self.conflict_map[subject].get(s, False)
        ]

    def free_partners(self, subject: str) -> list[str]:
        """All groups that can share a slot with the given group."""
        return [
            s for s in self.subjects
            if s != subject and not self.conflict_map[subject].get(s, False)
        ]

    def conflict_pairs(self) -> list[tuple[str, str]]:
        """All pairs that cannot share a slot."""
        seen, pairs = set(), []
        for a in self.subjects:
            for b in self.subjects:
                if a != b and self.conflict_map[a][b]:
                    key = tuple(sorted([a, b]))
                    if key not in seen:
                        seen.add(key)
                        pairs.append(key)
        return sorted(pairs)

    def free_pairs(self) -> list[tuple[str, str]]:
        """All pairs that CAN share a slot."""
        seen, pairs = set(), []
        for a in self.subjects:
            for b in self.subjects:
                if a != b and not self.conflict_map[a][b]:
                    key = tuple(sorted([a, b]))
                    if key not in seen:
                        seen.add(key)
                        pairs.append(key)
        return sorted(pairs)

    def degrees(self) -> dict[str, int]:
        """
        Number of conflicts per group.
        Higher degree = more constrained = schedule first.
        """
        return {
            s: sum(1 for t in self.subjects if t != s and self.conflict_map[s][t])
            for s in self.subjects
        }

    def ordering(self) -> list[str]:
        """
        Groups sorted most-constrained first.
        Recommended scheduling order for backtracking / CSP search.
        """
        d = self.degrees()
        return sorted(self.subjects, key=lambda s: -d[s])

    def shared_members(self, a: str, b: str) -> set[str]:
        """Return the set of member IDs shared between groups a and b."""
        return self.shared.get(a, {}).get(b, set())

    # ------------------------------------------------------------------
    # Display
    # ------------------------------------------------------------------

    def print_matrix(self):
        col_w = 6
        print(f"\n  Conflict matrix — {self.label}")
        header = f"  {'':12}" + "".join(f"{s:>{col_w}}" for s in self.subjects)
        print(header)
        print(f"  {'-' * (12 + col_w * len(self.subjects))}")

        for a in self.subjects:
            row = f"  {a:<12}"
            for b in self.subjects:
                if a == b:
                    row += f"{'—':>{col_w}}"
                elif self.conflict_map[a][b]:
                    row += f"{'X':>{col_w}}"
                else:
                    row += f"{'·':>{col_w}}"
            print(row)

        d = self.degrees()
        print(f"\n  Scheduling order (most constrained first):")
        for s in self.ordering():
            fp = self.free_partners(s)
            print(f"    {s:<8} degree={d[s]:<3}  "
                  f"free with: {fp if fp else '—'}")

        fp = self.free_pairs()
        if fp:
            print(f"\n  Free pairs (can share a slot):")
            for a, b in fp:
                n = len(self.shared_members(a, b))
                print(f"    {a} + {b}")
        else:
            print(f"\n  No free pairs — all subjects conflict with each other.")