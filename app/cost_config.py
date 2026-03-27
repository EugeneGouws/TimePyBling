from __future__ import annotations

from dataclasses import dataclass


@dataclass
class CostConfig:
    # Soft constraint weights
    same_week_penalty:      int = 1    # per same-week pair (subject+grade)
    teacher_load_penalty:   int = 1    # per teacher clash slot
    day_density_factor:     int = 5    # k*(k-1)*factor per student per day
    week_density_base:      int = 6    # (k-1)*(k+base)//2 per student per week

    # Hard constraint toggles (for display only — not editable weight)
    enforce_student_clash:  bool = True
    enforce_constraint_code: bool = True

    # User-facing priority weights (0-100, sum = 100)
    student_stress_weight:  int = 50   # relative priority: student stress load
    teacher_load_weight:    int = 50   # relative priority: teacher marking load
