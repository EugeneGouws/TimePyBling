from datetime import date

DEFAULT_EXCLUSIONS   = {"ST", "LIB", "PE", "RDI"}
SESSIONS             = ["AM", "PM"]
TEACHER_SUBJECT_COLS = ["sua", "sub", "suc"]

# Rotation timetable grid: 8 days × 7 periods.
# Each entry is the subblock that falls in that (day, period) slot.
TIMETABLE_GRID = [
    ["A1","B2","C3","D4","E5","F6","G7"],  # Day 1
    ["G1","H2","B3","C4","D5","E6","F7"],  # Day 2
    ["F1","A2","H3","B4","C5","D6","E7"],  # Day 3
    ["E1","F2","A3","H4","B5","C6","D7"],  # Day 4
    ["D1","E2","F3","A4","H5","B6","C7"],  # Day 5
    ["C1","D2","E3","F4","A5","H6","B7"],  # Day 6
    ["B1","C2","D3","E4","F5","A6","H7"],  # Day 7
    ["H1","B2","C3","D4","E5","F6","A7"],  # Day 8
]

CLR_GRID_CELL   = "#FFFBF5"
CLR_GRID_HEADER = "#DBEAFE"
CLR_GRID_ACTIVE = "#FED7AA"

DEFAULT_EXAM_START = date(2026, 6, 1)
DEFAULT_EXAM_END   = date(2026, 6, 23)

CLR_HEADER    = "#1E3A5F"
CLR_ORANGE    = "#EA6C0A"
CLR_BLUE      = "#1D5CB4"
CLR_PINK      = "#C2185B"
CLR_GREEN     = "#16A34A"
CLR_RED       = "#DC2626"
CLR_LIGHT     = "#EFF6FF"
CLR_MID       = "#93C5FD"
CLR_WHITE     = "white"
CLR_BG        = "#F0F4FF"
CLR_MORNING   = "#FFF3E0"
CLR_AFTERNOON = "#FCE4EC"
