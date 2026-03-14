import pandas as pd
import re
from pathlib import Path

def show_student_timetable_from_st1(st1_path, student_id: int) -> pd.DataFrame:
    st1_path = Path(st1_path)
    df = pd.read_excel(st1_path)

    if "Studentid" not in df.columns:
        raise ValueError("ST1.xlsx must contain a 'Studentid' column.")

    df = df.dropna(subset=["Studentid"]).copy()
    df["Studentid"] = df["Studentid"].astype(float).astype(int)

    timetable_cols = [
        str(col).strip()
        for col in df.columns
        if re.fullmatch(r"[A-H]\d+", str(col).strip())
    ]
    timetable_cols.sort(key=lambda x: (x[0], int(x[1:])))

    if not timetable_cols:
        raise ValueError("No timetable columns found. Expected columns like A1, B3, H7.")

    row_match = df.loc[df["Studentid"] == int(student_id)]
    if row_match.empty:
        raise ValueError(f"Student ID {student_id} not found in {st1_path.name}.")

    row = row_match.iloc[0]

    max_row = 10
    block_names = list("ABCDEFGH")
    out = pd.DataFrame("", index=range(1, max_row + 1), columns=block_names)

    for col in timetable_cols:
        block = col[0]
        period = int(col[1:])
        val = row[col]

        if pd.isna(val):
            out.at[period, block] = "-"
        else:
            raw = str(val).strip()
            out.at[period, block] = "-" if raw.upper() == "FREE" or raw == "" else raw

    print(f"\nTimetable for student {student_id}\n")
    print(out.to_string())

    return out

if __name__ == "__main__":
    import sys
    path = sys.argv[1]
    student_id = int(sys.argv[2])
    show_student_timetable_from_st1(path, student_id)