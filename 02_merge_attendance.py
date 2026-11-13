"""
02_merge_attendance.py  --  SOLUTION for Demo 2: "Merge these weekly files into one master"

Combines every data/attendance_week*.xlsx into a single workbook.
Gotcha built into the data: week 3 uses the column names "Student" and "Status"
instead of "Name" and "Present" -- a good moment to ask "does the output look right?"

Run: open this file in VS Code and click Run Python File.
Output:
    ../output/attendance_master.xlsx  (sheets: Master, By Week, Summary)

Rules used in the Summary sheet:
  - "Y" and "Late" both count as attended
  - Attendance % is calculated over the weeks where the student was actually listed
  - Students under 75% are highlighted
"""
import re
from pathlib import Path

import pandas as pd
from openpyxl.styles import PatternFill

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data"
OUT_DIR = BASE / "output"
OUT_DIR.mkdir(exist_ok=True)
OUT = OUT_DIR / "attendance_master.xlsx"

COLUMN_FIXES = {"Student": "Name", "Status": "Present"}  # alternate headers -> standard ones
THRESHOLD = 75


def load_week(path):
    df = pd.read_excel(path).rename(columns=COLUMN_FIXES)
    df["Week"] = int(re.search(r"week(\d+)", path.stem).group(1))  # week number from file name
    df["Name"] = df["Name"].str.strip()
    return df[["Week", "Name", "Present", "Notes"]]


def main():
    files = sorted(DATA.glob("attendance_week*.xlsx"))
    if not files:
        raise SystemExit("No attendance files found. Open the whole vibe-coding-kit folder in VS Code (File > Open Folder) and check the data folder.")

    master = pd.concat([load_week(f) for f in files], ignore_index=True)
    master["Notes"] = master["Notes"].fillna("")
    master = master.sort_values(["Name", "Week"]).reset_index(drop=True)

    by_week = master.pivot(index="Name", columns="Week", values="Present").fillna("-")
    by_week.columns = [f"Week {c}" for c in by_week.columns]
    by_week = by_week.reset_index()

    master["Attended"] = master["Present"].isin(["Y", "Late"])
    summary = master.groupby("Name").agg(
        Weeks_Listed=("Week", "count"),
        Attended=("Attended", "sum"),
        Late=("Present", lambda s: (s == "Late").sum()),
        Absent=("Present", lambda s: (s == "N").sum()),
    ).reset_index()
    summary["Attendance %"] = (summary["Attended"] / summary["Weeks_Listed"] * 100).round(1)
    summary = summary.sort_values("Attendance %").reset_index(drop=True)

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        master.drop(columns="Attended").to_excel(writer, sheet_name="Master", index=False)
        by_week.to_excel(writer, sheet_name="By Week", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        # Highlight low attendance
        ws = writer.book["Summary"]
        red = PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")
        pct_col = list(summary.columns).index("Attendance %") + 1
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=pct_col).value < THRESHOLD:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = red

        for sheet in writer.book.worksheets:
            for col in sheet.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                sheet.column_dimensions[col[0].column_letter].width = min(width + 3, 30)
            sheet.freeze_panes = "A2"

    low = summary[summary["Attendance %"] < THRESHOLD]
    print(f"Merged {len(files)} files, {len(master)} rows, {summary.shape[0]} students.")
    print(f"{len(low)} student(s) below {THRESHOLD}% attendance:")
    print(low[["Name", "Attendance %"]].to_string(index=False))
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
