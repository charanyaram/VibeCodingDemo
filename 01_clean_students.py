"""
01_clean_students.py  --  SOLUTION for Demo 1: "Clean up this messy spreadsheet"

What it fixes in data/messy_students.xlsx:
  - blank rows
  - stray spaces in column headers ("Email ")
  - names in mixed case / "Last, First" format / extra spaces
  - five different spellings of each major
  - five different date formats
  - GPA stored as numbers, text, or "N/A"
  - duplicate students (same student id entered twice)

Run: open this file in VS Code and click Run Python File.
Output:
    ../output/clean_students.xlsx  (sheets: Clean Data, Cleaning Log)
"""
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "data" / "messy_students.xlsx"
OUT_DIR = BASE / "output"
OUT_DIR.mkdir(exist_ok=True)
OUT = OUT_DIR / "clean_students.xlsx"

# Every spelling we saw -> the one official name
MAJOR_MAP = {
    "computer science": "Computer Science", "comp sci": "Computer Science", "cs": "Computer Science",
    "business": "Business", "bus.": "Business", "business admin": "Business",
    "engineering": "Engineering", "eng": "Engineering",
    "biology": "Biology", "bio": "Biology",
}

# Tried in order. Note: slashes = day/month/year, dashes with 2-digit year = month-day-year.
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d %B %Y", "%m-%d-%y", "%b %d, %Y"]


def clean_name(value):
    if pd.isna(value):
        return value
    text = str(value).strip()
    if "," in text:  # "Smith, Anna" -> "Anna Smith"
        last, first = [part.strip() for part in text.split(",", 1)]
        text = f"{first} {last}"
    return " ".join(text.split()).title()  # collapse double spaces, Title Case


def clean_major(value):
    if pd.isna(value):
        return value
    key = str(value).strip().lower()
    return MAJOR_MAP.get(key, str(value).strip().title())  # unknown -> tidy up and keep


def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return pd.NaT  # couldn't read it -> flag as missing rather than guess


def main():
    log = []
    df = pd.read_excel(SRC)
    log.append(("Rows read from file", len(df)))

    # 1. Tidy headers, then drop completely blank rows
    df.columns = df.columns.str.strip()
    before = len(df)
    df = df.dropna(how="all")
    log.append(("Blank rows removed", before - len(df)))

    # 2. Rename to consistent headers
    df = df.rename(columns={
        "Student Name": "Name", "student id": "Student ID",
        "Enrol Date": "Enrol Date",
    })

    # 3. Clean each column
    df["Name"] = df["Name"].apply(clean_name)
    df["Major"] = df["Major"].apply(clean_major)
    df["Enrol Date"] = df["Enrol Date"].apply(parse_date)
    df["GPA"] = pd.to_numeric(df["GPA"], errors="coerce")  # "N/A" / text -> number or blank
    df["Email"] = df["Email"].str.strip().str.lower()

    unreadable_dates = int(df["Enrol Date"].isna().sum())
    missing_gpa = int(df["GPA"].isna().sum())
    missing_email = int(df["Email"].isna().sum())

    # 4. Remove duplicates (same Student ID = same person)
    before = len(df)
    df = df.drop_duplicates(subset="Student ID", keep="first")
    log.append(("Duplicate students removed", before - len(df)))

    # 5. Sort and save
    df = df.sort_values("Name").reset_index(drop=True)
    log += [
        ("Final number of students", len(df)),
        ("Students with missing GPA", missing_gpa),
        ("Students with missing email", missing_email),
        ("Dates that could not be read", unreadable_dates),
    ]

    with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, sheet_name="Clean Data", index=False)
        pd.DataFrame(log, columns=["Step", "Count"]).to_excel(writer, sheet_name="Cleaning Log", index=False)

        for ws in writer.book.worksheets:  # readable column widths
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 40)
            ws.freeze_panes = "A2"

    print(f"Saved {OUT}")
    for step, count in log:
        print(f"  {step}: {count}")


if __name__ == "__main__":
    main()
