"""
instructor_generate_sample_data.py
Creates the deliberately messy Excel files used in the Vibe Coding session.

Instructor only: open this file in VS Code and click Run Python File to recreate the data files.

Output goes to ../data/
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)
DATA = Path(__file__).resolve().parent.parent / "data"
DATA.mkdir(exist_ok=True)

FIRST = ["Sokha", "Dara", "Maria", "Liam", "Aisha", "Chen", "Priya", "Noah",
         "Lina", "Omar", "Mei", "Jack", "Sara", "Ravi", "Emma", "Kai",
         "Nadia", "Tom", "Yuki", "Ben", "Sophie", "Arun", "Zoe", "Leo"]
LAST = ["Smith", "Patel", "Kim", "Nguyen", "Brown", "Lee", "Khan", "Garcia",
        "Tan", "Wong", "Singh", "Ali", "Jones", "Chea", "Sok", "Wilson"]


# --------------------------------------------------------------------------
# 1. messy_students.xlsx  ->  demo: "clean this up"
# --------------------------------------------------------------------------
def make_messy_students(n=40):
    majors_clean = ["Computer Science", "Business", "Engineering", "Biology"]
    major_variants = {
        "Computer Science": ["Computer Science", "comp sci", "CS", "computer science ", "COMPUTER SCIENCE"],
        "Business": ["Business", "business", "Bus.", "BUSINESS ", "Business Admin"],
        "Engineering": ["Engineering", "engineering", "Eng", "ENGINEERING"],
        "Biology": ["Biology", "biology ", "Bio", "BIOLOGY"],
    }

    def messy_date(d):
        style = random.choice(["iso", "dmy", "long", "us", "text"])
        if style == "iso":
            return d.strftime("%Y-%m-%d")
        if style == "dmy":
            return d.strftime("%d/%m/%Y")
        if style == "long":
            return d.strftime("%d %B %Y")
        if style == "us":
            return d.strftime("%m-%d-%y")
        return d.strftime("%b %d, %Y")

    rows = []
    for i in range(n):
        first, last = random.choice(FIRST), random.choice(LAST)
        name = f"{first} {last}"
        style = random.choice(["plain", "upper", "lower", "spaces", "swap"])
        if style == "upper":
            name = name.upper()
        elif style == "lower":
            name = name.lower()
        elif style == "spaces":
            name = f"  {name}   "
        elif style == "swap":
            name = f"{last}, {first}"

        major = random.choice(majors_clean)
        gpa = round(random.uniform(1.8, 4.0), 2)
        gpa_val = random.choice([gpa, gpa, str(gpa), None, "N/A"])

        email = f"{first}.{last}@uni.example.edu".lower()
        if random.random() < 0.15:
            email = None

        enrol = date(2024, 2, 1) + timedelta(days=random.randint(0, 400))
        rows.append({
            "Student Name": name,
            "student id": f"S{1000 + i}",
            "Email ": email,
            "Major": random.choice(major_variants[major]),
            "Enrol Date": messy_date(enrol),
            "GPA": gpa_val,
        })

    df = pd.DataFrame(rows)
    # Duplicates (same student entered twice, slightly different formatting)
    dupes = df.sample(5, random_state=1).copy()
    dupes["Student Name"] = dupes["Student Name"].astype(str).str.strip().str.upper()
    df = pd.concat([df, dupes], ignore_index=True).sample(frac=1, random_state=7).reset_index(drop=True)

    # Blank rows in the middle and at the end
    blank = pd.DataFrame([{c: None for c in df.columns}])
    df = pd.concat([df.iloc[:12], blank, df.iloc[12:30], blank, df.iloc[30:], blank, blank],
                   ignore_index=True)
    df.to_excel(DATA / "messy_students.xlsx", index=False)


# --------------------------------------------------------------------------
# 2. attendance_week1..4.xlsx  ->  demo: "merge these into one master file"
# --------------------------------------------------------------------------
def make_attendance(weeks=4, n=18):
    roster = [f"{random.choice(FIRST)} {random.choice(LAST)}" for _ in range(n * 2)]
    roster = list(dict.fromkeys(roster))[:n]  # unique, keep order
    for w in range(1, weeks + 1):
        rows = []
        for name in roster:
            if random.random() < 0.08:
                continue  # student not listed this week
            present = random.choices(["Y", "N", "Late"], weights=[75, 15, 10])[0]
            note = random.choice(["", "", "", "sick", "bus delay", "excused"]) if present != "Y" else ""
            rows.append({"Name": name, "Present": present, "Notes": note})
        df = pd.DataFrame(rows)
        if w == 3:
            # Gotcha: one teacher used different column names this week
            df = df.rename(columns={"Name": "Student", "Present": "Status"})
        df.to_excel(DATA / f"attendance_week{w}.xlsx", index=False)


# --------------------------------------------------------------------------
# 3. student_expenses.xlsx  ->  demo: "turn this into a report with charts"
# --------------------------------------------------------------------------
def make_expenses():
    cats = {
        "Food": (["Lunch", "Coffee", "Groceries", "Dinner out", "Snacks"], (2, 25)),
        "Transport": (["Bus pass", "Ride share", "Fuel", "Train ticket"], (3, 40)),
        "Study": (["Textbook", "Printing", "Stationery", "Online course"], (5, 60)),
        "Entertainment": (["Movie", "Streaming", "Game", "Concert"], (8, 50)),
        "Rent & Bills": (["Rent share", "Phone plan", "Internet"], (20, 120)),
    }
    weights = [40, 15, 10, 12, 8]
    rows = []
    start = date(2026, 9, 1)
    for _ in range(120):
        cat = random.choices(list(cats), weights=weights)[0]
        items, (lo, hi) = cats[cat]
        rows.append({
            "Date": start + timedelta(days=random.randint(0, 59)),
            "Category": cat,
            "Description": random.choice(items),
            "Amount": round(random.uniform(lo, hi), 2),
        })
    df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
    df.to_excel(DATA / "student_expenses.xlsx", index=False)


if __name__ == "__main__":
    make_messy_students()
    make_attendance()
    make_expenses()
    print(f"Sample files written to {DATA}")
    for f in sorted(DATA.glob("*.xlsx")):
        print("  -", f.name)
