"""
03_expense_report.py  --  SOLUTION for Demo 3: "Turn raw data into a report with charts"

Reads data/student_expenses.xlsx and builds a formatted report workbook with
a category summary, a monthly breakdown, top expenses, and two native Excel charts.

Run: open this file in VS Code and click Run Python File.
Output:
    ../output/expense_report.xlsx  (sheets: Summary, Monthly, Top 10, Data)
"""
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "data" / "student_expenses.xlsx"
OUT_DIR = BASE / "output"
OUT_DIR.mkdir(exist_ok=True)
OUT = OUT_DIR / "expense_report.xlsx"

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def autosize(ws, cap=30):
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(width + 3, cap)


def main():
    df = pd.read_excel(SRC)
    df["Date"] = pd.to_datetime(df["Date"])
    df["Month"] = df["Date"].dt.strftime("%Y-%m")

    by_cat = (df.groupby("Category")["Amount"]
                .agg(Total="sum", Count="count", Average="mean")
                .round(2).sort_values("Total", ascending=False).reset_index())
    by_cat["Share %"] = (by_cat["Total"] / by_cat["Total"].sum() * 100).round(1)

    monthly = df.pivot_table(index="Category", columns="Month", values="Amount",
                             aggfunc="sum", fill_value=0).round(2)
    monthly["Total"] = monthly.sum(axis=1)
    monthly = monthly.sort_values("Total", ascending=False).reset_index()

    top10 = df.nlargest(10, "Amount")[["Date", "Category", "Description", "Amount"]]

    with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        by_cat.to_excel(writer, sheet_name="Summary", index=False)
        monthly.to_excel(writer, sheet_name="Monthly", index=False)
        top10.to_excel(writer, sheet_name="Top 10", index=False)
        df.drop(columns="Month").to_excel(writer, sheet_name="Data", index=False)
        wb = writer.book

        # --- Summary sheet: total row (live formula) + charts
        ws = wb["Summary"]
        n = len(by_cat)
        total_row = n + 2
        ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=total_row, column=2, value=f"=SUM(B2:B{n + 1})").font = Font(bold=True)
        ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{n + 1})").font = Font(bold=True)
        for row in range(2, total_row + 1):
            ws.cell(row=row, column=2).number_format = "$#,##0.00"
            ws.cell(row=row, column=4).number_format = "$#,##0.00"

        bar = BarChart()
        bar.title, bar.y_axis.title, bar.legend = "Spending by category", "USD", None
        bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=n + 1), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n + 1))
        bar.width, bar.height = 16, 8
        ws.add_chart(bar, "G2")

        pie = PieChart()
        pie.title = "Share of total spending"
        pie.add_data(Reference(ws, min_col=2, min_row=1, max_row=n + 1), titles_from_data=True)
        pie.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n + 1))
        pie.width, pie.height = 12, 8
        ws.add_chart(pie, "G20")

        # --- Formatting on every sheet
        for sheet in wb.worksheets:
            style_header(sheet, sheet.max_column)
            autosize(sheet)
            sheet.freeze_panes = "A2"
        for row in wb["Top 10"].iter_rows(min_row=2, min_col=4, max_col=4):
            row[0].number_format = "$#,##0.00"

    print(f"Saved {OUT}")
    print(f"Total spend: ${df['Amount'].sum():,.2f} across {len(df)} transactions")
    print(by_cat[["Category", "Total", "Share %"]].to_string(index=False))


if __name__ == "__main__":
    main()
