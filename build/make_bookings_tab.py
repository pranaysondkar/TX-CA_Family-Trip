"""Writes ../../Bookings_tab_for_Google_Sheet.xlsx: import it into the family Google Sheet as a new tab.

Google Sheets: File → Import → Upload → "Insert new sheet(s)". The tab must be named "Bookings".
Only the Booked?, Actual paid and Paid by columns are meant to be edited. Keep the ID column as is.
"""
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import trip_data as T

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "Bookings_tab_for_Google_Sheet.xlsx")
PREBOOKED = {"Already paid": True}

wb = Workbook()
ws = wb.active
ws.title = T.BOOKINGS_TAB
cols = ["ID", "Category", "Item", "Dates", "In plan?", "Planned ($)", "Booked?", "Actual paid ($)", "Paid by", "Notes"]
widths = [30, 20, 52, 18, 9, 12, 10, 14, 14, 40]
for i, (c, w) in enumerate(zip(cols, widths), start=1):
    cell = ws.cell(1, i, c)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F3A5F")
    ws.column_dimensions[cell.column_letter].width = w
ws.freeze_panes = "A2"
edit = PatternFill("solid", fgColor="FFF59D")
yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
ws.add_data_validation(yn)
for r, (cat, item, dates, qty, unit, tax, status, lean, lk, note) in enumerate(T.BUDGET, start=2):
    planned = round(qty * unit * (1 + tax), 2)
    booked = "Yes" if cat == "Already paid" or status == "Booked" else "No"
    actual = planned if cat == "Already paid" else None
    row = [T.budget_id(cat, item), cat, item, dates, "N" if status == "Optional" else "Y", planned, booked, actual, None, None]
    for c, v in enumerate(row, start=1):
        cell = ws.cell(r, c, v)
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="top", wrap_text=c in (3, 10))
    for c in (5, 7, 8, 9, 10):
        ws.cell(r, c).fill = edit
    ws.cell(r, 6).number_format = ws.cell(r, 8).number_format = '$#,##0.00'
    yn.add(f"G{r}")
ws["A1"].comment = Comment("Don't edit IDs. The website matches rows by this column.", "trip")
ws["H1"].comment = Comment("Enter the real amount paid incl. taxes & fees. Leave blank until paid.", "trip")
wb.save(OUT)
print("wrote", os.path.abspath(OUT), len(T.BUDGET), "rows")
