# This is a sample Python script.
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name=f"Sheet 1")
    pdf = FPDF(orientation = "P", unit = "mm", format = "A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    pdf.set_font(family = "Times", size = 16, style = "B")
    pdf.cell(w=50, h=8, txt = f"Invoice number: {invoice_number}")
    pdf.output(f"PDFs/{filename}.pdf")
    print(df)


# Press Ctrl+F5 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

