# This is a sample Python script.
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
	pdf = FPDF(orientation = "P", unit = "mm", format = "A4")
	pdf.add_page()
	
	filename = Path(filepath).stem
	invoice_number, date = filename.split("-")
	
	pdf.set_font(family = "Times", size = 16, style = "B")
	pdf.cell(w = 50, h = 8, txt = f"Invoice number: {invoice_number}", ln = 1)
	
	pdf.set_font(family = "Times", size = 16, style = "B")
	pdf.cell(w = 50, h = 8, txt = f"Invoice date: {date}", ln = 1)
	
	rowLayout = ["product_id", "product_name", "amount_purchased",
				 "price_per_unit", "total_price"]
	
	
	# Code to make a new cell
	def create_cell(arg):
		pdf.cell(w = 40, h = 8, txt = str(arg).replace("_", " ").title(), border = 1)
	
	
	# Code to make a new end cell
	def create_end_cell(arg):
		pdf.cell(w = 40, h = 8, txt = str(arg).replace("_", " ").title(), border = 1, ln = 1)
	
	
	pdf.set_font(family = "Times", size = 10, style = "B")
	pdf.set_text_color(80, 80, 80)
	
	# Add a header
	for item in rowLayout:
		if item == "total_price":
			create_end_cell(item)
		else:
			create_cell(item)
	
	dataframe = pd.read_excel(filepath, sheet_name = f"Sheet 1")
	# Add rows
	for index, row in dataframe.iterrows():
		pdf.set_font(family = "Times", size = 10)
		# Add cells
		for item in rowLayout:
			if item == "total_price":
				create_end_cell(row[item])
			else:
				create_cell(row[item])
	
	pdf.output(f"PDFs/{filename}.pdf")
