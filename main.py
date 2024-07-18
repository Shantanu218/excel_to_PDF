# This is a sample Python script.
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


# Code to make a new cell
def create_cell(arg):
	pdf.cell(w = 40, h = 8, txt = str(arg).replace("_", " ").title(), border = 1)


# Code to make a new end cell
def create_end_cell(arg):
	pdf.cell(w = 40, h = 8, txt = str(arg).replace("_", " ").title(), border = 1, ln = 1)


# Code to make an entire row when given a list of values as an input
def create_all_cells(args):
	for index, item in enumerate(args):
		if index + 1 == len(args):
			break
		create_cell(item)
	create_end_cell(args[-1])


rowLayout = ["product_id", "product_name", "amount_purchased",
			 "price_per_unit", "total_price"]

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
	
	pdf.set_font(family = "Times", size = 10, style = "B")
	pdf.set_text_color(80, 80, 80)
	
	# Makes a header row
	create_all_cells(rowLayout)
	
	dataframe = pd.read_excel(filepath, sheet_name = f"Sheet 1")
	# Makes a row for each product
	for index, row in dataframe.iterrows():
		pdf.set_font(family = "Times", size = 10)
		# Adds all the cells for each product
		create_all_cells(row.tolist())
	
	total_sum = dataframe["total_price"].sum()
	sumLayout = [" "] * (len(rowLayout) - 1)
	sumLayout.append(str(total_sum))
	create_all_cells(sumLayout)
	
	pdf.set_font(family = "Times", size = 16, style = "B")
	pdf.cell(w = 50, h = 8, txt = f"The total price for all items: {total_sum}", ln = 1)
	
	pdf.set_font(family = "Times", size = 16, style = "B")
	pdf.cell(w = 30, h = 8, txt = f"PythonHow")
	pdf.image("pythonhow.png", w = 10)
	
	pdf.output(f"PDFs/{filename}.pdf")
