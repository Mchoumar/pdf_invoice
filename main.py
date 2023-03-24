import fpdf
from fpdf import FPDF
import pandas as pd
from glob import glob
from pathlib import Path

# access csv files
filepaths = glob("Invoices/*.xlsx")

for file in filepaths:
    # read Excel files
    df = pd.read_excel(file, sheet_name="Sheet 1")

    # add Excel file information into a pdf
    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    #
    filename = Path(file).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date. {date}")
    pdf.output(f"PDFs/{filename}.pdf")
