import fpdf
from fpdf import FPDF
import pandas as pd
from glob import glob
from pathlib import Path

# access csv files
filepaths = glob("Invoices/*.xlsx")
print(filepaths)
for file in filepaths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(file).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    pdf.output(f"PDFs/{invoice_nr}.pdf")