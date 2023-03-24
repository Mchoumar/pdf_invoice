import fpdf
import pandas as pd
from glob import glob
from pathlib import Path

# access csv files
filepaths = glob("Invoices/*.xlsx")

for file in filepaths:
    # add Excel file information into a pdf
    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    #
    filename = Path(file).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date. {date}", ln=1)

    # read Excel files
    df = pd.read_excel(file, sheet_name="Sheet 1")

    # add a header
    column_name = df.columns
    column_name = [item.replace("_", " ").title() for item in column_name]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=column_name[0], border=1)
    pdf.cell(w=50, h=10, txt=column_name[1], border=1)
    pdf.cell(w=40, h=10, txt=column_name[2], border=1)
    pdf.cell(w=30, h=10, txt=column_name[3], border=1)
    pdf.cell(w=30, h=10, txt=column_name[4], border=1, ln=1)

    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=10, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), border=1, ln=1)
    pdf.output(f"PDFs/{filename}.pdf")
