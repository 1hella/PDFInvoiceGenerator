import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    columns = df.columns
    columns = [column.replace('_', ' ').title() for column in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=True)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=True)
    pdf.cell(w=35, h=8, txt=str(columns[2]), border=True)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=True)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=True, ln=1)

    for index, row in df.iterrows():
        print(row)
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=True)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=True)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=True)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=True)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=True, ln=1)

    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=70, h=8, border=True)
    pdf.cell(w=35, h=8, border=True)
    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is ${total_sum}", ln=1)

    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=26, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
