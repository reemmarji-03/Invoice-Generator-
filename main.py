import pandas as pd
from fpdf import FPDF
import glob


paths = glob.glob("Invoices/*.xlsx")

for path in paths:
    df = pd.read_excel(path, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = path[9:14]
    Date = path[15:24]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {filename}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {Date}")

    pdf.output(f"PDFs/{filename}.pdf")
    
