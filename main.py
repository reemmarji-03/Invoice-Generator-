import pandas as pd
from fpdf import FPDF
import lookups
import glob


paths = glob.glob(lookups.invoices_path)

for path in paths:

    pdf = FPDF(orientation=lookups.pdf_orientation, unit=lookups.pdf_unit, format=lookups.pdf_format)
    pdf.add_page()

    filename = path[9:14]
    Date = path[15:24]

    pdf.set_font(family=lookups.family_font, size=16, style=lookups.font_style)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {filename}", ln=1)

    pdf.set_font(family=lookups.family_font, size=16, style=lookups.font_style)
    pdf.cell(w=50, h=8, txt=f"Date {Date}", ln=1)

    df = pd.read_excel(path, sheet_name=lookups.sheet_name)


    columns = df.columns
    columns = [c.replace("_", " ").title() for c in columns]
    pdf.set_font(family=lookups.family_font, style=lookups.font_style, size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    total_sum = 0
    
    for index, row in df.iterrows(): 
        pdf.set_font(family=lookups.family_font, size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

        total_sum = total_sum + row["total_price"]

    pdf.set_font(family=lookups.family_font, size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    pdf.set_font(family=lookups.family_font, size=10)
    pdf.cell(w=30, h=8, txt=f"The total due amount is {total_sum} Euros", ln=1)

    pdf.set_font(family=lookups.family_font, size=14)
    pdf.cell(w=24, h=8, txt=lookups.image_text)
    pdf.image(lookups.image_file_name, w=10)

    pdf.output(f"PDFs/{filename}.pdf")
    
