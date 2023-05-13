import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

print(filepaths)


def split_n_caps(txt):
    list_words = txt.split("_")
    list_words = [x.title() for x in list_words]
    final_text = " ".join(list_words)
    return final_text


for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add the header of the table
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    # print(columns)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=30, h=8, txt=str(row[row.index[0]]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Added the sum
    pdf.set_font(family="Times", size=8)
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(df["total_price"].sum()), border=1, ln=1)

    # Sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=100, h=15, align="Left", txt=f"The total due amount is {df['total_price'].sum()} GBP.", ln=1)

    # Company name and logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=40, h=10, txt="A Monkey Company", align="Left")
    pdf.image("monkey.png", x=45, h=6, w=4)

    pdf.output(f"PDFs/{filename}.pdf")
