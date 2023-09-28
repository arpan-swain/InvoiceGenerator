import pandas as pd
import glob
from fpdf import FPDF
import pathlib

filepaths = glob.glob('invoices/*.xlsx')
print(filepaths)

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm",format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()

    # Add Invoice No and Date
    filename,date = pathlib.Path(filepath).stem.split("-")

    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=0, h=12, txt=f"Invoice No. {filename}",ln=1, align="l")
    pdf.cell(w=0, h=12, txt=f"Date {date}", ln=1, align="l")

    pdf.ln(5)

    # Add Headers
    df = pd.read_excel(filepath)

    columns_ = list(df.columns)
    columns = [x.replace("_"," ").title() for x in columns_]
    pdf.set_font(family="arial",style="B" , size=10)
    pdf.set_text_color(80, 80, 80)

    pdf.cell(w=30, h=10, txt=columns[0], ln=0, border=1)
    pdf.cell(w=70, h=10, txt=columns[1], ln=0, border=1)
    pdf.cell(w=35, h=10, txt=columns[2], ln=0, border=1)
    pdf.cell(w=30, h=10, txt=columns[3], ln=0, border=1)
    pdf.cell(w=30, h=10, txt=columns[4], ln=1, border=1)

    # Add rows of data
    pdf.set_font(family="arial", size=10)
    pdf.set_text_color(80,80,80)

    for index, row in df.iterrows():
        pdf.cell(w=30,h=10,txt=str(row["product_id"]),ln=0,border =1)
        pdf.cell(w=70, h=10, txt=str(row["product_name"]), ln=0,border =1)
        pdf.cell(w=35, h=10, txt=str(row["amount_purchased"]), ln=0,border =1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), ln=0,border =1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), ln=1,border =1)

    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=10, txt="", ln=0, border=1)
    pdf.cell(w=70, h=10, txt="", ln=0, border=1)
    pdf.cell(w=35, h=10, txt="", ln=0, border=1)
    pdf.cell(w=30, h=10, txt="", ln=0, border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), ln=1, border=1)

    pdf.ln(5)

    pdf.set_font(family="Times", style="B", size=15)
    pdf.set_text_color(0,0,0)
    pdf.cell(w=0, h=12, txt=f"Your total Bill is : {total_sum}", ln=1, align="l")
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=50, h=12, txt=f"The Pokemon Company", ln=0, align="l")
    pdf.image("image.png",w=10)



    pdf.output(f"PDFs/{filename}.pdf")

# print(pathlib.Path(filepaths[1]).name)