import pandas as pd
import glob
from fpdf import FPDF
import pathlib

filepaths = glob.glob('invoices/*.xlsx')
print(filepaths)

for filepath in filepaths:
    df =pd.read_excel(filepath)
    pdf = FPDF(orientation="P", unit="mm",format="A4")
    pdf.add_page()
    filename = pathlib.Path(filepath).stem.split("-")[0]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=0, h=12, txt=f"Invoice No. {filename}",ln=1, align="l")
    pdf.output(f"PDFs/{filename}.pdf")