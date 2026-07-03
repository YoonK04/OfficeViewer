import os
from fpdf import FPDF

base = os.path.dirname(os.path.abspath(__file__))
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)

pdf.add_page()
pdf.set_font("Helvetica", "B", 20)
pdf.set_text_color(14, 99, 156)
pdf.cell(0, 15, "Office Viewer - PDF Test", new_x="LMARGIN", new_y="NEXT")
pdf.set_text_color(0, 0, 0)
pdf.set_font("Helvetica", "", 12)
pdf.ln(4)
pdf.multi_cell(0, 8, "This PDF is rendered by Chromium's built-in viewer inside the app. "
                     "You can scroll, zoom, search, and jump pages with the native toolbar.")
pdf.ln(4)
for i in range(1, 21):
    pdf.cell(0, 8, f"Line {i}: sample content for the PDF preview.", new_x="LMARGIN", new_y="NEXT")

pdf.add_page()
pdf.set_font("Helvetica", "B", 16)
pdf.cell(0, 12, "Page 2", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "", 12)
pdf.multi_cell(0, 8, "Second page confirms multi-page scrolling works.")

pdf.output(os.path.join(base, "샘플문서.pdf"))
print("pdf created")
