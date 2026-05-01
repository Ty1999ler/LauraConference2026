"""Converts 0_HOW_TO.md to 0_HOW_TO.pdf using fpdf2."""
import os
import re
from fpdf import FPDF

MD_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "0_HOW_TO.md")
PDF_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "0_HOW_TO.pdf")

COL_WIDTHS = [90, 95]   # two-column table widths


class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 8, "Alumo Conference 2026 - How To Guide", align="R")
        self.ln(6)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 8, f"Page {self.page_no()}", align="C")


def _strip_inline(text: str) -> str:
    """Remove markdown formatting and replace non-Latin-1 characters."""
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'`(.*?)`',       r'\1', text)
    text = (text
            .replace('→', '->')   # →
            .replace('—', '-')    # —
            .replace('–', '-')    # –
            .replace('“', '"')    # "
            .replace('”', '"')    # "
            .replace('‘', "'")    # '
            .replace('’', "'"))   # '
    # Drop anything still outside Latin-1
    return text.encode('latin-1', errors='replace').decode('latin-1')


def build_pdf():
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_margins(20, 20, 20)

    with open(MD_FILE, encoding="utf-8") as f:
        lines = f.readlines()

    in_code   = False
    in_table  = False
    table_rows = []

    def flush_table():
        nonlocal table_rows, in_table
        if not table_rows:
            return
        header_row = table_rows[0]
        data_rows  = table_rows[2:]  # skip separator row
        # header
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(180, 180, 180)
        for i, cell in enumerate(header_row):
            pdf.cell(COL_WIDTHS[i], 7, _strip_inline(cell.strip()), border=1,
                     fill=True, align="C")
        pdf.ln()
        # data
        pdf.set_font("Helvetica", "", 9)
        fill = False
        for row in data_rows:
            pdf.set_fill_color(230, 230, 230) if fill else pdf.set_fill_color(255, 255, 255)
            for i, cell in enumerate(row):
                pdf.cell(COL_WIDTHS[i], 6, _strip_inline(cell.strip()), border=1, fill=fill)
            pdf.ln()
            fill = not fill
        pdf.ln(3)
        table_rows = []
        in_table   = False

    for line in lines:
        raw = line.rstrip("\n")

        # Code fence
        if raw.strip().startswith("```"):
            in_code = not in_code
            continue

        if in_code:
            pdf.set_font("Courier", "", 9)
            pdf.set_fill_color(240, 240, 240)
            pdf.cell(0, 6, raw.replace("    ", "  "), fill=True)
            pdf.ln()
            continue

        # Table row
        if raw.strip().startswith("|"):
            cells = [c for c in raw.strip().split("|") if c != ""]
            table_rows.append(cells)
            in_table = True
            continue
        elif in_table:
            flush_table()

        # Headings
        if raw.startswith("# "):
            flush_table()
            pdf.set_font("Helvetica", "B", 16)
            pdf.set_text_color(30, 30, 30)
            pdf.ln(4)
            pdf.cell(0, 10, _strip_inline(raw[2:]), ln=True)
            pdf.set_draw_color(100, 100, 100)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 170, pdf.get_y())
            pdf.ln(3)
            continue

        if raw.startswith("## "):
            flush_table()
            pdf.set_font("Helvetica", "B", 13)
            pdf.set_text_color(50, 50, 50)
            pdf.ln(3)
            pdf.cell(0, 8, _strip_inline(raw[3:]), ln=True)
            continue

        if raw.startswith("### "):
            flush_table()
            pdf.set_font("Helvetica", "B", 11)
            pdf.set_text_color(70, 70, 70)
            pdf.ln(2)
            pdf.cell(0, 7, _strip_inline(raw[4:]), ln=True)
            continue

        # Horizontal rule
        if raw.strip() == "---":
            pdf.ln(2)
            pdf.set_draw_color(180, 180, 180)
            pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x() + 170, pdf.get_y())
            pdf.ln(4)
            continue

        # Numbered list
        m = re.match(r'^(\d+)\.\s+(.*)', raw)
        if m:
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(30, 30, 30)
            pdf.cell(8, 6, f"{m.group(1)}.")
            pdf.multi_cell(0, 6, _strip_inline(m.group(2)))
            continue

        # Bullet list
        if raw.strip().startswith("- "):
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(30, 30, 30)
            pdf.cell(8, 6, "-")
            pdf.multi_cell(0, 6, _strip_inline(raw.strip()[2:]))
            continue

        # Blank line
        if raw.strip() == "":
            pdf.ln(2)
            continue

        # Normal paragraph
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(30, 30, 30)
        pdf.multi_cell(0, 6, _strip_inline(raw.strip()))

    flush_table()
    pdf.output(PDF_FILE)
    print(f"PDF saved: {PDF_FILE}")


if __name__ == "__main__":
    build_pdf()
