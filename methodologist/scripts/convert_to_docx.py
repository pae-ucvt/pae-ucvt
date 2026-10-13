"""
Конвертер .md -> .docx для программ ДПП ПК, планов занятий и тестов.

Использование:
    python convert_to_docx.py файл.md               # сохранит файл.docx рядом
    python convert_to_docx.py файл.md --out путь.docx # указать имя на выходе
"""

import re
import sys
import argparse
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_cell(cell, text, bold=False, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    run.bold = bold


def md_to_docx(md_text):
    doc = Document()

    # --- default style ---
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(4)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    for level in range(1, 4):
        hs = doc.styles[f"Heading {level}"]
        hs.font.name = "Times New Roman"
        hs.font.color.rgb = RGBColor(0, 0, 0)
        hs.font.bold = True
        hs.font.size = Pt({1: 16, 2: 14, 3: 12}[level])

    lines = md_text.split("\n")
    i = 0
    in_table = False
    table_rows = []

    def flush_table():
        nonlocal table_rows, in_table
        if not table_rows:
            return
        t = doc.add_table(rows=len(table_rows), cols=len(table_rows[0]))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.style = "Table Grid"
        for ri, row_data in enumerate(table_rows):
            for cj, cell_text in enumerate(row_data):
                set_cell(t.cell(ri, cj), cell_text, bold=(ri == 0))
        table_rows = []
        in_table = False

    while i < len(lines):
        line = lines[i].rstrip()

        # --- horizontal rule ---
        if re.match(r"^-{3,}$", line):
            flush_table()
            doc.add_paragraph("")
            i += 1
            continue

        # --- detect table ---
        if re.match(r"^\|.+\|$", line):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            if not re.match(r"^[\s\-:]+\|", line):  # skip separator row
                table_rows.append(cells)
                in_table = True
            i += 1
            continue
        else:
            flush_table()

        # --- headings ---
        hm = re.match(r"^(#{1,3})\s+(.+)$", line)
        if hm:
            level = len(hm.group(1))
            text = hm.group(2).strip("**").strip()
            doc.add_heading(text, level=level)
            i += 1
            continue

        # --- bullet list ---
        bm = re.match(r"^[\s]*[-*]\s+(.+)$", line)
        if bm:
            text = bm.group(1)
            # bold prefix
            bold_prefix = re.match(r"^\*\*(.+?)\*\*[:.]?\s*(.*)$", text)
            p = doc.add_paragraph(style="List Bullet")
            if bold_prefix:
                run = p.add_run(bold_prefix.group(1))
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                if bold_prefix.group(2):
                    run2 = p.add_run(" " + bold_prefix.group(2))
                    run2.font.name = "Times New Roman"
                    run2.font.size = Pt(12)
            else:
                run = p.add_run(text)
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
            i += 1
            continue

        # --- numbered list ---
        nm = re.match(r"^\d+[.)]\s+(.+)$", line)
        if nm:
            text = nm.group(1)
            bold_prefix = re.match(r"^\*\*(.+?)\*\*[:.]?\s*(.*)$", text)
            p = doc.add_paragraph(style="List Number")
            if bold_prefix:
                run = p.add_run(bold_prefix.group(1))
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                if bold_prefix.group(2):
                    run2 = p.add_run(" " + bold_prefix.group(2))
                    run2.font.name = "Times New Roman"
                    run2.font.size = Pt(12)
            else:
                run = p.add_run(text)
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
            i += 1
            continue

        # --- empty line ---
        if not line.strip():
            flush_table()
            doc.add_paragraph("")
            i += 1
            continue

        # --- plain paragraph with bold markers ---
        flush_table()
        p = doc.add_paragraph()
        parts = re.split(r"(\*\*.+?\*\*)", line)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                run = p.add_run(part[2:-2])
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
            elif part.strip():
                run = p.add_run(part)
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

        i += 1

    flush_table()
    return doc


def main():
    parser = argparse.ArgumentParser(description="Convert .md to .docx")
    parser.add_argument("input", help="Input .md file path")
    parser.add_argument("--out", "-o", help="Output .docx file path")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: file not found: {input_path}")
        sys.exit(1)

    md_text = input_path.read_text(encoding="utf-8")
    doc = md_to_docx(md_text)

    output_path = Path(args.out) if args.out else input_path.with_suffix(".docx")
    doc.save(str(output_path))
    print(f"OK -> {output_path.resolve()}")


if __name__ == "__main__":
    main()
