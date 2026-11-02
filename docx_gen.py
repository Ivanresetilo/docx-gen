#!/usr/bin/env python3
"""CLI-генератор .docx документов по ГОСТ-стилю: Times New Roman 14, полуторный интервал."""
import argparse
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


def make_docx(title, author, sections, out, font="Times New Roman", size=14):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(size)
    pf = style.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Cm(1.25)

    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2)
        s.left_margin = Cm(3)
        s.right_margin = Cm(1.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(size + 2)

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"Выполнил: {author}").italic = True

    for heading, body in sections:
        doc.add_heading(heading, level=1)
        for para in body.split("\n\n"):
            if para.strip():
                doc.add_paragraph(para.strip())

    doc.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description="Generate .docx reports in GOST style")
    ap.add_argument("--title", required=True)
    ap.add_argument("--author", default="")
    ap.add_argument("--section", action="append", nargs=2,
                    metavar=("HEADING", "TEXT"), required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    print("Saved:", make_docx(args.title, args.author, args.section, args.out))


if __name__ == "__main__":
    main()
