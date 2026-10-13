#!/usr/bin/env python3
"""Build editable Thai DOCX + printable PDF from the public, privacy-safe Markdown templates.

Requirements: pandoc, libreoffice, python-docx, a Thai font (Noto Sans Thai preferred).
Usage: python tools/build_documents.py
"""
from __future__ import annotations
import os
import subprocess
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "published"
OUT.mkdir(exist_ok=True)
FONT = os.environ.get("NITI_THAI_FONT", "Noto Sans Thai")
NAVY = RGBColor(24, 47, 72)
ORANGE = RGBColor(185, 88, 32)
GRAY = RGBColor(93, 105, 118)

def run(*cmd):
    print("+", " ".join(map(str, cmd)), flush=True)
    subprocess.run([str(x) for x in cmd], check=True)

def set_font(run_obj, name=FONT, size=None, bold=None, color=None):
    run_obj.font.name = name
    if size is not None:
        run_obj.font.size = Pt(size)
    if bold is not None:
        run_obj.bold = bold
    if color is not None:
        run_obj.font.color.rgb = color
    rp = run_obj._element.get_or_add_rPr()
    fonts = rp.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rp.insert(0, fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn("w:" + key), name)

def style_font(style, size, color=None, bold=None, space_before=0, space_after=0):
    style.font.name = FONT
    style.font.size = Pt(size)
    if color: style.font.color.rgb = color
    if bold is not None: style.font.bold = bold
    style.paragraph_format.space_before = Pt(space_before)
    style.paragraph_format.space_after = Pt(space_after)
    style.paragraph_format.line_spacing = 1.17

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)

def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    trPr.append(el)

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)

def append_page_field(p):
    p.add_run("หน้า ")
    f = OxmlElement("w:fldSimple")
    f.set(qn("w:instr"), "PAGE")
    p._p.append(f)

def enhance_docx(path, is_quick=False):
    doc = Document(path)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = Cm(1.18 if is_quick else 2.0)
    sec.bottom_margin = Cm(1.15 if is_quick else 1.55)
    sec.left_margin = Cm(1.65)
    sec.right_margin = Cm(1.65)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.7)

    normal = doc.styles["Normal"]
    style_font(normal, 8.2 if is_quick else 10.0, RGBColor(34, 45, 54), space_after=(2.5 if is_quick else 4))
    for name, size, color, bold, before, after in (
        ("Title", (17 if is_quick else 21), NAVY, True, 6, 8),
        ("Heading 1", (12.5 if is_quick else 15), NAVY, True, 8, 3),
        ("Heading 2", (10 if is_quick else 11.5), ORANGE, True, 6, 2),
        ("Heading 3", 10.3, NAVY, True, 7, 2),
        ("Block Text", 9.7, GRAY, False, 2, 5),
    ):
        if name in doc.styles:
            style_font(doc.styles[name],size,color,bold,before,after)

    first_heading = True
    for p in doc.paragraphs:
        if p.style.name == "Heading 1" and first_heading:
            p.style = doc.styles["Title"]
            first_heading = False
        if p.style.name.startswith("Heading"):
            p.paragraph_format.keep_with_next = True
        # Ensure Thai glyph fallback in generated runs; pandoc may use theme fonts.
        for r in p.runs:
            set_font(r, size=None)

    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        for ridx, row in enumerate(table.rows):
            no_split(row)
            if ridx == 0:
                repeat_header(row)
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                shade(cell, "243F58" if ridx == 0 else ("F3F6F8" if ridx % 2 == 0 else "FFFFFF"))
                for p in cell.paragraphs:
                    p.paragraph_format.space_after = Pt(1.5)
                    p.paragraph_format.line_spacing = 1.12
                    for r in p.runs:
                        set_font(r,size=7.1 if is_quick else 8.3,bold=(ridx == 0),color=(RGBColor(255,255,255) if ridx == 0 else RGBColor(40,51,62)))
        # Avoid giant gaps before narrow tables.
        try:
            table.style = "Table Grid"
        except KeyError: pass

    head = sec.header.paragraphs[0]
    head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    head.add_run("NITI  |  คู่มือปฏิบัติการฉุกเฉิน  |  ฉบับร่าง")
    for r in head.runs: set_font(r,size=7.5,bold=True,color=GRAY)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("เอกสารตัวอย่าง — ต้องปรับตามผังและกำลังเจ้าหน้าที่จริง  •  ")
    append_page_field(footer)
    for r in footer.runs: set_font(r,size=7,color=GRAY)
    doc.core_properties.title = "แผนปฏิบัติการในภาวะฉุกเฉินสำหรับหมู่บ้านจัดสรร"
    doc.core_properties.subject = "ต้นแบบแผนสำหรับนิติบุคคลและ รปภ."
    doc.core_properties.author = "Emergency planning template"
    doc.save(path)

def build(stem, is_quick=False):
    src = ROOT / "docs" / stem
    if not src.exists():
        raise FileNotFoundError(src)
    target = OUT / ("NITI_Guard_Quick_Card_TH.docx" if is_quick else "NITI_Emergency_Plan_TH.docx")
    run("pandoc", str(src), "--from=markdown+pipe_tables+task_lists", "--to=docx",
        "--standalone", "--output", str(target))
    enhance_docx(target, is_quick)
    with tempfile.TemporaryDirectory(prefix="niti-lo-") as profile:
        run("libreoffice", "-env:UserInstallation=file://" + profile, "--headless",
            "--convert-to", "pdf", "--outdir", str(OUT), str(target))
    pdf = target.with_suffix(".pdf")
    if not pdf.exists() or pdf.stat().st_size < 5000:
        raise RuntimeError("PDF conversion failed: " + str(pdf))
    print("CREATED", target, target.stat().st_size, pdf, pdf.stat().st_size)

if __name__ == "__main__":
    build("EMERGENCY_PLAN_TH.md", False)
    build("QUICK_CARD_TH.md", True)
