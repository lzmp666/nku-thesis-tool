"""Generate the two small rule-analysis fixtures.

The documents are synthetic and contain no real student or thesis data. They are
intentionally minimal: one chapter, three body paragraphs and a references
section, so the structure classifier can establish a closed body boundary.
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent


def set_font(run, *, east_asia: str | None = None, latin: str | None = None,
             size: float | None = None, bold: bool | None = None) -> None:
    if east_asia is not None:
        run.font.name = latin or east_asia
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), east_asia)
    if latin is not None:
        run.font.name = latin
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:ascii"), latin)
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), latin)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def set_paragraph(paragraph, *, alignment, line_rule, line_value=None,
                  before: float = 0, after: float = 0,
                  first_line: float | None = None) -> None:
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing_rule = line_rule
    if line_value is not None:
        paragraph.paragraph_format.line_spacing = Pt(line_value)
    if first_line is not None:
        paragraph.paragraph_format.first_line_indent = Pt(first_line)


def set_outline(paragraph, level: int = 0) -> None:
    properties = paragraph._p.get_or_add_pPr()
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level))
    properties.append(outline)


def heading(document, text: str, *, size: float = 16, before: float = 24, after: float = 18) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    set_font(run, east_asia="黑体", latin="Times New Roman", size=size, bold=True)
    set_paragraph(paragraph, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  line_rule=WD_LINE_SPACING.SINGLE, before=before, after=after)
    set_outline(paragraph, 0)


def body(document, text: str, *, east_asia: str = "宋体", latin: str = "Times New Roman",
         size: float = 12, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
         line_rule=WD_LINE_SPACING.EXACTLY, line_value: float | None = 20,
         before: float = 0, after: float = 0, first_line: float | None = 24) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    set_font(run, east_asia=east_asia, latin=latin, size=size, bold=False)
    set_paragraph(paragraph, alignment=alignment, line_rule=line_rule,
                  line_value=line_value, before=before, after=after, first_line=first_line)


def base_document() -> Document:
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.8)
    section.bottom_margin = Cm(3.8)
    section.left_margin = Cm(3.2)
    section.right_margin = Cm(3.2)

    style = document.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(12)

    heading(document, "第一章  绪论")
    body(document, "这是一段用于规则分析回归的合成正文，Word 文档描述研究背景与方法，不包含真实数据。")
    body(document, "第二段继续说明研究问题与分析路径，保持宋体小四、固定值二十磅行距。")
    body(document, "第三段给出合成结论，仅用于验证段落级字体、字号、对齐、缩进和行距检查。")
    heading(document, "参考文献")
    body(document, "南开大学研究生院. 南开大学研究生学位论文写作规范（2026版）[S]. 2026.",
         east_asia="宋体", latin="Times New Roman", size=10.5,
         alignment=WD_ALIGN_PARAGRAPH.LEFT, line_rule=WD_LINE_SPACING.EXACTLY,
         line_value=16, first_line=None)
    return document


def main() -> None:
    compliant = base_document()
    compliant.save(ROOT / "analysis_compliant.docx")

    noncompliant = base_document()
    # Replace the first body paragraph's formatting with four deliberate deviations.
    target = noncompliant.paragraphs[1]
    for run in target.runs:
        set_font(run, east_asia="黑体", latin="Arial", size=10.5, bold=False)
    set_paragraph(target, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  line_rule=WD_LINE_SPACING.SINGLE, before=6, after=6, first_line=0)
    noncompliant.save(ROOT / "analysis_noncompliant.docx")


if __name__ == "__main__":
    main()
