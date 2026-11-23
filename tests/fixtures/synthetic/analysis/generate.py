"""analysis/ 的统一生成入口：规则分析回归样本 + 内存重建。

两份样本出自同一个基础文档，区别只在第一段正文：

* `analysis_compliant.docx` —— A4、3.8/3.2 cm 页边距、宋体/Times New Roman 12 pt、
  两端对齐、首行缩进 2 字符、固定值 20 pt。预期 0 条不符合、未检查清单非空。
* `analysis_noncompliant.docx` —— 第一段改成黑体/Arial 10.5 pt、左对齐、0 缩进、
  单倍行距、段前段后各 6 pt。预期 8 条不符合。

两份都由 `base_document()` 现造，不从磁盘读另一份：基础文档被改坏时两份会一起变形，
逐部件重建能同时抓出来，不会出现「源改了、另一份还是旧的，两份都算过」。

顶层只导入标准库，`verify_fixtures.py` 的清单核对照样能在干净检出下读到 `SAMPLES`；
真正重建时才加载 python-docx（`_require_docx()`）。

最终字节过一遍 `_normalize()`：条目顺序与解压后的部件原样保留，ZIP 条目时间戳统一成
`FIXED_TIME`（与 `format-samples/`、`full-thesis/` 同一个常量）。python-docx 的
`save()` 默认拿当前时刻当时间戳，不归一化的话同一个生成器换分钟跑出的整文件字节就不同，
`--check` 只能退化成逐部件比对。归一化之后 `--check` 直接比整文件字节。

用法（在本目录下）：

    python generate.py --check    # 重建两份，与磁盘比整文件字节，不写盘
    python generate.py            # 重建并写盘（覆盖本目录的已提交样本）

依赖 python-docx（`--check` 也要，因为要对重建结果算字节）。
"""
from __future__ import annotations

import argparse
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

HERE = Path(__file__).resolve().parent

# 与 format-samples/、full-thesis/ 同一个时间常量：ZIP 可存的最早时间。
FIXED_TIME = (2000, 1, 1, 0, 0, 0)

COMPLIANT = "analysis_compliant.docx"
NONCOMPLIANT = "analysis_noncompliant.docx"
SAMPLES = (COMPLIANT, NONCOMPLIANT)

Document = None


def _require_docx():
    """重建路径才加载第三方依赖；缺失时给出可执行的提示，不抛裸 traceback。"""
    global Document, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, OxmlElement, qn, Cm, Pt
    try:
        from docx import Document as _Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _ALIGN
        from docx.enum.text import WD_LINE_SPACING as _SPACING
        from docx.oxml import OxmlElement as _OxmlElement
        from docx.oxml.ns import qn as _qn
        from docx.shared import Cm as _Cm, Pt as _Pt
    except ImportError as exc:
        raise SystemExit(
            "重建样本需要 python-docx：\n"
            "  python -m pip install -r tests/fixtures/synthetic/analysis/requirements.txt\n"
            "原始错误：%s" % exc
        )
    Document = _Document
    WD_ALIGN_PARAGRAPH = _ALIGN
    WD_LINE_SPACING = _SPACING
    OxmlElement = _OxmlElement
    qn = _qn
    Cm = _Cm
    Pt = _Pt


def build_bytes(name=COMPLIANT):
    """返回指定样本的 docx 字节。未知样本名直接报错，不静默回落到主样本。"""
    if name not in SAMPLES:
        raise KeyError("未知样本：%s（可选 %s）" % (name, "、".join(SAMPLES)))
    _require_docx()
    document = base_document()
    if name == NONCOMPLIANT:
        deviate_first_body_paragraph(document)
    buffer = BytesIO()
    document.save(buffer)
    return _normalize(buffer.getvalue())


def _normalize(data):
    """固定 ZIP 条目时间戳；条目顺序与解压后的部件字节原样保留。

    与 `full-thesis/generate.py` 的 `_normalize()` 同形：只重写容器头部字段，
    不排序条目（排序会改顺序，反而对不上 python-docx 的写入顺序）。
    """
    buffer = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(buffer, "w") as target:
        for info in source.infolist():
            item = ZipInfo(info.filename, date_time=FIXED_TIME)
            item.compress_type = ZIP_DEFLATED
            item.external_attr = info.external_attr
            target.writestr(item, source.read(info.filename))
    return buffer.getvalue()


def set_font(run, *, east_asia=None, latin=None, size=None, bold=None):
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
                  before=0, after=0, first_line=None):
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing_rule = line_rule
    if line_value is not None:
        paragraph.paragraph_format.line_spacing = Pt(line_value)
    if first_line is not None:
        paragraph.paragraph_format.first_line_indent = Pt(first_line)


def set_outline(paragraph, level=0):
    properties = paragraph._p.get_or_add_pPr()
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level))
    properties.append(outline)


def heading(document, text, *, size=16, before=24, after=18):
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    set_font(run, east_asia="黑体", latin="Times New Roman", size=size, bold=True)
    set_paragraph(paragraph, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  line_rule=WD_LINE_SPACING.SINGLE, before=before, after=after)
    set_outline(paragraph, 0)


def body(document, text, *, east_asia="宋体", latin="Times New Roman",
         size=12, alignment=None, line_rule=None, line_value=20,
         before=0, after=0, first_line=24):
    # 默认值在函数体内取：python-docx 的枚举要延迟导入，写在默认参数里会在
    # import 阶段就求值，干净检出下直接 NameError。
    if alignment is None:
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if line_rule is None:
        line_rule = WD_LINE_SPACING.EXACTLY
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    set_font(run, east_asia=east_asia, latin=latin, size=size, bold=False)
    set_paragraph(paragraph, alignment=alignment, line_rule=line_rule,
                  line_value=line_value, before=before, after=after, first_line=first_line)


def base_document():
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.8)
    section.bottom_margin = Cm(3.8)
    section.left_margin = Cm(3.2)
    section.right_margin = Cm(3.2)
    # R1（4.1 表4.1，S2026 原文）：页眉、页脚距边界 3.0cm，装订线 0cm。
    # python-docx 不写这两项时 Word 默认 1.27cm，会让"合规样本"在页眉页脚规则上误报。
    section.header_distance = Cm(3.0)
    section.footer_distance = Cm(3.0)

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


def deviate_first_body_paragraph(document):
    """第一段正文故意偏离：字体、字号、对齐、缩进、行距、段前后六项。"""
    target = document.paragraphs[1]
    for run in target.runs:
        set_font(run, east_asia="黑体", latin="Arial", size=10.5, bold=False)
    set_paragraph(target, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  line_rule=WD_LINE_SPACING.SINGLE, before=6, after=6, first_line=0)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="只重建比对，不写盘")
    arguments = parser.parse_args()

    failed = 0
    for name in SAMPLES:
        target = HERE / name
        data = build_bytes(name)
        if arguments.check:
            if not target.exists():
                print("缺失：%s" % name)
                failed += 1
                continue
            if data == target.read_bytes():
                print("校验通过：%s（整文件字节一致）" % name)
            else:
                print("与重建结果不一致：%s" % name)
                failed += 1
            continue
        target.write_bytes(data)
        print("已写出：%s（%d 字节）" % (name, len(data)))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
