"""Issue #112 文档级版式检查的合成样本。

覆盖纸张、页边距、页眉页脚距离、装订线、页眉页脚有无内容、页码域、页码格式、
奇偶页不同、自动目录、疑似空白页。每个文件只偏离一个维度。

只用标准库生成。ZIP 条目顺序与时间戳固定，同一 Python 版本下重复生成可逐字节核对。

用法（在本目录下）：
    python generate.py            # 重新生成全部样本
    python generate.py --check    # 只比对，不写盘
"""
import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 1cm ≈ 567 twips
A4 = (11906, 16838)
LETTER = (12240, 15840)
# R1：上下 3.8cm、左右 3.2cm、页眉页脚 3.0cm、装订线 0
SPEC_MARGINS = dict(top=2155, right=1814, bottom=2155, left=1814)
SPEC_HF = dict(header=1701, footer=1701)

FIXED_TIME = (2026, 1, 1, 0, 0, 0)


def header_xml(text):
    body = "" if not text else (
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
        f'<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>{text}</w:t></w:r>'
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{body}</w:p></w:hdr>'
    )


def footer_xml(mode):
    """mode：'field' 用 PAGE 域，'text' 手打数字，'empty' 空页脚。"""
    if mode == "field":
        runs = (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )
    elif mode == "text":
        runs = '<w:r><w:t>1</w:t></w:r>'
    else:
        runs = ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{runs}</w:p></w:ftr>'
    )


def styles_xml():
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
        "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        "</w:styles>"
    )


def settings_xml(even_odd):
    extra = '<w:evenAndOddHeaders/>' if even_odd else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/>{extra}</w:settings>'
    )


def paragraph(text, page_break_before=False):
    ppr = "<w:ppr/>"
    if page_break_before:
        ppr = "<w:pPr><w:pageBreakBefore/></w:pPr>"
    else:
        ppr = ""
    if not text:
        return f"<w:p>{ppr}</w:p>"
    return (
        f'<w:p>{ppr}<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/>'
        f'<w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>'
    )


def empty_paragraph_with_page_break():
    """空段落里带分页符：Word 里会多出一个空白页。"""
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def toc_field_paragraph():
    return (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>第一章 引言\t1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )


def sect_pr(paper, margins, hf, gutter, number_format=None, title_pg=False,
            header_ref=True, footer_ref=True):
    parts = []
    if header_ref:
        parts.append('<w:headerReference w:type="default" r:id="rId10"/>')
    if footer_ref:
        parts.append('<w:footerReference w:type="default" r:id="rId11"/>')
    parts.append(f'<w:pgSz w:w="{paper[0]}" w:h="{paper[1]}"/>')
    parts.append(
        f'<w:pgMar w:top="{margins["top"]}" w:right="{margins["right"]}" '
        f'w:bottom="{margins["bottom"]}" w:left="{margins["left"]}" '
        f'w:header="{hf["header"]}" w:footer="{hf["footer"]}" w:gutter="{gutter}"/>'
    )
    if number_format:
        parts.append(f'<w:pgNumType w:fmt="{number_format}" w:start="1"/>')
    if title_pg:
        parts.append("<w:titlePg/>")
    return "<w:sectPr>" + "".join(parts) + "</w:sectPr>"


def document_xml(body_xml):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body_xml}</w:body></w:document>'
    )


def content_types(with_header_footer):
    extra = ""
    if with_header_footer:
        extra = (
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        f"{extra}</Types>"
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)


def document_rels(with_header_footer):
    extra = ""
    if with_header_footer:
        extra = (
            '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        f"{extra}</Relationships>"
    )


def build_docx(config):
    paper = config.get("paper", A4)
    margins = config.get("margins", SPEC_MARGINS)
    hf = config.get("hf", SPEC_HF)
    gutter = config.get("gutter", 0)
    header_text = config.get("header_text", "第一章 引言")
    footer_mode = config.get("footer", "field")
    number_format = config.get("number_format")
    header_ref = config.get("header_ref", True)
    footer_ref = config.get("footer_ref", True)
    has_parts = header_ref or footer_ref

    body = ""
    body += paragraph("合成样本正文第一段，用于占位。")
    if config.get("toc") == "auto":
        body += paragraph("目录")
        body += toc_field_paragraph()
    elif config.get("toc") == "manual":
        body += paragraph("目录")
        body += paragraph("第一章 引言    1")
    if config.get("blank_page"):
        body += empty_paragraph_with_page_break()
    body += paragraph("合成样本正文第二段。")

    if config.get("sections", 1) > 1:
        # 前一节 + 分节符，两节各自写 sectPr
        first_sect = sect_pr(paper, margins, hf, gutter, number_format, header_ref=header_ref, footer_ref=footer_ref)
        body += f"<w:p><w:pPr>{first_sect}</w:pPr></w:p>"
        body += paragraph("第二节正文。")
        body += paragraph("第二节第二段，设置段前分页。", page_break_before=config.get("blank_page", False))

    final_sect = sect_pr(paper, margins, hf, gutter, number_format, header_ref=header_ref, footer_ref=footer_ref)
    body += final_sect

    entries = [
        ("[Content_Types].xml", content_types(has_parts)),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(body)),
        ("word/_rels/document.xml.rels", document_rels(has_parts)),
        ("word/styles.xml", styles_xml()),
        ("word/settings.xml", settings_xml(config.get("even_odd", False))),
    ]
    if header_ref:
        entries.append(("word/header1.xml", header_xml(config.get("header_text", "第一章 引言"))))
    if footer_ref:
        entries.append(("word/footer1.xml", footer_xml(footer_mode)))

    return entries


SAMPLES = {
    # 合规基线：A4、规范四边距、规范页眉页脚距离、装订线 0、页眉有内容、页脚是 PAGE 域。
    "document_baseline_correct.docx": dict(),
    # 偏差：Letter 纸张
    "document_paper_wrong.docx": dict(paper=LETTER),
    # 偏差：上下页边距 2.5cm（两条）
    "document_margin_wrong.docx": dict(margins=dict(top=1417, right=1814, bottom=1417, left=1814)),
    # 偏差：页眉页脚距边界 1.27cm（两条）
    "document_hf_distance_wrong.docx": dict(hf=dict(header=720, footer=720)),
    # 偏差：装订线 1cm
    "document_gutter_wrong.docx": dict(gutter=567),
    # 偏差：整个文档没有页眉引用
    "document_no_header.docx": dict(header_ref=False),
    # 偏差：页眉引用了部件，但里面没字
    "document_empty_header.docx": dict(header_text=""),
    # 偏差：页脚是手打数字，不是 PAGE 域
    "document_footer_text.docx": dict(footer="text"),
    # 偏差：开了奇偶页不同（R2 要求奇偶页相同）
    "document_even_odd.docx": dict(even_odd=True),
    # 偏差：目录是手打的
    "document_toc_manual.docx": dict(toc="manual"),
    # 合规：目录是 TOC 域
    "document_toc_auto.docx": dict(toc="auto"),
    # 偏差：空段落里带分页符（疑似空白页，惯例项）
    "document_blank_page.docx": dict(blank_page=True),
    # 偏差：两节页码格式都是 decimal，没区分前置与正文
    "document_sections_same_numbering.docx": dict(sections=2, number_format="decimal"),
    # 合规：两节分别用罗马与阿拉伯
    "document_sections_mixed_numbering.docx": dict(sections=2),
}


def build_bytes(name):
    """两节且要求区分编号的那个样本，两节格式不同，需要单独处理。"""
    config = SAMPLES[name]

    if name == "document_sections_mixed_numbering.docx":
        paper, margins, hf = A4, SPEC_MARGINS, SPEC_HF
        body = paragraph("前置部分正文，页码用大写罗马数字。")
        first = sect_pr(paper, margins, hf, 0, "upperRoman")
        body += f"<w:p><w:pPr>{first}</w:pPr></w:p>"
        body += paragraph("正文第一章。")
        body += sect_pr(paper, margins, hf, 0, "decimal")
        entries = [
            ("[Content_Types].xml", content_types(True)),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", document_xml(body)),
            ("word/_rels/document.xml.rels", document_rels(True)),
            ("word/styles.xml", styles_xml()),
            ("word/settings.xml", settings_xml(False)),
            ("word/header1.xml", header_xml("第一章 引言")),
            ("word/footer1.xml", footer_xml("field")),
        ]
    else:
        entries = build_docx(config)

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for entry_name, data in entries:
            info = zipfile.ZipInfo(entry_name, date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8"))
    return buf.getvalue()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="只比对，不写盘")
    args = parser.parse_args()

    problems = 0
    for name in SAMPLES:
        data = build_bytes(name)
        target = HERE / name
        if args.check:
            if not target.exists():
                print(f"缺失：{name}")
                problems += 1
            elif target.read_bytes() != data:
                print(f"与生成结果不一致：{name}")
                problems += 1
        else:
            target.write_bytes(data)
            print(f"已生成：{name}（{len(data)} 字节）")

    if args.check:
        print("校验通过" if problems == 0 else f"校验失败：{problems} 个文件")
        return 1 if problems else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
