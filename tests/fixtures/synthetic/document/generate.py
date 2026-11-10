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


def rpr(font_cn="宋体", font_western="Times New Roman", size=21):
    return (
        f'<w:rPr><w:rFonts w:ascii="{font_western}" w:hAnsi="{font_western}" w:eastAsia="{font_cn}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
    )


def header_xml(text, *, font_cn="宋体", font_western="Times New Roman", size=21,
               align="center", via_style=False):
    """R2：篇眉宋体 10.5 磅（sz=21）居中。

    via_style=True 时刻意不在 run 上写 rPr，改用 pStyle 引用 styles.xml 里的
    Header 样式，用来验证检查器真的展开了样式链，不是只看显式 rPr。
    """
    body = ""
    if text:
        direct = "" if via_style else rpr(font_cn=font_cn, font_western=font_western, size=size)
        body = f'<w:r>{direct}<w:t>{text}</w:t></w:r>'
    ppr = ""
    if via_style:
        ppr = '<w:pPr><w:pStyle w:val="Header"/></w:pPr>'
    elif align != "center":
        ppr = f'<w:pPr><w:jc w:val="{align}"/></w:pPr>'
    else:
        ppr = '<w:pPr><w:jc w:val="center"/></w:pPr>'
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr xmlns:w="{W_NS}"><w:p>{ppr}{body}</w:p></w:hdr>'
    )


def footer_xml(mode, *, font_cn="宋体", size=21, align="center", via_style=False):
    """mode：'field' 用 PAGE 域，'text' 手打数字，'empty' 空页脚。

    页码显示结果那个 run 决定页码长什么样，rPr 写在它上面；
    fldChar / instrText 那些 run 不产生可见文字，检查器会跳过。
    """
    direct = "" if via_style else rpr(font_cn=font_cn, size=size)
    tpl = f'<w:r>{direct}<w:t>{{}}</w:t></w:r>'
    if mode == "field":
        runs = (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + tpl.format("1") +
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        )
    elif mode == "text":
        runs = tpl.format("1")
    else:
        runs = ""
    ppr = ""
    if via_style:
        ppr = '<w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
    elif align != "center":
        ppr = f'<w:pPr><w:jc w:val="{align}"/></w:pPr>'
    else:
        ppr = '<w:pPr><w:jc w:val="center"/></w:pPr>'
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr xmlns:w="{W_NS}"><w:p>{ppr}{runs}</w:p></w:ftr>'
    )


def styles_xml(with_hf_styles=False):
    """with_hf_styles=True 时把 R2/R3 的字体字号定义在样式里，页眉页脚靠 pStyle 继承。"""
    extra = ""
    if with_hf_styles:
        extra = (
            f'<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/>'
            f'<w:pPr><w:jc w:val="center"/></w:pPr>{rpr()}</w:style>'
            f'<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>'
            f'<w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr>{rpr()}</w:style>'
        )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
        "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
        f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        f"{extra}</w:styles>"
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


def content_types(with_header_footer, with_styles=True):
    extra = ""
    if with_header_footer:
        extra = (
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        )
    styles_override = (
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        if with_styles else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        f'{styles_override}'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        f"{extra}</Types>"
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)


def document_rels(with_header_footer, with_styles=True):
    extra = ""
    if with_header_footer:
        extra = (
            '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        )
    styles_rel = (
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        if with_styles else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{styles_rel}'
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
    no_styles = config.get("no_styles", False)

    header_kwargs = dict(
        font_cn=config.get("header_font", "宋体"),
        font_western=config.get("header_western", "Times New Roman"),
        size=config.get("header_size", 21),
        align=config.get("header_align", "center"),
        via_style=config.get("hf_via_style", False),
    )
    footer_kwargs = dict(
        font_cn=config.get("footer_font", "宋体"),
        size=config.get("footer_size", 21),
        align=config.get("footer_align", "center"),
        via_style=config.get("hf_via_style", False),
    )

    body = ""
    body += paragraph("合成样本正文第一段，用于占位。")
    if config.get("toc") == "auto":
        body += paragraph("目录")
        body += toc_field_paragraph()
    elif config.get("toc") == "manual":
        body += paragraph("目录")
        body += paragraph("第一章 引言    1")
    if config.get("true_blank_page"):
        # 两个连续的空段落各带一个分页符：Word 实测第 2 页整页空白。
        body += empty_paragraph_with_page_break()
        body += empty_paragraph_with_page_break()
    elif config.get("blank_page"):
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
        ("[Content_Types].xml", content_types(has_parts, with_styles=not no_styles)),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(body)),
        ("word/_rels/document.xml.rels", document_rels(has_parts, with_styles=not no_styles)),
        ("word/settings.xml", settings_xml(config.get("even_odd", False))),
    ]
    if not no_styles:
        entries.append(("word/styles.xml", styles_xml(config.get("hf_via_style", False))))
    if header_ref:
        entries.append(("word/header1.xml", header_xml(header_text, **header_kwargs)))
    if footer_ref:
        entries.append(("word/footer1.xml", footer_xml(footer_mode, **footer_kwargs)))

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

    # ---- 第二档（11/29）：R2 篇眉 / R3 页码的字体字号与居中 ----
    # 每个样本只偏离一个维度，正样本在基线里。
    "document_header_font_wrong.docx": dict(header_font="黑体"),
    "document_header_size_wrong.docx": dict(header_size=24),
    "document_header_left_aligned.docx": dict(header_align="left"),
    "document_footer_font_wrong.docx": dict(footer_font="黑体"),
    "document_footer_size_wrong.docx": dict(footer_size=24),
    "document_footer_left_aligned.docx": dict(footer_align="left"),
    # 合规：全西文篇眉按 R2 的 Abstract 条款用 Times New Roman
    "document_abstract_header_correct.docx": dict(header_text="Abstract"),
    # 偏差：全西文篇眉却用了 Arial
    "document_abstract_header_wrong.docx": dict(header_text="Abstract", header_western="Arial"),
    # 合规：字体字号居中都写在 styles.xml 的 Header / Footer 样式里，靠 pStyle 继承
    "document_hf_style_inherited.docx": dict(hf_via_style=True),
    # 0 条问题但不能算「查过」：没有 styles.xml，页眉页脚又没有显式 rPr/jc，
    # 三个维度都判不出来。报告必须把「压根没查」写出来。
    "document_no_styles_part.docx": dict(no_styles=True, hf_via_style=True),
    # 对照：同样没有 styles.xml，但 run 和段落显式写了值，这些维度照样能判
    "document_no_styles_explicit.docx": dict(no_styles=True),
    # 合规：字体写成 SimSun / TimesNewRoman（无空格），检查器要认这些别名
    "document_font_aliases.docx": dict(header_font="SimSun", header_western="TimesNewRoman", footer_font="SimSun"),
    # 偏差：连续两个空段落各带分页符，Word 实测会真的多出一整页空白
    "document_true_blank_page.docx": dict(true_blank_page=True),
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
