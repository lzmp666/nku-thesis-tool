"""R4/R5 中英文摘要格式检查的合成样本。

覆盖中文摘要标题（黑体18磅加粗居中、单倍行距、段前24段后18）、正文（宋体12磅、
固定值行距20磅、段前后0）、关键词加粗，以及英文摘要对应的 Arial / Times New Roman
条款。每个文件只偏离一个维度，便于断言检查器定位准确、不误报。

只用标准库生成，ZIP 条目顺序与时间戳固定，同版本下重复生成可逐字节核对。

用法（本目录下）：
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
FIXED_TIME = (2026, 1, 1, 0, 0, 0)

# R4/R5 规范值（general.yaml 4.5 表4.4）
TITLE_SIZE = 36          # 18pt -> 半磅
BODY_SIZE = 24           # 12pt -> 半磅
SINGLE_LINE = 240        # 单倍行距
EXACT20 = 400            # 固定值 20 磅 = 400 twips
TITLE_BEFORE, TITLE_AFTER = 480, 360   # 24 磅 / 18 磅
ZERO = 0


def run(text, *, eastAsia="宋体", ascii_font="Times New Roman", size=BODY_SIZE, bold=False):
    fonts = f'<w:rFonts w:ascii="{ascii_font}" w:hAnsi="{ascii_font}" w:eastAsia="{eastAsia}"/>'
    bold_tag = "<w:b/>" if bold else ""
    return (f'<w:r><w:rPr>{fonts}{bold_tag}<w:sz w:val="{size}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def paragraph(text_runs, *, align=None, before=None, after=None, line=None, line_rule=None):
    ppr_parts = []
    if align:
        ppr_parts.append(f'<w:jc w:val="{align}"/>')
    spacing_attrs = []
    if before is not None:
        spacing_attrs.append(f'w:before="{before}"')
    if after is not None:
        spacing_attrs.append(f'w:after="{after}"')
    if line is not None:
        spacing_attrs.append(f'w:line="{line}"')
    if line_rule is not None:
        spacing_attrs.append(f'w:lineRule="{line_rule}"')
    if spacing_attrs:
        ppr_parts.append(f'<w:spacing {" ".join(spacing_attrs)}/>')
    ppr = f"<w:pPr>{''.join(ppr_parts)}</w:pPr>" if ppr_parts else ""
    return f"<w:p>{ppr}{text_runs}</w:p>"


def cn_title(**options):
    align = options.pop("align", "center")
    before = options.pop("before", TITLE_BEFORE)
    after = options.pop("after", TITLE_AFTER)
    line = options.pop("line", SINGLE_LINE)
    line_rule = options.pop("line_rule", "auto")
    font = options.pop("font", "黑体")
    r = run(text="摘  要", eastAsia=font, ascii_font="Times New Roman",
            size=options.pop("size", TITLE_SIZE), bold=options.pop("bold", True))
    return paragraph(r, align=align, before=before, after=after, line=line, line_rule=line_rule)


def cn_body(text, **options):
    before = options.pop("before", ZERO)
    after = options.pop("after", ZERO)
    line = options.pop("line", EXACT20)
    line_rule = options.pop("line_rule", "exact")
    font = options.pop("font", "宋体")
    r = run(text=text, eastAsia=font, ascii_font="Times New Roman",
            size=options.pop("size", BODY_SIZE), bold=False)
    return paragraph(r, before=before, after=after, line=line, line_rule=line_rule)


def cn_keyword(**options):
    r = run(text="关  键  词：格式；规范", eastAsia="宋体", size=BODY_SIZE,
            bold=options.pop("bold", True))
    return paragraph(r, before=ZERO, after=ZERO, line=EXACT20, line_rule="exact")


def en_title(**options):
    align = options.pop("align", "center")
    before = options.pop("before", TITLE_BEFORE)
    after = options.pop("after", TITLE_AFTER)
    line = options.pop("line", SINGLE_LINE)
    line_rule = options.pop("line_rule", "auto")
    font = options.pop("font", "Arial")
    r = run(text="Abstract", eastAsia="宋体", ascii_font=font,
            size=options.pop("size", TITLE_SIZE), bold=options.pop("bold", True))
    return paragraph(r, align=align, before=before, after=after, line=line, line_rule=line_rule)


def en_body(text, **options):
    before = options.pop("before", ZERO)
    after = options.pop("after", ZERO)
    line = options.pop("line", EXACT20)
    line_rule = options.pop("line_rule", "exact")
    font = options.pop("font", "Times New Roman")
    r = run(text=text, eastAsia="宋体", ascii_font=font,
            size=options.pop("size", BODY_SIZE), bold=False)
    return paragraph(r, before=before, after=after, line=line, line_rule=line_rule)


def en_keyword(**options):
    r = run(text="Key words: format; specification", ascii_font="Times New Roman",
            size=BODY_SIZE, bold=options.pop("bold", True))
    return paragraph(r, before=ZERO, after=ZERO, line=EXACT20, line_rule="exact")


def styles_xml():
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>'
    )


def settings_xml():
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/></w:settings>')


def content_types():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '</Relationships>'
)


def document_xml(body_xml):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body_xml}</w:body></w:document>')


def build_body(config):
    """config 的键：cn_title / cn_body / cn_keyword / en_title / en_body / en_keyword，
    值是要传给对应构造函数的参数字典；None 表示省略该段落。"""
    parts = []
    if config.get("cn_title", {}) is not None:
        parts.append(cn_title(**config.get("cn_title", {})))
    if config.get("cn_body", {}) is not None:
        parts.append(cn_body("本研究采用样式驱动的排版方法。", **config.get("cn_body", {})))
    if config.get("cn_keyword", {}) is not None:
        parts.append(cn_keyword(**config.get("cn_keyword", {})))
    if config.get("en_title", {}) is not None:
        parts.append(en_title(**config.get("en_title", {})))
    if config.get("en_body", {}) is not None:
        parts.append(en_body("This study adopts a style-driven layout method.", **config.get("en_body", {})))
    if config.get("en_keyword", {}) is not None:
        parts.append(en_keyword(**config.get("en_keyword", {})))
    return "".join(parts)


SAMPLES = {
    # 合规基线：中英文摘要标题、正文、关键词全部符合 R4/R5。
    "abstract_compliant.docx": dict(),
    # 中文摘要标题偏差，每个只动一处。
    "abstract_cn_title_font_wrong.docx": dict(cn_title=dict(font="宋体")),
    "abstract_cn_title_size_wrong.docx": dict(cn_title=dict(size=32)),
    "abstract_cn_title_not_bold.docx": dict(cn_title=dict(bold=False)),
    "abstract_cn_title_left.docx": dict(cn_title=dict(align="left")),
    "abstract_cn_title_spacing_wrong.docx": dict(cn_title=dict(line=360)),
    # 中文摘要正文偏差。
    "abstract_cn_body_font_wrong.docx": dict(cn_body=dict(font="仿宋")),
    "abstract_cn_body_size_wrong.docx": dict(cn_body=dict(size=28)),
    "abstract_cn_body_spacing_wrong.docx": dict(cn_body=dict(line_rule="auto", line=240)),
    "abstract_cn_keyword_not_bold.docx": dict(cn_keyword=dict(bold=False)),
    # 英文摘要偏差。
    "abstract_en_title_font_wrong.docx": dict(en_title=dict(font="Times New Roman")),
    "abstract_en_body_font_wrong.docx": dict(en_body=dict(font="宋体")),
    "abstract_en_keyword_not_bold.docx": dict(en_keyword=dict(bold=False)),
    # 无摘要段落：报告必须写「未识别」，不能假装合规。
    "abstract_no_sections.docx": dict(cn_title=None, cn_body=None, cn_keyword=None,
                                      en_title=None, en_body=None, en_keyword=None),
}


def build_bytes(name):
    entries = [
        ("[Content_Types].xml", content_types()),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(build_body(SAMPLES[name]))),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/styles.xml", styles_xml()),
        ("word/settings.xml", settings_xml()),
    ]
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
        if problems == 0:
            print("全部样本逐字节一致")
        else:
            print(f"不一致样本：{problems} 个")
        return 1 if problems else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
