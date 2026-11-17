"""R14-R17 尾注部件格式检查的合成样本。

覆盖 general.yaml 4.8 表4.7：参考文献（宋体/TNR 10.5pt、固定16pt、段前后0）、
附录（宋体/TNR 12pt、两端对齐、首行缩进2汉字符、固定20pt、段前后0）、
符号表与个人简历（宋体/TNR 10.5pt、固定16pt、段前后0）、致谢（仿宋12pt、固定16pt、段前后0）。
每个文件只偏离一个维度，便于断言定位准确、不误报。

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
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
FIXED_TIME = (2026, 1, 1, 0, 0, 0)

# 规范值（general.yaml 4.8 表4.7）
SMALL = 21          # 10.5pt -> 半磅
MID = 24            # 12pt -> 半磅
EXACT16 = 320        # 固定值 16 磅
EXACT20 = 400        # 固定值 20 磅
ZERO = 0


def run(text, *, eastAsia="宋体", ascii_font="Times New Roman", size=SMALL):
    fonts = f'<w:rFonts w:ascii="{ascii_font}" w:hAnsi="{ascii_font}" w:eastAsia="{eastAsia}"/>'
    return (f'<w:r><w:rPr>{fonts}<w:sz w:val="{size}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def paragraph(runs, *, align=None, line=EXACT16, line_rule="exact",
              before=ZERO, after=ZERO, first_line_chars=None, first_line_twips=None):
    parts = []
    if align:
        parts.append(f'<w:jc w:val="{align}"/>')
    spacing = [f'w:before="{before}"', f'w:after="{after}"']
    if line is not None and line_rule is not None:
        spacing += [f'w:line="{line}"', f'w:lineRule="{line_rule}"']
    parts.append(f'<w:spacing {" ".join(spacing)}/>')
    if first_line_chars is not None:
        ind = [f'w:firstLineChars="{first_line_chars}"']
        if first_line_twips is not None:
            ind.append(f'w:firstLine="{first_line_twips}"')
        parts.append(f'<w:ind {" ".join(ind)}/>')
    return f"<w:p><w:pPr>{''.join(parts)}</w:pPr>{runs}</w:p>"


def title(text):
    return paragraph(run(text, size=MID))


def reference_body(overrides):
    return paragraph(
        run(overrides.get("ref_text", "参考文献正文示例"),
            eastAsia=overrides.get("ref_cn_font", "宋体"),
            ascii_font=overrides.get("ref_western_font", "Times New Roman"),
            size=overrides.get("ref_size", SMALL)),
        line=overrides.get("ref_line", EXACT16),
        line_rule=overrides.get("ref_line_rule", "exact"),
        after=overrides.get("ref_after", ZERO))


def appendix_body(overrides, index=1):
    first_chars = overrides.get("app_indent_chars", 200)
    first_twips = overrides.get("app_indent_twips", 480)
    size = overrides.get("app_size", MID)
    math = ('<m:oMath><m:r><m:t>a+b</m:t></m:r></m:oMath>'
            if overrides.get("app_math") else "")
    return paragraph(
        run(f"附录第{index}段正文", eastAsia=overrides.get("app_cn_font", "宋体"),
            ascii_font=overrides.get("app_western_font", "Times New Roman"), size=size) + math,
        align=overrides.get("app_align", "both"),
        line=overrides.get("app_line", EXACT20),
        line_rule=overrides.get("app_line_rule", "exact"),
        after=overrides.get("app_after", ZERO),
        first_line_chars=first_chars, first_line_twips=first_twips)


def symbol_body(overrides):
    return paragraph(
        run("符号说明正文", eastAsia=overrides.get("sym_cn_font", "宋体"),
            ascii_font=overrides.get("sym_western_font", "Times New Roman"),
            size=overrides.get("sym_size", SMALL)),
        line=overrides.get("sym_line", EXACT16),
        line_rule=overrides.get("sym_line_rule", "exact"),
        after=overrides.get("sym_after", ZERO))


def ack_body(overrides):
    return paragraph(
        run("致谢正文示例", eastAsia=overrides.get("ack_cn_font", "仿宋"),
            size=overrides.get("ack_size", MID)),
        line=overrides.get("ack_line", EXACT16),
        line_rule=overrides.get("ack_line_rule", "exact"),
        after=overrides.get("ack_after", ZERO))


def build_body(overrides):
    if overrides.get("no_sections"):
        return paragraph(run("这里只是普通正文，没有尾注部件标题")) + \
               paragraph(run("第二段普通正文"))
    parts = [
        title("参考文献"), reference_body(overrides),
        title("附录 A"), appendix_body(overrides, 1),
    ]
    if overrides.get("two_appendices"):
        parts += [title("附录 B"), appendix_body(overrides, 2)]
    parts += [
        title("主要符号表"), symbol_body(overrides),
        title("致谢"), ack_body(overrides),
    ]
    return "".join(parts)


SAMPLES = {
    "back_matter_compliant.docx": {},
    "back_matter_reference_cn_font_wrong.docx": {"ref_cn_font": "仿宋"},
    "back_matter_reference_size_wrong.docx": {"ref_size": MID},
    "back_matter_reference_spacing_wrong.docx": {"ref_line_rule": "auto", "ref_line": 240},
    "back_matter_appendix_alignment_left.docx": {"app_align": "left"},
    "back_matter_appendix_indent_wrong.docx": {"app_indent_chars": 100, "app_indent_twips": 240},
    "back_matter_appendix_size_wrong.docx": {"app_size": SMALL},
    "back_matter_appendix_math_line_exempt.docx": {"app_math": True, "app_line_rule": "auto", "app_line": 240},
    "back_matter_symbol_western_font_wrong.docx": {"sym_western_font": "Arial"},
    "back_matter_ack_cn_font_wrong.docx": {"ack_cn_font": "宋体"},
    "back_matter_ack_spacing_wrong.docx": {"ack_after": 240},
    "back_matter_two_appendices.docx": {"two_appendices": True},
    "back_matter_no_sections.docx": {"no_sections": True},
}

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>')
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/></Relationships>')
STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>'
    '<w:sz w:val="21"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style></w:styles>')
SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:settings xmlns:w="{W_NS}"/>')


def document_xml(body):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:m="{M_NS}">'
        f'<w:body>{body}<w:sectPr/></w:body></w:document>')


def build_bytes(name):
    overrides = SAMPLES[name]
    entries = [
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(build_body(overrides))),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", SETTINGS),
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
