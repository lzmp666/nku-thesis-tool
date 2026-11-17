"""R11 正文段落格式检查的合成样本（补齐 4 个排版维度）。

覆盖 general.yaml 4.7 表4.6 的正文排版值：两端对齐、首行缩进 2 汉字符、
固定值行距 20 磅、段前后 0 磅。字体/字号由 #43 的 R11 实现负责，本批样本
全部用合规的宋体/Times New Roman 12 磅，不制造该维度的差异。

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

# 规范值（general.yaml 4.7 表4.6）
MID = 24            # 12pt -> 半磅
EXACT20 = 400       # 固定值 20 磅
ZERO = 0


def run(text, *, size=MID):
    fonts = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
    return (f'<w:r><w:rPr>{fonts}<w:sz w:val="{size}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>')


def run_math(text):
    fonts = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
    return (f'<w:r><w:rPr>{fonts}<w:sz w:val="{MID}"/></w:rPr>'
            f'<w:t>{text}</w:t></w:r>'
            '<m:oMath><m:r><m:t>a+b=c</m:t></m:r></m:oMath>')


def paragraph(runs, *, align="both", line=EXACT20, line_rule="exact",
              before=ZERO, after=ZERO, first_line_chars=200, first_line_twips=480):
    parts = []
    if align:
        parts.append(f'<w:jc w:val="{align}"/>')
    spacing = [f'w:before="{before}"', f'w:after="{after}"']
    if line is not None and line_rule is not None:
        spacing += [f'w:line="{line}"', f'w:lineRule="{line_rule}"']
    parts.append(f'<w:spacing {" ".join(spacing)}/>')
    ind = []
    if first_line_chars is not None:
        ind.append(f'w:firstLineChars="{first_line_chars}"')
    if first_line_twips is not None:
        ind.append(f'w:firstLine="{first_line_twips}"')
    if ind:
        parts.append(f'<w:ind {" ".join(ind)}/>')
    return f"<w:p><w:pPr>{''.join(parts)}</w:pPr>{runs}</w:p>"


def chapter_title(text="第一章 绪论"):
    # 标题段落用左对齐、无首行缩进；分析器会跳过它，不拿标题格式判正文问题。
    return paragraph(run(text), align="left", first_line_chars=None, first_line_twips=None)


def section_title(text="1.1 研究背景"):
    return paragraph(run(text), align="left", first_line_chars=None, first_line_twips=None)


def end_title(text="参考文献"):
    return paragraph(run(text), align="left", first_line_chars=None, first_line_twips=None)


def body(text, o):
    return paragraph(
        run(text),
        align=o.get("align", "both"),
        line=o.get("line", EXACT20),
        line_rule=o.get("line_rule", "exact"),
        before=o.get("before", ZERO),
        after=o.get("after", ZERO),
        first_line_chars=o.get("indent_chars", 200),
        first_line_twips=o.get("indent_twips", 480))


def build_body(o):
    if o.get("no_boundary"):
        return paragraph(run("没有章标题的普通段落")) + paragraph(run("第二段"))
    parts = [
        chapter_title(),
        body("这是第一章的第一段正文，内容用于检查正文排版。", o),
        body("这是第一章的第二段正文，继续检查。", o),
    ]
    if o.get("math_exempt"):
        # 含公式段：行距故意用 auto/240（违规值），应被豁免；对齐/缩进/段距仍合规。
        parts.append(paragraph(
            run_math("含数学表达式的正文段"),
            align=o.get("align", "both"),
            line=240, line_rule="auto",
            before=ZERO, after=ZERO,
            first_line_chars=200, first_line_twips=480))
    else:
        parts.append(body("这是小节的正文段落，同样要检查。", o))
    parts.append(end_title())
    parts.append(paragraph(run("参考文献正文条目"), align="left",
                           first_line_chars=None, first_line_twips=None))
    return "".join(parts)


SAMPLES = {
    "body_compliant.docx": {},
    "body_align_left.docx": {"align": "left"},
    "body_indent_wrong.docx": {"indent_chars": 100, "indent_twips": 240},
    "body_line_spacing_wrong.docx": {"line_rule": "auto", "line": 240},
    "body_before_wrong.docx": {"before": 240},
    "body_after_wrong.docx": {"after": 240},
    "body_math_line_exempt.docx": {"math_exempt": True},
    "body_no_boundary.docx": {"no_boundary": True},
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
    entries = [
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml(build_body(SAMPLES[name]))),
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
                print(f"缺失：{name}"); problems += 1
            elif target.read_bytes() != data:
                print(f"与生成结果不一致：{name}"); problems += 1
        else:
            target.write_bytes(data)
            print(f"已生成：{name}（{len(data)} 字节）")
    if args.check:
        print("全部样本逐字节一致" if problems == 0 else f"不一致样本：{problems} 个")
        return 1 if problems else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
