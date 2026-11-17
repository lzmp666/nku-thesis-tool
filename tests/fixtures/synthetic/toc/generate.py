"""R6 目录格式检查的合成样本。

目录标题用文本「目录」作锚点；条目用 TOC1/TOC2/TOC3 样式（styles.xml 里带
"toc 1/2/3" 名称）让分析器分级为章/一级/二级条目。
规范值（general.yaml 4.6 表4.5）：
  目录标题 黑体16磅(32)加粗居中 单倍  段前24磅(480) 段后18磅(360)
  章条目   宋体14磅(28)          单倍  段前6磅(120)  段后0    左缩进0
  一级条目 宋体12磅(24)          单倍  段前6(120)    段后0    左缩进1汉字(leftChars=100)
  二级条目 宋体10.5磅(21)        单倍  段前6(120)    段后0    左缩进2汉字(leftChars=200)
条目的页码右对齐走制表位、依赖渲染（T6/RFC-7），样本不在该维度制造差异。

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

# 级别规范值
TITLE = dict(size=32, font="黑体", bold=True, align="center", before=480, after=360, style=None)
ENTRIES = {
    1: dict(size=28, font="宋体", style="TOC1", before=120, after=0, left_chars=0),
    2: dict(size=24, font="宋体", style="TOC2", before=120, after=0, left_chars=100),
    3: dict(size=21, font="宋体", style="TOC3", before=120, after=0, left_chars=200),
}


def para(text, *, style, size, font, bold=False, align=None,
         before, after, line=240, line_rule="auto", left_chars=None):
    b = "<w:b/>" if bold else ""
    run = (f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
           f'w:eastAsia="{font}"/><w:sz w:val="{size}"/>{b}</w:rPr><w:t>{text}</w:t></w:r>')
    parts = []
    if style:
        parts.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        parts.append(f'<w:jc w:val="{align}"/>')
    spacing = [f'w:before="{before}"', f'w:after="{after}"',
               f'w:line="{line}"', f'w:lineRule="{line_rule}"']
    parts.append(f'<w:spacing {" ".join(spacing)}/>')
    if left_chars is not None:
        parts.append(f'<w:ind w:leftChars="{left_chars}"/>')
    return f'<w:p><w:pPr>{"".join(parts)}</w:pPr>{run}</w:p>'


def toc_title(**over):
    o = dict(TITLE); o.update(over)
    return para("目录", style=o["style"], size=o["size"], font=o["font"], bold=o["bold"],
                align=o["align"], before=o["before"], after=o["after"])


def entry(level, text, **over):
    o = dict(ENTRIES[level]); o.update(over)
    return para(text, style=o["style"], size=o["size"], font=o["font"],
                before=o["before"], after=o["after"], left_chars=o["left_chars"])


def compliant_entries():
    return "".join([
        entry(1, "第一章 绪论\t1"),
        entry(2, "1.1 研究背景\t2"),
        entry(3, "1.1.1 子标题\t3"),
    ])


def build_body(name):
    if name == "toc_compliant.docx":
        return toc_title() + compliant_entries()
    if name == "toc_title_wrong.docx":
        # 目录标题 6 维偏离：C1字体/C2字号/C3加粗/C4对齐/C6段前/C7段后（行距保持合规）
        return toc_title(font="宋体", size=28, bold=False, align="left",
                         before=240, after=120) + compliant_entries()
    if name == "toc_entry_size_wrong.docx":
        # 一级条目字号 14 磅（应 12）→ 仅 C14
        return toc_title() + "".join([
            entry(1, "第一章 绪论\t1"),
            entry(2, "1.1 研究背景\t2", size=28),
            entry(3, "1.1.1 子标题\t3"),
        ])
    if name == "toc_entry_indent_wrong.docx":
        # 二级条目缩进 1 汉字（应 2）→ 仅 C24
        return toc_title() + "".join([
            entry(1, "第一章 绪论\t1"),
            entry(2, "1.1 研究背景\t2"),
            entry(3, "1.1.1 子标题\t3", left_chars=100),
        ])
    if name == "toc_field_unexpanded.docx":
        # 只有 TOC 域指令、域结果未展开（无 TOC 样式条目）→ 条目列未判定，不假装检查过。
        field = ('<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
                 '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
        return toc_title() + field
    if name == "toc_no_toc.docx":
        # 没有目录标题也没有 TOC 条目 → 全部未检查，不误报合规。
        return ('<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
                '<w:r><w:t>这是正文段落，不是目录。</w:t></w:r></w:p>')
    raise KeyError(name)


SAMPLES = {
    "toc_compliant.docx": {},
    "toc_title_wrong.docx": {},
    "toc_entry_size_wrong.docx": {},
    "toc_entry_indent_wrong.docx": {},
    "toc_field_unexpanded.docx": {},
    "toc_no_toc.docx": {},
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


def toc_style(style_id, level):
    return (f'<w:style w:type="paragraph" w:styleId="{style_id}">'
            f'<w:name w:val="toc {level}"/>'
            f'<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>')


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>'
    '<w:sz w:val="21"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    + toc_style("TOC1", 1)
    + toc_style("TOC2", 2)
    + toc_style("TOC3", 3)
    + '</w:styles>')
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
        ("word/document.xml", document_xml(build_body(name))),
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
