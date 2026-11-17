"""R7–R10 标题层级格式检查的合成样本。

标题级别通过 w:pStyle 指向 Heading1–Heading4（styles.xml 里带 outlineLvl 与
"heading N" 名称），分析器按大纲 0/1/2/3 映射到章/一级/二级/三级。
规范值（general.yaml 4.7 表4.6）：
  章   黑体16磅(32半点)加粗  单倍行距  段前24磅(480) 段后18磅(360)
  一级 黑体14磅(28)加粗      单倍      段前24(480)   段后6(120)
  二级 黑体13磅(26)不加粗    单倍      段前12(240)   段后6(120)
  三级 黑体12磅(24)不加粗    单倍      段前12(240)   段后6(120)
对齐与序号间距按条文冲突列未判定，样本不在这些维度制造差异。

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

# 级别规范值：style / 字号(半点) / 段前 / 段后 / 是否加粗
LEVELS = {
    "章":   dict(style="Heading1", size=32, before=480, after=360, bold=True),
    "一级": dict(style="Heading2", size=28, before=480, after=120, bold=True),
    "二级": dict(style="Heading3", size=26, before=240, after=120, bold=False),
    "三级": dict(style="Heading4", size=24, before=240, after=120, bold=False),
}


def title(text, level, *, size=None, font="黑体", bold=None,
          before=None, after=None, line=240, line_rule="auto"):
    spec = LEVELS[level]
    size = spec["size"] if size is None else size
    bold = spec["bold"] if bold is None else bold
    before = spec["before"] if before is None else before
    after = spec["after"] if after is None else after
    b = "<w:b/>" if bold else ""
    run = (f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
           f'w:eastAsia="{font}"/><w:sz w:val="{size}"/>{b}</w:rPr><w:t>{text}</w:t></w:r>')
    ppr = (f'<w:pPr><w:pStyle w:val="{spec["style"]}"/><w:jc w:val="left"/>'
           f'<w:spacing w:before="{before}" w:after="{after}" '
           f'w:line="{line}" w:lineRule="{line_rule}"/></w:pPr>')
    return f"<w:p>{ppr}{run}</w:p>"


def plain(text):
    """无大纲级别的普通段落（不会被识别为标题）。"""
    run = ('<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr>'
           f'<w:t>{text}</w:t></w:r>')
    return f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr>{run}</w:p>'


def compliant_titles():
    return "".join([
        title("第1章 绪论", "章"),
        title("1.1 研究背景", "一级"),
        title("1.1.1 子标题", "二级"),
        title("1.1.1.1 细节", "三级"),
    ])


def ending():
    # 结束标题也带大纲级别，分析器据此定位正文标题检查终点（本身不判格式）。
    return title("参考文献", "章")


def build_body(name):
    if name == "headings_compliant.docx":
        return compliant_titles() + ending() + plain("参考文献正文条目")
    if name == "headings_all_wrong.docx":
        # 章标题 6 个维度全错：H1 字体 / H2 字号 / H3 加粗 / H4 行距 / H5 段前 / H6 段后
        bad_chapter = title("第1章 绪论", "章", size=24, font="宋体", bold=False,
                            before=240, after=240, line=360, line_rule="exact")
        return ("".join([
            bad_chapter,
            title("1.1 研究背景", "一级"),
            title("1.1.1 子标题", "二级"),
            title("1.1.1.1 细节", "三级"),
        ]) + ending())
    if name == "headings_size_wrong.docx":
        # 一级标题字号 16 磅（应为 14）→ 仅 H8
        return ("".join([
            title("第1章 绪论", "章"),
            title("1.1 研究背景", "一级", size=32),
            title("1.1.1 子标题", "二级"),
            title("1.1.1.1 细节", "三级"),
        ]) + ending())
    if name == "headings_mode2.docx":
        # 模式二编号，格式全部合规；验证编号模式识别，不应产生问题。
        return "".join([
            title("第一章 绪论", "章"),
            title("第一节 研究背景", "一级"),
            title("一、研究方法", "二级"),
            title("(一) 数据来源", "三级"),
            ending(),
        ])
    if name == "headings_no_boundary.docx":
        # 没有章标题：只有一级标题，章起点识别不到 → 全部未检查。
        return title("1.1 研究背景", "一级") + title("1.2 研究意义", "一级")
    raise KeyError(name)


SAMPLES = {
    "headings_compliant.docx": {},
    "headings_all_wrong.docx": {},
    "headings_size_wrong.docx": {},
    "headings_mode2.docx": {},
    "headings_no_boundary.docx": {},
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


def heading_style(style_id, name, outline):
    return (f'<w:style w:type="paragraph" w:styleId="{style_id}">'
            f'<w:name w:val="{name}"/>'
            f'<w:pPr><w:outlineLvl w:val="{outline}"/></w:pPr></w:style>')


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>'
    '<w:sz w:val="21"/>'
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    + heading_style("Heading1", "heading 1", 0)
    + heading_style("Heading2", "heading 2", 1)
    + heading_style("Heading3", "heading 3", 2)
    + heading_style("Heading4", "heading 4", 3)
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
