"""R12 图题/表题格式检查的合成样本。

题注用 FigureTableParagraphs.Classify 的文本锚点识别（图/表/续表 + 编号），
不依赖图片/表格对象。规范值（general.yaml 4.7 表4.6）：
  图题（图下）宋体10.5磅(21)居中 单倍  段前6磅(120) 段后12磅(240)
  表题（表上）宋体10.5磅(21)居中 单倍  段前6磅(120) 段后6磅(120)
题注相对对象位置、序题间距（RFC-3）、R13 表达式编号（渲染依赖）样本不制造差异。

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

SIZE = 21       # 10.5 磅（五号）
BEFORE = 120    # 段前 6 磅
FIG_AFTER = 240  # 图题段后 12 磅
TBL_AFTER = 120  # 表题段后 6 磅


def caption(text, *, after, size=SIZE, font="宋体", align="center",
            before=BEFORE, line=240, line_rule="auto"):
    run = (f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
           f'w:eastAsia="{font}"/><w:sz w:val="{size}"/></w:rPr><w:t>{text}</w:t></w:r>')
    ppr = (f'<w:pPr><w:jc w:val="{align}"/>'
           f'<w:spacing w:before="{before}" w:after="{after}" '
           f'w:line="{line}" w:lineRule="{line_rule}"/></w:pPr>')
    return f"<w:p>{ppr}{run}</w:p>"


def figure(text="图 1-1 系统总体架构", **over):
    o = dict(after=FIG_AFTER); o.update(over)
    return caption(text, **o)


def table(text="表 2-3 实验结果对比", **over):
    o = dict(after=TBL_AFTER); o.update(over)
    return caption(text, **o)


def plain(text):
    run = ('<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr>'
           f'<w:t>{text}</w:t></w:r>')
    return f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr>{run}</w:p>'


def build_body(name):
    if name == "captions_compliant.docx":
        return figure() + plain("图下方说明正文。") + table() + plain("表格上方说明正文。")
    if name == "captions_figure_wrong.docx":
        # 图题 6 维全错：K1字体/K2字号/K3对齐/K4行距/K5段前/K6段后
        return figure(font="黑体", size=28, align="left", before=240,
                      after=120, line=360, line_rule="exact") + plain("正文。")
    if name == "captions_table_after_wrong.docx":
        # 表题段后 12 磅（应 6）→ 仅 K12
        return figure() + table(after=240) + plain("正文。")
    if name == "captions_unrecognized.docx":
        # 题注编号无法识别（“图 说明”无编号），Classify 判不出 → 无题注，列未检查。
        return caption("图 说明文字没有编号", after=FIG_AFTER) + plain("正文。")
    if name == "captions_none.docx":
        return plain("这是一段普通正文，没有任何题注。") + plain("第二段正文。")
    raise KeyError(name)


SAMPLES = {
    "captions_compliant.docx": {},
    "captions_figure_wrong.docx": {},
    "captions_table_after_wrong.docx": {},
    "captions_unrecognized.docx": {},
    "captions_none.docx": {},
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
