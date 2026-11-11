"""一份真正走完两级检查链路的干净合成样本。

背景：demo/clean_sample.docx 段落级输出「已检查 0 个正文文本片段」，
文档级又因为沿用 Word 默认页面设置报 8 处。它的 0 条来自「压根没查」，
不能拿来当「检查器工作正常」的证据。它保留下来作反面教材。

本脚本生成 clean_checked_sample.docx：
  - 段落级：可识别的章节边界 + 三段正文，字体字号全部按 demo 规则正确设置，
    所以报告会写明「已检查 3 个正文文本片段」且诊断 0。
  - 文档级：A4、R1 四边距与页眉页脚距、装订线 0、有篇眉、页脚是 PAGE 域，
    篇眉页码的字体字号与居中都能完整判定，所以 D1–D12 全过。

包骨架沿用 tests/fixtures/synthetic/document/generate.py 的写法，
只用标准库。ZIP 条目顺序与时间戳固定，同一 Python 版本下重复生成可逐字节核对。

用法（在本目录下）：
    python generate_checked_sample.py            # 重新生成
    python generate_checked_sample.py --check    # 只比对，不写盘
"""
import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent
# 与其它生成器同形：清单常量 + 接受样本名的 build_bytes()。
SAMPLES = ("clean_checked_sample.docx",)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 与 document/generate.py 保持同一组规范值：A4、上下 3.8cm、左右 3.2cm、
# 页眉页脚 3.0cm、装订线 0。规范值来自 rules/nku-graduate/general.yaml 的 R1。
A4 = (11906, 16838)
SPEC_MARGINS = dict(top=2155, right=1814, bottom=2155, left=1814)
SPEC_HF = dict(header=1701, footer=1701)

FIXED_TIME = (2026, 1, 1, 0, 0, 0)

# 10.5 磅 = sz 21，用于篇眉与页码（R2 / R3）。
PART_RPR = (
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
)
# 12 磅 = sz 24，是 demo/* 规则对正文的期望值。
BODY_RPR = (
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
    '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
)

BODY_TEXTS = [
    "本章为合成演示正文第一段，用于验证字体与字号检查链路。Latin Text 一并覆盖。",
    "第二段是合成文本，Times New Roman 与宋体都按演示规则显式设置。",
    "第三段用于确认正文边界结束前的片段也计入已检查范围。",
]


def heading(text):
    """章节边界：styleId 名字即可识别为标题，不需要 styles.xml 里再定义。"""
    return (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        f"<w:r><w:t>{text}</w:t></w:r></w:p>"
    )


def body_paragraph(text):
    return f"<w:p><w:r>{BODY_RPR}<w:t>{text}</w:t></w:r></w:p>"


def document_xml():
    body = heading("第一章 绪论")
    body += "".join(body_paragraph(text) for text in BODY_TEXTS)
    body += heading("参考文献")
    body += (
        "<w:sectPr>"
        '<w:headerReference w:type="default" r:id="rId10"/>'
        '<w:footerReference w:type="default" r:id="rId11"/>'
        f'<w:pgSz w:w="{A4[0]}" w:h="{A4[1]}"/>'
        f'<w:pgMar w:top="{SPEC_MARGINS["top"]}" w:right="{SPEC_MARGINS["right"]}" '
        f'w:bottom="{SPEC_MARGINS["bottom"]}" w:left="{SPEC_MARGINS["left"]}" '
        f'w:header="{SPEC_HF["header"]}" w:footer="{SPEC_HF["footer"]}" w:gutter="0"/>'
        "</w:sectPr>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body}</w:body></w:document>'
    )


def header_xml():
    """R2：篇眉宋体 10.5 磅居中，内容与该部分标题相同。"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:hdr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f"<w:r>{PART_RPR}<w:t>第一章 绪论</w:t></w:r></w:p></w:hdr>"
    )


def footer_xml():
    """R3：页码用 PAGE 域，显示结果那个 run 上的 rPr 决定字体字号。"""
    runs = (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"<w:r>{PART_RPR}<w:t>1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:ftr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f"{runs}</w:p></w:ftr>"
    )


def styles_xml():
    """样式链可用：docDefaults + Normal。PageFormatReader 需要它才能标「完整判定」。"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
        "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>'
        "</w:styles>"
    )


def settings_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/></w:settings>'
    )


def content_types():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        "</Types>"
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)


def document_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        "</Relationships>"
    )


def build_bytes(name=SAMPLES[0]):
    """返回本样本的 docx 字节。未知样本名直接报错，不静默回落到唯一那份。"""
    if name != SAMPLES[0]:
        raise KeyError("未知样本：%s（只有 %s）" % (name, SAMPLES[0]))
    entries = [
        ("[Content_Types].xml", content_types()),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", document_xml()),
        ("word/_rels/document.xml.rels", document_rels()),
        ("word/settings.xml", settings_xml()),
        ("word/styles.xml", styles_xml()),
        ("word/header1.xml", header_xml()),
        ("word/footer1.xml", footer_xml()),
    ]
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries:
            info = zipfile.ZipInfo(name, date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8"))
    return buf.getvalue()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="只比对，不写盘")
    args = parser.parse_args()

    failed = 0
    for name in SAMPLES:
        data = build_bytes(name)
        target = HERE / name
        if args.check:
            if not target.exists():
                print(f"缺失：{name}")
                failed += 1
                continue
            if target.read_bytes() != data:
                print(f"与生成结果不一致：{name}")
                failed += 1
                continue
            print(f"校验通过：{name}（{len(data)} 字节）")
        else:
            target.write_bytes(data)
            print(f"已生成：{name}（{len(data)} 字节）")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
