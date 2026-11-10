"""生成 demo 目录两份合成样本，只用标准库。

1. clean_sample.docx        极简样本：段落级 0 条，但「已检查 0 个正文文本片段」
2. known_issues_sample.docx 三项已知问题：东亚字体、西文字体、字号

内容为合成文本，规则 ID 用 demo/* 前缀，仅用于演示/测试，不代表合规结论。

历史情况：这两份包以前由 python-docx 生成，其中 known_issues_sample.docx 的
word/document.xml 后来被单独替换过（BOM、CRLF、缩进、时间戳都与其余部件不同），
它的 deflate 流也不是 Python zlib 写的，旧文件无法逐字节复现。本脚本按当前部件的
XML 原样重建模板，保留全部语义和风格差异，但把 ZIP 条目顺序、时间戳、压缩参数
固定下来，因此 --check 可以逐字节校验两份样本。

用法（在本目录下）：
    python generate_samples.py             # 重新生成
    python generate_samples.py --check    # 只比对，不写盘
"""
import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent
SAMPLES = ("clean_sample.docx", "known_issues_sample.docx")

# 1980-01-01 是 ZIP 时间戳下限，固定它可以避免把运行时间写进包。
FIXED_TIME = (1980, 1, 1, 0, 0, 0)
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

CONTENT_TYPES = XML_DECL + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""

ROOT_RELS = XML_DECL + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_RELS = XML_DECL + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

STYLES = XML_DECL + f"""<w:styles xmlns:w="{W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>
"""

SETTINGS = XML_DECL + f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/></w:settings>\n'

# Word 默认 A4 与 1 英寸页边距，沿用旧样本，不是校规值。
SECTPR_COMPACT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
)


def compact_rpr(east_asia, western, half_points):
    """12 磅 = w:sz 24；demo 规则按半点单位读字号。"""
    return (
        '<w:rPr><w:rFonts w:ascii="{w}" w:hAnsi="{w}" w:eastAsia="{e}"/>'
        '<w:sz w:val="{s}"/><w:szCs w:val="{s}"/></w:rPr>'
    ).format(w=western, e=east_asia, s=half_points)


def clean_document_xml():
    body = (
        f'<w:document xmlns:w="{W_NS}"><w:body>\n'
        "    <w:p>\n"
        "      <w:r>" + compact_rpr("宋体", "Times New Roman", 24)
        + "<w:t>这是一段合成演示文本，用于验证格式诊断链路。</w:t></w:r>\n"
        "    </w:p>\n"
        f"    {SECTPR_COMPACT}\n"
        "  </w:body></w:document>\n"
    )
    return XML_DECL + body


def known_issues_heading(text):
    return (
        "    <w:p>\n"
        "      <w:pPr>\n"
        '        <w:pStyle w:val="Heading1" />\n'
        "      </w:pPr>\n"
        "      <w:r>\n"
        f"        <w:t>{text}</w:t>\n"
        "      </w:r>\n"
        "    </w:p>\n"
    )


def known_issues_run(text, east_asia, western, half_points):
    return (
        "    <w:p>\n"
        "      <w:r>\n"
        "        <w:rPr>\n"
        f'          <w:rFonts w:ascii="{western}" w:hAnsi="{western}" w:eastAsia="{east_asia}" />\n'
        f'          <w:sz w:val="{half_points}" />\n'
        f'          <w:szCs w:val="{half_points}" />\n'
        "        </w:rPr>\n"
        f"        <w:t>{text}</w:t>\n"
        "      </w:r>\n"
        "    </w:p>\n"
    )


def known_issues_document_xml():
    """保留旧部件的 BOM / 小写 utf-8 / CRLF / 缩进 / 无末尾换行。

    这个部件的排版风格与其余部件不同，是历史造成的，不是规范要求；原样保留
    是为了让样本的诊断结果和文档级行为与重建前完全一致。
    """
    lines = (
        "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>\n"
        f'<w:document xmlns:w="{W_NS}">\n'
        "  <w:body>\n"
        + known_issues_heading("第一章 绪论")
        + known_issues_run("【问题1】东亚字体不符合预期。", "黑体", "Times New Roman", 24)
        + known_issues_run("【问题2】Latin 西文字体不符合预期。", "宋体", "Arial", 24)
        + known_issues_run("【问题3】字号不符合预期。", "宋体", "Times New Roman", 21)
        + known_issues_heading("参考文献")
        + "    <w:sectPr>\n"
        '      <w:pgSz w:w="11906" w:h="16838" />\n'
        '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="708" w:footer="708" w:gutter="0" />\n'
        "    </w:sectPr>\n"
        "  </w:body>\n"
        "</w:document>"
    )
    return "\ufeff" + lines.replace("\n", "\r\n")


def parts_for(name):
    """返回 (条目名, 文本) 列表；顺序也是产物的一部分，不能随意调整。"""
    shared = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
    ]
    if name == "clean_sample.docx":
        return shared + [
            ("word/document.xml", clean_document_xml()),
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
        ]
    if name == "known_issues_sample.docx":
        return shared + [
            ("word/_rels/document.xml.rels", DOCUMENT_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", known_issues_document_xml()),
        ]
    raise ValueError(f"未知样本：{name}")


def build_bytes(name):
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for entry_name, text in parts_for(name):
            info = zipfile.ZipInfo(entry_name, date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            # 显式固定，避免不同操作系统写出不同的 create_system / 权限位。
            info.create_system = 0
            info.external_attr = 0o600 << 16
            z.writestr(info, text.encode("utf-8"))
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
                print(f"校验通过：{name}（{len(data)} 字节）")
        else:
            target.write_bytes(data)
            print(f"已生成：{name}（{len(data)} 字节）")
    if args.check:
        return 1 if problems else 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
