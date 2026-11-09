"""Issue #114 第一档：公式排版检查的合成样本。

只用标准库生成，样本全部人工构造，不含真实论文内容。
包骨架照 demo/clean_sample.docx 的写法内联，避免手写包结构出错。
ZIP 时间戳与条目顺序固定，同一 Python 版本下重复生成可逐字节核对。

用法（在本目录下）：
    python generate.py            # 重新生成全部样本
    python generate.py --check    # 只校验磁盘上的样本与重新生成结果一致，不写盘
"""
import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# 规范页边距：上下 3.8cm、左右 3.2cm（R1）。1cm ≈ 567 twips。
PAGE_W, PAGE_H = 11906, 16838
MARGIN_TB, MARGIN_LR = 2155, 1814
# 版心宽度 = 页宽 - 左右页边距。右制表位要落在这里才叫右对齐。
TEXT_WIDTH = PAGE_W - MARGIN_LR * 2  # 8278

CONTENT_TYPES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
'''

SETTINGS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/></w:settings>
'''

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>
'''

SECT_PR = (
    f'<w:sectPr><w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
    f'<w:pgMar w:top="{MARGIN_TB}" w:right="{MARGIN_LR}" w:bottom="{MARGIN_TB}" '
    f'w:left="{MARGIN_LR}" w:header="1701" w:footer="1701" w:gutter="0"/></w:sectPr>'
)


def document_xml(inner):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body>{inner}{SECT_PR}</w:body></w:document>'
    )


def plain_paragraph(text):
    """普通正文段，用来把公式段夹在中间，让段落序号有实际意义。"""
    return (
        '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
        f'<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>{text}</w:t></w:r></w:p>'
    )


def inline_paragraph(before, math_tokens, after):
    """正文里夹行内公式：公式只是句子的一部分，不该被当成独立公式段。"""
    runs = "".join(
        f'<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>{t}</m:t></m:r>'
        for t in math_tokens
    )
    return (
        '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>'
        f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{before}</w:t></w:r>'
        f'<m:oMath>{runs}</m:oMath>'
        f'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{after}</w:t></w:r>'
        "</w:p>"
    )


def omath_para(tokens, align="center", math_font="Cambria Math", with_para_pr=True):
    runs = "".join(
        f'<m:r><w:rPr><w:rFonts w:ascii="{math_font}" w:hAnsi="{math_font}"/></w:rPr><m:t>{t}</m:t></m:r>'
        for t in tokens
    )
    para_pr = f'<m:oMathParaPr><m:jc m:val="{align}"/></m:oMathParaPr>' if with_para_pr else ""
    return f'<m:oMathPara>{para_pr}<m:oMath>{runs}</m:oMath></m:oMathPara>'


def number_run(number, font="Times New Roman", size=21):
    return (
        '<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'</w:rPr><w:t>{number}</w:t></w:r>'
    )


def block_formula(
    tokens,
    number="(1)",
    align="center",
    tabs=True,
    spacing=(0, 0),
    math_font="Cambria Math",
    number_font="Times New Roman",
    number_size=21,
    with_para_pr=True,
):
    """独立成段的公式：居中 + 右制表位 + 式号，三项都能单独关掉，用来造偏差。"""
    ppr_parts = [f'<w:jc w:val="{align}"/>']
    if tabs:
        ppr_parts.append(f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_WIDTH}"/></w:tabs>')
    if spacing is not None:
        ppr_parts.append(f'<w:spacing w:before="{spacing[0]}" w:after="{spacing[1]}"/>')
    ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>"

    body = omath_para(tokens, align=align, math_font=math_font, with_para_pr=with_para_pr)
    if number is None:
        return f"<w:p>{ppr}{body}</w:p>"

    tab_run = "<w:r><w:tab/></w:r>" if tabs else ""
    return (
        f"<w:p>{ppr}{body}{tab_run}"
        + number_run(number, font=number_font, size=number_size)
        + "</w:p>"
    )


def word_native_formula(tokens, number="(1)", align="center", tabs=True):
    """Word 16 存出来的真实结构：式号抄在公式对象里，靠 m:r 里的 w:tab 推到右制表位。

    实测（Word 16.0，2026-11-09）的写法是：

        <w:pPr><w:tabs><w:tab w:val="right" w:pos="8278"/></w:tabs><w:jc w:val="center"/></w:pPr>
        <m:oMathPara><m:oMath>
          <m:r><w:rPr><w:rFonts w:ascii="Cambria Math" …/><w:sz w:val="21"/></w:rPr><m:t>E=mc^2</m:t></m:r>
          <m:r><w:rPr><w:rFonts w:ascii="Cambria Math" …/><w:sz w:val="21"/></w:rPr><w:tab/><m:t>(1)</m:t></m:r>
        </m:oMath></m:oMathPara>

    注意式号是 m:t 不是 w:t，而且它的字体是 Cambria Math——这是公式引擎写的，
    不是用户选的。只读 w:t 的实现会认不出式号，F2/F3 就永远不检查。
    """
    math_rpr = (
        '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>'
        '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    )
    body_runs = "".join(f'<m:r>{math_rpr}<m:t>{t}</m:t></m:r>' for t in tokens)
    tab = "<w:tab/>" if tabs else ""
    number_run = f'<m:r>{math_rpr}{tab}<m:t>{number}</m:t></m:r>'

    ppr_parts = []
    if tabs:
        ppr_parts.append(f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_WIDTH}"/></w:tabs>')
    if align:
        ppr_parts.append(f'<w:jc w:val="{align}"/>')
    ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>"

    return f'<w:p>{ppr}<m:oMathPara><m:oMath>{body_runs}{number_run}</m:oMath></m:oMathPara></w:p>'


def table_row_formula(tokens, number="(3)", number_align="right", formula_align="center"):
    """两列表格排版：公式一格、式号一格。论文里很常见，得能认出来。"""
    formula_cell = (
        '<w:tc><w:tcPr><w:tcW w:w="7600" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:jc w:val="{formula_align}"/></w:pPr>{omath_para(tokens, align=formula_align)}</w:p>'
        "</w:tc>"
    )
    number_cell = (
        f'<w:tc><w:tcPr><w:tcW w:w="678" w:type="dxa"/></w:tcPr>'
        f'<w:p><w:pPr><w:jc w:val="{number_align}"/></w:pPr>{number_run(number)}</w:p>'
        "</w:tc>"
    )
    return (
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
        '<w:tblGrid><w:gridCol w:w="7600"/><w:gridCol w:w="678"/></w:tblGrid>'
        f"<w:tr>{formula_cell}{number_cell}</w:tr></w:tbl>"
    )


SAMPLES = {
    # 合规样本：两种居中写法各一个（pPr/jc 与 m:oMathParaPr/jc），
    # 外加两个行内公式，确认行内公式不会被算成独立公式段。
    "formula_baseline_correct.docx": document_xml(
        plain_paragraph("式（1）给出能量与质量的关系，式（2）给出二次项展开。")
        + block_formula(["E", "=", "m", "c", "2"], number="(1)")
        + inline_paragraph("当", ["x", ">", "0"], "时上式成立，这是行内公式，不参与排版检查。")
        + block_formula(["a", "+", "b", "2"], number="(2)", with_para_pr=False)
        + inline_paragraph("取", ["n", "=", "1"], "代入即可。")
    ),
    # 偏差一：未居中，其余合规。
    "formula_not_centered.docx": document_xml(
        plain_paragraph("本段用于说明公式位置。")
        + block_formula(["E", "=", "m", "c", "2"], number="(1)", align="left")
    ),
    # 偏差二：式号没有右对齐（无右制表位，段落也是居中而非右对齐）。
    "formula_number_not_right_aligned.docx": document_xml(
        plain_paragraph("本段用于说明式号位置。")
        + block_formula(["E", "=", "m", "c", "2"], number="(1)", tabs=False)
    ),
    # 偏差三：式号写法与字体都不对——裸数字、宋体、小四。
    "formula_number_format_wrong.docx": document_xml(
        plain_paragraph("本段用于说明式号写法。")
        + block_formula(["E", "=", "m", "c", "2"], number="3", number_font="宋体", number_size=24)
    ),
    # 偏差四：公式段带段前段后间距（R11 要求 0 磅）。
    "formula_spacing_wrong.docx": document_xml(
        plain_paragraph("本段用于说明公式段间距。")
        + block_formula(["E", "=", "m", "c", "2"], number="(1)", spacing=(120, 120))
    ),
    # 表格排版：公式一格、式号一格，都合规。
    "formula_table_row_correct.docx": document_xml(
        plain_paragraph("本段用于说明两列表格排版。")
        + table_row_formula(["E", "=", "m", "c", "2"], number="(3)")
    ),
    # Word 16 原生写法：式号在公式对象内部（m:t + w:tab），全部合规。
    "formula_word_native_correct.docx": document_xml(
        plain_paragraph("本段用于说明 Word 原生公式写法。")
        + word_native_formula(["E", "=", "m", "c", "2"], number="(1)")
    ),
    # Word 原生写法但段落左对齐、没有右制表位：式号仍要能被认出来并报 F1/F2。
    # 这条是实机验证后补的——只覆盖 w:t 式号的实现会把它当成「无式号」静静放过。
    "formula_word_native_number_not_aligned.docx": document_xml(
        plain_paragraph("本段用于说明 Word 原生写法下的对齐问题。")
        + word_native_formula(["E", "=", "m", "c", "2"], number="(1)", align=None, tabs=False)
    ),
    # 只有行内公式：独立公式数应为 0，问题数应为 0。
    "formula_inline_only.docx": document_xml(
        inline_paragraph("设", ["f", "(", "x", ")"], "在区间上连续，则下列结论成立。")
        + inline_paragraph("若", ["x", ">", "0"], "，函数单调递增。")
    ),
}

FIXED_TIME = (2026, 1, 1, 0, 0, 0)


def build(sample_name):
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in [
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", SAMPLES[sample_name]),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
        ]:
            info = zipfile.ZipInfo(name, date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8"))
    return buf.getvalue()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="只比对，不写盘")
    args = parser.parse_args()

    problems = 0
    for name in SAMPLES:
        data = build(name)
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
