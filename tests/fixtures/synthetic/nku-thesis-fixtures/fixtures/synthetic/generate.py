# -*- coding: utf-8 -*-
"""
nku-thesis-tool 合成格式样本生成脚本（issue #5）

用途：
    生成 tests/fixtures/synthetic/ 下全部合成 docx 样本，用于格式检查器的
    正确 / 错误 用例测试。

设计原则：
    1. 每个错误样本只引入【一个】可说明的格式错误，其余属性保持基线正确；
    2. 全部文本为合成占位文本，不复制任何真实论文内容；
    3. 样本属性可通过脚本参数复现（可重复生成，逐字节稳定，除时间戳外）。

基线格式约定（结构样本阶段假设值，最终数值以规范审核结果为准，
见 README.md「基线格式约定」一节）：
    - 中文正文字体：宋体；西文字体：Times New Roman
    - 正文字号：小四（12pt）
    - 一级标题：黑体，三号（16pt），居中
    - 正文行距：1.5 倍
    - 页边距：上 2.5cm / 下 2.5cm / 左 3.0cm / 右 2.5cm

运行：
    python generate.py            # 在脚本所在目录生成全部样本
    python generate.py --check    # 生成后重新读取并校验各样本的关键属性
"""

import argparse
import os
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# ---------------------------------------------------------------------------
# 基线格式约定（假设值，待规范审核确认后统一调整）
# ---------------------------------------------------------------------------
BASE = {
    "cn_font": "宋体",
    "en_font": "Times New Roman",
    "body_size_pt": 12.0,        # 小四
    "heading_font": "黑体",
    "heading_size_pt": 16.0,     # 三号
    "line_spacing": 1.5,
    "margin_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
}

SYNTH_PARAS = [
    "本段为合成测试文本，仅用于格式检查工具的开发与回归验证，不包含任何真实论文内容。"
    "样本中的所有文字均为占位性质，可直接替换或删除。",
    "合成文本第二段：格式检查器应当能够识别字体、字号、行距与页边距等版式属性，"
    "并在属性偏离基线约定时给出可定位的提示。本段与上一段使用相同段落样式。",
    "合成文本第三段：The quick brown fox jumps over the lazy dog. "
    "0123456789. 本句用于覆盖西文字符与数字的字体渲染场景。",
]


# ---------------------------------------------------------------------------
# 底层工具
# ---------------------------------------------------------------------------
def set_run_font(run, cn=BASE["cn_font"], en=BASE["en_font"], size=None):
    """同时设置 run 的西文字体与中文字体（eastAsia）。"""
    run.font.name = en  # 设置 ascii / hAnsi
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cn)
    if size is not None:
        run.font.size = Pt(size)


def set_style_font(style, cn, en, size_pt):
    """在样式（如 Normal）级别设置中西文字体与字号。"""
    style.font.name = en
    style.font.size = Pt(size_pt)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cn)


def set_section_margins(section, cm=None):
    cm = cm or BASE["margin_cm"]
    section.top_margin = Cm(cm["top"])
    section.bottom_margin = Cm(cm["bottom"])
    section.left_margin = Cm(cm["left"])
    section.right_margin = Cm(cm["right"])


def set_line_spacing(paragraph, multiple):
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = multiple


def configure_baseline(doc, body_cn=BASE["cn_font"], body_en=BASE["en_font"],
                       body_size=BASE["body_size_pt"],
                       line_spacing=BASE["line_spacing"]):
    """配置 Normal 样式与默认节为基线正确状态。返回 (normal_style, doc)。"""
    normal = doc.styles["Normal"]
    set_style_font(normal, body_cn, body_en, body_size)
    for p in doc.paragraphs:
        set_line_spacing(p, line_spacing)
    if doc.sections:
        set_section_margins(doc.sections[0])
    return normal


def add_heading_cn(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(p, BASE["line_spacing"])
    run = p.add_run(text)
    set_run_font(run, cn=BASE["heading_font"], en=BASE["heading_font"],
                 size=BASE["heading_size_pt"])
    run.bold = False  # 标题字重由字体与字号体现，不叠加加粗
    return p


def add_body_paragraph(doc, text, cn=None, en=None, size=None,
                       line_spacing=None):
    """添加正文段。cn/en/size/line_spacing 为 None 时继承 Normal 样式。"""
    p = doc.add_paragraph()
    set_line_spacing(p, line_spacing if line_spacing is not None
                     else BASE["line_spacing"])
    run = p.add_run(text)
    if cn is not None or en is not None or size is not None:
        set_run_font(run,
                     cn=cn or BASE["cn_font"],
                     en=en or BASE["en_font"],
                     size=size or BASE["body_size_pt"])
    else:
        # 显式写入与样式一致的 run 级字体和字号，
        # 避免不同 Word 版本的主题字体/默认字号干扰，保证样本确定性
        set_run_font(run, size=BASE["body_size_pt"])
    return p


# ---------------------------------------------------------------------------
# 样本构造器（每个函数对应一个 docx）
# ---------------------------------------------------------------------------
def build_baseline_correct(doc):
    """01 全部属性符合基线（对照组）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（正确基线）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t)


def build_font_family_cn_wrong(doc):
    """02 唯一错误：中文正文字体用了楷体（应为宋体）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（字体错误）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t, cn="楷体")  # 唯一注入的错误


def build_font_ascii_wrong(doc):
    """03 唯一错误：西文字体用了 Arial（应为 Times New Roman）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（西文字体错误）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t, en="Arial")  # 唯一注入的错误


def build_font_size_wrong(doc):
    """04 唯一错误：正文字号为五号 10.5pt（应为小四 12pt）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（字号错误）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t, size=10.5)  # 唯一注入的错误


def build_line_spacing_wrong(doc):
    """05 唯一错误：正文行距为单倍（应为 1.5 倍）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（行距错误）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t, line_spacing=1.0)  # 唯一注入的错误


def build_margin_wrong(doc):
    """06 唯一错误：左边距 2.0cm（应为 3.0cm）。"""
    configure_baseline(doc)
    margins = dict(BASE["margin_cm"], left=2.0)
    set_section_margins(doc.sections[0], margins)  # 唯一注入的错误
    add_heading_cn(doc, "第一章 合成样本标题（页边距错误）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t)


def build_style_override_control(doc):
    """07 样式覆盖对照组：Normal 正确，run 不带任何偏离的直接格式。"""
    normal = configure_baseline(doc)
    set_style_font(normal, BASE["cn_font"], BASE["en_font"],
                   BASE["body_size_pt"])
    add_heading_cn(doc, "第一章 合成样本标题（无直接覆盖）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t)
    # run 级字体与样式一致，检测器不应报告任何错误


def build_style_override_wrong(doc):
    """08 样式覆盖错误：Normal 样式正确，但第 2 段 run 直接格式改楷体。"""
    normal = configure_baseline(doc)
    set_style_font(normal, BASE["cn_font"], BASE["en_font"],
                   BASE["body_size_pt"])
    add_heading_cn(doc, "第一章 合成样本标题（样式被直接格式覆盖）")
    add_body_paragraph(doc, SYNTH_PARAS[0])
    # 唯一注入的错误：直接格式覆盖样式字体
    add_body_paragraph(doc, SYNTH_PARAS[1], cn="楷体")
    add_body_paragraph(doc, SYNTH_PARAS[2])


def build_multisection_control(doc):
    """09 多节对照：第 1、2 节页边距与正文格式均正确。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（多节正确）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t)
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    set_section_margins(sec2)  # 与基线一致
    add_heading_cn(doc, "第二章 合成样本标题（第二节，正确）")
    for t in SYNTH_PARAS[:2]:
        add_body_paragraph(doc, t)


def build_multisection_margin_wrong(doc):
    """10 多节错误：第 1 节正确；第 2 节唯一错误——左边距 2.5cm（应 3.0cm）。"""
    configure_baseline(doc)
    add_heading_cn(doc, "第一章 合成样本标题（第一节，正确）")
    for t in SYNTH_PARAS:
        add_body_paragraph(doc, t)
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    set_section_margins(sec2, dict(BASE["margin_cm"], left=2.5))  # 唯一错误
    add_heading_cn(doc, "第二章 合成样本标题（第二节页边距错误）")
    for t in SYNTH_PARAS[:2]:
        add_body_paragraph(doc, t)


# 样本注册表：文件名 -> (构造器, 场景说明, 注入的错误, 人工预期)
SAMPLES = [
    ("baseline_correct.docx", build_baseline_correct,
     "全部属性符合基线的对照组",
     "无",
     "检查器不报告任何格式问题"),
    ("font_family_cn_wrong.docx", build_font_family_cn_wrong,
     "中文正文字体错误",
     "中文字体：宋体 → 楷体（仅此一项）",
     "检查器报告正文中文字体为楷体；Word 中选中正文可见字体为楷体"),
    ("font_ascii_wrong.docx", build_font_ascii_wrong,
     "西文字体错误",
     "西文字体：Times New Roman → Arial（仅此一项）",
     "检查器报告正文西文字体为 Arial；Word 中西文与数字显示为 Arial"),
    ("font_size_wrong.docx", build_font_size_wrong,
     "正文字号错误",
     "字号：小四 12pt → 五号 10.5pt（仅此一项）",
     "检查器报告正文字号为 10.5pt；Word 中字号框显示 10.5"),
    ("line_spacing_wrong.docx", build_line_spacing_wrong,
     "正文行距错误",
     "行距：1.5 倍 → 单倍（仅此一项）",
     "检查器报告行距为单倍；Word 段落对话框行距显示单倍行距"),
    ("margin_wrong.docx", build_margin_wrong,
     "页边距错误",
     "左边距：3.0cm → 2.0cm（仅此一项）",
     "检查器报告第 1 节左边距 2.0cm；Word 页面设置可见左 2.0cm"),
    ("style_override_control.docx", build_style_override_control,
     "样式统一、无直接覆盖（样式覆盖对照组）",
     "无",
     "检查器不报告错误；确认检测逻辑基于有效字体（含样式继承）而非误报"),
    ("style_override_wrong.docx", build_style_override_wrong,
     "样式正确但被直接格式覆盖",
     "第 2 段 run 直接格式将中文字体改为楷体（仅此一项，样式本身仍为宋体）",
     "检查器报告第 2 段中文字体为楷体；Word 中修改 Normal 样式该段外观不变"),
    ("multisection_control.docx", build_multisection_control,
     "多节文档，两节均正确",
     "无",
     "检查器不报告错误；确认检测器遍历所有节"),
    ("multisection_margin_wrong.docx", build_multisection_margin_wrong,
     "多节文档，仅第 2 节页边距错误",
     "第 2 节左边距：3.0cm → 2.5cm（仅此一项，第 1 节正确）",
     "检查器仅报告第 2 节左边距 2.5cm，第 1 节不报错"),
]


def generate(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    written = []
    for name, builder, _, _, _ in SAMPLES:
        doc = Document()
        builder(doc)
        path = os.path.join(out_dir, name)
        doc.save(path)
        written.append(path)
    return written


# ---------------------------------------------------------------------------
# --check：重新读取样本并校验关键属性是否与设计一致
# ---------------------------------------------------------------------------
def _run_fonts(run):
    rpr = run._element.find(qn("w:rPr"))
    cn = None
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            cn = rfonts.get(qn("w:eastAsia"))
    return cn, run.font.name, run.font.size


def check(out_dir):
    failures = []

    def expect(cond, msg):
        if not cond:
            failures.append(msg)

    def body_runs(doc, skip_heading=True):
        runs = []
        for p in doc.paragraphs:
            if not p.runs:
                continue
            if skip_heading and p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                continue
            runs.extend(p.runs)
        return runs

    def margins(sec):
        return {k: round(getattr(sec, f"{k}_margin").cm, 2)
                for k in ("top", "bottom", "left", "right")}

    d = lambda n: os.path.join(out_dir, n)

    # 01 基线
    doc = Document(d("baseline_correct.docx"))
    for r in body_runs(doc):
        cn, en, size = _run_fonts(r)
        expect(cn == "宋体", f"baseline: 中文={cn}")
        expect(en == "Times New Roman", f"baseline: 西文={en}")
        expect(size and size.pt == 12.0, f"baseline: 字号={size}")
    for p in doc.paragraphs:
        expect(p.paragraph_format.line_spacing == 1.5,
               f"baseline: 行距={p.paragraph_format.line_spacing}")
    expect(margins(doc.sections[0]) == BASE["margin_cm"],
           f"baseline: 边距={margins(doc.sections[0])}")

    # 02 中文字体错误
    doc = Document(d("font_family_cn_wrong.docx"))
    for r in body_runs(doc):
        cn, en, size = _run_fonts(r)
        expect(cn == "楷体", f"font_cn: 中文={cn}")
        expect(en == "Times New Roman", f"font_cn: 西文={en}")
        expect(size and size.pt == 12.0, f"font_cn: 字号={size}")
    expect(margins(doc.sections[0]) == BASE["margin_cm"], "font_cn: 边距意外改变")

    # 03 西文字体错误
    doc = Document(d("font_ascii_wrong.docx"))
    for r in body_runs(doc):
        cn, en, size = _run_fonts(r)
        expect(cn == "宋体", f"font_en: 中文={cn}")
        expect(en == "Arial", f"font_en: 西文={en}")
        expect(size and size.pt == 12.0, f"font_en: 字号={size}")

    # 04 字号错误
    doc = Document(d("font_size_wrong.docx"))
    for r in body_runs(doc):
        cn, en, size = _run_fonts(r)
        expect(cn == "宋体" and en == "Times New Roman", "font_size: 字体意外改变")
        expect(size and size.pt == 10.5, f"font_size: 字号={size}")

    # 05 行距错误
    doc = Document(d("line_spacing_wrong.docx"))
    body = [p for p in doc.paragraphs
            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER and p.runs]
    for p in body:
        expect(p.paragraph_format.line_spacing == 1.0,
               f"spacing: 行距={p.paragraph_format.line_spacing}")
    expect(margins(doc.sections[0]) == BASE["margin_cm"], "spacing: 边距意外改变")

    # 06 页边距错误
    doc = Document(d("margin_wrong.docx"))
    expect(margins(doc.sections[0])["left"] == 2.0,
           f"margin: 左={margins(doc.sections[0])}")
    for r in body_runs(doc):
        expect(_run_fonts(r)[0] == "宋体", "margin: 字体意外改变")

    # 07 样式覆盖对照
    doc = Document(d("style_override_control.docx"))
    for r in body_runs(doc):
        expect(_run_fonts(r)[0] == "宋体", "override_ctrl: 中文意外偏离")

    # 08 样式覆盖错误：仅第 2 个正文段为楷体
    doc = Document(d("style_override_wrong.docx"))
    paras = [p for p in doc.paragraphs
             if p.alignment != WD_ALIGN_PARAGRAPH.CENTER and p.runs]
    expect(len(paras) == 3, f"override: 段数={len(paras)}")
    fonts = [_run_fonts(r)[0] for p in paras for r in p.runs]
    expect(fonts == ["宋体", "楷体", "宋体"], f"override: 序列={fonts}")

    # 09 多节对照
    doc = Document(d("multisection_control.docx"))
    expect(len(doc.sections) == 2, f"multi_ctrl: 节数={len(doc.sections)}")
    for i, sec in enumerate(doc.sections):
        expect(margins(sec) == BASE["margin_cm"],
               f"multi_ctrl: 第{i+1}节边距={margins(sec)}")

    # 10 多节错误：仅第 2 节左边距错误
    doc = Document(d("multisection_margin_wrong.docx"))
    expect(len(doc.sections) == 2, f"multi: 节数={len(doc.sections)}")
    expect(margins(doc.sections[0]) == BASE["margin_cm"],
           f"multi: 第1节边距={margins(doc.sections[0])}")
    expect(margins(doc.sections[1])["left"] == 2.5,
           f"multi: 第2节左边距={margins(doc.sections[1])['left']}")

    return failures


def main():
    parser = argparse.ArgumentParser(description="生成/校验合成 docx 样本")
    parser.add_argument("--check", action="store_true",
                        help="生成后重新读取并校验关键属性")
    args = parser.parse_args()

    out_dir = os.path.dirname(os.path.abspath(__file__))
    written = generate(out_dir)
    print(f"已生成 {len(written)} 个样本：")
    for path in written:
        print(f"  - {os.path.basename(path)}")

    if args.check:
        failures = check(out_dir)
        if failures:
            print(f"\n校验失败 {len(failures)} 项：")
            for f in failures:
                print(f"  [FAIL] {f}")
            sys.exit(1)
        print("\n校验通过：全部样本关键属性与设计一致。")


if __name__ == "__main__":
    main()
