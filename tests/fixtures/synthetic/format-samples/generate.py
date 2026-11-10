"""Issue #5: isolated synthetic fixtures.

校验已提交样本只用 Python 标准库；重建 .docx 才需要 python-docx==1.2.0（MIT）。
这样干净检出后不必先 pip install 也能核对 10 份样本的各个维度。
"""
import argparse
import io
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

BASE = dict(cn="宋体", en="Times New Roman", size=12.0, spacing=1.5, left=3.0)
TEXT = ["完全自建合成正文第一段。", "完全自建合成正文第二段。", "合成第三段 ABC 0123456789。"]
SAMPLES = {
    "baseline_correct": {},
    "font_family_cn_wrong": {"cn": "楷体"},
    "font_ascii_wrong": {"en": "Arial"},
    "font_size_wrong": {"size": 10.5},
    "line_spacing_wrong": {"spacing": 1.0},
    "margin_wrong": {"left": 2.0},
    "style_override_control": {"inherit": True},
    "style_override_wrong": {"inherit": True, "override": True},
    "multisection_control": {"multi": True},
    "multisection_margin_wrong": {"multi": True, "second_left": 2.5},
}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC = "{http://purl.org/dc/elements/1.1/}"
DCTERMS = "{http://purl.org/dc/terms/}"
FIXED_TIME = (2000, 1, 1, 0, 0, 0)
FIXED_ISO_TIME = "2000-01-01T00:00:00Z"
HEADING = "合成格式样本"


def _require_docx():
    """重建路径才加载第三方依赖；缺失时给出可执行的提示，不抛裸 traceback。"""
    global Document, WD_SECTION, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, qn, Cm, Pt
    try:
        from docx import Document as _Document
        from docx.enum.section import WD_SECTION as _WD_SECTION
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN_PARAGRAPH
        from docx.enum.text import WD_LINE_SPACING as _WD_LINE_SPACING
        from docx.oxml.ns import qn as _qn
        from docx.shared import Cm as _Cm, Pt as _Pt
    except ImportError as exc:
        raise SystemExit(
            "重建样本需要 python-docx==1.2.0（MIT）：\n"
            "  python -m pip install -r tests/fixtures/synthetic/format-samples/requirements.txt\n"
            "只校验已提交样本不需要安装，请改用 --check-only。\n"
            "原始错误：%s" % exc
        )
    Document = _Document
    WD_SECTION = _WD_SECTION
    WD_ALIGN_PARAGRAPH = _WD_ALIGN_PARAGRAPH
    WD_LINE_SPACING = _WD_LINE_SPACING
    qn = _qn
    Cm = _Cm
    Pt = _Pt


def fonts(element, cn, en):
    element.font.name = en
    rpr = (element.element if hasattr(element, "element") else element._element).get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    for attr in list(rf.attrib):
        if attr.endswith("Theme"):
            del rf.attrib[attr]
    rf.set(qn("w:eastAsia"), cn)


def margins(section, left):
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(left)
    section.right_margin = Cm(2.5)


def build(config):
    doc = Document()
    normal = doc.styles["Normal"]
    fonts(normal, BASE["cn"], BASE["en"])
    normal.font.size = Pt(BASE["size"])
    normal.paragraph_format.line_spacing = BASE["spacing"]
    margins(doc.sections[0], config.get("left", BASE["left"]))

    def heading():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = BASE["spacing"]
        r = p.add_run(HEADING)
        fonts(r, "黑体", "黑体")
        r.font.size = Pt(16)
        r.bold = False

    def body(text, index):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = config.get("spacing", BASE["spacing"])
        r = p.add_run(text)
        if not config.get("inherit"):
            fonts(r, config.get("cn", BASE["cn"]), config.get("en", BASE["en"]))
            r.font.size = Pt(config.get("size", BASE["size"]))
        elif config.get("override") and index == 1:
            # Only this eastAsia attribute overrides Normal; size/ascii still inherit.
            rf = r._element.get_or_add_rPr().makeelement(qn("w:rFonts"), {})
            rf.set(qn("w:eastAsia"), "楷体")
            r._element.rPr.append(rf)

    heading()
    for i, text in enumerate(TEXT):
        body(text, i)
    if config.get("multi"):
        section = doc.add_section(WD_SECTION.NEW_PAGE)
        margins(section, config.get("second_left", BASE["left"]))
        heading()
        for i, text in enumerate(TEXT[:2]):
            body(text, i)
    cp = doc.core_properties
    cp.author = "Synthetic fixture generator"
    cp.last_modified_by = ""
    cp.created = cp.modified = datetime(2000, 1, 1, tzinfo=timezone.utc)
    return doc


def build_bytes(config):
    """返回该样本的完整 .docx 字节。重建路径，需要 python-docx。

    与其它生成器的 build_bytes() 同名同义，便于统一校验脚本按同一套方式调用。
    """
    _require_docx()
    stream = io.BytesIO()
    build(config).save(stream)
    buffer = io.BytesIO()
    with ZipFile(stream) as source, ZipFile(buffer, "w") as target:
        for part in sorted(source.namelist()):
            item = ZipInfo(part, date_time=FIXED_TIME)
            item.compress_type = ZIP_DEFLATED
            target.writestr(item, source.read(part))
    return buffer.getvalue()


def generate(directory):
    directory.mkdir(parents=True, exist_ok=True)
    for name, config in SAMPLES.items():
        (directory / (name + ".docx")).write_bytes(build_bytes(config))


def _attr(element, name):
    return element.get(W + name) if element is not None else None


def _twips(cm):
    return round(cm * 1440 / 2.54)


def _read(path):
    """返回 (条目名列表, {条目名: 解压字节}, ZIP 信息, 首个损坏条目名)。"""
    with ZipFile(path) as z:
        names = z.namelist()
        bad = z.testzip()
        infos = z.infolist()
        data = {name: z.read(name) for name in names}
    return names, data, infos, bad


def _normal_style(data):
    root = ET.fromstring(data["word/styles.xml"])
    for style in root.iter(W + "style"):
        if _attr(style, "styleId") == "Normal":
            return style
    return None


def _sections(body):
    """段落 pPr 里的 sectPr 是前序节，body 末尾的 sectPr 是最后一节。"""
    sections = []
    for para in body.iter(W + "p"):
        ppr = para.find(W + "pPr")
        if ppr is not None:
            sect = ppr.find(W + "sectPr")
            if sect is not None:
                sections.append(sect)
    tail = body.find(W + "sectPr")
    if tail is not None:
        sections.append(tail)
    return sections


def _paragraphs(body):
    """按 README 口径分标题与正文：空段落（分节符所在段）不计，居中段是标题。"""
    titles, bodies = [], []
    for para in body.iter(W + "p"):
        text = "".join(node.text or "" for node in para.iter(W + "t"))
        if not text:
            continue
        ppr = para.find(W + "pPr")
        jc = ppr.find(W + "jc") if ppr is not None else None
        if jc is not None and _attr(jc, "val") == "center":
            titles.append((para, text))
        else:
            bodies.append((para, text, ppr))
    return titles, bodies


def verify(directory):
    """返回问题描述列表；空列表表示全部维度通过。只用标准库，不写盘。"""
    problems = []
    baseline = directory / "baseline_correct.docx"
    if not baseline.exists():
        return ["baseline_correct.docx 缺失"]
    base_names, base_data, _, base_bad = _read(baseline)
    if base_bad is not None:
        problems.append("baseline_correct.docx 部件损坏：" + base_bad)

    for name, config in SAMPLES.items():
        label = name + ".docx"
        path = directory / label
        if not path.exists():
            problems.append(label + " 缺失")
            continue
        try:
            names, data, infos, bad = _read(path)
        except Exception as exc:
            problems.append("%s 无法读取：%s" % (label, exc))
            continue
        if bad is not None:
            problems.append("%s 部件损坏：%s" % (label, bad))
        if len(names) != len(set(names)):
            problems.append(label + " 存在重复 ZIP 条目名")
        if names != sorted(names):
            problems.append(label + " ZIP 条目未按名称排序")
        for info in infos:
            if info.date_time != FIXED_TIME:
                problems.append("%s 条目 %s 时间戳不是 %s" % (label, info.filename, FIXED_TIME))
            if info.compress_type != ZIP_DEFLATED:
                problems.append("%s 条目 %s 不是 deflate" % (label, info.filename))
        if set(names) != set(base_names):
            problems.append(label + " ZIP 部件集合与 baseline 不同")
            continue
        for part, payload in base_data.items():
            if part != "word/document.xml" and data.get(part) != payload:
                problems.append("%s 的 %s 与 baseline 不一致（只应改 document.xml）" % (label, part))

        try:
            doc_root = ET.fromstring(data["word/document.xml"])
            core = ET.fromstring(data["docProps/core.xml"])
        except ET.ParseError as exc:
            problems.append("%s XML 解析失败：%s" % (label, exc))
            continue
        body = doc_root.find(W + "body")
        if body is None:
            problems.append(label + " 缺少 w:body")
            continue

        normal = _normal_style(data)
        if normal is None:
            problems.append(label + " styles.xml 缺少 Normal 样式")
            continue
        rpr = normal.find(W + "rPr")
        rf = rpr.find(W + "rFonts") if rpr is not None else None
        sz = rpr.find(W + "sz") if rpr is not None else None
        if rf is None or sz is None:
            problems.append(label + " Normal 样式缺少 rFonts 或 sz")
            continue
        if _attr(rf, "eastAsia") != BASE["cn"] or _attr(rf, "ascii") != BASE["en"]:
            problems.append(label + " Normal 样式中西文字体不符合基线")
        if any(key.endswith("Theme") for key in rf.attrib):
            problems.append(label + " Normal 样式残留 Theme 字体属性")
        if _attr(sz, "val") != str(int(BASE["size"] * 2)):
            problems.append(label + " Normal 样式字号不是 12pt")
        base_cn, base_en = _attr(rf, "eastAsia"), _attr(rf, "ascii")

        creator = core.find(DC + "creator")
        if creator is None or creator.text != "Synthetic fixture generator":
            problems.append(label + " core.xml 作者不是固定合成标识")
        for tag in ("created", "modified"):
            node = core.find(DCTERMS + tag)
            if node is None or node.text != FIXED_ISO_TIME:
                problems.append(label + " core.xml %s 不是固定时间" % tag)
        modified_by = core.find(CP + "lastModifiedBy")
        if modified_by is None or (modified_by.text or "") != "":
            problems.append(label + " core.xml lastModifiedBy 不为空")

        multi = bool(config.get("multi"))
        sections = _sections(body)
        if len(sections) != (2 if multi else 1):
            problems.append("%s 节数应为 %d，实际 %d" % (label, 2 if multi else 1, len(sections)))
        for index, sect in enumerate(sections):
            mar = sect.find(W + "pgMar")
            if mar is None:
                problems.append("%s 第 %d 节缺少 pgMar" % (label, index + 1))
                continue
            left_cm = config.get("left", BASE["left"]) if index == 0 else config.get("second_left", BASE["left"])
            for side, cm in (("top", 2.5), ("bottom", 2.5), ("right", 2.5), ("left", left_cm)):
                raw = _attr(mar, side)
                if raw is None or abs(int(raw) - _twips(cm)) > 5:
                    problems.append("%s 第 %d 节 %s 边距不符合 %scm" % (label, index + 1, side, cm))

        titles, bodies = _paragraphs(body)
        if len(titles) != (2 if multi else 1):
            problems.append("%s 标题段数量不对" % label)
        for para, text in titles:
            if text != HEADING:
                problems.append("%s 标题文字不对：%r" % (label, text))
            hppr = para.find(W + "pPr")
            hspacing = hppr.find(W + "spacing") if hppr is not None else None
            heading_line = round(BASE["spacing"] * 240)
            if _attr(hspacing, "lineRule") != "auto" or _attr(hspacing, "line") != str(heading_line):
                problems.append("%s 标题行距不是基线 %s 倍" % (label, BASE["spacing"]))
            runs = [c for c in para if c.tag == W + "r"]
            if len(runs) != 1:
                problems.append(label + " 标题段不是单个 run")
                continue
            hrpr = runs[0].find(W + "rPr")
            hsz = hrpr.find(W + "sz") if hrpr is not None else None
            bold = hrpr.find(W + "b") if hrpr is not None else None
            if hsz is None or _attr(hsz, "val") != "32":
                problems.append(label + " 标题字号不是 16pt")
            if bold is None or _attr(bold, "val") not in ("0", "false"):
                problems.append(label + " 标题未显式设为不加粗")

        expected_texts = TEXT + (TEXT[:2] if multi else [])
        if [text for _, text, _ in bodies] != expected_texts:
            problems.append(label + " 正文文字或段落顺序不符合清单")
        want_line = round(config.get("spacing", BASE["spacing"]) * 240)
        for index, (para, _, ppr) in enumerate(bodies):
            spacing = ppr.find(W + "spacing") if ppr is not None else None
            line = _attr(spacing, "line")
            if _attr(spacing, "lineRule") != "auto" or line is None or int(line) != want_line:
                problems.append("%s 正文第 %d 段行距不是 %s 倍" % (label, index + 1, config.get("spacing", BASE["spacing"])))
            runs = [c for c in para if c.tag == W + "r"]
            if len(runs) != 1:
                problems.append("%s 正文第 %d 段不是单个 run" % (label, index + 1))
                continue
            run_rpr = runs[0].find(W + "rPr")
            inherit = bool(config.get("inherit"))
            override = bool(config.get("override")) and index == 1
            if inherit and not override and run_rpr is not None:
                problems.append("%s 继承对照第 %d 段出现直接格式" % (label, index + 1))
            if override:
                if run_rpr is None or len(list(run_rpr)) != 1:
                    problems.append(label + " 覆盖错误的第二段 rPr 应只含一个子元素")
                local = run_rpr.find(W + "rFonts") if run_rpr is not None else None
                if local is None or dict(local.attrib) != {W + "eastAsia": "楷体"}:
                    problems.append(label + " 覆盖错误应只在第二段写 eastAsia=楷体")
            local = run_rpr.find(W + "rFonts") if run_rpr is not None else None
            local_sz = run_rpr.find(W + "sz") if run_rpr is not None else None
            effective_cn = _attr(local, "eastAsia") or base_cn
            effective_en = _attr(local, "ascii") or base_en
            size_raw = _attr(local_sz, "val")
            effective_size = int(size_raw) / 2 if size_raw is not None else BASE["size"]
            want_cn = "楷体" if override else config.get("cn", BASE["cn"])
            want_en = config.get("en", BASE["en"])
            want_size = config.get("size", BASE["size"])
            if (effective_cn, effective_en, effective_size) != (want_cn, want_en, want_size):
                problems.append(
                    "%s 正文第 %d 段有效字体/字号为 %r，预期 %r"
                    % (label, index + 1, (effective_cn, effective_en, effective_size),
                       (want_cn, want_en, want_size))
                )
    return problems


def check(directory):
    problems = verify(directory)
    if problems:
        for issue in problems:
            print("不一致：" + issue)
        raise SystemExit("校验失败：%d 处" % len(problems))
    print(
        "校验通过：10 份结构样本，ZIP 完整性/排序/时间戳、Normal 样式、全部节边距、"
        "正文文本/行距/继承/唯一覆盖、标题与固定元数据（仅标准库）"
    )


def main():
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="重新生成后校验（需要 python-docx）")
    mode.add_argument("--check-only", action="store_true", help="只校验已提交样本，零依赖")
    args = parser.parse_args()
    directory = Path(__file__).resolve().parent
    if args.check_only:
        check(directory)
        return
    generate(directory)
    if args.check:
        check(directory)


if __name__ == "__main__":
    main()
