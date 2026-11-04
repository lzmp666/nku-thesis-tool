"""Generate a long, fully synthetic thesis fixture for Word add-in acceptance."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "synthetic_complete_thesis.docx"
CN_FONT = "宋体"
LATIN_FONT = "Times New Roman"


def set_run_font(run, east_asia: str = CN_FONT, western: str = LATIN_FONT, size: float = 12,
                 bold: bool | None = None, italic: bool | None = None) -> None:
    run.font.name = western
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    r_pr = run._element.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, fonts)
    fonts.set(qn("w:ascii"), western)
    fonts.set(qn("w:hAnsi"), western)
    fonts.set(qn("w:eastAsia"), east_asia)


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:fill"), fill)
    tc_pr.insert_element_before(shading, "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:tcPrChange")


def set_cell_borders(cell, color="9AA9BA", size="5") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)
        borders.append(tag)
    tc_pr.insert_element_before(borders, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:tcPrChange")


def set_para_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent=Cm(0.74),
                    before=0, after=0, line=1.5) -> None:
    paragraph.alignment = alignment
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = first_indent
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.widow_control = True


def add_text(paragraph, text: str, *, east_asia=CN_FONT, western=LATIN_FONT, size=12,
             bold=None, italic=None):
    run = paragraph.add_run(text)
    set_run_font(run, east_asia, western, size, bold, italic)
    return run


def add_body(doc: Document, text: str, marker: str | None = None) -> None:
    p = doc.add_paragraph()
    set_para_format(p)
    if marker == "wrong_cn":
        before, after = text.split("【中文格式测试点】", 1)
        add_text(p, before)
        add_text(p, "【中文格式测试点】", east_asia="黑体")
        add_text(p, after)
    elif marker == "wrong_en":
        before, after = text.split("SYNTHETIC ENGLISH FORMAT CHECK", 1)
        add_text(p, before)
        add_text(p, "SYNTHETIC ENGLISH FORMAT CHECK", western="Arial")
        add_text(p, after)
    elif marker == "wrong_size":
        before, after = text.split("【字号格式测试点】", 1)
        add_text(p, before)
        add_text(p, "【字号格式测试点】", size=10.5)
        add_text(p, after)
    else:
        add_text(p, text)


def add_heading(doc: Document, text: str, level=1) -> None:
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(14 if level == 1 else 9)
    p.paragraph_format.space_after = Pt(6)
    add_text(p, text, east_asia="黑体", western=LATIN_FONT,
             size=16 if level == 1 else 13, bold=True)


def add_caption(doc: Document, text: str, *, wrong_style: bool = False) -> None:
    p = doc.add_paragraph()
    is_figure = text.startswith("图")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if wrong_style else WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0 if wrong_style else 6)
    p.paragraph_format.space_after = Pt(0 if wrong_style else (12 if is_figure else 6))
    p.paragraph_format.line_spacing = 1.5 if wrong_style else 1.0
    add_text(p, text, east_asia="黑体" if wrong_style else "宋体", size=12 if wrong_style else 10.5)


def add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for col, text in enumerate(headers):
        cell = table.rows[0].cells[col]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, "DCE8F5")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            add_text(p, text, size=10.5, bold=True)
    for row in rows:
        cells = table.add_row().cells
        for col, text in enumerate(row):
            cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cells[col].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col == 0 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_after = Pt(2)
                add_text(p, text, size=10.5)
    for row in table.rows:
        for cell in row.cells:
            set_cell_borders(cell)


def make_figure(title: str, labels: list[str], values: list[int]) -> BytesIO:
    width, height = 1200, 620
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font_path = Path("C:/Windows/Fonts/msyh.ttc")
    font = ImageFont.truetype(str(font_path), 30) if font_path.exists() else ImageFont.load_default()
    small = ImageFont.truetype(str(font_path), 23) if font_path.exists() else ImageFont.load_default()
    draw.text((60, 30), title, fill="#18324F", font=font)
    left, top, chart_w, chart_h = 120, 120, 980, 400
    draw.line((left, top, left, top + chart_h), fill="#60758C", width=3)
    draw.line((left, top + chart_h, left + chart_w, top + chart_h), fill="#60758C", width=3)
    max_value = max(values)
    gap = 46
    bar_w = (chart_w - gap * (len(values) + 1)) // len(values)
    colors = ["#3577B8", "#54A6A1", "#E6A344", "#7C70B7", "#D56F62"]
    for i, (label, value) in enumerate(zip(labels, values)):
        x = left + gap + i * (bar_w + gap)
        bar_h = int((chart_h - 60) * value / max_value)
        draw.rounded_rectangle((x, top + chart_h - bar_h, x + bar_w, top + chart_h), radius=10,
                               fill=colors[i % len(colors)])
        draw.text((x + bar_w // 3, top + chart_h - bar_h - 34), str(value), fill="#26384A", font=small)
        draw.text((x + 5, top + chart_h + 12), label, fill="#35485D", font=small)
    draw.text((left, height - 40), "说明：所有数值为本研究合成演示数据。", fill="#667789", font=small)
    stream = BytesIO()
    image.save(stream, format="PNG", dpi=(160, 160))
    stream.seek(0)
    return stream


def add_page_number(section) -> None:
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Cm(0)
    add_text(footer, "— ", size=9)
    run = footer.add_run()
    set_run_font(run, size=9)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    add_text(footer, " —", size=9)


CHAPTERS = [
    ("第1章 绪论", [
        ("1.1 研究背景", [
            "高校图书馆在学期中承担课程学习、资料检索、协作讨论和短时休息等多种功能。随着移动设备普及，学生可以在教学楼、宿舍和图书馆之间灵活安排学习活动，但空间选择仍受座位供给、课程时间、设备条件和同行协作需求影响。管理者通常根据年度统计或人工巡查调整开放区域，难以及时发现一天之内不同区域的需求变化。本段含有【中文格式测试点】，供单项修复和悬停联动验收。",
            "本研究关注一个范围有限但具有实践意义的问题：在不收集个人身份信息的条件下，能否通过短周期、低成本的观察记录，形成可解释的空间需求判断。论文中的“北岚大学”及其全部调查记录均为虚构数据，仅用于检验论文格式分析、定位和修复流程，不描述任何真实学校或真实学生。"], "wrong_cn"),
        ("1.2 研究问题与目标", [
            "研究围绕三个问题展开。第一，座位使用的时间变化是否足以支持对开放区域进行分时管理；第二，区域类型、课程周期和设备条件如何共同影响学生的空间选择；第三，如何把观察所得转化为便于管理人员检查的规则，而不是不可解释的单一评分。研究目标是构造一套可复现的记录方案，并用合成样本验证分析流程。此处设置【字号格式测试点】，用于验证修改和再次分析。",
            "在软件测试层面，本章还设置了明确的格式对照样本：一处汉字文本采用不同于目标值的东亚字体，一处英文句子采用不同的西文字体，一处正文片段采用不同字号。读者可在 Word 中分别查看分析卡片、原文高亮、单项修复预览以及重新分析后的剩余问题。"], "wrong_size"),
        ("1.3 研究范围与论文结构", [
            "研究范围限定为图书馆公共学习空间的合成使用记录，不采集姓名、学号、设备标识和精确轨迹。分析单位为区域与时间段，结果只用于说明方法能否复现。论文不尝试从少量观察推断个体偏好，也不将空间占用率解释为学习质量或学业表现。",
            "全文共六章。第一章说明问题背景、研究目标和范围；第二章回顾空间管理与需求测量的相关思路；第三章给出变量、流程和计算方法；第四章报告合成数据上的流程验证；第五章讨论误差、隐私和适用边界；第六章总结结论并提出后续工作。"], None),
    ]),
    ("第2章 相关研究与概念框架", [
        ("2.1 学习空间需求的观察方法", [
            "公共空间需求常通过预约记录、人工计数、短问卷或设施使用日志进行描述。不同方法观察到的对象并不相同：预约反映计划使用，现场计数反映某一时点的占用，问卷反映自我报告的偏好，设施日志则可能只记录设备状态。将这些信息合并时，必须保留时间窗口和采集边界，否则同一指标可能代表完全不同的含义。",
            "已有管理实践常把容量利用率作为空间配置的起点，但单一总量无法解释座位之间的差异。例如，安静阅读区和小组讨论区即使具有相同占用率，仍可能对应不同的拥挤感和噪声容忍度。因此本研究将空间类型作为分层变量，并同时报告人数、可用座位和活动类型。"], None),
        ("2.2 需求分层与服务设计", [
            "服务设计强调将使用场景、服务对象和资源限制联系起来。对图书馆而言，管理者面对的不是抽象的“更多座位”，而是特定时段、特定区域、特定设备组合下的服务需求。需求分层可以帮助组织观察，也可避免将所有活动压缩为一个难以解释的综合分数。",
            "本研究采用“场景—时段—区域”三层框架。场景区分个人学习与协作活动，时段按课程前后和晚间划分，区域区分安静区、讨论区和数字资源区。框架的目的不是给学生贴标签，而是使管理建议可以追溯到观测条件，并明确哪些推断仍需人工核实。"], None),
        ("2.3 隐私保护与数据最小化", [
            "空间研究可能涉及敏感行为数据。若管理目的只是估计区域使用强度，通常不需要记录个体身份或连续移动路径。数据最小化原则要求只收集回答研究问题所需的信息，并在设计阶段规定保留期限、汇总粒度和访问范围。即使数据经过汇总，也应检查小样本时段是否可能使个体被间接识别。",
            "本文采用虚构的汇总数据，目的是让测试样本具有论文结构而不包含真实调查内容。实际部署时还需由学校伦理和数据管理流程审查采集方式；本论文的演示材料不能替代伦理审批、告知或真实场地授权。"], None),
        ("2.4 概念框架与研究假设", [
            "综合前述研究，本研究把区域需求表示为时段内的平均占用、达到高占用阈值的次数、活动类型比例和设备可用率。课程周期、区域类型和服务条件作为解释变量。研究假设不预设某一种区域必然更受欢迎，而是检验这些变量是否为观察到的差异提供了可复核解释。",
            "分析框架包含两个层次：描述层回答“何时、何地、出现何种活动”；解释层比较不同场景下的需求模式。管理建议只在描述结果稳定且解释边界清楚时提出，并与数据来源和不确定性一起报告。"], None),
    ]),
    ("第3章 研究设计与方法", [
        ("3.1 研究流程", [
            "研究流程分为问题界定、字段设计、观察记录、数据校验、分层汇总和建议复核六步。每一步都保留输入和输出说明，使另一名分析者能够按照相同口径重算指标。记录表在正式使用前先用虚构数据进行试填，检查字段定义是否存在相互矛盾或难以判断的情况。",
            "观察员仅记录预先定义的区域级数量和活动类别，不追踪个人。每个时间段采用相同长度的观察窗口，遇到临时闭馆、设备维护或特殊活动时，将该时段标记为不可比较，而不以零占用代替缺失值。"], None),
        ("3.2 变量操作化", [
            "区域占用率定义为观察时点在用座位数与可使用座位数之比。设备可用率定义为可正常使用设备数与计划开放设备数之比。协作活动比例则用记录为讨论或共同操作的座位数除以可分类活动总数。三项指标的分母和缺失处理规则在计算前固定，避免结果出来后调整口径。",
            "所有指标均按区域类型和时段分别汇总。若某一时段的有效观察次数不足，则只报告原始计数与缺失说明，不进行趋势外推。课程周与非课程周分开呈现，避免把日历结构变化误判为长期需求变化。"], None),
        ("3.3 数据质量控制", [
            "数据校验包括范围检查、字段完整性检查和重复记录检查。占用数不得超过可用座位数；开放时间之外的记录需要备注；区域代码必须来自固定字典。两名观察员对抽取的记录进行独立分类，分歧先依据操作手册复核，无法一致时保留“不确定”类别。",
            "记录完成后由脚本生成不可变的原始导出文件和清洗日志。清洗步骤只进行格式统一、重复项核验和缺失标记，不对个别记录进行推测性补全。论文中的流程测试进一步把正文、题注和表格单元格分开，防止将正文规则误用在对象范围不同的段落上。"], None),
        ("3.4 指标计算与比较", [
            "对区域 r 在时段 t 的占用率计算式如下。该式仅用于合成数据的演示，不对真实图书馆提供容量规划结论。",
        ], None),
    ]),
    ("第4章 合成数据结果与流程验证", [
        ("4.1 样本结构", [
            "为验证统计流程，构造了六周的合成观察记录，覆盖三个虚构区域和四种时段。每周安排相同数量的工作日观察点，另设置部分设备维护和临时关闭记录，用于测试缺失标记是否被错误解释为低需求。所有数值均由脚本生成，未从真实场地采集。",
            "样本中安静区的可用座位较多，讨论区的协作活动比例较高，数字资源区的设备可用率具有周际变化。该结构仅用于让表格和图示具有不同数据形态，不代表任何真实学校的服务现状。"], None),
        ("4.2 区域使用概况", [
            "表4-1展示样本期各类区域的汇总指标。安静区的合成平均占用率为0.72，讨论区为0.66，数字资源区为0.59。高占用时段次数与平均占用率并不完全一致，说明只观察平均值会遮蔽短时拥挤情况。",
            "表中可用率按计划开放设备计算。若设备维护导致某一时段不可用，该时段不应计入“无人使用”的观察。不同指标在样本规模较小时容易受到观察安排影响，因此结果仅作流程核对，不进行显著性或因果解释。"], None),
        ("4.3 时段变化", [
            "图4-1显示四个时段的合成平均占用水平。下午时段较高，午间和晚间次之，上午相对较低。趋势的作用是测试图像插入、题注识别和文字引用是否可以共同保存在 Word 文件中；不应用于调整真实开放时间。",
            "在实际研究中，还需要结合课程表、天气、考试周和开放区域变化判断趋势。若只选择观察方便的日期，估计值可能系统性偏高或偏低；因此后续实施应预先登记观察窗口，并保存因故缺测的说明。"], None),
        ("4.4 流程检验", [
            "合成测试覆盖了正文中文字体、英文西文字体和字号三个已实现的诊断入口，并在文档中加入多段正文、图、表、公式、参考文献和附录。SYNTHETIC ENGLISH FORMAT CHECK is an intentionally malformed English run for verifying Latin-font detection. 表格单元格与图表题注被排除在正文规则之外，主要用于检查段落索引和修复定位不会因复杂结构而静默偏移。",
            "本轮演示的验收重点是：所选段落位置与 Word 中的实际文字相同；问题卡片给出期望值和实际值；单项与批量修复预览列出将修改的段落；修复写入当前文档后重新分析，已修复位置不再出现同一诊断。对规则未覆盖的图像清晰度或表格样式，系统应显示未覆盖，而不是生成自动修复。"], "wrong_en"),
    ]),
    ("第5章 讨论", [
        ("5.1 结果解释", [
            "合成样本展示了分层记录和简单汇总的可操作性，但并未证明这些指标足以代表空间服务质量。占用率高可能意味着需求旺盛，也可能意味着替代空间不足；占用率低可能与时间安排、设备条件或区域定位有关。解释时需要把使用数据与服务目标同时纳入讨论。",
            "研究流程的价值主要在于建立透明的数据口径。管理者可以从原始观察追溯到汇总指标，也可检查维护记录或缺测情况是否影响结论。若缺少这样的过程信息，单个数字即使计算正确，也可能被赋予超出数据能力的意义。"], None),
        ("5.2 方法局限", [
            "第一，固定观察窗口不能捕捉连续变化，可能错过高峰的开始和结束。第二，活动分类仍需要观察员判断，不同人员对“讨论”和“个人协作”的边界可能理解不同。第三，样本只为格式和流程验收而合成，不可据此评估任何真实场地。",
            "此外，数据表格中的分类和值需要专业人员复核；自动统计不理解具体学科、课程安排和无障碍需求。更完整的研究应进行观察员一致性培训，开展小规模预测试，并将方案提交给相应的伦理与数据管理审查。"], None),
        ("5.3 隐私与治理", [
            "即使研究不记录姓名，摄像、无线网络日志或门禁记录仍可能形成可识别轨迹，因此本研究不建议把此类数据作为默认输入。若管理任务确需使用已有系统数据，应先论证必要性，缩短保留时间，限制访问人员，并优先输出不可逆的区域级汇总。",
            "数据治理还包括对模型建议的约束。管理者不应仅根据某个区域的平均值减少服务供给，也不应把统计标签用于评价个人。论文呈现了数据的生成和汇总方法，但实际部署仍应由场地负责部门判断影响并向使用者说明。"], None),
        ("5.4 实施建议", [
            "建议按小步试点推进：先由管理人员共同定义区域和时段，再使用不含个人标识的观察表开展短周期测试；核对字段是否容易理解后，才扩大到更多日期。每轮结果都记录数据质量、服务调整和后续观察，避免一次性统计长期固化为不适用的规则。",
            "涉及区域用途改变、设备采购或开放时间调整时，应让空间使用者和一线工作人员参与讨论。对样本不足、来源冲突或无法比较的时段，应保留不确定性说明，而不是给出看似精确的结论。"], None),
    ]),
    ("第6章 结论与展望", [
        ("6.1 主要结论", [
            "本研究构造了一套面向公共学习空间的场景—时段—区域记录框架，并说明了占用、活动和设备指标的计算边界。合成数据表明，按区域和时段组织记录有助于检查流程的一致性；多个指标并列呈现也比单一综合分数更容易解释。",
            "研究没有使用真实学生或真实学校数据，不能被引用为任何机构的现状评估。论文样例中的格式错误是有意设置的测试点，修复结果只能说明插件针对当前草案规则执行了对应格式变更，不代表论文其他部分符合学校规范。"], None),
        ("6.2 后续工作", [
            "后续首先需要由规则负责人完成标准范围与规则值审核，并为图题、表题、表格内部格式、页边距、段落间距和参考文献等内容分别定义清晰的适用范围。其次要扩充独立合成样例和人工黄金结果，再由非开发人员在受支持的 Word 版本中复核分析、定位、修复和撤销行为。",
            "在数据研究方面，未来可在获得必要审批后设计小规模观察试点；在软件方面，可探索题注专属规则和表格专属规则，但必须先明确标准语义及例外条件。对图片清晰度、图表自明性和论文论证质量，应继续保留人工判断，不把视觉模型的猜测作为格式修复依据。"], None),
    ]),
]


def main() -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.4)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)
    add_page_number(section)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN_FONT
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(0)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)
    normal._element.rPr.rFonts.set(qn("w:ascii"), LATIN_FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), LATIN_FONT)
    for style_name in ("Heading 1", "Heading 2", "Title", "Subtitle"):
        style = doc.styles[style_name]
        style.font.name = LATIN_FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        style._element.rPr.rFonts.set(qn("w:ascii"), LATIN_FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), LATIN_FONT)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(65)
    p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "海岚大学研究生学位论文（合成测试样稿）", size=18, bold=True, east_asia="黑体")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(78)
    p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "高校图书馆学习空间需求的分时观察与管理方法研究", size=22, bold=True, east_asia="黑体")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "A Synthetic Study of Time-Segmented Observation and Management for Learning Spaces in University Libraries", size=14, western=LATIN_FONT)
    for label, value in [("学科门类", "管理学（虚构）"), ("研究方向", "公共服务与信息管理"),
                         ("研究生", "合成测试数据"), ("指导教师", "合成测试数据"), ("完成日期", "2026年10月")]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(12)
        add_text(p, f"{label}：{value}")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(38)
    p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "本文件为软件测试用虚构论文，不含真实学生研究或真实调查数据。", size=10.5, bold=True)
    doc.add_page_break()

    add_heading(doc, "摘  要", 1)
    add_body(doc, "高校公共学习空间的管理需要理解不同区域、时段和活动类型之间的需求差异。本文构造一套不使用个人身份信息的分时观察框架，围绕区域占用、协作活动和设备可用情况建立记录字段，并通过合成数据演示质量检查、分层汇总和结果复核流程。研究将“场景—时段—区域”作为组织结构，分别呈现安静学习区、讨论区和数字资源区的模拟使用情况。结果显示，该结构有助于追溯指标定义和缺失处理，但平均占用率不能单独代表服务质量，管理建议仍需结合课程周期、设施条件和使用者反馈。本文所有学校名称、观察记录、统计结果与参考条目均为虚构，仅用于软件格式诊断和修复验收；不构成任何真实机构的研究结论。")
    p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "关键词：学习空间；分时观察；公共服务；数据最小化；合成测试", bold=True)

    doc.add_page_break()
    add_heading(doc, "ABSTRACT", 1)
    p = doc.add_paragraph(); set_para_format(p, first_indent=Cm(0))
    add_text(p, "University libraries provide quiet study areas, collaborative rooms, and digital-resource spaces. This synthetic thesis develops a time-segmented observation framework that records area-level occupancy, activity categories, and equipment availability without collecting personal identifiers. A fictional six-week dataset is used to demonstrate validation, stratified summaries, and review of management suggestions. The results illustrate how a scene-time-area structure makes indicator definitions and missing-data decisions traceable. They do not establish the quality of any real library service. All institutions, observations, measurements, and references in this document are fictional and are included solely for software testing and document-format acceptance.", western=LATIN_FONT)
    p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Cm(0)
    add_text(p, "Key words: learning spaces; time-segmented observation; public services; data minimization; synthetic test", western=LATIN_FONT, bold=True)

    doc.add_page_break()
    add_heading(doc, "目  录", 1)
    for item in ["第1章 绪论", "第2章 相关研究与概念框架", "第3章 研究设计与方法", "第4章 合成数据结果与流程验证", "第5章 讨论", "第6章 结论与展望", "参考文献", "附录A 观察记录字段", "附录B 合成样本说明", "致谢"]:
        p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Cm(0); p.paragraph_format.space_after = Pt(8)
        add_text(p, item)
    doc.add_page_break()

    for chapter_index, (chapter_title, sections) in enumerate(CHAPTERS):
        add_heading(doc, chapter_title, 1)
        for section_title, paragraphs, marker in sections:
            add_heading(doc, section_title, 2)
            for paragraph_index, text in enumerate(paragraphs):
                selected_marker = marker if paragraph_index == 0 else None
                add_body(doc, text, selected_marker)
            if chapter_title.startswith("第3章") and section_title.startswith("3.4"):
                equation = doc.add_paragraph()
                equation.alignment = WD_ALIGN_PARAGRAPH.CENTER
                equation.paragraph_format.first_line_indent = Cm(0)
                add_text(equation, "U(r,t) = N(r,t) / C(r,t)", western=LATIN_FONT, italic=True)
                p = doc.add_paragraph(); set_para_format(p)
                add_text(p, "其中，U(r,t)表示区域r在时段t的座位占用率，N(r,t)表示该时段观察到的使用座位数，C(r,t)表示同时段可使用座位数。分母为零或区域未开放时，该记录标记为不适用，不赋值为零。")
        if chapter_title.startswith("第3章"):
            add_caption(doc, "表3.1　观察变量与计算口径")
            add_table(doc, ["变量", "定义", "汇总单位"], [
                ["区域占用率", "使用座位数/开放座位数", "区域×时段"],
                ["设备可用率", "正常设备数/计划开放设备数", "区域×观察日"],
                ["协作活动比例", "协作活动数/可分类活动总数", "区域×时段"],
                ["记录完整度", "有效字段数/应填字段数", "单条观察记录"],
            ])
        if chapter_title.startswith("第4章"):
            add_caption(doc, "表4.1　合成样本的区域汇总指标")
            add_table(doc, ["虚构区域", "平均占用率", "高占用时段数", "设备可用率"], [
                ["安静学习区", "0.72", "18", "0.98"],
                ["协作讨论区", "0.66", "14", "0.96"],
                ["数字资源区", "0.59", "9", "0.91"],
            ])
            image_paragraph = doc.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.add_run().add_picture(make_figure("不同时段的模拟占用记录", ["上午", "午间", "下午", "晚间"], [54, 68, 82, 61]), width=Inches(5.4))
            add_caption(doc, "图4.1　不同时段的合成平均占用水平", wrong_style=True)
            add_caption(doc, "表4.2　数据质量检查结果（合成样本）")
            add_table(doc, ["检查项", "记录数", "处理方式"], [
                ["字段完整", "486", "进入汇总"],
                ["设备维护", "12", "标记后单独说明"],
                ["临时关闭", "6", "排除占用率分母"],
                ["重复记录", "0", "无须删除"],
            ])
            image_paragraph = doc.add_paragraph()
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.add_run().add_picture(make_figure("可追溯的分析流程", ["记录", "校验", "汇总", "复核"], [48, 40, 35, 28]), width=Inches(5.4))
            add_caption(doc, "图4.2　合成数据处理与复核流程")

    add_heading(doc, "参考文献", 1)
    refs = [
        "[1] 合成条目：学习空间分时观察方法示例，测试资料，2026。",
        "[2] 合成条目：公共服务数据最小化实践示例，测试资料，2026。",
        "[3] Synthetic reference: fictional methods for library space observation, test material, 2026.",
        "注：以上均为本测试样稿中的虚构条目，不得引用为真实出版物。",
    ]
    for ref in refs:
        p = doc.add_paragraph(); set_para_format(p, first_indent=Cm(0), line=1.2)
        add_text(p, ref, western=LATIN_FONT, size=10.5)

    add_heading(doc, "附录A 观察记录字段", 1)
    add_body(doc, "字段字典用于说明合成观察表的最小结构。每条记录包含虚构区域代码、观察日期类别、时段代码、开放座位数、使用座位数、设备状态计数和活动类别计数。字段不包含姓名、学号、设备地址或精确个人移动信息。")
    add_caption(doc, "表A.1　合成记录字段字典")
    add_table(doc, ["字段名", "类型", "说明"], [
        ["area_code", "枚举", "虚构区域编号"],
        ["time_slot", "枚举", "上午/午间/下午/晚间"],
        ["open_seats", "整数", "计划开放的可用座位数"],
        ["occupied_seats", "整数", "观察时点的使用座位数"],
        ["equipment_status", "分类", "正常/维护/不适用"],
    ])

    add_heading(doc, "附录B 合成样本说明", 1)
    add_body(doc, "本论文的全部数值、图表、人物和机构名称均为程序构造的虚构测试数据。格式异常测试点分别位于正文中文字体、英文西文字体、字号和一条图题的段落格式。使用者可以在 Word 的问题检索列表中筛选对应问题，预览单项或批量修复，并重新分析。当前版本按南开大学 2026 版表 4.6 检查图表题注字体、字号、对齐、行距和段前段后距；题名空格、图表位置和表格单元格结构仍需人工检查。")
    add_body(doc, "测试结束后，请由作者决定是否保存修改后的当前文档。插件不应为每次修复另建输出副本；如需保留原始状态，可在第一次操作前由用户自行另存一份。")

    add_heading(doc, "致谢", 1)
    add_body(doc, "本测试样稿由软件验收流程自动生成。感谢参与格式规则审阅、合成样例设计和交互测试的虚构角色。此致谢仅为文档结构测试内容，不指向真实个人或团队。")

    doc.core_properties.title = "高校图书馆学习空间需求的分时观察与管理方法研究（合成测试样稿）"
    doc.core_properties.subject = "NKU Thesis Tool full synthetic Word acceptance fixture"
    doc.core_properties.author = "Synthetic test fixture"
    doc.core_properties.keywords = "synthetic, test-only, Word add-in"
    doc.core_properties.comments = "All research content and measurements are fictional."
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
