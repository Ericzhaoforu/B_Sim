#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT_DIR = Path("out")
OUT_DIR.mkdir(exist_ok=True)
OUT_FILE = OUT_DIR / "ITU-R_M.2046与RS.2537-0详细解读与工程应用指南.docx"

CJK = "Noto Sans CJK SC"
LATIN = "Arial"
BLUE = "1F4E78"
MID_BLUE = "2F75B5"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EEF5FB"
DARK = "1F1F1F"
GRAY = "666666"
LIGHT_GRAY = "F2F2F2"
PALE_YELLOW = "FFF2CC"
PALE_GREEN = "E2F0D9"
PALE_RED = "FCE4D6"
WHITE = "FFFFFF"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_run_font(run, size=10.5, bold=False, color=DARK, italic=False, name=None):
    font_name = name or LATIN
    run.font.name = font_name
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), CJK if name is None else font_name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def add_field(paragraph, instruction: str):
    run = paragraph.add_run()
    fld_char = OxmlElement("w:fldChar")
    fld_char.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instruction
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "separate")
    fld_char3 = OxmlElement("w:fldChar")
    fld_char3.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char)
    run._r.append(instr_text)
    run._r.append(fld_char2)
    run._r.append(fld_char3)
    return run


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = add_field(paragraph, " PAGE ")
    set_run_font(run, 9, color=GRAY)


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def add_bookmark(paragraph, name: str, bid: int):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def configure_document(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(2.1)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(0.7)
    sec.footer_distance = Cm(0.7)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)
    normal.font.size = Pt(10.5)
    normal.paragraph_format.line_spacing = 1.35
    normal.paragraph_format.space_after = Pt(4)

    style_specs = {
        "Title": (24, BLUE, True),
        "Subtitle": (12, GRAY, False),
        "Heading 1": (16, BLUE, True),
        "Heading 2": (13.5, MID_BLUE, True),
        "Heading 3": (11.5, "365F91", True),
    }
    for name, (size, color, bold) in style_specs.items():
        st = doc.styles[name]
        st.font.name = LATIN
        st._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor.from_string(color)
        st.font.bold = bold
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.space_before = Pt(10)
        st.paragraph_format.space_after = Pt(5)

    custom = {
        "BodyTextCN": (10.5, DARK, False),
        "SmallTextCN": (9, GRAY, False),
        "EquationCN": (10.5, DARK, False),
        "SourceNote": (8.5, GRAY, False),
        "Callout": (10, DARK, False),
        "CaptionCN": (9, GRAY, False),
    }
    for name, (size, color, bold) in custom.items():
        if name not in doc.styles:
            doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st = doc.styles[name]
        st.font.name = LATIN
        st._element.rPr.rFonts.set(qn("w:eastAsia"), CJK)
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor.from_string(color)
        st.font.bold = bold
        st.paragraph_format.space_after = Pt(4)
        if name == "BodyTextCN":
            st.paragraph_format.line_spacing = 1.35
            st.paragraph_format.first_line_indent = Cm(0.74)
        elif name == "EquationCN":
            st.paragraph_format.line_spacing = 1.15
            st.paragraph_format.space_before = Pt(4)
            st.paragraph_format.space_after = Pt(6)

    header = sec.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = header.add_run("ITU-R M.2046-0 与 Report ITU-R RS.2537-0 详细解读")
    set_run_font(r, 8.5, color=GRAY)
    add_page_number(sec.footer.paragraphs[0])


def add_body(doc, text: str, bold_lead: str | None = None):
    p = doc.add_paragraph(style="BodyTextCN")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        set_run_font(r1, 10.5, bold=True, color=BLUE)
        r2 = p.add_run(text[len(bold_lead):])
        set_run_font(r2)
    else:
        r = p.add_run(text)
        set_run_font(r)
    return p


def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.left_indent = Cm(0.65 + 0.55 * level)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run("• " + item)
        set_run_font(r, 10.3)


def add_numbered(doc, items):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.left_indent = Cm(0.72)
        p.paragraph_format.first_line_indent = Cm(-0.55)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(f"{i}. {item}")
        set_run_font(r, 10.3)


def add_equation(doc, formula: str, explanation: str | None = None):
    p = doc.add_paragraph(style="EquationCN")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(formula)
    set_run_font(r, 11, bold=True, color=BLUE, name="Cambria Math")
    if explanation:
        q = doc.add_paragraph(style="SmallTextCN")
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr = q.add_run(explanation)
        set_run_font(rr, 9, color=GRAY)


def add_callout(doc, title: str, text: str, fill=PALE_YELLOW):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, 130, 160, 130, 160)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.style = doc.styles["Callout"]
    r1 = p.add_run(title + "：")
    set_run_font(r1, 10, bold=True, color=BLUE)
    r2 = p.add_run(text)
    set_run_font(r2, 10)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_table(doc, headers, rows, widths=None, font_size=8.8, first_col_bold=False):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False if widths else True
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_shading(cell, BLUE)
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h))
        set_run_font(r, font_size, bold=True, color=WHITE)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths:
            cell.width = Cm(widths[i])
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        for ci, value in enumerate(row):
            if ri % 2 == 1:
                set_cell_shading(cells[ci], PALE_BLUE)
            set_cell_margins(cells[ci])
            cells[ci].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cells[ci].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(value))
            set_run_font(r, font_size, bold=(first_col_bold and ci == 0))
            if widths:
                cells[ci].width = Cm(widths[ci])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def add_source(doc, text: str):
    p = doc.add_paragraph(style="SourceNote")
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    r = p.add_run("来源与定位：" + text)
    set_run_font(r, 8.5, color=GRAY, italic=True)


def add_flow_table(doc, headers, texts):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, (h, t) in enumerate(zip(headers, texts)):
        c = table.cell(0, i)
        set_cell_shading(c, [LIGHT_BLUE, PALE_GREEN, PALE_YELLOW, PALE_RED][i % 4])
        set_cell_margins(c, 150, 120, 150, 120)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run(h + "\n")
        set_run_font(r1, 10, bold=True, color=BLUE)
        r2 = p.add_run(t)
        set_run_font(r2, 9.2)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def main():
    doc = Document()
    configure_document(doc)

    # Cover
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ITU-R M.2046-0 与\nReport ITU-R RS.2537-0\n详细解读与工程应用指南")
    set_run_font(r, 24, bold=True, color=BLUE)
    p2 = doc.add_paragraph(style="Subtitle")
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run("非GSO移动卫星系统保护准则 · 星载SAR对RNSS脉冲干扰评估")
    set_run_font(r, 12, color=GRAY)
    doc.add_paragraph()
    meta = doc.add_table(rows=4, cols=2)
    meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta.style = "Table Grid"
    cover_rows = [
        ("文献一", "Recommendation ITU-R M.2046-0（12/2013，现行主版本）"),
        ("文献二", "Report ITU-R RS.2537-0（09/2023，现行报告）"),
        ("编制目的", "用易于理解的方式梳理研究对象、参数、公式、判据、算例与工程使用边界"),
        ("版本核对日期", "2026年9月17日"),
    ]
    for i, (a, b) in enumerate(cover_rows):
        for j, value in enumerate((a, b)):
            c = meta.cell(i, j)
            set_cell_shading(c, LIGHT_BLUE if j == 0 else WHITE)
            set_cell_margins(c, 130, 150, 130, 150)
            rr = c.paragraphs[0].add_run(value)
            set_run_font(rr, 10, bold=(j == 0), color=BLUE if j == 0 else DARK)
    doc.add_paragraph()
    note = doc.add_paragraph(style="SmallTextCN")
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = note.add_run("说明：本文为技术研读与工程方法提炼，不替代ITU正式文本、现行《无线电规则》或主管部门的频率指配与协调结论。")
    set_run_font(rr, 9, color=GRAY)
    doc.add_page_break()

    # TOC
    h = doc.add_paragraph(style="Heading 1")
    h.add_run("目录")
    toc = doc.add_paragraph()
    add_field(toc, 'TOC \\o "1-3" \\h \\z \\u')
    tip = doc.add_paragraph(style="SmallTextCN")
    rr = tip.add_run("提示：在Word中右击目录并选择“更新域”，即可显示或刷新页码。")
    set_run_font(rr, 9, color=GRAY)
    doc.add_page_break()

    # 1 Overview
    doc.add_heading("1  两份文献的定位与关系", level=1)
    add_callout(doc, "先建立正确认识", "M.2046-0是ITU-R建议书，给出一个具体非GSO移动卫星系统（ARGOS4）在399.9–400.05 MHz上行频段的系统描述和保护准则；RS.2537-0是ITU-R报告，给出1215–1300 MHz星载SAR对RNSS地面接收机脉冲干扰的代表性参数、分析方法和算例。两者频段、系统和干扰机理不同，数值不能相互套用，但“从接收机机理推导保护准则”的方法可以互相借鉴。", PALE_YELLOW)
    add_flow_table(doc,
        ["M.2046", "核心机理", "RS.2537", "共同方法"],
        ["399.9–400.05 MHz\n非GSO MSS上行", "宽带噪声抬升噪声底；窄带谱线占用检测资源", "1215–1300 MHz\nSAR脉冲干扰RNSS", "接收功率→接收机响应→性能劣化→时间概率/保护判据"])
    add_body(doc, "M.2046的重点是“受扰系统能承受多少干扰”。它从ARGOS4的噪声温度、最低可用信号、误码率目标和载波检测结构出发，分别推导宽带噪声和窄带干扰门限。")
    add_body(doc, "RS.2537的重点是“如何把星载SAR脉冲参数转换为RNSS接收机性能劣化”。它先检查前端损坏与饱和，再根据脉冲占空比、接收机恢复时间和频谱重叠计算有效信号损失，并与RNSS允许劣化比比较。")
    add_table(doc,
        ["比较维度", "ITU-R M.2046-0", "Report ITU-R RS.2537-0"],
        [
            ("文献性质", "建议书：给出推荐保护准则", "报告：给出技术参数、分析方法与算例"),
            ("频段", "399.9–400.05 MHz", "1215–1300 MHz"),
            ("受扰系统", "非GSO MSS ARGOS4星上接收机", "RNSS空对地接收地球站"),
            ("干扰方向", "地面发射源→低轨卫星", "星载SAR→地面RNSS接收机"),
            ("主要干扰类型", "宽带噪声、窄带谱线", "强/弱脉冲RFI，叠加连续干扰"),
            ("最终指标", "spfd或pfd＋不超过1%时间", "峰值功率、生存/饱和、PDC、等效噪声劣化比"),
            ("工程作用", "规定保护ARGOS4所需的输入干扰上限", "给出SAR–RNSS初步兼容性评估流程"),
        ], widths=[3.2, 6.8, 6.8], font_size=8.7, first_col_bold=True)

    # 2 M.2046
    doc.add_heading("2  ITU-R M.2046-0：ARGOS4非GSO移动卫星系统的特性与保护准则", level=1)
    doc.add_heading("2.1  文献要解决的问题", level=2)
    add_body(doc, "M.2046-0描述了一个在399.9–400.05 MHz频段工作、方向为地对空的非GSO移动卫星业务系统，并给出保护该系统免受宽带噪声和窄带干扰的准则。建议书采用的代表系统是ARGOS4数据采集系统，因此文中门限首先是ARGOS4专用门限，而不是所有400 MHz卫星接收机的通用门限。")
    add_source(doc, "Recommendation ITU-R M.2046-0，Scope、recommends 1–2及Annex 1。ITU官方目录将M.2046-0列为现行主版本。")
    add_callout(doc, "频段理解", "该建议书形成于2013年。正文中关于RNSS在该频段的历史性划分日期应按当时《无线电规则》理解；开展当前项目时仍需另行核对最新《无线电规则》和国家频率划分。", PALE_RED)

    doc.add_heading("2.2  ARGOS4系统怎样工作", level=2)
    add_flow_table(doc,
        ["数据采集平台DCP", "399.9–400.05 MHz上行", "低轨ARGOS4卫星", "任务数据下传"],
        ["传感器、浮标、动物跟踪器等\n通常低增益天线", "400 bit/s\nManchester编码PSK\n消息前含160 ms纯载波", "星上解调并复用\n载荷接收机搜索弱载波", "1670–1710、7750–7850、8025–8400 MHz；另有465.9875 MHz专用下行"])
    add_body(doc, "ARGOS数据采集平台通常功率较低、天线增益较小，且环境和安装方式差异大。卫星接收机必须从大范围覆盖区内同时发现并处理许多弱信号。系统采用星上解调，因此上行接收性能可以与后续下行链路性能分开分析。")
    add_table(doc,
        ["项目", "原文代表值", "易懂解释"],
        [
            ("上行频段", "399.9–400.05 MHz", "DCP向低轨卫星发送数据"),
            ("数据速率", "400 bit/s", "低速遥测，强调覆盖与低功耗而非高速率"),
            ("调制/编码", "Manchester编码、PSK", "便于时钟恢复与星上解调"),
            ("消息前导", "160 ms未调制载波", "供星上频谱分析器发现载波并建立锁定"),
            ("DCP天线", "典型低增益；40°仰角最大约3 dBi", "平台形态和安装环境使接收功率差别较大"),
            ("卫星接收天线", "RHCP/LHCP增益随卫星天底角变化", "必须用动态方向增益，不能始终使用峰值增益"),
        ], widths=[3.2, 5.0, 8.6], font_size=8.8, first_col_bold=True)

    doc.add_heading("2.3  为什么要分别规定宽带和窄带干扰", level=2)
    add_table(doc,
        ["干扰类型", "接收机看到的现象", "主要后果", "保护量"],
        [
            ("宽带噪声", "整个接收带宽内噪声底升高", "Eb/N0下降、误码率上升", "频谱功率通量密度spfd"),
            ("窄带干扰", "FFT中出现离散谱线，类似消息前导纯载波", "触发虚假检测、占用星上数据恢复单元，降低并发容量", "19 Hz分辨带宽内pfd"),
        ], widths=[2.8, 5.2, 5.1, 3.7], font_size=8.9, first_col_bold=True)
    add_body(doc, "这一区分十分关键。宽带噪声主要是“把噪声底抬高”；窄带干扰即使总功率不大，只要在一个很窄的频率单元中形成明显谱线，也可能被接收机误认为ARGOS信标前导，从而消耗处理资源。")

    doc.add_heading("2.4  宽带噪声保护准则如何推导", level=2)
    add_numbered(doc, [
        "建立接收机噪声基线：噪声系数3 dB，最坏背景噪声温度1200 K，天线至接收机损耗1.6 dB，得到接收机输入系统噪声温度1214 K。",
        "将系统噪声温度换算为噪声谱密度：N₀ = −197.8 dB(W/Hz)。",
        "采用最低接收信号C = −160 dBW。该条件下有效Eb/N₀约为8.3 dB。",
        "为达到BER = 2×10⁻⁴，需要的最低Eb/N₀约为8 dB，因此新增干扰最多只能消耗0.3 dB余量。",
        "把0.3 dB退化换算为新增噪声与热噪声之比，得到I₀/N₀约为−11.5 dB，I₀ = −209.3 dB(W/Hz)。",
        "根据天线峰值方向的有效孔径0.105 m²，并补偿1.6 dB馈线损耗，将接收机输入噪声密度换算为空间中的spfd门限。",
    ])
    add_equation(doc, "0.3 dB = 10·log₁₀[(N₀ + I₀)/N₀]", "新增宽带干扰只允许使总噪声增加0.3 dB")
    add_equation(doc, "I₀/N₀ = 10^(0.3/10) − 1 ≈ 0.0715 ≈ −11.5 dB", "新增干扰噪声约为热噪声的7.15%")
    add_equation(doc, "Aₑ = G·λ²/(4π)", "用接收天线增益把接收功率密度换算为空间功率通量密度")
    add_callout(doc, "宽带保护结果", "ARGOS4接收天线处的最大可接受聚合宽带噪声spfd为 −197.9 dB(W/(m²·Hz))，在卫星视场内超过该值的时间不得超过1%。对典型1600 Hz ARGOS4信号带宽，等效pfd约为 −165.8 dB(W/m²)。", PALE_GREEN)
    add_body(doc, "原文还指出，按ITU-R M.1475的性能目标推导，相关不可用时间可能只有0.018%；M.2046最终采用1%作为建议书中的时间判据。文献没有进一步解释为何选择更宽松的1%，因此项目中不应自行把这个1%解释为接收机物理极限。")

    doc.add_heading("2.5  窄带干扰保护准则如何推导", level=2)
    add_body(doc, "ARGOS4消息以160 ms纯载波开始。星上接收机用频谱分析器和FFT持续搜索离散载波，频率分辨率为19 Hz。任何超过检测门限的窄带谱线都会被分配给一个星上数据恢复单元（DRU）继续处理。若窄带干扰被误判为信标，就会占用有限的DRU资源，直接降低系统同时处理真实消息的能力。")
    add_numbered(doc, [
        "接收机要求最弱可检测载波满足Cmin/N₀ > 21 dB(Hz)。",
        "采用N₀ = −197.8 dB(W/Hz)，得到接收机输入端Cmin = −176.8 dBW。",
        "补偿天线到接收机之间1.6 dB损耗，换算到天线端为−175.2 dBW。",
        "采用最大天线有效孔径换算为pfd，得到−165.4 dB(W/m²)。",
    ])
    add_equation(doc, "Cmin = N₀ + 21 dB(Hz) = −176.8 dBW", "19 Hz谱线只要越过接收机检测门限，就可能触发虚假信标处理")
    add_callout(doc, "窄带保护结果", "在ARGOS4接收天线处，每个窄带干扰在19 Hz分辨带宽内的最大pfd为 −165.4 dB(W/m²)，超过该值的时间同样不得超过卫星视场时间的1%。", PALE_GREEN)
    add_body(doc, "建议书正文使用“每个窄带干扰”来表述门限；附件在结论中同时强调窄带谱线的聚合影响。工程上应逐谱线检查峰值，也要检查多个谱线同时触发处理资源后造成的总容量损失。后一项是基于接收机工作机理的工程解释，原文没有给出统一的DRU占用数量门限。")

    doc.add_heading("2.6  M.2046最终可直接使用的保护准则", level=2)
    add_table(doc,
        ["干扰类别", "保护准则", "统计时间", "保护对象"],
        [
            ("宽带噪声", "聚合spfd ≤ −197.9 dB(W/(m²·Hz))", "超过门限不得多于卫星视场时间的1%", "ARGOS4误码率/噪声余量"),
            ("窄带干扰", "19 Hz内每个干扰pfd ≤ −165.4 dB(W/m²)", "超过门限不得多于卫星视场时间的1%", "载波检测与星上处理容量"),
        ], widths=[3.0, 6.3, 4.4, 3.7], font_size=8.8, first_col_bold=True)

    doc.add_heading("2.7  参数取值依据与不能直接外推的地方", level=2)
    add_table(doc,
        ["参数", "文献取值", "取值依据/含义", "项目使用建议"],
        [
            ("背景噪声温度", "1200 K", "欧洲工业噪声条件下的最坏实测背景", "应按任务区域和在轨测量修正"),
            ("噪声系数", "3 dB", "ARGOS4典型接收机参数", "不能代替其他星上接收机规格"),
            ("最低接收信号", "−160 dBW", "ARGOS4设计最弱工作点", "应换成本项目灵敏度/链路预算值"),
            ("允许Eb/N₀退化", "0.3 dB", "8.3 dB工作点与8 dB BER要求之差", "属于具体波形和译码目标的剩余余量"),
            ("FFT分辨带宽", "19 Hz", "ARGOS4纯载波搜索结构", "其他接收机的窄带门限会随检测带宽改变"),
            ("1%时间", "卫星视场内", "建议书规定的统计判据", "应明确统计窗口、卫星视场定义和聚合场景"),
        ], widths=[3.1, 3.0, 6.2, 5.1], font_size=8.4, first_col_bold=True)
    add_callout(doc, "M.2046的使用边界", "它只描述一个ARGOS4系统，并不提供399.9–400.05 MHz所有MSS系统的统一参数，也没有给出多星座动态聚合仿真的完整流程。使用时应把它视为“接收机专用保护准则”，而不是频段通用I/N门限。", PALE_RED)

    # 3 RS.2537
    doc.add_heading("3  Report ITU-R RS.2537-0：星载SAR对RNSS接收机的脉冲干扰评估", level=1)
    doc.add_heading("3.1  报告的目的和适用范围", level=2)
    add_body(doc, "RS.2537-0面向1215–1300 MHz频段，给出六类代表性星载SAR、若干RNSS地面接收机的参数，以及把SAR脉冲干扰换算为RNSS性能劣化的通用分析方法和算例。它用于单个SAR对单个RNSS空对地接收地球站的初步评估，尚未验证用于航天器上的RNSS空对空接收机。")
    add_body(doc, "原报告正文引用Recommendation ITU-R RS.2160作为上层程序文件；该报告发布后，相关正式建议书以RS.2165-0编号批准。本文在介绍原报告时保留其原始引用，在工程关系说明中按现行RS.2165理解。")
    add_source(doc, "Report ITU-R RS.2537-0，Scope、Annex 1与Annex 2；现行ITU-R报告目录仍列有RS.2537。")
    add_callout(doc, "报告不是最终协调结论", "RS.2537给出的是“初步、代表性、偏保守”的分析方法。若初评超限，应使用真实轨道、真实天线方向图、实测接收滤波器、实际恢复时间和多源聚合条件进行详细分析。", PALE_YELLOW)

    doc.add_heading("3.2  六类代表性SAR参数", level=2)
    sar_rows = [
        ("SAR1", "400", "3200", "36.4 / 71.5", "40", "33.8", "1736", "5.87", "20–55"),
        ("SAR2", "568", "1200", "33.0 / 63.8", "15", "35", "1607", "5.62", "35"),
        ("SAR3", "757", "3200", "35.0 / 68.4", "78", "78", "2400", "18.7", "30"),
        ("SAR4", "628", "3950", "34.7 / 70.7", "84", "43–71", "1620–2670", "11.5", "7.2–59"),
        ("SAR5", "628", "6120", "36.6 / 74.5", "14或28", "37–67", "1050–1860", "7.0", "7.2–59"),
        ("SAR6", "628", "6120", "36.6 / 74.5", "28", "18–43", "1550–3640", "6.8", "7.2–59"),
    ]
    add_table(doc,
        ["系统", "高度 km", "峰值功率 W", "天线增益/峰值EIRP dB", "带宽 MHz", "脉宽 μs", "最大PRF Hz", "占空比 %", "视角 °"],
        sar_rows, widths=[1.5, 1.8, 2.0, 3.1, 2.0, 1.8, 2.0, 1.7, 2.0], font_size=7.8, first_col_bold=True)
    add_body(doc, "六类系统均采用线性调频脉冲和太阳同步轨道。它们覆盖不同高度、天线孔径、峰值EIRP、带宽、脉宽和占空比，用于代表现实SAR的宽参数范围。报告明确说明，算例选择的是对所研究RNSS接收机产生最坏干扰的特性，而不是每个SAR的日常平均工作状态。")
    add_body(doc, "报告不仅给出峰值增益，还给出SAR1至SAR6的分段二维天线方向图方程。这样在详细分析时可以根据卫星姿态、波束视角和接收机方向计算瞬时增益，而不是把峰值天线增益应用到整个过境过程。")

    doc.add_heading("3.3  RNSS接收机为什么必须分类型分析", level=2)
    add_body(doc, "不同RNSS接收机在天线增益、射频带宽、预相关滤波器、系统噪声温度、ADC量化、AGC、脉冲消隐、输入饱和点和过载恢复时间上差别很大。因此，同一SAR脉冲对不同接收机的影响可能相差几十dB或数倍占空比。")
    add_table(doc,
        ["接收机示例", "输入饱和电平", "生存电平", "恢复时间", "NLIM", "允许新增脉冲劣化"],
        [
            ("SBAS地面参考接收机", "−135 dBW（1 MHz）", "−10 dBW", "1 μs", "1", "0.2 dB"),
            ("高精度半无码接收机", "−120 dBW", "−20 dBW", "1 μs", "2", "0.2 dB"),
            ("航空导航FDMA接收机", "−80 dBW", "−1 dBW", "1–30 μs", "1", "0.1 dB"),
            ("室内定位接收机", "−70或−100 dBW（类型相关）", "−20或−17 dBW", "30 μs", "表9未给完整基线", "需详细分析"),
        ], widths=[4.0, 3.2, 2.8, 2.2, 1.5, 3.6], font_size=8.4, first_col_bold=True)
    add_body(doc, "NLIM表示ADC饱和电压相对于AGC建立的1σ噪声电压之比。1 bit硬限幅接收机常取NLIM=1；具有脉冲消隐的接收机可取NLIM=0。这个参数决定同一脉冲占空比如何映射为相关器后的等效噪声。")

    doc.add_heading("3.4  干扰机理：损坏、饱和和性能劣化不是一回事", level=2)
    add_table(doc,
        ["层级", "物理现象", "判据", "后果"],
        [
            ("前端生存", "LNA、滤波器、限幅器等承受过高峰值或平均功率", "与receiver survival level比较", "可能造成永久损坏"),
            ("输入饱和/压缩", "增益压缩、限幅或ADC削顶", "与input saturation level比较", "干扰脉冲期间及恢复时间内有用信号丢失"),
            ("性能劣化", "强脉冲造成时间空白；弱脉冲抬高等效噪声", "PDC、RI和允许劣化比", "C/N₀下降、捕获/跟踪性能恶化"),
        ], widths=[3.0, 6.0, 4.2, 4.0], font_size=8.8, first_col_bold=True)
    add_callout(doc, "关键区别", "“不会烧毁”不等于“不会干扰”。RS.2537的算例中，SAR峰值通常远低于RNSS生存电平，却经常高于输入饱和点，因此仍需继续计算性能劣化。", PALE_YELLOW)

    doc.add_heading("3.5  两级评估流程", level=2)
    add_flow_table(doc,
        ["第一级：峰值功率", "比较生存电平", "比较饱和区", "第二级：性能劣化"],
        ["最坏几何、主瓣、极化、自由空间损耗", "超过→存在硬件风险\n停止简单评估", "低于生存但高于“饱和点−15 dB”→继续", "计算频谱重叠、PDC、RI、基线干扰和允许劣化"])
    add_equation(doc, "Pᵣ,pk = Pₜ,pk + Gₜ + Gᵣ − LFS − Lpol − Lother", "所有量采用dB制；损坏/饱和初筛一般不计接收机窄带滤波带来的有利抑制")
    add_body(doc, "报告在初筛中采用共频、SAR主瓣照射RNSS接收机的最坏情况。自由空间损耗通常按1250 MHz计算；若SAR波形与某一RNSS具体通带重叠，则详细计算应在对应RNSS信号中心频率上求损耗。")
    add_body(doc, "如果接收峰值低于生存电平，但高于“输入饱和点以下15 dB”的筛查线，就进入性能劣化计算。原报告没有给出15 dB的器件级推导，因此应把它理解为保守的工程筛查带，而不是新的物理饱和定义。")

    doc.add_heading("3.6  强脉冲参数PDC：接收机有多少时间看不到有用信号", level=2)
    add_equation(doc, "PDCⱼ = (PWⱼ,eff + τᵣ) · PRFⱼ", "强脉冲超过饱和/消隐门限时，脉冲宽度和过载恢复时间都计入丢失时间")
    add_equation(doc, "PWₑff = PW · (Δf / Chirpwidth)", "只有落入RNSS预相关通带的chirp部分计入有效脉宽")
    add_body(doc, "PDC不是简单照搬SAR发射占空比。它同时受接收机滤波器、SAR中心频率、chirp带宽和接收机恢复时间影响。若频谱完全不重叠，Δf=0，初步模型中的PDC为0；若缺少滤波器衰减数据，报告建议在初评中采用频谱重叠因子1，作为最坏情况。")
    add_equation(doc, "PDCtotal = 1 − ∏ⱼ(1 − PDCⱼ)", "多个异构强脉冲源的时间覆盖率不能简单相加，应计算“至少一个脉冲出现”的概率")

    doc.add_heading("3.7  弱脉冲参数RI：未触发饱和的脉冲如何抬高噪声", level=2)
    add_equation(doc, "Rᵢ = Σ [Pᵢ · dcᵢ / (N₀ · BW)]", "将低于门限脉冲的平均功率密度与接收机热噪声谱密度比较")
    add_equation(doc, "RItotal = Σ Rⱼ", "弱脉冲的平均噪声贡献按功率线性相加")
    add_body(doc, "RS.2537把超过门限的脉冲归入PDC，把低于门限的脉冲归入RI。再与热噪声N₀、连续宽带干扰I₀,WB和ADC参数NLIM共同代入一般有效噪声密度公式。这样可以在同一框架内同时处理已有地面雷达、航空雷达、SAR和连续干扰。")

    doc.add_heading("3.8  从PDC到性能劣化", level=2)
    add_body(doc, "一般情况下，应使用报告式(7)，把基线PDC、基线RI、连续宽带干扰比I₀,WB/N₀、NLIM以及新增脉冲源共同带入。对于NLIM=1的硬限幅接收机，且弱脉冲贡献可忽略时，式(7)可简化为下面的直观表达。")
    add_equation(doc, "Dlinear = 1 / (1 − PDCY)²", "D表示加入新脉冲源后等效噪声与基线等效噪声之比")
    add_equation(doc, "DdB = −20·log₁₀(1 − PDCY)", "PDC越大，信号被切掉的时间越长，等效性能劣化越大")
    add_body(doc, "已有基线干扰不能被忽略。新增源加入后，强脉冲空白时间按未受扰时间相乘，弱脉冲噪声贡献线性相加：")
    add_equation(doc, "(1 − PDCbase+Y) = (1 − PDCbase)(1 − PDCY)")
    add_equation(doc, "RIbase+Y = RIbase + RY")
    add_callout(doc, "允许劣化不是总容限", "表9中的0.1 dB或0.2 dB是“新规划脉冲源在既有干扰基线之外可占用的新增预算”，不是接收机允许的总干扰劣化。", PALE_RED)

    doc.add_heading("3.9  三组峰值功率算例说明了什么", level=2)
    add_table(doc,
        ["RNSS接收机", "SAR最大接收峰值范围", "饱和电平", "生存电平", "初筛结论"],
        [
            ("QZSS室内定位", "约−83.16至−70.05 dBW", "−70 dBW", "−20 dBW", "不致损坏；多数结果位于饱和点15 dB范围内，应做性能评估"),
            ("GPS SBAS地面参考", "约−91.16至−78.05 dBW", "−135 dBW", "−10 dBW", "不致损坏；远高于饱和点，必须做性能评估"),
            ("GLONASS航空导航", "约−82.16至−69.05 dBW", "−80 dBW", "−1 dBW", "不致损坏；多数结果高于饱和点，必须做性能评估"),
        ], widths=[4.1, 4.2, 2.6, 2.6, 5.0], font_size=8.5, first_col_bold=True)

    doc.add_heading("3.10  SAR5对SBAS地面参考接收机的算例", level=2)
    add_body(doc, "算例假定SAR5在一段名义时间内使SBAS接收机进入饱和。SBAS预相关滤波器采用以1227.6 MHz为中心、20.5 MHz矩形等效带宽，恢复时间取1 μs。SAR5的chirp中心频率可选1236.5、1257.5和1278.5 MHz，带宽可选14或28 MHz。")
    add_table(doc,
        ["SAR5中心频率", "chirp带宽", "频谱重叠比", "PDC范围", "计算劣化", "与0.2 dB比较"],
        [
            ("1236.5 MHz", "14 MHz", "0.5964", "4.280%–4.361%", "最大0.387 dB", "超限"),
            ("1236.5 MHz", "28 MHz", "0.5482", "3.942%–4.023%", "最小0.349 dB", "超限"),
            ("1257.5 MHz", "14/28 MHz", "0", "0", "0（简化矩形模型）", "不重叠"),
            ("1278.5 MHz", "14/28 MHz", "0", "0", "0（简化矩形模型）", "不重叠"),
        ], widths=[3.1, 2.4, 2.7, 3.2, 3.6, 2.5], font_size=8.4, first_col_bold=True)
    add_callout(doc, "算例的核心启示", "中心频率是否与受扰接收机的实际预相关通带重叠，可能比单纯降低峰值功率更有效。频率规划必须基于接收机滤波响应，而不能只看整个1215–1300 MHz频段是否重叠。", PALE_GREEN)
    add_body(doc, "报告强调，超出0.2 dB只代表初评未通过，需要进一步详细分析；它不是在没有真实轨道和方向图的情况下直接判定系统不可运行。")

    doc.add_heading("3.11  SAR5对GLONASS接收机的算例", level=2)
    add_body(doc, "由于缺少GLONASS预相关滤波器在所需衰减电平处的有效带宽，算例采用频谱重叠因子1，即假定整个脉冲都有效进入接收机。这是明确的最坏情况假设。")
    add_table(doc,
        ["恢复时间", "PRF/脉宽组合", "PDC", "劣化", "结论"],
        [
            ("1 μs", "1050 Hz / 67 μs", "7.14%", "约0.64 dB（最大）", "超过0.1 dB，需要详细分析"),
            ("1 μs", "1860 Hz / 37 μs", "7.07%", "同量级", "超过允许值"),
            ("30 μs", "1050 Hz / 67 μs", "10.18%", "约0.93 dB量级", "恢复时间显著增加劣化"),
            ("30 μs", "1860 Hz / 37 μs", "12.46%", "约1.16 dB", "超过允许值，需要详细分析"),
        ], widths=[2.4, 4.2, 2.2, 3.4, 5.4], font_size=8.5, first_col_bold=True)
    add_body(doc, "这个算例说明，过载恢复时间不是次要参数。脉冲本身可能只有几十微秒，但如果接收机每次过载后还要恢复30 μs，累计PDC会明显增大。实际项目必须通过设备规范或射频注入试验获得恢复时间，不能默认使用1 μs。")

    doc.add_heading("3.12  RS.2537的局限与详细分析应补充的内容", level=2)
    add_bullets(doc, [
        "报告主要是单个SAR对单个RNSS接收机的初步评估；多个SAR同时照射的聚合影响需要另行计算。",
        "许多算例用“主瓣照射、共频、持续数分钟饱和”作为保守条件；实际动态场景应由轨道、姿态、波束时序和站址共同决定。",
        "表中RNSS接收机不覆盖所有已部署设备，尤其不能代表所有多星座、宽带或具有专有消隐算法的接收机。",
        "缺少真实滤波器衰减曲线时采用重叠因子1，会高估有效脉宽；详细分析应进行发射谱与接收滤波器的积分。",
        "接收机天线方向图、极化、下半球增益、AGC、ADC位数、限幅器和恢复时间均可能显著改变结果。",
        "初评超限后，应采用真实的峰值功率时间历程，而不是只用一个固定最坏值。",
    ])
    add_callout(doc, "报告给出的聚合缓解方向", "当多个星载有源传感器可能同时照射同一RNSS接收机时，一个直接的缓解方法是EESS运营方进行运行协调，例如错开成像时段、区域、中心频率或工作模式。", PALE_GREEN)

    # 4 Comparison
    doc.add_heading("4  两篇文献的共同分析思想", level=1)
    add_numbered(doc, [
        "先定义受扰接收机，而不是先假设一个通用I/N门限。M.2046从ARGOS4接收结构出发，RS.2537从不同RNSS接收机的饱和、恢复和量化结构出发。",
        "把干扰划分为不同机理。M.2046分宽带噪声与窄带谱线；RS.2537分硬件生存、强脉冲空白、弱脉冲噪声和连续干扰。",
        "保护准则必须对应性能指标。M.2046对应BER和处理容量；RS.2537对应RNSS等效噪声和允许性能劣化比。",
        "频谱重叠和接收滤波必须显式建模。只有进入实际接收通带的能量才参与性能计算。",
        "时间统计不可缺少。M.2046用视场内1%时间；RS.2537用PDC并要求最终详细分析恢复真实时间历程。",
        "聚合干扰不能只做单源相加。强脉冲用概率并集，弱脉冲用功率相加，多个卫星还需运行协调和动态仿真。",
    ])
    add_table(doc,
        ["通用工程步骤", "M.2046中的体现", "RS.2537中的体现"],
        [
            ("1. 定义系统与业务", "ARGOS4非GSO MSS上行", "SAR与RNSS接收地球站"),
            ("2. 建立接收机参数", "Tsys、NF、检测带宽、Cmin", "饱和、生存、恢复、NLIM、滤波器"),
            ("3. 识别干扰机制", "宽带噪声/窄带虚警", "损坏/饱和/强脉冲/弱脉冲"),
            ("4. 计算耦合", "spfd/pfd与天线有效面积", "峰值链路预算、天线方向图、极化"),
            ("5. 映射性能", "Eb/N₀、BER、DRU容量", "PDC、RI、等效噪声劣化"),
            ("6. 加入统计", "视场时间1%", "脉冲时间占比、动态轨道时间历程"),
            ("7. 判定与缓解", "满足两类pfd门限", "低于允许劣化；否则详细分析/协调"),
        ], widths=[4.1, 6.4, 6.4], font_size=8.7, first_col_bold=True)

    # 5 Project application
    doc.add_heading("5  可直接用于项目的分析模板", level=1)
    doc.add_heading("5.1  输入参数清单", level=2)
    add_table(doc,
        ["类别", "至少需要的参数"],
        [
            ("干扰发射机", "中心频率、峰值/平均功率、发射谱、脉宽、PRF、占空比、chirp方向和带宽、极化"),
            ("发射天线", "二维/三维方向图、峰值增益、扫描角、波束时序、旁瓣和后瓣"),
            ("传播与几何", "轨道/站址、斜距、可见窗口、姿态、自由空间损耗、附加传播损耗"),
            ("受扰天线", "方向图、极化、有效孔径、上/下半球增益、馈线损耗"),
            ("接收机前端", "RF滤波器、压缩/饱和点、生存电平、限幅器、AGC、ADC位数和满量程"),
            ("接收处理", "预相关带宽、检测分辨带宽、恢复时间、脉冲消隐、积分时间、处理资源容量"),
            ("性能与判据", "BER、C/N₀、允许劣化、不可用时间、处理容量或任务性能指标"),
            ("背景环境", "已有连续干扰、已有脉冲源、同时可见源数量和运行相关性"),
        ], widths=[4.0, 13.0], font_size=8.8, first_col_bold=True)

    doc.add_heading("5.2  推荐计算流程", level=2)
    add_numbered(doc, [
        "核对频率划分和业务地位，明确受保护业务、干扰方向和统计时间窗口。",
        "用真实发射谱与接收滤波器计算频谱重叠；不要只比较中心频率。",
        "完成峰值链路预算，先检查硬件生存，再检查饱和/压缩。",
        "按接收机门限将脉冲分为强脉冲和弱脉冲，计算PDC与RI。",
        "叠加已有连续干扰与脉冲基线，计算新增源引起的性能劣化。",
        "进行轨道—姿态—波束动态仿真，得到峰值功率、PDC或spfd随时间的序列及CDF。",
        "进行多源聚合，重点分析同时照射、最坏频率组合和异常恢复时间。",
        "如果初评超限，逐项测试缓解措施：频率调整、带宽/脉宽/PRF优化、旁瓣降低、接收滤波、脉冲消隐和运行协调。",
        "用射频注入试验验证饱和点、恢复时间、消隐效率、虚警和真实性能退化。",
    ])

    doc.add_heading("5.3  结论表达应避免的三种错误", level=2)
    add_callout(doc, "错误一", "把M.2046的−197.9 dB(W/(m²·Hz))用于任何400 MHz卫星接收机。该值由ARGOS4具体噪声与BER余量推导。", PALE_RED)
    add_callout(doc, "错误二", "只证明峰值功率低于生存电平，就声称“无干扰”。接收机可能不会损坏，但会长期饱和并丢失信号。", PALE_RED)
    add_callout(doc, "错误三", "把RS.2537的最坏情况初评超限直接写成“系统不可共存”。正确结论应是“需要真实动态参数和设备特性的详细分析”。", PALE_RED)

    # 6 Glossary
    doc.add_heading("6  术语与符号速查", level=1)
    add_table(doc,
        ["缩写/符号", "中文含义", "本文中的作用"],
        [
            ("MSS", "移动卫星业务", "M.2046中的ARGOS4业务类别"),
            ("non-GSO", "非对地静止卫星轨道", "低轨等非静止轨道系统"),
            ("DCP", "数据采集平台", "向ARGOS4卫星发送低速数据"),
            ("spfd", "频谱功率通量密度", "单位面积、单位带宽的入射干扰功率"),
            ("pfd", "功率通量密度", "单位面积内的干扰功率"),
            ("EESS（active）", "卫星地球探测业务（有源）", "星载SAR所属业务"),
            ("RNSS", "卫星无线电导航业务", "GPS、GLONASS、Galileo、QZSS等"),
            ("PDC", "超过门限脉冲的有效占空比", "强脉冲造成信号空白的时间比例"),
            ("RI", "低于门限脉冲的平均噪声密度比", "弱脉冲对热噪声的相对贡献"),
            ("NLIM", "ADC饱和电压与1σ噪声电压之比", "描述接收机量化/限幅结构"),
            ("τr", "过载恢复时间", "脉冲结束后接收机仍不可用的时间"),
            ("Δf", "chirp与接收通带的重叠带宽", "决定有效脉宽"),
        ], widths=[3.1, 5.2, 8.7], font_size=8.8, first_col_bold=True)

    # 7 Sources
    doc.add_heading("7  主要来源与版本说明", level=1)
    add_body(doc, "1）Recommendation ITU-R M.2046-0，Characteristics and protection criteria for non-geostationary mobile-satellite service systems operating in the band 399.9-400.05 MHz，12/2013。ITU官方目录显示该版本为现行主版本。")
    p = doc.add_paragraph(style="SourceNote")
    r = p.add_run("官方页面：https://www.itu.int/rec/R-REC-M.2046/en")
    set_run_font(r, 8.8, color=MID_BLUE)
    add_body(doc, "2）Report ITU-R RS.2537-0，Representative system characteristics and examples of evaluating interference into receiving earth stations in the RNSS from spaceborne SAR sensors in the EESS（active）in the 1215–1300 MHz band，09/2023。")
    p = doc.add_paragraph(style="SourceNote")
    r = p.add_run("官方报告目录：https://www.itu.int/pub/R-REP-RS/en")
    set_run_font(r, 8.8, color=MID_BLUE)
    add_body(doc, "3）用于理解RS.2537程序关系的配套文件包括ITU-R M.1902、M.2030、M.2220、M.2305及RS.2165。本文仅在RS.2537明确引用的范围内说明其作用，没有用配套文件替换原报告中的参数或结论。")
    add_callout(doc, "最终使用原则", "正式频率协调、许可证申请和系统验收必须以ITU正式文本、最新《无线电规则》、所在国家频率划分、主管部门要求和设备实测数据为准。", PALE_YELLOW)

    # Core properties
    props = doc.core_properties
    props.title = "ITU-R M.2046-0与RS.2537-0详细解读与工程应用指南"
    props.subject = "非GSO移动卫星系统保护准则与星载SAR对RNSS脉冲干扰评估"
    props.author = "OpenAI"
    props.keywords = "ITU-R, M.2046, RS.2537, MSS, ARGOS4, SAR, RNSS, 脉冲干扰, 频谱共用"
    props.comments = "基于ITU-R正式文献形成的中文技术研读材料。"

    doc.save(OUT_FILE)
    print(OUT_FILE)


if __name__ == "__main__":
    main()
