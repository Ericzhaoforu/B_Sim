from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


OUTPUT_DIR = Path("generated")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
DOCX_PATH = OUTPUT_DIR / "ITU-R初评超限后详细分析相关文献主要内容与工程应用指南.docx"

BODY_FONT = "Noto Sans CJK SC"
HEADING_FONT = "Noto Serif CJK SC"
MATH_FONT = "Cambria Math"
ACCENT = "1F4E78"
ACCENT_LIGHT = "D9EAF7"
ACCENT_VERY_LIGHT = "EEF5FA"
GRAY = "666666"
LIGHT_GRAY = "F2F2F2"
DARK = "1F1F1F"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text_color(cell, color: str) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor.from_string(color)


def set_cell_margins(cell, top=90, start=110, bottom=90, end=110) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    tr_pr.append(cant)


def set_run_font(run, font_name: str = BODY_FONT, size: float | None = None,
                 bold: bool | None = None, color: str | None = None) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_spacing(paragraph, before: float = 0, after: float = 6,
                          line: float = 1.35) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("第 ")
    set_run_font(run, size=9, color=GRAY)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run2 = paragraph.add_run()
    run2._r.append(fld_char1)
    run2._r.append(instr_text)
    run2._r.append(fld_char2)
    run3 = paragraph.add_run(" 页")
    set_run_font(run3, size=9, color=GRAY)


def add_toc(paragraph) -> None:
    run = paragraph.add_run()
    fld_char = OxmlElement("w:fldChar")
    fld_char.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "打开文档后右键目录并选择“更新域”以刷新页码。"
    fld_sep.append(text)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char, instr, fld_sep, fld_end])


def add_hyperlink(paragraph, text: str, url: str) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.extend([color, underline])
    new_run.append(r_pr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_body(doc: Document, text: str, *, bold_prefix: str | None = None,
             indent_cm: float = 0, keep_with_next: bool = False) -> None:
    p = doc.add_paragraph()
    if bold_prefix and text.startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        set_run_font(r1, bold=True)
        r2 = p.add_run(text[len(bold_prefix):])
        set_run_font(r2)
    else:
        r = p.add_run(text)
        set_run_font(r)
    p.paragraph_format.first_line_indent = Cm(indent_cm) if indent_cm else Cm(0.74)
    p.paragraph_format.keep_with_next = keep_with_next
    set_paragraph_spacing(p, after=6, line=1.45)


def add_bullet(doc: Document, text: str, level: int = 0) -> None:
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    r = p.add_run(text)
    set_run_font(r)
    p.paragraph_format.left_indent = Cm(0.74 + 0.55 * level)
    p.paragraph_format.first_line_indent = Cm(-0.45)
    set_paragraph_spacing(p, after=3, line=1.3)


def add_numbered(doc: Document, text: str, level: int = 0) -> None:
    p = doc.add_paragraph(style="List Number" if level == 0 else "List Number 2")
    r = p.add_run(text)
    set_run_font(r)
    p.paragraph_format.left_indent = Cm(0.74 + 0.55 * level)
    p.paragraph_format.first_line_indent = Cm(-0.45)
    set_paragraph_spacing(p, after=3, line=1.3)


def add_formula(doc: Document, text: str, note: str | None = None) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    cell = table.cell(0, 0)
    set_cell_shading(cell, ACCENT_VERY_LIGHT)
    set_cell_margins(cell, top=110, start=180, bottom=110, end=180)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, font_name=MATH_FONT, size=11.5, color=DARK)
    if note:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(note)
        set_run_font(r2, size=9, color=GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_callout(doc: Document, title: str, body: str, fill: str = ACCENT_LIGHT) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=120, start=160, bottom=120, end=160)
    p = cell.paragraphs[0]
    r = p.add_run(title)
    set_run_font(r, bold=True, color=ACCENT)
    p2 = cell.add_paragraph()
    r2 = p2.add_run(body)
    set_run_font(r2)
    set_paragraph_spacing(p2, after=0, line=1.35)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_table(doc: Document, headers: list[str], rows: Iterable[Iterable[str]],
              widths_cm: list[float] | None = None, font_size: float = 9) -> None:
    rows = list(rows)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_shading(cell, ACCENT)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        set_run_font(r, size=font_size, bold=True, color="FFFFFF")
        if widths_cm:
            cell.width = Cm(widths_cm[i])
    for ridx, row_data in enumerate(rows):
        row = table.add_row()
        set_cant_split(row)
        for i, text in enumerate(row_data):
            cell = row.cells[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if ridx % 2 == 1:
                set_cell_shading(cell, ACCENT_VERY_LIGHT)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i > 0 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(text))
            set_run_font(r, size=font_size)
            if widths_cm:
                cell.width = Cm(widths_cm[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_paragraph(style=f"Heading {level}")
    r = p.add_run(text)
    set_run_font(r, font_name=HEADING_FONT, bold=True, color=ACCENT)
    if level == 1:
        r.font.size = Pt(16)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(8)
    elif level == 2:
        r.font.size = Pt(13)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(5)
    else:
        r.font.size = Pt(11.5)
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True


def configure_document(doc: Document) -> None:
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.1)
    sec.left_margin = Cm(2.35)
    sec.right_margin = Cm(2.15)
    sec.header_distance = Cm(1.0)
    sec.footer_distance = Cm(1.0)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = BODY_FONT
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.font.size = Pt(10.5)
    normal.paragraph_format.line_spacing = 1.45
    normal.paragraph_format.space_after = Pt(6)

    for level, size in ((1, 16), (2, 13), (3, 11.5)):
        st = styles[f"Heading {level}"]
        st.font.name = HEADING_FONT
        st._element.rPr.rFonts.set(qn("w:eastAsia"), HEADING_FONT)
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(ACCENT)

    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("ITU-R初评超限后详细分析相关文献主要内容与工程应用指南")
    set_run_font(hr, size=8.5, color=GRAY)
    add_page_number(sec.footer.paragraphs[0])


def add_cover(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(55)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ITU-R初评超限后详细分析相关文献")
    set_run_font(r, font_name=HEADING_FONT, size=25, bold=True, color=ACCENT)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("主要内容与工程应用指南")
    set_run_font(r2, font_name=HEADING_FONT, size=22, bold=True, color=ACCENT)

    line = doc.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = line.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    set_run_font(run, size=14, color=ACCENT)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(20)
    r3 = p3.add_run("围绕1215–1300 MHz星载SAR/有源传感器对RNSS接收机的\n脉冲干扰初评、详细分析、动态聚合与实机验证")
    set_run_font(r3, size=13, color=DARK)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    info = [
        ("核心建议书", "ITU-R M.2030-0、ITU-R RS.2165-0、ITU-R RS.1347-0"),
        ("核心研究报告", "ITU-R M.2220-1、M.2305-1、M.2496-0、RS.2311-0、RS.2537-0、RS.2566-0"),
        ("整理日期", "2026年10月18日"),
    ]
    for i, (a, b) in enumerate(info):
        c0, c1 = table.rows[i].cells
        set_cell_shading(c0, ACCENT_LIGHT)
        set_cell_margins(c0, top=110, start=140, bottom=110, end=140)
        set_cell_margins(c1, top=110, start=140, bottom=110, end=140)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr0 = p0.add_run(a)
        set_run_font(rr0, bold=True, color=ACCENT)
        rr1 = c1.paragraphs[0].add_run(b)
        set_run_font(rr1)

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(55)
    r4 = p4.add_run("技术总结文档｜供方案论证、项目评审和内部培训使用")
    set_run_font(r4, size=10, color=GRAY)
    doc.add_page_break()


def build_document() -> Document:
    doc = Document()
    configure_document(doc)
    add_cover(doc)

    add_heading(doc, "文档说明", 1)
    add_body(doc, "本文将前期整理的九份ITU-R文献按“初步评估—参数生成—单源算例—详细分析触发—轨道动态聚合—接收机实测—多源验证”的工程逻辑重新组织，目的在于说明每份文献解决什么问题、输出什么参数，以及它们如何组合成一套可执行的详细干扰分析流程。")
    add_callout(doc, "重要编号说明", "Report ITU-R RS.2537-0正文的适用范围引用Recommendation ITU-R RS.2160；本资料包中包含的是Recommendation ITU-R RS.2165-0。两者编号不同，本文按各自文件的实际标题和内容分别表述，不将其视为同一文件。")
    add_body(doc, "本文属于技术总结与工程方法提炼，不替代ITU-R正式文本。对外提交频率协调、许可申请或系统合规结论时，应回到正式建议书、报告、无线电规则及项目实测参数逐项核验。")

    add_heading(doc, "目录", 1)
    toc_p = doc.add_paragraph()
    add_toc(toc_p)
    doc.add_page_break()

    add_heading(doc, "一、九份文献的总体关系", 1)
    add_body(doc, "九份文献不是彼此独立的资料，而是分别承担数学建模、输入参数生成、应用算例、决策程序、动态仿真和硬件验证等角色。最核心的关系可以概括为：")
    add_formula(doc, "M.2220（生成PDC、R_I） → M.2030（计算劣化） → RS.2537（单SAR初评算例）")
    add_formula(doc, "RS.2165（详评触发与协调） → M.2305（多星动态聚合） → M.2496 / RS.2311（接收机实测） → RS.2566（多SAR聚合算例）")
    add_body(doc, "RS.1347则是一份较早但链条较完整的端到端兼容性案例，覆盖峰值链路预算、前端限幅、AGC/ADC、轨道统计和硬件注入试验，可作为理解接收机内部机理的补充材料。")

    overview_rows = [
        ("M.2030-0", "建议书", "数学判定核心", "已知PDC、R_I等参数后，计算RNSS接收机有效噪声和新增劣化。"),
        ("M.2220-1", "研究报告", "参数生成基础", "把真实雷达、传播和接收机条件转换为PDC、R_I及基线干扰参数。"),
        ("RS.2537-0", "研究报告", "单SAR应用实例", "演示1215–1300 MHz星载SAR对RNSS地面接收机的最坏峰值和性能劣化初评。"),
        ("RS.2165-0", "建议书", "决策与协调程序", "规定初评如何使用、何时进入详评、参与方及多星聚合要求。"),
        ("M.2305-1", "研究报告", "多星动态分析", "判断多颗SAR/散射计是否同时照射同一RNSS接收机，并统计事件次数与持续时间。"),
        ("M.2496-0", "研究报告", "接收机参数实测", "通过射频注入确定饱和点、恢复特性和扫描脉冲响应。"),
        ("RS.2311-0", "研究报告", "射频注入验证", "使用真实/代表性SAR与散射计波形测试特定RNSS接收机的C/N0或S/N劣化。"),
        ("RS.1347-0", "建议书", "早期端到端案例", "从前端安全、限幅和AGC/ADC一直分析到失锁、定位性能与硬件兼容测试。"),
        ("RS.2566-0", "研究报告", "多SAR聚合算例", "说明单颗分别合格并不保证多颗同时工作时仍满足允许劣化。"),
    ]
    add_table(doc, ["文献", "类型", "主要定位", "核心作用"], overview_rows, [2.6, 2.2, 3.2, 9.0], font_size=8.6)

    add_heading(doc, "二、ITU-R M.2030-0：脉冲干扰评估的数学核心", 1)
    add_heading(doc, "2.1 研究对象与适用范围", 2)
    add_body(doc, "M.2030用于初步评估RNSS以外的相关脉冲无线电源，对1164–1215 MHz、1215–1300 MHz和1559–1610 MHz频段RNSS系统及网络的影响。其输出不是简单的“干扰功率是否超门限”，而是把脉冲干扰、持续宽带干扰和接收机非线性共同转换为有效后相关噪声密度，再计算新增干扰造成的劣化比。")
    add_heading(doc, "2.2 两类基本接收机模型", 2)
    add_bullet(doc, "脉冲熄减接收机：检测到强脉冲后，在脉冲期间将相关器输入置零；高于熄减门限的脉冲用PDC描述，低于门限的脉冲用R_I描述。")
    add_bullet(doc, "饱和/限幅接收机：没有理想脉冲熄减功能，强脉冲可能使LNA、混频器、AGC或ADC进入饱和；模型引入N_LIM描述A/D饱和电平相对于1σ噪声电压的比值。")
    add_heading(doc, "2.3 关键公式与物理意义", 2)
    add_formula(doc, "PDC_i = (PW_i + τ_REC) · PRF_i", "单个强脉冲源造成的有效信号丢失占空比；恢复时间必须计入。")
    add_formula(doc, "PDC_total = 1 − ∏(1 − PDC_i)", "多个强脉冲源按“未受干扰时间的乘积”聚合，不能简单相加。")
    add_formula(doc, "R_I,total = Σ R_I,i", "低于门限的弱脉冲按平均功率密度比线性求和。")
    add_formula(doc, "D = 10·log10[N0,EFF+Y / N0,EFF]", "新增脉冲源相对于既有基线环境造成的增量劣化。")
    add_heading(doc, "2.4 1215–1300 MHz典型参数", 2)
    add_table(doc, ["接收机类型", "N_LIM", "基线PDC", "基线R_I", "基线I0,WB/N0", "允许新增劣化"], [
        ("SBAS地面参考接收机", "1", "0.0793", "0", "0.3925", "0.2 dB"),
        ("高精度半无代码接收机", "2", "0.0765", "0", "0.3983", "0.2 dB"),
        ("航空导航接收机（1 μs恢复）", "1", "0.1327", "0", "0.455", "0.1 dB"),
        ("航空导航接收机（30 μs恢复）", "1", "0.1723", "0", "0.455", "0.1 dB"),
    ], [4.4, 1.4, 1.9, 1.5, 2.4, 2.4], font_size=8.4)
    add_callout(doc, "理解重点", "允许劣化值不是接收机全部干扰容限，而是在已有基线脉冲和持续干扰环境之上，新规划脉冲源可以占用的剩余性能预算。")

    add_heading(doc, "三、Report ITU-R M.2220-1：从真实场景生成PDC和R_I", 1)
    add_heading(doc, "3.1 文献解决的问题", 2)
    add_body(doc, "M.2220回答的是“如何由发射机、传播路径和接收机参数得到M.2030需要的输入”。它采用干扰源—传播路径—受扰接收机的完整链路框架，要求收集发射源位置、频率、峰值功率、脉宽、PRF、天线方向图和发射谱，并结合距离、到达角、传播损耗、接收滤波器、饱和门限和恢复时间进行计算。")
    add_heading(doc, "3.2 分析对象与步骤", 2)
    for text in [
        "计算每个脉冲源到RNSS接收机的峰值接收功率。",
        "与接收机熄减门限或饱和门限比较，将脉冲分成强脉冲与弱脉冲。",
        "对强脉冲计算有效脉宽和PDC；对弱脉冲计算等效平均功率密度比R_I。",
        "将接收机过载恢复时间并入强脉冲信号丢失时间。",
        "对异构脉冲源进行PDC补集乘积聚合，对R_I直接求和。",
        "将聚合结果输入M.2030，计算相对于基线环境的新增劣化。",
    ]:
        add_numbered(doc, text)
    add_heading(doc, "3.3 典型案例：SBAS基线PDC来源", 2)
    add_body(doc, "报告用两部地面监视雷达构造SBAS接收机的基线脉冲环境。两个源分别得到PDC_A=0.001074和PDC_B=0.0783，聚合后为：")
    add_formula(doc, "PDC_agg = 1 − (1−0.001074)(1−0.0783) = 0.0793")
    add_body(doc, "0.0793随后被M.2030和RS.2537作为SBAS地面参考接收机的保守基线PDC。该例说明M.2220是整套方法中的“参数生成器”，没有它，M.2030中的PDC和R_I就难以从实际电磁环境中得到。")

    add_heading(doc, "四、Report ITU-R RS.2537-0：单颗星载SAR的初评应用实例", 1)
    add_heading(doc, "4.1 研究对象", 2)
    add_body(doc, "RS.2537针对1215–1300 MHz频段，列出六类代表性星载SAR和多类RNSS地面接收机参数，并给出M.2030方法的应用实例。该报告面向RNSS空对地接收地球站，未验证其对航天器上的RNSS空对空接收机是否适用。")
    add_heading(doc, "4.2 两阶段评估", 2)
    add_numbered(doc, "前端安全与饱和筛查：假设接收机位于SAR主波束内，计算最坏峰值接收功率，依次与生存电平、饱和电平和“饱和点以下15 dB”的工程筛查线比较。")
    add_numbered(doc, "性能劣化计算：考虑SAR调频谱与RNSS预相关通带的重叠，计算有效脉宽、PDC以及最终有效噪声劣化。")
    add_formula(doc, "P_RFI,pk = P_t + G_t + G_r − L_fs − L_pol − L_misc")
    add_formula(doc, "PW_eff = PW · (Δf / B_chirp)")
    add_formula(doc, "PDC = (PW_eff + τ_r) · PRF")
    add_heading(doc, "4.3 典型结果", 2)
    add_bullet(doc, "SAR5在1236.5 MHz工作时，14 MHz调频波形最大劣化约0.387 dB，28 MHz调频波形约0.349 dB，均超过SBAS接收机0.2 dB允许新增劣化。")
    add_bullet(doc, "SAR5改用1257.5 MHz或1278.5 MHz时，在报告采用的理想矩形接收滤波器近似下不再与SBAS通带重叠，初评PDC降为零。")
    add_bullet(doc, "对GLONASS接收机，1 μs恢复时间对应约0.64 dB劣化；30 μs恢复时间对应约1.16 dB，说明恢复时间可能直接改变评估结论。")
    add_heading(doc, "4.4 方法局限", 2)
    add_body(doc, "RS.2537主要采用单源、静态最坏条件和代表性矩形滤波器进行初评。它没有完整描述真实过境期间的斜距、天线增益、照射时长、多星同时照射概率以及接收机真实滤波器裙边，因此初评超限后必须转入更详细的轨道动态、频谱耦合和实机验证。")

    add_heading(doc, "五、ITU-R RS.2165-0：初评到详评的决策与协调程序", 1)
    add_heading(doc, "5.1 文献定位", 2)
    add_body(doc, "RS.2165不是新的劣化公式，而是规定应如何组合M.2030、M.2220、M.1901、M.1902和M.1787开展初评，并规定何时进入详细分析、由谁参与以及如何处理多星聚合影响。")
    add_heading(doc, "5.2 详细分析触发条件", 2)
    add_bullet(doc, "应用初评方法后，结果超过M.2030规定的允许劣化比。")
    add_bullet(doc, "受影响RNSS系统或网络的主管部门提出开展详细分析的要求。")
    add_heading(doc, "5.3 详细分析应包含的内容", 2)
    add_bullet(doc, "让目标EESS（有源）传感器的操作主管部门和受影响RNSS系统/网络主管部门共同参与。")
    add_bullet(doc, "采用所研究SAR的真实运行参数，而不是代表性最坏模板。")
    add_bullet(doc, "采用详细RNSS接收机特性，包括天线、滤波器、AGC/ADC、饱和点、恢复时间和脉冲抑制算法。")
    add_bullet(doc, "纳入所有已知的规划中和正在运行、可能同时照射同一接收机的星载有源传感器。")
    add_heading(doc, "5.4 聚合干扰缓解", 2)
    add_body(doc, "RS.2165明确提出，多颗星载有源传感器的潜在聚合干扰可通过运营方之间的运行协调加以缓解，例如错开工作时段、照射区域、成像模式或波束计划。")

    add_heading(doc, "六、Report ITU-R M.2305-1：轨道级、多星级动态聚合分析", 1)
    add_heading(doc, "6.1 研究问题", 2)
    add_body(doc, "M.2305补足RS.2537缺少的轨道和时间维度，研究多颗SAR或散射计在什么时间、什么地理位置会同时照射同一RNSS接收机，以及这种事件发生多少次、持续多久。")
    add_heading(doc, "6.2 动态仿真方法", 2)
    for text in [
        "建立太阳同步轨道、波束观察角、二维/三维发射方向图和散射计旋转扫描规律。",
        "在不同纬度布设RNSS地面接收机，并引入接收天线方向图和强脉冲判定门限。",
        "逐时刻计算每颗传感器的星地斜距、离轴角和接收功率时间序列。",
        "采用足够小的时间步长：SAR组合可用秒级步长；快速旋转散射计需在共同过境窗口内使用几十毫秒级精细步长。",
        "逐时刻将信号划分为强—强、强—弱、弱—强和弱—弱组合。",
        "统计双星/多星事件次数、最长持续时间、中位持续时间及纬度分布。",
    ]:
        add_numbered(doc, text)
    add_formula(doc, "P_r,i(t) = EIRP_i + G_r,i(t) − L_fs,i(t) − L_other,i")
    add_heading(doc, "6.3 输出与局限", 2)
    add_body(doc, "M.2305的主要输出是多星同时照射事件的概率和持续时间，而不是所有接收机的最终导航性能。事件统计结果仍需与M.2220生成的PDC/R_I和M.2030劣化公式结合，才能给出最终兼容性判断。")

    add_heading(doc, "七、Report ITU-R M.2496-0：接收机饱和参数与扫描脉冲响应实测", 1)
    add_heading(doc, "7.1 研究重点", 2)
    add_body(doc, "M.2496聚焦RNSS接收机内部实现，包括LNA压缩、混频与中频链路、AGC、ADC动态范围、输入饱和电平、恢复时间、脉冲熄减以及跟踪环路带宽。报告强调：饱和电平和恢复特性高度依赖设备实现，最可靠的办法是实测。")
    add_heading(doc, "7.2 饱和门限测试方法", 2)
    add_numbered(doc, "将可调功率脉冲干扰与模拟RNSS信号合成，直接注入接收机射频前端。")
    add_numbered(doc, "逐级提高干扰峰值功率并记录C/N0。")
    add_numbered(doc, "识别低于噪声区、非饱和渐进劣化区和饱和平台区。")
    add_numbered(doc, "将非饱和区向饱和平台区的转折点定义为详细分析所用的实际输入饱和电平。")
    add_heading(doc, "7.3 扫描波束与动态占空比", 2)
    add_body(doc, "报告指出，将旋转波束干扰简单按照射时间比例做长期平均的动态占空比折减并不普遍可靠。瞬时C/N0可能在照射期间明显下降，长时间平均会掩盖峰值性能损失。因此应同时输出平均值、最小瞬时值、劣化持续时间及恢复过程。")

    add_heading(doc, "八、Report ITU-R RS.2311-0：真实RNSS接收机的射频注入试验", 1)
    add_heading(doc, "8.1 研究目的", 2)
    add_body(doc, "RS.2311将代表性SAR和散射计波形硬耦合进入特定RNSS接收机的射频前端，直接测量S/N或C/N0变化，用于验证M.2030/M.2220解析模型并研究可能的缓解技术。")
    add_heading(doc, "8.2 试验变量", 2)
    add_table(doc, ["试验量", "代表取值"], [
        ("脉宽", "10、15、20、40、50 μs"),
        ("调频带宽", "1、5、20、40、78 MHz"),
        ("PRF", "1750 Hz、3500 Hz"),
        ("干扰电平", "−125、−105、−85、−75 dBW"),
        ("扫描波束门控", "导通100–200 ms，关闭约4 s，以模拟旋转波束"),
    ], [4.0, 12.5], font_size=9)
    add_heading(doc, "8.3 测试系统与评价量", 2)
    add_bullet(doc, "干扰波形直接注入天线之后的RF前端，并用频谱仪、示波器和峰值功率计进行校准。")
    add_bullet(doc, "评价量包括C/N0、S/N、捕获和跟踪状态、失锁、前端滤波、IF滤波、ADC和脉冲抑制的综合作用。")
    add_bullet(doc, "频谱耦合应采用“干扰波形实测频谱 × 接收机及天线综合滤波响应”，而不是只使用理想矩形通带。")
    add_heading(doc, "8.4 结论边界", 2)
    add_body(doc, "RS.2311只测试了特定接收机和特定强脉冲条件，测试结论不能自动推广到所有RNSS设备。不同前端滤波、ADC位数、AGC、脉冲检测、跟踪环路和接收带宽都可能导致明显不同的响应。")

    add_heading(doc, "九、ITU-R RS.1347-0：较早但完整的端到端兼容性案例", 1)
    add_heading(doc, "9.1 场景与分析链", 2)
    add_body(doc, "RS.1347分析1215–1260 MHz内两类标准星载SAR对GPS和GLONASS-M接收机的影响，覆盖峰值功率、前端烧毁与限幅、信号丢失占空比、AGC/ADC、跟踪环路、轨道统计及硬件兼容试验。")
    add_heading(doc, "9.2 主要分析结果", 2)
    add_bullet(doc, "最坏峰值低于前置放大器烧毁门限；GPS前端饱和门限不被超过。")
    add_bullet(doc, "GLONASS-M最坏情况下可能高于限幅门限约1.5 dB；通过15000轨道仿真，超限时间估计少于全天0.0019%，即不足约2 s/day。")
    add_bullet(doc, "按(PW+τ_r)·PRF计算，单颗SAR可造成约6%的信号时间丢失；恢复时间1–30 μs时，SNR劣化约为−0.1至−0.6 dB。")
    add_bullet(doc, "针对4 bit ADC和慢AGC，分析了ADC饱和、AGC重新设定、量化噪声和相关时间损失。")
    add_heading(doc, "9.3 硬件注入试验", 2)
    add_body(doc, "文献将SIR-C雷达信号注入GPS接收机，测试40 MHz带宽、1736 Hz及其他PRF、33.8 μs等脉宽，以及−75、−65和−43 dBW等输入电平。即使干扰提高到比预计空间最坏值高约32 dB的−43 dBW，接收机仍保持锁定，未观察到定位精度下降。")
    add_callout(doc, "使用注意", "RS.1347基于1990年代接收机结构和参数，更适合学习端到端分析方法，不应把其共存结论直接推广到现代多星座、多频点和宽带接收机。")

    add_heading(doc, "十、Report ITU-R RS.2566-0：多颗SAR聚合超限的算例", 1)
    add_heading(doc, "10.1 核心案例", 2)
    add_body(doc, "RS.2566给出多个星载SAR同时影响同一SBAS地面参考接收机的聚合计算实例：SAR1单独造成约0.019 dB劣化，SARA单独造成约0.196 dB劣化，二者分别低于0.2 dB允许值；但同时作用时总劣化约为0.215 dB，从而超过允许值。")
    add_formula(doc, "0.019 dB + 0.196 dB = 0.215 dB > 0.2 dB")
    add_callout(doc, "核心结论", "单颗系统分别合格，并不意味着多颗系统同时运行时仍然合格。对新SAR的评估必须纳入所有规划中和运行中的同时照射源。")
    add_body(doc, "报告同时指出，如何系统化避免或缓解多SAR聚合干扰，仍需进一步研究。因此它提供了聚合超限的定量实例，但不是完整的协调优化程序。")

    add_heading(doc, "十一、九份文献共同形成的详细分析流程", 1)
    add_heading(doc, "11.1 第一步：确认实际系统参数", 2)
    add_table(doc, ["对象", "最少需要的参数"], [
        ("SAR/EESS侧", "精确星历、轨道、工作模式、中心频率、实测发射谱、调频带宽、脉宽、PRF、三维方向图、观察角、极化、运行时段、照射区域及多星协同关系。"),
        ("RNSS侧", "接收天线三维方向图、RF/IF/预相关滤波响应、噪声温度、AGC、ADC位数、饱和/生存电平、恢复时间、脉冲熄减算法、捕获/跟踪环路和基线干扰环境。"),
    ], [3.2, 13.5], font_size=8.8)
    add_heading(doc, "11.2 第二步：开展轨道—姿态—波束动态仿真", 2)
    add_formula(doc, "P_r,i(t) = P_t,i + G_t,i(t) + G_r,i(t) − L_fs,i(t) − L_pol,i − L_misc,i")
    add_body(doc, "时间步长必须足以解析最窄波束和最快扫描运动；输出至少包括每颗卫星的接收峰值时间序列、照射窗口、主瓣/旁瓣状态和强脉冲门限超限事件。")
    add_heading(doc, "11.3 第三步：采用真实频谱耦合", 2)
    add_formula(doc, "P_in,i(t) = ∫ S_i(f,t) · |H_RNSS(f)|² df")
    add_body(doc, "S_i(f,t)应优先采用实测或高保真发射谱，H_RNSS(f)应包括天线滤波、RF、IF及预相关滤波器的综合响应。由此确定实际重叠带宽、有效脉宽和进入接收机的干扰功率。")
    add_heading(doc, "11.4 第四步：强弱脉冲分类与多源聚合", 2)
    add_bullet(doc, "峰值高于饱和/熄减门限：按强脉冲计算PDC。")
    add_bullet(doc, "峰值低于门限：按弱脉冲计算R_I。")
    add_bullet(doc, "接近门限：应考虑门限测量误差、温度漂移和批次差异。")
    add_formula(doc, "PDC_agg(t) = 1 − ∏[1 − PDC_i(t)]")
    add_formula(doc, "R_I,agg(t) = Σ R_I,i(t)")
    add_heading(doc, "11.5 第五步：计算随时间变化的性能劣化", 2)
    add_formula(doc, "D(t) = 10·log10[N0,EFF+Y(t) / N0,EFF]")
    add_body(doc, "详评不应只给最大值，还应统计超限总时长、单次最长超限时间、每次过境超限次数、年度/任务周期超限概率、不同纬度和地区风险、捕获与跟踪阶段差异以及多星同时照射事件数。")
    add_heading(doc, "11.6 第六步：射频注入与接收机验证", 2)
    add_table(doc, ["试验维度", "建议扫描内容", "主要输出"], [
        ("功率", "从低于噪声区扫描至饱和平台区", "饱和转折点、C/N0曲线、恢复时间"),
        ("频率/带宽", "中心频率、调频带宽、滤波裙边", "实际频谱耦合和有效脉宽"),
        ("时域", "脉宽、PRF、单次照射和扫描门控", "PDC、瞬时最低C/N0、恢复过程"),
        ("多源", "单源、双源和不同相对时序", "聚合PDC、R_I和超限组合"),
        ("任务性能", "捕获、跟踪、伪距、载波相位、定位", "失锁、定位误差和业务可用性"),
    ], [2.8, 6.6, 7.3], font_size=8.6)
    add_heading(doc, "11.7 第七步：缓解措施与复算", 2)
    for text in [
        "调整SAR中心频率或工作子带，降低与RNSS预相关通带的重叠。",
        "缩短脉宽、降低PRF或有效占空比。",
        "限制高风险波束观察角、照射区域和运行时段。",
        "避免多颗SAR在同一地区、同一频段同时工作。",
        "通过运营方任务协调错开成像计划。",
        "改善RNSS滤波、动态范围、脉冲检测和过载恢复速度。",
        "对缓解后的剩余最坏工况重新执行动态仿真和硬件注入。",
    ]:
        add_numbered(doc, text)

    add_heading(doc, "十二、项目实施中的文献分工与交付物", 1)
    add_table(doc, ["分析阶段", "主参考文献", "项目应形成的交付物"], [
        ("基线模型", "M.2220、M.1902、RS.2105", "发射源数据库、接收机参数库、基线PDC/R_I与持续干扰预算"),
        ("单源初评", "M.2030、RS.2537", "峰值生存/饱和表、频谱重叠表、有效脉宽、PDC、劣化值和初评结论"),
        ("详评决策", "RS.2165", "超限触发记录、双方参数确认清单、评审门与责任分工"),
        ("动态聚合", "M.2305、RS.2566", "轨道时序、同时照射事件、超限时间统计、地理风险图和多源聚合结果"),
        ("接收机实测", "M.2496、RS.2311、RS.1347", "饱和门限、恢复时间、C/N0曲线、失锁/定位结果和模型校准参数"),
        ("缓解与复核", "RS.2165及各算例", "频率/波形/波束/时序优化方案、复算结果及剩余风险"),
    ], [3.0, 4.2, 9.5], font_size=8.5)

    add_heading(doc, "十三、推荐阅读顺序", 1)
    reading = [
        "M.2030-0：先理解PDC、R_I、N_LIM和劣化公式。",
        "M.2220-1：学习如何由真实场景得到上述参数。",
        "RS.2537-0：查看单颗SAR的完整初评算例。",
        "RS.2165-0：理解初评与详评之间的评审门和参与方。",
        "M.2305-1：学习轨道、多星和同时照射统计。",
        "M.2496-0：学习真实接收机饱和参数和扫描脉冲响应的确定方法。",
        "RS.2311-0：学习射频硬注入试验和实测频谱耦合。",
        "RS.2566-0：理解单颗分别合格但聚合后超限的情况。",
        "RS.1347-0：作为从RF输入到导航任务性能的早期端到端案例补充阅读。",
    ]
    for item in reading:
        add_numbered(doc, item)
    add_callout(doc, "最核心的五份文献", "M.2030 + M.2220 + RS.2537 + M.2305 + M.2496，分别覆盖劣化公式、参数生成、单源算例、多星动态和接收机实测。")

    add_heading(doc, "十四、结论", 1)
    add_body(doc, "这组ITU-R文献共同说明，RS.2537式初评超限后，详细分析不应只是把链路预算做得更精细，而应同时完成四类升级：一是用真实发射谱和真实接收滤波器替代矩形带宽近似；二是用轨道、姿态和波束时序替代静态主瓣最坏值；三是把所有规划中和运行中的脉冲源纳入聚合；四是通过接收机射频注入获得饱和点、恢复时间和任务性能响应。只有当解析模型、动态统计和实机试验相互闭合时，才能形成可信的兼容性结论及缓解方案。")

    add_heading(doc, "附录A：核心公式速查", 1)
    add_table(doc, ["公式", "用途"], [
        ("PDC_i=(PW_i+τ_REC)·PRF_i", "单个强脉冲源的有效信号丢失占空比"),
        ("PDC_total=1−∏(1−PDC_i)", "多个强脉冲源的异构聚合"),
        ("R_I,total=ΣR_I,i", "多个弱脉冲源的等效噪声功率密度聚合"),
        ("PW_eff=PW·Δf/B_chirp", "线性调频信号的有效重叠脉宽"),
        ("P_r=P_t+G_t+G_r−L_fs−L_pol−L_misc", "峰值干扰链路预算"),
        ("P_in=∫S(f)|H(f)|²df", "真实发射谱与接收滤波器的频谱耦合"),
        ("D=10log10(N0,EFF+Y/N0,EFF)", "新增脉冲源相对于基线环境的劣化"),
    ], [7.4, 9.4], font_size=9)

    add_heading(doc, "附录B：来源文件清单", 1)
    sources = [
        ("Recommendation ITU-R M.2030-0", "Evaluation method for pulsed interference ... to RNSS systems and networks", "12/2012"),
        ("Report ITU-R M.2220-1", "Calculation method to determine aggregate interference parameters of pulsed RF systems ...", "2021"),
        ("Report ITU-R RS.2537-0", "Representative system characteristics and examples of evaluating interference ... from spaceborne SAR sensors", "09/2023"),
        ("Recommendation ITU-R RS.2165-0", "Evaluation of the potential for pulsed interference from planned and future spaceborne SAR sensors ...", "12/2023"),
        ("Report ITU-R M.2305-1", "Consideration of aggregate RFI event potentials from multiple EESS systems ...", "2025/2026版"),
        ("Report ITU-R M.2496-0", "RNSS receiver pulsed interference response and related parameter measurement", "2021"),
        ("Report ITU-R RS.2311-0", "Pulsed RF signal impact measurements and possible mitigation techniques ...", "09/2014"),
        ("Recommendation ITU-R RS.1347-0", "Feasibility of sharing between RNSS receivers and EESS(active)/SRS(active)", "1998"),
        ("Report ITU-R RS.2566-0", "Aggregate pulsed interference examples from multiple spaceborne SAR sensors", "03/2026"),
    ]
    add_table(doc, ["文件", "英文主题（简写）", "版本/年份"], sources, [4.4, 9.6, 2.8], font_size=8.2)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    r = p.add_run("—— 文档结束 ——")
    set_run_font(r, bold=True, color=ACCENT)

    return doc


def main() -> None:
    doc = build_document()
    doc.save(DOCX_PATH)

    # Basic structural validation: reopen and verify minimum content.
    check = Document(DOCX_PATH)
    assert len(check.paragraphs) > 120, "Document unexpectedly short"
    assert len(check.tables) > 15, "Expected tables were not generated"
    print(f"Generated: {DOCX_PATH}")
    print(f"Paragraphs: {len(check.paragraphs)}")
    print(f"Tables: {len(check.tables)}")
    print(f"Size: {DOCX_PATH.stat().st_size} bytes")


if __name__ == "__main__":
    main()
