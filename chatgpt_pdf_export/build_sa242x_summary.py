#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a Chinese technical summary of ITU-R Reports SA.2425/2426/2427."""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

OUT = Path("generated/ITU-R_SA2425_SA2426_SA2427_Chinese_Summary.pdf")

PAGE_W, PAGE_H = A4
LEFT = 20 * mm
RIGHT = 18 * mm
TOP = 18 * mm
BOTTOM = 17 * mm

NAVY = colors.HexColor("#17365D")
BLUE = colors.HexColor("#2F5597")
LIGHT_BLUE = colors.HexColor("#D9EAF7")
PALE_BLUE = colors.HexColor("#EEF5FB")
DARK = colors.HexColor("#222222")
GREY = colors.HexColor("#666666")
LIGHT_GREY = colors.HexColor("#F2F2F2")
LINE = colors.HexColor("#B8C4CE")
GREEN = colors.HexColor("#2E7D32")
PALE_GREEN = colors.HexColor("#EAF4E8")
ORANGE = colors.HexColor("#C65911")
PALE_ORANGE = colors.HexColor("#FCE4D6")

pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
FONT = "STSong-Light"

styles = getSampleStyleSheet()
TITLE = ParagraphStyle(
    "TitleZH",
    fontName=FONT,
    fontSize=23,
    leading=31,
    textColor=NAVY,
    alignment=TA_CENTER,
    spaceAfter=8 * mm,
)
SUBTITLE = ParagraphStyle(
    "SubtitleZH",
    fontName=FONT,
    fontSize=12,
    leading=18,
    textColor=GREY,
    alignment=TA_CENTER,
)
H1 = ParagraphStyle(
    "H1ZH",
    fontName=FONT,
    fontSize=16,
    leading=22,
    textColor=NAVY,
    spaceBefore=6 * mm,
    spaceAfter=3 * mm,
    keepWithNext=True,
)
H2 = ParagraphStyle(
    "H2ZH",
    fontName=FONT,
    fontSize=13,
    leading=19,
    textColor=BLUE,
    spaceBefore=4 * mm,
    spaceAfter=2 * mm,
    keepWithNext=True,
)
H3 = ParagraphStyle(
    "H3ZH",
    fontName=FONT,
    fontSize=11.2,
    leading=17,
    textColor=DARK,
    spaceBefore=3 * mm,
    spaceAfter=1.5 * mm,
    keepWithNext=True,
)
BODY = ParagraphStyle(
    "BodyZH",
    fontName=FONT,
    fontSize=9.8,
    leading=16,
    textColor=DARK,
    alignment=TA_JUSTIFY,
    firstLineIndent=2 * FONT_SIZE if (FONT_SIZE := 9.8) else 0,
    spaceAfter=2.2 * mm,
)
BODY_NO_INDENT = ParagraphStyle(
    "BodyNoIndent",
    parent=BODY,
    firstLineIndent=0,
)
BULLET = ParagraphStyle(
    "BulletZH",
    parent=BODY,
    firstLineIndent=0,
    leftIndent=6 * mm,
    bulletIndent=1.5 * mm,
    spaceAfter=1.2 * mm,
)
NOTE = ParagraphStyle(
    "NoteZH",
    fontName=FONT,
    fontSize=8.5,
    leading=13,
    textColor=GREY,
    alignment=TA_LEFT,
    spaceAfter=2 * mm,
)
SMALL = ParagraphStyle(
    "SmallZH",
    fontName=FONT,
    fontSize=8.2,
    leading=12,
    textColor=DARK,
    alignment=TA_LEFT,
)
TABLE_HEAD = ParagraphStyle(
    "TableHeadZH",
    fontName=FONT,
    fontSize=8.6,
    leading=12,
    textColor=colors.white,
    alignment=TA_CENTER,
)
TABLE_CELL = ParagraphStyle(
    "TableCellZH",
    fontName=FONT,
    fontSize=8.2,
    leading=12,
    textColor=DARK,
    alignment=TA_LEFT,
)
TABLE_CENTER = ParagraphStyle(
    "TableCenterZH",
    parent=TABLE_CELL,
    alignment=TA_CENTER,
)
QUOTE = ParagraphStyle(
    "QuoteZH",
    fontName=FONT,
    fontSize=10.2,
    leading=17,
    textColor=NAVY,
    leftIndent=8 * mm,
    rightIndent=8 * mm,
    alignment=TA_CENTER,
    spaceBefore=3 * mm,
    spaceAfter=3 * mm,
)


def P(text: str, style=BODY) -> Paragraph:
    return Paragraph(text, style)


def bullets(items: list[str]) -> list[Paragraph]:
    return [Paragraph(item, BULLET, bulletText="•") for item in items]


def source_note(text: str) -> Paragraph:
    return P(f"资料定位：{text}", NOTE)


def make_table(data, widths, header=True, alignments=None, font_size=8.2):
    formatted = []
    for r, row in enumerate(data):
        fr = []
        for c, cell in enumerate(row):
            if isinstance(cell, Paragraph):
                fr.append(cell)
            else:
                if r == 0 and header:
                    fr.append(P(str(cell), TABLE_HEAD))
                else:
                    style = TABLE_CENTER if alignments and alignments[c] == "center" else TABLE_CELL
                    fr.append(P(str(cell), style))
        formatted.append(fr)
    table = Table(formatted, colWidths=widths, repeatRows=1 if header else 0, hAlign="LEFT")
    ts = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    if header:
        ts += [("BACKGROUND", (0, 0), (-1, 0), BLUE)]
        if len(data) > 2:
            ts += [("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE_BLUE])]
    table.setStyle(TableStyle(ts))
    return table


def box(text: str, fill=PALE_BLUE, border=BLUE):
    t = Table([[P(text, BODY_NO_INDENT)]], colWidths=[PAGE_W - LEFT - RIGHT])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), fill),
        ("BOX", (0, 0), (-1, -1), 0.8, border),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


def flow_diagram():
    cells = [
        P("SA.2426<br/><font size='8'>建立轨道、射频、天线和杂散参数基线</font>", TABLE_CENTER),
        P("→", TABLE_CENTER),
        P("SA.2425<br/><font size='8'>由系统规模和动态同频容量推导频谱需求</font>", TABLE_CENTER),
        P("→", TABLE_CENTER),
        P("SA.2427<br/><font size='8'>评估现有划分并开展同频、邻频和聚合共用研究</font>", TABLE_CENTER),
    ]
    widths = [48 * mm, 8 * mm, 48 * mm, 8 * mm, 55 * mm]
    t = Table([cells], colWidths=widths, hAlign="CENTER")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), PALE_GREEN),
        ("BACKGROUND", (2, 0), (2, 0), LIGHT_BLUE),
        ("BACKGROUND", (4, 0), (4, 0), PALE_ORANGE),
        ("BOX", (0, 0), (0, 0), 0.8, GREEN),
        ("BOX", (2, 0), (2, 0), 0.8, BLUE),
        ("BOX", (4, 0), (4, 0), 0.8, ORANGE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    return t


class NumberedDocTemplate(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(
            filename,
            pagesize=A4,
            leftMargin=LEFT,
            rightMargin=RIGHT,
            topMargin=TOP,
            bottomMargin=BOTTOM,
            title="ITU-R SA.2425/2426/2427内容总结与工程方法提炼",
            author="OpenAI",
            subject="ITU-R short-duration NGSO spectrum studies summary",
        )
        frame = Frame(LEFT, BOTTOM, PAGE_W - LEFT - RIGHT, PAGE_H - TOP - BOTTOM, id="normal")
        self.addPageTemplates(PageTemplate(id="main", frames=[frame], onPage=self.draw_header_footer))

    def draw_header_footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT, 8)
        canvas.setFillColor(GREY)
        if doc.page > 1:
            canvas.drawString(LEFT, PAGE_H - 10 * mm, "ITU-R SA.2425-0、SA.2426-0、SA.2427-0 内容总结")
            canvas.drawRightString(PAGE_W - RIGHT, PAGE_H - 10 * mm, "中文技术研读资料")
            canvas.setStrokeColor(LINE)
            canvas.setLineWidth(0.4)
            canvas.line(LEFT, PAGE_H - 12 * mm, PAGE_W - RIGHT, PAGE_H - 12 * mm)
        canvas.drawCentredString(PAGE_W / 2, 9 * mm, str(doc.page))
        canvas.restoreState()


story = []

# Cover
story.append(Spacer(1, 30 * mm))
story.append(P("ITU-R SA.2425-0、SA.2426-0、SA.2427-0", TITLE))
story.append(P("内容总结与工程方法提炼", TITLE))
story.append(Spacer(1, 5 * mm))
story.append(P("1 GHz以下短期任务非GSO卫星TT&C频谱需求、系统参数及共用兼容性研究", SUBTITLE))
story.append(Spacer(1, 18 * mm))
story.append(box(
    "三篇文件均为ITU-R研究报告，而非建议书。它们构成一套连续研究链：SA.2426给出系统参数，"
    "SA.2425据此估算频谱需求，SA.2427再判断现有划分能否满足需求，并对可能新增或升级的频率划分开展详细共用研究。",
    PALE_BLUE,
    NAVY,
))
story.append(Spacer(1, 18 * mm))
story.append(P("编制日期：2026年9月17日", SUBTITLE))
story.append(P("用途：技术研读、方案讨论与干扰分析方法参考", SUBTITLE))
story.append(Spacer(1, 24 * mm))
story.append(P(
    "说明：本文件根据三份ITU-R英文原报告及此前形成的中文参考译本整理。其作用是帮助理解报告的技术逻辑，"
    "不替代ITU正式文本；法规条款、频率划分、数值、保护准则和最终工程结论应回到英文原文复核。",
    NOTE,
))
story.append(PageBreak())

story.append(P("执行摘要", H1))
story.append(flow_diagram())
story.append(Spacer(1, 5 * mm))
story.append(P(
    "三篇报告研究的核心对象，是1 GHz以下短期任务非GSO卫星的遥测、跟踪与遥控（TT&C）。"
    "所谓“短期任务”主要指频率指配或任务有效期通常不超过3年，而不是简单以单颗卫星硬件寿命判断。"
    "若运营方持续补充替代卫星，使同一频率指配长期存在，则不属于报告中的短期任务范畴。"
))
story.append(P(
    "三篇报告得出的总体认识是：单颗小卫星以较低功率、25 kHz量级带宽和简单天线建立TT&C链路并不困难；"
    "真正制约频谱使用的是大量系统同时运行后的聚合干扰、地球站地域集中、受保护业务的极低干扰门限、"
    "频段的业务地位以及协调程序是否与短期任务的研制和运行周期匹配。"
))
story.append(P("核心结果速览", H2))
summary_table = [
    ["报告", "主要作用", "代表性结论"],
    ["SA.2426-0", "给出典型轨道、上下行射频、天线、噪声温度、链路预算及非必要发射参数", "典型VHF/UHF链路可建立正链路余量；25 kHz带宽、1 W星上发射、50 W地面上行为主要参数基线"],
    ["SA.2425-0", "用动态轨道和聚合干扰仿真估算300组卫星—地球站链路所需频谱", "主仿真得到空对地0.625～2.5 MHz、地对空0.682～0.938 MHz；地球站分布和协调显著改变容量"],
    ["SA.2427-0", "评估现有划分及150.05～174 MHz、400.15～420 MHz候选频段的共用可能性", "没有证明存在一个可简单全球统一新增的频段；多数候选子频段受安全、气象、射电天文、移动或其他空间业务强约束"],
]
story.append(make_table(summary_table, [27 * mm, 66 * mm, 79 * mm]))
story.append(PageBreak())

# SA.2425
story.append(P("一、Report ITU-R SA.2425-0", H1))
story.append(P("研究目标：由系统数量和同频容量推导TT&C频谱需求", H2))
story.append(P(
    "SA.2425回答的问题不是“某一频段能否共用”，而是“不断增加的短期任务非GSO卫星需要多少TT&C频谱”。"
    "报告先估计未来同时运行的系统数量，再用SA.2426提供的射频参数开展动态干扰仿真，求出一个25 kHz信道内"
    "可在无专门协调条件下同时承载多少组卫星—地球站链路。"
))
story.append(source_note("SA.2425-0第1节、第3节及表1～7。"))

story.append(P("1. 系统规模假设", H2))
story.extend(bullets([
    "统计2000年至2017年质量约1～10 kg的小卫星发射情况，发现2010年以后数量明显增长。",
    "对于不属于长期商业星座的单颗小卫星，2013年以后平均每年约发射60颗。",
    "按三年有效期估计，至少需要同时维持约180颗；为考虑增长和年度波动，报告采用300组卫星—地球站组合进行容量规划。",
]))
story.append(box(
    "300不是精确预测值，而是频谱需求研究中的保守容量假设。报告关注的是：当系统规模达到这一数量级时，"
    "无协调同频复用能力能否支撑其运行。",
    PALE_GREEN,
    GREEN,
))

story.append(P("2. 动态仿真的主要输入", H2))
params_2425 = [
    ["参数类别", "报告采用值"],
    ["下行代表频率", "137.5 MHz、401 MHz"],
    ["上行代表频率", "149 MHz、450 MHz"],
    ["轨道高度", "300～1000 km，并按数据库统计分段抽样"],
    ["轨道倾角", "0°～100°，按数据库统计分段抽样"],
    ["星上天线", "近似全向，增益≤3 dBi"],
    ["地球站天线", "跟踪天线；VHF 12 dBi，UHF 16 dBi"],
    ["接收噪声温度", "地面端VHF 1500 K、UHF 500 K；星上500～1000 K"],
    ["单链路必要带宽", "25 kHz"],
    ["仿真规模", "300颗卫星、300个地球站，一星一站随机配对"],
    ["仿真时长与步长", "7天，10 s"],
]
story.append(make_table(params_2425, [49 * mm, 123 * mm]))
story.append(P(
    "地球站没有按全球均匀分布，而是依据实际小卫星地球站数据库布设，因而集中在欧洲、美国和亚洲部分地区。"
    "这项处理很关键，因为可见锥的空间重叠和地球站地域聚集，是聚合干扰的重要来源。"
))

story.append(P("3. 基准可持续系数与频谱需求公式", H2))
story.append(P(
    "报告定义“基准可持续系数”，表示在没有TDMA、CDMA、轨道规划或跨运营方协调的情况下，"
    "同一个25 kHz信道可以支持的卫星—地球站组合数量。总频谱需求估算为：",
    BODY_NO_INDENT,
))
story.append(P("B<sub>req</sub> = (N<sub>total</sub> / N<sub>cochannel</sub>) × B<sub>single</sub>", QUOTE))
story.append(P(
    "例如，300组系统中若同频只能支持3组，则需要300/3×25 kHz=2.5 MHz。TDMA、CDMA、集中式星座管理、"
    "轨道规划和地面站协调都可能提高实际同频容量，因此这里得到的是未优化基线。"
))

story.append(P("4. 保护准则与仿真输出", H2))
story.extend(bullets([
    "地对空：载波干扰比C/I不得在每日超过1%的链路时间内低于20 dB。",
    "空对地：在1 kHz参考带宽内，接收机输入端总干扰功率不得在每日超过1%的时间内越过门限；报告使用137 MHz约−166.7 dBW/kHz、401 MHz约−176 dBW/kHz。",
    "每个时步计算可见性、斜距、天线离轴角、聚合干扰功率或C/I，再形成累积分布函数（CDF）。",
    "对同频系统数n=2～40逐档仿真，每档进行100次随机抽样，统计至少一组链路超限的概率。",
]))

story.append(P("5. 频谱需求结果", H2))
results_2425 = [
    ["场景", "频谱需求"],
    ["空对地：卫星在全部可见时段发射", "约1.875～2.5 MHz（主仿真条件不同）"],
    ["空对地：每日仅最长一次过境发射", "约0.625～0.938 MHz"],
    ["地对空：每日最长一次过境", "约0.682～0.938 MHz"],
    ["独立验证：VHF空对地", "约1.5～2.5 MHz"],
    ["独立验证：UHF空对地", "约1.07～3.75 MHz"],
]
story.append(make_table(results_2425, [85 * mm, 87 * mm]))
story.append(P(
    "结果的离散范围不是矛盾，而是说明同频容量高度依赖发射时序、地球站分布及采用的统计可接受条件。"
    "如果地球站可见区域大量重叠，超限几乎不可避免；若站点地理分散，则同一信道可支持更多系统。"
))
story.append(P("SA.2425可迁移的方法", H3))
story.append(box(
    "系统数量 → 轨道与站址分布 → 动态聚合干扰 → 同频可承载数量 → 总频谱需求。"
    "这套方法适用于卫星通信、卫星雷达和星座系统的频谱容量规划。",
    PALE_BLUE,
    BLUE,
))
story.append(PageBreak())

# SA.2426
story.append(P("二、Report ITU-R SA.2426-0", H1))
story.append(P("研究目标：建立可用于需求和共用分析的统一系统参数基线", H2))
story.append(P(
    "SA.2426不直接决定频率划分，而是定义“研究对象是什么样”。它为后续报告提供轨道、上下行射频、"
    "天线、接收机噪声、链路预算以及带外和杂散发射数据。"
))
story.append(source_note("SA.2426-0第2节及附件链路预算表。"))

story.append(P("1. 典型轨道和下行参数", H2))
params_2426_down = [
    ["参数", "典型值"],
    ["远地点/近地点高度", "300～1000 km"],
    ["轨道倾角", "0°～100°"],
    ["必要带宽", "≤25 kHz"],
    ["星上发射功率", "≤1 W"],
    ["星上天线", "近似全向，增益≤3 dBi，线极化"],
    ["地球站天线", "八木或抛物面；VHF典型12 dBi、UHF典型16 dBi"],
    ["地球站系统噪声温度", "VHF 1500 K；UHF 500 K"],
    ["最小仰角", "5°"],
    ["C/N目标", "12 dB"],
    ["接触期间占空比", "最高100%"],
]
story.append(make_table(params_2426_down, [62 * mm, 110 * mm]))
story.append(P(
    "25 kHz主要依据当时短期任务系统常见的不高于9.6 kbit/s数据率和至少约0.5 bit/s/Hz的频谱效率。"
    "报告假设链路只承担TT&C，而不承担大容量载荷数据传输。"
))

story.append(P("2. 典型上行参数", H2))
params_2426_up = [
    ["参数", "典型值"],
    ["必要带宽", "≤25 kHz"],
    ["地球站发射功率", "≤50 W（17 dBW）"],
    ["地球站天线增益", "VHF典型12 dBi、UHF典型16 dBi"],
    ["最小仰角/指向损耗", "5° / 1 dB"],
    ["星上系统噪声温度", "500～1000 K"],
    ["星上接收天线增益", "≤3 dBi"],
    ["C/N目标", "20 dB"],
    ["接触期间占空比", "最高100%"],
]
story.append(make_table(params_2426_up, [62 * mm, 110 * mm]))
story.append(P(
    "上行C/N目标高于下行，反映遥控链路对卫星安全控制的可靠性要求更高。"
))

story.append(P("3. 链路预算结论", H2))
link_table = [
    ["链路与场景", "报告给出的代表性链路余量"],
    ["空对地，300 km、5°仰角", "137 MHz约7.3 dB；400 MHz约8.8 dB"],
    ["空对地，1000 km、5°仰角", "137 MHz约0.7 dB；400 MHz约2.2 dB"],
    ["地对空，300 km、5°仰角", "148 MHz约20.8 dB；450 MHz约16.7 dB"],
    ["地对空，1000 km、5°仰角", "148 MHz约14.3 dB；450 MHz约10.1 dB"],
]
story.append(make_table(link_table, [84 * mm, 88 * mm]))
story.append(P(
    "最不利的高轨、低仰角下行链路仍为正余量，但裕度很小；上行因地面发射功率较大，余量明显更充足。"
    "因此，系统在链路层面可行，并不意味着它在聚合频谱环境中一定可协调。"
))

story.append(P("4. 带外和杂散发射", H2))
story.extend(bullets([
    "当时缺少1 GHz以下空间业务专用带外模板，报告借用ITU-R SM.1541中1～20 GHz空间业务模板。",
    "按ITU-R SM.1539，带外域与杂散域边界取必要带宽的250%；25 kHz带宽对应距中心频率62.5 kHz。",
    "实测星上和地面发射机均能满足SM.329的−43 dBc/4 kHz杂散“安全网”限值。",
    "从约120 kHz频偏起，实测设备可以达到更严格的−60 dBc/4 kHz，报告认为该值更适合作为工程参考。",
    "对射电天文连续谱观测，仅用4 kHz参考带宽的窄带模板不够，还需在更宽带宽内评价宽带噪声。",
]))
story.append(P("SA.2426可迁移的方法", H3))
story.append(box(
    "在任何共存研究中，先建立可追溯的参数基线：轨道、功率、带宽、功率谱密度、方向图、占空比、"
    "噪声温度、最小仰角、保护准则、带外谱和杂散谱。缺少这些参数时，后续干扰结论很难复核。",
    PALE_GREEN,
    GREEN,
))
story.append(PageBreak())

# SA.2427
story.append(P("三、Report ITU-R SA.2427-0", H1))
story.append(P("研究目标：判断现有频率划分能否满足需求，并评估新增或升级划分", H2))
story.append(P(
    "SA.2427使用SA.2425的频谱需求和SA.2426的系统参数，先审查1 GHz以下现有空间操作业务划分，"
    "再对150.05～174 MHz和400.15～420 MHz候选范围开展同频、邻频、单源和聚合兼容性研究。"
))
story.append(source_note("SA.2427-0第5～10节及附录A。"))

story.append(P("1. 现有频率划分的主要矛盾", H2))
alloc_table = [
    ["用途", "报告归纳的现有条件"],
    ["卫星识别", "0.005 MHz主要业务划分"],
    ["地对空遥控", "2.4 MHz主要业务划分，但受RR No. 9.21协调条件约束"],
    ["空对地遥测", "3 MHz主要业务划分；另有5.85 MHz次要业务资源"],
]
story.append(make_table(alloc_table, [56 * mm, 116 * mm]))
story.append(P(
    "报告认为，受RR No. 9.21约束的频段不适合短期任务：现有业务可以反对新系统进入，且协调过程可能"
    "远长于短期任务的研制与运行周期。若148～149.9 MHz不受该条件约束，其1.9 MHz带宽可能满足上行需求，"
    "但取消或改变协调条件所造成的影响仍需研究。"
))
story.append(P(
    "下行方面，137～138 MHz具有一定潜力；272～273 MHz和401～402 MHz已被其他系统广泛使用，"
    "不能仅凭频率划分表中的名义带宽判定可用。"
))

story.append(P("2. 候选频段中的受保护业务", H2))
business_table = [
    ["频率范围", "主要被研究业务"],
    ["150.05～174 MHz", "无线电定位雷达、固定、陆地/水上移动、航空移动、GMDSS、射电天文、移动卫星、广播"],
    ["400.15～420 MHz", "气象辅助、气象卫星、EESS数据采集、空间研究、固定/移动、射电天文、移动卫星、COSPAS-SARSAT、国际空间站链路"],
]
story.append(make_table(business_table, [47 * mm, 125 * mm]))

story.append(P("3. 报告采用的分析工具", H2))
story.extend(bullets([
    "最小耦合损耗和链路预算；",
    "自由空间损耗以及ITU-R P.452传播、地形、绕射、散射和地物杂波；",
    "单干扰源、多个地球站和多个卫星的聚合干扰；",
    "动态轨道仿真、超限时间统计和干扰功率CDF；",
    "同频、邻频、带外和杂散发射分析；",
    "隔离距离、保护带宽、EIRP降低、天线指向回避等缓解措施；",
]))
story.append(P(
    "SA.2427的突出价值，是把法规地位、系统参数、传播、动态轨道、方向图和接收机保护准则放在同一研究框架中。"
))

story.append(P("4. VHF候选频段的主要结论", H2))
story.extend(bullets([
    "150.05～153 MHz受到射电天文极低干扰门限限制；代表性计算得到地球站与射电天文站可能需要数百公里隔离。",
    "154～156 MHz存在空间监视/无线电定位雷达，共用研究显示互扰风险显著。",
    "156～162 MHz包含GMDSS遇险、安全和AIS相关频率，保护要求具有生命安全属性。",
    "162～174 MHz受到陆地和水上移动业务密集使用限制。",
    "最终汇总表中，150.05～174 MHz大多数候选子频段的地对空和空对地共用结论均为不可行。",
]))

story.append(P("5. UHF候选频段的主要结论", H2))
uhf_table = [
    ["子频段", "报告结论与原因"],
    ["400.15～401 MHz", "已有空间操作业务次要划分，但升级为主要业务会对气象卫星和空间研究接收系统产生过量干扰"],
    ["401～402 MHz", "已有空对地主要划分；地对空新增使用与EESS/MetSat数据采集系统不兼容，空对地也可能需要专门缓解"],
    ["402～403 MHz", "对气象卫星和EESS数据采集系统，共频地对空和空对地总体均不可行"],
    ["403～405 MHz", "部分固定探空仪站在有限区域和严格条件下可能局部共存；对下投式探空仪和全球统一划分不具普遍可行性"],
    ["405～406.1 MHz", "405.9～406 MHz需保护带，405～405.9 MHz应避免；406～406.1 MHz为COSPAS-SARSAT生命安全频段，应排除"],
    ["406.1～420 MHz", "受射电天文、固定、移动、空间研究等业务约束，汇总结论为地对空和空对地均不可行"],
]
story.append(make_table(uhf_table, [38 * mm, 134 * mm]))

story.append(P("6. 两个典型量化案例", H2))
story.append(P("射电天文隔离", H3))
story.append(P(
    "报告使用ITU-R P.452等传播模型估算地球站对射电天文站的干扰，代表性结果给出VHF约697 km、"
    "UHF约560 km的隔离距离；对空对地非必要发射，还提出候选频率至少远离射电天文频段边缘约1.5 MHz的估算。"
))
story.append(P("下投式探空仪", H3))
story.append(P(
    "即使假设地球站以最低功率、最低侧瓣增益并位于最大可见距离，下投式探空仪接收机仍可能超过保护准则"
    "约43.2～60.6 dB，影响距离可超过500 km。报告因此认为403～405 MHz不能被视为全球普遍可用的"
    "空间操作业务候选频段。"
))

story.append(P("7. SA.2427的总体结论", H2))
story.append(box(
    "报告没有找到一个可以简单、全球统一地新增给短期任务非GSO卫星使用的频段。"
    "名义带宽并非唯一瓶颈；业务地位、协调条件、现有使用密度、生命安全属性和受保护系统门限往往更关键。",
    PALE_ORANGE,
    ORANGE,
))
story.append(PageBreak())

# Combined
story.append(P("四、三篇报告合起来说明了什么", H1))
story.append(P("1. 链路可建立，不等于频谱可获得", H2))
story.append(P(
    "SA.2426表明，以1 W星上发射、25 kHz带宽和简单天线可以建立小卫星TT&C链路；"
    "SA.2425和SA.2427进一步表明，当系统数量上升到数百组时，聚合干扰、频率协调和既有业务保护才是决定因素。"
))

story.append(P("2. 频谱需求不能用“单星带宽×卫星数量”直接计算", H2))
story.append(P(
    "300颗卫星每颗25 kHz直接相乘为7.5 MHz，但轨道、可见时间和地域隔离提供了空间复用；"
    "反过来，也不能假定所有系统可以无限同频复用。应先通过动态仿真求出同频容量，再推导频谱需求。"
))

story.append(P("3. 静态最坏值与动态统计必须结合", H2))
story.append(P("静态链路预算 + 轨道动态仿真 + 干扰CDF + 超限时间概率", QUOTE))
story.append(P(
    "静态分析适合快速淘汰明显不兼容场景；动态分析用于判断最坏几何出现多久、多频繁，以及多个系统同时形成高耦合的概率。"
))

story.append(P("4. 总带宽足够，不等于频段可用", H2))
story.extend(bullets([
    "主要业务还是次要业务；",
    "是否受RR No. 9.21等协调程序约束；",
    "是否涉及生命安全、遇险通信或搜救系统；",
    "是否为全球统一划分，还是只在特定国家或区域可用；",
    "现有系统是否已高度拥挤；",
]))

story.append(P("5. 局部可行不等于全球可行", H2))
story.append(P(
    "某些固定地球站或固定探空仪接收站，在限定地理区域、限制功率和方向、保持足够隔离距离时可能实现局部共存。"
    "但移动、临时、全球分布或生命安全系统无法依赖同样的协调条件，因此不能把局部案例直接上升为全球划分结论。"
))

story.append(P("五、对干扰分析项目的工程启示", H1))
story.append(P(
    "尽管三篇报告研究的是1 GHz以下小卫星TT&C，而不是1200～1400 MHz天基雷达，"
    "其分析组织方式具有直接参考价值。可以提炼出以下标准流程："
))
workflow_data = [
    [P("① 系统参数", TABLE_CENTER), P("② 频谱与链路", TABLE_CENTER), P("③ 受扰准则", TABLE_CENTER), P("④ 静态筛查", TABLE_CENTER)],
    [P("轨道、功率、波形、带宽、方向图、占空比、接收机参数", TABLE_CENTER), P("发射谱、接收滤波、传播、链路预算", TABLE_CENTER), P("I/N、C/I、功率通量密度、超限时间或任务性能", TABLE_CENTER), P("最坏几何、MCL、带内/邻频/过载", TABLE_CENTER)],
    [P("⑤ 动态仿真", TABLE_CENTER), P("⑥ 聚合统计", TABLE_CENTER), P("⑦ 缓解设计", TABLE_CENTER), P("⑧ 结论边界", TABLE_CENTER)],
    [P("轨道、姿态、波束、可见窗口和时间步长", TABLE_CENTER), P("CDF、百分位、超限时长、多源同时出现", TABLE_CENTER), P("频率、功率、时序、波束、站址和协调", TABLE_CENTER), P("局部/全球、单源/聚合、示例值/项目值", TABLE_CENTER)],
]
wf = Table(workflow_data, colWidths=[43 * mm] * 4, hAlign="CENTER")
wf.setStyle(TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.5, LINE),
    ("BACKGROUND", (0, 0), (-1, 0), BLUE),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("BACKGROUND", (0, 2), (-1, 2), NAVY),
    ("TEXTCOLOR", (0, 2), (-1, 2), colors.white),
    ("BACKGROUND", (0, 1), (-1, 1), PALE_BLUE),
    ("BACKGROUND", (0, 3), (-1, 3), LIGHT_GREY),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ("TOPPADDING", (0, 0), (-1, -1), 7),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
]))
story.append(wf)
story.append(Spacer(1, 4 * mm))
story.append(P(
    "对1200～1400 MHz天基预警雷达项目，SA.2426式参数清单可用于建立雷达及干扰源数据库；"
    "SA.2425式动态容量分析可用于星座或多干扰源聚合；SA.2427式报告结构可用于组织法规、业务、"
    "静态链路、轨道动态、邻频、带外、缓解措施和剩余风险。"
))

story.append(P("六、结论速记", H1))
conclusion_table = [
    ["报告", "一句话总结"],
    ["SA.2426-0", "先把研究对象描述清楚：参数不完整，干扰分析就不可复核。"],
    ["SA.2425-0", "先算同频能承载多少系统，再由系统总数反推频谱，而不是简单带宽相乘。"],
    ["SA.2427-0", "频谱是否可用由法规地位、既有业务、传播和动态聚合共同决定，名义空闲带宽并不等于可分配资源。"],
]
story.append(make_table(conclusion_table, [32 * mm, 140 * mm]))
story.append(Spacer(1, 6 * mm))
story.append(box(
    "三篇报告真正值得借鉴的，是一套从“参数—需求—兼容性—缓解—结论边界”逐层推进的研究方法。"
    "对任何新的频谱项目，都应把示例参数与项目真实参数区分开，把单源结论与聚合结论区分开，把局部协调可行性与全球划分可行性区分开。",
    PALE_BLUE,
    NAVY,
))

story.append(PageBreak())
story.append(P("资料来源与使用说明", H1))
sources = [
    ["文件", "英文题名与用途"],
    ["Report ITU-R SA.2425-0 (2018)", "Studies to accommodate spectrum requirements in the space operation service for non-geostationary satellites with short duration missions；用于估算TT&C频谱需求。"],
    ["Report ITU-R SA.2426-0 (2018)", "Technical characteristics for telemetry, tracking and command in the space operation service below 1 GHz for non-GSO satellites with short duration missions；用于建立系统参数和链路预算。"],
    ["Report ITU-R SA.2427-0 (2018)", "Studies on the suitability of existing allocations to the space operation service below 1 GHz and additional sharing studies on possible new and/or upgraded allocations；用于评估现有及候选频段的共用兼容性。"],
]
story.append(make_table(sources, [52 * mm, 120 * mm]))
story.append(Spacer(1, 5 * mm))
story.append(P(
    "本总结保持三份原报告的研究逻辑和主要数值，不把报告未论证的内容补写成ITU结论。"
    "为便于阅读，部分表格和复杂研究案例进行了归纳；需要用于正式频率申报、协调或工程验收时，"
    "应依据原报告、现行《无线电规则》以及项目所在国家/地区的频率划分和许可条件重新核验。"
))

OUT.parent.mkdir(parents=True, exist_ok=True)
doc = NumberedDocTemplate(str(OUT))
doc.build(story)
print(f"Created {OUT} ({OUT.stat().st_size:,} bytes)")
