# -*- coding: utf-8 -*-
"""Shared project data and Word helpers for filing documents."""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

PROJECT_NAME = "山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目"
BID_NO = "ZLZX 招【2021】0194 号"
PROJECT_CODE = "A1400000007001110001"
OWNER = "山西神龙能源焦化有限责任公司"
AGENT = "山西重力工程咨询有限公司"
SUPERVISOR = "太原煤炭气化（集团）有限责任公司"

APPROVAL_ROWS = [
    ("项目名称", PROJECT_NAME),
    ("项目代码", PROJECT_CODE),
    ("项目单位", OWNER),
    ("服务地点", "山西省晋中市（招标人指定地点）"),
    ("资金来源", "企业自筹（100%），已落实"),
    ("核准/批复状态", "本项目已具备招标条件"),
    ("核准/批复机关", "待补充"),
    ("核准/批复文号", "待补充"),
    ("核准主要内容", "同意开展危险废物收集、运输、处置及设施清理清洗等服务采购"),
]

FILING_ROWS = [
    ("1", "建设单位", OWNER),
    ("2", "项目名称", "厂区废液等危险废物处置项目"),
    ("3", "招标编号", BID_NO),
    ("4", "建设性质", "服务采购（危险废物处置）"),
    ("5", "建设规模", "约800余吨废液及污泥等；滤袋1.5吨；废油约1吨"),
    ("6", "服务期限", "合同签订后45天内完成"),
    ("7", "招标代理机构", AGENT),
    ("8", "监督部门", SUPERVISOR),
    ("9", "备案号", "待补充"),
    ("10", "备案日期", "待补充"),
]


def set_cell_font(cell, text, size=10.5):
    cell.text = str(text)
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.name = "宋体"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            r.font.size = Pt(size)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.name = "黑体"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    r.font.size = Pt(18)
    r.bold = True
    doc.add_paragraph()


def add_meta(doc):
    for line in [
        f"项目名称：{PROJECT_NAME}",
        f"招标编号：{BID_NO}",
        f"投资项目代码：{PROJECT_CODE}",
        "编制说明：结构参考采购需求.doc；核准/备案栏参考核准批复及报建备案表。",
    ]:
        add_body(doc, line, indent=False)


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.name = "黑体"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")


def add_body(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = "宋体"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    r.font.size = Pt(12)
    r.bold = bold
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)


def add_table_2col(doc, rows):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Table Grid"
    for i, (a, b) in enumerate(rows):
        set_cell_font(t.rows[i].cells[0], a)
        set_cell_font(t.rows[i].cells[1], b)
    doc.add_paragraph()


def add_table_ncol(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for j, h in enumerate(headers):
        set_cell_font(t.rows[0].cells[j], h, 10)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            set_cell_font(t.rows[i + 1].cells[j], v, 10)
    doc.add_paragraph()


def add_common_sections(doc):
    add_heading(doc, "项目核准及批复信息", 1)
    add_table_2col(doc, APPROVAL_ROWS)
    add_heading(doc, "项目备案及报建信息登记表", 1)
    add_table_ncol(doc, ["序号", "栏目", "填写内容"], FILING_ROWS)


def new_filing_doc(title):
    doc = Document()
    add_title(doc, title)
    add_meta(doc)
    add_common_sections(doc)
    return doc
