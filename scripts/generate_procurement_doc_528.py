# -*- coding: utf-8 -*-
"""Generate 采购需求 Word for 项目528-工程施建."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

OUT_DIR = Path(r"d:\Document\LAB\AI\test\test\备案文件\项目528-工程施建")
OUT_DOCX = OUT_DIR / "采购需求.docx"

PROJECT = "古交市汾河下游三期河道治理工程（2 标段）"
BID_NO = "CEITCL-SX-ZCGC-200119"
APPROVAL = "古发改审批字【2019】34 号"
OWNER = "古交市基建水务工程项目建设管理有限公司"
AGENT = "中经国际招标集团有限公司"
SUPERVISOR = "古交市水务局"


def set_cell_font(cell, text, size=10.5):
    cell.text = str(text)
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.name = "宋体"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            r.font.size = Pt(size)


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.name = "黑体"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    return h


def add_body(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(12)
    run.bold = bold
    p.paragraph_format.first_line_indent = Cm(0.74)
    return p


def add_table_ncol(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for j, h in enumerate(headers):
        set_cell_font(t.rows[0].cells[j], h, 10)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            set_cell_font(t.rows[i + 1].cells[j], v, 10)
    doc.add_paragraph()


def add_table_2col(doc, rows):
    add_table_ncol(doc, ["项目", "内容"], rows)


def build_document():
    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("采购需求")
    tr.font.name = "黑体"
    tr._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    tr.font.size = Pt(22)
    tr.bold = True

    add_body(doc, f"项目名称：{PROJECT}")
    add_body(doc, f"招标编号：{BID_NO}")
    add_body(doc, f"投资项目批复：{APPROVAL}")
    add_body(doc, "编制说明：结构参考采购需求模板；内容来源于项目528招标文件正文。")

    add_heading(doc, "项目核准及批复信息", 1)
    add_table_2col(doc, [
        ("项目名称", "古交市汾河下游三期河道治理工程"),
        ("本次采购标段", "002标段（桩号C9+050～C9+443）"),
        ("项目单位", OWNER),
        ("建设地点", "汾河古交市段（火山引黄桥至古交市第二污水厂段）"),
        ("资金来源", "上级投资及古交自筹，已落实"),
        ("核准/批复文号", APPROVAL),
        ("核准主要内容", "同意实施河道治理，全长6.5km，防洪20年一遇，工程等别Ⅳ等"),
    ])

    add_heading(doc, "项目备案及报建信息登记表", 1)
    add_table_ncol(doc, ["序号", "栏目", "填写内容"], [
        ("1", "建设单位", OWNER),
        ("2", "项目名称", PROJECT),
        ("3", "建设性质", "水利工程施工"),
        ("4", "建设规模", "本标段C9+050～C9+443；全项目清淤6.5km及堤防、钢坝等"),
        ("5", "全项目投资", "约22470.77万元"),
        ("6", "计划工期", "16个月"),
        ("7", "招标代理", AGENT),
        ("8", "监督部门", SUPERVISOR),
        ("9", "联系人", "周先生 13934665458"),
    ])

    add_heading(doc, "（一）需求清单", 1)
    add_heading(doc, "（一）项目概况", 2)
    add_body(
        doc,
        "汾河下游三期河道治理工程治理总长6.5km，防洪标准20年一遇，质量达到SL176-2007合格标准。"
        "本次2标段实施桩号C9+050～C9+443段，范围包括施工图纸及工程量清单全部内容。",
    )

    add_heading(doc, "（二）采购项目预（概）算", 2)
    add_table_2col(doc, [
        ("全项目投资", "约22470.77万元"),
        ("招标控制价", "以招标文件及清单公布值为准"),
        ("投标保证金", "20万元"),
        ("招标文件售价", "1000元/标段；图纸1000元"),
    ])

    add_heading(doc, "（三）采购标的汇总表", 2)
    add_table_ncol(
        doc,
        ["包号", "序号", "标的名称", "品目分类", "计量单位", "数量", "是否进口"],
        [("002", "1", "古交市汾河下游三期河道治理工程2标段", "水利水电工程施工", "项", "1", "否")],
    )

    add_heading(doc, "（四）技术商务要求", 2)
    add_heading(doc, "（1）技术要求", 3)
    add_table_2col(doc, [
        ("招标范围", "施工图纸及工程量清单内全部工程"),
        ("计划工期", "16个月"),
        ("质量要求", "SL176-2007合格标准"),
        ("企业资质", "水利水电施工总承包贰级及以上或河湖整治贰级及以上"),
        ("项目经理", "水利水电贰级及以上建造师+安全考核证"),
        ("技术负责人", "水利水电工程建筑或施工专业工程师职称"),
        ("分包/联合体", "不允许分包；不接受联合体"),
    ])

    add_heading(doc, "（2）商务要求", 3)
    for line in [
        "投标有效期90日历天；保证金20万元（基本账户）。",
        "全流程电子投标：山西新点交易平台+山西CA。",
        "报价不得超过招标控制价；近五年类似业绩及ISO9001认证。",
        "付款及验收按合同条款及水利验收规程执行。",
    ]:
        add_body(doc, line)

    add_heading(doc, "综合评估法评分细则", 1)
    add_table_ncol(doc, ["评审大类", "因素", "分值", "标准摘要"], [
        ("投标报价", "投标总价", "45", "复合基准价±扣分规则"),
        ("施工组织", "方案/质量/安全/环保/进度/资源", "35", "分项打分"),
        ("管理机构", "主要岗位人员", "5", "各1分"),
        ("其他", "业绩/荣誉/财务/信用", "15", "见招标文件第三章"),
    ])
    add_body(doc, "监督部门：古交市水务局。平台：山西新点交易平台。")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUT_DOCX)
    print(f"Created: {OUT_DOCX}")


if __name__ == "__main__":
    build_document()
