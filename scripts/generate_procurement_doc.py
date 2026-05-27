# -*- coding: utf-8 -*-
"""Generate 采购需求 Word document - merged template from 3 reference files."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

OUT_DIR = Path(r"d:\Document\LAB\AI\test\test\备案文件")
OUT_DOCX = OUT_DIR / "采购需求.docx"
OUT_DOC = OUT_DIR / "采购需求.doc"


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


def add_body(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(12)
    run.bold = bold
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    return p


def add_table_ncol(doc, headers, rows):
    cols = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=cols)
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

    add_body(doc, "项目名称：山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目")
    add_body(doc, "招标编号：ZLZX 招【2021】0194 号")
    add_body(doc, "投资项目代码：A1400000007001110001")
    add_body(doc, "编制说明：结构参考《采购需求.doc》；核准/备案栏目参考核准批复及报建备案表；内容来源于招标文件。")

    add_heading(doc, "项目核准及批复信息", 1)
    add_table_2col(doc, [
        ("项目名称", "山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目"),
        ("项目代码", "A1400000007001110001"),
        ("项目单位", "山西神龙能源焦化有限责任公司"),
        ("服务地点", "山西省晋中市（招标人指定地点）"),
        ("资金来源", "企业自筹（100%），已落实"),
        ("核准/批复状态", "本项目已具备招标条件"),
        ("核准/批复机关", "待补充"),
        ("核准/批复文号", "待补充"),
        ("核准主要内容", "同意开展危险废物收集、运输、处置及设施清理清洗等服务采购"),
    ])

    add_heading(doc, "项目备案及报建信息登记表", 1)
    add_table_ncol(doc, ["序号", "栏目", "填写内容"], [
        ("1", "建设单位", "山西神龙能源焦化有限责任公司"),
        ("2", "项目名称", "厂区废液等危险废物处置项目"),
        ("3", "建设性质", "服务采购"),
        ("4", "建设规模", "约800余吨废液及污泥等；滤袋1.5吨；废油约1吨"),
        ("5", "预算口径", "单价×过磅量；最高限价4000元/吨（含税）"),
        ("6", "服务期限", "合同签订后45天内完成"),
        ("7", "招标代理", "山西重力工程咨询有限公司"),
        ("8", "监督部门", "太原煤炭气化（集团）有限责任公司"),
        ("9", "备案号", "待补充"),
    ])

    add_heading(doc, "（一）需求清单", 1)
    add_heading(doc, "（一）项目概况", 2)
    add_body(doc, "本项目选择具备危险废物经营及危废道路运输资质的服务单位，提供收集、装卸、运输、水泥窑协同处置（C1）、设施清理清洗及完结报告等环境服务。焦化厂现存新事故池、新调节池、脱硫及粗苯工段废液收集池约800余吨废液及污泥、滤袋1.5吨、废矿物油及废油桶约1吨等（以出厂过磅为准）。")

    add_heading(doc, "（二）采购项目预（概）算", 2)
    add_body(doc, "预算金额：按实际过磅量×中标单价计算")
    add_body(doc, "最高限价：肆仟元整/吨（￥4000.00/吨，含税）")
    add_table_2col(doc, [
        ("计价方式", "处置服务费单价×实际过磅量"),
        ("预估规模", "约800余吨"),
        ("投标保证金", "60000元"),
    ])

    add_heading(doc, "（三）采购标的汇总表", 2)
    add_table_ncol(doc, ["包号", "序号", "标的名称", "品目分类", "计量单位", "数量", "是否进口"], [
        ("001", "1", "厂区废液等危险废物处置项目", "危险废物处置服务", "项", "1", "否"),
    ])

    add_heading(doc, "（四）技术商务要求", 2)
    add_heading(doc, "（1）技术要求", 3)

    for title_text, rows in [
        ("服务内容一：危险废物处置技术服务", [
            ("服务条款", "危险废物无害化处置"),
            ("采购人", "山西神龙能源焦化有限责任公司"),
            ("危废类别", "HW11、HW49、HW08"),
            ("处置方式", "水泥窑协同处置（C1）"),
            ("最高限价", "4000元/吨（含税）"),
            ("服务期限", "合同签订后45天内"),
        ]),
        ("服务内容二：道路运输服务", [
            ("资质要求", "道路危险货物运输许可证"),
            ("责任", "装车出厂至合法处置完毕前安全环保责任"),
        ]),
        ("服务内容三：装卸及清理清洗", [
            ("范围", "装车卸车；清理池壁、涵洞管沟等"),
        ]),
    ]:
        add_body(doc, title_text, bold=True)
        add_table_2col(doc, rows)

    add_body(doc, "投标人须提交：处置计划、清理清洗方案、应急预案、安全防护方案、运输方案、日处理及贮存能力证明、质量管理措施。", bold=True)

    add_heading(doc, "（2）商务要求", 3)
    for line in [
        "（一）服务期限：合同签订后45天内完成并出具完结报告。",
        "（二）服务地点：山西省晋中市，招标人指定地点。",
        "（三）服务标准：国家、行业合格标准。",
        "（四）须具备危废经营许可证、危运许可证、近五年业绩；不接受分包；质保金最低10%。",
        "（五）付款：完成工作并出具报告后5日内支付（以合同为准）。",
        "（六）验收：以处置完结报告及过磅数据为准。",
        "（七）投标保证金60000元；有效期90天；不要求履约保证金。",
    ]:
        add_body(doc, line)

    add_heading(doc, "综合评分法评分细则", 1)
    add_table_ncol(doc, ["评审内容", "评审因素", "分值", "评审标准"], [
        ("技术60分", "技术指标响应", "0-15", "合理完整符合国标"),
        ("技术60分", "设备投入计划", "0-15", "方案完善程度"),
        ("技术60分", "运输及贮存计划", "0-10", "能力与方案"),
        ("技术60分", "应急预案", "0-10", "针对本次处置物品"),
        ("技术60分", "质量管理", "0-10", "描述详细程度"),
        ("商务40分", "业绩", "0-25", "每项5分最多25分"),
        ("商务40分", "服务承诺", "0-10", "排序评分"),
        ("商务40分", "质保金", "0-5", "最低10%"),
        ("价格", "经评审投标价", "—", "从低到高；限价4000元/吨"),
    ])
    add_body(doc, "五、其他：投标人须保持登录山西新点招投标交易平台，及时响应评标询问。")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUT_DOCX)
    print(f"Created: {OUT_DOCX}")

    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        wdoc = word.Documents.Open(str(OUT_DOCX.resolve()))
        wdoc.SaveAs(str(OUT_DOC.resolve()), FileFormat=0)
        wdoc.Close()
        word.Quit()
        print(f"Created: {OUT_DOC}")
    except Exception as e:
        print(f".doc skip: {e}")


if __name__ == "__main__":
    build_document()
