# -*- coding: utf-8 -*-
"""Generate 采购需求 Word document from tender-derived content."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

OUT_DIR = Path(r"d:\Document\LAB\AI\test\test\备案文件")
OUT_DOCX = OUT_DIR / "采购需求.docx"
OUT_DOC = OUT_DIR / "采购需求.doc"


def set_cell_font(cell, text, size=10.5):
    cell.text = text
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


def add_table_2col(doc, rows):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Table Grid"
    for i, (a, b) in enumerate(rows):
        set_cell_font(t.rows[i].cells[0], a)
        set_cell_font(t.rows[i].cells[1], b)
    doc.add_paragraph()


def add_table_7col(doc, headers, data):
    t = doc.add_table(rows=2, cols=7)
    t.style = "Table Grid"
    for j, h in enumerate(headers):
        set_cell_font(t.rows[0].cells[j], h, 10)
    for j, v in enumerate(data):
        set_cell_font(t.rows[1].cells[j], v, 10)
    doc.add_paragraph()


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
    add_body(doc, "编制说明：本文档按《采购需求》标准章节结构编制，内容来源于招标文件。")

    add_heading(doc, "（一）需求清单", 1)

    add_heading(doc, "（一）项目概况", 2)
    add_body(
        doc,
        "山西神龙能源焦化有限责任公司厂区现存新事故池、新调节池、脱硫工段和粗苯工段废液收集池的约 800 余吨废液及污泥、滤袋 1.5 吨、废矿物油和废油桶约 1 吨等危险废物，需委托具备资质的单位完成收集、装卸、道路运输、处置及清理清洗（含池壁、涵洞管沟等）服务，并出具危险废物转移处置完结报告。项目资金来源为企业自筹，服务地点为招标人指定地点（山西省晋中市）。",
    )

    add_heading(doc, "（二）采购项目预（概）算", 2)
    add_table_2col(
        doc,
        [
            ("最高投标限价", "4000 元/吨（含税）"),
            ("计价方式", "处置服务费单价 × 实际过磅量"),
            ("预估处置规模", "约 800 余吨废液及污泥等（以出厂过磅为准）"),
            ("预估合同金额", "按实际过磅量结算，投标报价不得超过最高限价"),
        ],
    )

    add_heading(doc, "（三）采购标的汇总表", 2)
    add_table_7col(
        doc,
        ["包号", "序号", "标的名称", "品目分类", "计量单位", "数量", "是否进口"],
        [
            "001",
            "1",
            "山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目",
            "危险废物处置服务",
            "项",
            "1",
            "否",
        ],
    )

    add_heading(doc, "（四）技术商务要求", 2)
    add_heading(doc, "（1）技术要求", 3)
    add_body(doc, "标的：厂区废液等危险废物处置服务", bold=True)

    add_table_2col(
        doc,
        [
            ("采购人/委托人", "山西神龙能源焦化有限责任公司"),
            ("服务地点", "委托方厂区或指定地点（山西省晋中市）"),
            ("服务标的", "危险废物收集、装卸、运输、处置、清理清洗及完结报告编制"),
            ("主要废物来源", "新事故池、新调节池、脱硫工段、粗苯工段废液收集池等"),
            ("预估规模", "废液及污泥约 800 余吨；滤袋 1.5 吨；废矿物油及废油桶约 1 吨"),
            ("结算依据", "以神龙能源焦化公司出厂过磅数据为准"),
            ("危废类别", "HW11、HW49、HW08"),
            ("处置方式", "水泥窑协同处置（代码 C1）"),
            ("服务期限", "自合同签订后 45 天内完成全部处置并出具完结报告"),
            ("质量要求", "达到国家、行业相关规范、规定合格标准"),
            ("环保要求", "符合国家、山西省有关环保、安全、职业健康等法律法规及行业标准"),
        ],
    )

    add_body(doc, "服务内容一：处置技术服务", bold=True)
    add_body(
        doc,
        "对危险废物中有毒、有害物质进行定性、定量分析，根据其理化性质及危险特性，通过相应处置系统输送至水泥回转窑进行高温、无害化处置；按危险废物经营许可证核准方式实施处置；处置完毕后出具《神龙焦化危险废物转移处置完结》报告。",
    )

    add_body(doc, "服务内容二：运输服务", bold=True)
    add_body(
        doc,
        "采用具有《道路危险货物运输许可证》且经营范围包含本次危废类别的车辆运输；装车出厂后至合法处置完毕前的安全环保责任由受托方承担。",
    )

    add_body(doc, "服务内容三：装卸及清理清洗服务", bold=True)
    add_body(
        doc,
        "负责装车、卸车；清理清洗危废设施内部（含池壁、涵洞管沟等），达到发包人要求。",
    )

    add_body(doc, "投标人须提交的方案（纳入评标）：", bold=True)
    add_body(
        doc,
        "1. 处置计划及储存废液设施内部清理清洗方案；2. 应急预案、安全防护方案、运输方案；3. 日处理能力及贮存能力证明材料；4. 全过程质量管理和控制措施说明。",
    )

    add_body(doc, "特别约定：", bold=True)
    add_body(
        doc,
        "委托方负责转移联单申请及管理；转移废物须与检验合格准入样品相符；不得混入易制毒、剧毒、放射性、爆炸性物质；受托方对技术资料保密。",
    )

    add_body(doc, "五、其他", bold=True)
    add_body(
        doc,
        "投标人须按招标文件编制电子投标文件并参加线上开标；辅助纸质投标文件正本 1 份、副本 4 份；须保持关注山西新点招投标交易平台。",
    )

    add_heading(doc, "（2）商务要求", 3)
    add_body(doc, "（一）服务期限：自合同签订后 45 天内完成全部处置并出具完结报告。")
    add_body(doc, "（二）服务地点：招标人指定地点（山西省晋中市）。")
    add_body(doc, "（三）服务标准：达到国家、行业相关规范、规定合格标准。")
    add_body(doc, "（四）服务要求：", bold=True)
    add_body(
        doc,
        "1. 须具有省级《危险废物经营许可证》（HW11、HW49、HW08）；2. 须具有《道路危险货物运输许可证》；3. 近五年同类型业绩；4. 不允许分包；联合体牵头人须具备危废经营许可证；5. 投标报价不超过 4000 元/吨（含税）；6. 质保金预留不低于 10%。",
    )
    add_body(doc, "（五）付款方式：以合同约定为准；完成工作并出具完结报告后 5 日内支付（示范合同）。")
    add_body(doc, "（六）验收方式及验收标准：按合同约定及国家行业标准；验收成果为危险废物处置完结报告；结算以过磅为准。")
    add_body(
        doc,
        "（七）其他要求：投标保证金 60000 元；投标有效期 90 天；不要求履约保证金；须提供财务报告、社保及纳税凭证；信用记录符合要求。",
    )

    add_heading(doc, "综合评分法评分细则（摘要）", 1)
    add_body(doc, "评标办法：经评审的最低投标价法；技术 60 分 + 商务 40 分，折算计入评标基准价。")
    add_body(doc, "商务服务部分（40 分）：业绩 0～25 分；服务承诺 0～10 分；质保金承诺 0～5 分（最低 10%）。", bold=True)
    add_body(doc, "技术部分（60 分）：技术指标 0～15；设备计划 0～15；运输及贮存 0～10；应急预案 0～10；质量管理 0～10。", bold=True)
    add_body(doc, "价格部分：经评审投标价从低到高推荐中标候选人；最高限价 4000 元/吨（含税）。", bold=True)

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
        print(f"Note: .doc conversion skipped ({e}). Use .docx or open docx in Word and Save As .doc.")


if __name__ == "__main__":
    build_document()
