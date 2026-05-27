# Generate 采购需求.doc from structured content via Word COM
$ErrorActionPreference = "Stop"
$outDir = "d:\Document\LAB\AI\test\test\备案文件"
$outDoc = Join-Path $outDir "采购需求.doc"
$outDocx = Join-Path $outDir "采购需求.docx"

function Add-Heading($doc, $text, $level) {
    $p = $doc.Paragraphs.Add()
    $p.Range.Text = $text
    $p.Range.Style = "Heading $level"
    $p.Range.Font.Name = "宋体"
    $p.Range.InsertParagraphAfter()
}

function Add-Body($doc, $text, $bold = $false) {
    $p = $doc.Paragraphs.Add()
    $p.Range.Text = $text
    $p.Range.Font.Name = "宋体"
    $p.Range.Font.Size = 12
    $p.Range.Font.Bold = $bold
    $p.Range.ParagraphFormat.FirstLineIndent = 24
    $p.Range.InsertParagraphAfter()
}

function Add-Table2Col($doc, $rows) {
    $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1), $rows.Count, 2)
    $t.Range.Font.Name = "宋体"
    $t.Range.Font.Size = 11
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $t.Cell($i + 1, 1).Range.Text = $rows[$i][0]
        $t.Cell($i + 1, 2).Range.Text = $rows[$i][1]
    }
    $t.Borders.Enable = $true
    $doc.Content.InsertParagraphAfter()
}

function Add-Table7Col($doc, $headers, $dataRow) {
    $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1), 2, 7)
    $t.Range.Font.Name = "宋体"
    $t.Range.Font.Size = 10
    for ($c = 1; $c -le 7; $c++) { $t.Cell(1, $c).Range.Text = $headers[$c - 1] }
    for ($c = 1; $c -le 7; $c++) { $t.Cell(2, $c).Range.Text = $dataRow[$c - 1] }
    $t.Borders.Enable = $true
    $doc.Content.InsertParagraphAfter()
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()

    # Title
    $p = $doc.Paragraphs.Add()
    $p.Range.Text = "采购需求"
    $p.Range.Font.Name = "黑体"
    $p.Range.Font.Size = 22
    $p.Range.Font.Bold = $true
    $p.Alignment = 1
    $p.Range.InsertParagraphAfter()

    Add-Body $doc "项目名称：山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目"
    Add-Body $doc "招标编号：ZLZX 招【2021】0194 号"
    Add-Body $doc "编制说明：本文档按《采购需求》标准章节结构编制，内容来源于招标文件。"

    Add-Heading $doc "（一）需求清单" 1

    Add-Heading $doc "（一）项目概况" 2
    Add-Body $doc "山西神龙能源焦化有限责任公司厂区现存新事故池、新调节池、脱硫工段和粗苯工段废液收集池的约 800 余吨废液及污泥、滤袋 1.5 吨、废矿物油和废油桶约 1 吨等危险废物，需委托具备资质的单位完成收集、装卸、道路运输、处置及清理清洗（含池壁、涵洞管沟等）服务，并出具危险废物转移处置完结报告。项目资金来源为企业自筹，服务地点为招标人指定地点（山西省晋中市）。"

    Add-Heading $doc "（二）采购项目预（概）算" 2
    Add-Table2Col $doc @(
        @("最高投标限价", "4000 元/吨（含税）"),
        @("计价方式", "处置服务费单价 × 实际过磅量"),
        @("预估处置规模", "约 800 余吨废液及污泥等（以出厂过磅为准）"),
        @("预估合同金额", "按实际过磅量结算，投标报价不得超过最高限价")
    )

    Add-Heading $doc "（三）采购标的汇总表" 2
    Add-Table7Col $doc @("包号", "序号", "标的名称", "品目分类", "计量单位", "数量", "是否进口") @(
        "001", "1", "山西神龙能源焦化有限责任公司厂区废液等危险废物处置项目", "危险废物处置服务", "项", "1", "否"
    )

    Add-Heading $doc "（四）技术商务要求" 2
    Add-Heading $doc "（1）技术要求" 3
    Add-Body $doc "标的：厂区废液等危险废物处置服务" $true

    Add-Table2Col $doc @(
        @("采购人/委托人", "山西神龙能源焦化有限责任公司"),
        @("服务地点", "委托方厂区或指定地点（山西省晋中市）"),
        @("服务标的", "危险废物收集、装卸、运输、处置、清理清洗及完结报告编制"),
        @("主要废物来源", "新事故池、新调节池、脱硫工段、粗苯工段废液收集池等"),
        @("预估规模", "废液及污泥约 800 余吨；滤袋 1.5 吨；废矿物油及废油桶约 1 吨"),
        @("结算依据", "以神龙能源焦化公司出厂过磅数据为准"),
        @("危废类别", "HW11、HW49、HW08"),
        @("处置方式", "水泥窑协同处置（代码 C1）"),
        @("服务期限", "自合同签订后 45 天内完成全部处置并出具完结报告"),
        @("质量要求", "达到国家、行业相关规范、规定合格标准"),
        @("环保要求", "符合国家、山西省有关环保、安全、职业健康等法律法规及行业标准")
    )

    Add-Body $doc "服务内容一：处置技术服务" $true
    Add-Body $doc "对危险废物中有毒、有害物质进行定性、定量分析，根据其理化性质及危险特性，通过相应处置系统输送至水泥回转窑进行高温、无害化处置；按危险废物经营许可证核准方式实施处置；处置完毕后出具《神龙焦化危险废物转移处置完结》报告。"

    Add-Body $doc "服务内容二：运输服务" $true
    Add-Body $doc "采用具有《道路危险货物运输许可证》且经营范围包含本次危废类别的车辆运输；装车出厂后至合法处置完毕前的安全环保责任由受托方承担。"

    Add-Body $doc "服务内容三：装卸及清理清洗服务" $true
    Add-Body $doc "负责装车、卸车；清理清洗危废设施内部（含池壁、涵洞管沟等），达到发包人要求。"

    Add-Body $doc "投标人须提交的方案（纳入评标）：" $true
    Add-Body $doc "1. 处置计划及储存废液设施内部清理清洗方案；2. 应急预案、安全防护方案、运输方案；3. 日处理能力及贮存能力证明材料；4. 全过程质量管理和控制措施说明。"

    Add-Body $doc "特别约定：" $true
    Add-Body $doc "委托方负责转移联单申请及管理；转移废物须与检验合格准入样品相符；不得混入易制毒、剧毒、放射性、爆炸性物质；受托方对技术资料保密。"

    Add-Body $doc "五、其他" $true
    Add-Body $doc "投标人须按招标文件编制电子投标文件并参加线上开标；辅助纸质投标文件正本 1 份、副本 4 份；须保持关注山西新点招投标交易平台。"

    Add-Heading $doc "（2）商务要求" 3
    Add-Body $doc "（一）服务期限：自合同签订后 45 天内完成全部处置并出具完结报告。"
    Add-Body $doc "（二）服务地点：招标人指定地点（山西省晋中市）。"
    Add-Body $doc "（三）服务标准：达到国家、行业相关规范、规定合格标准。"
    Add-Body $doc "（四）服务要求：" $true
    Add-Body $doc "1. 须具有省级《危险废物经营许可证》（HW11、HW49、HW08）；2. 须具有《道路危险货物运输许可证》；3. 近五年同类型业绩；4. 不允许分包；联合体牵头人须具备危废经营许可证；5. 投标报价不超过 4000 元/吨（含税）；6. 质保金预留不低于 10%。"
    Add-Body $doc "（五）付款方式：以合同约定为准；完成工作并出具完结报告后 5 日内支付（示范合同）。"
    Add-Body $doc "（六）验收方式及验收标准：按合同约定及国家行业标准；验收成果为危险废物处置完结报告；结算以过磅为准。"
    Add-Body $doc "（七）其他要求：投标保证金 60000 元；投标有效期 90 天；不要求履约保证金；须提供财务报告、社保及纳税凭证；信用记录符合要求。"

    Add-Heading $doc "综合评分法评分细则（摘要）" 1
    Add-Body $doc "评标办法：经评审的最低投标价法；技术 60 分 + 商务 40 分，折算计入评标基准价。"

    Add-Body $doc "商务服务部分（40 分）：业绩 0～25 分；服务承诺 0～10 分；质保金承诺 0～5 分（最低 10%）。" $true
    Add-Body $doc "技术部分（60 分）：技术指标 0～15；设备计划 0～15；运输及贮存 0～10；应急预案 0～10；质量管理 0～10。" $true
    Add-Body $doc "价格部分：经评审投标价从低到高推荐中标候选人；最高限价 4000 元/吨（含税）。" $true

    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    # wdFormatDocument = 0 (.doc), wdFormatXMLDocument = 12 (.docx)
    $doc.SaveAs([ref]$outDoc, [ref]0)
    $doc.SaveAs([ref]$outDocx, [ref]12)
    $doc.Close()
    Write-Output "Created: $outDoc"
    Write-Output "Created: $outDocx"
}
finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
