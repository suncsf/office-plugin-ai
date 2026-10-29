/**
 * 附件功能离线测试（不依赖 Office / 外网）
 *  1. docx/xlsx 解析（内容 + 格式）
 *  2. attachmentRead 各模式
 *  3. Agent 循环：attachment_* 服务端工具直接执行（不触发任务窗格中断）
 *  4. 历史消息附件标注
 */
import { zipSync, strToU8 } from "fflate";
import { parseDocument, attachmentRead, attachmentOverview } from "../server/docparse.js";
import { runAgentTurn, historyToLcMessages } from "../server/agent.js";
import http from "node:http";

let failed = false;
const check = (name, cond, detail = "") => {
  console.log(`${cond ? "✅" : "❌"} ${name}${cond || !detail ? "" : " " + detail}`);
  if (!cond) failed = true;
};
const j = (v) => JSON.stringify(v);

/* ---------------- 构造最小 docx ---------------- */
const docxXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="480"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="黑体" w:eastAsia="黑体"/><w:b/><w:color w:val="C00000"/><w:sz w:val="32"/></w:rPr><w:t>项目&lt;报告&gt;</w:t></w:r></w:p>
<w:p><w:r><w:t>这是正文第一段。</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>苹果</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
</w:body></w:document>`;
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`;
const docxBytes = zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "word/document.xml": strToU8(docxXml),
  "word/styles.xml": strToU8(stylesXml),
});

/* ---------------- 构造最小 xlsx ---------------- */
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:C3"/><cols><col min="1" max="1" width="18.5"/></cols>
<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" s="1"><v>0.25</v></c><c r="C2"><v>42</v></c></row>
</sheetData></worksheet>`;
const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="销售数据" sheetId="1" r:id="rId1"/></sheets></workbook>`;
const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2"><si><t>月份</t></si><si><t>一季度</t></si></sst>`;
const xStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="176" formatCode="0.00%"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="12"/><name val="微软雅黑"/></font></fonts>
<cellXfs count="2"><xf numFmtId="0" fontId="0"/><xf numFmtId="176" fontId="1"/></cellXfs>
</styleSheet>`;
const xlsxBytes = zipSync({
  "xl/workbook.xml": strToU8(wbXml),
  "xl/_rels/workbook.xml.rels": strToU8(wbRels),
  "xl/sharedStrings.xml": strToU8(sstXml),
  "xl/styles.xml": strToU8(xStylesXml),
  "xl/worksheets/sheet1.xml": strToU8(sheetXml),
});

console.log("\n[docx 解析]");
{
  const p = parseDocument("测试.docx", docxBytes);
  check("类型与段落计数", p.type === "docx" && p.paragraphCount === 2, j({ type: p.type, n: p.paragraphCount }));
  const h = p.paragraphs[0];
  check("标题段样式解析", h.style === "heading 1", j(h.style));
  check("对齐/大纲", h.align === "Centered");
  check("字体解析（名称/字号/加粗/颜色）", h.font?.name === "黑体" && h.font?.sizePt === 16 && h.font?.bold === true && h.font?.color === "#C00000", j(h.font));
  check("间距与缩进", h.spacing?.beforePt === 12 && h.spacing?.lineSpacing === "1.50 倍行距" && h.indent?.firstLinePt === 24, j({ s: h.spacing, i: h.indent }));
  check("XML 实体反转义", h.text === "项目<报告>", j(h.text));
  check("表格解析", p.tables.length === 1 && j(p.tables[0].rows) === j([["名称", "数量"], ["苹果", "3"]]), j(p.tables));
}

console.log("\n[xlsx 解析]");
{
  const p = parseDocument("数据.xlsx", xlsxBytes);
  check("工作表定位（rid→路径）", p.sheets?.[0]?.name === "销售数据", j(p.sheets?.map((s) => s.name)));
  const rows = p.sheets[0].rows;
  check("共享字符串与数值", rows[0].cells[0].v === "月份" && rows[1].cells[2].v === "42", j(rows));
  const b2 = rows[1].cells[1];
  check("单元格格式（数字格式+加粗+字体）", b2.numFmt === "0.00%" && b2.bold === true && b2.font?.name === "微软雅黑", j(b2));
  check("合并单元格与列宽", j(p.sheets[0].merges) === j(["A1:C1"]) && p.sheets[0].colWidths?.[0]?.widthChars === 18.5);
  check("dimension", p.sheets[0].ref === "A1:C3");
}

console.log("\n[attachmentRead 模式]");
{
  const p = parseDocument("测试.docx", docxBytes);
  check("structure 返回标题大纲", attachmentRead(p, { mode: "structure" }).headings[0].text === "项目<报告>");
  check("text 不含格式详情", !("font" in attachmentRead(p, { mode: "text", offset: 0, limit: 10 }).paragraphs[0]));
  check("format 含字体详情", attachmentRead(p, { mode: "format" }).paragraphs[0].font?.name === "黑体");
  check("tables 返回表格", attachmentRead(p, { mode: "tables" }).tables[0].rows[0][0] === "名称");
  check("概览（attachment_list 用）", attachmentOverview(p).paragraphCount === 2 && attachmentOverview(p).tableCount === 1);
  const t = parseDocument("说明.txt", Buffer.from("第一行\n第二行", "utf8"));
  check("纯文本", t.type === "text" && attachmentRead(t, {}).text.includes("第二行"));
}

console.log("\n[Agent 循环：服务端工具直执行]（假 OpenAI SSE 服务 + 真实 ChatOpenAI）");
{
  let callCount = 0;
  const chunk = (delta, finish = null) =>
    `data: ${JSON.stringify({ id: "c1", object: "chat.completion.chunk", created: 1, model: "m", choices: [{ index: 0, delta, finish_reason: finish }] })}\n\n`;
  const server = http.createServer((req, res) => {
    let body = "";
    req.on("data", (d) => (body += d));
    req.on("end", () => {
      callCount++;
      res.writeHead(200, { "content-type": "text/event-stream" });
      if (callCount === 1) {
        res.write(chunk({ role: "assistant", tool_calls: [{ index: 0, id: "call_1", type: "function", function: { name: "attachment_list", arguments: "{}" } }] }));
        res.write(chunk({ role: "assistant", tool_calls: [{ index: 1, id: "call_2", type: "function", function: { name: "attachment_read", arguments: JSON.stringify({ id: "att_x1", mode: "format" }) } }] }));
        res.write(chunk({}, "tool_calls"));
      } else {
        res.write(chunk({ role: "assistant", content: "已按附件格式了解完毕" }));
        res.write(chunk({}, "stop", { usage: { prompt_tokens: 9, completion_tokens: 3, total_tokens: 12 } }));
      }
      res.write("data: [DONE]\n\n");
      res.end();
    });
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const port = server.address().port;

  const parsed = parseDocument("测试.docx", docxBytes);
  const atts = { att_x1: { id: "att_x1", conversationId: "conv_test", name: "测试.docx", size: docxBytes.length } };
  const executor = async (name, args) => {
    if (name === "attachment_list") return { count: 1, attachments: [{ ...atts.att_x1, overview: attachmentOverview(parsed) }] };
    if (name === "attachment_read") {
      const att = atts[String(args?.id)];
      if (!att) throw new Error("附件不存在");
      return attachmentRead(parsed, args);
    }
    throw new Error(`未知的服务端工具：${name}`);
  };

  const messages = [];
  let interruptCalled = false;
  const { createChatModel } = await import("../server/agent.js");
  const modelCfg = { name: "fake", model: "test-model", baseUrl: `http://127.0.0.1:${port}/v1`, apiKey: "test", contextTokens: 1000, supportsTools: true, supportsImages: false, streamUsage: false };
  await runAgentTurn({
    turnId: `att_${Date.now()}`,
    historyRows: [{ role: "user", content: "读取附件", images: [], toolCalls: [], meta: { attachments: [{ id: "att_x1", name: "测试.docx" }] } }],
    model: modelCfg,
    general: { maxToolTurns: 5 },
    appType: "word",
    modelInstance: createChatModel(modelCfg, { maxToolTurns: 5 }),
    onDelta: () => {},
    onMessage: (m) => messages.push(m),
    onInterrupt: () => { interruptCalled = true; },
    waitForResults: async () => { interruptCalled = true; return []; },
    onUsage: () => {},
    serverToolExecutor: executor,
  });
  server.close();

  const tools = messages.filter((m) => m.role === "tool");
  check("未触发任务窗格中断（服务端直执行）", interruptCalled === false);
  check("两条工具结果都有落库形态", tools.length === 2, j(messages.map((m) => m.role)));
  const readMsg = tools.find((m) => m.name === "attachment_read");
  check("attachment_read 返回内容与格式", (readMsg?.content || "").includes("项目<报告>") && (readMsg?.content || "").includes("黑体"), (readMsg?.content || "").slice(0, 120));
  check("attachment_list 返回概览", (tools[0]?.content || "").includes("paragraphCount"));
  check("最终回复正常", messages.some((m) => m.role === "assistant" && m.content === "已按附件格式了解完毕"));
}

console.log("\n[历史消息附件标注]");
{
  const rows = [{ role: "user", content: "帮我排版", images: [], toolCalls: [], meta: { attachments: [{ id: "att_x1", name: "测试.docx" }] } }];
  const lc = historyToLcMessages(rows, { supportsImages: false });
  const text = typeof lc[0].content === "string" ? lc[0].content : lc[0].content[0].text;
  check("标注包含附件 id 与读取指引", text.includes("测试.docx") && text.includes("att_x1") && text.includes("attachment_read"), text.slice(-80));
  const noAtt = historyToLcMessages([{ role: "user", content: "无附件", images: [], toolCalls: [] }], { supportsImages: false });
  check("无附件不加标注", noAtt[0].content === "无附件");
}

console.log(failed ? "\n❌ 存在失败项" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
