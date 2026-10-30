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
<dimension ref="A1:D3"/><cols><col min="1" max="1" width="18.5"/></cols>
<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" s="1"><v>0.25</v></c><c r="C2"><v>42</v></c><c r="D2" s="2"><v>45413</v></c></row>
<row r="3"><c r="A3" s="3" t="s"><v>2</v></c><c r="B3"><f>B2*100</f><v>25</v></c></row>
</sheetData></worksheet>`;
const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="销售数据" sheetId="1" r:id="rId1"/></sheets></workbook>`;
const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3"><si><t>月份</t></si><si><t>一季度</t></si><si><t>汇总</t></si></sst>`;
const xStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="176" formatCode="0.00%"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="等线"/></font><font><b/><sz val="12"/><name val="微软雅黑"/></font><font><i/><color rgb="FF7030A0"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>
<cellXfs count="4"><xf numFmtId="0" fontId="0"/><xf numFmtId="176" fontId="1"/><xf numFmtId="14" fontId="0"/><xf numFmtId="0" fontId="2" fillId="1" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs>
</styleSheet>`;
const xlsxBytes = zipSync({
  "xl/workbook.xml": strToU8(wbXml),
  "xl/_rels/workbook.xml.rels": strToU8(wbRels),
  "xl/sharedStrings.xml": strToU8(sstXml),
  "xl/styles.xml": strToU8(xStylesXml),
  "xl/worksheets/sheet1.xml": strToU8(sheetXml),
});

/* ---------------- 构造最小 pptx ---------------- */
const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="等线"/><a:cs typeface=""/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;
const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="8000000" cy="1300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
<a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="4400" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="黑体"/><a:ea typeface="黑体"/></a:rPr><a:t>项目汇报</a:t></a:r></a:p>
<a:p><a:r><a:rPr lang="zh-CN" sz="2400"><a:latin typeface="+mn-lt"/></a:rPr><a:t>副标题行</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="内容占位符 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="3500000"/><a:ext cx="8000000" cy="3000000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr lang="zh-CN" sz="2000"/><a:t>要点一</a:t></a:r></a:p>
<a:p><a:r><a:rPr lang="zh-CN" sz="2000"/><a:t>要点二</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="表格 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="1000000" y="500000"/><a:ext cx="6000000" cy="1200000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="3000000"/></a:tblGrid>
<a:tr h="500000"><a:tc><a:txBody><a:p><a:r><a:t>名称</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>数量</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
<a:tr h="500000"><a:tc><a:txBody><a:p><a:r><a:t>叉车</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>3</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>
<p:pic><p:nvPicPr><p:cNvPr id="5" name="图片 4"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr><a:xfrm><a:off x="7000000" y="500000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="图表 5"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="2000000" y="6000000"/><a:ext cx="5000000" cy="3000000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId4"/></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>`;
const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="标题幻灯片"><p:spTree/></p:cSld></p:sldLayout>`;
const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>季度销售</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:plotArea><c:barChart>
<c:ser><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>销售额</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>一季度</c:v></c:pt><c:pt idx="1"><c:v>二季度</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>200</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
const notesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:sp><p:txBody><a:p><a:r><a:t>备注内容</a:t></a:r></a:p></p:txBody></p:sp></p:cSld></p:notes>`;
const pptxBytes = zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "ppt/theme/theme1.xml": strToU8(themeXml),
  "ppt/slides/slide1.xml": strToU8(slideXml),
  "ppt/slides/_rels/slide1.xml.rels": strToU8(slideRels),
  "ppt/slideLayouts/slideLayout1.xml": strToU8(layoutXml),
  "ppt/charts/chart1.xml": strToU8(chartXml),
  "ppt/notesSlides/notesSlide1.xml": strToU8(notesXml),
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
  check("共享字符串与数值类型化", rows[0].cells[0].v === "月份" && rows[1].cells[2].v === 42 && typeof rows[1].cells[2].v === "number", j(rows[1].cells));
  const b2 = rows[1].cells[1];
  check("单元格格式（数字格式+加粗+字体）", b2.numFmt === "0.00%" && b2.font?.bold === true && b2.font?.name === "微软雅黑", j(b2));
  const d2 = rows[1].cells[3];
  check("日期序列值换算（builtin 14）", d2.v === "2024-05-01" && d2.raw === 45413 && d2.numFmt === "yyyy-mm-dd", j(d2));
  const b3 = rows[2].cells[1];
  check("公式解析", b3.f === "=B2*100" && b3.v === 25, j(b3));
  const a3 = rows[2].cells[0];
  check("样式（斜体/颜色/填充/边框/对齐/换行）", a3.font?.italic === true && a3.font?.color === "#7030A0" && a3.fill === "#FF0000" && a3.border === "thin #BFBFBF" && a3.align?.h === "center" && a3.align?.v === "center" && a3.align?.wrap === true, j(a3));
  check("合并单元格与列宽", j(p.sheets[0].merges) === j(["A1:C1"]) && p.sheets[0].colWidths?.[0]?.widthChars === 18.5);
  check("dimension", p.sheets[0].ref === "A1:D3");
}

console.log("\n[pptx 解析]");
{
  const p = parseDocument("演示.pptx", pptxBytes);
  check("幻灯片与版式名", p.type === "pptx" && p.slideCount === 1 && p.slides[0].layout === "标题幻灯片", j({ type: p.type, layout: p.slides[0]?.layout }));
  const sp = p.slides[0].shapes;
  check("形状数量与类型（2 文本+表格+图片+图表）", sp.length === 5 && sp.filter((s) => s.kind === "text").length === 2 && sp.some((s) => s.kind === "table") && sp.some((s) => s.kind === "picture") && sp.some((s) => s.kind === "chart"), j(sp.map((s) => s.kind)));
  const title = sp[0];
  check("形状名/占位符/位置尺寸（EMU→磅）", title.name === "标题 1" && title.ph === "ctrTitle" && title.x === 78.7 && title.y === 157.5 && title.w === 629.9 && title.h === 102.4, j({ x: title.x, y: title.y, w: title.w, h: title.h }));
  check("填充色（主题色+lumMod）与边框色线宽", title.fill === "#BFBFBF" && title.line?.color === "#4472C4" && title.line?.wPt === 1, j({ fill: title.fill, line: title.line }));
  check("段落文字格式（字体/字号/加粗/颜色/对齐）", title.paragraphs[0] === "项目汇报" && title.paraFonts[0].name === "黑体" && title.paraFonts[0].sizePt === 44 && title.paraFonts[0].bold === true && title.paraFonts[0].color === "#C00000" && title.paraFonts[0].align === "Center", j(title.paraFonts[0]));
  check("主题字体引用解析（+mn-lt→Calibri）", title.paraFonts[1].name === "Calibri" && title.paraFonts[1].sizePt === 24, j(title.paraFonts[1]));
  check("垂直锚点 anchor=ctr", title.anchor === "ctr");
  const tbl = p.slides[0].tables[0];
  check("表格数据与首行/条带样式标志", j(tbl.rows) === j([["名称", "数量"], ["叉车", "3"]]) && tbl.headerRow === true && tbl.bandedRows === true, j(tbl));
  const chart = p.slides[0].charts[0];
  check("图表标题与序列数据", chart.title === "季度销售" && chart.series[0].name === "销售额" && j(chart.series[0].categories) === j(["一季度", "二季度"]) && j(chart.series[0].values) === j(["100", "200"]), j(chart));
  check("图片形状含媒体引用", sp.some((s) => s.kind === "picture" && s.image === "media/image1.png"), j(sp.find((s) => s.kind === "picture")));
  check("备注解析", p.slides[0].notes === "备注内容");
  check("页级字体聚合（兼容保留）", p.slides[0].fonts.includes("黑体") && p.slides[0].sizesPt.includes(44) && p.slides[0].boldUsed === true, j({ fonts: p.slides[0].fonts, sizes: p.slides[0].sizesPt }));
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

console.log("\n[attachmentRead：xlsx / pptx 模式]");
{
  const x = parseDocument("数据.xlsx", xlsxBytes);
  const xt = attachmentRead(x, { mode: "text" });
  check("xlsx text 含公式且值已类型化", xt.rows[2].cells.some((c) => c.f === "=B2*100") && xt.rows[1].cells[2].v === 42, j(xt.rows[2].cells));
  const xf = attachmentRead(x, { mode: "format" });
  const xa3 = xf.rows[2].cells[0];
  check("xlsx format 含填充/边框/对齐/raw 日期", xa3.fill === "#FF0000" && xa3.border === "thin #BFBFBF" && xa3.align?.h === "center" && xf.rows[1].cells[3].raw === 45413, j(xa3));

  const p = parseDocument("演示.pptx", pptxBytes);
  const st = attachmentRead(p, { mode: "structure" });
  check("pptx structure 形状清单", st.slides[0].shapes.length === 5 && st.slides[0].shapes[0].placeholder === "ctrTitle" && st.slides[0].shapes[2].rows === 2, j(st.slides[0].shapes));
  const fmt = attachmentRead(p, { mode: "format" });
  check("pptx format 形状位置与逐段格式", fmt.slides[0].shapes[0].posPt?.x === 78.7 && fmt.slides[0].shapes[0].paragraphs[0].font?.sizePt === 44 && fmt.slides[0].shapes[0].paragraphs[0].align === "Center", j(fmt.slides[0].shapes[0]));
  const txt = attachmentRead(p, { mode: "text" });
  check("pptx text 含表格与图表数据", j(txt.slides[0].tables?.[0]?.rows) === j([["名称", "数量"], ["叉车", "3"]]) && j(txt.slides[0].charts?.[0]?.series?.[0]?.values) === j(["100", "200"]), j({ t: txt.slides[0].tables, c: txt.slides[0].charts }));
  const tbls = attachmentRead(p, { mode: "tables" });
  check("pptx tables 模式", tbls.tables[0].slide === 1 && tbls.tables[0].rows[1][0] === "叉车", j(tbls.tables));
  const ov = attachmentOverview(p);
  check("pptx 概览含形状/表格/图表计数", ov.slides[0].shapes === 5 && ov.tableCount === 1 && ov.chartCount === 1, j(ov));
  const xov = attachmentOverview(x);
  check("xlsx 概览注记更新", String(xov.note).includes("填充色") && String(xov.note).includes("公式"), xov.note);
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
