/**
 * 附件文档解析（纯 JS，fflate 解压 OOXML + 轻量 XML 提取）
 * 支持：docx（段落/表格 + 样式/字体/对齐/缩进/行距）、xlsx（工作表/单元格 + 数字格式/加粗/合并）、
 *       pptx（幻灯片/形状文本 + 版式/字体字号）、txt/md/csv/json/log
 * 不支持：旧版二进制 .doc/.xls/.ppt、PDF（明确报错提示转换）
 */
import { unzipSync, strFromU8 } from "fflate";

/* ---------------- XML 基础 ---------------- */

export function unescXml(s) {
  return String(s || "")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/** 从开标签串中取属性值（属性名按"去前缀"匹配，锚定空白避免误命中长属性名） */
function attr(tagText, name) {
  const m = tagText.match(new RegExp(`(?:^|\\s)(?:[\\w.]+:)?${name}="([^"]*)"`, "i"));
  return m ? m[1] : undefined;
}

function toU8(buf) {
  return buf instanceof Uint8Array ? buf : new Uint8Array(buf);
}

function openZip(buf) {
  try {
    return unzipSync(toU8(buf));
  } catch {
    throw new Error("无法解压文档（文件可能损坏或不是 OOXML 格式）");
  }
}

/* ---------------- docx ---------------- */

const ALIGN_MAP = { left: "Left", start: "Left", center: "Centered", right: "Right", end: "Right", both: "Justified", justify: "Justified", distribute: "Justified" };

function docxText(fragment) {
  let out = "";
  const re = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>|<(?:[\w]+:)?tab\b[^>]*\/>|<(?:[\w]+:)?br\b[^>]*\/>/g;
  let m;
  while ((m = re.exec(fragment))) {
    if (m[1] !== undefined) out += unescXml(m[1]);
    else if (/tab/.test(m[0])) out += "\t";
    else out += "\n";
  }
  return out;
}

function parseRunProps(rPr) {
  if (!rPr) return {};
  const font = {};
  const fontsTag = rPr.match(/<(?:[\w]+:)?rFonts\b[^>]*>/)?.[0];
  if (fontsTag) {
    const ascii = attr(fontsTag, "ascii") || attr(fontsTag, "hAnsi") || attr(fontsTag, "eastAsia");
    if (ascii) font.name = ascii;
  }
  const sz = rPr.match(/<(?:[\w]+:)?sz\b[^>]*val="(\d+)"/)?.[1];
  if (sz) font.sizePt = Number(sz) / 2;
  const bTag = rPr.match(/<(?:[\w]+:)?b\b[^>]*\/?>(?:<\/(?:[\w]+:)?b>)?/)?.[0];
  if (bTag) font.bold = !(attr(bTag, "val") === "0" || attr(bTag, "val") === "false");
  const iTag = rPr.match(/<(?:[\w]+:)?i\b[^>]*\/?>(?:<\/(?:[\w]+:)?i>)?/)?.[0];
  if (iTag) font.italic = !(attr(iTag, "val") === "0" || attr(iTag, "val") === "false");
  const uTag = rPr.match(/<(?:[\w]+:)?u\b[^>]*>/)?.[0];
  if (uTag) font.underline = attr(uTag, "val") || "single";
  const cTag = rPr.match(/<(?:[\w]+:)?color\b[^>]*>/)?.[0];
  if (cTag) {
    const v = attr(cTag, "val");
    if (v && v !== "auto") font.color = `#${v}`;
  }
  const hTag = rPr.match(/<(?:[\w]+:)?highlight\b[^>]*>/)?.[0];
  if (hTag) font.highlightColor = attr(hTag, "val");
  return Object.keys(font).length ? font : {};
}

function parseDocx(buf) {
  const zip = openZip(buf);
  const docXml = zip["word/document.xml"] ? strFromU8(zip["word/document.xml"]) : "";
  if (!docXml) throw new Error("文档缺少 word/document.xml，无法解析");
  const stylesXml = zip["word/styles.xml"] ? strFromU8(zip["word/styles.xml"]) : "";

  // 样式表：styleId → 显示名
  const styleNames = new Map();
  const styleRe = /<(?:[\w]+:)?style\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?style>/g;
  let m;
  while ((m = styleRe.exec(stylesXml))) {
    if (/type="paragraph"/.test(m[1])) {
      const id = attr(m[1], "styleId");
      const name = m[2].match(/(?:[\w]+:)?name\b[^>]*val="([^"]*)"/)?.[1];
      if (id) styleNames.set(id, name || id);
    }
  }
  // 文档默认字体
  let defaultFont = null;
  {
    const d = stylesXml.match(/<(?:[\w]+:)?rPrDefault\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPrDefault>/)?.[1];
    if (d) defaultFont = parseRunProps(d) || null;
  }

  const body = docXml.match(/<(?:[\w]+:)?body\b[^>]*>([\s\S]*)<\/(?:[\w]+:)?body>/)?.[1] || docXml;
  const blockRe = /<(?:[\w]+:)?tbl\b[\s\S]*?<\/(?:[\w]+:)?tbl>|<(?:[\w]+:)?p\b[^>]*\/>|<(?:[\w]+:)?p\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?p>/g;

  const paragraphs = [];
  const tables = [];
  while ((m = blockRe.exec(body))) {
    const block = m[0];
    if (block.startsWith("<") && /<(?:[\w]+:)?tbl\b/.test(block.slice(0, 30))) {
      const rows = [];
      const trRe = /<(?:[\w]+:)?tr\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?tr>/g;
      let tr;
      while ((tr = trRe.exec(block))) {
        const cells = [];
        const tcRe = /<(?:[\w]+:)?tc\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?tc>/g;
        let tc;
        while ((tc = tcRe.exec(tr[0]))) cells.push(docxText(tc[0]).trim());
        rows.push(cells);
      }
      tables.push({ anchorParagraph: paragraphs.length, rows });
      continue;
    }
    // 段落
    const text = docxText(block);
    const pPr = block.match(/<(?:[\w]+:)?pPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?pPr>/)?.[1] || "";
    const styleId = pPr.match(/<(?:[\w]+:)?pStyle\b[^>]*>/)?.[0];
    const style = styleId ? styleNames.get(attr(styleId, "val")) || attr(styleId, "val") : undefined;
    const jcTag = pPr.match(/<(?:[\w]+:)?jc\b[^>]*>/)?.[0];
    const align = jcTag ? ALIGN_MAP[attr(jcTag, "val")] : undefined;
    const ol = pPr.match(/<(?:[\w]+:)?outlineLvl\b[^>]*val="(\d+)"/)?.[1];

    const spacingTag = pPr.match(/<(?:[\w]+:)?spacing\b[^>]*>/)?.[0];
    let spacing;
    if (spacingTag) {
      const before = attr(spacingTag, "before");
      const after = attr(spacingTag, "after");
      const line = attr(spacingTag, "line");
      const lineRule = attr(spacingTag, "lineRule") || "auto";
      spacing = {};
      if (before != null) spacing.beforePt = Number(before) / 20;
      if (after != null) spacing.afterPt = Number(after) / 20;
      if (line != null) spacing.lineSpacing = lineRule === "auto" ? `${(Number(line) / 240).toFixed(2)} 倍行距` : `${Number(line) / 20} 磅`;
      if (!Object.keys(spacing).length) spacing = undefined;
    }
    const indTag = pPr.match(/<(?:[\w]+:)?ind\b[^>]*>/)?.[0];
    let indent;
    if (indTag) {
      const firstLine = attr(indTag, "firstLine");
      const left = attr(indTag, "left");
      const hanging = attr(indTag, "hanging");
      indent = {};
      if (firstLine != null) indent.firstLinePt = Number(firstLine) / 20;
      if (left != null) indent.leftPt = Number(left) / 20;
      if (hanging != null) indent.hangingPt = Number(hanging) / 20;
      if (!Object.keys(indent).length) indent = undefined;
    }
    // 字体：取第一个有文字的 run 的 rPr，其次段落标记 rPr
    let font = null;
    const runRe = /<(?:[\w]+:)?r\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?r>/g;
    let run;
    while ((run = runRe.exec(block))) {
      if (!docxText(run[0]).trim()) continue;
      font = parseRunProps(run[0].match(/<(?:[\w]+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPr>/)?.[1]);
      break;
    }
    if (!font || !Object.keys(font).length) {
      const markRPr = pPr.match(/<(?:[\w]+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPr>/)?.[1];
      font = parseRunProps(markRPr) || null;
    }
    const p = { index: paragraphs.length, style, text };
    if (align) p.align = align;
    if (ol != null) p.outlineLevel = Number(ol);
    if (font && Object.keys(font).length) p.font = font;
    if (spacing) p.spacing = spacing;
    if (indent) p.indent = indent;
    paragraphs.push(p);
  }
  return { type: "docx", kind: "Word 文档", paragraphCount: paragraphs.length, paragraphs, tables, defaultFont: defaultFont && Object.keys(defaultFont).length ? defaultFont : null };
}

/* ---------------- xlsx ---------------- */

// OOXML 内置数字格式（常用子集）
const BUILTIN_NUMFMT = {
  0: "General", 1: "0", 2: "0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "yyyy-mm-dd", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM",
  20: "h:mm", 21: "h:mm:ss", 22: "yyyy-mm-dd h:mm", 37: "#,##0;(#,##0)", 38: "#,##0;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 43: '#,##0.00;(#,##0.00)', 44: '_("$"* #,##0.00_)',
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

function colToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function numToCol(n) {
  n += 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseXlsx(buf) {
  const zip = openZip(buf);
  const wbXml = zip["xl/workbook.xml"] ? strFromU8(zip["xl/workbook.xml"]) : "";
  if (!wbXml) throw new Error("工作簿缺少 xl/workbook.xml，无法解析");

  // rid → 路径
  const relMap = new Map();
  if (zip["xl/_rels/workbook.xml.rels"]) {
    const rels = strFromU8(zip["xl/_rels/workbook.xml.rels"]);
    const re = /<(?:[\w]+:)?Relationship\b[^>]*>/g;
    let m;
    while ((m = re.exec(rels))) {
      relMap.set(attr(m[0], "Id"), attr(m[0], "Target"));
    }
  }
  // 共享字符串
  const sst = [];
  if (zip["xl/sharedStrings.xml"]) {
    const s = strFromU8(zip["xl/sharedStrings.xml"]);
    const siRe = /<(?:[\w]+:)?si\b>([\s\S]*?)<\/(?:[\w]+:)?si>/g;
    let m, x;
    while ((m = siRe.exec(s))) {
      let t = "";
      const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
      while ((x = tRe.exec(m[1]))) t += unescXml(x[1]);
      sst.push(t);
    }
  }
  // 样式：numFmt 与字体加粗
  const numFmts = { ...BUILTIN_NUMFMT };
  const fontBold = [];
  if (zip["xl/styles.xml"]) {
    const st = strFromU8(zip["xl/styles.xml"]);
    let m;
    const nfRe = /<(?:[\w]+:)?numFmt\b[^>]*>/g;
    while ((m = nfRe.exec(st))) numFmts[attr(m[0], "numFmtId")] = unescXml(attr(m[0], "formatCode") || "");
    const fRe = /<(?:[\w]+:)?font\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?font>/g;
    while ((m = fRe.exec(st))) {
      const b = /<(?:[\w]+:)?b\b[^>]*\/?>/.test(m[1]) && !/b\s[^>]*val="(?:0|false)"/.test(m[1]);
      const name = m[1].match(/<(?:[\w]+:)?name\b[^>]*val="([^"]*)"/)?.[1];
      const sz = m[1].match(/<(?:[\w]+:)?sz\b[^>]*val="([\d.]+)"/)?.[1];
      const color = m[1].match(/<(?:[\w]+:)?color\b[^>]*rgb="([^"]*)"/)?.[1];
      fontBold.push({
        bold: b,
        ...(name ? { name } : {}),
        ...(sz ? { sizePt: Number(sz) } : {}),
        ...(color && color !== "auto" && color.length >= 6 ? { color: `#${color.slice(-6)}` } : {}),
      });
    }
  }
  // cellXfs: s → {numFmtId, fontId}
  const xfs = [];
  {
    const xfsXml = strFromU8(zip["xl/styles.xml"]).match(/<(?:[\w]+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?cellXfs>/)?.[1] || "";
    const re = /<(?:[\w]+:)?xf\b[^>]*\/?>/g;
    let m;
    while ((m = re.exec(xfsXml))) {
      xfs.push({ numFmtId: attr(m[0], "numFmtId"), fontId: Number(attr(m[0], "fontId") || 0) });
    }
  }

  const sheetTags = [...wbXml.matchAll(/<(?:[\w]+:)?sheet\b[^>]*>/g)].map((x) => x[0]);
  const sheets = sheetTags.map((tag, i) => {
    const name = unescXml(attr(tag, "name") || `Sheet${i + 1}`);
    const rid = attr(tag, "id") || attr(tag, "r:id");
    let target = relMap.get(rid) || "";
    target = target.replace(/^\/?xl\//, "").replace(/^\//, "");
    const file = target.startsWith("worksheets/") ? `xl/${target}` : target.startsWith("xl/") ? target : `xl/${target}`;
    const sheet = { name, path: file };
    return sheet;
  });

  const parsedSheets = [];
  for (const s of sheets) {
    const data = zip[s.path];
    if (!data) {
      parsedSheets.push({ name: s.name, error: "找不到工作表数据文件" });
      continue;
    }
    const xml = strFromU8(data);
    const dimension = xml.match(/<(?:[\w]+:)?dimension\b[^>]*>/)?.[0];
    const ref = dimension ? attr(dimension, "ref") : null;
    // 合并单元格
    const merges = [];
    {
      const mcXml = xml.match(/<(?:[\w]+:)?mergeCells\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?mergeCells>/)?.[1] || "";
      const re = /<(?:[\w]+:)?mergeCell\b[^>]*>/g;
      let m;
      while ((m = re.exec(mcXml))) merges.push(attr(m[0], "ref"));
    }
    // 列宽
    const cols = [];
    {
      const colsXml = xml.match(/<(?:[\w]+:)?cols\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?cols>/)?.[1] || "";
      const re = /<(?:[\w]+:)?col\b[^>]*\/?>/g;
      let m;
      while ((m = re.exec(colsXml))) {
        const min = Number(attr(m[0], "min"));
        const max = Number(attr(m[0], "max") || min);
        const width = attr(m[0], "width");
        if (min && width) cols.push({ from: numToCol(min - 1), to: numToCol(max - 1), widthChars: Number(width) });
      }
    }
    // 行与单元格
    const rows = [];
    const rowRe = /<(?:[\w]+:)?row\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?row>|<(?:[\w]+:)?row\b[^>]*\/>/g;
    let rm, cm;
    while ((rm = rowRe.exec(xml))) {
      const rowNo = Number(attr(rm[0].match(/<(?:[\w]+:)?row\b[^>]*>/)?.[0] || "", "r") || rows.length + 1);
      const cells = [];
      const cRe = /<(?:[\w]+:)?c\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?c>|<(?:[\w]+:)?c\b([^>]*)\/>/g;
      while ((cm = cRe.exec(rm[1]))) {
        const cTag = cm[1] != null ? cm[1] : cm[3] || "";
        const inner = cm[2] || "";
        const ref = attr(`<a ${cTag}>`, "r") || "";
        const colLetters = (ref.match(/^([A-Z]+)/) || [])[1];
        const t = attr(`<a ${cTag}>`, "t");
        const sId = attr(`<a ${cTag}>`, "s");
        let v = "";
        const vRaw = inner.match(/<(?:[\w]+:)?v\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?v>/)?.[1];
        const inline = inner.match(/<(?:[\w]+:)?is\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?is>/)?.[1];
        if (t === "s" && vRaw != null) v = sst[Number(vRaw)] ?? "";
        else if (t === "inlineStr" && inline) {
          const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
          let x;
          while ((x = tRe.exec(inline))) v += unescXml(x[1]);
        } else if (t === "b" && vRaw != null) v = vRaw === "1";
        else if (vRaw != null) v = unescXml(vRaw);
        let fmt;
        let bold;
        let fontInfo;
        if (sId != null && xfs[Number(sId)]) {
          const xf = xfs[Number(sId)];
          const nf = xf.numFmtId != null ? numFmts[xf.numFmtId] : null;
          if (nf && nf !== "General") fmt = nf;
          const f = fontBold[xf.fontId];
          if (f) {
            if (f.bold) bold = true;
            if (f.name || f.sizePt || f.color) fontInfo = f;
          }
        }
        const cell = { ref, v };
        if (colLetters) cell.col = colLetters;
        if (fmt) cell.numFmt = fmt;
        if (bold) cell.bold = true;
        if (fontInfo) cell.font = fontInfo;
        cells.push(cell);
      }
      rows.push({ r: rowNo, cells });
    }
    parsedSheets.push({
      name: s.name,
      ref,
      rowCount: rows.length,
      merges,
      ...(cols.length ? { colWidths: cols } : {}),
      rows,
    });
  }
  return { type: "xlsx", kind: "Excel 工作簿", sheetCount: parsedSheets.length, sheets: parsedSheets };
}

/* ---------------- pptx ---------------- */

function parsePptx(buf) {
  const zip = openZip(buf);
  const slideFiles = Object.keys(zip)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  if (!slideFiles.length) throw new Error("演示文稿中没有幻灯片，无法解析");

  const slides = slideFiles.map((file, i) => {
    const xml = strFromU8(zip[file]);
    // 版式名（通过 rels → layout 文件）
    let layout;
    const relFile = `ppt/slides/_rels/${file.split("/").pop()}.rels`;
    if (zip[relFile]) {
      const rels = strFromU8(zip[relFile]);
      const target = rels.match(/Target="([^"]*slideLayout[^"]*)"/)?.[1];
      if (target) {
        const lp = ("ppt/" + target.replace(/^\.\.\//, "").replace(/^\//, "")).replace("ppt/ppt/", "ppt/");
        const lxml = zip[lp] ? strFromU8(zip[lp]) : "";
        layout = lxml?.match(/<(?:[\w]+:)?cSld\b[^>]*name="([^"]*)"/)?.[1] || undefined;
      }
    }
    // 形状文本
    const shapes = [];
    const spRe = /<(?:[\w]+:)?sp\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?sp>/g;
    let m;
    while ((m = spRe.exec(xml))) {
      const paras = [];
      const pRe = /<(?:[\w]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?p>/g;
      let pm, tm;
      while ((pm = pRe.exec(m[0]))) {
        let t = "";
        const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
        while ((tm = tRe.exec(pm[1]))) t += unescXml(tm[1]);
        if (t) paras.push(t);
      }
      if (paras.length) shapes.push({ paragraphs: paras });
    }
    // 字体信息（format 模式用）
    const fonts = new Set();
    const sizes = new Set();
    let boldUsed = false;
    const fRe = /<(?:[\w]+:)?latin\b[^>]*typeface="([^"]*)"/g;
    while ((m = fRe.exec(xml))) if (m[1] && m[1] !== "+mn-lt" && m[1] !== "+mj-lt") fonts.add(m[1]);
    const szRe = /<(?:[\w]+:)?rPr\b[^>]*\ssz="(\d+)"/g;
    while ((m = szRe.exec(xml))) sizes.add(Number(m[1]) / 100);
    if (/<(?:[\w]+:)?rPr\b[^>]*\sb="1"/.test(xml)) boldUsed = true;
    const n = i + 1;
    return {
      n,
      ...(layout ? { layout } : {}),
      shapes,
      fonts: [...fonts],
      sizesPt: [...sizes].sort((a, b) => a - b),
      boldUsed,
      notesPath: zip[`ppt/notesSlides/notesSlide${n}.xml`] ? 1 : 0,
    };
  });
  // 备注
  for (const s of slides) {
    if (s.notesPath) {
      const nx = strFromU8(zip[`ppt/notesSlides/notesSlide${s.n}.xml`]);
      let notes = "";
      const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
      let m;
      while ((m = tRe.exec(nx))) notes += unescXml(m[1]);
      s.notes = notes.trim() || undefined;
    }
    delete s.notesPath;
  }
  return { type: "pptx", kind: "PowerPoint 演示文稿", slideCount: slides.length, slides };
}

/* ---------------- 纯文本 ---------------- */

function parseText(buf) {
  return { type: "text", kind: "文本文件", text: buf.toString("utf8") };
}

/* ---------------- 入口 ---------------- */

export const SUPPORTED_EXTS = ["docx", "dotx", "xlsx", "xlsm", "pptx", "txt", "md", "csv", "json", "log"];

export function parseDocument(name, buffer) {
  const ext = String(name || "").toLowerCase().split(".").pop();
  switch (ext) {
    case "docx":
    case "dotx":
      return parseDocx(buffer);
    case "xlsx":
    case "xlsm":
      return parseXlsx(buffer);
    case "pptx":
      return parsePptx(buffer);
    case "txt":
    case "md":
    case "csv":
    case "json":
    case "log":
      return parseText(buffer);
    default:
      throw new Error(
        `不支持的附件格式 .${ext}。支持：${SUPPORTED_EXTS.join(" / ")}；旧版 .doc/.xls/.ppt 与 PDF 请先另存为对应的新格式`
      );
  }
}

/* ---------------- 读取工具（attachment_list / attachment_read 的执行逻辑） ---------------- */

const clamp = (n, def, max) => Math.min(Math.max(Number.isFinite(Number(n)) ? Number(n) : def, 0), max);

export function attachmentOverview(parsed) {
  if (parsed.type === "docx") {
    const headings = parsed.paragraphs
      .filter((p) => (p.outlineLevel != null && p.outlineLevel <= 3) || /^heading\s*\d$/i.test(p.style || "") || /^(标题)\s*\d+$/.test(p.style || ""))
      .slice(0, 40)
      .map((p) => ({ index: p.index, style: p.style, text: (p.text || "").slice(0, 60) }));
    return {
      kind: parsed.kind,
      paragraphCount: parsed.paragraphCount,
      tableCount: parsed.tables.length,
      defaultFont: parsed.defaultFont || undefined,
      headings: headings.length ? headings : undefined,
      note: "mode=text 读段落文本；mode=format 读字体/对齐/缩进/行距；mode=tables 读表格",
    };
  }
  if (parsed.type === "xlsx") {
    return {
      kind: parsed.kind,
      sheets: parsed.sheets.map((s) => ({ name: s.name, ref: s.ref, rowCount: s.rowCount, merges: s.merges?.length })),
      note: "sheet 参数选表（名称或序号），mode=text/format 读单元格（含数字格式/加粗/字体）",
    };
  }
  if (parsed.type === "pptx") {
    return {
      kind: parsed.kind,
      slideCount: parsed.slideCount,
      slides: parsed.slides.map((s) => ({ n: s.n, layout: s.layout, firstText: (s.shapes[0]?.paragraphs?.[0] || "").slice(0, 50) })),
    };
  }
  return { kind: parsed.kind, chars: parsed.text.length, lines: parsed.text.split(/\r?\n/).length };
}

export function attachmentRead(parsed, args) {
  const mode = ["structure", "text", "format", "tables"].includes(args?.mode) ? args.mode : "text";

  if (parsed.type === "text") {
    const offset = clamp(args?.offset, 0, 1e9);
    const limit = clamp(args?.limit, 4000, 40000);
    return { type: parsed.type, mode: "text", total: parsed.text.length, offset, text: parsed.text.slice(offset, offset + limit) };
  }

  if (parsed.type === "docx") {
    if (mode === "tables") {
      const list = parsed.tables.map((t, i) => ({ table: i + 1, anchorParagraph: t.anchorParagraph, rows: t.rows.slice(0, 40), rowCount: t.rows.length }));
      return { type: "docx", mode, tableCount: parsed.tables.length, tables: list };
    }
    if (mode === "structure") {
      const headings = parsed.paragraphs
        .filter((p) => p.outlineLevel != null || /^heading\s*\d$/i.test(p.style || "") || /^(标题)\s*\d+$/.test(p.style || "") || /^(Title|标题)$/.test(p.style || ""))
        .map((p) => ({ index: p.index, style: p.style, outlineLevel: p.outlineLevel, text: (p.text || "").slice(0, 80) }));
      return { type: "docx", mode, total: parsed.paragraphCount, headings };
    }
    const offset = clamp(args?.offset, 0, parsed.paragraphCount);
    const cap = mode === "format" ? 100 : 400;
    const limit = clamp(args?.limit, mode === "format" ? 50 : 120, cap);
    const slice = parsed.paragraphs.slice(offset, offset + limit);
    return {
      type: "docx",
      mode,
      total: parsed.paragraphCount,
      offset,
      ...(parsed.defaultFont ? { defaultFont: parsed.defaultFont } : {}),
      paragraphs: mode === "format" ? slice : slice.map((p) => ({ index: p.index, style: p.style, text: p.text })),
    };
  }

  if (parsed.type === "xlsx") {
    let sheet = parsed.sheets[0];
    if (args?.sheet != null && args.sheet !== "") {
      const key = String(args.sheet);
      sheet = parsed.sheets.find((s) => s.name === key) || parsed.sheets[Number(key) - 1] || parsed.sheets[0];
    }
    if (!sheet) throw new Error("工作簿中没有工作表");
    if (mode === "structure") {
      return {
        type: "xlsx",
        mode,
        sheets: parsed.sheets.map((s) => ({ name: s.name, ref: s.ref, rowCount: s.rowCount, merges: s.merges })),
      };
    }
    const offset = clamp(args?.offset, 0, sheet.rowCount);
    const limit = clamp(args?.limit, 100, 500);
    const rows = sheet.rows.slice(offset, offset + limit);
    return {
      type: "xlsx",
      mode,
      sheet: sheet.name,
      ref: sheet.ref,
      rowCount: sheet.rowCount,
      offset,
      merges: sheet.merges,
      ...(sheet.colWidths ? { colWidths: sheet.colWidths } : {}),
      // mode=format 保留单元格格式信息；text 仅取值
      rows: rows.map((r) => ({
        r: r.r,
        cells: r.cells.map((c) =>
          mode === "format"
            ? c
            : { ref: c.ref, v: c.v }
        ),
      })),
    };
  }

  if (parsed.type === "pptx") {
    const offset = clamp(args?.offset, 1, parsed.slideCount) || 1;
    const limit = clamp(args?.limit, 10, 50);
    const slice = parsed.slides.filter((s) => s.n >= offset && s.n < offset + limit);
    if (mode === "structure") {
      return { type: "pptx", mode, slideCount: parsed.slideCount, slides: slice.map((s) => ({ n: s.n, layout: s.layout, texts: s.shapes.map((sp) => sp.paragraphs[0]).filter(Boolean).slice(0, 3) })) };
    }
    if (mode === "format") {
      return { type: "pptx", mode, slideCount: parsed.slideCount, slides: slice.map((s) => ({ n: s.n, layout: s.layout, fonts: s.fonts, sizesPt: s.sizesPt, boldUsed: s.boldUsed })) };
    }
    return { type: "pptx", mode: "text", slideCount: parsed.slideCount, slides: slice.map((s) => ({ n: s.n, layout: s.layout, shapes: s.shapes, notes: s.notes })) };
  }

  throw new Error("未知的附件类型");
}
