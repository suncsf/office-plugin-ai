/**
 * 附件文档解析（纯 JS，fflate 解压 OOXML + 轻量 XML 提取）
 * 支持：docx（段落/表格 + 样式/字体/对齐/缩进/行距）、xlsx（工作表/单元格 + 数字格式/日期换算/公式/字体/对齐/填充/边框/合并）、
 *       pptx（形状清单/文本/表格数据/图表数据 + 位置尺寸/填充边框/逐段文字格式/版式/主题色）、txt/md/csv/json/log
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

/** 从开标签串中取属性值（属性名按"去前缀"匹配，锚定空白避免误命中长属性名）；tagText 为空时返回 undefined */
function attr(tagText, name) {
  const m = String(tagText || "").match(new RegExp(`(?:^|\\s)(?:[\\w.]+:)?${name}="([^"]*)"`, "i"));
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
    // 主题字体（真实 Word 文档常用）：asciiTheme/eastAsiaTheme 指向 theme1.xml，由解析器随后解析
    const asciiTheme = attr(fontsTag, "asciiTheme") || attr(fontsTag, "hAnsiTheme");
    const eastAsiaTheme = attr(fontsTag, "eastAsiaTheme");
    if (!ascii && (asciiTheme || eastAsiaTheme)) font.themeName = { asciiTheme, eastAsiaTheme };
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

const mergeDefined = (base, over) => {
  const out = { ...base };
  for (const k of Object.keys(over || {})) if (over[k] !== undefined) out[k] = over[k];
  return out;
};

/** 解析段落属性原始 XML（pPr）→ 结构化（直接格式与样式定义共用） */
function parseParaProps(pPrRaw) {
  if (!pPrRaw) return {};
  const out = {};
  const jcTag = pPrRaw.match(/<(?:[\w]+:)?jc\b[^>]*>/)?.[0];
  if (jcTag) out.align = ALIGN_MAP[attr(jcTag, "val")];
  const ol = pPrRaw.match(/<(?:[\w]+:)?outlineLvl\b[^>]*val="(\d+)"/)?.[1];
  if (ol != null) out.outlineLevel = Number(ol);
  const spacingTag = pPrRaw.match(/<(?:[\w]+:)?spacing\b[^>]*>/)?.[0];
  if (spacingTag) {
    const before = attr(spacingTag, "before");
    const after = attr(spacingTag, "after");
    const line = attr(spacingTag, "line");
    const lineRule = attr(spacingTag, "lineRule") || "auto";
    const spacing = {};
    if (before != null) spacing.beforePt = Number(before) / 20;
    if (after != null) spacing.afterPt = Number(after) / 20;
    if (line != null) {
      spacing.lineSpacing = lineRule === "auto" ? `${(Number(line) / 240).toFixed(2)} 倍行距` : `${Number(line) / 20} 磅`;
    }
    if (Object.keys(spacing).length) out.spacing = spacing;
  }
  const indTag = pPrRaw.match(/<(?:[\w]+:)?ind\b[^>]*>/)?.[0];
  if (indTag) {
    const firstLine = attr(indTag, "firstLine");
    const left = attr(indTag, "left");
    const hanging = attr(indTag, "hanging");
    const indent = {};
    if (firstLine != null) indent.firstLinePt = Number(firstLine) / 20;
    if (left != null) indent.leftPt = Number(left) / 20;
    if (hanging != null) indent.hangingPt = Number(hanging) / 20;
    if (Object.keys(indent).length) out.indent = indent;
  }
  return out;
}

function parseDocx(buf) {
  const zip = openZip(buf);
  const docXml = zip["word/document.xml"] ? strFromU8(zip["word/document.xml"]) : "";
  if (!docXml) {
    throw new Error("文档缺少 word/document.xml，无法解析（文件内条目: " + Object.keys(zip).slice(0, 6).join(", ") + "）");
  }
  const stylesXml = zip["word/styles.xml"] ? strFromU8(zip["word/styles.xml"]) : "";

  // 主题字体（theme1.xml）：真实文档的 rFonts 常用 asciiTheme/eastAsiaTheme 引用
  const themeFonts = {};
  {
    const themeXml = zip["word/theme/theme1.xml"] ? strFromU8(zip["word/theme/theme1.xml"]) : "";
    if (themeXml) {
      const major = themeXml.match(/<a:majorFont>[\s\S]*?<\/a:majorFont>/)?.[0] || "";
      const minor = themeXml.match(/<a:minorFont>[\s\S]*?<\/a:minorFont>/)?.[0] || "";
      themeFonts.majorLatin = major.match(/<a:latin typeface="([^"]*)"/)?.[1];
      themeFonts.majorEa = major.match(/<a:ea typeface="([^"]*)"/)?.[1];
      themeFonts.minorLatin = minor.match(/<a:latin typeface="([^"]*)"/)?.[1];
      themeFonts.minorEa = minor.match(/<a:ea typeface="([^"]*)"/)?.[1];
    }
  }
  const resolveThemeFont = (font) => {
    if (!font || !font.themeName) return font;
    const t = font.themeName;
    const pick = (v) => (v && !v.startsWith("+") ? v : undefined);
    const latin = t.asciiTheme ? pick(t.asciiTheme.includes("major") ? themeFonts.majorLatin : themeFonts.minorLatin) : undefined;
    const ea = t.eastAsiaTheme ? pick(t.eastAsiaTheme.includes("major") ? themeFonts.majorEa : themeFonts.minorEa) : undefined;
    const name = ea || latin;
    if (name) font.name = name;
    delete font.themeName;
    return font;
  };

  // 文档默认（docDefaults）：所有样式的最底层
  let defaultFont = {};
  let defaultPara = {};
  {
    const rPrDefault = stylesXml.match(/<(?:[\w]+:)?rPrDefault\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPrDefault>/)?.[1];
    if (rPrDefault) defaultFont = resolveThemeFont(parseRunProps(rPrDefault));
    const pPrDefault = stylesXml.match(/<(?:[\w]+:)?pPrDefault\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?pPrDefault>/)?.[1];
    if (pPrDefault) defaultPara = parseParaProps(pPrDefault);
  }

  // 样式表：id → { type, name, basedOn, pPrRaw, rPrRaw }
  const styles = new Map();
  const styleRe = /<(?:[\w]+:)?style\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?style>/g;
  let m;
  while ((m = styleRe.exec(stylesXml))) {
    const tag = m[1];
    const styleBody = m[2];
    const id = attr(tag, "styleId");
    if (!id) continue;
    const basedOnTag = styleBody.match(/<(?:[\w]+:)?basedOn\b[^>]*>/)?.[0];
    styles.set(id, {
      type: /type="(\w+)"/.exec(tag)?.[1] || "paragraph",
      name: styleBody.match(/<(?:[\w]+:)?name\b[^>]*val="([^"]*)"/)?.[1] || id,
      basedOn: basedOnTag ? attr(basedOnTag, "val") : undefined,
      pPrRaw: styleBody.match(/<(?:[\w]+:)?pPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?pPr>/)?.[1] || "",
      rPrRaw: styleBody.match(/<(?:[\w]+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPr>/)?.[1] || "",
    });
  }

  /**
   * 样式有效格式：docDefaults → basedOn 链（基类先应用、派生覆盖）。
   * 真实 Word 文档的格式主要定义在样式里，段落往往只有 pStyle 引用。
   */
  const styleChain = (styleId) => {
    let font = { ...defaultFont };
    let para = { ...defaultPara };
    const seen = new Set();
    const chain = [];
    let cur = styleId ? styles.get(styleId) : null;
    while (cur && cur.type === "paragraph" && !seen.has(cur)) {
      seen.add(cur);
      chain.push(cur);
      cur = cur.basedOn ? styles.get(cur.basedOn) : null;
    }
    for (const s of chain.reverse()) {
      font = mergeDefined(font, parseRunProps(s.rPrRaw));
      para = mergeDefined(para, parseParaProps(s.pPrRaw));
    }
    return { font, para };
  };

  const body = docXml.match(/<(?:[\w]+:)?body\b[^>]*>([\s\S]*)<\/(?:[\w]+:)?body>/)?.[1] || docXml;
  const blockRe = /<(?:[\w]+:)?tbl\b[\s\S]*?<\/(?:[\w]+:)?tbl>|<(?:[\w]+:)?p\b[^>]*\/>|<(?:[\w]+:)?p\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?p>/g;

  const paragraphs = [];
  const tables = [];
  let skipped = 0;
  while ((m = blockRe.exec(body))) {
    const block = m[0];
    try {
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
      const pStyleTag = pPr.match(/<(?:[\w]+:)?pStyle\b[^>]*>/)?.[0];
      const styleId = pStyleTag ? attr(pStyleTag, "val") : undefined;
      const styleName = styleId ? styles.get(styleId)?.name || styleId : undefined;
      const chain = styleChain(styleId);
      const direct = parseParaProps(pPr);

      // 字体优先级：run 直接格式（叠加字符样式）> 段落标记 > 样式链
      let font = null;
      const runRe = /<(?:[\w]+:)?r\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?r>/g;
      let run;
      while ((run = runRe.exec(block))) {
        if (!docxText(run[0]).trim()) continue;
        const rPrRaw = run[0].match(/<(?:[\w]+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPr>/)?.[1] || "";
        const rStyleTag = rPrRaw.match(/<(?:[\w]+:)?rStyle\b[^>]*>/)?.[0];
        const rStyleId = rStyleTag ? attr(rStyleTag, "val") : undefined;
        let f = chain.font;
        if (rStyleId) {
          const cs = styles.get(rStyleId);
          if (cs && cs.type === "character") f = mergeDefined(f, parseRunProps(cs.rPrRaw));
        }
        f = mergeDefined(f, parseRunProps(rPrRaw));
        font = f;
        break;
      }
      if (!font) {
        const markRPr = pPr.match(/<(?:[\w]+:)?rPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?rPr>/)?.[1];
        font = markRPr ? mergeDefined(chain.font, parseRunProps(markRPr)) : chain.font;
      }
      font = resolveThemeFont(font);

      const p = { index: paragraphs.length, style: styleName, text };
      const effAlign = direct.align ?? chain.para.align;
      if (effAlign) p.align = effAlign;
      const effOutline = direct.outlineLevel ?? chain.para.outlineLevel;
      if (effOutline != null) p.outlineLevel = effOutline;
      if (font && Object.keys(font).length) p.font = font;
      const effSpacing = direct.spacing ?? chain.para.spacing;
      if (effSpacing) p.spacing = effSpacing;
      const effIndent = direct.indent ?? chain.para.indent;
      if (effIndent) p.indent = effIndent;
      paragraphs.push(p);
    } catch (e) {
      // 单块解析失败不拖垮整体：内容仍尽力保留
      skipped++;
      try {
        paragraphs.push({ index: paragraphs.length, text: docxText(block) });
      } catch { /* 忽略 */ }
    }
  }
  const result = {
    type: "docx",
    kind: "Word 文档",
    paragraphCount: paragraphs.length,
    paragraphs,
    tables,
    defaultFont: Object.keys(defaultFont).length ? defaultFont : null,
  };
  if (skipped) result.parseWarnings = `${skipped} 个块解析异常已跳过（内容尽力保留）`;
  return result;
}

/* ---------------- 主题色与颜色（xlsx / pptx 共用） ---------------- */

/** 解析 theme 的 a:clrScheme → { dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink }（#RRGGBB） */
function parseThemeColors(themeXml) {
  if (!themeXml) return null;
  const scheme = themeXml.match(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/)?.[0];
  if (!scheme) return null;
  const out = {};
  for (const n of ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
    const seg = scheme.match(new RegExp(`<a:${n}>([\\s\\S]*?)</a:${n}>`))?.[1] || "";
    const v = seg.match(/val="([0-9A-Fa-f]{6})"/)?.[1] || seg.match(/lastClr="([0-9A-Fa-f]{6})"/)?.[1];
    if (v) out[n] = `#${v.toUpperCase()}`;
  }
  return Object.keys(out).length ? out : null;
}

function rgbToHsl(r, g, b) {
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 510;
  const d = max - min;
  const s = d === 0 ? 0 : d / 255 / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d !== 0) {
    if (max === r) h = (((g - b) / d) % 6 + 6) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
  }
  return [h, s, l];
}

function hslToHex(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = ((h % 6) + 6) % 6;
  const x = c * (1 - Math.abs(hp - 1));
  const m = l - c / 2;
  const [r, g, b] = [[c, x, 0], [x, c, 0], [0, c, x], [0, x, c], [x, 0, c], [c, 0, x]][Math.floor(hp)];
  return [r, g, b].map((v) => Math.round((v + m) * 255).toString(16).padStart(2, "0")).join("");
}

/**
 * HSL 亮度调整：tint>0 向白、<0 向黑（SpreadsheetML tint 语义）；
 * lumMod/lumOff 为 DrawingML 亮度缩放/偏移（Office「深色/浅色 x%」的实现机制，如 白色·深色25% = lumMod 75%）。
 */
function lumAdjusted(hex6, { tint = 0, shade = 0, lumMod = 0, lumOff = 0 } = {}) {
  if (!tint && !shade && !lumMod && !lumOff) return hex6.toUpperCase();
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex6.slice(i, i + 2), 16));
  const [h, s, l] = rgbToHsl(r, g, b);
  let L = l;
  if (tint) L = L + (1 - L) * tint;
  if (shade) L *= shade;
  if (lumMod || lumOff) L = Math.min(1, Math.max(0, L * (lumMod || 1) + lumOff));
  return hslToHex(h, s, L).toUpperCase();
}

/** DrawingML 颜色片段（srgbClr/schemeClr/sysClr + tint/shade/lumMod/lumOff 修饰）→ #RRGGBB */
function drawColor(fragment, themeColors) {
  if (!fragment) return undefined;
  const m = fragment.match(/<a:(srgbClr|schemeClr|sysClr|prstClr)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:\1>)/);
  if (!m) return undefined;
  const attrs = `<a ${m[2]}>`;
  const inner = m[3] || "";
  let hex;
  if (m[1] === "srgbClr" || m[1] === "sysClr") {
    hex = (attr(attrs, m[1] === "sysClr" ? "lastClr" : "val") || "").toUpperCase();
  } else if (m[1] === "schemeClr") {
    const v = attr(attrs, "val");
    hex = (themeColors?.[{ bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }[v] || v] || "").slice(1);
  }
  if (!hex || !/^[0-9A-F]{6}$/.test(hex)) return undefined;
  // tint/shade/lumMod/lumOff 均为十万分比（val="75000" = 0.75）
  const pct = (tag, name) => Number(tag.match(new RegExp(`<a:${name}\\b[^>]*val="(-?[\\d.]+)"`))?.[1] || 0) / 100000;
  return `#${lumAdjusted(hex, {
    tint: pct(inner, "tint"),
    shade: pct(inner, "shade"),
    lumMod: pct(inner, "lumMod"),
    lumOff: pct(inner, "lumOff"),
  })}`;
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

// 内置日期/时间类数字格式 id（ECMA-376 §18.8.30）
const DATE_NUMFMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

/** 数字格式是否为日期/时间类（内置 id 表 + 自定义格式码启发式：剔除颜色段/引号/转义后含 y m d h s） */
function isDateNumFmt(numFmtId, code) {
  if (numFmtId == null || numFmtId === "") return false;
  const id = Number(numFmtId);
  if (DATE_NUMFMT_IDS.has(id)) return true;
  if (!code) return false;
  const cleaned = String(code).replace(/\[[^\]]*\]/g, "").replace(/"[^"]*"/g, "").replace(/\\./g, "");
  return /[ymdhs]/i.test(cleaned);
}

/** Excel 日期序列值 → "yyyy-mm-dd[ hh:mm:ss]"（1900 系统；serial 60 为假闰日按 2-28 归并；date1904 系统 +1462 天换算） */
function serialToDateTime(v, date1904) {
  const serial = Number(v) + (date1904 ? 1462 : 0);
  if (!Number.isFinite(serial)) return String(v);
  const days = Math.floor(serial);
  const ms = Date.UTC(1900, 0, 1) + (days - (days >= 60 ? 2 : 1)) * 86400000 + Math.round((serial - days) * 86400000);
  const d = new Date(ms);
  const p2 = (n) => String(n).padStart(2, "0");
  const date = `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`;
  if (serial - days < 1e-6) return date;
  const secs = Math.round((serial - days) * 86400);
  return `${date} ${p2(Math.floor(secs / 3600))}:${p2(Math.floor((secs % 3600) / 60))}:${p2(secs % 60)}`;
}

// SpreadsheetML 颜色的 theme 索引顺序（0=lt1、1=dk1、2=lt2、3=dk2、4-9=accent1-6）
const XLSX_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** SpreadsheetML 颜色元素（<color>/<fgColor>，rgb/theme/tint 属性）→ #RRGGBB */
function xlsxColor(colorTag, themeColors) {
  if (!colorTag) return undefined;
  const tint = Number(attr(colorTag, "tint") || 0);
  const rgb = attr(colorTag, "rgb");
  if (rgb && rgb.length >= 6) return `#${lumAdjusted(rgb.slice(-6), { tint })}`;
  const themeIdx = attr(colorTag, "theme");
  if (themeIdx != null) {
    const base = themeColors?.[XLSX_THEME_ORDER[Number(themeIdx)]];
    if (base) return `#${lumAdjusted(base.slice(1), { tint })}`;
  }
  return undefined;
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
  // 主题色（xl/theme/theme1.xml 的 clrScheme）
  const themeColors = parseThemeColors(zip["xl/theme/theme1.xml"] ? strFromU8(zip["xl/theme/theme1.xml"]) : "");
  // 1904 日期系统（罕见，Mac 旧工作簿）
  const date1904 = /date1904="(?:1|true)"/.test(wbXml);

  // 样式表：numFmt / fonts / fills / borders / cellXfs
  const numFmts = { ...BUILTIN_NUMFMT };
  const fonts = [];
  const fills = [];
  const borders = [];
  const xfs = [];
  const stXml = zip["xl/styles.xml"] ? strFromU8(zip["xl/styles.xml"]) : "";
  {
    let m;
    const nfRe = /<(?:[\w]+:)?numFmt\b[^>]*>/g;
    while ((m = nfRe.exec(stXml))) numFmts[attr(m[0], "numFmtId")] = unescXml(attr(m[0], "formatCode") || "");
    const fRe = /<(?:[\w]+:)?font\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?font>/g;
    while ((m = fRe.exec(stXml))) {
      const b = /<(?:[\w]+:)?b\b[^>]*\/?>/.test(m[1]) && !/b\s[^>]*val="(?:0|false)"/.test(m[1]);
      const it = /<(?:[\w]+:)?i\b[^>]*\/?>/.test(m[1]) && !/i\s[^>]*val="(?:0|false)"/.test(m[1]);
      const name = m[1].match(/<(?:[\w]+:)?name\b[^>]*val="([^"]*)"/)?.[1];
      const sz = m[1].match(/<(?:[\w]+:)?sz\b[^>]*val="([\d.]+)"/)?.[1];
      const color = xlsxColor(m[1].match(/<(?:[\w]+:)?color\b[^>]*\/?>/)?.[0], themeColors);
      const uTag = m[1].match(/<(?:[\w]+:)?u\b[^>]*\/?>/)?.[0];
      const u = uTag ? attr(uTag, "val") || "single" : undefined;
      const font = {
        ...(b ? { bold: true } : {}),
        ...(it ? { italic: true } : {}),
        ...(name ? { name } : {}),
        ...(sz ? { sizePt: Number(sz) } : {}),
        ...(color ? { color } : {}),
        ...(u && u !== "none" ? { underline: u } : {}),
      };
      fonts.push(font);
    }
    // 填充（patternFill 的 fgColor）
    const fillsXml = stXml.match(/<(?:[\w]+:)?fills\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?fills>/)?.[1] || "";
    const fillRe = /<(?:[\w]+:)?fill\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?fill>/g;
    while ((m = fillRe.exec(fillsXml))) {
      const pfTag = m[1].match(/<(?:[\w]+:)?patternFill\b[^>]*>/)?.[0];
      const pattern = pfTag ? attr(pfTag, "patternType") : undefined;
      const pfBody = m[1].match(/<(?:[\w]+:)?patternFill\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?patternFill>/)?.[1] || "";
      const color = xlsxColor(pfBody.match(/<(?:[\w]+:)?fgColor\b[^>]*\/?>/)?.[0], themeColors);
      fills.push(pattern && pattern !== "none" && color ? { pattern, color } : null);
    }
    // 边框（四边 style + color）
    const bordersXml = stXml.match(/<(?:[\w]+:)?borders\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?borders>/)?.[1] || "";
    const bdRe = /<(?:[\w]+:)?border\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?border>/g;
    while ((m = bdRe.exec(bordersXml))) {
      const sides = {};
      for (const side of ["left", "right", "top", "bottom"]) {
        const sideRe = new RegExp(`<(?:[\\w]+:)?${side}\\b([^>]*)>([\\s\\S]*?)</(?:[\\w]+:)?${side}>|<(?:[\\w]+:)?${side}\\b([^>]*)/>`, "i");
        const sm = m[1].match(sideRe);
        if (!sm) continue;
        const sideAttrs = `<a ${sm[1] ?? sm[3]}>`;
        const style = attr(sideAttrs, "style");
        if (!style || style === "none") continue;
        const color = xlsxColor((sm[2] || "").match(/<(?:[\w]+:)?color\b[^>]*\/?>/)?.[0], themeColors);
        sides[side] = color ? `${style} ${color}` : style;
      }
      borders.push(sides);
    }
    // cellXfs: s → { numFmtId, fontId, fillId, borderId, align }
    const xfsXml = stXml.match(/<(?:[\w]+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?cellXfs>/)?.[1] || "";
    const xfRe = /<(?:[\w]+:)?xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w]+:)?xf>)/g;
    while ((m = xfRe.exec(xfsXml))) {
      const tagAttrs = `<a ${m[1]}>`;
      const alignTag = (m[2] || "").match(/<(?:[\w]+:)?alignment\b[^>]*\/?>/)?.[0];
      const h = alignTag ? attr(alignTag, "horizontal") : undefined;
      const v = alignTag ? attr(alignTag, "vertical") : undefined;
      xfs.push({
        numFmtId: attr(tagAttrs, "numFmtId"),
        fontId: Number(attr(tagAttrs, "fontId") || 0),
        fillId: Number(attr(tagAttrs, "fillId") || 0),
        borderId: Number(attr(tagAttrs, "borderId") || 0),
        align: alignTag
          ? {
              ...(h && h !== "general" ? { h } : {}),
              ...(v && v !== "bottom" ? { v } : {}),
              ...(attr(alignTag, "wrapText") === "1" ? { wrap: true } : {}),
              ...(attr(alignTag, "indent") ? { indent: Number(attr(alignTag, "indent")) } : {}),
              ...(attr(alignTag, "textRotation") && attr(alignTag, "textRotation") !== "0" ? { rotate: Number(attr(alignTag, "textRotation")) } : {}),
            }
          : null,
      });
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
        let raw; // 日期类数值保留原始序列值（format 模式输出）
        // 公式：<f>A1+B1</f>；共享公式从单元格取主公式文本（从属格无文本则忽略）
        const fMatch = inner.match(/<(?:[\w]+:)?f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w]+:)?f>)/);
        const formula = fMatch && fMatch[2] != null && fMatch[2].trim() ? unescXml(fMatch[2]).trim() : null;
        const vRaw = inner.match(/<(?:[\w]+:)?v\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?v>/)?.[1];
        const inline = inner.match(/<(?:[\w]+:)?is\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?is>/)?.[1];
        if (t === "s" && vRaw != null) v = sst[Number(vRaw)] ?? "";
        else if (t === "inlineStr" && inline) {
          const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
          let x;
          while ((x = tRe.exec(inline))) v += unescXml(x[1]);
        } else if (t === "b" && vRaw != null) v = vRaw === "1";
        else if (vRaw != null) {
          // 数值类型化：数字 → number，其余（错误值 #DIV/0!、公式文本结果等）保持字符串
          const s = unescXml(vRaw);
          const n = Number(s);
          v = s !== "" && Number.isFinite(n) ? n : s;
        }
        const cell = { ref, v };
        let fmt;
        if (sId != null && xfs[Number(sId)]) {
          const xf = xfs[Number(sId)];
          const nf = xf.numFmtId != null ? numFmts[xf.numFmtId] : null;
          if (nf && nf !== "General") fmt = nf;
          // 日期/时间格式 + 数值 → 可读日期字符串
          if (typeof v === "number" && isDateNumFmt(xf.numFmtId, nf)) {
            raw = v;
            v = serialToDateTime(v, date1904);
            cell.v = v;
          }
          const f = fonts[xf.fontId];
          if (f && Object.keys(f).length) cell.font = f;
          const fill = fills[xf.fillId];
          if (fill) cell.fill = fill.color;
          const bd = borders[xf.borderId];
          if (bd && Object.keys(bd).length) {
            const vals = Object.values(bd);
            cell.border = vals.length === 4 && new Set(vals).size === 1 ? vals[0] : bd;
          }
          if (xf.align && Object.keys(xf.align).length) cell.align = xf.align;
        }
        if (colLetters) cell.col = colLetters;
        if (fmt) cell.numFmt = fmt;
        if (formula) cell.f = `=${formula}`;
        if (raw !== undefined) cell.raw = raw;
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

const PPT_ALIGN = { l: "Left", ctr: "Center", r: "Right", just: "Justified", dist: "Distributed" };
const emuToPt = (v) => (v != null && Number.isFinite(Number(v)) ? Math.round((Number(v) / 12700) * 10) / 10 : undefined);

/** 解析主题字体（majorFont/minorFont），用于还原 +mn-lt / +mj-ea 等主题字体引用 */
function parseThemeFonts(themeXml) {
  const major = themeXml.match(/<a:majorFont>[\s\S]*?<\/a:majorFont>/)?.[0] || "";
  const minor = themeXml.match(/<a:minorFont>[\s\S]*?<\/a:minorFont>/)?.[0] || "";
  return {
    "+mj-lt": major.match(/<a:latin typeface="([^"]*)"/)?.[1],
    "+mj-ea": major.match(/<a:ea typeface="([^"]*)"/)?.[1],
    "+mn-lt": minor.match(/<a:latin typeface="([^"]*)"/)?.[1],
    "+mn-ea": minor.match(/<a:ea typeface="([^"]*)"/)?.[1],
  };
}

/** 解析运行属性（a:rPr 标签属性 + 内部子元素）→ 文字格式 { sizePt, bold, italic, underline, color, name } */
function pptRunFont(runXml, themeColors, themeFonts) {
  const tag = runXml.match(/<(?:[\w]+:)?rPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w]+:)?rPr>)/);
  if (!tag) return null;
  const attrs = `<a ${tag[1]}>`;
  const inner = tag[2] || "";
  const font = {};
  const sz = attr(attrs, "sz");
  if (sz) font.sizePt = Number(sz) / 100;
  if (attr(attrs, "b") === "1") font.bold = true;
  if (attr(attrs, "i") === "1") font.italic = true;
  const u = attr(attrs, "u");
  if (u && u !== "none") font.underline = u === "sng" ? "single" : u;
  const color = drawColor(inner.match(/<(?:[\w]+:)?solidFill\b[\s\S]*?<\/(?:[\w]+:)?solidFill>/)?.[0], themeColors);
  if (color) font.color = color;
  const latin = inner.match(/<(?:[\w]+:)?latin\b[^>]*typeface="([^"]+)"/)?.[1];
  const ea = inner.match(/<(?:[\w]+:)?ea\b[^>]*typeface="([^"]+)"/)?.[1];
  const resolve = (v) => (v && v.startsWith("+") ? themeFonts?.[v] : v);
  const name = resolve(ea) || (latin && !/^\+/.test(latin) ? latin : resolve(latin));
  if (name) font.name = name;
  return Object.keys(font).length ? font : null;
}

function parsePptx(buf) {
  const zip = openZip(buf);
  const slideFiles = Object.keys(zip)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  if (!slideFiles.length) throw new Error("演示文稿中没有幻灯片，无法解析");

  const themeXml = zip["ppt/theme/theme1.xml"] ? strFromU8(zip["ppt/theme/theme1.xml"]) : "";
  const themeColors = parseThemeColors(themeXml);
  const themeFonts = parseThemeFonts(themeXml);

  const slides = slideFiles.map((file, i) => {
    const xml = strFromU8(zip[file]);
    // 版式名与关系表（rels：layout / 图片媒体 / 图表数据）
    let layout;
    let relsXml = "";
    const relFile = `ppt/slides/_rels/${file.split("/").pop()}.rels`;
    if (zip[relFile]) relsXml = strFromU8(zip[relFile]);
    if (relsXml) {
      const target = relsXml.match(/Target="([^"]*slideLayout[^"]*)"/)?.[1];
      if (target) {
        const lp = ("ppt/" + target.replace(/^\.\.\//, "").replace(/^\//, "")).replace("ppt/ppt/", "ppt/");
        const lxml = zip[lp] ? strFromU8(zip[lp]) : "";
        layout = lxml?.match(/<(?:[\w]+:)?cSld\b[^>]*name="([^"]*)"/)?.[1] || undefined;
      }
    }
    const resolveRel = (rid) => {
      if (!rid || !relsXml) return null;
      const target =
        relsXml.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`))?.[1] ||
        relsXml.match(new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`))?.[1];
      if (!target) return null;
      return ("ppt/" + target.replace(/^\.\.\//, "").replace(/^\//, "")).replace("ppt/ppt/", "ppt/");
    };

    const shapes = [];
    const tables = [];
    const charts = [];
    const fonts = new Set();
    const sizes = new Set();
    let boldUsed = false;

    // 形状枚举：sp（文本）/ pic（图片）/ graphicFrame（表格/图表）；组合 grpSp 内的子形状被平铺枚举
    const elRe = /<(?:[\w]+:)?(sp|pic|graphicFrame)\b[^>]*>[\s\S]*?<\/(?:[\w]+:)?\1>/g;
    let m, pm, rmm;
    while ((m = elRe.exec(xml))) {
      if (shapes.length >= 80) break;
      const block = m[0];
      const name = unescXml(block.match(/<(?:[\w]+:)?cNvPr\b[^>]*name="([^"]*)"/)?.[1] || "");
      const phTag = block.match(/<(?:[\w]+:)?ph\b[^>]*\/?>/)?.[0];
      const phType = phTag ? attr(phTag, "type") : undefined;
      const off = block.match(/<(?:[\w]+:)?off\b[^>]*\/?>/)?.[0];
      const ext = block.match(/<(?:[\w]+:)?ext\b[^>]*\/?>/)?.[0];
      const x = emuToPt(attr(off, "x"));
      const y = emuToPt(attr(off, "y"));
      const w = emuToPt(attr(ext, "cx"));
      const h = emuToPt(attr(ext, "cy"));

      if (m[1] === "graphicFrame") {
        const tblXml = block.match(/<(?:[\w]+:)?tbl\b[\s\S]*?<\/(?:[\w]+:)?tbl>/)?.[0];
        const chartRid = block.match(/<(?:[\w]+:)?chart\b[^>]*r:id="([^"]+)"/)?.[1];
        if (tblXml) {
          const rows = [];
          const trRe = /<(?:[\w]+:)?tr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?tr>/g;
          let tr;
          while ((tr = trRe.exec(tblXml)) && rows.length < 100) {
            const cells = [];
            const tcRe = /<(?:[\w]+:)?tc\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?tc>|<(?:[\w]+:)?tc\b[^>]*\/>/g;
            let tc;
            while ((tc = tcRe.exec(tr[1])) && cells.length < 40) {
              let text = "";
              const tRe = /<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g;
              let tm;
              while ((tm = tRe.exec(tc[1] || ""))) text += unescXml(tm[1]);
              cells.push(text.trim());
            }
            rows.push(cells);
          }
          const tblPr = tblXml.match(/<(?:[\w]+:)?tblPr\b[^>]*\/?>/)?.[0] || "";
          tables.push({
            name,
            ...(x != null ? { x, y, w, h } : {}),
            ...(attr(tblPr, "firstRow") === "1" ? { headerRow: true } : {}),
            ...(attr(tblPr, "bandRow") === "1" ? { bandedRows: true } : {}),
            rows,
            rowCount: rows.length,
          });
          shapes.push({ name, kind: "table", ph: phType, x, y, w, h, rows: rows.length, cols: rows[0]?.length || 0, paragraphs: [], paraFonts: [] });
        } else if (chartRid) {
          // 图表：rels → chart xml → 标题与序列缓存数据
          const chartPath = resolveRel(chartRid);
          const cxml = chartPath && zip[chartPath] ? strFromU8(zip[chartPath]) : "";
          const title = cxml.match(/<c:title\b[\s\S]*?<\/c:title>/)?.[0] || "";
          const series = [];
          const serRe = /<c:ser\b[\s\S]*?<\/c:ser>/g;
          let ser;
          while ((ser = serRe.exec(cxml)) && series.length < 10) {
            const body = ser[0];
            const cachePts = (seg) => {
              const pts = [];
              const ptRe = /<c:pt\b[^>]*>([\s\S]*?)<\/c:pt>/g;
              let pt;
              while ((pt = ptRe.exec(seg)) && pts.length < 100) pts.push(unescXml(pt[1].match(/<c:v\b[^>]*>([\s\S]*?)<\/c:v>/)?.[1] || ""));
              return pts;
            };
            const nameV = body.match(/<c:tx\b[\s\S]*?<\/c:tx>/)?.[0];
            const cats = body.match(/<c:cat\b[\s\S]*?<\/c:cat>/)?.[0];
            const vals = body.match(/<c:val\b[\s\S]*?<\/c:val>/)?.[0];
            series.push({
              name: nameV ? unescXml(nameV.match(/<c:v\b[^>]*>([\s\S]*?)<\/c:v>/)?.[1] || "") || `系列 ${series.length + 1}` : `系列 ${series.length + 1}`,
              ...(cats ? { categories: cachePts(cats) } : {}),
              values: vals ? cachePts(vals) : [],
            });
          }
          charts.push({
            name,
            ...(x != null ? { x, y, w, h } : {}),
            ...(title ? { title: unescXml(title.match(/<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/)?.[1] || "") } : {}),
            series,
          });
          shapes.push({ name, kind: "chart", ph: phType, x, y, w, h, paragraphs: [], paraFonts: [] });
        }
        continue;
      }

      if (m[1] === "pic") {
        const embed = block.match(/<(?:[\w]+:)?blip\b[^>]*r:embed="([^"]+)"/)?.[1];
        const media = embed ? resolveRel(embed) : null;
        shapes.push({
          name,
          kind: "picture",
          ph: phType,
          ...(x != null ? { x, y, w, h } : {}),
          ...(media ? { image: media.replace(/^ppt\//, "") } : {}),
          paragraphs: [],
          paraFonts: [],
        });
        continue;
      }

      // 文本形状：spPr 作用域取填充/线条，txBody 取段落与文字格式
      const spPr = block.match(/<(?:[\w]+:)?spPr\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?spPr>/)?.[1] || "";
      const fill = /<(?:[\w]+:)?noFill\b/.test(spPr) ? "无" : drawColor(spPr.match(/<(?:[\w]+:)?solidFill\b[\s\S]*?<\/(?:[\w]+:)?solidFill>/)?.[0], themeColors);
      const lnTag = spPr.match(/<(?:[\w]+:)?ln\b[^>]*>/)?.[0];
      const lnXml = spPr.match(/<(?:[\w]+:)?ln\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?ln>/)?.[0];
      const lineColor = drawColor(lnXml, themeColors);
      const lineW = emuToPt(attr(lnTag, "w"));
      const line = lineColor || lineW != null ? { ...(lineColor ? { color: lineColor } : {}), ...(lineW != null ? { wPt: lineW } : {}) } : null;
      const anchor = attr(block.match(/<(?:[\w]+:)?bodyPr\b[^>]*>/)?.[0], "anchor");

      const paragraphs = [];
      const paraFonts = [];
      const pRe = /<(?:[\w]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?p>/g;
      while ((pm = pRe.exec(block.match(/<(?:[\w]+:)?txBody\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?txBody>/)?.[1] || ""))) {
        if (paragraphs.length >= 60) break;
        const pBody = pm[1];
        let text = "";
        let font = null;
        const rRe = /<(?:[\w]+:)?r\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?r>/g;
        while ((rmm = rRe.exec(pBody))) {
          const t = unescXml(rmm[1].match(/<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/)?.[1] || "");
          if (!t.trim() && !font) continue;
          if (!font) font = pptRunFont(rmm[1], themeColors, themeFonts);
          text += t;
        }
        if (!text) continue;
        const algn = attr(pBody.match(/<(?:[\w]+:)?pPr\b[^>]*>/)?.[0], "algn");
        paragraphs.push(text);
        paraFonts.push({ ...(algn && PPT_ALIGN[algn] ? { align: PPT_ALIGN[algn] } : {}), ...(font || {}) });
        if (font?.name) fonts.add(font.name);
        if (font?.sizePt) sizes.add(font.sizePt);
        if (font?.bold) boldUsed = true;
      }
      shapes.push({
        name,
        kind: "text",
        ph: phType,
        ...(x != null ? { x, y, w, h } : {}),
        ...(fill ? { fill } : {}),
        ...(line ? { line } : {}),
        ...(anchor && anchor !== "t" ? { anchor } : {}),
        paragraphs,
        paraFonts,
      });
    }

    const n = i + 1;
    return {
      n,
      ...(layout ? { layout } : {}),
      shapes,
      tables,
      charts,
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
      note: "sheet 参数选表（名称或序号），mode=text 读单元格值/公式（数字、日期已转换），mode=format 读数字格式/字体/对齐/填充色/边框",
    };
  }
  if (parsed.type === "pptx") {
    const tableCount = parsed.slides.reduce((n, s) => n + s.tables.length, 0);
    const chartCount = parsed.slides.reduce((n, s) => n + s.charts.length, 0);
    return {
      kind: parsed.kind,
      slideCount: parsed.slideCount,
      ...(tableCount ? { tableCount } : {}),
      ...(chartCount ? { chartCount } : {}),
      slides: parsed.slides.map((s) => ({
        n: s.n,
        layout: s.layout,
        shapes: s.shapes.length,
        firstText: (s.shapes.find((sp) => sp.paragraphs.length)?.paragraphs[0] || "").slice(0, 50),
      })),
      note: "mode=structure 读形状清单；mode=text 读文本/表格数据/图表数据；mode=format 读形状位置尺寸、填充边框、逐段文字格式；mode=tables 读表格",
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
      // mode=format 保留单元格全部格式信息（数字格式/字体/对齐/填充/边框/raw 日期值）；text 仅取值与公式
      rows: rows.map((r) => ({
        r: r.r,
        cells: r.cells.map((c) =>
          mode === "format"
            ? c
            : { ref: c.ref, v: c.v, ...(c.f ? { f: c.f } : {}) }
        ),
      })),
    };
  }

  if (parsed.type === "pptx") {
    const offset = clamp(args?.offset, 1, parsed.slideCount) || 1;
    const limit = clamp(args?.limit, 10, 50);
    const slice = parsed.slides.filter((s) => s.n >= offset && s.n < offset + limit);
    if (mode === "structure") {
      // 形状清单：名称/类型/占位符/尺寸 + 首行文本
      return {
        type: "pptx",
        mode,
        slideCount: parsed.slideCount,
        slides: slice.map((s) => ({
          n: s.n,
          layout: s.layout,
          shapes: s.shapes.slice(0, 40).map((sp) => ({
            name: sp.name,
            kind: sp.kind,
            ...(sp.ph ? { placeholder: sp.ph } : {}),
            ...(sp.w != null ? { sizePt: [sp.w, sp.h] } : {}),
            ...(sp.kind === "table" ? { rows: sp.rows, cols: sp.cols } : {}),
            ...(sp.image ? { image: sp.image } : {}),
            firstText: (sp.paragraphs[0] || "").slice(0, 60) || undefined,
          })),
          ...(s.tables.length ? { tableCount: s.tables.length } : {}),
          ...(s.charts.length ? { chartCount: s.charts.length } : {}),
        })),
      };
    }
    if (mode === "tables") {
      const tables = [];
      for (const s of slice)
        for (const t of s.tables)
          tables.push({ slide: s.n, name: t.name, ...(t.headerRow ? { headerRow: true } : {}), rows: t.rows.slice(0, 40), rowCount: t.rowCount });
      return { type: "pptx", mode, slideCount: parsed.slideCount, tables };
    }
    if (mode === "format") {
      // 形状级格式：位置尺寸/填充/边框 + 逐段文字格式（段落数组与 paraFonts 并行合并）
      return {
        type: "pptx",
        mode,
        slideCount: parsed.slideCount,
        slides: slice.map((s) => ({
          n: s.n,
          layout: s.layout,
          shapes: s.shapes.slice(0, 25).map((sp) => ({
            name: sp.name,
            kind: sp.kind,
            ...(sp.ph ? { placeholder: sp.ph } : {}),
            ...(sp.x != null ? { posPt: { x: sp.x, y: sp.y, w: sp.w, h: sp.h } } : {}),
            ...(sp.fill ? { fill: sp.fill } : {}),
            ...(sp.line ? { line: sp.line } : {}),
            ...(sp.anchor ? { anchor: sp.anchor } : {}),
            ...(sp.kind === "table" ? { style: { ...(sp.rows ? { rows: sp.rows } : {}), ...(s.tables.find((t) => t.name === sp.name)?.headerRow ? { headerRow: true } : {}) } } : {}),
            paragraphs: sp.paragraphs.slice(0, 25).map((text, j) => {
              const f = sp.paraFonts[j] || {};
              return {
                text,
                ...(f.align ? { align: f.align } : {}),
                ...(f.sizePt || f.bold || f.italic || f.underline || f.color || f.name
                  ? { font: { ...(f.name ? { name: f.name } : {}), ...(f.sizePt ? { sizePt: f.sizePt } : {}), ...(f.bold ? { bold: true } : {}), ...(f.italic ? { italic: true } : {}), ...(f.underline ? { underline: f.underline } : {}), ...(f.color ? { color: f.color } : {}) } }
                  : {}),
              };
            }),
          })),
          ...(s.tables.length ? { tables: s.tables.map((t) => ({ name: t.name, ...(t.headerRow ? { headerRow: true } : {}), ...(t.bandedRows ? { bandedRows: true } : {}) })) } : {}),
        })),
      };
    }
    // text：文本 + 表格数据 + 图表数据 + 备注
    return {
      type: "pptx",
      mode: "text",
      slideCount: parsed.slideCount,
      slides: slice.map((s) => ({
        n: s.n,
        layout: s.layout,
        shapes: s.shapes.slice(0, 40).map((sp) => ({
          name: sp.name,
          kind: sp.kind,
          ...(sp.ph ? { placeholder: sp.ph } : {}),
          ...(sp.image ? { image: sp.image } : {}),
          ...(sp.paragraphs.length ? { paragraphs: sp.paragraphs } : {}),
        })),
        ...(s.tables.length ? { tables: s.tables.map((t) => ({ name: t.name, ...(t.headerRow ? { headerRow: true } : {}), rows: t.rows.slice(0, 40), rowCount: t.rowCount })) } : {}),
        ...(s.charts.length ? { charts: s.charts.map((ch) => ({ name: ch.name, ...(ch.title ? { title: ch.title } : {}), ...(ch.x != null ? { posPt: { x: ch.x, y: ch.y, w: ch.w, h: ch.h } } : {}), series: ch.series })) } : {}),
        ...(s.notes ? { notes: s.notes } : {}),
      })),
    };
  }

  throw new Error("未知的附件类型");
}
