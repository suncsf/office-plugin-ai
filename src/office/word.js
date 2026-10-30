/**
 * Word 执行器（Office.js Word API）
 * 每个 executor 返回对象/字符串；异常向上抛出由 executeTool 统一包装。
 *
 * ⚠️ insertLocation 约定（官方签名核实，WordApi 1.8）：
 *  - Range/Selection: insertText → Replace/Start/End/Before/After；
 *    insertParagraph / insertTable / insertBreak → 仅 Before/After
 *  - Body: insertText → Replace/Start/End；insertParagraph/insertTable/insertBreak → 仅 Start/End
 *  - select(selectionMode) → 仅 "Select" | "Start" | "End"
 */
import { stripDataUrl } from "./dataurl";
import { resolveStyleRequest, normalizeAlignment, BUILT_IN_STYLES } from "./style-map";

/* ---------------- 内部工具函数 ---------------- */

function textOf(content) {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map((p) => (typeof p === "string" ? p : p?.text || "")).join("");
  return String(content ?? "");
}

function countWords(text) {
  const t = String(text || "");
  const noSpace = t.replace(/\s+/g, "");
  const cjk = (noSpace.match(/[\u3000-\u9fff\uff00-\uffef]/g) || []).length;
  const latin = (t.match(/[A-Za-z0-9]+/g) || []).length;
  return { chars: t.length, charsNoSpace: noSpace.length, words: cjk + latin };
}

async function loadParagraphs(ctx, select) {
  const paras = ctx.document.body.paragraphs;
  paras.load(select);
  await ctx.sync();
  return paras;
}

/** 定位包含 selection 的段落 index（0 起始） */
async function findSelectionParagraphIndex(ctx, sel, paras) {
  try {
    const cmps = paras.items.map((p) => p.getRange("Whole").compareLocationWith(sel));
    await ctx.sync();
    const vals = cmps.map((c) => c.value);
    let idx = vals.findIndex((v) => v === "Contains" || v === "Equal" || v === "Inside" || v === "Contained");
    if (idx < 0) idx = vals.findIndex((v) => v === "Unknown");
    return idx;
  } catch {
    // 兜底：按选区文本匹配
    const t = textOf(sel.text).slice(0, 60);
    if (!t) return -1;
    return paras.items.findIndex((p) => (p.text || "").includes(t));
  }
}

function fontInfo(font, style, styleBuiltIn, alignment) {
  const info = {
    style: style ?? undefined,
    styleBuiltIn: styleBuiltIn ?? undefined,
    alignment: alignment ?? undefined,
    fontName: font?.name, fontSize: font?.size,
    bold: font?.bold, italic: font?.italic, underline: font?.underline,
    color: font?.color, highlightColor: font?.highlightColor,
  };
  Object.keys(info).forEach((k) => info[k] === undefined && delete info[k]);
  return info;
}

/** 提取 Office.js 错误的可读信息（含错误码与出错位置，便于诊断） */
function describeErr(err) {
  const msg = err?.message || String(err);
  const code = err?.code && err.code !== msg ? String(err.code) : err?.name && err.name !== "Error" && err.name !== msg ? String(err.name) : null;
  const loc = err?.debugInfo?.errorLocation || err?.innerError?.debugInfo?.errorLocation;
  const extra = [code, loc ? `位置: ${loc}` : null].filter(Boolean).join(" @ ");
  return extra ? `${msg}（${extra}）` : msg;
}

/** 颜色归一化：缺 # 时补上 */
function normalizeColor(c) {
  const s = String(c || "").trim();
  return s && !s.startsWith("#") && !s.startsWith("rgb") ? `#${s}` : s;
}

/** 解析可选的 JSON 对象参数（接受对象或 JSON 字符串） */
function parseJsonObj(v, label) {
  if (v == null || v === "") return null;
  if (typeof v === "object") return v;
  try {
    const o = JSON.parse(v);
    if (o && typeof o === "object") return o;
  } catch { /* fallthrough */ }
  throw new Error(`${label} 必须是 JSON 对象`);
}

/** 压缩 Style 代理的已加载属性为普通对象（剔除 undefined/null） */
function pickStyleInfo(s) {
  const out = {
    name: s.nameLocal,
    builtIn: s.builtIn,
    type: s.type,
    inUse: s.inUse,
    quickStyle: s.quickStyle,
  };
  const f = {};
  if (s.font) {
    if (s.font.name) f.fontName = s.font.name;
    if (s.font.size != null) f.fontSizePt = s.font.size;
    if (s.font.bold != null) f.bold = s.font.bold;
    if (s.font.italic != null) f.italic = s.font.italic;
    if (s.font.color) f.color = s.font.color;
    if (s.font.underline != null) f.underline = s.font.underline;
  }
  const pf = {};
  if (s.paragraphFormat) {
    if (s.paragraphFormat.alignment != null) pf.alignment = s.paragraphFormat.alignment;
    if (s.paragraphFormat.outlineLevel != null) pf.outlineLevel = s.paragraphFormat.outlineLevel;
    if (s.paragraphFormat.lineSpacing != null) pf.lineSpacingPt = s.paragraphFormat.lineSpacing;
    if (s.paragraphFormat.spaceBefore != null) pf.spaceBeforePt = s.paragraphFormat.spaceBefore;
    if (s.paragraphFormat.spaceAfter != null) pf.spaceAfterPt = s.paragraphFormat.spaceAfter;
    if (s.paragraphFormat.leftIndent != null) pf.leftIndentPt = s.paragraphFormat.leftIndent;
    if (s.paragraphFormat.rightIndent != null) pf.rightIndentPt = s.paragraphFormat.rightIndent;
    if (s.paragraphFormat.firstLineIndent != null) pf.firstLineIndentPt = s.paragraphFormat.firstLineIndent;
  }
  if (Object.keys(f).length) out.font = f;
  if (Object.keys(pf).length) out.paragraphFormat = pf;
  return out;
}

/* ---------------- 执行器 ---------------- */

const registry = {};

registry["word.getDocumentInfo"] = async () => {
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    const all = paras.items.map((p) => p.text || "").join("\n");
    const wc = countWords(all);
    return {
      paragraphCount: paras.items.length,
      sectionCount: sections.items.length,
      charCount: wc.chars,
      wordCount: wc.words,
    };
  });
};

registry["word.getDocumentText"] = async (args) => {
  const max = Math.max(200, Number(args?.max_length) || 8000);
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    let text = paras.items.map((p) => p.text || "").join("\n");
    const total = text.length;
    let note = "";
    if (total > max) {
      text = text.slice(0, max);
      note = `…[已截断，全文共 ${total} 字符，可用 word_get_paragraphs 分段读取]`;
    }
    return { totalChars: total, text: text + note };
  });
};

registry["word.getParagraphs"] = async (args) => {
  const offset = Math.max(0, Number(args?.offset) || 0);
  const limit = Math.min(200, Math.max(1, Number(args?.limit) || 50));
  const includeStyle = args?.include_style !== false;
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(
      ctx,
      includeStyle ? "text,style,styleBuiltIn,isListItem,outlineLevel" : "text"
    );
    const items = paras.items.slice(offset, offset + limit).map((p, i) => ({
      index: offset + i,
      text: String(p.text || "").slice(0, 200),
      style: includeStyle ? p.style : undefined,
      styleBuiltIn: includeStyle ? p.styleBuiltIn : undefined,
      isListItem: includeStyle ? p.isListItem : undefined,
      outlineLevel: includeStyle ? p.outlineLevel : undefined,
    }));
    return { total: paras.items.length, offset, returned: items.length, paragraphs: items };
  });
};

registry["word.getParagraphText"] = async (args) => {
  const index = Number(args?.index);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围（共 ${paras.items.length} 段）`);
    return { index, text: paras.items[index].text || "" };
  });
};

registry["word.getSelection"] = async () => {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text,isEmpty,style");
    sel.font.load("name,size,bold,italic,underline,color");
    await ctx.sync();
    return {
      text: sel.text || "",
      isEmpty: sel.isEmpty,
      ...fontInfo(sel.font, sel.style),
    };
  });
};

registry["word.getCursorPosition"] = async () => {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text,isEmpty");
    const paras = await loadParagraphs(ctx, "text");
    await ctx.sync();
    const idx = await findSelectionParagraphIndex(ctx, sel, paras);
    return {
      paragraphIndex: idx,
      paragraphPreview: idx >= 0 ? String(paras.items[idx].text || "").slice(0, 80) : null,
      totalParagraphs: paras.items.length,
      selection: { text: sel.text || "", isEmpty: sel.isEmpty },
    };
  });
};

registry["word.moveCursor"] = async (args) => {
  const index = Number(args?.paragraph_index);
  if (!Number.isInteger(index) || index < 0) throw new Error("paragraph_index 必须是非负整数");
  const position = args?.position === "end" ? "End" : "Start";
  const searchText = args?.search_text;
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围（共 ${paras.items.length} 段）`);
    if (searchText) {
      try {
        const results = paras.items[index].getRange("Whole").search(String(searchText), { matchCase: false });
        results.load("text");
        await ctx.sync();
        if (results.items.length) {
          results.items[0].select("Start");
          await ctx.sync();
          return { ok: true, paragraphIndex: index, movedTo: `“${searchText}”处` };
        }
      } catch { /* 忽略，走段落定位 */ }
      throw new Error(`在段落 ${index} 中未找到“${searchText}”`);
    }
    paras.items[index].select(position);
    await ctx.sync();
    return { ok: true, paragraphIndex: index, position };
  });
};

registry["word.selectText"] = async (args) => {
  const query = String(args?.query || "");
  if (!query) throw new Error("query 不能为空");
  const occurrence = Math.max(1, Number(args?.occurrence) || 1);
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(query, {
      matchCase: !!args?.match_case,
      matchWholeWord: !!args?.match_whole_word,
    });
    results.load("text");
    await ctx.sync();
    if (!results.items.length) throw new Error(`未找到“${query}”`);
    const idx = Math.min(occurrence, results.items.length) - 1;
    results.items[idx].select();
    await ctx.sync();
    return { ok: true, totalMatches: results.items.length, selected: results.items[idx].text, occurrence: idx + 1 };
  });
};

registry["word.selectParagraph"] = async (args) => {
  const index = Number(args?.index);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围（共 ${paras.items.length} 段）`);
    // 官方枚举: "Select" | "Start" | "End"（没有 "Selected"）
    paras.items[index].select("Select");
    await ctx.sync();
    return { ok: true, index, text: String(paras.items[index].text || "").slice(0, 120) };
  });
};

registry["word.selectAll"] = async () => {
  return Word.run(async (ctx) => {
    ctx.document.body.select("Select");
    await ctx.sync();
    return { ok: true };
  });
};

async function insertTextCore(args) {
  const text = String(args?.text ?? "");
  if (!text) throw new Error("text 不能为空");
  const location = args?.location || "cursor";
  return Word.run(async (ctx) => {
    if (location === "cursor") {
      const sel = ctx.document.getSelection();
      const r = sel.insertText(text, "Replace");
      await ctx.sync();
      return { ok: true, inserted: text.length, at: "cursor" };
    }
    const body = ctx.document.body;
    body.insertText(text, location === "start" ? "Start" : "End");
    await ctx.sync();
    return { ok: true, inserted: text.length, at: location };
  });
}

registry["word.insertText"] = insertTextCore;

registry["word.insertParagraph"] = async (args) => {
  const text = String(args?.text ?? "");
  return Word.run(async (ctx) => {
    if (args?.location === "end") {
      // Body.insertParagraph 仅允许 "Start" | "End"
      ctx.document.body.insertParagraph(text, "End");
    } else {
      // Range/Selection.insertParagraph 仅允许 "Before" | "After"（"Replace" 非法）
      // 选区为空（光标）时 "After" 即在光标处插入新段落
      ctx.document.getSelection().insertParagraph(text, "After");
    }
    await ctx.sync();
    return { ok: true, text };
  });
};

registry["word.insertPageBreak"] = async () => {
  return Word.run(async (ctx) => {
    // Range.insertBreak 仅允许 "Before" | "After"；光标处用 "After"
    ctx.document.getSelection().insertBreak("Page", "After");
    await ctx.sync();
    return { ok: true, action: "已插入分页符（新建页面）" };
  });
};

registry["word.insertLineBreak"] = async () => {
  return Word.run(async (ctx) => {
    // BreakType 合法值（官方）: "Page"|"Next"|"SectionNext"|"SectionContinuous"|"SectionEven"|"SectionOdd"|"Line"
    ctx.document.getSelection().insertBreak("Line", "After");
    await ctx.sync();
    return { ok: true, action: "已插入换行符" };
  });
};

registry["word.insertSymbol"] = async (args) => {
  const symbol = String(args?.symbol || "");
  if (!symbol) throw new Error("symbol 不能为空");
  const count = Math.min(500, Math.max(1, Number(args?.count) || 1));
  return insertTextCore({ text: symbol.repeat(count), location: args?.location || "cursor" });
};

registry["word.insertHyperlink"] = async (args) => {
  const url = String(args?.url || "");
  const text = String(args?.text || "");
  if (!url) throw new Error("url 不能为空");
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("isEmpty");
    await ctx.sync();
    if (sel.isEmpty) {
      if (!text) throw new Error("选区为空时必须提供 text");
      // 先插入文字，再通过 hyperlink 属性设置链接（Range 没有 insertHyperlink 方法）
      const r = sel.insertText(text, "Replace");
      r.hyperlink = url;
    } else {
      // 直接把选中文本设为链接（Range.hyperlink 可写）
      sel.hyperlink = url;
    }
    await ctx.sync();
    return { ok: true, url };
  });
};

registry["word.insertImage"] = async (args) => {
  const b64 = stripDataUrl(args?.base64);
  if (!b64) throw new Error("base64 不能为空");
  return Word.run(async (ctx) => {
    const pic = ctx.document.getSelection().insertInlinePictureFromBase64(b64, "Replace");
    if (args?.width) pic.width = Number(args.width);
    if (args?.height) pic.height = Number(args.height);
    await ctx.sync();
    return { ok: true };
  });
};

registry["word.insertTable"] = async (args) => {
  const rows = Math.max(1, Number(args?.rows) || 1);
  const cols = Math.max(1, Number(args?.cols) || 1);
  const location = args?.location === "end" ? "end" : "cursor";
  return Word.run(async (ctx) => {
    let data;
    if (Array.isArray(args?.data)) {
      data = new Array(rows).fill(null).map((_, r) =>
        new Array(cols).fill(null).map((__, c) => String(args.data?.[r]?.[c] ?? ""))
      );
    }
    if (location === "end") {
      // Body.insertTable 仅允许 "Start" | "End"
      ctx.document.body.insertTable(rows, cols, "End", data);
    } else {
      // Range/Selection.insertTable 仅允许 "Before" | "After"（"End" 非法）
      ctx.document.getSelection().insertTable(rows, cols, "After", data);
    }
    await ctx.sync();
    return { ok: true, rows, cols };
  });
};

registry["word.replaceText"] = async (args) => {
  const search = String(args?.search || "");
  if (!search) throw new Error("search 不能为空");
  const replace = String(args?.replace ?? "");
  const replaceAll = args?.replace_all !== false;
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(search, { matchCase: !!args?.match_case });
    results.load("text");
    await ctx.sync();
    const total = results.items.length;
    if (!total) return { ok: true, replaced: 0, totalFound: 0, note: "未找到匹配文本" };
    const targets = replaceAll ? results.items : results.items.slice(0, 1);
    for (const r of targets) r.insertText(replace, "Replace");
    await ctx.sync();
    return { ok: true, replaced: targets.length, totalFound: total };
  });
};

registry["word.deleteSelection"] = async () => {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("isEmpty");
    await ctx.sync();
    if (sel.isEmpty) throw new Error("当前没有选中内容");
    try {
      sel.delete();
    } catch {
      sel.insertText("", "Replace");
    }
    await ctx.sync();
    return { ok: true };
  });
};

registry["word.deleteParagraph"] = async (args) => {
  const index = Number(args?.index);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Word.run(async (ctx) => {
    const paras = await loadParagraphs(ctx, "text");
    if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围（共 ${paras.items.length} 段）`);
    const text = String(paras.items[index].text || "").slice(0, 80);
    paras.items[index].delete();
    await ctx.sync();
    return { ok: true, deletedIndex: index, deletedTextPreview: text };
  });
};

registry["word.getFormatting"] = async (args) => {
  const target = args?.target === "paragraph" ? "paragraph" : "selection";
  return Word.run(async (ctx) => {
    // 文档默认字体（正文/Normal 的基础字体，全局格式）
    const bodyFont = ctx.document.body.font;
    bodyFont.load("name,size,color");

    const PARA_PROPS = "alignment,leftIndent,rightIndent,firstLineIndent,lineSpacing,spaceBefore,spaceAfter,outlineLevel";
    const fmtOf = (p) => {
      const out = {};
      for (const k of PARA_PROPS.split(",")) {
        const v = p[k];
        if (v != null) out[k] = v;
      }
      return out;
    };

    if (target === "paragraph") {
      const index = Number(args?.index);
      if (!Number.isInteger(index) || index < 0) throw new Error("target=paragraph 时必须提供 index");
      const paras = await loadParagraphs(
        ctx,
        `text,style,styleBuiltIn,${PARA_PROPS}`
      );
      if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围`);
      const p = paras.items[index];
      p.font.load("name,size,bold,italic,underline,color,highlightColor");
      await ctx.sync();
      return {
        target,
        index,
        documentDefaultFont: { name: bodyFont.name, sizePt: bodyFont.size, color: bodyFont.color },
        ...fontInfo(p.font, p.style, p.styleBuiltIn, p.alignment),
        paragraphFormat: fmtOf(p),
      };
    }
    const sel = ctx.document.getSelection();
    sel.load("text,isEmpty,style,styleBuiltIn");
    sel.font.load("name,size,bold,italic,underline,color,highlightColor");
    const p = sel.paragraphs.getFirst();
    p.load(PARA_PROPS);
    await ctx.sync();
    return {
      target,
      selectionText: sel.text || "",
      documentDefaultFont: { name: bodyFont.name, sizePt: bodyFont.size, color: bodyFont.color },
      ...fontInfo(sel.font, sel.style, sel.styleBuiltIn, p.alignment),
      paragraphFormat: fmtOf(p),
    };
  });
};

registry["word.setFormatting"] = async (args) => {
  const align = args?.alignment ? normalizeAlignment(args.alignment) : null;
  if (args?.alignment && !align) {
    throw new Error(`无效的 alignment: "${args.alignment}"，合法值: Left / Centered / Right / Justified`);
  }
  const hasAny =
    args?.font_name || args?.font_size != null || args?.bold != null || args?.italic != null ||
    args?.underline || args?.color || args?.highlight_color !== undefined || align;
  if (!hasAny) throw new Error("未提供任何格式属性");
  const index = args?.paragraph_index != null ? Number(args.paragraph_index) : null;
  return Word.run(async (ctx) => {
    let font, paragraphs;
    if (index !== null) {
      const paras = await loadParagraphs(ctx, "text");
      if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围`);
      font = paras.items[index].font;
      paragraphs = [paras.items[index]];
    } else {
      const sel = ctx.document.getSelection();
      font = sel.font;
      paragraphs = sel.paragraphs;
      paragraphs.load("items");
    }
    if (args?.font_name) font.name = String(args.font_name);
    if (args?.font_size != null) font.size = Number(args.font_size);
    if (args?.bold != null) font.bold = !!args.bold;
    if (args?.italic != null) font.italic = !!args.italic;
    if (args?.underline) font.underline = args.underline === true ? "Single" : String(args.underline);
    if (args?.color) font.color = String(args.color);
    if (args?.highlight_color !== undefined) {
      try {
        font.highlightColor = args.highlight_color === null ? "null" : String(args.highlight_color);
      } catch { /* 部分版本不支持清除 */ }
    }
    if (align) {
      await ctx.sync();
      for (const p of paragraphs.items || paragraphs) p.alignment = align;
    }
    await ctx.sync();
    return { ok: true, appliedTo: index !== null ? `段落 ${index}` : "当前选区" };
  });
};

/** 解析目标段落：指定 index 或当前选区内的段落 */
async function resolveTargetParagraphs(ctx, index) {
  if (index !== null) {
    const paras = await loadParagraphs(ctx, "text");
    if (index >= paras.items.length) throw new Error(`段落 index ${index} 超出范围（共 ${paras.items.length} 段）`);
    return [paras.items[index]];
  }
  const sel = ctx.document.getSelection();
  const paras = sel.paragraphs;
  paras.load("text");
  await ctx.sync();
  if (!paras.items.length) throw new Error("当前选区内没有段落");
  return paras.items;
}

registry["word.applyStyle"] = async (args) => {
  const req = resolveStyleRequest(args?.style);
  if (!req) throw new Error("style 不能为空");
  const index = args?.paragraph_index != null ? Number(args.paragraph_index) : null;
  if (index !== null && (!Number.isInteger(index) || index < 0)) {
    throw new Error("paragraph_index 必须是非负整数");
  }

  /**
   * 尝试序列（每次尝试使用独立的 Word.run，失败互不影响）：
   *  内置样式：① styleBuiltIn = 官方枚举值（如 "Heading1"，跨语言可移植，主路径）
   *            ② style = 中文显示名（如 "标题 1"，style 属性按本地化名匹配）
   *            ③ style = 英文显示名
   *  自定义样式：style = 原样名称（文档中必须已存在该样式）
   * 注意：Word.js 属性赋值只是排队命令，错误在 ctx.sync() 时才抛出，
   *       因此不能在赋值处 try/catch，必须包裹 sync。
   */
  const attempts = req.builtin
    ? [
        { prop: "styleBuiltIn", value: req.builtin.value, label: `styleBuiltIn="${req.builtin.value}"（官方枚举）` },
        { prop: "style", value: req.builtin.zh, label: `style="${req.builtin.zh}"（本地化名）` },
        { prop: "style", value: req.builtin.en, label: `style="${req.builtin.en}"（英文名）` },
      ]
    : [{ prop: "style", value: req.custom, label: `style="${req.custom}"（自定义样式）` }];

  const errors = [];
  for (const at of attempts) {
    try {
      const actual = await Word.run(async (ctx) => {
        const paras = await resolveTargetParagraphs(ctx, index);
        for (const p of paras) p[at.prop] = at.value;
        await ctx.sync();
        // 回读验证，确认真实生效的样式
        const first = paras[0];
        first.load("style,styleBuiltIn");
        await ctx.sync();
        return { style: first.style, styleBuiltIn: first.styleBuiltIn };
      });
      return {
        ok: true,
        requested: String(args?.style),
        appliedVia: at.label,
        actualStyle: actual,
        appliedTo: index !== null ? `段落 ${index}` : "当前选区",
        paragraphCount: undefined,
        failedAttempts: errors.length ? errors : undefined,
      };
    } catch (e) {
      errors.push(`${at.label} → ${describeErr(e)}`);
    }
  }
  throw new Error(`应用样式失败（已尝试 ${attempts.length} 种方式）：${errors.join("；")}`);
};

const STYLE_LOAD_PROPS =
  "nameLocal,builtIn,type,inUse,quickStyle,visibility,priority," +
  "font/name,font/size,font/bold,font/italic,font/underline,font/color," +
  "paragraphFormat/alignment,paragraphFormat/outlineLevel,paragraphFormat/lineSpacing," +
  "paragraphFormat/spaceBefore,paragraphFormat/spaceAfter,paragraphFormat/leftIndent," +
  "paragraphFormat/rightIndent,paragraphFormat/firstLineIndent";

registry["word.listStyles"] = async (args) => {
  const name = args?.name ? String(args.name) : null;
  return Word.run(async (ctx) => {
    // 旧版本 Word JS 无 getStyles（WordApi 1.5+）：回退到内置样式表
    if (typeof ctx.document.getStyles !== "function") {
      return {
        source: "fallback-builtin",
        note: "当前 Word JS API 不支持枚举样式库（需要 WordApi 1.5+，Office 2024/新版 365 支持），以下为 Word 内置样式（跨语言可用）",
        total: BUILT_IN_STYLES.length,
        styles: BUILT_IN_STYLES.map((s) => ({
          name: s.value,
          zhName: s.zh,
          builtIn: true,
          type: "Paragraph",
          applyHint: `word_apply_style 的 style 可传：${s.value} / ${s.zh}`,
        })),
      };
    }

    const styles = ctx.document.getStyles();

    // 查询单个样式详情
    if (name) {
      const s = styles.getByNameOrNullObject(name);
      s.load(`isNullObject,${STYLE_LOAD_PROPS}`);
      await ctx.sync();
      if (s.isNullObject) {
        return {
          found: false,
          name,
          note: "文档样式中不存在该名称；可先不带 name 调用本工具查看全部可用样式，或用 word_create_style 创建自定义样式",
        };
      }
      return { found: true, style: pickStyleInfo(s) };
    }

    // 枚举样式库
    styles.load(`items/${STYLE_LOAD_PROPS.split(",").join(",items/")}`);
    await ctx.sync();
    let items = styles.items.slice();
    const typeFilter = args?.type && args.type !== "all" ? String(args.type) : null;
    if (typeFilter) items = items.filter((s) => s.type === typeFilter);
    if (args?.only_in_use) items = items.filter((s) => s.inUse);
    // 排序：快速样式区 > 已使用 > 内置 > 其他
    items.sort((a, b) => (b.quickStyle - a.quickStyle) || (b.inUse - a.inUse) || (a.builtIn - b.builtIn));
    const limit = Math.min(200, Math.max(1, Number(args?.limit) || 60));
    return {
      total: styles.items.length,
      matched: items.length,
      returned: Math.min(items.length, limit),
      note: "name 为本地化名称，可直接传给 word_apply_style",
      styles: items.slice(0, limit).map(pickStyleInfo),
    };
  });
};

registry["word.createStyle"] = async (args) => {
  const name = String(args?.name || "").trim();
  if (!name) throw new Error("name 不能为空");
  const type = ["Paragraph", "Character", "Table", "List"].includes(args?.type) ? args.type : "Paragraph";
  const fontSpec = parseJsonObj(args?.font, "font");
  const pfSpec = parseJsonObj(args?.paragraph_format, "paragraph_format");
  const align = pfSpec?.alignment ? normalizeAlignment(pfSpec.alignment) : null;
  if (pfSpec?.alignment && !align) {
    throw new Error(`无效的 alignment: "${pfSpec.alignment}"，合法值: Left / Centered / Right / Justified`);
  }

  return Word.run(async (ctx) => {
    if (typeof ctx.document.addStyle !== "function") {
      throw new Error("当前 Word JS API 不支持创建样式（需要 WordApi 1.5+，Office 2024/新版 365 支持）");
    }
    const styles = ctx.document.getStyles();
    const existing = styles.getByNameOrNullObject(name);
    existing.load("isNullObject,nameLocal,type");
    await ctx.sync();

    let style;
    let action;
    if (!existing.isNullObject) {
      if (!args?.update_if_exists) {
        throw new Error(`样式“${name}”已存在；如需覆盖其格式请传 update_if_exists=true，或换一个名称`);
      }
      if (existing.type !== type) {
        throw new Error(`已存在同名样式（类型 ${existing.type}），与请求类型 ${type} 不一致，请换名`);
      }
      style = existing;
      action = "updated";
    } else {
      style = ctx.document.addStyle(name, type);
      action = "created";
    }

    // 样式链与行为属性（baseStyle/nextParagraphStyle 需要 WordApi 1.6+）
    if (args?.base_style) style.baseStyle = String(args.base_style);
    if (args?.next_paragraph_style) style.nextParagraphStyle = String(args.next_paragraph_style);
    if (args?.quick_style != null) style.quickStyle = !!args.quick_style;
    await ctx.sync();

    // 字体定义（官方示例模式：直接写 style.font.*）
    if (fontSpec) {
      const f = style.font;
      if (fontSpec.name) f.name = String(fontSpec.name);
      if (fontSpec.size != null) f.size = Number(fontSpec.size);
      if (fontSpec.bold != null) f.bold = !!fontSpec.bold;
      if (fontSpec.italic != null) f.italic = !!fontSpec.italic;
      if (fontSpec.underline) f.underline = String(fontSpec.underline);
      if (fontSpec.color) f.color = normalizeColor(fontSpec.color);
    }
    // 段落格式定义（单位：磅）
    if (pfSpec) {
      const pf = style.paragraphFormat;
      if (align) pf.alignment = align;
      for (const k of ["lineSpacing", "spaceBefore", "spaceAfter", "leftIndent", "rightIndent", "firstLineIndent"]) {
        if (pfSpec[k] != null) pf[k] = Number(pfSpec[k]);
      }
    }
    await ctx.sync();

    // 回读验证
    style.load("nameLocal,type,quickStyle,font/name,font/size,font/bold,font/italic,font/color,paragraphFormat/alignment,paragraphFormat/lineSpacing,paragraphFormat/spaceBefore,paragraphFormat/spaceAfter,paragraphFormat/leftIndent,paragraphFormat/firstLineIndent");
    await ctx.sync();
    return {
      ok: true,
      action,
      style: pickStyleInfo(style),
      applyHint: `应用该样式：word_apply_style 且 style="${style.nameLocal}"`,
    };
  });
};

/* ---------------- 全量 API：通用代码执行（word_run_code） ---------------- */

/**
 * 在 Word.run 回调内执行 AI 生成的任意 Word.js 代码，开放完整 Word 对象模型，
 * 作为内置 word_* 工具未覆盖能力的兜底入口。
 * - 作用域注入 Word 与 context；代码体支持 await / return
 * - 返回值在 run 结束前完成序列化（已 load 的代理对象经 toJSON 自动展开为普通数据）
 * - 带超时保护，避免 AI 代码死循环导致工具调用永久挂起
 */
registry["word.runCode"] = async (args) => {
  const code = String(args?.code ?? "").trim();
  if (!code) throw new Error("缺少必填参数 code（在 Word.run 回调内执行的 JavaScript 源码）");
  const timeoutMs = Math.min(Math.max(Number(args?.timeout_ms) || 30000, 1000), 120000);

  let serialized;
  await Word.run(async (ctx) => {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    const fn = new AsyncFunction("Word", "context", `"use strict";\n${code}`);
    let timer;
    const work = (async () => {
      const result = await fn(Word, ctx);
      await ctx.sync(); // 收尾同步：确保队列中未提交的写入全部落地
      let out;
      try {
        out = JSON.stringify(result ?? null);
      } catch {
        out = JSON.stringify(String(result));
      }
      return out;
    })();
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(
        () => reject(new Error(`代码执行超时（${timeoutMs}ms）：请检查是否存在死循环、或读取属性前未 load/未 await context.sync()`)),
        timeoutMs
      );
    });
    try {
      serialized = await Promise.race([work, timeout]);
    } finally {
      clearTimeout(timer);
      work.catch(() => {}); // 超时后代码仍在后台收尾，吞掉迟到的 rejection
    }
  });

  try {
    return { ok: true, result: JSON.parse(serialized) };
  } catch {
    return { ok: true, result: serialized };
  }
};

export { registry };
