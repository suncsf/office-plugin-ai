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
import { resolveImageSource, saveWorkspaceFile, imageExtFromBase64 } from "./bridge";

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

registry["word.insertImage"] = async (args, ctxInfo) => {
  const b64 = await resolveImageSource(args, ctxInfo);
  if (!b64) throw new Error("base64 / file 至少提供一项");
  return Word.run(async (ctx) => {
    const pic = ctx.document.getSelection().insertInlinePictureFromBase64(b64, "Replace");
    if (args?.width) pic.width = Number(args.width);
    if (args?.height) pic.height = Number(args.height);
    await ctx.sync();
    return { ok: true };
  });
};

/* ---------------- 图片执行器（列表/提取/调整大小） ---------------- */

const round1 = (v) => Math.round(Number(v) * 10) / 10;

registry["word.listImages"] = async () => {
  return Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    pics.load("width,height");
    await ctx.sync();
    return {
      count: pics.items.length,
      images: pics.items.map((p, i) => ({ index: i, width: round1(p.width), height: round1(p.height) })),
      note:
        "index 按文档出现顺序（0 起始）。调整大小用 word_resize_image；导出为工作区文件用 word_get_image。" +
        "浮动/衬于文字下方的图片不在列表中（Word JS API 限制）。",
    };
  });
};

registry["word.getImage"] = async (args, ctxInfo) => {
  const index = Number(args?.index);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数（来自 word_list_images）");
  return Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    pics.load("width,height");
    await ctx.sync();
    if (index >= pics.items.length) {
      throw new Error(`图片 index ${index} 超出范围（共 ${pics.items.length} 张），先用 word_list_images 查看`);
    }
    const pic = pics.items[index];
    if (typeof pic.getBase64ImageSrc !== "function") {
      throw new Error("当前 Word JS API 不支持提取图片内容（getBase64ImageSrc），请更新 Office 版本");
    }
    const result = pic.getBase64ImageSrc();
    await ctx.sync();
    const b64 = stripDataUrl(result?.value || "");
    if (!b64) throw new Error("未能读取图片数据");
    const ext = imageExtFromBase64(b64) || "png";
    const rel = `images/word_img${index + 1}_${Date.now().toString(36)}.${ext}`;
    const saved = await saveWorkspaceFile(ctxInfo?.conversationId, rel, b64);
    return {
      ok: true,
      index,
      width: round1(pic.width),
      height: round1(pic.height),
      file: saved.file,
      size: saved.size,
      note: "已保存到对话工作区。可用 image_process 裁剪/缩放后，用 word_insert_image 的 file 参数插回文档。",
    };
  });
};

registry["word.resizeImage"] = async (args) => {
  const target = args?.target === "selection" ? "selection" : "index";
  const width = args?.width != null ? Number(args.width) : null;
  const height = args?.height != null ? Number(args.height) : null;
  const scale = args?.scale != null ? Number(args.scale) : null;
  const checkSize = (name, v) => {
    if (!Number.isFinite(v) || v <= 0 || v > 10000) throw new Error(`${name} 必须是 0~10000 的磅值`);
  };
  if (width != null) checkSize("width", width);
  if (height != null) checkSize("height", height);
  if (scale != null && (!Number.isFinite(scale) || scale <= 0 || scale > 20)) throw new Error("scale 必须是 0~20 的正数");
  if (width == null && height == null && scale == null) throw new Error("至少提供 width / height / scale 之一");

  return Word.run(async (ctx) => {
    let pic;
    if (target === "selection") {
      const selPics = ctx.document.getSelection().inlinePictures;
      selPics.load("width,height");
      await ctx.sync();
      if (!selPics.items.length) {
        throw new Error("当前选区中没有图片：请点击选中目标图片后重试，或改用 target=index + index 定位");
      }
      pic = selPics.items[0];
    } else {
      const index = Number(args?.index);
      if (!Number.isInteger(index) || index < 0) throw new Error("target=index 时必须提供 index（非负整数，来自 word_list_images）");
      const pics = ctx.document.body.inlinePictures;
      pics.load("width,height");
      await ctx.sync();
      if (index >= pics.items.length) {
        throw new Error(`图片 index ${index} 超出范围（共 ${pics.items.length} 张），先用 word_list_images 查看`);
      }
      pic = pics.items[index];
    }
    pic.load("width,height");
    await ctx.sync();
    const oldW = Number(pic.width);
    const oldH = Number(pic.height);
    if (!oldW || !oldH) throw new Error("未能读取图片当前尺寸");
    let newW = width;
    let newH = height;
    if (scale != null) {
      newW = oldW * scale;
      newH = oldH * scale;
    }
    // 只给一边 → 等比缩放
    if (newW != null && newH == null) newH = oldH * (newW / oldW);
    if (newH != null && newW == null) newW = oldW * (newH / oldH);
    pic.width = newW;
    pic.height = newH;
    await ctx.sync();
    return {
      ok: true,
      target,
      oldSize: { width: round1(oldW), height: round1(oldH) },
      newSize: { width: round1(newW), height: round1(newH) },
      note: "只给宽或高时按原比例等比缩放；如需非等比请同时提供 width 与 height。",
    };
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

/* ---------------- 表格编辑（WordApi 1.3+） ---------------- */

/** 校验非负整数参数 */
function requireInt(v, label) {
  const n = Number(v);
  if (!Number.isInteger(n) || n < 0) throw new Error(`${label} 必须是非负整数`);
  return n;
}

/** 校验非空一维数组 */
function parse1D(v, label) {
  if (!Array.isArray(v) || !v.length) throw new Error(`${label} 必须是非空的一维数组`);
  return v;
}

/** 校验非空二维数组 */
function parse2D(v, label) {
  if (!Array.isArray(v) || !v.length || !v.every((r) => Array.isArray(r))) {
    throw new Error(`${label} 必须是非空的二维数组（行优先，如 [["a","b"],["c","d"]]）`);
  }
  return v;
}

/** 一维数组 → 定长字符串数组（不足补空，超出截断） */
function padRow(v, len) {
  const out = v.map((x) => String(x ?? "")).slice(0, len);
  while (out.length < len) out.push("");
  return out;
}

/** 校验 (row, col) 在表格 values 范围内 */
function assertCellInRange(values, row, col) {
  const rows = values ? values.length : 0;
  if (row >= rows) throw new Error(`行号 ${row} 超出范围（表格共 ${rows} 行）`);
  const cols = values[row] ? values[row].length : 0;
  if (col >= cols) throw new Error(`列号 ${col} 超出范围（表格共 ${cols} 列）`);
}

/** 定位第 index 个表格并加载指定属性（index 0 起始，按文档出现顺序） */
async function loadTableAt(ctx, index, select) {
  const tables = ctx.document.body.tables;
  tables.load(`items/${select}`);
  await ctx.sync();
  if (!Number.isInteger(index) || index < 0 || index >= tables.items.length) {
    throw new Error(
      `表格 index ${index} 超出范围（文档中共 ${tables.items.length} 个表格）` +
        (tables.items.length ? "，可先用 word_get_tables 查看列表" : "，可先用 word_insert_table 插入表格")
    );
  }
  return tables.items[index];
}

registry["word.getTables"] = async () => {
  return Word.run(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load("items/values,items/isUniform,items/nestingLevel,items/rowCount");
    await ctx.sync();
    if (!tables.items.length) {
      return { total: 0, tables: [], note: "文档中没有表格，可先用 word_insert_table 插入" };
    }
    return {
      total: tables.items.length,
      tables: tables.items.map((t, i) => {
        const firstRow = (t.values && t.values[0]) || [];
        return {
          index: i,
          rows: t.rowCount,
          cols: firstRow.length,
          isUniform: t.isUniform,
          nestingLevel: t.nestingLevel,
          preview: firstRow.join(" | ").slice(0, 120),
        };
      }),
      note: "index 按文档出现顺序（0 起始）；表格结构变化后请重新调用本工具",
    };
  });
};

registry["word.getTable"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const maxRows = Math.max(1, Number(args?.max_rows) || 100);
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,isUniform");
    const values = t.values || [];
    const out = {
      index,
      rowCount: values.length,
      colCount: (values[0] && values[0].length) || 0,
      isUniform: t.isUniform,
      values: values.slice(0, maxRows),
    };
    const notes = [];
    if (t.isUniform === false) notes.push("该表含合并单元格（非规整表），部分行可能长短不一，批量写入前请谨慎核对");
    if (values.length > maxRows) notes.push(`…[仅返回前 ${maxRows} 行（共 ${values.length} 行），可用 max_rows 调整]`);
    if (notes.length) out.note = notes.join(" ");
    return out;
  });
};

registry["word.setTableCell"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const row = requireInt(args?.row, "row");
  const col = requireInt(args?.col, "col");
  const text = String(args?.text ?? "");
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values");
    assertCellInRange(t.values, row, col);
    t.getCell(row, col).value = text;
    await ctx.sync();
    return { ok: true, index, row, col, text };
  });
};

registry["word.setTableCells"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const startRow = Math.max(0, Number(args?.start_row) || 0);
  const startCol = Math.max(0, Number(args?.start_col) || 0);
  const data = parse2D(args?.data, "data");
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values");
    const values = t.values || [];
    const endRow = startRow + data.length - 1;
    const endCol = startCol + Math.max(0, ...data.map((r) => r.length)) - 1;
    if (endRow >= values.length) {
      throw new Error(`写入区域越界：结束行 ${endRow} 超出表格行数（共 ${values.length} 行）；扩行请用 word_add_table_row`);
    }
    const cols = (values[0] && values[0].length) || 0;
    if (endCol >= cols) {
      throw new Error(`写入区域越界：结束列 ${endCol} 超出表格列数（共 ${cols} 列）；扩列请用 word_add_table_column`);
    }
    let updated = 0;
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        t.getCell(startRow + r, startCol + c).value = String(data[r][c] ?? "");
        updated++;
      }
    }
    await ctx.sync();
    return { ok: true, index, updatedCells: updated, range: `行 ${startRow}-${endRow} × 列 ${startCol}-${endCol}` };
  });
};

registry["word.addTableRow"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const at = ["top", "bottom", "above", "below"].includes(args?.at) ? args.at : "bottom";
  const data = args?.data == null ? null : parse1D(args.data, "data");
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    const colCount = (t.values && t.values[0] ? t.values[0].length : 0) || 0;
    // addRows/insertRows 的 values 为 string[][]（本次只插 1 行）
    const rowValues = data ? [padRow(data, colCount)] : undefined;
    let insertedAt;
    if (at === "top") {
      t.addRows("Start", 1, rowValues);
      insertedAt = 0;
    } else if (at === "bottom") {
      t.addRows("End", 1, rowValues);
      insertedAt = t.rowCount; // load 于插入前，即原行数
    } else {
      const ref = requireInt(args?.row_index, "at=above/below 时的 row_index");
      if (ref >= t.rowCount) throw new Error(`参照行 ${ref} 超出范围（表格共 ${t.rowCount} 行）`);
      t.getCell(ref, 0).insertRows(at === "above" ? "Before" : "After", 1, rowValues);
      insertedAt = at === "above" ? ref : ref + 1;
    }
    await ctx.sync();
    return { ok: true, index, at, insertedAt, note: `新行位于第 ${insertedAt} 行（0 起始）` };
  });
};

registry["word.addTableColumn"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const at = ["left", "right", "before", "after"].includes(args?.at) ? args.at : "right";
  const data = args?.data == null ? null : parse1D(args.data, "data");
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    const values = t.values || [];
    const rowCount = values.length;
    const colCount = (values[0] && values[0].length) || 0;
    let insertedAt;
    if (at === "left") {
      t.addColumns("Start", 1);
      insertedAt = 0;
    } else if (at === "right") {
      t.addColumns("End", 1);
      insertedAt = colCount;
    } else {
      const ref = requireInt(args?.col_index, "at=before/after 时的 col_index");
      if (ref >= colCount) throw new Error(`参照列 ${ref} 超出范围（表格共 ${colCount} 列）`);
      t.getCell(0, ref).insertColumns(at === "before" ? "Before" : "After", 1);
      insertedAt = at === "before" ? ref : ref + 1;
    }
    await ctx.sync();
    let filled = 0;
    if (data) {
      for (let r = 0; r < rowCount; r++) {
        if (data[r] == null) continue;
        t.getCell(r, insertedAt).value = String(data[r]);
        filled++;
      }
      await ctx.sync();
    }
    return { ok: true, index, at, insertedAt, filledCells: filled };
  });
};

registry["word.deleteTableRow"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const rowIndex = requireInt(args?.row_index, "row_index");
  const count = Math.max(1, Number(args?.count) || 1);
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    if (rowIndex >= t.rowCount) throw new Error(`行号 ${rowIndex} 超出范围（表格共 ${t.rowCount} 行）`);
    if (t.rowCount - count < 1) {
      throw new Error(`不能删除全部行（表格共 ${t.rowCount} 行）；要删除整个表格请用 word_delete_table`);
    }
    const deletedPreview = (t.values || [])
      .slice(rowIndex, rowIndex + count)
      .map((r) => (r || []).join(" | "))
      .join("；")
      .slice(0, 200);
    t.deleteRows(rowIndex, count);
    await ctx.sync();
    return { ok: true, index, deletedRows: count, remainingRows: t.rowCount - count, deletedPreview };
  });
};

registry["word.deleteTableColumn"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const colIndex = requireInt(args?.col_index, "col_index");
  const count = Math.max(1, Number(args?.count) || 1);
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    const cols = (t.values && t.values[0] ? t.values[0].length : 0) || 0;
    if (colIndex >= cols) throw new Error(`列号 ${colIndex} 超出范围（表格共 ${cols} 列）`);
    if (cols - count < 1) {
      throw new Error(`不能删除全部列（表格共 ${cols} 列）；要删除整个表格请用 word_delete_table`);
    }
    const deletedPreview = (t.values || [])
      .slice(0, 10)
      .map((r) => ((r || []).slice(colIndex, colIndex + count) || []).join(" | "))
      .join("；")
      .slice(0, 200);
    t.deleteColumns(colIndex, count);
    await ctx.sync();
    return { ok: true, index, deletedColumns: count, remainingColumns: cols - count, deletedPreview };
  });
};

registry["word.deleteTable"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    const firstRowPreview = ((t.values && t.values[0]) || []).join(" | ").slice(0, 120);
    const rows = t.rowCount;
    t.delete();
    await ctx.sync();
    return {
      ok: true,
      deletedIndex: index,
      deletedTable: { rows, firstRowPreview },
      note: "表格已删除；其余表格的 index 已变化，请重新调用 word_get_tables",
    };
  });
};

registry["word.setTableFormat"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const rowIndex = args?.row_index != null ? requireInt(args.row_index, "row_index") : null;
  const colIndex = args?.col_index != null ? requireInt(args.col_index, "col_index") : null;
  const style = args?.style ? String(args.style) : null;
  if (style && (rowIndex !== null || colIndex !== null)) {
    throw new Error("表格样式 style 仅支持整表设置，请不要与 row_index/col_index 同时使用");
  }
  const hAlign = args?.horizontal_alignment ? normalizeAlignment(args.horizontal_alignment) : null;
  if (args?.horizontal_alignment && !hAlign) {
    throw new Error(`无效的 horizontal_alignment: "${args.horizontal_alignment}"，合法值: Left / Centered / Right / Justified`);
  }
  let vAlign = null;
  if (args?.vertical_alignment) {
    const m = { top: "Top", center: "Center", middle: "Center", bottom: "Bottom" };
    vAlign = m[String(args.vertical_alignment).trim().toLowerCase()] || null;
    if (!vAlign) throw new Error(`无效的 vertical_alignment: "${args.vertical_alignment}"，合法值: Top / Center / Bottom`);
  }
  // shading_color 传 null 表示清除（Office.js 置空约定为字符串 "null"）
  const shading =
    args?.shading_color === undefined ? null : args?.shading_color === null ? "null" : normalizeColor(args.shading_color);
  const bold = args?.bold != null ? !!args.bold : null;
  const fontSize = args?.font_size != null ? Number(args.font_size) : null;
  const fontColor = args?.font_color ? normalizeColor(args.font_color) : null;
  const hasAny = style || hAlign || vAlign || shading !== null || bold !== null || fontSize != null || fontColor;
  if (!hasAny) throw new Error("未提供任何格式属性");

  const target =
    rowIndex !== null && colIndex !== null ? "cell" : rowIndex !== null ? "row" : colIndex !== null ? "column" : "table";

  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values,rowCount");
    const warnings = [];
    // 底纹单独同步：清除底纹（"null"）在部分版本不受支持，失败不应影响其他属性
    if (shading !== null) {
      try {
        if (target === "table") {
          t.shadingColor = shading;
        } else {
          const values = t.values || [];
          applyToCells(t, target, rowIndex, colIndex, values, (cell) => {
            cell.shadingColor = shading;
          });
        }
        await ctx.sync();
      } catch {
        warnings.push("底纹设置未生效（当前版本可能不支持清除/设置底纹）");
      }
    }
    if (target === "table") {
      if (style) t.style = style;
      if (hAlign) t.horizontalAlignment = hAlign;
      if (vAlign) t.verticalAlignment = vAlign;
      if (bold !== null) t.font.bold = bold;
      if (fontSize != null) t.font.size = fontSize;
      if (fontColor) t.font.color = fontColor;
    } else {
      applyToCells(t, target, rowIndex, colIndex, t.values || [], (cell) => {
        if (hAlign) cell.horizontalAlignment = hAlign;
        if (vAlign) cell.verticalAlignment = vAlign;
        if (bold !== null) cell.body.font.bold = bold;
        if (fontSize != null) cell.body.font.size = fontSize;
        if (fontColor) cell.body.font.color = fontColor;
      });
    }
    await ctx.sync();
    const targetLabel =
      target === "table"
        ? "整表"
        : target === "row"
          ? `第 ${rowIndex} 行`
          : target === "column"
            ? `第 ${colIndex} 列`
            : `单元格 (${rowIndex}, ${colIndex})`;
    return { ok: true, index, appliedTo: targetLabel, ...(warnings.length ? { warnings } : {}) };
  });
};

/** 按目标范围（row/column/cell）解析表格单元格代理并逐个应用 */
function applyToCells(t, target, rowIndex, colIndex, values, apply) {
  const rowCount = values.length;
  const cols = (values[0] && values[0].length) || 0;
  let cells;
  if (target === "cell") {
    assertCellInRange(values, rowIndex, colIndex);
    cells = [t.getCell(rowIndex, colIndex)];
  } else if (target === "row") {
    if (rowIndex >= rowCount) throw new Error(`行号 ${rowIndex} 超出范围（表格共 ${rowCount} 行）`);
    cells = Array.from({ length: cols }, (_, c) => t.getCell(rowIndex, c));
  } else {
    if (colIndex >= cols) throw new Error(`列号 ${colIndex} 超出范围（表格共 ${cols} 列）`);
    cells = Array.from({ length: rowCount }, (_, r) => t.getCell(r, colIndex));
  }
  for (const cell of cells) apply(cell);
}

registry["word.mergeTableCells"] = async (args) => {
  const index = requireInt(args?.table_index, "table_index");
  const topRow = requireInt(args?.top_row, "top_row");
  const firstCol = requireInt(args?.first_col, "first_col");
  const bottomRow = requireInt(args?.bottom_row, "bottom_row");
  const lastCol = requireInt(args?.last_col, "last_col");
  if (bottomRow < topRow || lastCol < firstCol) {
    throw new Error("合并范围无效：需满足 bottom_row ≥ top_row 且 last_col ≥ first_col");
  }
  return Word.run(async (ctx) => {
    const t = await loadTableAt(ctx, index, "values");
    const values = t.values || [];
    if (topRow >= values.length) throw new Error(`行号 ${topRow} 超出范围（表格共 ${values.length} 行）`);
    const cols = (values[0] && values[0].length) || 0;
    if (lastCol >= cols) throw new Error(`列号 ${lastCol} 超出范围（表格共 ${cols} 列）`);
    // mergeCells 为 WordApiDesktop 1.4（Office 2024/新版 365），旧版本做特性检测
    if (typeof t.mergeCells !== "function") {
      throw new Error("当前 Word 版本不支持合并单元格（需要 WordApiDesktop 1.4：Office 2024/新版 365 桌面版支持）");
    }
    t.mergeCells(topRow, firstCol, bottomRow, lastCol);
    await ctx.sync();
    return {
      ok: true,
      index,
      mergedRange: { topRow, firstCol, bottomRow, lastCol },
      note: "合并后表格结构已变化，建议重新 word_get_table 确认",
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
