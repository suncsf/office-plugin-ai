/**
 * PowerPoint 执行器（Office.js PowerPoint API）
 * 说明：PowerPoint JS API 的部分能力按 API 集合分批提供：
 *  - 删除形状 Shape.delete()：PowerPointApi 1.3+
 *  - 形状移动/缩放/名称/填充/线条、添加几何形状/线条/文本框：PowerPointApi 1.4+
 *  - 文本框/图片/表格的添加：PowerPointApi 1.8（addPicture/addTable）
 *  - 组合/解组：PowerPointApi 1.8+
 * （Office 2024 / 新版 Microsoft 365 均满足），全部做了特性检测，不支持时返回明确的错误信息。
 */
import { stripDataUrl } from "./dataurl";
import { resolveImageSource, saveWorkspaceFile, imageExtFromBase64 } from "./bridge";

/* ---------------- 内部工具函数 ---------------- */

function hasFn(obj, name) {
  return !!obj && typeof obj[name] === "function";
}

async function loadSlides(ctx) {
  const slides = ctx.presentation.slides;
  slides.load("id");
  await ctx.sync();
  return slides;
}

function shapePos(s) {
  return {
    id: s.id,
    name: s.name,
    type: s.type,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  };
}

/* ---------------- 形状级操作公共辅助 ---------------- */

/** 读取一页的全部形状（核心属性 + 文本；层级 zOrder 按 API 支持情况补充） */
async function getSlideShapes(ctx, index) {
  const n = Number(index);
  if (!Number.isInteger(n) || n < 0) throw new Error("slide_index 必须是非负整数");
  const slides = await loadSlides(ctx);
  if (n >= slides.items.length) throw new Error(`幻灯片 index ${n} 超出范围（共 ${slides.items.length} 页）`);
  const shapes = slides.items[n].shapes;
  shapes.load("id,name,type,left,top,width,height,textFrame/textRange/text");
  await ctx.sync();
  let hasZOrder = false;
  try {
    shapes.load("zOrderPosition");
    await ctx.sync();
    hasZOrder = true;
  } catch { /* 旧版 API 无层级属性 */ }
  return { shapes, hasZOrder, index: n };
}

function shapeInfo(s, hasZOrder = false) {
  const info = { ...shapePos(s) };
  if (hasZOrder && typeof s.zOrderPosition === "number") info.zOrder = s.zOrderPosition;
  const text = s.textFrame?.textRange?.text;
  if (text) info.text = text;
  return info;
}

/**
 * 按 shape_id / shape_ids / match_text 选定目标形状
 * @returns {{ list: Array, matchedCount: number }} list=命中的形状；matchedCount=实际匹配总数
 */
function pickTargets(shapes, args = {}) {
  const items = shapes.items;
  if (args.shape_id != null || args.shape_ids != null) {
    const raw = args.shape_ids != null ? args.shape_ids : [args.shape_id];
    const ids = (Array.isArray(raw) ? raw : [raw]).map(Number).filter((v) => Number.isFinite(v));
    if (!ids.length) throw new Error("shape_id / shape_ids 无效");
    const missing = ids.filter((id) => !items.some((s) => Number(s.id) === id));
    if (missing.length) throw new Error(`未找到 id=${missing.join("、")} 的形状（可先用 ppt_list_shapes 查看本页形状）`);
    return { list: items.filter((s) => ids.includes(Number(s.id))), matchedCount: ids.length };
  }
  if (args.match_text) {
    const needle = String(args.match_text);
    const hits = items.filter((s) => (s.textFrame?.textRange?.text || "").includes(needle));
    if (!hits.length) throw new Error(`未找到包含“${needle}”的形状（可先用 ppt_list_shapes 查看本页形状）`);
    return { list: args.match_all ? hits : [hits[0]], matchedCount: hits.length };
  }
  throw new Error("必须提供 shape_id / shape_ids / match_text 之一");
}

/** 单目标操作公共校验：只接受 shape_id 或 match_text */
function requireSingleSelector(args) {
  if (args?.shape_id == null && !args?.match_text) throw new Error("必须提供 shape_id 或 match_text");
  if (args?.shape_ids != null) throw new Error("本操作一次只能作用于一个形状，请改用 shape_id；批量请用 ppt_delete_shape");
}

/** 枚举幻灯片母版下的全部版式（id/name，顺序即 ppt_add_slide 的 layout_index） */
async function listLayouts(ctx) {
  const masters = ctx.presentation.slideMasters;
  masters.load("items");
  await ctx.sync();
  const layouts = [];
  for (const m of masters.items) {
    m.layouts.load("id,name");
    await ctx.sync();
    for (const l of m.layouts.items) layouts.push({ id: l.id, name: l.name });
  }
  return layouts;
}

/* ---------------- 执行器 ---------------- */

const registry = {};

registry["ppt.getPresentationInfo"] = async () => {
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    const info = { slideCount: slides.items.length };
    try {
      ctx.presentation.load("slideWidth,slideHeight");
      await ctx.sync();
      info.slideWidth = ctx.presentation.slideWidth;
      info.slideHeight = ctx.presentation.slideHeight;
      info.sizeUnit = "磅";
    } catch { /* 旧版 API 无页面尺寸 */ }
    return info;
  });
};

registry["ppt.listSlides"] = async (args) => {
  const includeText = args?.include_text !== false;
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    const list = [];
    for (let i = 0; i < slides.items.length; i++) {
      const item = { index: i, id: slides.items[i].id };
      if (includeText) {
        try {
          const shapes = slides.items[i].shapes;
          shapes.load("id,textFrame/textRange/text");
          await ctx.sync();
          const texts = shapes.items.map((s) => s.textFrame?.textRange?.text || "").filter(Boolean);
          item.shapeCount = shapes.items.length;
          item.textPreview = texts.join(" | ").slice(0, 160);
        } catch {
          item.note = "读取文本需要较新的 PowerPoint API";
        }
      }
      list.push(item);
    }
    return { slideCount: list.length, slides: list };
  });
};

async function slideText(ctx, index) {
  const slides = await loadSlides(ctx);
  if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围（共 ${slides.items.length} 页）`);
  const slide = slides.items[index];
  const shapes = slide.shapes;
  shapes.load("id,name,type,left,top,width,height,textFrame/textRange/text");
  await ctx.sync();
  return {
    slideIndex: index,
    shapes: shapes.items.map((s) => ({
      ...shapePos(s),
      text: s.textFrame?.textRange?.text ?? undefined,
    })),
  };
}

registry["ppt.getSlideText"] = async (args) => {
  const index = Number(args?.slide_index);
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  return PowerPoint.run(async (ctx) => slideText(ctx, index));
};

registry["ppt.getAllText"] = async (args) => {
  const maxSlides = Math.min(100, Math.max(1, Number(args?.max_slides) || 50));
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    const count = Math.min(slides.items.length, maxSlides);
    const pages = [];
    for (let i = 0; i < count; i++) {
      try {
        const r = await slideText(ctx, i);
        pages.push({
          slideIndex: i,
          text: r.shapes.map((s) => s.text).filter(Boolean).join("\n"),
        });
      } catch {
        pages.push({ slideIndex: i, text: "", note: "该页文本读取失败" });
      }
    }
    return { slideCount: slides.items.length, pages, note: slides.items.length > count ? `仅读取前 ${count} 页` : undefined };
  });
};

registry["ppt.addSlide"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    const before = slides.items.length;

    let layoutId;
    let layouts = null;
    if (args?.layout_index != null) {
      try {
        layouts = await listLayouts(ctx);
        const chosen = layouts[Number(args.layout_index)];
        if (!chosen) throw new Error(`layout_index 超出范围（共 ${layouts.length} 个版式）`);
        layoutId = chosen.id;
      } catch (e) {
        if (e?.message?.includes("超出范围")) throw e;
        layouts = null; // 当前 API 不支持读取版式
      }
    }

    let newSlide;
    try {
      newSlide = layoutId ? slides.add({ layoutId }) : slides.add({});
    } catch (e) {
      throw new Error(`新建幻灯片失败（可能需要较新的 PowerPoint API）：${e?.message || e}`);
    }
    newSlide.load("id");
    await ctx.sync();

    let note = "";
    if (layouts) note = `可用版式: ${layouts.map((l, i) => `${i}=${l.name}`).join(", ")}`;

    if (args?.title_text) {
      try {
        newSlide.shapes.addTextBox(String(args.title_text), { left: 60, top: 30, width: 500, height: 60 });
        await ctx.sync();
      } catch {
        note += " 标题文本框添加失败（需要 PowerPointApi 1.8+）";
      }
    }
    return { ok: true, slideIndex: before, newSlideCount: before + 1, layouts: layouts ? layouts.map((l) => l.name) : undefined, note: note.trim() || undefined };
  });
};

registry["ppt.deleteSlide"] = async (args) => {
  const index = Number(args?.slide_index);
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[index];
    if (hasFn(slide, "delete")) {
      slide.delete();
    } else if (hasFn(slides, "getItemAt") && hasFn(slides.getItemAt(index), "delete")) {
      slides.getItemAt(index).delete();
    } else {
      throw new Error("当前 PowerPoint JS API 不支持删除幻灯片，请手动操作");
    }
    await ctx.sync();
    return { ok: true, deletedIndex: index, remaining: slides.items.length - 1 };
  });
};

registry["ppt.moveSlide"] = async (args) => {
  const from = Number(args?.from_index);
  const to = Number(args?.to_index);
  if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < 0) {
    throw new Error("from_index / to_index 必须是非负整数");
  }
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (from >= slides.items.length || to >= slides.items.length) {
      throw new Error(`索引超出范围（共 ${slides.items.length} 页）`);
    }
    const slide = slides.items[from];
    if (hasFn(slide, "moveTo")) {
      try {
        slide.moveTo(to);
        await ctx.sync();
        return { ok: true, from, to };
      } catch (e) {
        throw new Error(`移动幻灯片失败：${e?.message || e}`);
      }
    }
    throw new Error("当前 PowerPoint JS API 版本不支持移动/重排幻灯片，请在 PowerPoint 中手动拖拽，或告知用户手动调整");
  });
};

registry["ppt.insertTextbox"] = async (args) => {
  const text = String(args?.text || "");
  if (!text) throw new Error("text 不能为空");
  const index = args?.slide_index != null ? Number(args.slide_index) : null;
  if (index === null) throw new Error("请提供 slide_index（可先用 ppt_list_slides 查看）");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[index];
    if (!hasFn(slide.shapes, "addTextBox")) {
      throw new Error("当前 PowerPoint JS API 不支持添加文本框（需要 PowerPointApi 1.8+，Office 2024/新版 365 支持）");
    }
    const shape = slide.shapes.addTextBox(text, {
      left: args?.left != null ? Number(args.left) : 60,
      top: args?.top != null ? Number(args.top) : 60,
      width: args?.width != null ? Number(args.width) : 300,
      height: args?.height != null ? Number(args.height) : 80,
    });
    if (args?.font_size != null || args?.bold != null || args?.color || args?.align) {
      const tr = shape.textFrame.textRange;
      if (args?.font_size != null) tr.font.size = Number(args.font_size);
      if (args?.bold != null) tr.font.bold = !!args.bold;
      if (args?.color) tr.font.color = String(args.color);
      if (args?.align) tr.paragraphFormat.horizontalAlignment = String(args.align);
    }
    shape.load("id");
    await ctx.sync();
    return { ok: true, slideIndex: index, shapeId: shape.id };
  });
};

registry["ppt.setShapeText"] = async (args) => {
  const index = Number(args?.slide_index);
  const text = String(args?.text ?? "");
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  if (args?.shape_id == null && !args?.match_text) throw new Error("必须提供 shape_id 或 match_text");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围`);
    const shapes = slides.items[index].shapes;
    shapes.load("id,textFrame/textRange/text");
    await ctx.sync();
    let target = null;
    if (args?.shape_id != null) {
      target = shapes.items.find((s) => s.id === Number(args.shape_id));
      if (!target) throw new Error(`未找到 id=${args.shape_id} 的形状`);
    } else {
      const needle = String(args.match_text);
      target = shapes.items.find((s) => (s.textFrame?.textRange?.text || "").includes(needle));
      if (!target) throw new Error(`未找到包含“${needle}”的形状`);
    }
    if (!target.textFrame?.textRange) {
      throw new Error("该形状不含可编辑文本（或当前 API 不支持），可删除后用 ppt_insert_textbox 重建");
    }
    try {
      target.textFrame.textRange.text = text;
    } catch {
      throw new Error("当前 PowerPoint API 不支持直接改写文本，可删除该形状后用 ppt_insert_textbox 重建");
    }
    await ctx.sync();
    return { ok: true, slideIndex: index, shapeId: target.id };
  });
};

registry["ppt.insertImage"] = async (args, ctxInfo) => {
  const b64 = await resolveImageSource(args, ctxInfo);
  if (!b64) throw new Error("base64 / file 至少提供一项");
  const index = args?.slide_index != null ? Number(args.slide_index) : null;
  if (index === null) throw new Error("请提供 slide_index");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围`);
    const shapes = slides.items[index].shapes;
    if (!hasFn(shapes, "addImage")) {
      throw new Error("当前 PowerPoint JS API 不支持插入图片（需要 PowerPointApi 1.8+）");
    }
    const shape = shapes.addImage(b64, {
      left: args?.left != null ? Number(args.left) : 60,
      top: args?.top != null ? Number(args.top) : 60,
      width: args?.width != null ? Number(args.width) : 300,
      height: args?.height != null ? Number(args.height) : 200,
    });
    shape.load("id");
    await ctx.sync();
    return { ok: true, slideIndex: index, shapeId: shape.id };
  });
};

registry["ppt.getImage"] = async (args, ctxInfo) => {
  requireSingleSelector(args);
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    if (String(target.type) !== "picture") {
      throw new Error(`目标形状类型是 ${target.type}，不是图片（picture）。用 ppt_list_shapes 查看 type=picture 的形状`);
    }
    if (!hasFn(target, "getImageAsBase64")) {
      throw new Error("当前 PowerPoint JS API 不支持导出图片内容（Shape.getImageAsBase64 需要 PowerPointApi 1.10+；Office 2024 永久版可能不支持，较新版 Microsoft 365 支持）");
    }
    const result = target.getImageAsBase64();
    await ctx.sync();
    const b64 = stripDataUrl(result?.value || "");
    if (!b64) throw new Error("未能读取图片数据");
    const ext = imageExtFromBase64(b64) || "png";
    const rel = `images/ppt_shape${target.id}_${Date.now().toString(36)}.${ext}`;
    const saved = await saveWorkspaceFile(ctxInfo?.conversationId, rel, b64);
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      file: saved.file,
      size: saved.size,
      note: "已保存到对话工作区。可用 image_process 裁剪/缩放后，用 ppt_insert_image 的 file 参数插回页面；调尺寸/位置用 ppt_move_shape。",
    };
  });
};

registry["ppt.insertTable"] = async (args) => {
  const rows = Math.max(1, Number(args?.rows) || 1);
  const cols = Math.max(1, Number(args?.cols) || 1);
  const index = args?.slide_index != null ? Number(args.slide_index) : null;
  if (index === null) throw new Error("请提供 slide_index");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围`);
    const shapes = slides.items[index].shapes;
    if (!hasFn(shapes, "addTable")) {
      throw new Error("当前 PowerPoint JS API 不支持插入表格（需要 PowerPointApi 1.8+）");
    }
    const data = Array.isArray(args?.data)
      ? new Array(rows).fill(null).map((_, r) =>
          new Array(cols).fill(null).map((__, c) => String(args.data?.[r]?.[c] ?? ""))
        )
      : undefined;
    const shape = shapes.addTable(rows, cols, {
      left: args?.left != null ? Number(args.left) : 60,
      top: args?.top != null ? Number(args.top) : 120,
      width: args?.width != null ? Number(args.width) : 500,
    }, data);
    shape.load("id");
    await ctx.sync();
    return { ok: true, slideIndex: index, shapeId: shape.id, rows, cols };
  });
};

registry["ppt.setSlideNotes"] = async (args) => {
  const index = Number(args?.slide_index);
  const notes = String(args?.notes ?? "");
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围`);
    const slide = slides.items[index];
    if (!slide.notesPage) {
      throw new Error("当前 PowerPoint JS API 不支持备注页（需要较新的 API 集合）");
    }
    const shapes = slide.notesPage.shapes;
    shapes.load("items");
    await ctx.sync();
    // 备注占位符通常是第 2 个形状（index 1）
    const candidates = [shapes.items[1], shapes.items[0]].filter(Boolean);
    let done = false;
    for (const s of candidates) {
      try {
        if (s.textFrame?.textRange) {
          s.textFrame.textRange.text = notes;
          done = true;
          break;
        }
      } catch { /* 尝试下一个 */ }
    }
    if (!done) throw new Error("未能写入备注（当前 API 不支持修改备注文本）");
    await ctx.sync();
    return { ok: true, slideIndex: index };
  });
};

/* ---------------- 形状级操作执行器 ---------------- */

registry["ppt.listShapes"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const { shapes, hasZOrder, index } = await getSlideShapes(ctx, args?.slide_index);
    return {
      slideIndex: index,
      shapeCount: shapes.items.length,
      shapes: shapes.items.map((s) => shapeInfo(s, hasZOrder)),
      note: "type=group 为组合形状，作为整体列出与操作；组合内的子形状需先 ppt_ungroup_shape 解散组合",
    };
  });
};

registry["ppt.deleteShape"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list, matchedCount } = pickTargets(shapes, args);
    if (!hasFn(list[0], "delete")) {
      throw new Error("当前 PowerPoint JS API 不支持删除形状（需要 PowerPointApi 1.3+）");
    }
    const deletedIds = list.map((t) => Number(t.id));
    for (const t of list) t.delete();
    await ctx.sync();
    const leftover = args?.match_text && !args?.match_all ? matchedCount - list.length : 0;
    return {
      ok: true,
      slideIndex: index,
      deletedShapeIds: deletedIds,
      deletedCount: deletedIds.length,
      remainingShapes: shapes.items.length - list.length,
      note: leftover > 0
        ? `本页还有 ${leftover} 个包含“${args.match_text}”的形状未删除，可再次调用并配合 match_all=true`
        : undefined,
    };
  });
};

registry["ppt.moveShape"] = async (args) => {
  requireSingleSelector(args);
  const updates = {};
  for (const k of ["left", "top", "width", "height"]) {
    if (args?.[k] != null) {
      const v = Number(args[k]);
      if (!Number.isFinite(v) || v < 0) throw new Error(`${k} 必须是非负数字（磅）`);
      updates[k] = v;
    }
  }
  if (!Object.keys(updates).length) throw new Error("至少提供 left / top / width / height 之一");
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    for (const [k, v] of Object.entries(updates)) target[k] = v;
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`移动/缩放形状失败（需要 PowerPointApi 1.4+，Office 2024/新版 365 支持）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      updated: updates,
      note: list.length > 1 ? `匹配到 ${list.length} 个形状，已作用于第一个（其余请用 shape_id 指定）` : undefined,
    };
  });
};

registry["ppt.formatShape"] = async (args) => {
  requireSingleSelector(args);
  if (args?.fill_color == null && !args?.fill_clear && args?.line_color == null && args?.line_weight == null && !args?.name) {
    throw new Error("至少提供 fill_color / fill_clear / line_color / line_weight / name 之一");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    const applied = {};
    if (args?.name) {
      target.name = String(args.name);
      applied.name = String(args.name);
    }
    if (args?.fill_clear || args?.fill_color != null) {
      if (!hasFn(target.fill, "setSolidColor") && !hasFn(target.fill, "clear")) {
        throw new Error("当前 PowerPoint JS API 不支持设置形状填充（需要 PowerPointApi 1.4+）");
      }
      if (args.fill_clear) {
        target.fill.clear();
        applied.fill = "已清除（透明）";
      } else {
        target.fill.setSolidColor(String(args.fill_color));
        applied.fill = String(args.fill_color);
      }
    }
    if (args?.line_color != null || args?.line_weight != null) {
      if (!target.lineFormat) throw new Error("该形状不支持线条/边框格式（如图片）");
      if (args.line_color != null) target.lineFormat.color = String(args.line_color);
      if (args.line_weight != null) target.lineFormat.weight = Number(args.line_weight);
      target.lineFormat.visible = true; // 设置边框颜色/粗细时自动显示边框
      applied.line = { color: args.line_color ?? undefined, weight: args.line_weight ?? undefined };
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置形状格式失败（需要 PowerPointApi 1.4+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      applied,
      note: list.length > 1 ? `匹配到 ${list.length} 个形状，已作用于第一个（其余请用 shape_id 指定）` : undefined,
    };
  });
};

registry["ppt.formatText"] = async (args) => {
  requireSingleSelector(args);
  const fmtKeys = ["font_name", "font_size", "bold", "italic", "underline", "color", "align"];
  if (!fmtKeys.some((k) => args?.[k] != null)) {
    throw new Error("至少提供一个格式属性（font_name/font_size/bold/italic/underline/color/align）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    const textRange = target.textFrame?.textRange;
    if (!textRange) throw new Error("该形状不含可编辑文本（图片、线条、表格等不适用本工具）");
    const f = textRange.font;
    if (args?.font_name != null) f.name = String(args.font_name);
    if (args?.font_size != null) f.size = Number(args.font_size);
    if (args?.bold != null) f.bold = !!args.bold;
    if (args?.italic != null) f.italic = !!args.italic;
    if (args?.underline != null) f.underline = String(args.underline);
    if (args?.color != null) f.color = String(args.color);
    if (args?.align != null) textRange.paragraphFormat.horizontalAlignment = String(args.align);
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置文字格式失败（需要 PowerPointApi 1.4+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      applied: Object.fromEntries(fmtKeys.filter((k) => args[k] != null).map((k) => [k, args[k]])),
      note: list.length > 1 ? `匹配到 ${list.length} 个形状，已作用于第一个（其余请用 shape_id 指定）` : undefined,
    };
  });
};

registry["ppt.addShape"] = async (args) => {
  const index = Number(args?.slide_index);
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  const kind = args?.kind === "line" ? "line" : "geometric";
  const opts = {
    left: args?.left != null ? Number(args.left) : 60,
    top: args?.top != null ? Number(args.top) : 60,
    width: args?.width != null ? Number(args.width) : 200,
    height: args?.height != null ? Number(args.height) : 100,
  };
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围（共 ${slides.items.length} 页）`);
    const shapes = slides.items[index].shapes;
    let shape;
    if (kind === "line") {
      if (!hasFn(shapes, "addLine")) {
        throw new Error("当前 PowerPoint JS API 不支持添加线条（需要 PowerPointApi 1.4+）");
      }
      shape = shapes.addLine(String(args?.connector || "straight"), opts);
    } else {
      if (!hasFn(shapes, "addGeometricShape")) {
        throw new Error("当前 PowerPoint JS API 不支持添加几何形状（需要 PowerPointApi 1.4+）");
      }
      shape = shapes.addGeometricShape(String(args?.geometry || "rectangle"), opts);
    }
    if (args?.name) shape.name = String(args.name);
    const notes = [];
    if (kind === "geometric") {
      if (args?.fill_color != null) shape.fill.setSolidColor(String(args.fill_color));
      const text = String(args?.text ?? "");
      if (text) {
        shape.textFrame.textRange.text = text;
        if (args?.font_size != null) shape.textFrame.textRange.font.size = Number(args.font_size);
        if (args?.font_color != null) shape.textFrame.textRange.font.color = String(args.font_color);
      }
    } else if (args?.fill_color != null) {
      notes.push("线条无填充，fill_color 已忽略（请用 line_color）");
    }
    if (args?.line_color != null) shape.lineFormat.color = String(args.line_color);
    if (args?.line_weight != null) shape.lineFormat.weight = Number(args.line_weight);
    shape.load("id");
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`插入形状失败（geometry/connector 需为 PowerPoint 枚举值，API 需要 PowerPointApi 1.4+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(shape.id),
      kind,
      geometry: kind === "geometric" ? String(args?.geometry || "rectangle") : undefined,
      connector: kind === "line" ? String(args?.connector || "straight") : undefined,
      note: notes.join("；") || undefined,
    };
  });
};

registry["ppt.setShapeZorder"] = async (args) => {
  const position = String(args?.position || "");
  if (!["BringToFront", "SendToBack", "BringForward", "SendBackward"].includes(position)) {
    throw new Error("position 必须是 BringToFront / SendToBack / BringForward / SendBackward 之一");
  }
  requireSingleSelector(args);
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    if (!hasFn(target, "setZOrder")) {
      throw new Error("当前 PowerPoint JS API 不支持调整形状层级（需要较新的 PowerPoint API）");
    }
    target.setZOrder(position);
    await ctx.sync();
    return { ok: true, slideIndex: index, shapeId: Number(target.id), position };
  });
};

registry["ppt.ungroupShape"] = async (args) => {
  requireSingleSelector(args);
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    try {
      target.load("group");
      await ctx.sync();
    } catch (e) {
      throw new Error(`读取组合信息失败（需要 PowerPointApi 1.8+）：${e?.message || e}`);
    }
    const group = target.group;
    if (!group || !hasFn(group, "ungroup")) {
      throw new Error("该形状不是组合形状，无法解散（type=group 的才是组合，可先用 ppt_list_shapes 查看）");
    }
    group.ungroup();
    await ctx.sync();
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      note: "已解散组合，子形状变为独立形状，可用 ppt_list_shapes 重新查看",
    };
  });
};

registry["ppt.getSlideLayouts"] = async () => {
  return PowerPoint.run(async (ctx) => {
    let layouts;
    try {
      layouts = await listLayouts(ctx);
    } catch (e) {
      throw new Error(`读取版式失败（需要 PowerPointApi 1.5+）：${e?.message || e}`);
    }
    return { layoutCount: layouts.length, layouts: layouts.map((l, i) => ({ index: i, id: l.id, name: l.name })) };
  });
};

/* ---------------- 文本框专用工具（读取含文字格式 / 调整大小 / 删除） ---------------- */

const TEXTBOX_FONT_LOAD =
  "textFrame/textRange/font/name,textFrame/textRange/font/size,textFrame/textRange/font/bold," +
  "textFrame/textRange/font/italic,textFrame/textRange/font/underline,textFrame/textRange/font/color," +
  "textFrame/textRange/paragraphFormat/horizontalAlignment";

function isTextbox(s) {
  return String(s.type).toLowerCase() === "textbox";
}

registry["ppt.getTextboxes"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    let withFormat = true;
    try {
      shapes.load(TEXTBOX_FONT_LOAD);
      await ctx.sync();
    } catch { withFormat = false; /* 旧版 API 无文字格式属性 */ }
    const boxes = [];
    for (const s of shapes.items) {
      const text = s.textFrame?.textRange?.text;
      if (!isTextbox(s) && !text) continue; // 只列文本框与带文字的形状
      const entry = { ...shapePos(s), text: text ?? undefined };
      if (withFormat && s.textFrame?.textRange?.font) {
        const f = s.textFrame.textRange.font;
        entry.font = {
          name: f.name ?? undefined,
          size: f.size ?? undefined,
          bold: f.bold ?? undefined,
          italic: f.italic ?? undefined,
          underline: f.underline ?? undefined,
          color: f.color ?? undefined,
          align: s.textFrame.textRange.paragraphFormat?.horizontalAlignment ?? undefined,
        };
      }
      boxes.push(entry);
    }
    return {
      slideIndex: index,
      textboxCount: boxes.filter((b) => String(b.type).toLowerCase() === "textbox").length,
      textboxes: boxes,
      note: withFormat
        ? "type=textBox 为文本框；font.color 为 #RRGGBB，混合格式时对应项可能为 null"
        : "文字格式详情需要 PowerPointApi 1.4+，已省略",
    };
  });
};

registry["ppt.adjustTextbox"] = async (args) => {
  requireSingleSelector(args);
  const updates = {};
  for (const k of ["left", "top", "width", "height"]) {
    if (args?.[k] != null) {
      const v = Number(args[k]);
      if (!Number.isFinite(v) || v < 0) throw new Error(`${k} 必须是非负数字（磅）`);
      updates[k] = v;
    }
  }
  if (!Object.keys(updates).length) {
    throw new Error("至少提供 left / top（位置）或 width / height（大小）之一");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    for (const [k, v] of Object.entries(updates)) target[k] = v;
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`调整文本框失败（需要 PowerPointApi 1.4+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      updated: updates,
      note: !isTextbox(target)
        ? `目标 type=${target.type}（非文本框），已按通用形状处理；通用形状建议用 ppt_move_shape`
        : undefined,
    };
  });
};

registry["ppt.deleteTextbox"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const boxes = shapes.items.filter(isTextbox);
    let targets = [];
    if (args?.shape_id != null || args?.shape_ids != null) {
      const raw = args.shape_ids != null ? args.shape_ids : [args.shape_id];
      const ids = (Array.isArray(raw) ? raw : [raw]).map(Number).filter((v) => Number.isFinite(v));
      if (!ids.length) throw new Error("shape_id / shape_ids 无效");
      const byId = new Map(shapes.items.map((s) => [Number(s.id), s]));
      const notFound = ids.filter((id) => !byId.has(id));
      if (notFound.length) throw new Error(`未找到 id=${notFound.join("、")} 的形状（可先用 ppt_get_textboxes 查看）`);
      const notBox = ids.filter((id) => !isTextbox(byId.get(id)));
      if (notBox.length) {
        throw new Error(`id=${notBox.join("、")} 不是文本框（type=${notBox.map((id) => byId.get(id).type).join("/")}），请改用 ppt_delete_shape`);
      }
      targets = ids.map((id) => byId.get(id));
    } else if (args?.match_text) {
      const needle = String(args.match_text);
      targets = boxes.filter((s) => (s.textFrame?.textRange?.text || "").includes(needle));
      if (!targets.length) {
        const other = shapes.items.find(
          (s) => !isTextbox(s) && (s.textFrame?.textRange?.text || "").includes(needle)
        );
        if (other) {
          throw new Error(`文本“${needle}”出现在 type=${other.type} 的形状中（非文本框），请改用 ppt_delete_shape`);
        }
        throw new Error(`未找到包含“${needle}”的文本框（可先用 ppt_get_textboxes 查看）`);
      }
      if (!args?.match_all) targets = [targets[0]];
    } else {
      throw new Error("必须提供 shape_id / shape_ids / match_text 之一");
    }
    if (!hasFn(targets[0], "delete")) {
      throw new Error("当前 PowerPoint JS API 不支持删除形状（需要 PowerPointApi 1.3+）");
    }
    const deletedIds = targets.map((t) => Number(t.id));
    for (const t of targets) t.delete();
    await ctx.sync();
    const leftover =
      args?.match_text && !args?.match_all
        ? boxes.filter((s) => (s.textFrame?.textRange?.text || "").includes(String(args.match_text))).length - targets.length
        : 0;
    return {
      ok: true,
      slideIndex: index,
      deletedShapeIds: deletedIds,
      deletedCount: deletedIds.length,
      remainingTextboxes: boxes.length - targets.length,
      note: leftover > 0
        ? `本页还有 ${leftover} 个包含“${args.match_text}”的文本框未删除，可再次调用并配合 match_all=true`
        : undefined,
    };
  });
};

registry["ppt.getSlideNotes"] = async (args) => {
  return PowerPoint.run(async (ctx) => {
    const { index } = await getSlideShapes(ctx, args?.slide_index);
    const slide = (await loadSlides(ctx)).items[index];
    if (!slide.notesPage) {
      throw new Error("当前 PowerPoint JS API 不支持备注页（需要较新的 API 集合）");
    }
    const shapes = slide.notesPage.shapes;
    shapes.load("textFrame/textRange/text");
    await ctx.sync();
    const texts = shapes.items.map((s) => s.textFrame?.textRange?.text || "").filter(Boolean);
    return { ok: true, slideIndex: index, notes: texts.join("\n"), note: texts.length ? undefined : "该页暂无备注" };
  });
};

registry["ppt.runCode"] = async (args) => {
  const code = String(args?.code ?? "").trim();
  if (!code) throw new Error("缺少必填参数 code（在 PowerPoint.run 回调内执行的 JavaScript 源码）");
  if (code.length > 200000) throw new Error("code 过长（>200K 字符）");
  const timeoutMs = Math.min(Math.max(Number(args?.timeout_ms) || 30000, 1000), 120000);

  let serialized;
  await PowerPoint.run(async (ctx) => {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    const fn = new AsyncFunction("PowerPoint", "context", `"use strict";\n${code}`);
    let timer;
    const work = (async () => {
      const result = await fn(PowerPoint, ctx);
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
