/**
 * PowerPoint 执行器（Office.js PowerPoint API）
 * 说明：PowerPoint JS API 的部分能力按 API 集合分批提供：
 *  - 删除形状 Shape.delete()：PowerPointApi 1.3+
 *  - 形状移动/缩放/名称/填充/线条（含虚线/双线/透明度）、添加几何形状/线条/文本框：PowerPointApi 1.4+
 *  - 选中（Slide.setSelected 视图跟随）：PowerPointApi 1.5+
 *  - 文本框/图片/表格的添加：PowerPointApi 1.8（addPicture/addTable）
 *  - 组合/解组：PowerPointApi 1.8+
 *  - 表格美化（单元格填充/字体/对齐/边框/合并/行列尺寸）：PowerPointApi 1.9+
 *  - 形状旋转/显隐/无障碍文本/超链接/调整值、导出形状图片：PowerPointApi 1.10+
 *  - 页面背景（纯色/图片/渐变，SlideBackground）：PowerPointApi 1.10+
 *  - 编号/项目符号类型样式、段落缩进级别（BulletFormat.type/style、indentLevel）：PowerPointApi 1.10+
 * （正式版 Office 以 isSetSupported 运行时检测为准，见 detectPptCapabilities；
 *   全部做了特性检测，不支持时返回明确的错误信息。）
 * 注意：插入图片（addImage/addPicture）仅 Preview API 提供，正式版会明确报错；
 *       隐藏幻灯片（Slide.hidden）、形状渐变填充、形状阴影、线条箭头、段落行距
 *       在 Office.js 中均无对应 API（见 PPT_UNSUPPORTED）。
 */
import { stripDataUrl } from "./dataurl";
import { resolveImageSource, saveWorkspaceFile, imageExtFromBase64 } from "./bridge";

/* ---------------- 内部工具函数 ---------------- */

function hasFn(obj, name) {
  return !!obj && typeof obj[name] === "function";
}

/** 运行时检测 PowerPointApi 需求集是否可用（自动识别当前 Office 可用接口的依据） */
function supportsApi(version) {
  try {
    return !!(Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("PowerPointApi", version));
  } catch {
    return false;
  }
}

/** 版本号比较：1.8 -> 1008，1.10 -> 1010（避免 parseFloat 把 1.10 变成 1.1） */
function verNum(v) {
  const [a, b] = String(v || "").split(".").map((n) => parseInt(n, 10) || 0);
  return a * 1000 + b;
}

/** 读取页面尺寸（磅）：优先 pageSetup（1.10），回退旧版 slideWidth/slideHeight 属性 */
async function slideSize(ctx) {
  try {
    ctx.presentation.load("pageSetup/slideWidth,pageSetup/slideHeight");
    await ctx.sync();
    const w = ctx.presentation.pageSetup?.slideWidth;
    const h = ctx.presentation.pageSetup?.slideHeight;
    if (Number.isFinite(w) && Number.isFinite(h)) return { w, h };
  } catch { /* 旧版无 pageSetup，走回退 */ }
  ctx.presentation.load("slideWidth,slideHeight");
  await ctx.sync();
  if (!Number.isFinite(ctx.presentation.slideWidth) || !Number.isFinite(ctx.presentation.slideHeight)) {
    throw new Error("无法读取页面尺寸（pageSetup 需要 PowerPointApi 1.10+，本机 Office 不支持）");
  }
  return { w: ctx.presentation.slideWidth, h: ctx.presentation.slideHeight };
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
    // 文本框级属性（TextFrame，1.4+）
    if (args?.vertical_alignment || args?.word_wrap != null || args?.margin_left != null || args?.margin_right != null || args?.margin_top != null || args?.margin_bottom != null) {
      const tf = shape.textFrame;
      if (args?.vertical_alignment) tf.verticalAlignment = String(args.vertical_alignment);
      if (args?.word_wrap != null) tf.wordWrap = !!args.word_wrap;
      if (args?.margin_left != null) tf.leftMargin = Number(args.margin_left);
      if (args?.margin_right != null) tf.rightMargin = Number(args.margin_right);
      if (args?.margin_top != null) tf.topMargin = Number(args.margin_top);
      if (args?.margin_bottom != null) tf.bottomMargin = Number(args.margin_bottom);
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
    // Shape.type 运行时返回 "Image"（枚举 PowerPoint.ShapeType.image）；旧写法误判为 "picture"，兼容两者
    const t = String(target.type || "").toLowerCase();
    if (t !== "image" && t !== "picture") {
      throw new Error(`目标形状类型是 ${target.type}，不是图片（image）。用 ppt_list_shapes 查看 type=image 的形状`);
    }
    if (!hasFn(target, "getImageAsBase64")) {
      throw new Error("当前 PowerPoint JS API 不支持导出图片内容（Shape.getImageAsBase64 需要 PowerPointApi 1.10+；Office 2024 永久版可能不支持，较新版 Microsoft 365 支持）");
    }
    const opts = {};
    if (args?.format != null) {
      const fmt = String(args.format).toLowerCase();
      if (fmt !== "png" && fmt !== "jpeg") throw new Error("format 仅支持 png / jpeg");
      opts.format = fmt.toUpperCase();
    }
    if (args?.width != null) {
      const w = Number(args.width);
      if (!Number.isFinite(w) || w <= 0) throw new Error("width 必须是正数（像素）");
      opts.width = Math.round(w);
    }
    if (args?.height != null) {
      const h = Number(args.height);
      if (!Number.isFinite(h) || h <= 0) throw new Error("height 必须是正数（像素）");
      opts.height = Math.round(h);
    }
    const result = target.getImageAsBase64(Object.keys(opts).length ? opts : undefined);
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
  if (args?.rotation != null) {
    const r = Number(args.rotation);
    if (!Number.isFinite(r)) throw new Error("rotation 必须是数字（度，顺时针为正）");
    if (!supportsApi("1.10")) {
      throw new Error("设置旋转（rotation）需要 PowerPointApi 1.10+（本机 Office 不支持，可升级 Microsoft 365；可先用 ppt_get_api_info 查看本机能力）");
    }
    updates.rotation = r;
  }
  const dx = args?.dx != null ? Number(args.dx) : null;
  const dy = args?.dy != null ? Number(args.dy) : null;
  if ((dx != null && !Number.isFinite(dx)) || (dy != null && !Number.isFinite(dy))) {
    throw new Error("dx / dy 必须是数字（磅，可为负，表示相对当前位置的位移）");
  }
  const hAlign = args?.h_align ? String(args.h_align).toLowerCase() : null;
  const vAlign = args?.v_align ? String(args.v_align).toLowerCase() : null;
  if (hAlign && !["left", "center", "right"].includes(hAlign)) throw new Error("h_align 仅支持 left / center / right");
  if (vAlign && !["top", "middle", "bottom"].includes(vAlign)) throw new Error("v_align 仅支持 top / middle / bottom");
  if (!Object.keys(updates).length && dx == null && dy == null && !hAlign && !vAlign) {
    throw new Error("至少提供 left / top / width / height / rotation / dx / dy / h_align / v_align 之一");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    // dx / dy：相对位移（getSlideShapes 已加载 left/top 当前值）
    if (dx != null && Number.isFinite(target.left)) updates.left = Math.max(0, target.left + dx);
    if (dy != null && Number.isFinite(target.top)) updates.top = Math.max(0, target.top + dy);
    // h_align / v_align：页面级对齐（需要页面尺寸）
    if (hAlign || vAlign) {
      const size = await slideSize(ctx);
      if (hAlign && Number.isFinite(target.width)) {
        updates.left = hAlign === "left" ? 0 : hAlign === "center" ? Math.max(0, (size.w - target.width) / 2) : Math.max(0, size.w - target.width);
      }
      if (vAlign && Number.isFinite(target.height)) {
        updates.top = vAlign === "top" ? 0 : vAlign === "middle" ? Math.max(0, (size.h - target.height) / 2) : Math.max(0, size.h - target.height);
      }
    }
    // 先应用位置/尺寸（1.4），同步失败互不影响旋转
    const posApplied = {};
    for (const k of ["left", "top", "width", "height"]) {
      if (updates[k] != null) {
        target[k] = updates[k];
        posApplied[k] = updates[k];
      }
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`移动/缩放形状失败（需要 PowerPointApi 1.4+，Office 2024/新版 365 支持）：${e?.message || e}`);
    }
    let rotationApplied;
    if (updates.rotation != null) {
      try {
        target.rotation = updates.rotation;
        await ctx.sync();
        rotationApplied = updates.rotation;
      } catch (e) {
        throw new Error(`旋转设置失败（需要 PowerPointApi 1.10+）：${e?.message || e}`);
      }
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      updated: rotationApplied != null ? { ...posApplied, rotation: rotationApplied } : posApplied,
      note: list.length > 1 ? `匹配到 ${list.length} 个形状，已作用于第一个（其余请用 shape_id 指定）` : undefined,
    };
  });
};

registry["ppt.formatShape"] = async (args) => {
  requireSingleSelector(args);
  const fmtKeys = ["fill_color", "fill_clear", "fill_transparency", "line_color", "line_weight", "line_dash", "line_style", "line_transparency", "name"];
  if (!fmtKeys.some((k) => args?.[k] != null)) {
    throw new Error("至少提供 fill_color / fill_clear / fill_transparency / line_color / line_weight / line_dash / line_style / line_transparency / name 之一");
  }
  for (const k of ["fill_transparency", "line_transparency"]) {
    if (args?.[k] != null) {
      const v = Number(args[k]);
      if (!Number.isFinite(v) || v < 0 || v > 1) throw new Error(`${k} 必须是 0（不透明）~ 1（全透明）之间的数字`);
    }
  }
  if (args?.line_style != null && !["Single", "ThickThin", "ThinThick", "ThickBetweenThin", "ThinThin"].includes(String(args.line_style))) {
    throw new Error("line_style 仅支持 Single（单线）/ ThickThin（粗细双线）/ ThinThick（细粗双线）/ ThickBetweenThin（粗间细三线）/ ThinThin（细细双线）");
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
    if (args?.fill_clear || args?.fill_color != null || args?.fill_transparency != null) {
      if (!hasFn(target.fill, "setSolidColor") && !hasFn(target.fill, "clear")) {
        throw new Error("当前 PowerPoint JS API 不支持设置形状填充（需要 PowerPointApi 1.4+）");
      }
      if (args.fill_clear) {
        target.fill.clear();
        applied.fill = "已清除（透明）";
      } else {
        if (args.fill_color != null) {
          target.fill.setSolidColor(String(args.fill_color));
          applied.fill = String(args.fill_color);
        }
        if (args.fill_transparency != null) {
          target.fill.transparency = Number(args.fill_transparency);
          applied.fill_transparency = Number(args.fill_transparency);
        }
      }
    }
    if (args?.line_color != null || args?.line_weight != null || args?.line_dash != null || args?.line_style != null || args?.line_transparency != null) {
      if (!target.lineFormat) throw new Error("该形状不支持线条/边框格式（如图片）");
      if (args.line_color != null) target.lineFormat.color = String(args.line_color);
      if (args.line_weight != null) target.lineFormat.weight = Number(args.line_weight);
      if (args.line_dash != null) target.lineFormat.dashStyle = String(args.line_dash);
      if (args.line_style != null) target.lineFormat.style = String(args.line_style);
      if (args.line_transparency != null) target.lineFormat.transparency = Number(args.line_transparency);
      target.lineFormat.visible = true; // 设置边框属性时自动显示边框
      applied.line = { color: args.line_color ?? undefined, weight: args.line_weight ?? undefined, dash: args.line_dash ?? undefined, style: args.line_style ?? undefined, transparency: args.line_transparency ?? undefined };
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
  const fmtKeys = [
    "font_name", "font_size", "bold", "italic", "underline", "color", "align",
    "subscript", "superscript", "strikethrough", "small_caps", "all_caps", "indent_level",
    "bullet_visible", "bullet_type", "bullet_style",
    "vertical_alignment", "word_wrap", "auto_size",
    "margin_left", "margin_right", "margin_top", "margin_bottom",
  ];
  if (!fmtKeys.some((k) => args?.[k] != null)) {
    throw new Error("至少提供一个格式属性（font_name/font_size/bold/italic/underline/color/align/subscript/superscript/strikethrough/small_caps/all_caps/indent_level/bullet_*/vertical_alignment/word_wrap/auto_size/margin_*）");
  }
  if (args?.subscript != null && args?.superscript != null && !!args.subscript && !!args.superscript) {
    throw new Error("不能同时设置上标（superscript）和下标（subscript）");
  }
  if (args?.auto_size != null && !["None", "TextToFitShape", "ShapeToFitText"].includes(String(args.auto_size))) {
    throw new Error("auto_size 仅支持 None（不自动调整）/ TextToFitShape（文字缩放适应形状）/ ShapeToFitText（形状适应文字）");
  }
  if (args?.bullet_type != null && !["None", "Numbered", "Unnumbered"].includes(String(args.bullet_type))) {
    throw new Error("bullet_type 仅支持 None（无符号）/ Numbered（数字编号）/ Unnumbered（项目符号）");
  }
  // 字符级扩展格式需要 PowerPointApi 1.8（ShapeFont），缩进级别/编号类型与样式需要 1.10（ParagraphFormat.indentLevel、BulletFormat.type/style）
  const charExtKeys = ["subscript", "superscript", "strikethrough", "small_caps", "all_caps"];
  if (charExtKeys.some((k) => args?.[k] != null) && !supportsApi("1.8")) {
    throw new Error("下标/上标/删除线/小型大写/全部大写需要 PowerPointApi 1.8+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  const paraExtKeys = ["indent_level", "bullet_type", "bullet_style"];
  if (paraExtKeys.some((k) => args?.[k] != null) && !supportsApi("1.10")) {
    throw new Error("缩进级别/编号类型/编号样式（indent_level/bullet_type/bullet_style）需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
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
    if (args?.subscript != null) f.subscript = !!args.subscript;
    if (args?.superscript != null) f.superscript = !!args.superscript;
    if (args?.strikethrough != null) f.strikethrough = !!args.strikethrough;
    if (args?.small_caps != null) f.smallCaps = !!args.small_caps;
    if (args?.all_caps != null) f.allCaps = !!args.all_caps;
    if (args?.indent_level != null) textRange.paragraphFormat.indentLevel = Number(args.indent_level);
    // 段落级项目符号/编号（BulletFormat：visible 1.4+，type/style 1.10+，门禁见上方）
    const pf = textRange.paragraphFormat;
    if (args?.bullet_visible != null) pf.bulletFormat.visible = !!args.bullet_visible;
    if (args?.bullet_type != null) pf.bulletFormat.type = String(args.bullet_type);
    if (args?.bullet_style != null) pf.bulletFormat.style = String(args.bullet_style);
    // 文本框级属性（TextFrame，1.4+）
    const tf = target.textFrame;
    if (args?.vertical_alignment != null) tf.verticalAlignment = String(args.vertical_alignment);
    if (args?.word_wrap != null) tf.wordWrap = !!args.word_wrap;
    if (args?.auto_size != null) {
      const map = { None: "AutoSizeNone", TextToFitShape: "AutoSizeTextToFitShape", ShapeToFitText: "AutoSizeShapeToFitText" };
      tf.autoSizeSetting = map[String(args.auto_size)];
    }
    if (args?.margin_left != null) tf.leftMargin = Number(args.margin_left);
    if (args?.margin_right != null) tf.rightMargin = Number(args.margin_right);
    if (args?.margin_top != null) tf.topMargin = Number(args.margin_top);
    if (args?.margin_bottom != null) tf.bottomMargin = Number(args.margin_bottom);
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

/* ---------------- 页面（幻灯片）属性扩展 ---------------- */

registry["ppt.getSlideInfo"] = async (args) => {
  const n = Number(args?.slide_index);
  if (!Number.isInteger(n) || n < 0) throw new Error("slide_index 必须是非负整数");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (n >= slides.items.length) throw new Error(`幻灯片 index ${n} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[n];
    const info = { slideIndex: n, id: slide.id };
    // 页码（PowerPointApi 1.8+ 才有 Slide.index）
    if (Number.isFinite(slide.index)) info.index = slide.index;
    // 版式名称（1.3+）
    try {
      slide.load("layout/name");
      await ctx.sync();
      info.layout = slide.layout?.name ?? undefined;
    } catch {
      info.layoutNote = "读取版式需要 PowerPointApi 1.3+，可用 ppt_get_slide_layouts 查看全部版式";
    }
    // 背景填充类型（1.10+）
    try {
      slide.background.fill.load("type");
      await ctx.sync();
      info.backgroundFill = slide.background.fill.type ?? undefined;
    } catch { /* 1.10 之前无背景 API */ }
    // 形状统计
    const shapes = slide.shapes;
    shapes.load("id,type");
    await ctx.sync();
    info.shapeCount = shapes.items.length;
    const byType = {};
    for (const s of shapes.items) byType[s.type] = (byType[s.type] || 0) + 1;
    info.shapeTypes = byType;
    // 标签（1.3+）
    try {
      slide.tags.load("key,value");
      await ctx.sync();
      if (slide.tags.items.length) info.tags = Object.fromEntries(slide.tags.items.map((t) => [t.key, t.value]));
    } catch { /* 旧版无 tags */ }
    // 备注
    try {
      if (slide.notesPage) {
        const ns = slide.notesPage.shapes;
        ns.load("textFrame/textRange/text");
        await ctx.sync();
        const notes = ns.items.map((s) => s.textFrame?.textRange?.text || "").filter(Boolean).join("\n");
        if (notes) info.notes = notes;
      }
    } catch { /* 备注读取失败不影响整体 */ }
    return info;
  });
};

/** 枚举全部版式的 SlideLayout 对象（顺序即 layout_index，供 applyLayout 使用） */
async function layoutObjects(ctx) {
  const masters = ctx.presentation.slideMasters;
  masters.load("items");
  await ctx.sync();
  const objs = [];
  for (const m of masters.items) {
    m.layouts.load("id,name");
    await ctx.sync();
    for (const l of m.layouts.items) objs.push(l);
  }
  return objs;
}

registry["ppt.applySlideLayout"] = async (args) => {
  const n = Number(args?.slide_index);
  if (!Number.isInteger(n) || n < 0) throw new Error("slide_index 必须是非负整数");
  const li = Number(args?.layout_index);
  if (!Number.isInteger(li) || li < 0) throw new Error("layout_index 必须是非负整数（可先用 ppt_get_slide_layouts 查看）");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (n >= slides.items.length) throw new Error(`幻灯片 index ${n} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[n];
    if (!hasFn(slide, "applyLayout") || !supportsApi("1.8")) {
      throw new Error("应用版式（Slide.applyLayout）需要 PowerPointApi 1.8+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
    }
    const objs = await layoutObjects(ctx);
    const chosen = objs[li];
    if (!chosen) throw new Error(`layout_index 超出范围（共 ${objs.length} 个版式）：${objs.map((l, i) => `${i}=${l.name}`).join(", ")}`);
    const name = chosen.name;
    try {
      slide.applyLayout(chosen);
      await ctx.sync();
    } catch (e) {
      throw new Error(`应用版式失败：${e?.message || e}`);
    }
    return { ok: true, slideIndex: n, layoutIndex: li, layoutName: name, note: "应用版式会按新版式重排占位符，自定义元素位置保留" };
  });
};

registry["ppt.getSlideImage"] = async (args, ctxInfo) => {
  const n = Number(args?.slide_index);
  if (!Number.isInteger(n) || n < 0) throw new Error("slide_index 必须是非负整数");
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (n >= slides.items.length) throw new Error(`幻灯片 index ${n} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[n];
    if (!hasFn(slide, "getImageAsBase64") || !supportsApi("1.8")) {
      throw new Error("渲染页面为图片（Slide.getImageAsBase64）需要 PowerPointApi 1.8+（本机 Office 不支持）");
    }
    const opts = {};
    if (args?.width != null) {
      const w = Number(args.width);
      if (!Number.isFinite(w) || w <= 0) throw new Error("width 必须是正数（像素）");
      opts.width = Math.round(w);
    }
    if (args?.height != null) {
      const h = Number(args.height);
      if (!Number.isFinite(h) || h <= 0) throw new Error("height 必须是正数（像素）");
      opts.height = Math.round(h);
    }
    const result = slide.getImageAsBase64(Object.keys(opts).length ? opts : undefined);
    await ctx.sync();
    const b64 = stripDataUrl(result?.value || "");
    if (!b64) throw new Error("未能渲染页面图片");
    const rel = `images/ppt_slide${n}_${Date.now().toString(36)}.png`;
    const saved = await saveWorkspaceFile(ctxInfo?.conversationId, rel, b64);
    return {
      ok: true,
      slideIndex: n,
      file: saved.file,
      size: saved.size,
      note: "页面已渲染为 PNG 并保存到对话工作区（内容缩略图/预览用途）。",
    };
  });
};

registry["ppt.setSlideBackground"] = async (args, ctxInfo) => {
  const n = Number(args?.slide_index);
  if (!Number.isInteger(n) || n < 0) throw new Error("slide_index 必须是非负整数");
  const wantImage = args?.base64 != null || args?.file != null;
  const wantSolid = args?.color != null;
  const wantGradient = args?.gradient_type != null;
  const wantReset = args?.reset === true;
  if (!wantImage && !wantSolid && !wantGradient && !wantReset) throw new Error("至少提供 color（纯色背景）/ gradient_type（渐变背景）/ base64 或 file（图片背景）/ reset=true（重置为母版背景）之一");
  if (wantGradient && !["Linear", "Radial", "Rectangular", "Path", "ShadeFromTitle"].includes(String(args.gradient_type))) {
    throw new Error("gradient_type 仅支持 Linear（线性）/ Radial（径向）/ Rectangular（矩形）/ Path（路径）/ ShadeFromTitle（标题底纹）");
  }
  let transparency;
  if (args?.transparency != null) {
    transparency = Number(args.transparency);
    if (!Number.isFinite(transparency) || transparency < 0 || transparency > 1) throw new Error("transparency 必须是 0（不透明）~ 1（全透明）之间的数字");
  }
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (n >= slides.items.length) throw new Error(`幻灯片 index ${n} 超出范围（共 ${slides.items.length} 页）`);
    const slide = slides.items[n];
    const fill = slide.background?.fill;
    if (!fill || (!hasFn(fill, "setSolidFill") && !hasFn(slide.background, "reset"))) {
      throw new Error("设置页面背景需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
    }
    const applied = { slideIndex: n };
    if (wantReset) {
      slide.background.reset();
      applied.reset = true;
    } else if (wantImage) {
      const b64 = await resolveImageSource(args, ctxInfo);
      if (!b64) throw new Error("base64 / file 至少提供一项");
      fill.setPictureOrTextureFill({ imageBase64: b64, ...(transparency != null ? { transparency } : {}) });
      applied.image = true;
    } else if (wantGradient) {
      if (!hasFn(fill, "setGradientFill")) {
        throw new Error("渐变背景（SlideBackgroundFill.setGradientFill）需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
      }
      fill.setGradientFill({ type: String(args.gradient_type) });
      applied.gradient = String(args.gradient_type);
    } else {
      fill.setSolidFill({ color: String(args.color), ...(transparency != null ? { transparency } : {}) });
      applied.color = String(args.color);
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置页面背景失败（需要 PowerPointApi 1.10+）：${e?.message || e}`);
    }
    if (transparency != null) applied.transparency = transparency;
    if (applied.gradient) {
      applied.note = "渐变色标与角度由当前主题决定（Office.js 仅支持设置渐变类型）；如需精确配色建议用纯色背景或 image_process 生成渐变图后以 base64 设为背景";
    }
    return { ok: true, ...applied };
  });
};

/* ---------------- 形状级扩展（读取格式 / 显隐 / 图片填充 / 超链接 / 无障碍文本 / 组合） ---------------- */

registry["ppt.getShapeFormat"] = async (args) => {
  requireSingleSelector(args);
  return PowerPoint.run(async (ctx) => {
    const { shapes, index, hasZOrder } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const s = list[0];
    const info = { slideIndex: index, shapeId: Number(s.id), name: s.name, type: s.type, left: s.left, top: s.top, width: s.width, height: s.height };
    // 填充 / 边框 / 层级（1.4+）
    let fmtOk = false;
    try {
      s.load("fill/type,fill/color,fill/transparency,lineFormat/visible,lineFormat/color,lineFormat/weight");
      await ctx.sync();
      fmtOk = true;
    } catch { /* 旧版 API 无格式属性 */ }
    if (fmtOk) {
      info.fill = { type: s.fill?.type ?? undefined, color: s.fill?.color ?? undefined, transparency: s.fill?.transparency ?? undefined };
      if (s.lineFormat) {
        info.line = { visible: s.lineFormat.visible ?? undefined, color: s.lineFormat.color ?? undefined, weight: s.lineFormat.weight ?? undefined };
      }
    }
    // 旋转 / 显隐 / 无障碍文本（1.10+）
    try {
      s.load("rotation,visible,altTextTitle,altTextDescription,isDecorative");
      await ctx.sync();
      info.rotation = s.rotation;
      info.visible = s.visible;
      info.altText = { title: s.altTextTitle ?? undefined, description: s.altTextDescription ?? undefined, decorative: s.isDecorative ?? undefined };
    } catch { info.extNote = "旋转/显隐/无障碍文本需要 PowerPointApi 1.10+，已省略"; }
    if (hasZOrder && typeof s.zOrderPosition === "number") info.zOrder = s.zOrderPosition;
    const text = s.textFrame?.textRange?.text;
    if (text) info.text = text;
    return info;
  });
};

registry["ppt.setShapeVisible"] = async (args) => {
  requireSingleSelector(args);
  if (typeof args?.visible !== "boolean") throw new Error("visible 必须是布尔值（true=显示，false=隐藏）");
  if (!supportsApi("1.10")) {
    throw new Error("形状显示/隐藏（Shape.visible）需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    target.visible = args.visible;
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置形状可见性失败（需要 PowerPointApi 1.10+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      visible: args.visible,
      note: list.length > 1 ? `匹配到 ${list.length} 个形状，已作用于第一个（其余请用 shape_id 指定）` : undefined,
    };
  });
};

registry["ppt.setShapeFillImage"] = async (args, ctxInfo) => {
  requireSingleSelector(args);
  const b64 = await resolveImageSource(args, ctxInfo);
  if (!b64) throw new Error("base64 / file 至少提供一项");
  if (!supportsApi("1.8")) {
    throw new Error("图片填充（ShapeFill.setImage）需要 PowerPointApi 1.8+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    if (!hasFn(target.fill, "setImage")) {
      throw new Error("该形状不支持设置填充图片（ShapeFill.setImage 需要 PowerPointApi 1.8+）");
    }
    target.fill.setImage(b64);
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置图片填充失败：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      note: "已把图片设为形状的填充内容。对图片形状（type=image）这相当于替换其显示内容；调整显示尺寸仍用 ppt_move_shape。",
    };
  });
};

registry["ppt.setShapeHyperlink"] = async (args) => {
  requireSingleSelector(args);
  const remove = args?.remove === true;
  const address = String(args?.address || "").trim();
  if (!remove && !address) throw new Error("必须提供 address（链接地址，URL/文件路径/mailto），或 remove=true 清除已有链接");
  if (!supportsApi("1.10")) {
    throw new Error("形状超链接需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    if (remove) {
      // 清除：遍历本页超链接，删除链接到该形状的项（Hyperlink.delete 为 1.10）
      const slide = target.getParentSlide();
      slide.load("id");
      await ctx.sync();
      const links = slide.hyperlinks;
      links.load("items");
      await ctx.sync();
      const pairs = [];
      for (const h of links.items) {
        try {
          const ls = h.getLinkedShapeOrNullObject();
          ls.load("id");
          pairs.push([h, ls]);
        } catch { /* 无法关联到形状的超链接，跳过 */ }
      }
      await ctx.sync();
      const toDelete = pairs.filter(([, ls]) => Number(ls.id) === Number(target.id)).map(([h]) => h);
      if (!toDelete.length) return { ok: true, slideIndex: index, shapeId: Number(target.id), removed: 0, note: "该形状没有超链接" };
      for (const h of toDelete) {
        if (!hasFn(h, "delete")) throw new Error("清除超链接需要 PowerPointApi 1.10+");
        h.delete();
      }
      await ctx.sync();
      return { ok: true, slideIndex: index, shapeId: Number(target.id), removed: toDelete.length };
    }
    if (!hasFn(target, "setHyperlink")) {
      throw new Error("当前 PowerPoint JS API 不支持设置形状超链接（Shape.setHyperlink 需要 PowerPointApi 1.10+）");
    }
    const screenTip = String(args?.screen_tip || "").trim();
    try {
      target.setHyperlink({ address, ...(screenTip ? { screenTip } : {}) });
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置超链接失败：${e?.message || e}`);
    }
    return { ok: true, slideIndex: index, shapeId: Number(target.id), address, screenTip: screenTip || undefined, note: "设置会覆盖该形状已有的超链接" };
  });
};

registry["ppt.setShapeAltText"] = async (args) => {
  requireSingleSelector(args);
  if (args?.title == null && args?.description == null && args?.decorative == null) {
    throw new Error("至少提供 title（无障碍标题）/ description（无障碍描述）/ decorative（是否装饰元素）之一");
  }
  if (!supportsApi("1.10")) {
    throw new Error("无障碍文本（altText）需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    const applied = {};
    if (args?.title != null) { target.altTextTitle = String(args.title); applied.title = String(args.title); }
    if (args?.description != null) { target.altTextDescription = String(args.description); applied.description = String(args.description); }
    if (args?.decorative != null) { target.isDecorative = !!args.decorative; applied.decorative = !!args.decorative; }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`设置无障碍文本失败（需要 PowerPointApi 1.10+）：${e?.message || e}`);
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

registry["ppt.groupShapes"] = async (args) => {
  const index = Number(args?.slide_index);
  if (!Number.isInteger(index) || index < 0) throw new Error("slide_index 必须是非负整数");
  const raw = Array.isArray(args?.shape_ids) ? args.shape_ids : [];
  const ids = raw.map(Number).filter((v) => Number.isFinite(v));
  if (ids.length < 2) throw new Error("shape_ids 必须是至少 2 个形状 id 的数组（先用 ppt_list_shapes 查看本页形状）");
  if (new Set(ids).size !== ids.length) throw new Error("shape_ids 中存在重复的形状 id");
  if (!supportsApi("1.8")) {
    throw new Error("组合形状（ShapeCollection.addGroup）需要 PowerPointApi 1.8+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (index >= slides.items.length) throw new Error(`幻灯片 index ${index} 超出范围（共 ${slides.items.length} 页）`);
    const shapes = slides.items[index].shapes;
    if (!hasFn(shapes, "addGroup")) {
      throw new Error("当前 PowerPoint JS API 不支持组合形状（ShapeCollection.addGroup 需要 PowerPointApi 1.8+）");
    }
    shapes.load("id");
    await ctx.sync();
    const have = new Set(shapes.items.map((s) => Number(s.id)));
    const missing = ids.filter((id) => !have.has(id));
    if (missing.length) throw new Error(`未找到 id=${missing.join("、")} 的形状（可先用 ppt_list_shapes 查看本页形状）`);
    const group = shapes.addGroup(ids.map(String));
    group.load("id");
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`组合形状失败：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      groupShapeId: Number(group.id),
      memberCount: ids.length,
      note: "已创建组合（type=group），后续可对组合整体移动/缩放；解散用 ppt_ungroup_shape",
    };
  });
};

/* ---------------- 美化增强（形状调整值 / 多形状对齐分布 / 表格格式化） ---------------- */

registry["ppt.setShapeAdjustments"] = async (args) => {
  requireSingleSelector(args);
  const read = args?.read === true;
  const adj = Array.isArray(args?.adjustments) ? args.adjustments : [];
  if (!read && !adj.length) {
    throw new Error("adjustments 必须是 [{index, value}] 数组（如 [{\"index\":0,\"value\":0.15}]）；只读取当前调整值时传 read=true");
  }
  for (const a of adj) {
    if (a == null || !Number.isInteger(Number(a.index)) || Number(a.index) < 0) {
      throw new Error('adjustments 每项必须包含非负整数 index（0 起始），如 {"index": 0, "value": 0.15}');
    }
    if (!Number.isFinite(Number(a.value))) throw new Error("adjustments 每项必须包含数字 value（含义随形状类型不同，多为 0~1 的比例值）");
  }
  if (!supportsApi("1.10")) {
    throw new Error("形状调整值（Shape.adjustments）需要 PowerPointApi 1.10+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    target.load("adjustments/count");
    await ctx.sync();
    const count = Number(target.adjustments.count) || 0;
    if (!read) {
      for (const a of adj) {
        const i = Number(a.index);
        if (i >= count) {
          throw new Error(`调整点 index=${i} 超出范围（该形状共 ${count} 个调整点，index 从 0 开始；可先传 read=true 读取）`);
        }
        target.adjustments.set(i, Number(a.value));
      }
      try {
        await ctx.sync();
      } catch (e) {
        throw new Error(`设置形状调整值失败：${e?.message || e}`);
      }
    }
    // 读取全部调整值（ClientResult 批量取值，一次 sync）
    const readers = [];
    for (let i = 0; i < count; i++) readers.push(target.adjustments.get(i));
    try {
      await ctx.sync();
    } catch { /* 个别形状读取失败不影响设置结果 */ }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      count,
      values: readers.map((r) => { try { return r.value; } catch { return null; } }),
      applied: read ? undefined : Object.fromEntries(adj.map((a) => [String(a.index), Number(a.value)])),
      note: read ? "以上为该形状各调整点当前值（index 从 0 开始）" : "调整值含义随形状类型不同：圆角矩形 index0=圆角半径比例（0~0.5），箭头/饼图等各有含义；可先用 read=true 对比修改前后",
    };
  });
};

registry["ppt.alignShapes"] = async (args) => {
  const mode = String(args?.mode || "");
  if (!["left", "center", "right", "top", "middle", "bottom", "distribute_h", "distribute_v"].includes(mode)) {
    throw new Error("mode 必须是 left（左对齐）/ center（水平居中对齐）/ right（右对齐）/ top（顶对齐）/ middle（垂直居中对齐）/ bottom（底对齐）/ distribute_h（水平等距分布）/ distribute_v（垂直等距分布）之一");
  }
  if (args?.shape_ids == null) throw new Error("必须提供 shape_ids（对齐/分布是多个形状之间的操作，至少 2 个 id）");
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    if (list.length < 2) throw new Error(`对齐/分布至少需要 2 个形状（当前匹配 ${list.length} 个，请用 shape_ids 提供 2 个以上）`);
    const size = await slideSize(ctx);
    const items = list.map((s) => ({ s, left: Number(s.left), top: Number(s.top), w: Number(s.width), h: Number(s.height) }));
    const bad = items.find((i) => !Number.isFinite(i.left) || !Number.isFinite(i.top));
    if (bad) throw new Error(`形状 ${bad.s.id} 的位置信息读取失败（需要 PowerPointApi 1.4+）`);
    const L = Math.min(...items.map((i) => i.left));
    const R = Math.max(...items.map((i) => i.left + i.w));
    const T = Math.min(...items.map((i) => i.top));
    const B = Math.max(...items.map((i) => i.top + i.h));
    const avgCX = items.reduce((acc, i) => acc + i.left + i.w / 2, 0) / items.length;
    const avgCY = items.reduce((acc, i) => acc + i.top + i.h / 2, 0) / items.length;
    for (const it of items) {
      if (mode === "left") it.s.left = L;
      else if (mode === "right") it.s.left = R - it.w;
      else if (mode === "center") it.s.left = Math.max(0, avgCX - it.w / 2);
      else if (mode === "top") it.s.top = T;
      else if (mode === "bottom") it.s.top = B - it.h;
      else if (mode === "middle") it.s.top = Math.max(0, avgCY - it.h / 2);
    }
    let distributed;
    if (mode === "distribute_h" || mode === "distribute_v") {
      const horiz = mode === "distribute_h";
      const sorted = items.slice().sort((a, b) => (horiz ? a.left - b.left : a.top - b.top));
      const total = sorted.reduce((acc, i) => acc + (horiz ? i.w : i.h), 0);
      const span = horiz ? R - L : B - T;
      const gap = (span - total) / (sorted.length - 1);
      let cursor = horiz ? L : T;
      for (const it of sorted) {
        if (horiz) { it.s.left = cursor; cursor += it.w + gap; }
        else { it.s.top = cursor; cursor += it.h + gap; }
      }
      distributed = { gap: Math.round(gap * 100) / 100, order: sorted.map((i) => Number(i.s.id)) };
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`对齐/分布形状失败（位置写入需要 PowerPointApi 1.4+）：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      mode,
      count: items.length,
      shapeIds: items.map((i) => Number(i.s.id)),
      bounds: { left: Math.round(L * 100) / 100, top: Math.round(T * 100) / 100, right: Math.round(R * 100) / 100, bottom: Math.round(B * 100) / 100, pageWidth: size.w, pageHeight: size.h },
      distributed,
      note: mode.startsWith("distribute")
        ? "已按选定方向等距分布（首尾形状位置不变，中间均匀排布）"
        : "已按选定方式对齐（以形状组的整体边界/中心为基准，可配合 ppt_move_shape 微调）",
    };
  });
};

registry["ppt.formatTable"] = async (args) => {
  requireSingleSelector(args);
  const fmtKeys = [
    "header_fill", "header_color", "header_bold", "header_size",
    "cell_fill", "cell_color", "cell_font_size", "cell_bold",
    "band_fill", "band_start", "h_align", "v_align",
    "border_color", "border_weight", "border_dash",
    "merges", "col_widths", "row_heights", "delete_rows", "delete_cols",
  ];
  if (!fmtKeys.some((k) => args?.[k] != null)) {
    throw new Error("至少提供一个格式化参数（header_* / cell_* / band_fill / h_align / v_align / border_* / merges / col_widths / row_heights / delete_rows / delete_cols）");
  }
  if (!supportsApi("1.9")) {
    throw new Error("表格美化（单元格填充/字体/对齐/边框/合并/行高列宽/删行列）需要 PowerPointApi 1.9+（本机 Office 不支持，可先用 ppt_get_api_info 查看本机能力）");
  }
  const merges = args?.merges != null ? args.merges : null;
  if (merges != null) {
    if (!Array.isArray(merges) || !merges.length) throw new Error("merges 必须是非空数组，每项 [起始行, 起始列, 行跨度, 列跨度]");
    for (const m of merges) {
      if (!Array.isArray(m) || m.length !== 4 || m.some((v) => !Number.isInteger(Number(v)) || Number(v) < 1)) {
        throw new Error("merges 每项必须是 [起始行, 起始列, 行跨度, 列跨度]（整数，跨度 ≥1）");
      }
    }
  }
  const numList = (v, name) => {
    if (!Array.isArray(v)) throw new Error(`${name} 必须是数字数组`);
    return v.map((x) => {
      const n = Number(x);
      if (!Number.isFinite(n) || n <= 0) throw new Error(`${name} 每项必须是正数（磅）`);
      return n;
    });
  };
  return PowerPoint.run(async (ctx) => {
    const { shapes, index } = await getSlideShapes(ctx, args?.slide_index);
    const { list } = pickTargets(shapes, args);
    const target = list[0];
    if (String(target.type).toLowerCase() !== "table") {
      throw new Error(`目标形状（type=${target.type}）不是表格；可先用 ppt_list_shapes 找到 type=table 的形状再用本工具`);
    }
    if (!hasFn(target, "getTable")) throw new Error("读取表格（Shape.getTable）需要 PowerPointApi 1.8+");
    const table = target.getTable();
    table.load("rowCount,columnCount");
    await ctx.sync();
    const rows = Number(table.rowCount);
    const cols = Number(table.columnCount);
    const applied = { rows, cols };
    // 预取全部单元格并检测合并占位：*OrNullObject 规范要求 isNullObject 先 load 再 sync 才能读取，
    // 合并区域的非左上角单元格会返回 isNullObject=true 的对象（官方 Table.getCellOrNullObject 文档）
    const grid = [];
    for (let r = 0; r < rows; r++) {
      const rowCells = [];
      for (let c = 0; c < cols; c++) {
        const cell = table.getCellOrNullObject(r, c);
        cell.load("isNullObject");
        rowCells.push(cell);
      }
      grid.push(rowCells);
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`读取表格单元格失败（${rows} 行 × ${cols} 列）：${e?.message || e}`);
    }
    const eachCell = (r0, cb) => {
      for (let r = r0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (!grid[r][c].isNullObject) cb(grid[r][c], r, c); // 跳过合并区域的非左上角单元格
        }
      }
    };
    // 1) 数据区字体/填充（全表），2) 表头行覆盖第 0 行（后设置优先，二者可同时用）
    if (args?.cell_fill != null || args?.cell_color != null || args?.cell_font_size != null || args?.cell_bold != null) {
      eachCell(0, (cell) => {
        if (args.cell_fill != null) cell.fill.setSolidColor(String(args.cell_fill));
        if (args.cell_color != null) cell.font.color = String(args.cell_color);
        if (args.cell_font_size != null) cell.font.size = Number(args.cell_font_size);
        if (args.cell_bold != null) cell.font.bold = !!args.cell_bold;
      });
      applied.cell = { fill: args.cell_fill, color: args.cell_color, size: args.cell_font_size, bold: args.cell_bold };
    }
    if (args?.header_fill != null || args?.header_color != null || args?.header_bold != null || args?.header_size != null) {
      for (let c = 0; c < cols; c++) {
        const cell = grid[0][c];
        if (cell.isNullObject) continue;
        if (args.header_fill != null) cell.fill.setSolidColor(String(args.header_fill));
        if (args.header_color != null) cell.font.color = String(args.header_color);
        if (args.header_bold != null) cell.font.bold = !!args.header_bold;
        if (args.header_size != null) cell.font.size = Number(args.header_size);
      }
      applied.header = { fill: args.header_fill, color: args.header_color, bold: args.header_bold, size: args.header_size };
    }
    // 3) 隔行条纹（band_start 起每隔一行；默认从第 1 行开始，跳过表头）
    if (args?.band_fill != null) {
      const start = Math.min(Math.max(Number(args.band_start ?? 1) || 0, 0), Math.max(rows - 1, 0));
      for (let r = start; r < rows; r += 2) {
        for (let c = 0; c < cols; c++) {
          const cell = grid[r][c];
          if (!cell.isNullObject) cell.fill.setSolidColor(String(args.band_fill));
        }
      }
      applied.band = { color: String(args.band_fill), startRow: start };
    }
    // 4) 全表对齐
    if (args?.h_align != null || args?.v_align != null) {
      eachCell(0, (cell) => {
        if (args.h_align != null) cell.horizontalAlignment = String(args.h_align);
        if (args.v_align != null) cell.verticalAlignment = String(args.v_align);
      });
      applied.align = { h: args.h_align ?? undefined, v: args.v_align ?? undefined };
    }
    // 5) 边框（四边；Border.color/weight/dashStyle，1.9）
    if (args?.border_color != null || args?.border_weight != null || args?.border_dash != null) {
      eachCell(0, (cell) => {
        for (const side of ["top", "bottom", "left", "right"]) {
          const b = cell.borders?.[side];
          if (!b) continue;
          if (args.border_color != null) b.color = String(args.border_color);
          if (args.border_weight != null) b.weight = Number(args.border_weight);
          if (args.border_dash != null) b.dashStyle = String(args.border_dash);
        }
      });
      applied.border = { color: args.border_color, weight: args.border_weight, dash: args.border_dash };
    }
    // 6) 合并单元格（Table.mergeCells）
    if (merges) {
      for (const [r, c, rc, cc] of merges) {
        if (r - 1 + rc > rows || c - 1 + cc > cols) {
          throw new Error(`合并区域 [${r},${c},${rc},${cc}] 超出表格范围（${rows} 行 × ${cols} 列；行列从 1 起）`);
        }
        table.mergeCells(r - 1, c - 1, rc, cc);
      }
      applied.merges = merges;
    }
    // 7) 行高/列宽（TableRow.height / TableColumn.width）
    if (args?.col_widths != null) {
      const ws = numList(args.col_widths, "col_widths");
      if (ws.length > cols) throw new Error(`col_widths 长度（${ws.length}）超过表格列数（${cols}）`);
      table.columns.load("items");
      await ctx.sync();
      ws.forEach((w, c) => { table.columns.items[c].width = w; });
      applied.colWidths = ws;
    }
    if (args?.row_heights != null) {
      const hs = numList(args.row_heights, "row_heights");
      if (hs.length > rows) throw new Error(`row_heights 长度（${hs.length}）超过表格行数（${rows}）`);
      table.rows.load("items");
      await ctx.sync();
      hs.forEach((h, r) => { table.rows.items[r].height = h; });
      applied.rowHeights = hs;
    }
    // 8) 删除行列（从大到小删，避免索引移动；TableClear/TableRow.delete 1.9）
    const delIdx = (v, name, max) => {
      if (!Array.isArray(v)) throw new Error(`${name} 必须是索引数组（0 起始）`);
      const ids = v.map((x) => Number(x));
      if (ids.some((n) => !Number.isInteger(n) || n < 0 || n >= max)) {
        throw new Error(`${name} 中存在越界索引（表格共 ${max} 行/列，0 起始）`);
      }
      return ids.sort((a, b) => b - a);
    };
    if (args?.delete_cols != null) {
      const cs = delIdx(args.delete_cols, "delete_cols", cols);
      table.columns.load("items");
      await ctx.sync();
      cs.forEach((c) => { table.columns.items[c].delete(); });
      applied.deletedCols = cs.slice().sort((a, b) => a - b);
    }
    if (args?.delete_rows != null) {
      const rs = delIdx(args.delete_rows, "delete_rows", rows);
      table.rows.load("items");
      await ctx.sync();
      rs.forEach((r) => { table.rows.items[r].delete(); });
      applied.deletedRows = rs.slice().sort((a, b) => a - b);
    }
    try {
      await ctx.sync();
    } catch (e) {
      throw new Error(`表格格式化失败：${e?.message || e}`);
    }
    return {
      ok: true,
      slideIndex: index,
      shapeId: Number(target.id),
      ...applied,
      note: "表格美化完成；单元格填充/字体/边框为 PowerPointApi 1.9+ 能力，行高列宽按磅设置（小于最小值会被 Office 自动提升）",
    };
  });
};

/* ---------------- 能力自检（自动识别当前 Office 可用的 PowerPoint 接口） ----------------
 * 依据：Office.context.requirements.isSetSupported("PowerPointApi", v)（官方运行时检测）
 * 结果供三处使用：
 *  1) ppt_get_api_info 工具（模型可随时查询本机能力）
 *  2) 任务窗格启动时检测 → 随 /api/chat 的 caps 上报
 *  3) 服务端据此从 LLM 绑定中剔除本机不可用的工具，并在系统提示词注入能力摘要
 */

/** PowerPointApi 探测项：api=最低需求集版本，label=注入提示词的中文能力名 */
const PPT_CAP_CHECKS = [
  { key: "delete_shape", label: "删除形状", api: "1.3" },
  { key: "shape_format", label: "形状移动/缩放/填充/边框/文字格式", api: "1.4" },
  { key: "add_shape", label: "添加几何形状/线条/文本框", api: "1.4" },
  { key: "select_shapes", label: "选中形状/文本", api: "1.5" },
  { key: "slide_follow", label: "视图跟随（修改后自动切换到目标页）", api: "1.5" },
  { key: "zorder_group", label: "形状层级/组合/解组", api: "1.8" },
  { key: "add_table", label: "插入表格", api: "1.8" },
  { key: "slide_ops", label: "新建/移动幻灯片、应用版式、导出页面图", api: "1.8" },
  { key: "fill_image", label: "图片填充形状", api: "1.8" },
  { key: "table_format", label: "表格美化（单元格填充/字体/对齐/边框/合并/行列尺寸）", api: "1.9" },
  { key: "shape_image", label: "导出形状图片内容", api: "1.10" },
  { key: "slide_background", label: "页面背景设置（纯色/图片/渐变）", api: "1.10" },
  { key: "rotation", label: "形状旋转", api: "1.10" },
  { key: "visible", label: "形状显示/隐藏", api: "1.10" },
  { key: "alt_text", label: "无障碍文本", api: "1.10" },
  { key: "hyperlink", label: "形状/文字超链接", api: "1.10" },
  { key: "shape_adjustments", label: "形状调整值（圆角半径等）", api: "1.10" },
  { key: "bullet_style", label: "编号与项目符号类型/样式/缩进（含中文编号）", api: "1.10" },
];

/** Office.js 明确没有对应 API 的能力（避免模型反复尝试） */
const PPT_UNSUPPORTED = [
  { key: "slide_hidden", label: "隐藏幻灯片", reason: "PowerPoint.js 未提供幻灯片隐藏 API（Slide.hidden 不存在，Preview 也没有）；只能用 ppt_set_shape_visible 隐藏单个形状" },
  { key: "crop_image", label: "文档内裁剪图片", reason: "PowerPoint.js 无裁剪 API；流程为 ppt_get_image 导出 → image_process 裁剪 → ppt_insert_image 插回" },
  { key: "insert_image", label: "插入图片", reason: "仅 Preview API 提供（addImage/addPicture），正式版 Office 上会明确报错；导出图片内容（ppt_get_image）不受影响" },
  { key: "shape_gradient", label: "形状渐变填充", reason: "ShapeFill 无写入渐变的 API（官方功能请求排队中）；替代：用 image_process 生成渐变 PNG 后 ppt_set_shape_fill_image，或页面背景用 ppt_set_slide_background 的 gradient_type" },
  { key: "shape_shadow", label: "形状阴影", reason: "PowerPoint.js 的 Shape 无 shadow 属性（Excel/Word 有，PowerPoint 没有）" },
  { key: "line_arrow", label: "线条箭头样式", reason: "ShapeLineFormat 无 begin/end 箭头属性；需要箭头请用 kind=geometric 的箭头形状（rightArrow/leftArrow 等）或 image_process 生成箭头图片" },
  { key: "para_spacing", label: "段落行距/段距", reason: "ParagraphFormat 无行距（spaceWithin）与段前段后距属性；可用 auto_size + 形状尺寸间接控制文本密度" },
  { key: "read_geometry", label: "读取已有几何形状的具体类型", reason: "GeometricShapeType 只写不可读（官方功能请求排队中）；AI 添加形状时已知道自己用了什么 geometry，读取未知形状时只有 type=geometricShape" },
];

/**
 * 检测本机 PowerPoint API 能力（不修改文档，可在启动时与工具内随时调用）
 * @returns {{host:"ppt", maxVersion:string|null, hostVersion:string|undefined, platform:string|undefined, capabilities:Array, unsupported:Array, detectedAt:number}}
 */
export async function detectPptCapabilities() {
  const result = {
    host: "ppt",
    maxVersion: null,
    hostVersion: undefined,
    platform: undefined,
    capabilities: [],
    unsupported: PPT_UNSUPPORTED,
    detectedAt: Date.now(),
  };
  try {
    const d = Office.context?.diagnostics;
    result.hostVersion = d?.version;
    result.platform = d?.platform;
  } catch { /* noop */ }
  // isSetSupported 累积向后兼容：从 1.1 逐级测到 1.10，取最后一个支持的版本
  const versions = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"];
  for (const v of versions) {
    if (supportsApi(v)) result.maxVersion = v;
    else break;
  }
  const maxN = result.maxVersion ? verNum(result.maxVersion) : 0;
  result.capabilities = PPT_CAP_CHECKS.map((c) => ({ ...c, ok: maxN > 0 && verNum(c.api) <= maxN }));
  return result;
}

registry["ppt.getApiInfo"] = async () => {
  const caps = await detectPptCapabilities();
  const okList = caps.capabilities.filter((c) => c.ok).map((c) => c.label);
  const badList = caps.capabilities.filter((c) => !c.ok).map((c) => `${c.label}（需 ${c.api}+）`);
  return {
    ok: true,
    maxVersion: caps.maxVersion,
    hostVersion: caps.hostVersion,
    platform: caps.platform,
    available: okList,
    unavailable: badList,
    notInOfficeJs: caps.unsupported,
    note: "maxVersion 为 null 表示 PowerPointApi 1.1 也不可用（Office 过旧），大部分 ppt_* 工具无法工作",
  };
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
// detectPptCapabilities 已在上方内联 export

/* ---------------- 视图跟随（AI 调整哪页，Office 自动切换到哪页） ----------------
 * 需求：AI 对某一页执行修改类操作后，把该页设为当前选中页，用户能实时看到 AI 正在编辑的位置。
 * 实现：Slide.setSelected()（PowerPointApi 1.5 选择类 API，官方"What's new in 1.5"）；
 * 在下列修改类执行器成功返回后，再开一次 PowerPoint.run 调用 setSelected。
 * 切换失败（本机不支持 1.5+ / 演示文稿视图受限等）静默忽略，绝不影响主操作结果。
 */
const SLIDE_FOLLOW_KEYS = new Set([
  "ppt.addSlide", "ppt.deleteSlide", "ppt.moveSlide", "ppt.applySlideLayout",
  "ppt.insertTextbox", "ppt.insertImage", "ppt.insertTable", "ppt.addShape",
  "ppt.setShapeText", "ppt.deleteShape", "ppt.moveShape", "ppt.formatShape", "ppt.formatText",
  "ppt.adjustTextbox", "ppt.deleteTextbox",
  "ppt.setSlideBackground", "ppt.setSlideNotes",
  "ppt.setShapeZorder", "ppt.ungroupShape", "ppt.groupShapes",
  "ppt.setShapeVisible", "ppt.setShapeFillImage", "ppt.setShapeHyperlink", "ppt.setShapeAltText",
  "ppt.setShapeAdjustments", "ppt.alignShapes", "ppt.formatTable",
]);

/** 把指定页设为当前选中页（失败静默：旧版 Office 无 setSelected / sync 出错时直接跳过） */
async function selectSlideQuiet(slideIndex) {
  try {
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(slideIndex);
      if (!hasFn(slide, "setSelected")) return;
      slide.setSelected();
      await ctx.sync();
    });
  } catch { /* 视图切换失败不影响主操作 */ }
}

for (const [key, fn] of Object.entries(registry)) {
  if (!SLIDE_FOLLOW_KEYS.has(key)) continue;
  registry[key] = async (args, ctxInfo) => {
    const result = await fn(args, ctxInfo);
    try {
      let idx;
      if (key === "ppt.deleteSlide") {
        // 删除后切到被删位置的相邻页；删光（remaining=0）时不切
        idx = Number(result?.remaining) > 0 ? Math.min(Number(result?.deletedIndex) || 0, Number(result?.remaining) - 1) : null;
      } else if (key === "ppt.moveSlide") {
        idx = Number(result?.to);
      } else {
        idx = Number(result?.slideIndex);
      }
      if (Number.isInteger(idx) && idx >= 0) await selectSlideQuiet(idx);
    } catch { /* 跟随切换失败不影响工具结果 */ }
    return result;
  };
}

export { registry };
