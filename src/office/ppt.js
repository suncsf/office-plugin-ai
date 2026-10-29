/**
 * PowerPoint 执行器（Office.js PowerPoint API）
 * 说明：PowerPoint JS API 的部分能力（文本读写、添加形状等）要求较新的 API 集合
 * （PowerPointApi 1.8+，Office 2024 / 新版 Microsoft 365 满足），全部做了特性检测，
 * 不支持时返回明确的错误信息。
 */
import { stripDataUrl } from "./index";

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
        const masters = ctx.presentation.slideMasters;
        masters.load("items");
        await ctx.sync();
        layouts = [];
        for (const m of masters.items) {
          m.layouts.load("id,name");
          await ctx.sync();
          for (const l of m.layouts.items) layouts.push({ id: l.id, name: l.name });
        }
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

registry["ppt.insertImage"] = async (args) => {
  const b64 = stripDataUrl(args?.base64);
  if (!b64) throw new Error("base64 不能为空");
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

export { registry };
