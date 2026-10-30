/**
 * 图片处理核心（服务端）
 * ---------------------------------------------------------------
 * - resolveWorkspacePath: 对话工作区内路径的安全解析（拒绝绝对路径与 .. 越界）
 * - sniffImageExt: 按魔数识别图片格式（png/jpeg/gif/bmp/webp/tiff）
 * - processImage: 基于 jimp 的图片处理（缩放/裁剪/旋转/翻转/转格式/压缩）
 *
 * 使用方：
 *  - server/index.js 的 image_process 工具（serverToolExecutor 分支）
 *  - /api/workspace/file 端点（路径安全 + 格式嗅探）
 *  - 附件上传的图片格式识别与贴图落盘
 */
import path from "node:path";
import fs from "node:fs";
import { Jimp } from "jimp";

/** jimp v1 可解码/编码的图片扩展名（webp 不支持） */
export const IMAGE_EXTS = ["png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff"];
/** 附件/贴图额外接受的图片格式（可存盘，但 image_process 不能处理 webp/tiff 之外的限制见 IMAGE_EXTS） */
export const IMAGE_FILE_EXTS = [...IMAGE_EXTS, "webp"];
export const MAX_IMAGE_MB = 20;

/**
 * 对话工作区内相对路径的安全解析
 * @param {string} convDir 对话工作区绝对路径
 * @param {string} relPath 相对路径（正/反斜杠均可）
 * @returns {{ abs: string, rel: string }} abs=绝对路径；rel=规整后的相对路径（/ 分隔）
 */
export function resolveWorkspacePath(convDir, relPath) {
  const raw = String(relPath || "").trim().replace(/\\/g, "/");
  if (!raw) throw new Error("缺少文件路径");
  if (/^[a-zA-Z]:/.test(raw) || raw.startsWith("//") || raw.startsWith("/")) {
    throw new Error(`必须是工作区内的相对路径：${raw}`);
  }
  const parts = raw.split("/").filter((p) => p.length && p !== ".");
  if (!parts.length) throw new Error("路径无效");
  if (parts.some((p) => p === "..")) throw new Error("路径不允许包含 ..（不能越出对话工作区）");
  const abs = path.resolve(convDir, ...parts);
  const normConv = path.resolve(convDir);
  if (abs !== normConv && !abs.startsWith(normConv + path.sep)) {
    throw new Error("路径越界（不在对话工作区内）");
  }
  return { abs, rel: parts.join("/") };
}

/** 按文件头魔数识别图片格式，返回小写扩展名或 null */
export function sniffImageExt(buf) {
  const b = buf instanceof Buffer ? buf : Buffer.from(buf || []);
  if (b.length < 12) return null;
  // PNG: 89 50 4E 47 0D 0A 1A 0A
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "png";
  // JPEG: FF D8 FF
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "jpg";
  // GIF: GIF87a / GIF89a
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return "gif";
  // BMP: 42 4D
  if (b[0] === 0x42 && b[1] === 0x4d) return "bmp";
  // WEBP: RIFF....WEBP
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return "webp";
  // TIFF: II*\0 / MM\0*
  if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00)) return "tif";
  return null;
}

/** 客户端侧：按 base64 前缀嗅探图片格式（解码首部 32 字节判断魔数） */
export function imageExtFromBase64(b64) {
  const s = String(b64 || "");
  const head = s.slice(0, 48);
  try {
    return sniffImageExt(Buffer.from(head, "base64"));
  } catch {
    return null;
  }
}

/**
 * 处理对话工作区内的图片文件（image_process 工具执行核心）
 * @param {object} args 工具参数（见 tools-catalog.js 中 image_process 定义）
 * @param {{ convDir: string }} binding 服务端绑定（对话工作区）
 * @returns {Promise<{ok:boolean, file:string, width:number, height:number, size:number, format:string, operations:object}>}
 */
export async function processImage(args = {}, { convDir } = {}) {
  if (!convDir) throw new Error("缺少对话工作区绑定");
  const file = String(args.file || "").trim();
  if (!file) throw new Error("缺少 file（工作区内图片文件的相对路径）");

  const ops = {};
  const hasResize = args.resize_width != null || args.resize_height != null || args.resize_scale != null;
  const hasCrop = [args.crop_x, args.crop_y, args.crop_w, args.crop_h].some((v) => v != null);
  if (hasResize) ops.resize = true;
  if (hasCrop) ops.crop = true;
  if (args.rotate != null) ops.rotate = Number(args.rotate);
  if (args.flip_h) ops.flip_h = true;
  if (args.flip_v) ops.flip_v = true;
  if (args.format) ops.format = String(args.format).toLowerCase();
  if (!Object.keys(ops).length) {
    throw new Error("未指定任何处理操作：至少提供 resize_* / crop_* / rotate / flip_* / format 之一");
  }

  const resolved = resolveWorkspacePath(convDir, file);
  if (!fs.existsSync(resolved.abs) || !fs.statSync(resolved.abs).isFile()) {
    throw new Error(`文件不存在：${resolved.rel}（可先在工作区中确认，如通过 shell 工具 ls）`);
  }
  const input = fs.readFileSync(resolved.abs);
  if (input.length > MAX_IMAGE_MB * 1024 * 1024) {
    throw new Error(`图片过大（>${MAX_IMAGE_MB}MB）：${resolved.rel}`);
  }
  const realExt = sniffImageExt(input);
  if (!realExt) throw new Error(`不是可识别的图片文件：${resolved.rel}（支持 png/jpg/gif/bmp/webp/tiff）`);
  if (realExt === "webp") {
    throw new Error(`webp 输入暂不支持处理（图片库限制）：${resolved.rel}。请先用其他方式转为 png/jpg（例如 shell 调用 ffmpeg），或改用其他图片`);
  }

  let img;
  try {
    img = await Jimp.read(input);
  } catch (e) {
    throw new Error(`图片解码失败（${realExt}）：${e?.message || e}`);
  }
  const origW = img.bitmap.width;
  const origH = img.bitmap.height;
  const applied = {};

  // 缩放：优先 scale，其次 width/height（只给一边则等比）
  if (hasResize) {
    let w = null;
    let h = null;
    if (args.resize_scale != null) {
      const scale = Number(args.resize_scale);
      if (!Number.isFinite(scale) || scale <= 0 || scale > 10) throw new Error("resize_scale 必须是 0~10 的正数");
      w = Math.max(1, Math.round(origW * scale));
      h = Math.max(1, Math.round(origH * scale));
    } else {
      if (args.resize_width != null) {
        w = Number(args.resize_width);
        if (!Number.isFinite(w) || w < 1 || w > 20000) throw new Error("resize_width 必须是 1~20000 的像素值");
        w = Math.round(w);
      }
      if (args.resize_height != null) {
        h = Number(args.resize_height);
        if (!Number.isFinite(h) || h < 1 || h > 20000) throw new Error("resize_height 必须是 1~20000 的像素值");
        h = Math.round(h);
      }
      if (w != null && h == null) h = Math.max(1, Math.round(origH * (w / origW)));
      if (h != null && w == null) w = Math.max(1, Math.round(origW * (h / origH)));
    }
    img.resize({ w, h });
    applied.resize = { width: w, height: h };
  }

  // 裁剪：默认像素，unit=ratio 时为 0~1 相对比例
  if (hasCrop) {
    const unit = args.crop_unit === "ratio" ? "ratio" : "px";
    const baseW = origW;
    const baseH = origH;
    const pick = (v, base, name) => {
      if (v == null) throw new Error(`裁剪缺少 crop_${name}`);
      const n = Number(v);
      if (!Number.isFinite(n)) throw new Error(`crop_${name} 必须是数字`);
      if (unit === "ratio") {
        if (n < 0 || n > 1) throw new Error(`crop_${name} 在 ratio 模式下必须是 0~1 的比例值`);
        return Math.round(n * base);
      }
      return Math.round(n);
    };
    let x = pick(args.crop_x, baseW, "x");
    let y = pick(args.crop_y, baseH, "y");
    let w = pick(args.crop_w, baseW, "w");
    let h = pick(args.crop_h, baseH, "h");
    const curW = img.bitmap.width;
    const curH = img.bitmap.height;
    x = Math.min(Math.max(0, x), curW - 1);
    y = Math.min(Math.max(0, y), curH - 1);
    w = Math.min(Math.max(1, w), curW - x);
    h = Math.min(Math.max(1, h), curH - y);
    img.crop({ x, y, w, h });
    applied.crop = { x, y, width: w, height: h };
  }

  // 旋转（顺时针，度）
  if (args.rotate != null) {
    const deg = Number(args.rotate);
    if (!Number.isFinite(deg)) throw new Error("rotate 必须是数字（度）");
    img.rotate(deg);
    applied.rotate = deg;
  }
  if (args.flip_h) {
    img.flip({ horizontal: true, vertical: false });
    applied.flip_h = true;
  }
  if (args.flip_v) {
    img.flip({ horizontal: false, vertical: true });
    applied.flip_v = true;
  }

  // 输出格式：优先 format 参数 → 输出文件扩展名 → 保持输入格式
  const fmtMap = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", bmp: "image/bmp", gif: "image/gif", tif: "image/tiff", tiff: "image/tiff" };
  let ext = realExt;
  if (ops.format) {
    if (!fmtMap[ops.format]) throw new Error(`format 仅支持 png/jpeg/bmp/gif/tiff，收到：${ops.format}`);
    ext = ops.format === "jpeg" ? "jpg" : ops.format;
  }

  // 输出文件：默认 <stem>_<op>.<ext>，自动避免覆盖；显式 output 允许覆盖
  let outRel;
  if (args.output != null && String(args.output).trim()) {
    const o = resolveWorkspacePath(convDir, String(args.output).trim());
    if (!/\.[a-z0-9]+$/i.test(o.rel)) throw new Error(`output 需要带扩展名（如 out.png）：${o.rel}`);
    outRel = o;
    const outExt = (o.rel.split(".").pop() || "").toLowerCase();
    if (fmtMap[outExt]) ext = outExt === "jpeg" ? "jpg" : outExt;
  } else {
    const stem = resolved.rel.replace(/\.[^.]+$/, "").replace(/[\\/]/g, "_").slice(-40) || "image";
    const tag = Object.keys(applied).map((k) => (k === "resize" ? `r${applied.resize.width}x${applied.resize.height}` : k === "crop" ? `crop` : k === "rotate" ? `rot${applied.rotate}` : k)).join("-");
    let name = `${stem}_${tag}.${ext}`;
    let n = 2;
    while (fs.existsSync(path.join(convDir, name))) name = `${stem}_${tag}_${n++}.${ext}`;
    outRel = { abs: path.join(convDir, name), rel: name };
  }

  const mime = fmtMap[ext] || "image/png";
  const quality = args.quality != null ? Math.min(100, Math.max(1, Number(args.quality))) : null;
  let outBuf;
  try {
    outBuf = quality != null && mime === "image/jpeg" ? await img.getBuffer(mime, { quality }) : await img.getBuffer(mime);
  } catch (e) {
    throw new Error(`图片编码失败（${ext}）：${e?.message || e}`);
  }
  fs.mkdirSync(path.dirname(outRel.abs), { recursive: true });
  fs.writeFileSync(outRel.abs, outBuf);

  return {
    ok: true,
    file: outRel.rel,
    width: img.bitmap.width,
    height: img.bitmap.height,
    size: outBuf.length,
    format: ext,
    sourceFile: resolved.rel,
    sourceSize: { width: origW, height: origH },
    operations: applied,
    note: "输出为工作区新文件（未覆盖原图）。可用 word/ppt/excel 插入图片工具的 file 参数把该文件插入文档。",
  };
}
