/**
 * 图片能力离线自测（不需要真实 Office / LLM）
 *   node test/image-process-test.mjs
 * 覆盖：jimp 处理链（缩放/裁剪/旋转/转格式/压缩/翻转）、工作区路径安全、格式嗅探、输出命名防覆盖。
 */
import os from "node:os";
import path from "node:path";
import fs from "node:fs";
import { Jimp } from "jimp";
import {
  processImage, resolveWorkspacePath, sniffImageExt, imageExtFromBase64,
  IMAGE_EXTS, IMAGE_FILE_EXTS,
} from "../server/image.js";

let failed = false;
function check(name, cond, detail = "") {
  if (cond) console.log(`  ✅ ${name}`);
  else {
    failed = true;
    console.error(`  ❌ ${name} ${detail}`);
  }
}

/* 临时"对话工作区" */
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "officeai-img-"));

/* 造 200x100 纯色 PNG + JPEG 测试图 */
const srcPng = new Jimp({ width: 200, height: 100, color: 0x3366aaff });
const pngBuf = await srcPng.getBuffer("image/png");
fs.writeFileSync(path.join(tmp, "test.png"), pngBuf);
const jpegBuf = await new Jimp({ width: 80, height: 60, color: 0xff8800ff }).getBuffer("image/jpeg");
fs.writeFileSync(path.join(tmp, "photo.jpg"), jpegBuf);

console.log("\n[场景1] 缩放");
{
  const r = await processImage({ file: "test.png", resize_width: 100 }, { convDir: tmp });
  check("等比缩放 200x100 → 100x50", r.width === 100 && r.height === 50, JSON.stringify(r));
  check("输出到工作区新文件（不覆盖原图）", r.file !== "test.png" && fs.existsSync(path.join(tmp, r.file)) && fs.existsSync(path.join(tmp, "test.png")));
  check("返回尺寸与实际文件一致", (async () => {
    const back = await Jimp.read(path.join(tmp, r.file));
    return back.bitmap.width === 100 && back.bitmap.height === 50;
  })());

  const r2 = await processImage({ file: "test.png", resize_height: 200 }, { convDir: tmp });
  check("只给高 → 等比 400x200", r2.width === 400 && r2.height === 200, JSON.stringify(r2));

  const r3 = await processImage({ file: "test.png", resize_scale: 0.5 }, { convDir: tmp });
  check("scale 0.5 → 100x50", r3.width === 100 && r3.height === 50, JSON.stringify(r3));

  const r4 = await processImage({ file: "test.png", resize_width: 120, resize_height: 30 }, { convDir: tmp });
  check("宽高同时给定 → 非等比 120x30", r4.width === 120 && r4.height === 30, JSON.stringify(r4));
}

console.log("\n[场景2] 裁剪");
{
  const c = await processImage({ file: "test.png", crop_x: 10, crop_y: 10, crop_w: 50, crop_h: 40 }, { convDir: tmp });
  check("像素裁剪 50x40", c.width === 50 && c.height === 40, JSON.stringify(c));

  const c2 = await processImage({ file: "test.png", crop_x: 0, crop_y: 0, crop_w: 0.5, crop_h: 0.5, crop_unit: "ratio" }, { convDir: tmp });
  check("比例裁剪（左上四分之一）100x50", c2.width === 100 && c2.height === 50, JSON.stringify(c2));

  const c3 = await processImage({ file: "test.png", crop_x: 0, crop_y: 0, crop_w: 9999, crop_h: 1, crop_y2: null, crop_unit: "px" }, { convDir: tmp });
  check("越界裁剪收敛到图内", c3.width === 200 && c3.height === 1, JSON.stringify(c3));

  const bad = await processImage({ file: "test.png", crop_x: 10, crop_y: 10 }, { convDir: tmp }).catch((e) => e);
  check("缺裁剪尺寸时明确报错", bad instanceof Error && /crop_w|crop_h/.test(bad.message), String(bad));
}

console.log("\n[场景3] 旋转 / 翻转 / 格式转换 / 压缩");
{
  const rot = await processImage({ file: "test.png", rotate: 90 }, { convDir: tmp });
  check("顺时针旋转 90° 交换宽高 → 100x200", rot.width === 100 && rot.height === 200, JSON.stringify(rot));

  const conv = await processImage({ file: "test.png", format: "jpeg", quality: 60 }, { convDir: tmp });
  check("转 jpeg 输出 .jpg", /\.(jpg|jpeg)$/.test(conv.file) && conv.format === "jpg", JSON.stringify(conv));
  const q90 = await processImage({ file: "test.png", format: "jpeg", quality: 95, output: "q95.jpg" }, { convDir: tmp });
  const q20 = await processImage({ file: "test.png", format: "jpeg", quality: 20, output: "q20.jpg" }, { convDir: tmp });
  check("quality 生效（q20 < q95）", q20.size < q90.size, JSON.stringify({ q90: q90.size, q20: q20.size }));

  const flip = await processImage({ file: "test.png", flip_h: true }, { convDir: tmp });
  check("水平翻转保持尺寸", flip.width === 200 && flip.height === 100, JSON.stringify(flip));

  const combo = await processImage({ file: "photo.jpg", resize_scale: 0.5, rotate: 90, format: "png" }, { convDir: tmp });
  check("组合操作（缩放+旋转+转png）30x40", combo.width === 30 && combo.height === 40, JSON.stringify(combo));
}

console.log("\n[场景4] 输出命名与覆盖");
{
  const a = await processImage({ file: "test.png", resize_scale: 0.25 }, { convDir: tmp });
  const b = await processImage({ file: "test.png", resize_scale: 0.25 }, { convDir: tmp });
  check("默认命名防覆盖（自动加序号）", a.file !== b.file, JSON.stringify({ a: a.file, b: b.file }));

  const ov = await processImage({ file: "test.png", resize_scale: 0.25, output: "fixed.png" }, { convDir: tmp });
  const ov2 = await processImage({ file: "test.png", resize_scale: 0.25, output: "fixed.png" }, { convDir: tmp });
  check("显式 output 允许覆盖同名", ov.file === "fixed.png" && ov2.file === "fixed.png");

  const noOp = await processImage({ file: "test.png" }, { convDir: tmp }).catch((e) => e);
  check("无操作参数报错", noOp instanceof Error && /操作/.test(noOp.message), String(noOp));
}

console.log("\n[场景5] 输入校验与格式嗅探");
{
  const nf = await processImage({ file: "nope.png", resize_scale: 0.5 }, { convDir: tmp }).catch((e) => e);
  check("文件不存在报错", nf instanceof Error && /不存在/.test(nf.message), String(nf));

  fs.writeFileSync(path.join(tmp, "fake.png"), Buffer.from("hello world, not an image"));
  const fake = await processImage({ file: "fake.png", resize_scale: 0.5 }, { convDir: tmp }).catch((e) => e);
  check("非图片文件报错", fake instanceof Error && /图片/.test(fake.message), String(fake));

  // webp：真实 RIFF/WEBP 头（jimp 不支持处理，应明确报错）
  const webpHeader = Buffer.alloc(16);
  webpHeader.write("RIFF", 0, "ascii");
  webpHeader.writeUInt32LE(8, 4);
  webpHeader.write("WEBPVP8 ", 8, "ascii");
  fs.writeFileSync(path.join(tmp, "pic.webp"), webpHeader);
  check("sniff 识别 webp", sniffImageExt(fs.readFileSync(path.join(tmp, "pic.webp"))) === "webp");
  const wp = await processImage({ file: "pic.webp", resize_scale: 0.5 }, { convDir: tmp }).catch((e) => e);
  check("webp 输入明确报错（非静默失败）", wp instanceof Error && /webp/i.test(wp.message), String(wp));

  check("sniff png", sniffImageExt(pngBuf) === "png");
  check("sniff jpeg", sniffImageExt(jpegBuf) === "jpg");
  check("sniff gif（合成头）", sniffImageExt(Buffer.from("GIF89a" + "x".repeat(16), "utf8")) === "gif");
  check("sniff bmp", sniffImageExt(await new Jimp({ width: 4, height: 4 }).getBuffer("image/bmp")) === "bmp");
  check("sniff 拒绝任意文本", sniffImageExt(Buffer.from("plain text")) === null);
  check("base64 嗅探 png", imageExtFromBase64(pngBuf.toString("base64")) === "png");
  check("IMAGE_EXTS 白名单", IMAGE_EXTS.includes("png") && IMAGE_EXTS.includes("jpg") && !IMAGE_EXTS.includes("webp"));
  check("附件白名单含 webp", IMAGE_FILE_EXTS.includes("webp"));
}

console.log("\n[场景6] 工作区路径安全");
{
  const safe = (p) => {
    try {
      resolveWorkspacePath(tmp, p);
      return "ok";
    } catch (e) {
      return "rejected";
    }
  };
  check("正常相对路径放行", safe("images/a.png") === "ok");
  check("反斜杠相对路径放行", safe("images\\a.png") === "ok");
  check("单层 .. 拒绝", safe("../x.png") === "rejected");
  check("深层 .. 拒绝", safe("a/b/../../../x.png") === "rejected");
  check("盘符绝对路径拒绝", safe("C:\\Windows\\win.ini") === "rejected");
  check("根路径拒绝", safe("/Windows/win.ini") === "rejected");
  check("UNC 路径拒绝", safe("\\\\server\\share\\x.png") === "rejected");
  check("空路径拒绝", safe("") === "rejected");
  check("当前目录点路径拒绝", safe(".") === "rejected");
  const r = resolveWorkspacePath(tmp, "images/sub/pic.PNG");
  check("正常路径解析到工作区内", r.abs.startsWith(path.resolve(tmp)) && r.rel === "images/sub/pic.PNG", JSON.stringify(r));
  const resolvedInside = resolveWorkspacePath(tmp, "images/./b.png");
  check("./ 段被过滤", resolvedInside.rel === "images/b.png", JSON.stringify(resolvedInside));
}

console.log("\n[场景7] 跨目录隔离（不同对话工作区互不可见）");
{
  const other = fs.mkdtempSync(path.join(os.tmpdir(), "officeai-img2-"));
  fs.writeFileSync(path.join(other, "secret.png"), pngBuf);
  const cross = await processImage({ file: "../" + path.basename(other) + "/secret.png", resize_scale: 0.5 }, { convDir: tmp }).catch((e) => e);
  check("跨工作区读取被拒绝", cross instanceof Error && /\.\./.test(cross.message), String(cross));
  fs.rmSync(other, { recursive: true, force: true });
}

fs.rmSync(tmp, { recursive: true, force: true });
console.log(failed ? "\n❌ 存在失败项" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
