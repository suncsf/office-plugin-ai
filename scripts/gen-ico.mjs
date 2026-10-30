#!/usr/bin/env node
/**
 * 由 launcher/icon-source.png（1024² AI 图标）生成 launcher/app.ico
 *  - 去除右下角「AI生 WORKBUDDY」水印（按行左移取干净渐变，保留目标 alpha 以维持圆角透明）
 *  - 输出多尺寸 ICO（PNG 内嵌，Vista+ 支持）：16/24/32/48/64/128/256
 * 依赖：jimp（生产依赖已含）
 * 编译时由 scripts/package.mjs 引用 /win32icon:launcher\app.ico
 */
import { Jimp } from "jimp";
import { writeFileSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SRC = join(ROOT, "launcher", "icon-source.png");
const OUT = join(ROOT, "launcher", "app.ico");

/* ---------- 1. 读取并去水印 ---------- */
const img = await Jimp.read(SRC);
const W = img.bitmap.width, H = img.bitmap.height;
const data = img.bitmap.data; // RGBA
console.log(`源图 ${W}x${H}`);

if (W !== 1024 || H !== 1024) {
  // 水印坐标按 1024 设计，其它尺寸按比例换算
  const scale = W / 1024;
  // 下面坐标按换算后处理
  img.bitmap.width = W; // 保持
}

// 水印区（按 1024 坐标系，左移 SHIFT 采样）
const shift = (x, y) => Math.round((x - 160) * (W / 1024));
const X0 = Math.round(870 * (W / 1024));
const X1 = W; // 到右边缘
const Y0 = Math.round(940 * (H / 1024));
const Y1 = H; // 到下边缘
const SS = W / 1024;

for (let y = Y0; y < Y1; y++) {
  for (let x = X0; x < X1; x++) {
    const sx = x - Math.round(160 * SS); // 左移 160px 采样（干净渐变，无水印）
    if (sx < 0) continue;
    const di = (y * W + x) * 4;
    const si = (y * W + sx) * 4;
    // 保留目标 alpha（圆角边缘），RGB 取自左侧干净区
    data[di] = data[si];
    data[di + 1] = data[si + 1];
    data[di + 2] = data[si + 2];
    // alpha 保持不变
  }
}

/* ---------- 2. 生成各尺寸 PNG ---------- */
const SIZES = [16, 24, 32, 48, 64, 128, 256];
async function renderPng(size) {
  const clone = img.clone();
  clone.resize({ w: size, h: size });
  return await clone.getBuffer("image/png");
}

const entries = [];
for (const s of SIZES) {
  const png = await renderPng(s);
  entries.push({ size: s, data: png });
  console.log(`  ${s}x${s} → ${png.length}B PNG`);
}

/* ---------- 3. 组装 ICO ---------- */
const count = entries.length;
const headerSize = 6 + count * 16;
const total = headerSize + entries.reduce((a, e) => a + e.data.length, 0);
const out = Buffer.alloc(total);
let off = 0;
out.writeUInt16LE(0, off); off += 2; // reserved
out.writeUInt16LE(1, off); off += 2; // type = icon
out.writeUInt16LE(count, off); off += 2; // count
const dataOffsets = [];
for (let i = 0; i < count; i++) {
  const e = entries[i];
  const pngSize = e.data.length;
  dataOffsets.push(headerSize + off - 0); // 先占位，下面重算
  out.writeUInt8(e.size >= 256 ? 0 : e.size, off); off += 1;
  out.writeUInt8(e.size >= 256 ? 0 : e.size, off); off += 1;
  out.writeUInt8(0, off); off += 1; // color count
  out.writeUInt8(0, off); off += 1; // reserved
  out.writeUInt16LE(1, off); off += 2; // planes
  out.writeUInt16LE(32, off); off += 2; // bit count
  out.writeUInt32LE(pngSize, off); off += 4;
  // offset 占位，后面回填
  off += 4;
}
// 回填 offset
let dataOff = headerSize;
for (let i = 0; i < count; i++) {
  out.writeUInt32LE(dataOff, 6 + i * 16 + 12);
  const e = entries[i];
  e.data.copy(out, dataOff);
  dataOff += e.data.length;
}
writeFileSync(OUT, out);
console.log(`\n✅ 已生成 ${OUT}（${out.length}B，${count} 尺寸）`);
