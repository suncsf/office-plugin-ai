/**
 * 生成加载项图标（public/assets/icon-{16,32,64,80}.png）
 * 纯 Node 实现 PNG 编码，无第三方依赖。
 */
import { deflateSync } from "node:zlib";
import { writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const outDir = join(dirname(fileURLToPath(import.meta.url)), "..", "public", "assets");
mkdirSync(outDir, { recursive: true });

const CRC_TABLE = (() => {
  const t = new Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const t = Buffer.from(type, "ascii");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([t, data])));
  return Buffer.concat([len, t, data, crc]);
}

function makePng(size) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // RGBA
  ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;

  const raw = Buffer.alloc(size * (size * 4 + 1));
  const c1 = [37, 99, 235];    // 蓝
  const c2 = [16, 185, 129];   // 绿
  const margin = Math.max(1, Math.round(size * 0.04));
  const s = size - margin * 2;
  const radius = Math.round(s * 0.24);
  const cx = size / 2, cy = size / 2;
  const arm = s * 0.32;
  const thick = Math.max(2, s * 0.16);

  const insideRounded = (x, y) => {
    const x0 = x - margin, y0 = y - margin;
    if (x0 < 0 || y0 < 0 || x0 >= s || y0 >= s) return false;
    const nx = Math.min(x0, s - 1 - x0);
    const ny = Math.min(y0, s - 1 - y0);
    if (nx >= radius || ny >= radius) return true;
    const ccx = x0 < radius ? radius : s - 1 - radius;
    const ccy = y0 < radius ? radius : s - 1 - radius;
    return (x0 - ccx) ** 2 + (y0 - ccy) ** 2 <= radius * radius;
  };
  const insidePlus = (x, y) => {
    const dx = Math.abs(x - cx), dy = Math.abs(y - cy);
    return (dx <= thick / 2 && dy <= arm) || (dy <= thick / 2 && dx <= arm);
  };

  let p = 0;
  for (let y = 0; y < size; y++) {
    raw[p++] = 0; // filter: none
    const t = y / (size - 1);
    for (let x = 0; x < size; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      if (insidePlus(x, y) && insideRounded(x, y)) {
        r = 255; g = 255; b = 255; a = 255;
      } else if (insideRounded(x, y)) {
        r = Math.round(c1[0] + (c2[0] - c1[0]) * t);
        g = Math.round(c1[1] + (c2[1] - c1[1]) * t);
        b = Math.round(c1[2] + (c2[2] - c1[2]) * t);
        a = 255;
      }
      raw[p++] = r; raw[p++] = g; raw[p++] = b; raw[p++] = a;
    }
  }

  const ihdrChunk = chunk("IHDR", ihdr);
  const idat = chunk("IDAT", deflateSync(raw, { level: 9 }));
  const iend = chunk("IEND", Buffer.alloc(0));
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    ihdrChunk, idat, iend,
  ]);
}

for (const size of [16, 32, 64, 80]) {
  writeFileSync(join(outDir, `icon-${size}.png`), makePng(size));
}
console.log(`[icons] 已生成: ${outDir}`);
