/**
 * 图片格式嗅探（客户端侧，按 base64 首部魔数判断）
 * 独立无依赖模块：word/ppt 执行器与 bridge.js 共用，兼容 IE11 兼容版复用。
 */

/** atob 的安全封装（拿不到时返回空串） */
function decodeHead(b64) {
  try {
    const s = String(b64 || "").slice(0, 64);
    if (typeof atob === "function") return atob(s);
    return "";
  } catch (e) {
    return "";
  }
}

/** 按 base64 前缀识别图片格式，返回小写扩展名（png/jpg/gif/bmp/webp/tif）或 null */
export function imageExtFromBase64(b64) {
  const bin = decodeHead(b64);
  if (bin.length < 12) return null;
  const b = [];
  for (let i = 0; i < 12; i++) b.push(bin.charCodeAt(i));
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "png";
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "jpg";
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return "gif";
  if (b[0] === 0x42 && b[1] === 0x4d) return "bmp";
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return "webp";
  if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00)) return "tif";
  return null;
}
