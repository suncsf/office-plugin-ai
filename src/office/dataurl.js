/**
 * dataURL → 纯 base64 工具（独立文件）
 * 原位于 office/index.js；因 word/excel/ppt 执行器需要被 IE11 兼容版（src-legacy）复用，
 * 而后者不能连带引入 Vue3 版 store，故拆出此无依赖模块。
 */
export function stripDataUrl(b64) {
  const s = String(b64 || "");
  const idx = s.indexOf("base64,");
  return idx >= 0 ? s.slice(idx + 7) : s;
}
