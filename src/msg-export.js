/* 消息导出辅助：复制文本 / 复制为图片 / 导出 PDF（仅现代版，legacy IE11 不引入） */
import { toBlob, toCanvas } from "html-to-image";
import { jsPDF } from "jspdf";

// Chromium 画布单边上限 16384px，留余量；内容过高时降低像素比防止导出失败
const MAX_CAPTURE_PX = 16000;

function stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

function captureRatio(el) {
  return Math.min(2, Math.max(1, MAX_CAPTURE_PX / el.scrollHeight));
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/* 复制纯文本；navigator.clipboard 不可用或被拒时走 execCommand 后备 */
export async function copyText(text) {
  const s = String(text || "");
  if (!s) return false;
  try {
    await navigator.clipboard.writeText(s);
    return true;
  } catch (e) {
    const ta = document.createElement("textarea");
    ta.value = s;
    ta.style.cssText = "position:fixed;top:-9999px;opacity:0";
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (e2) {
      ok = false;
    }
    document.body.removeChild(ta);
    return ok;
  }
}

/* 把元素渲染成 PNG 写入剪贴板；剪贴板不可用时降级为下载图片文件 */
export async function copyElementAsImage(el) {
  const blob = await toBlob(el, { pixelRatio: captureRatio(el), backgroundColor: "#fff" });
  if (!blob) throw new Error("渲染失败");
  if (window.ClipboardItem && navigator.clipboard && navigator.clipboard.write) {
    try {
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      return "clipboard";
    } catch (e) {
      /* 剪贴板被拒 → 落到下载 */
    }
  }
  downloadBlob(blob, `AI回答_${stamp()}.png`);
  return "download";
}

/* 把元素渲染成分页 A4 PDF（图片版式：内容按宽度缩放，超高自动分页） */
export async function exportElementToPdf(el) {
  const canvas = await toCanvas(el, { pixelRatio: captureRatio(el), backgroundColor: "#fff" });
  const imgW = canvas.width;
  const imgH = canvas.height;
  if (!imgW || !imgH) throw new Error("渲染失败");
  const MARGIN = 10; // 页边距 mm
  const contentW = 210 - MARGIN * 2; // A4 内容宽 190mm
  const contentH = 297 - MARGIN * 2; // A4 内容高 277mm
  const pxPerMm = imgW / contentW;
  const pageHpx = Math.floor(contentH * pxPerMm);
  const pages = Math.max(1, Math.ceil(imgH / pageHpx));
  const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait", compress: true });
  for (let i = 0; i < pages; i++) {
    const h = Math.min(pageHpx, imgH - i * pageHpx);
    const page = document.createElement("canvas");
    page.width = imgW;
    page.height = h;
    const ctx = page.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, imgW, h);
    ctx.drawImage(canvas, 0, i * pageHpx, imgW, h, 0, 0, imgW, h);
    if (i > 0) pdf.addPage();
    pdf.addImage(page.toDataURL("image/jpeg", 0.92), "JPEG", MARGIN, MARGIN, contentW, h / pxPerMm, undefined, "FAST");
  }
  pdf.save(`AI回答_${stamp()}.pdf`);
  return pages;
}
