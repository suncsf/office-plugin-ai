import MarkdownIt from "markdown-it";

const md = new MarkdownIt({
  html: false,
  linkify: true,
  breaks: true,
});

export function renderMarkdown(text) {
  return md.render(String(text || ""));
}

export function formatTokens(n) {
  n = Number(n) || 0;
  if (n >= 1000000) return (n / 1000000).toFixed(1) + "M";
  if (n >= 1000) return (n / 1000).toFixed(1) + "k";
  return String(n);
}

export function formatTime(ts) {
  if (!ts) return "";
  const d = new Date(ts);
  const now = new Date();
  const sameDay = d.toDateString() === now.toDateString();
  const hm = `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  if (sameDay) return hm;
  return `${d.getMonth() + 1}/${d.getDate()} ${hm}`;
}

export function truncate(text, max) {
  const s = String(text ?? "");
  if (s.length <= max) return s;
  return s.slice(0, max) + `…[已截断，原始长度 ${s.length}]`;
}

// 检测任务窗格当前运行环境（WebView / 浏览器内核与版本），用于「关于」卡片显示
export function getRuntimeInfo() {
  const ua = navigator.userAgent || "";
  // IE11 兼容内核（Office 2019 走 src-legacy 构建时的承载环境）
  if (/MSIE |Trident\//.test(ua)) return "IE11 兼容内核（Trident）";
  const rules = [
    ["Edge", /Edg(?:e|A|iOS)?\/([\d.]+)/],
    ["Opera", /(?:OPR|Opera)\/([\d.]+)/],
    ["Chrome", /(?:Chrome|CriOS)\/([\d.]+)/],
    ["Firefox", /(?:Firefox|FxiOS)\/([\d.]+)/],
    ["Safari", /Version\/([\d.]+)[\s\S]*Safari/],
  ];
  let name = "未知浏览器";
  let version = "";
  for (const [n, re] of rules) {
    const m = ua.match(re);
    if (m) {
      name = n;
      version = m[1];
      break;
    }
  }
  // Office 桌面端任务窗格由 WebView2（Edge Chromium 内核）承载
  let host = "";
  try {
    const diag = window.Office && Office.context && Office.context.diagnostics;
    if (diag && diag.host) host = diag.host;
  } catch (e) { /* Office.js 未加载（纯浏览器开发模式） */ }
  const core = version ? `${name} ${version}` : name;
  return host ? `${host}（WebView2 · ${core}）` : `${core}（普通浏览器）`;
}
