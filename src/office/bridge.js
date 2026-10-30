/**
 * 工作区文件桥接（任务窗格执行器 ↔ 服务端）
 * ---------------------------------------------------------------
 * 图片等二进制数据在「服务端对话工作区 ↔ 任务窗格」之间经本地 HTTP 直接传输，
 * 不经过 LLM 上下文（工具调用参数/结果只有文件名与元信息）。
 *
 * 传输层用 XMLHttpRequest：word/excel/ppt 执行器被 IE11 兼容版（src-legacy）复用，
 * IE11 无 fetch，XHR 全平台可用。
 */
import { imageExtFromBase64 } from "./sniff.js";

function xhrJson(method, url, body) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(method, url, true);
    if (body !== undefined) xhr.setRequestHeader("Content-Type", "application/json");
    xhr.onload = () => {
      let data = null;
      try {
        data = xhr.responseText ? JSON.parse(xhr.responseText) : null;
      } catch (e) {
        data = null;
      }
      if (xhr.status >= 200 && xhr.status < 300) resolve(data);
      else reject(new Error((data && data.error) || xhr.responseText || "HTTP " + xhr.status));
    };
    xhr.onerror = () => reject(new Error("无法连接 AI 助手服务（工作区文件读写失败，请确认服务已启动）"));
    xhr.send(body !== undefined ? JSON.stringify(body) : undefined);
  });
}

/** 读取对话工作区内的图片文件 → { name, file, ext, size, base64 } */
export function fetchWorkspaceFile(conversationId, relPath) {
  if (!conversationId) {
    return Promise.reject(new Error("无法定位对话工作区（conversationId 为空），请改用 base64 方式提供图片"));
  }
  const qs = "?conv=" + encodeURIComponent(String(conversationId)) + "&path=" + encodeURIComponent(String(relPath));
  return xhrJson("GET", "/api/workspace/file" + qs);
}

/** 保存图片到对话工作区 → { ok, file, ext, size } */
export function saveWorkspaceFile(conversationId, relPath, base64) {
  if (!conversationId) {
    return Promise.reject(new Error("无法定位对话工作区（conversationId 为空），图片未能保存"));
  }
  return xhrJson("POST", "/api/workspace/file", {
    conv: conversationId,
    path: relPath,
    base64: String(base64 || ""),
  });
}

/**
 * 解析插入图片工具的来源参数（base64 / file 二选一，base64 优先）
 * @returns {Promise<string>} 纯 base64（无 dataURL 前缀）
 */
export async function resolveImageSource(args, ctx) {
  const raw = String(args?.base64 || "").trim();
  if (raw) return raw.indexOf("base64,") >= 0 ? raw.slice(raw.indexOf("base64,") + 7) : raw;
  const file = String(args?.file || "").trim();
  if (!file) throw new Error("必须提供 base64 或 file 之一（file 为对话工作区内图片文件的相对路径）");
  const data = await fetchWorkspaceFile(ctx?.conversationId, file);
  if (!data?.base64) throw new Error(`工作区文件读取失败：${file}`);
  return data.base64;
}

export { imageExtFromBase64 };
