/**
 * 后端 REST + SSE 客户端（IE11 版）
 * 与现代版 src/api.js 的接口保持一致，但传输层换成 XMLHttpRequest：
 * IE11 无 fetch / ReadableStream / TextDecoder / AbortController。
 * SSE 通过 xhr.onprogress 增量读取 responseText 实现，协议解析与现代版相同。
 */

function request(method, url, body) {
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
      if (xhr.status >= 200 && xhr.status < 300) {
        resolve(data);
      } else {
        const msg = (data && data.error) || xhr.responseText || "HTTP " + xhr.status;
        reject(new Error(msg));
      }
    };
    xhr.onerror = () => reject(new Error("网络错误（请确认 AI 助手服务已启动、证书已安装）"));
    xhr.send(body !== undefined ? JSON.stringify(body) : undefined);
  });
}

export const api = {
  health: () => request("GET", "/api/health"),
  getSettings: () => request("GET", "/api/settings"),
  saveSettings: (settings) => request("PUT", "/api/settings", settings),
  testModel: (model) => request("POST", "/api/settings/test-model", model),

  listConversations: () => request("GET", "/api/conversations"),
  createConversation: (title) => request("POST", "/api/conversations", { title }),
  getConversation: (id) => request("GET", "/api/conversations/" + id),
  renameConversation: (id, title) => request("PATCH", "/api/conversations/" + id, { title }),
  deleteConversation: (id) => request("DELETE", "/api/conversations/" + id),

  resume: (turnId, results) => request("POST", "/api/chat/resume", { turnId, results }),
};

/**
 * POST 并以 SSE 读取事件流（IE11：XHR 增量读取）
 * @param {string} url
 * @param {object} body
 * @param {(evt: object) => void} onEvent
 * @returns {{ promise: Promise<void>, abort: () => void }}
 */
export function sseStream(url, body, onEvent) {
  const xhr = new XMLHttpRequest();
  let seen = 0;
  let buf = "";

  function feed(chunk) {
    buf += chunk;
    let idx;
    while ((idx = buf.indexOf("\n\n")) >= 0) {
      const raw = buf.slice(0, idx);
      buf = buf.slice(idx + 2);
      const lines = raw.split("\n");
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.indexOf("data:") !== 0) continue;
        const data = line.slice(5).trim();
        if (!data) continue;
        try {
          onEvent(JSON.parse(data));
        } catch (e) {
          console.warn("SSE 解析失败", data, e);
        }
      }
    }
  }

  const promise = new Promise((resolve, reject) => {
    xhr.open("POST", url, true);
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.onprogress = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        const text = xhr.responseText;
        if (text.length > seen) {
          feed(text.slice(seen));
          seen = text.length;
        }
      }
    };
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        const text = xhr.responseText;
        if (text.length > seen) {
          feed(text.slice(seen));
          seen = text.length;
        }
        resolve();
      } else {
        let msg = "HTTP " + xhr.status;
        const t = xhr.responseText;
        if (t) {
          try {
            const j = JSON.parse(t);
            msg = (j && j.error) || t;
          } catch (e) {
            msg = t;
          }
        }
        reject(new Error(msg));
      }
    };
    xhr.onerror = () => reject(new Error("网络错误（请确认 AI 助手服务已启动、证书已安装）"));
    xhr.onabort = () => {
      const err = new Error("已停止");
      err.name = "AbortError";
      reject(err);
    };
    xhr.send(JSON.stringify(body));
  });

  return {
    promise: promise,
    abort: () => xhr.abort(),
  };
}
