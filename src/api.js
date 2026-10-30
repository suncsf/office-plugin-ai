/** 后端 REST + SSE 客户端 */

async function jfetch(url, { method = "GET", body } = {}) {
  const res = await fetch(url, {
    method,
    headers: body !== undefined ? { "Content-Type": "application/json" } : undefined,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = null;
  }
  if (!res.ok) throw new Error(data?.error || text || `HTTP ${res.status}`);
  return data;
}

export const api = {
  health: () => jfetch("/api/health"),
  getSettings: () => jfetch("/api/settings"),
  saveSettings: (settings) => jfetch("/api/settings", { method: "PUT", body: settings }),
  testModel: (model) => jfetch("/api/settings/test-model", { method: "POST", body: model }),

  listConversations: () => jfetch("/api/conversations"),
  createConversation: (title) => jfetch("/api/conversations", { method: "POST", body: { title } }),
  getConversation: (id) => jfetch(`/api/conversations/${id}`),
  renameConversation: (id, title) => jfetch(`/api/conversations/${id}`, { method: "PATCH", body: { title } }),
  deleteConversation: (id) => jfetch(`/api/conversations/${id}`, { method: "DELETE" }),

  listMemories: (host) => jfetch(`/api/memories${host && host !== "all" ? `?host=${host}` : ""}`),
  createMemory: (memory) => jfetch("/api/memories", { method: "POST", body: memory }),
  updateMemory: (id, memory) => jfetch(`/api/memories/${id}`, { method: "PATCH", body: memory }),
  deleteMemory: (id) => jfetch(`/api/memories/${id}`, { method: "DELETE" }),

  resume: (turnId, results) => jfetch("/api/chat/resume", { method: "POST", body: { turnId, results } }),
};

/**
 * POST 并以 SSE 读取事件流
 * @param {string} url
 * @param {object} body
 * @param {(evt: object) => void} onEvent
 * @param {AbortSignal} [signal]
 */
export async function sseStream(url, body, onEvent, signal) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok || !res.body) {
    let msg = `HTTP ${res.status}`;
    try {
      const t = await res.text();
      if (t) msg = t;
    } catch { /* noop */ }
    throw new Error(msg);
  }
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    let idx;
    while ((idx = buf.indexOf("\n\n")) >= 0) {
      const raw = buf.slice(0, idx);
      buf = buf.slice(idx + 2);
      for (const line of raw.split("\n")) {
        if (!line.startsWith("data:")) continue;
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
}
