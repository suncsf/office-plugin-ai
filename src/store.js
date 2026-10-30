import { reactive } from "vue";
import { api } from "./api";
import { DEFAULT_GENERAL_CLIENT } from "./defaults";

export const store = reactive({
  ready: false,
  host: "unknown", // word | excel | ppt | unknown
  activeTab: "chat",

  settings: { models: [], general: { ...DEFAULT_GENERAL_CLIENT } },
  activeModelId: localStorage.getItem("activeModelId") || null,

  conversations: [],
  conversationId: null,
  messages: [],
  ctxUsed: 0,

  running: false,
  live: { content: "", thinking: "" },
  interrupt: null, // { turnId, toolCalls, results: [] }
  pendingOpenId: null, // 历史页请求打开的对话 id

  // 本机 Office 能力自检（ppt 宿主启动时自动检测，随 /api/chat 的 caps 上报）
  caps: { ppt: null },
});

export function activeModel() {
  const list = store.settings.models;
  return list.find((m) => m.id === store.activeModelId) || list[0] || null;
}

export async function loadSettings() {
  const s = await api.getSettings();
  store.settings = {
    models: Array.isArray(s.models) ? s.models : [],
    general: { ...DEFAULT_GENERAL_CLIENT, ...s.general },
  };
  if (store.activeModelId && !store.settings.models.some((m) => m.id === store.activeModelId)) {
    store.activeModelId = null;
  }
  if (!store.activeModelId && store.settings.models.length) {
    store.activeModelId = store.settings.models[0].id;
  }
}

export async function loadConversations() {
  store.conversations = await api.listConversations();
}

export function setActiveModel(id) {
  store.activeModelId = id;
  localStorage.setItem("activeModelId", id || "");
}
