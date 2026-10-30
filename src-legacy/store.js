/**
 * 全局状态（IE11 版）：Vue 2.7 的 Vue.observable 等价于现代版的 Vue3 reactive，
 * 字段与 src/store.js 完全一致；localStorage 键名相同（同一后端、同一套设置）。
 */
import Vue from "vue";
import { api } from "./api";
import { DEFAULT_GENERAL_CLIENT } from "../src/defaults";

function readLocal(key) {
  try {
    return localStorage.getItem(key);
  } catch (e) {
    return null;
  }
}

function writeLocal(key, value) {
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    /* 隐私模式等场景忽略 */
  }
}

export const store = Vue.observable({
  ready: false,
  host: "unknown", // word | excel | ppt | unknown
  activeTab: "chat",

  settings: { models: [], general: { ...DEFAULT_GENERAL_CLIENT } },
  activeModelId: readLocal("activeModelId") || null,

  conversations: [],
  conversationId: null,
  messages: [],
  ctxUsed: 0,

  running: false,
  live: { content: "", thinking: "" },
  interrupt: null, // { turnId, toolCalls, results: [] }
  pendingOpenId: null, // 历史页请求打开的对话 id
});

export function activeModel() {
  const list = store.settings.models;
  return (
    list.filter((m) => m.id === store.activeModelId)[0] || list[0] || null
  );
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
  writeLocal("activeModelId", id || "");
}
