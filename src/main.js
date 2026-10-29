import { createApp } from "vue";
import ElementPlus from "element-plus";
import zhCn from "element-plus/dist/locale/zh-cn.mjs";
import "element-plus/dist/index.css";

import App from "./App.vue";
import { initOffice } from "./office";
import { store, loadSettings, loadConversations } from "./store";
import { api } from "./api";
import "./styles.css";

(async () => {
  // 1. 初始化 Office.js（浏览器直接打开时降级为 unknown）
  try {
    await initOffice();
  } catch (e) {
    store.host = "unknown";
  }

  const app = createApp(App);
  app.use(ElementPlus, { locale: zhCn });
  app.mount("#app");

  // 2. 加载设置与历史对话
  try {
    await loadSettings();
    await loadConversations();
  } catch (e) {
    console.error("初始化加载失败", e);
  }

  // 浏览器直接打开（非 Office 宿主）时给出提示
  if (store.host === "unknown") {
    api.health().catch(() => {});
  }
})();
