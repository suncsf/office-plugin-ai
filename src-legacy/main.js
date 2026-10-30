import Vue from "vue";
import ElementUI from "element-ui";
import zhCn from "element-ui/lib/locale/lang/zh-CN";
import "element-ui/lib/theme-chalk/index.css";

import App from "./App.vue";
import { initOffice } from "./office";
import { store, loadSettings, loadConversations } from "./store";
import "./styles.css";

Vue.use(ElementUI, { locale: zhCn });
Vue.config.productionTip = false;

(async () => {
  // 1. 初始化 Office.js（浏览器直接打开时降级为 unknown）
  try {
    await initOffice();
  } catch (e) {
    store.host = "unknown";
  }

  new Vue({ render: (h) => h(App) }).$mount("#app");

  // 2. 加载设置与历史对话
  try {
    await loadSettings();
    await loadConversations();
  } catch (e) {
    console.error("初始化加载失败", e);
  }
})();
