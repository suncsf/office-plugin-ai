<script setup>
import { computed } from "vue";
import { store, activeModel } from "./store";
import { HOST_LABELS } from "./defaults";
import ChatView from "./components/ChatView.vue";
import HistoryView from "./components/HistoryView.vue";
import SettingsView from "./components/SettingsView.vue";
import AboutView from "./components/AboutView.vue";

const hostLabel = computed(() => HOST_LABELS[store.host] || HOST_LABELS.unknown);
const modelOk = computed(() => store.settings.models.length > 0);
</script>

<template>
  <div class="app">
    <header class="app-header">
      <div class="brand">
        <span class="logo">✦</span>
        <span class="title">Office AI 助手</span>
      </div>
      <el-tag :type="store.host === 'unknown' ? 'info' : 'success'" size="small" effect="plain">
        {{ hostLabel }}
      </el-tag>
    </header>

    <el-alert
      v-if="store.host === 'unknown'"
      type="warning"
      :closable="false"
      show-icon
      class="host-alert"
      title="未检测到 Office 宿主"
      description="请在 Word / Excel / PowerPoint 的「加载项」中打开本面板；此时仍可使用对话与设置，文档工具不可用。"
    />
    <el-alert
      v-else-if="!modelOk"
      type="info"
      :closable="false"
      show-icon
      class="host-alert"
      title="开始前请先配置 AI 模型"
      description="到「设置 → 模型管理」添加一个 OpenAI 兼容接口（url + key + model）。"
    />

    <main class="app-body">
      <el-tabs v-model="store.activeTab" class="app-tabs">
        <el-tab-pane label="对话" name="chat">
          <ChatView />
        </el-tab-pane>
        <el-tab-pane label="历史" name="history">
          <HistoryView />
        </el-tab-pane>
        <el-tab-pane label="设置" name="settings">
          <SettingsView />
        </el-tab-pane>
        <el-tab-pane label="关于" name="about">
          <AboutView />
        </el-tab-pane>
      </el-tabs>
    </main>
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  background: #f5f7fa;
}
.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 12px;
  background: #fff;
  border-bottom: 1px solid #e4e7ed;
  flex: none;
}
.brand {
  display: flex;
  align-items: center;
  gap: 6px;
  font-weight: 600;
  font-size: 14px;
  color: #303133;
}
.logo {
  color: #2563eb;
  font-size: 16px;
}
.host-alert {
  margin: 6px 8px 0;
  flex: none;
}
.app-body {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
}
.app-tabs {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
  padding: 0 8px;
}
.app-tabs :deep(.el-tabs__content) {
  flex: 1;
  min-height: 0;
  overflow: hidden;
}
.app-tabs :deep(.el-tab-pane) {
  height: 100%;
}
</style>
