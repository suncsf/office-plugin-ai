<script setup>
import { ref, computed } from "vue";
import { ElMessage, ElMessageBox } from "element-plus";
import { Search, Plus, Delete, Edit } from "@element-plus/icons-vue";
import { store, loadConversations } from "../store";
import { api } from "../api";
import { formatTime, formatTokens } from "../utils";

const keyword = ref("");

const list = computed(() => {
  const k = keyword.value.trim().toLowerCase();
  if (!k) return store.conversations;
  return store.conversations.filter((c) => (c.title || "").toLowerCase().includes(k));
});

async function remove(conv, e) {
  e.stopPropagation();
  try {
    await ElMessageBox.confirm(`确定删除对话「${conv.title}」？该操作不可恢复。`, "删除对话", {
      type: "warning",
      confirmButtonText: "删除",
      cancelButtonText: "取消",
    });
  } catch {
    return;
  }
  try {
    await api.deleteConversation(conv.id);
    if (store.conversationId === conv.id) {
      store.conversationId = null;
      store.messages = [];
    }
    await loadConversations();
    ElMessage.success("已删除");
  } catch (err) {
    ElMessage.error(err.message);
  }
}

async function rename(conv, e) {
  e.stopPropagation();
  try {
    const { value } = await ElMessageBox.prompt("修改对话标题", "重命名", {
      inputValue: conv.title,
      confirmButtonText: "保存",
      cancelButtonText: "取消",
    });
    if (!value?.trim()) return;
    await api.renameConversation(conv.id, value.trim());
    await loadConversations();
  } catch { /* 取消 */ }
}

function onOpen(conv) {
  // 通过 store 中转，ChatView 监听后加载
  store.pendingOpenId = conv.id;
  store.activeTab = "chat";
}
</script>

<template>
  <div class="history">
    <div class="toolbar">
      <el-input v-model="keyword" placeholder="搜索对话标题" size="small" clearable :prefix-icon="Search" />
      <el-button size="small" type="primary" :icon="Plus" @click="store.activeTab = 'chat'">新对话</el-button>
    </div>

    <el-scrollbar class="list">
      <el-empty v-if="!list.length" description="暂无历史对话" :image-size="80" />
      <div
        v-for="c in list"
        :key="c.id"
        class="item"
        :class="{ active: c.id === store.conversationId }"
        @click="onOpen(c)"
      >
        <div class="item-main">
          <div class="item-title">{{ c.title }}</div>
          <div class="item-meta">
            {{ formatTime(c.updatedAt) }} · {{ c.messageCount ?? 0 }} 条消息
            <template v-if="c.ctxUsed"> · 上下文 {{ formatTokens(c.ctxUsed) }}</template>
          </div>
        </div>
        <div class="item-actions">
          <el-button :icon="Edit" text size="small" @click="rename(c, $event)" />
          <el-button :icon="Delete" text size="small" type="danger" @click="remove(c, $event)" />
        </div>
      </div>
    </el-scrollbar>
  </div>
</template>

<style scoped>
.history {
  height: 100%;
  display: flex;
  flex-direction: column;
  padding-top: 8px;
}
.toolbar {
  display: flex;
  gap: 6px;
  flex: none;
  padding-bottom: 8px;
}
.toolbar .el-input {
  flex: 1;
}
.list {
  flex: 1;
  min-height: 0;
}
.item {
  display: flex;
  align-items: center;
  padding: 8px 10px;
  border-radius: 8px;
  cursor: pointer;
  margin-bottom: 2px;
}
.item:hover {
  background: #eceff3;
}
.item.active {
  background: #e0e9ff;
}
.item-main {
  flex: 1;
  min-width: 0;
}
.item-title {
  font-weight: 600;
  font-size: 13px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.item-meta {
  font-size: 11px;
  color: #909399;
  margin-top: 2px;
}
.item-actions {
  flex: none;
  display: none;
}
.item:hover .item-actions {
  display: flex;
}
</style>
