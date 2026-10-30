<script>
import { ElMessage, ElMessageBox } from "element-ui";
import { store, loadConversations } from "../store";
import { api } from "../api";
import { formatTime, formatTokens } from "../utils";

export default {
  name: "HistoryViewLegacy",
  data() {
    return {
      store,
      keyword: "",
    };
  },
  computed: {
    list() {
      const k = this.keyword.trim().toLowerCase();
      if (!k) return store.conversations;
      return store.conversations.filter((c) => (c.title || "").toLowerCase().indexOf(k) >= 0);
    },
  },
  methods: {
    msgCount(c) {
      return c.messageCount || 0;
    },
    async remove(conv, e) {
      if (e && e.stopPropagation) e.stopPropagation();
      try {
        await ElMessageBox.confirm("确定删除对话「" + conv.title + "」？该操作不可恢复。", "删除对话", {
          type: "warning",
          confirmButtonText: "删除",
          cancelButtonText: "取消",
        });
      } catch (err) {
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
    },
    async rename(conv, e) {
      if (e && e.stopPropagation) e.stopPropagation();
      try {
        const result = await ElMessageBox.prompt("修改对话标题", "重命名", {
          inputValue: conv.title,
          confirmButtonText: "保存",
          cancelButtonText: "取消",
        });
        const value = result.value;
        if (!value || !String(value).trim()) return;
        await api.renameConversation(conv.id, String(value).trim());
        await loadConversations();
      } catch (err) {
        /* 取消 */
      }
    },
    onOpen(conv) {
      // 通过 store 中转，ChatView 监听后加载
      store.pendingOpenId = conv.id;
      store.activeTab = "chat";
    },
  },
};
</script>

<template>
  <div class="history">
    <div class="toolbar">
      <el-input v-model="keyword" placeholder="搜索对话标题" size="small" clearable prefix-icon="el-icon-search" />
      <el-button size="small" type="primary" icon="el-icon-plus" @click="store.activeTab = 'chat'">新对话</el-button>
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
            {{ formatTime(c.updatedAt) }} · {{ msgCount(c) }} 条消息
            <template v-if="c.ctxUsed"> · 上下文 {{ formatTokens(c.ctxUsed) }}</template>
          </div>
        </div>
        <div class="item-actions">
          <el-button icon="el-icon-edit" type="text" size="small" @click="rename(c, $event)" />
          <el-button icon="el-icon-delete" type="text" size="small" class="danger" @click="remove(c, $event)" />
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
  flex: none;
  padding-bottom: 8px;
  overflow: hidden;
}
.toolbar .el-input {
  float: left;
  width: 65%;
}
.toolbar .el-button {
  float: right;
}
.list {
  flex: 1 1 0;
  height: 0;
}
.item {
  padding: 8px 10px;
  border-radius: 8px;
  cursor: pointer;
  margin-bottom: 2px;
  overflow: hidden;
}
.item:hover {
  background: #eceff3;
}
.item.active {
  background: #e0e9ff;
}
.item-main {
  float: left;
  width: 65%;
}
.item-actions {
  float: right;
  width: 35%;
  text-align: right;
  display: none;
  padding-top: 4px;
}
.item:hover .item-actions {
  display: block;
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
.item-actions .danger {
  color: #f56c6c;
}
</style>
