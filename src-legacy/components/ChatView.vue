<script>
import { ElMessage } from "element-ui";
import { store, activeModel, setActiveModel, loadConversations } from "../store";
import { api, sseStream } from "../api";
import { executeTool } from "../office";
import { formatTokens, renderMarkdown } from "../utils";
import MessageBubble from "./MessageBubble.vue";
import ChatInput from "./ChatInput.vue";

export default {
  name: "ChatViewLegacy",
  components: { MessageBubble, ChatInput },
  data() {
    return {
      store,
      stream: null, // { promise, abort }
      interruptState: null, // { turnId, toolCalls, results: [] }
      pinned: true,
    };
  },
  computed: {
    model() {
      return activeModel();
    },
    ctxPct() {
      const m = this.model;
      if (!m || !m.contextTokens) return 0;
      return Math.min(100, Math.round(((store.ctxUsed || 0) / m.contextTokens) * 1000) / 10);
    },
    ctxColor() {
      return this.ctxPct > 90 ? "#ef4444" : this.ctxPct > 70 ? "#f59e0b" : "#22c55e";
    },
  },
  watch: {
    "store.live.thinking"() {
      if (!this.pinned) return;
      const el = this.$refs.liveThinkEl;
      if (el) el.scrollTop = el.scrollHeight;
    },
    "store.messages.length"() {
      this.scrollBottom();
    },
    "store.live.content"() {
      this.scrollBottom();
    },
    "store.pendingOpenId"(id) {
      if (id && store.activeTab === "chat" && !store.running) {
        store.pendingOpenId = null;
        this.openConversation(id);
      }
    },
  },
  methods: {
    renderLive(t) {
      return renderMarkdown(t);
    },

    /* ---------------- 滚动（智能跟随） ---------------- */
    nearBottom() {
      const el = this.$refs.msgsEl;
      if (!el) return true;
      // 阈值放宽到 150px：流式渲染时高度先增长后滚动，小阈值会被大增量误判
      return el.scrollHeight - el.scrollTop - el.clientHeight < 150;
    },
    onScroll() {
      this.pinned = this.nearBottom();
    },
    scrollBottom(force) {
      if (force) this.pinned = true;
      if (!this.pinned) return;
      this.$nextTick(() => {
        const el = this.$refs.msgsEl;
        if (el) el.scrollTop = el.scrollHeight;
        const think = this.$refs.liveThinkEl;
        if (think) think.scrollTop = think.scrollHeight;
      });
    },

    /* ---------------- 对话管理 ---------------- */
    newChat() {
      if (store.running) {
        ElMessage.warning("当前正在生成回复");
        return;
      }
      store.conversationId = null;
      store.messages = [];
      store.ctxUsed = 0;
      this.interruptState = null;
      this.pinned = true;
    },

    async openConversation(id) {
      if (store.running) {
        ElMessage.warning("当前正在生成回复");
        return;
      }
      try {
        const data = await api.getConversation(id);
        store.conversationId = data.conversation.id;
        store.messages = data.messages;
        store.ctxUsed = data.conversation.ctxUsed || 0;
        this.restoreToolStatuses();
        store.activeTab = "chat";
        this.scrollBottom(true);
      } catch (e) {
        ElMessage.error(e.message);
      }
    },

    restoreToolStatuses() {
      const answered = {};
      store.messages.forEach((m) => {
        if (m.role === "tool" && m.toolCallId) answered[m.toolCallId] = m;
      });
      store.messages.forEach((m) => {
        const tcs = m.toolCalls || [];
        tcs.forEach((tc) => {
          const r = answered[tc.id];
          tc.status = r ? ((r.meta && r.meta.isError) ? "error" : "ok") : "interrupted";
        });
      });
    },

    findToolCall(id) {
      for (let i = store.messages.length - 1; i >= 0; i--) {
        const m = store.messages[i];
        const tcs = (m.role === "assistant" && m.toolCalls) || [];
        for (let j = 0; j < tcs.length; j++) {
          if (tcs[j].id === id) return tcs[j];
        }
      }
      return null;
    },

    /* ---------------- 发送 & SSE ---------------- */
    async onSend(payload) {
      const content = payload.content;
      const images = payload.images || [];
      const attachments = payload.attachments || [];
      if (store.running) return;
      if (!this.model) {
        ElMessage.warning("请先在「设置」中添加模型");
        store.activeTab = "settings";
        return;
      }
      store.running = true;
      store.live = { content: "", thinking: "" };
      const stream = sseStream(
        "/api/chat",
        {
          conversationId: store.conversationId,
          content,
          images,
          attachments,
          appType: store.host,
          modelId: this.model.id,
          // IE11 兼容版标识：服务端据此剔除依赖 async/await 的 run_code 工具
          caps: { legacy: true },
        },
        this.handleEvent
      );
      this.stream = stream;
      try {
        await stream.promise;
      } catch (e) {
        if (!e || e.name !== "AbortError") {
          store.messages.push({
            id: "err_" + Date.now(),
            role: "assistant",
            content: "⚠️ 请求失败：" + e.message,
            toolCalls: [],
            thinking: "",
            images: [],
            meta: { error: true },
          });
        }
      } finally {
        store.running = false;
        store.live = { content: "", thinking: "" };
        this.interruptState = null;
        this.stream = null;
        loadConversations().catch(() => {});
      }
    },

    onStop() {
      if (this.stream) this.stream.abort();
    },

    handleEvent(evt) {
      switch (evt.type) {
        case "start": {
          store.conversationId = evt.conversationId;
          if (evt.userMessage) store.messages.push(evt.userMessage);
          break;
        }
        case "content_delta":
          store.live.content += evt.text;
          break;
        case "thinking_delta":
          store.live.thinking += evt.text;
          break;
        case "message": {
          store.messages.push(evt.message);
          store.live = { content: "", thinking: "" };
          if (evt.message.role === "tool" && evt.message.toolCallId) {
            const tc = this.findToolCall(evt.message.toolCallId);
            if (tc) tc.status = (evt.message.meta && evt.message.meta.isError) ? "error" : "ok";
          }
          this.scrollBottom();
          break;
        }
        case "usage":
          if (evt.usage && evt.usage.input) store.ctxUsed = Math.max(store.ctxUsed || 0, evt.usage.input);
          break;
        case "interrupt": {
          const toolCalls = evt.toolCalls || [];
          this.interruptState = {
            turnId: evt.turnId,
            toolCalls,
            results: toolCalls.map(() => null),
          };
          const self = this;
          toolCalls.forEach((tc) => {
            const t = self.findToolCall(tc.id);
            if (t) t.status = "pending";
          });
          this.autoRunInterrupt();
          break;
        }
        case "done":
          if (evt.ctxUsed != null) store.ctxUsed = evt.ctxUsed;
          break;
        case "error":
          store.live = { content: "", thinking: "" };
          store.messages.push({
            id: "err_" + Date.now(),
            role: "assistant",
            content: "⚠️ " + evt.message,
            toolCalls: [],
            thinking: "",
            images: [],
            meta: { error: true },
          });
          this.scrollBottom();
          break;
      }
    },

    /* ---------------- 工具中断：自动执行 + 表单提问 ---------------- */
    fillResult(toolCallId, result) {
      const st = this.interruptState;
      if (!st) return;
      const idx = st.toolCalls.findIndex((t) => t.id === toolCallId);
      if (idx < 0) return;
      st.results[idx] = result;
      this.trySubmitResume();
    },

    async autoRunInterrupt() {
      const st = this.interruptState;
      if (!st) return;
      for (const tc of st.toolCalls) {
        if (tc.name === "ask_user") continue; // 等用户作答
        const t = this.findToolCall(tc.id);
        if (t) t.status = "running";
        const result = await executeTool(tc.name, tc.args);
        if (t) t.status = result.isError ? "error" : "ok";
        this.fillResult(tc.id, result);
      }
    },

    onAskAnswer(payload) {
      const t = this.findToolCall(payload.toolCallId);
      if (t) t.status = "running";
      this.fillResult(payload.toolCallId, { content: payload.answer, isError: false });
    },

    trySubmitResume() {
      const st = this.interruptState;
      if (!st || st.results.some((r) => r === null)) return;
      this.interruptState = null;
      const results = st.toolCalls.map((tc, i) => ({
        toolCallId: tc.id,
        name: tc.name,
        content: (st.results[i] && st.results[i].content) || "",
        isError: !!(st.results[i] && st.results[i].isError),
      }));
      api.resume(st.turnId, results).catch((e) => {
        // 会话已失效（服务重启等）
        for (const r of results) {
          const t = this.findToolCall(r.toolCallId);
          if (t && (t.status === "pending" || t.status === "running")) t.status = "interrupted";
        }
        ElMessage.error(e.message);
      });
    },
  },
};
</script>

<template>
  <div class="chat">
    <!-- 模型栏 -->
    <div class="chat-head">
      <el-select
        :value="model ? model.id : null"
        size="small"
        placeholder="选择模型"
        class="model-select"
        @input="setActiveModel"
      >
        <el-option v-for="m in store.settings.models" :key="m.id" :value="m.id" :label="m.name">
          <div class="model-option">
            <span class="mo-name">{{ m.name }}</span>
            <el-tag v-if="m.supportsTools !== false" size="small" type="success" effect="plain">工具</el-tag>
            <el-tag v-if="m.supportsImages" size="small" type="primary" effect="plain">图片</el-tag>
            <el-tag size="small" type="info" effect="plain">{{ formatTokens(m.contextTokens) }}</el-tag>
          </div>
        </el-option>
      </el-select>
      <el-button size="small" @click="newChat">＋ 新对话</el-button>
    </div>

    <!-- 上下文额度 -->
    <div v-if="model" class="ctx-bar">
      <el-progress
        :percentage="ctxPct"
        :color="ctxColor"
        :stroke-width="5"
        :show-text="false"
        class="ctx-progress"
      />
      <span class="ctx-text" :style="{ color: ctxColor }">
        上下文 {{ formatTokens(store.ctxUsed) }} / {{ formatTokens(model.contextTokens) }}
      </span>
    </div>

    <!-- 消息列表 -->
    <div ref="msgsEl" class="msgs" @scroll="onScroll">
      <el-empty
        v-if="!store.messages.length && !store.running"
        description="向 AI 描述你想对文档做的事"
        :image-size="80"
      />
      <MessageBubble
        v-for="m in store.messages"
        :key="m.id"
        :message="m"
        @ask-answer="onAskAnswer"
      />

      <!-- 流式输出中 -->
      <div v-if="store.running" class="live">
        <div v-if="store.live.thinking" class="thinking-block">
          <div class="thinking-title">💭 思考中…</div>
          <div ref="liveThinkEl" class="thinking-content md-body" v-html="renderLive(store.live.thinking)"></div>
        </div>
        <div v-if="store.live.content" class="bubble assistant live-bubble">
          <div class="md-body" v-html="renderLive(store.live.content)"></div>
        </div>
        <div v-if="!store.live.content && !store.live.thinking" class="live-hint">
          <span class="dot-spin"></span> AI 正在处理…
        </div>
      </div>
    </div>

    <!-- 已离开底部：一键回最新（悬浮，不随内容滚动） -->
    <div v-if="!pinned" class="jump-latest" @click="scrollBottom(true)">↓ 回到最新</div>

    <ChatInput :disabled="store.running" @send="onSend" @stop="onStop" />
  </div>
</template>

<style scoped>
.chat {
  height: 100%;
  display: flex;
  flex-direction: column;
  position: relative;
}
.jump-latest {
  position: absolute;
  right: 18px;
  bottom: 128px;
  z-index: 5;
  background: #fff;
  border: 1px solid #c7d6ea;
  border-radius: 14px;
  padding: 3px 12px;
  font-size: 12px;
  color: #2563eb;
  cursor: pointer;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}
.jump-latest:hover {
  background: #eff6ff;
}
.chat-head {
  display: flex;
  align-items: center;
  padding: 8px 0 4px;
  flex: none;
}
.model-select {
  flex: 1;
  margin-right: 6px;
}
.model-option:after {
  content: "";
  display: block;
  clear: both;
}
.mo-name {
  float: left;
  margin-right: 6px;
}
.model-option .el-tag {
  float: left;
  margin-right: 4px;
}
.ctx-bar {
  display: flex;
  align-items: center;
  flex: none;
  padding: 0 2px 6px;
}
.ctx-progress {
  flex: 1;
}
.ctx-text {
  font-size: 11px;
  white-space: nowrap;
  margin-left: 8px;
}
.msgs {
  flex: 1 1 0;
  height: 0;
  overflow-y: auto;
  padding: 4px 2px 10px;
}
.live {
  margin: 6px 0;
}
.live-bubble {
  background: #fff;
}
.live-hint {
  color: #909399;
  font-size: 12px;
  padding: 8px 4px;
}
.dot-spin {
  display: inline-block;
  width: 10px;
  height: 10px;
  border: 2px solid #cbd5e1;
  border-top-color: #2563eb;
  border-radius: 50%;
  margin-right: 6px;
  animation: spin 0.8s linear infinite;
}
</style>
