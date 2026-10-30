<script setup>
import { ref, computed, watch, nextTick } from "vue";
import { ElMessage } from "element-plus";
import { store, activeModel, setActiveModel, loadConversations } from "../store";
import { api, sseStream } from "../api";
import { executeTool } from "../office";
import { formatTokens, renderMarkdown } from "../utils";
import MessageBubble from "./MessageBubble.vue";
import ChatInput from "./ChatInput.vue";

const msgsEl = ref(null);
const controller = ref(null);
const interruptState = ref(null); // { turnId, toolCalls, results: [] }

const model = computed(() => activeModel());
const ctxPct = computed(() => {
  const m = model.value;
  if (!m?.contextTokens) return 0;
  return Math.min(100, Math.round(((store.ctxUsed || 0) / m.contextTokens) * 1000) / 10);
});
const ctxColor = computed(() => (ctxPct.value > 90 ? "#ef4444" : ctxPct.value > 70 ? "#f59e0b" : "#22c55e"));

function renderLive(t) {
  return renderMarkdown(t);
}

/* ---------------- 滚动（智能跟随） ----------------
 * pinned = 用户"钉住"底部：跟随思考/输出自动滚动；
 * 用户向上翻阅 → 取消钉住，自由浏览历史；
 * 用户拉回最底部 → 恢复钉住，继续自动滚动到最新。
 */
const pinned = ref(true);
function nearBottom() {
  const el = msgsEl.value;
  if (!el) return true;
  // 阈值放宽到 150px：流式渲染时高度先增长后滚动，小阈值会被大增量误判为"已离开底部"
  return el.scrollHeight - el.scrollTop - el.clientHeight < 150;
}
function onScroll() {
  pinned.value = nearBottom();
}
function scrollBottom(force = false) {
  if (force) pinned.value = true;
  if (!pinned.value) return;
  nextTick(() => {
    const el = msgsEl.value;
    if (el) el.scrollTop = el.scrollHeight;
    const think = liveThinkEl.value;
    if (think) think.scrollTop = think.scrollHeight; // 思考框内部也跟随
  });
}
const liveThinkEl = ref(null); // 流式思考框内部跟随
watch(
  () => store.live.thinking,
  () => {
    if (!pinned.value) return;
    nextTick(() => {
      const el = liveThinkEl.value;
      if (el) el.scrollTop = el.scrollHeight;
    });
  }
);
watch(() => [store.messages.length, store.live.content, store.live.thinking], () => scrollBottom());

/* ---------------- 对话管理 ---------------- */
function newChat() {
  if (store.running) {
    ElMessage.warning("当前正在生成回复");
    return;
  }
  store.conversationId = null;
  store.messages = [];
  store.ctxUsed = 0;
  interruptState.value = null;
  pinned.value = true;
}

// 历史页请求打开某个对话
watch(
  () => [store.pendingOpenId, store.activeTab],
  async ([id, tab]) => {
    if (id && tab === "chat" && !store.running) {
      store.pendingOpenId = null;
      await openConversation(id);
    }
  }
);

async function openConversation(id) {
  if (store.running) {
    ElMessage.warning("当前正在生成回复");
    return;
  }
  try {
    const data = await api.getConversation(id);
    store.conversationId = data.conversation.id;
    store.messages = data.messages;
    store.ctxUsed = data.conversation.ctxUsed || 0;
    restoreToolStatuses();
    store.activeTab = "chat";
    scrollBottom(true);
  } catch (e) {
    ElMessage.error(e.message);
  }
}

function restoreToolStatuses() {
  const answered = new Map();
  for (const m of store.messages) {
    if (m.role === "tool" && m.toolCallId) answered.set(m.toolCallId, m);
  }
  for (const m of store.messages) {
    for (const tc of m.toolCalls || []) {
      const r = answered.get(tc.id);
      tc.status = r ? (r.meta?.isError ? "error" : "ok") : "interrupted";
    }
  }
}

function findToolCall(id) {
  for (let i = store.messages.length - 1; i >= 0; i--) {
    const m = store.messages[i];
    if (m.role === "assistant" && (m.toolCalls || []).length) {
      const tc = m.toolCalls.find((t) => t.id === id);
      if (tc) return tc;
    }
  }
  return null;
}

/* ---------------- 发送 & SSE ---------------- */
async function onSend({ content, images, attachments = [] }) {
  if (store.running) return;
  if (!model.value) {
    ElMessage.warning("请先在「设置」中添加模型");
    store.activeTab = "settings";
    return;
  }
  store.running = true;
  store.live = { content: "", thinking: "" };
  controller.value = new AbortController();
  try {
    await sseStream(
      "/api/chat",
      {
        conversationId: store.conversationId,
        content,
        images,
        attachments,
        appType: store.host,
        modelId: model.value.id,
        // 客户端能力声明：legacy=IE11 兼容版；ppt=PPT 宿主启动时检测的 PowerPointApi 能力
        // 服务端据此剔除本机不可用的 ppt 工具，并在系统提示词注入能力摘要
        caps: { legacy: false, ppt: store.caps?.ppt || null },
      },
      handleEvent,
      controller.value.signal
    );
  } catch (e) {
    if (e?.name !== "AbortError") {
      store.messages.push({
        id: `err_${Date.now()}`,
        role: "assistant",
        content: `⚠️ 请求失败：${e.message}`,
        toolCalls: [],
        thinking: "",
        images: [],
        meta: { error: true },
      });
    }
  } finally {
    store.running = false;
    store.live = { content: "", thinking: "" };
    interruptState.value = null;
    loadConversations().catch(() => {});
  }
}

function onStop() {
  controller.value?.abort();
}

function handleEvent(evt) {
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
      // 工具结果消息 → 更新对应调用状态
      if (evt.message.role === "tool" && evt.message.toolCallId) {
        const tc = findToolCall(evt.message.toolCallId);
        if (tc) tc.status = evt.message.meta?.isError ? "error" : "ok";
      }
      scrollBottom();
      break;
    }
    case "usage":
      if (evt.usage?.input) store.ctxUsed = Math.max(store.ctxUsed || 0, evt.usage.input);
      break;
    case "interrupt": {
      interruptState.value = {
        turnId: evt.turnId,
        toolCalls: evt.toolCalls || [],
        results: new Array((evt.toolCalls || []).length).fill(null),
      };
      for (const tc of evt.toolCalls || []) {
        const t = findToolCall(tc.id);
        if (t) t.status = "pending";
      }
      autoRunInterrupt();
      break;
    }
    case "done":
      if (evt.ctxUsed != null) store.ctxUsed = evt.ctxUsed;
      break;
    case "error":
      store.live = { content: "", thinking: "" };
      store.messages.push({
        id: `err_${Date.now()}`,
        role: "assistant",
        content: `⚠️ ${evt.message}`,
        toolCalls: [],
        thinking: "",
        images: [],
        meta: { error: true },
      });
      scrollBottom();
      break;
  }
}

/* ---------------- 工具中断：自动执行 + 表单提问 ---------------- */
function fillResult(toolCallId, result) {
  const st = interruptState.value;
  if (!st) return;
  const i = st.toolCalls.findIndex((t) => t.id === toolCallId);
  if (i < 0) return;
  st.results[i] = result;
  trySubmitResume();
}

async function autoRunInterrupt() {
  const st = interruptState.value;
  if (!st) return;
  for (const tc of st.toolCalls) {
    if (tc.name === "ask_user") continue; // 等用户作答
    const t = findToolCall(tc.id);
    if (t) t.status = "running";
    const result = await executeTool(tc.name, tc.args);
    if (t) t.status = result.isError ? "error" : "ok";
    fillResult(tc.id, result);
  }
}

async function onAskAnswer({ toolCallId, answer }) {
  const t = findToolCall(toolCallId);
  if (t) t.status = "running";
  fillResult(toolCallId, { content: answer, isError: false });
}

function trySubmitResume() {
  const st = interruptState.value;
  if (!st || st.results.some((r) => r === null)) return;
  interruptState.value = null;
  const results = st.toolCalls.map((tc, i) => ({
    toolCallId: tc.id,
    name: tc.name,
    content: st.results[i]?.content ?? "",
    isError: !!st.results[i]?.isError,
  }));
  api.resume(st.turnId, results).catch((e) => {
    // 会话已失效（服务重启等）
    for (const r of results) {
      const t = findToolCall(r.toolCallId);
      if (t && (t.status === "pending" || t.status === "running")) t.status = "interrupted";
    }
    ElMessage.error(e.message);
  });
}

defineExpose({ openConversation, newChat });
</script>

<template>
  <div class="chat">
    <!-- 模型栏 -->
    <div class="chat-head">
      <el-select
        :model-value="model?.id"
        size="small"
        placeholder="选择模型"
        class="model-select"
        @update:model-value="setActiveModel"
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
      <MessageBubble v-for="m in store.messages" :key="m.id" :message="m" @ask-answer="onAskAnswer" />

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
  user-select: none;
}
.jump-latest:hover {
  background: #eff6ff;
}
.chat-head {
  display: flex;
  gap: 6px;
  align-items: center;
  padding: 8px 0 4px;
  flex: none;
}
.model-select {
  flex: 1;
}
.model-option {
  display: flex;
  align-items: center;
  gap: 4px;
}
.mo-name {
  flex: 1;
  margin-right: 6px;
}
.ctx-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  flex: none;
  padding: 0 2px 6px;
}
.ctx-progress {
  flex: 1;
}
.ctx-text {
  font-size: 11px;
  white-space: nowrap;
}
.msgs {
  flex: 1;
  min-height: 0;
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
  display: flex;
  align-items: center;
  gap: 6px;
  color: #909399;
  font-size: 12px;
  padding: 8px 4px;
}
.dot-spin {
  width: 10px;
  height: 10px;
  border: 2px solid #cbd5e1;
  border-top-color: #2563eb;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}
</style>
