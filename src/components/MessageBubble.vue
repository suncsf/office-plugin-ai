<script setup>
import { ref, computed } from "vue";
import { renderMarkdown } from "../utils";

const props = defineProps({
  message: { type: Object, required: true },
});
const emit = defineEmits(["ask-answer"]);

const showArgs = ref({});   // toolCallId -> bool
const showResult = ref(false);

const isUser = computed(() => props.message.role === "user");
const isTool = computed(() => props.message.role === "tool");
const isError = computed(() => !!props.message.meta?.error);

const html = computed(() => (isUser.value ? "" : renderMarkdown(props.message.content)));

function toggleArgs(id) {
  showArgs.value[id] = !showArgs.value[id];
}

/* ask_user 表单 */
const form = ref("");
const choices = ref([]);

function statusIcon(tc) {
  switch (tc.status) {
    case "ok": return "✅";
    case "error": return "❌";
    case "running": return "🔄";
    case "pending": return "⏳";
    case "interrupted": return "⏸";
    default: return "🔧";
  }
}

function submitAsk(tc) {
  let answer = "";
  if (tc.args?.form_type === "single_choice") {
    if (!choices.value.length) return;
    answer = choices.value[0];
  } else if (tc.args?.form_type === "multiple_choice") {
    if (!choices.value.length) return;
    answer = choices.value.join("、");
  } else {
    if (!String(form.value).trim()) return;
    answer = String(form.value).trim();
  }
  emit("ask-answer", { toolCallId: tc.id, answer });
}
</script>

<template>
  <!-- 用户消息 -->
  <div v-if="isUser" class="row user-row">
    <div class="bubble user">
      <div v-if="message.images?.length" class="imgs">
        <img v-for="(img, i) in message.images" :key="i" :src="img" alt="附件图片" />
      </div>
      <div v-if="message.meta?.attachments?.length" class="att-list">
        <span v-for="a in message.meta.attachments" :key="a.id" class="att-chip" :title="a.name">
          📄 {{ a.name }}
        </span>
      </div>
      <div v-if="message.content" class="content pre-wrap">{{ message.content }}</div>
    </div>
  </div>

  <!-- 工具结果 -->
  <div v-else-if="isTool" class="row tool-row">
    <div class="tool-result" :class="{ error: message.meta?.isError }">
      <div class="tr-head" @click="showResult = !showResult">
        <span>{{ message.meta?.isError ? "❌" : "✅" }}</span>
        <span class="tr-name">{{ message.name || "tool" }}</span>
        <span v-if="message.name === 'ask_user'" class="tr-answer-inline">回答：{{ message.content }}</span>
        <span v-else class="tr-toggle">{{ showResult ? "收起" : "详情" }}</span>
      </div>
      <pre v-if="showResult && message.name !== 'ask_user'">{{ message.content }}</pre>
    </div>
  </div>

  <!-- 助手消息 -->
  <div v-else class="row assistant-row">
    <div class="bubble assistant" :class="{ error: isError }">
      <!-- 思考过程 -->
      <div v-if="message.thinking" class="thinking-block">
        <div class="thinking-title">💭 思考过程</div>
        <div class="thinking-content md-body" v-html="renderMarkdown(message.thinking)"></div>
      </div>

      <!-- 正文 markdown -->
      <div v-if="message.content" class="md-body" v-html="html"></div>

      <!-- 工具调用 -->
      <div v-for="tc in message.toolCalls || []" :key="tc.id" class="tool-card">
        <!-- 表单提问 -->
        <template v-if="tc.name === 'ask_user'">
          <div class="tool-head">
            <span>{{ statusIcon(tc) }}</span>
            <span class="tool-name">❓ 向你提问</span>
          </div>
          <div v-if="tc.status === 'pending'" class="tool-body ask-form">
            <div class="ask-question">{{ tc.args?.question }}</div>

            <template v-if="tc.args?.form_type === 'confirm'">
              <div class="ask-actions">
                <el-button size="small" type="primary" @click="emit('ask-answer', { toolCallId: tc.id, answer: '是' })">是</el-button>
                <el-button size="small" type="danger" plain @click="emit('ask-answer', { toolCallId: tc.id, answer: '否' })">否</el-button>
              </div>
            </template>

            <template v-else-if="tc.args?.form_type === 'single_choice'">
              <el-radio-group v-model="choices[0]" class="ask-options">
                <el-radio v-for="(op, i) in tc.args?.options || []" :key="i" :value="op">{{ op }}</el-radio>
              </el-radio-group>
              <div class="ask-actions">
                <el-button size="small" type="primary" :disabled="!choices[0]" @click="submitAsk(tc)">确定</el-button>
              </div>
            </template>

            <template v-else-if="tc.args?.form_type === 'multiple_choice'">
              <el-checkbox-group v-model="choices" class="ask-options">
                <el-checkbox v-for="(op, i) in tc.args?.options || []" :key="i" :value="op">{{ op }}</el-checkbox>
              </el-checkbox-group>
              <div class="ask-actions">
                <el-button size="small" type="primary" :disabled="!choices.length" @click="submitAsk(tc)">提交</el-button>
              </div>
            </template>

            <template v-else>
              <el-input
                v-model="form"
                type="textarea"
                :rows="2"
                placeholder="输入你的回答"
                @keydown.enter.exact.prevent="submitAsk(tc)"
              />
              <div class="ask-actions">
                <el-button size="small" type="primary" @click="submitAsk(tc)">发送</el-button>
              </div>
            </template>
          </div>
          <div v-else class="tool-body">
            <span class="muted">{{ tc.status === "running" ? "已提交，等待处理…" : "已回答（见下方记录）" }}</span>
          </div>
        </template>

        <!-- 普通工具 -->
        <template v-else>
          <div class="tool-head" @click="toggleArgs(tc.id)">
            <span>{{ statusIcon(tc) }}</span>
            <span class="tool-name">{{ tc.name }}</span>
            <span class="tool-toggle">{{ showArgs[tc.id] ? "收起参数" : "参数" }}</span>
          </div>
          <div v-if="showArgs[tc.id]" class="tool-body">
            <pre>{{ JSON.stringify(tc.args, null, 2) }}</pre>
          </div>
        </template>
      </div>
    </div>
  </div>
</template>

<style scoped>
.row {
  display: flex;
  margin: 8px 0;
}
.user-row {
  justify-content: flex-end;
}
.assistant-row {
  justify-content: flex-start;
}
.tool-row {
  justify-content: center;
}
.bubble {
  max-width: 92%;
  border-radius: 10px;
  padding: 8px 12px;
}
.bubble.user {
  background: #2563eb;
  color: #fff;
  border-bottom-right-radius: 3px;
}
.bubble.user .content {
  white-space: pre-wrap;
  word-break: break-word;
}
.bubble.assistant {
  background: #fff;
  border: 1px solid #e4e7ed;
  border-bottom-left-radius: 3px;
  flex: 1;
  max-width: 100%;
}
.bubble.error {
  border-color: #fca5a5;
  background: #fef2f2;
}
.imgs {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-bottom: 4px;
}
.imgs img {
  max-width: 120px;
  max-height: 90px;
  border-radius: 6px;
}
.att-list {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-bottom: 4px;
}
.att-chip {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  max-width: 200px;
  padding: 2px 8px;
  background: rgba(255, 255, 255, 0.75);
  border: 1px solid #dbe4ee;
  border-radius: 5px;
  font-size: 12px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.pre-wrap {
  white-space: pre-wrap;
}
.tool-result {
  width: 100%;
  border: 1px solid #e8ebf0;
  background: #f8fafc;
  border-radius: 6px;
  font-size: 12px;
}
.tool-result.error {
  border-color: #fecaca;
  background: #fef2f2;
}
.tr-head {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 4px 10px;
  cursor: pointer;
}
.tr-name {
  font-family: Consolas, Monaco, monospace;
  color: #57606a;
}
.tr-answer-inline {
  color: #2563eb;
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.tr-toggle {
  margin-left: auto;
  color: #909399;
}
.tool-result pre {
  margin: 0;
  padding: 6px 10px;
  border-top: 1px solid #eef0f3;
  white-space: pre-wrap;
  word-break: break-all;
  color: #57606a;
  max-height: 200px;
  overflow-y: auto;
}
.ask-form .ask-question {
  font-weight: 600;
  margin-bottom: 6px;
  color: #303133;
}
.ask-options {
  display: flex;
  flex-direction: column;
  gap: 2px;
  margin-bottom: 6px;
}
.ask-actions {
  display: flex;
  justify-content: flex-end;
  margin-top: 4px;
}
.muted {
  color: #909399;
}
</style>
