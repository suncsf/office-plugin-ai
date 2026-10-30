<script setup>
import { ref, reactive, onMounted } from "vue";
import { ElMessage, ElMessageBox } from "element-plus";
import { Plus, Delete, Edit, Connection } from "@element-plus/icons-vue";
import { store, loadSettings } from "../store";
import { api } from "../api";
import { HOST_LABELS } from "../defaults";

/* ---------------- 通用设置 ---------------- */
const generalForm = reactive({ ...store.settings.general });

async function saveGeneral() {
  try {
    await api.saveSettings({ models: store.settings.models, general: { ...generalForm } });
    await loadSettings();
    Object.assign(generalForm, store.settings.general);
    ElMessage.success("通用设置已保存");
  } catch (e) {
    ElMessage.error(e.message);
  }
}

/* ---------------- 模型管理 ---------------- */
const dialogVisible = ref(false);
const editingId = ref(null);
const saving = ref(false);
const testing = ref(false);
const testResult = ref(null);

const blankModel = () => ({
  name: "",
  model: "",
  baseUrl: "https://api.deepseek.com/v1",
  apiKey: "",
  contextTokens: 128000,
  temperature: null,
  supportsTools: true,
  supportsImages: false,
  streamUsage: true,
});
const modelForm = reactive(blankModel());

function openAdd() {
  editingId.value = null;
  Object.assign(modelForm, blankModel());
  testResult.value = null;
  dialogVisible.value = true;
}

function openEdit(m) {
  editingId.value = m.id;
  Object.assign(modelForm, JSON.parse(JSON.stringify(m)));
  testResult.value = null;
  dialogVisible.value = true;
}

async function saveModel() {
  if (!modelForm.name.trim() || !modelForm.model.trim()) {
    ElMessage.warning("请填写名称与模型 ID");
    return;
  }
  saving.value = true;
  try {
    let models = JSON.parse(JSON.stringify(store.settings.models));
    if (editingId.value) {
      const i = models.findIndex((m) => m.id === editingId.value);
      if (i >= 0) models[i] = { ...models[i], ...modelForm, id: editingId.value };
      else models.push({ ...modelForm, id: editingId.value });
    } else {
      models.push({ ...modelForm, id: undefined });
    }
    await api.saveSettings({ models, general: store.settings.general });
    await loadSettings();
    dialogVisible.value = false;
    ElMessage.success("模型已保存");
  } catch (e) {
    ElMessage.error(e.message);
  } finally {
    saving.value = false;
  }
}

async function removeModel(m) {
  try {
    await ElMessageBox.confirm(`删除模型「${m.name}」？`, "删除模型", { type: "warning" });
  } catch {
    return;
  }
  try {
    const models = store.settings.models.filter((x) => x.id !== m.id);
    await api.saveSettings({ models, general: store.settings.general });
    await loadSettings();
    ElMessage.success("已删除");
  } catch (e) {
    ElMessage.error(e.message);
  }
}

async function testModel() {
  testing.value = true;
  testResult.value = null;
  try {
    const r = await api.testModel({ ...modelForm });
    testResult.value = r;
  } catch (e) {
    testResult.value = { ok: false, error: e.message };
  } finally {
    testing.value = false;
  }
}

/* ---------------- 服务信息 ---------------- */
const health = ref(null);
onMounted(async () => {
  try {
    health.value = await api.health();
  } catch { /* 忽略 */ }
});
</script>

<template>
  <div class="settings">
    <el-scrollbar class="scroll">
      <!-- 服务信息 -->
      <el-card shadow="never" class="card">
        <template #header>服务状态</template>
        <div class="kv">
          <span>宿主应用</span><b>{{ HOST_LABELS[store.host] }}</b>
        </div>
        <div class="kv" v-if="health">
          <span>工具数量</span>
          <b>
            Word {{ health.tools.word }} · Excel {{ health.tools.excel }} · PPT {{ health.tools.ppt }} · 通用
            {{ health.tools.common }}
          </b>
        </div>
        <div class="kv" v-else>
          <span>服务状态</span><b class="status-error">未连接（请确认 npm run start 已运行）</b>
        </div>
      </el-card>

      <!-- 通用设置 -->
      <el-card shadow="never" class="card">
        <template #header>通用设置</template>
        <div class="field">
          <div class="label">系统提示词（{APP} 会自动替换为当前应用名）</div>
          <el-input v-model="generalForm.systemPrompt" type="textarea" :rows="8" />
        </div>
        <div class="field row2">
          <div class="field">
            <div class="label">温度 temperature：{{ generalForm.temperature }}</div>
            <el-slider v-model="generalForm.temperature" :min="0" :max="2" :step="0.1" />
          </div>
          <div class="field">
            <div class="label">最大工具轮次</div>
            <el-input-number v-model="generalForm.maxToolTurns" :min="3" :max="100" size="small" />
          </div>
        </div>
        <div class="field">
          <div class="label">工具结果截断长度（字符）</div>
          <el-input-number v-model="generalForm.maxToolResultChars" :min="500" :max="50000" :step="500" size="small" />
        </div>
        <el-button type="primary" @click="saveGeneral">保存通用设置</el-button>
      </el-card>

      <!-- 模型管理 -->
      <el-card shadow="never" class="card">
        <template #header>
          <div class="card-head">
            <span>模型管理（可配置多个，对话中随时切换）</span>
            <el-button size="small" type="primary" :icon="Plus" @click="openAdd">添加模型</el-button>
          </div>
        </template>
        <el-empty v-if="!store.settings.models.length" description="尚未添加模型" :image-size="60" />
        <div v-for="m in store.settings.models" :key="m.id" class="model-item">
          <div class="mi-main">
            <div class="mi-name">
              {{ m.name }}
              <el-tag v-if="m.id === store.activeModelId" size="small" type="success" effect="plain">当前使用</el-tag>
            </div>
            <div class="mi-sub">
              {{ m.model }} · {{ m.baseUrl }}
            </div>
            <div class="mi-tags">
              <el-tag size="small" type="info" effect="plain">上下文 {{ m.contextTokens }}</el-tag>
              <el-tag size="small" :type="m.supportsTools !== false ? 'success' : 'danger'" effect="plain">
                {{ m.supportsTools !== false ? "支持工具" : "不支持工具" }}
              </el-tag>
              <el-tag size="small" :type="m.supportsImages ? 'success' : 'info'" effect="plain">
                {{ m.supportsImages ? "支持图片" : "不支持图片" }}
              </el-tag>
            </div>
          </div>
          <div class="mi-actions">
            <el-button size="small" :icon="Edit" circle @click="openEdit(m)" />
            <el-button size="small" :icon="Delete" circle type="danger" plain @click="removeModel(m)" />
          </div>
        </div>
      </el-card>
    </el-scrollbar>

    <!-- 模型编辑对话框 -->
    <el-dialog v-model="dialogVisible" :title="editingId ? '编辑模型' : '添加模型'" width="95%">
      <el-form label-position="top" size="small">
        <el-form-item label="显示名称（用于对话面板切换）" required>
          <el-input v-model="modelForm.name" placeholder="如：DeepSeek V3" />
        </el-form-item>
        <el-form-item label="模型 ID（API 的 model 参数）" required>
          <el-input v-model="modelForm.model" placeholder="如：deepseek-chat / qwen-max / gpt-4o" />
        </el-form-item>
        <el-form-item label="API 地址（OpenAI 兼容 baseURL）">
          <el-input v-model="modelForm.baseUrl" placeholder="https://api.deepseek.com/v1" />
        </el-form-item>
        <el-form-item label="API Key（仅保存在本机 SQLite）">
          <el-input v-model="modelForm.apiKey" type="password" show-password placeholder="sk-..." />
        </el-form-item>
        <el-form-item label="上下文 Token 长度（用于额度显示）">
          <el-input-number v-model="modelForm.contextTokens" :min="2000" :step="1000" />
        </el-form-item>
        <el-form-item label="温度（留空则使用通用设置）">
          <el-input v-model="modelForm.temperature" placeholder="如 0.7，留空用通用设置" style="width: 200px" />
        </el-form-item>
        <div class="switches">
          <el-switch v-model="modelForm.supportsTools" active-text="允许调用工具（操作文档）" />
          <el-switch v-model="modelForm.supportsImages" active-text="支持图片输入" />
          <el-switch v-model="modelForm.streamUsage" active-text="请求返回 Token 用量（部分服务不支持）" />
        </div>
      </el-form>

      <div v-if="testResult" class="test-result" :class="{ ok: testResult.ok }">
        <template v-if="testResult.ok">✅ 连接成功（{{ testResult.ms }}ms）：{{ testResult.sample }}</template>
        <template v-else>❌ 连接失败：{{ testResult.error || testResult.message }}</template>
      </div>

      <template #footer>
        <div class="dialog-footer">
          <el-button size="small" :icon="Connection" :loading="testing" @click="testModel">测试连接</el-button>
          <span style="flex: 1"></span>
          <el-button size="small" @click="dialogVisible = false">取消</el-button>
          <el-button size="small" type="primary" :loading="saving" @click="saveModel">保存</el-button>
        </div>
      </template>
    </el-dialog>
  </div>
</template>

<style scoped>
.settings {
  height: 100%;
  padding-top: 8px;
}
.scroll {
  height: 100%;
}
.card {
  margin-bottom: 10px;
}
.card-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.field {
  margin-bottom: 12px;
}
.label {
  font-size: 12px;
  color: #606266;
  margin-bottom: 4px;
}
.row2 {
  display: flex;
  gap: 16px;
  align-items: flex-end;
}
.row2 .field {
  flex: 1;
}
.kv {
  display: flex;
  justify-content: space-between;
  font-size: 12px;
  padding: 3px 0;
  color: #606266;
}
.kv b {
  color: #303133;
  text-align: right;
}
.model-item {
  display: flex;
  align-items: center;
  padding: 8px 4px;
  border-bottom: 1px solid #f0f2f5;
}
.mi-main {
  flex: 1;
  min-width: 0;
}
.mi-name {
  font-weight: 600;
  display: flex;
  gap: 6px;
  align-items: center;
}
.mi-sub {
  font-size: 11px;
  color: #909399;
  margin: 2px 0 4px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.mi-tags {
  display: flex;
  gap: 4px;
  flex-wrap: wrap;
}
.switches {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 8px;
}
.test-result {
  margin-top: 8px;
  font-size: 12px;
  padding: 6px 10px;
  border-radius: 6px;
  background: #fef2f2;
  color: #b91c1c;
  word-break: break-all;
}
.test-result.ok {
  background: #f0fdf4;
  color: #15803d;
}
.dialog-footer {
  display: flex;
  align-items: center;
}
</style>
