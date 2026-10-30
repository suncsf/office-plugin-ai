<script>
import { ElMessage, ElMessageBox } from "element-ui";
import { store, loadSettings } from "../store";
import { api } from "../api";
import { HOST_LABELS } from "../../src/defaults";
import { developer, department } from "../about-info";

function blankModel() {
  return {
    name: "",
    model: "",
    baseUrl: "https://api.deepseek.com/v1",
    apiKey: "",
    contextTokens: 128000,
    temperature: null,
    supportsTools: true,
    supportsImages: false,
    streamUsage: true,
  };
}

export default {
  name: "SettingsViewLegacy",
  data() {
    return {
      store,
      developer,
      department,
      generalForm: { ...store.settings.general },
      dialogVisible: false,
      editingId: null,
      saving: false,
      testing: false,
      testResult: null,
      health: null,
      modelForm: blankModel(),
      hostLabels: HOST_LABELS,
    };
  },
  mounted() {
    const self = this;
    api
      .health()
      .then((h) => {
        self.health = h;
      })
      .catch(() => {});
  },
  methods: {
    /* ---------------- 通用设置 ---------------- */
    async saveGeneral() {
      try {
        await api.saveSettings({ models: this.store.settings.models, general: { ...this.generalForm } });
        await loadSettings();
        this.generalForm = { ...this.store.settings.general };
        ElMessage.success("通用设置已保存");
      } catch (e) {
        ElMessage.error(e.message);
      }
    },

    /* ---------------- 模型管理 ---------------- */
    openAdd() {
      this.editingId = null;
      this.modelForm = blankModel();
      this.testResult = null;
      this.dialogVisible = true;
    },

    openEdit(m) {
      this.editingId = m.id;
      this.modelForm = JSON.parse(JSON.stringify(m));
      this.testResult = null;
      this.dialogVisible = true;
    },

    async saveModel() {
      if (!this.modelForm.name.trim() || !this.modelForm.model.trim()) {
        ElMessage.warning("请填写名称与模型 ID");
        return;
      }
      this.saving = true;
      try {
        const models = JSON.parse(JSON.stringify(this.store.settings.models));
        if (this.editingId) {
          let found = false;
          for (let i = 0; i < models.length; i++) {
            if (models[i].id === this.editingId) {
              models[i] = { ...models[i], ...this.modelForm, id: this.editingId };
              found = true;
            }
          }
          if (!found) models.push({ ...this.modelForm, id: this.editingId });
        } else {
          models.push({ ...this.modelForm, id: undefined });
        }
        await api.saveSettings({ models: models, general: this.store.settings.general });
        await loadSettings();
        this.dialogVisible = false;
        ElMessage.success("模型已保存");
      } catch (e) {
        ElMessage.error(e.message);
      } finally {
        this.saving = false;
      }
    },

    async removeModel(m) {
      try {
        await ElMessageBox.confirm("删除模型「" + m.name + "」？", "删除模型", { type: "warning" });
      } catch (e) {
        return;
      }
      try {
        const models = this.store.settings.models.filter((x) => x.id !== m.id);
        await api.saveSettings({ models: models, general: this.store.settings.general });
        await loadSettings();
        ElMessage.success("已删除");
      } catch (e) {
        ElMessage.error(e.message);
      }
    },

    async testModel() {
      this.testing = true;
      this.testResult = null;
      try {
        const r = await api.testModel({ ...this.modelForm });
        this.testResult = r;
      } catch (e) {
        this.testResult = { ok: false, error: e.message };
      } finally {
        this.testing = false;
      }
    },
  },
};
</script>

<template>
  <div class="settings">
    <el-scrollbar class="scroll">
      <!-- 服务信息 -->
      <el-card shadow="never" class="card">
        <div slot="header">服务状态</div>
        <div class="kv">
          <span>宿主应用</span><b>{{ hostLabels[store.host] }}</b>
        </div>
        <div class="kv" v-if="health">
          <span>工具数量</span>
          <b>
            Word {{ health.tools.word }} · Excel {{ health.tools.excel }} · PPT {{ health.tools.ppt }} · 通用
            {{ health.tools.common }}
          </b>
        </div>
        <div class="kv" v-else>
          <span>服务状态</span><b class="status-error">未连接（请确认 AI 助手服务已启动）</b>
        </div>
        <div class="kv">
          <span>界面版本</span><b>IE11 兼容版（Office 2019）</b>
        </div>
      </el-card>

      <!-- 通用设置 -->
      <el-card shadow="never" class="card">
        <div slot="header">通用设置</div>
        <div class="field">
          <div class="label">系统提示词（{APP} 会自动替换为当前应用名）</div>
          <el-input v-model="generalForm.systemPrompt" type="textarea" :rows="8" />
        </div>
        <div class="field">
          <div class="label">温度 temperature：{{ generalForm.temperature }}</div>
          <el-slider v-model="generalForm.temperature" :min="0" :max="2" :step="0.1" />
        </div>
        <div class="field">
          <div class="label">最大工具轮次</div>
          <el-input-number v-model="generalForm.maxToolTurns" :min="3" :max="100" size="small" />
        </div>
        <div class="field">
          <div class="label">工具结果截断长度（字符）</div>
          <el-input-number v-model="generalForm.maxToolResultChars" :min="500" :max="50000" :step="500" size="small" />
        </div>
        <div class="field">
          <div class="label">允许 AI 执行 Shell 命令（shell 工具：在本机运行命令行，关闭后 AI 不可见也不可执行）</div>
          <el-switch v-model="generalForm.enableShell" active-text="开启" inactive-text="关闭" />
        </div>
        <el-button type="primary" @click="saveGeneral">保存通用设置</el-button>
      </el-card>

      <!-- 模型管理 -->
      <el-card shadow="never" class="card">
        <div slot="header" class="card-head">
          <span>模型管理（可配置多个，对话中随时切换）</span>
          <el-button size="small" type="plain" icon="el-icon-plus" @click="openAdd">添加模型</el-button>
        </div>
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
            <el-button size="small" icon="el-icon-edit" circle @click="openEdit(m)" />
            <el-button size="small" icon="el-icon-delete" circle type="danger" plain @click="removeModel(m)" />
          </div>
        </div>
      </el-card>

      <!-- 关于 -->
      <el-card shadow="never" class="card">
        <div class="about-head">
          <span class="about-title">AI 助手（Office 加载项）</span>
          <el-tag size="small" effect="plain">v1.0.1</el-tag>
        </div>
        <div class="about-sub" v-if="developer || department">
          <span v-if="developer">开发人员：<b>{{ developer }}</b></span>
          <span v-if="department">部门：{{ department }}</span>
        </div>
      </el-card>
    </el-scrollbar>

    <!-- 模型编辑对话框 -->
    <el-dialog :visible.sync="dialogVisible" :title="editingId ? '编辑模型' : '添加模型'" width="95%">
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

      <div slot="footer" class="dialog-footer">
        <el-button size="small" icon="el-icon-connection" :loading="testing" @click="testModel">测试连接</el-button>
        <el-button size="small" @click="dialogVisible = false">取消</el-button>
        <el-button size="small" type="primary" :loading="saving" @click="saveModel">保存</el-button>
      </div>
    </el-dialog>
  </div>
</template>

<style scoped>
.settings {
  height: 100%;
  padding-top: 8px;
}
.about-head {
  font-weight: 600;
}
.about-head .el-tag {
  margin-left: 8px;
}
.about-sub {
  margin-top: 6px;
  font-size: 12px;
  color: #6b7280;
}
.scroll {
  height: 100%;
}
.card {
  margin-bottom: 10px;
}
.card-head {
  overflow: hidden;
}
.card-head span {
  float: left;
}
.card-head .el-button {
  float: right;
}
.field {
  margin-bottom: 12px;
}
.label {
  font-size: 12px;
  color: #606266;
  margin-bottom: 4px;
}
.kv {
  font-size: 12px;
  padding: 3px 0;
  color: #606266;
  overflow: hidden;
}
.kv span {
  float: left;
}
.kv b {
  float: right;
  color: #303133;
  text-align: right;
}
.model-item {
  padding: 8px 4px;
  border-bottom: 1px solid #f0f2f5;
  overflow: hidden;
}
.mi-main {
  float: left;
  width: 70%;
}
.mi-actions {
  float: right;
  width: 30%;
  text-align: right;
  padding-top: 12px;
}
.mi-name {
  font-weight: 600;
}
.mi-name .el-tag {
  margin-left: 6px;
}
.mi-sub {
  font-size: 11px;
  color: #909399;
  margin: 2px 0 4px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.mi-tags .el-tag {
  margin-right: 4px;
  margin-bottom: 2px;
}
.switches .el-switch {
  display: block;
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
  text-align: right;
}
.dialog-footer .el-button {
  margin-left: 8px;
}
</style>
