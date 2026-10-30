<script setup>
import { ref, reactive, computed, onMounted, watch } from "vue";
import { ElMessage, ElMessageBox } from "element-plus";
import { Plus, Delete, Edit, Search, Refresh } from "@element-plus/icons-vue";
import { store } from "../store";
import { api } from "../api";
import { formatTime } from "../utils";

/* ---------------- 记忆分类 ---------------- */
// host 取值与服务端一致：word | excel | ppt | common
const HOST_OPTIONS = [
  { value: "all", label: "全部" },
  { value: "word", label: "Word" },
  { value: "excel", label: "Excel" },
  { value: "ppt", label: "PPT" },
  { value: "common", label: "通用" },
];
const HOST_TAG_TYPE = { word: "primary", excel: "success", ppt: "warning", common: "info" };
const hostLabel = (h) => HOST_OPTIONS.find((o) => o.value === h)?.label || h;

/* ---------------- 列表加载 ---------------- */
const loading = ref(false);
const items = ref([]);
const filterHost = ref("all");
const keyword = ref("");

const filtered = computed(() => {
  const k = keyword.value.trim().toLowerCase();
  if (!k) return items.value;
  return items.value.filter(
    (m) => (m.title || "").toLowerCase().includes(k) || (m.content || "").toLowerCase().includes(k)
  );
});

async function load() {
  loading.value = true;
  try {
    items.value = await api.listMemories(filterHost.value);
  } catch (e) {
    ElMessage.error(e.message);
  } finally {
    loading.value = false;
  }
}

onMounted(load);
watch(filterHost, load);
// 切到本 tab 时刷新（AI 可能在对话中刚保存了新记忆）
watch(
  () => store.activeTab,
  (tab) => { if (tab === "memory") load(); }
);

/* ---------------- 新增 / 编辑 ---------------- */
const dialogVisible = ref(false);
const editingId = ref(null);
const saving = ref(false);
const form = reactive({ host: "common", title: "", content: "" });

function openAdd() {
  editingId.value = null;
  Object.assign(form, { host: filterHost.value !== "all" ? filterHost.value : "common", title: "", content: "" });
  dialogVisible.value = true;
}

function openEdit(m) {
  editingId.value = m.id;
  Object.assign(form, { host: m.host, title: m.title, content: m.content });
  dialogVisible.value = true;
}

async function save() {
  if (!form.title.trim()) { ElMessage.warning("请填写标题"); return; }
  if (!form.content.trim()) { ElMessage.warning("请填写内容"); return; }
  saving.value = true;
  try {
    const body = { host: form.host, title: form.title.trim(), content: form.content.trim() };
    if (editingId.value) await api.updateMemory(editingId.value, body);
    else await api.createMemory(body);
    dialogVisible.value = false;
    await load();
    ElMessage.success(editingId.value ? "记忆已更新" : "记忆已保存");
  } catch (e) {
    ElMessage.error(e.message);
  } finally {
    saving.value = false;
  }
}

async function remove(m) {
  try {
    await ElMessageBox.confirm(`删除记忆「${m.title}」？该操作不可恢复。`, "删除记忆", {
      type: "warning",
      confirmButtonText: "删除",
      cancelButtonText: "取消",
    });
  } catch {
    return;
  }
  try {
    await api.deleteMemory(m.id);
    await load();
    ElMessage.success("已删除");
  } catch (e) {
    ElMessage.error(e.message);
  }
}

/* ---------------- 内容展开 / 收起 ---------------- */
const expanded = reactive({}); // id -> bool
const isLong = (m) => (m.content || "").length > 120 || (m.content || "").includes("\n");
</script>

<template>
  <div class="memory">
    <div class="toolbar">
      <el-radio-group v-model="filterHost" size="small">
        <el-radio-button v-for="o in HOST_OPTIONS" :key="o.value" :value="o.value">{{ o.label }}</el-radio-button>
      </el-radio-group>
      <el-input v-model="keyword" placeholder="搜索标题或内容" size="small" clearable :prefix-icon="Search" class="search" />
      <el-button size="small" :icon="Refresh" circle @click="load" title="刷新" />
      <el-button size="small" type="primary" :icon="Plus" @click="openAdd">新增</el-button>
    </div>

    <div class="tip">
      AI 在对话中可通过 memory_save 工具把模板格式、公式、规范等存入这里（跨对话持久）；相关记忆会自动注入后续对话。
    </div>

    <el-scrollbar class="list">
      <el-empty v-if="!filtered.length" :description="loading ? '加载中…' : '暂无记忆'" :image-size="80" />
      <el-card v-for="m in filtered" :key="m.id" shadow="never" class="item">
        <div class="item-head">
          <el-tag :type="HOST_TAG_TYPE[m.host] || 'info'" size="small" effect="plain">{{ hostLabel(m.host) }}</el-tag>
          <span class="item-title">{{ m.title }}</span>
          <span class="item-actions">
            <el-button :icon="Edit" text size="small" @click="openEdit(m)" />
            <el-button :icon="Delete" text size="small" type="danger" @click="remove(m)" />
          </span>
        </div>
        <div class="item-content" :class="{ collapsed: isLong(m) && !expanded[m.id] }">{{ m.content }}</div>
        <div class="item-foot">
          <el-button
            v-if="isLong(m)"
            link size="small" type="primary"
            @click="expanded[m.id] = !expanded[m.id]"
          >{{ expanded[m.id] ? "收起" : "展开全部" }}</el-button>
          <span class="item-time">更新于 {{ formatTime(m.updatedAt) }}</span>
        </div>
      </el-card>
    </el-scrollbar>

    <!-- 新增 / 编辑弹窗 -->
    <el-dialog
      v-model="dialogVisible"
      :title="editingId ? '编辑记忆' : '新增记忆'"
      width="90%"
      :close-on-click-modal="false"
    >
      <el-form label-width="56px" label-position="left">
        <el-form-item label="分类">
          <el-radio-group v-model="form.host">
            <el-radio-button v-for="o in HOST_OPTIONS.slice(1)" :key="o.value" :value="o.value">{{ o.label }}</el-radio-button>
          </el-radio-group>
        </el-form-item>
        <el-form-item label="标题">
          <el-input v-model="form.title" maxlength="100" placeholder="一句话概括，如「公司公文标准模板」" />
        </el-form-item>
        <el-form-item label="内容">
          <el-input
            v-model="form.content"
            type="textarea"
            :rows="10"
            placeholder="完整细节：模板结构、样式参数、公式原文、使用步骤等"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" :loading="saving" @click="save">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<style scoped>
.memory {
  height: 100%;
  display: flex;
  flex-direction: column;
  padding-top: 8px;
}
.toolbar {
  display: flex;
  gap: 6px;
  align-items: center;
  flex: none;
  padding-bottom: 8px;
}
.toolbar .search {
  flex: 1;
  min-width: 0;
}
.tip {
  flex: none;
  font-size: 11px;
  color: #909399;
  padding-bottom: 8px;
  line-height: 1.5;
}
.list {
  flex: 1;
  min-height: 0;
}
.item {
  margin-bottom: 8px;
}
.item-head {
  display: flex;
  align-items: center;
  gap: 6px;
}
.item-title {
  flex: 1;
  min-width: 0;
  font-weight: 600;
  font-size: 13px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.item-actions {
  flex: none;
}
.item-content {
  margin-top: 6px;
  font-size: 12px;
  color: #606266;
  white-space: pre-wrap;
  word-break: break-all;
  line-height: 1.6;
}
.item-content.collapsed {
  display: -webkit-box;
  -webkit-line-clamp: 4;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.item-foot {
  margin-top: 4px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  min-height: 20px;
}
.item-time {
  font-size: 11px;
  color: #c0c4cc;
}
</style>
