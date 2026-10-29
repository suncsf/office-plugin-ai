<script setup>
import { ref, computed } from "vue";
import { ElMessage } from "element-plus";
import { activeModel } from "../store";

const props = defineProps({
  disabled: { type: Boolean, default: false },
});
const emit = defineEmits(["send", "stop"]);

const content = ref("");
const images = ref([]);
const files = ref([]); // [{ name, size, dataUrl }]
const fileInput = ref(null);
const docInput = ref(null);

const supportsImages = computed(() => !!activeModel()?.supportsImages);

const DOC_ACCEPT = ".docx,.dotx,.xlsx,.xlsm,.pptx,.txt,.md,.csv,.json,.log";
const MAX_FILE_MB = 15;
const MAX_FILES = 5;

function fmtSize(n) {
  n = Number(n) || 0;
  if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + "MB";
  if (n >= 1024) return Math.round(n / 1024) + "KB";
  return n + "B";
}

function send() {
  const text = content.value.trim();
  if (!text && !images.value.length && !files.value.length) return;
  if (props.disabled) return;
  emit("send", { content: text, images: [...images.value], attachments: files.value.map((f) => ({ name: f.name, size: f.size, dataUrl: f.dataUrl })) });
  content.value = "";
  images.value = [];
  files.value = [];
}

function onKeydown(e) {
  if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    send();
  }
}

function pickImages() {
  if (!supportsImages.value) {
    ElMessage.warning("当前模型不支持图片");
    return;
  }
  fileInput.value?.click();
}

function pickDocs() {
  docInput.value?.click();
}

async function onFilesChosen(e) {
  const fileList = Array.from(e.target.files || []);
  e.target.value = "";
  for (const f of fileList) {
    if (images.value.length >= 4) {
      ElMessage.warning("最多附带 4 张图片");
      break;
    }
    try {
      const dataUrl = await downscale(await fileToDataUrl(f));
      images.value.push(dataUrl);
    } catch (err) {
      ElMessage.error(`读取图片失败：${err.message}`);
    }
  }
}

async function onDocsChosen(e) {
  const fileList = Array.from(e.target.files || []);
  e.target.value = "";
  for (const f of fileList) {
    if (files.value.length >= MAX_FILES) {
      ElMessage.warning(`最多附带 ${MAX_FILES} 个文档`);
      break;
    }
    if (f.size > MAX_FILE_MB * 1024 * 1024) {
      ElMessage.error(`「${f.name}」超过 ${MAX_FILE_MB}MB 上限`);
      continue;
    }
    try {
      const dataUrl = await fileToDataUrl(f);
      files.value.push({ name: f.name, size: f.size, dataUrl });
    } catch (err) {
      ElMessage.error(`读取文件失败：${err.message}`);
    }
  }
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("文件读取失败"));
    r.readAsDataURL(file);
  });
}

function downscale(dataUrl, max = 1280) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, max / Math.max(img.width, img.height));
      if (scale >= 1) return resolve(dataUrl);
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      canvas.getContext("2d").drawImage(img, 0, 0, img.width, img.height);
      resolve(canvas.toDataURL("image/jpeg", 0.85));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}
</script>

<template>
  <div class="input-area">
    <!-- 图片预览 -->
    <div v-if="images.length" class="img-previews">
      <div v-for="(img, i) in images" :key="i" class="img-chip">
        <img :src="img" alt="" />
        <span class="remove" @click="images.splice(i, 1)">×</span>
      </div>
    </div>

    <!-- 文档附件 -->
    <div v-if="files.length" class="doc-chips">
      <div v-for="(f, i) in files" :key="f.name + i" class="doc-chip" :title="f.name">
        <span class="doc-icon">📄</span>
        <span class="doc-name">{{ f.name }}</span>
        <span class="doc-size">{{ fmtSize(f.size) }}</span>
        <span class="remove" @click="files.splice(i, 1)">×</span>
      </div>
    </div>

    <div class="input-row">
      <el-input
        v-model="content"
        type="textarea"
        :autosize="{ minRows: 2, maxRows: 6 }"
        resize="none"
        placeholder="描述你要做的操作，可附加文档（docx/xlsx/pptx/txt），Enter 发送"
        :disabled="disabled"
        @keydown="onKeydown"
      />
      <div class="btns">
        <el-button
          size="small"
          circle
          :type="supportsImages ? 'default' : 'info'"
          title="插入图片"
          @click="pickImages"
        >
          📷
        </el-button>
        <el-button size="small" circle title="附加文档（AI 可读取内容与格式）" @click="pickDocs">📎</el-button>
        <el-button v-if="disabled" size="small" type="danger" plain @click="emit('stop')">停止</el-button>
        <el-button v-else size="small" type="primary" @click="send">发送</el-button>
      </div>
    </div>
    <input ref="fileInput" type="file" accept="image/*" multiple hidden @change="onFilesChosen" />
    <input ref="docInput" type="file" :accept="DOC_ACCEPT" multiple hidden @change="onDocsChosen" />
  </div>
</template>

<style scoped>
.input-area {
  flex: none;
  padding: 6px 0 10px;
  border-top: 1px solid #e4e7ed;
  background: #f5f7fa;
}
.input-row {
  display: flex;
  gap: 6px;
  align-items: flex-end;
}
.btns {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.img-previews {
  display: flex;
  gap: 6px;
  padding: 4px 0;
  flex-wrap: wrap;
}
.img-chip {
  position: relative;
}
.img-chip img {
  width: 56px;
  height: 56px;
  object-fit: cover;
  border-radius: 6px;
  border: 1px solid #e4e7ed;
}
.img-chip .remove {
  position: absolute;
  top: -6px;
  right: -6px;
  width: 16px;
  height: 16px;
  line-height: 14px;
  text-align: center;
  background: #ef4444;
  color: #fff;
  border-radius: 50%;
  cursor: pointer;
  font-size: 12px;
}
.doc-chips {
  display: flex;
  gap: 6px;
  padding: 4px 0;
  flex-wrap: wrap;
}
.doc-chip {
  display: flex;
  align-items: center;
  gap: 5px;
  max-width: 240px;
  padding: 3px 22px 3px 8px;
  background: #fff;
  border: 1px solid #d9e2ec;
  border-radius: 6px;
  position: relative;
  font-size: 12px;
  color: #374151;
}
.doc-chip .doc-name {
  max-width: 140px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.doc-chip .doc-size {
  color: #9ca3af;
  flex: none;
}
.doc-chip .remove {
  position: absolute;
  top: -6px;
  right: -6px;
  width: 16px;
  height: 16px;
  line-height: 14px;
  text-align: center;
  background: #ef4444;
  color: #fff;
  border-radius: 50%;
  cursor: pointer;
  font-size: 12px;
}
</style>
