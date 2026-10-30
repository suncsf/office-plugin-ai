<script>
import { ElMessage } from "element-ui";
import { activeModel } from "../store";

const MAX_FILE_MB = 15;
const MAX_FILES = 5;

export default {
  name: "ChatInputLegacy",
  props: {
    disabled: { type: Boolean, default: false },
  },
  data() {
    return {
      content: "",
      images: [], // [dataUrl]
      files: [], // [{ name, size, dataUrl }]
      composing: false, // 中文输入法组合输入中（IE11 无 isComposing，需自行跟踪）
      DOC_ACCEPT: ".docx,.dotx,.xlsx,.xlsm,.pptx,.txt,.md,.csv,.json,.log",
    };
  },
  computed: {
    supportsImages() {
      const m = activeModel();
      return !!(m && m.supportsImages);
    },
  },
  methods: {
    fmtSize(n) {
      n = Number(n) || 0;
      if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + "MB";
      if (n >= 1024) return Math.round(n / 1024) + "KB";
      return n + "B";
    },
    send() {
      const text = this.content.trim();
      if (!text && !this.images.length && !this.files.length) return;
      if (this.disabled) return;
      this.$emit("send", {
        content: text,
        images: this.images.slice(0),
        attachments: this.files.map((f) => ({ name: f.name, size: f.size, dataUrl: f.dataUrl })),
      });
      this.content = "";
      this.images = [];
      this.files = [];
    },
    onKeydown(e) {
      // IE11 输入法组合期间 keydown 为 229，且无 isComposing 属性
      if (e.keyCode === 229 || this.composing) return;
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        this.send();
      }
    },
    pickImages() {
      if (!this.supportsImages) {
        ElMessage.warning("当前模型不支持图片");
        return;
      }
      if (this.$refs.fileInput) this.$refs.fileInput.click();
    },
    pickDocs() {
      if (this.$refs.docInput) this.$refs.docInput.click();
    },
    async onFilesChosen(e) {
      const target = e.target;
      const fileList = Array.prototype.slice.call(target.files || []);
      target.value = "";
      for (const f of fileList) {
        if (this.images.length >= 4) {
          ElMessage.warning("最多附带 4 张图片");
          break;
        }
        try {
          const dataUrl = await this.downscale(await this.fileToDataUrl(f));
          this.images.push(dataUrl);
        } catch (err) {
          ElMessage.error("读取图片失败：" + err.message);
        }
      }
    },
    async onDocsChosen(e) {
      const target = e.target;
      const fileList = Array.prototype.slice.call(target.files || []);
      target.value = "";
      for (const f of fileList) {
        if (this.files.length >= MAX_FILES) {
          ElMessage.warning("最多附带 " + MAX_FILES + " 个文档");
          break;
        }
        if (f.size > MAX_FILE_MB * 1024 * 1024) {
          ElMessage.error("「" + f.name + "」超过 " + MAX_FILE_MB + "MB 上限");
          continue;
        }
        try {
          const dataUrl = await this.fileToDataUrl(f);
          this.files.push({ name: f.name, size: f.size, dataUrl: dataUrl });
        } catch (err) {
          ElMessage.error("读取文件失败：" + err.message);
        }
      }
    },
    fileToDataUrl(file) {
      return new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.onerror = () => reject(new Error("文件读取失败"));
        r.readAsDataURL(file);
      });
    },
    downscale(dataUrl, max = 1280) {
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
    },
    removeImage(i) {
      this.images.splice(i, 1);
    },
    removeFile(i) {
      this.files.splice(i, 1);
    },
  },
};
</script>

<template>
  <div class="input-area">
    <!-- 图片预览 -->
    <div v-if="images.length" class="img-previews">
      <div v-for="(img, i) in images" :key="i" class="img-chip">
        <img :src="img" alt="" />
        <span class="remove" @click="removeImage(i)">×</span>
      </div>
    </div>

    <!-- 文档附件 -->
    <div v-if="files.length" class="doc-chips">
      <div v-for="(f, i) in files" :key="f.name + i" class="doc-chip" :title="f.name">
        <span class="doc-icon">📄</span>
        <span class="doc-name">{{ f.name }}</span>
        <span class="doc-size">{{ fmtSize(f.size) }}</span>
        <span class="remove" @click="removeFile(i)">×</span>
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
        @compositionstart="composing = true"
        @compositionend="composing = false"
        @keydown.native="onKeydown"
      />
      <div class="btns">
        <el-button
          size="small"
          circle
          title="插入图片"
          @click="pickImages"
        >
          📷
        </el-button>
        <el-button size="small" circle title="附加文档（AI 可读取内容与格式）" @click="pickDocs">📎</el-button>
        <el-button v-if="disabled" size="small" type="danger" plain @click="$emit('stop')">停止</el-button>
        <el-button v-else size="small" type="primary" @click="send">发送</el-button>
      </div>
    </div>
    <input ref="fileInput" type="file" accept="image/*" multiple style="display: none" @change="onFilesChosen" />
    <input ref="docInput" type="file" :accept="DOC_ACCEPT" multiple style="display: none" @change="onDocsChosen" />
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
  align-items: flex-end;
}
.input-row .el-textarea {
  flex: 1;
}
.btns {
  margin-left: 6px;
}
.btns .el-button {
  display: block;
  margin: 0 0 4px 0;
}
.btns .el-button:last-child {
  margin-bottom: 0;
}
.img-previews {
  padding: 4px 0;
}
.img-chip {
  position: relative;
  display: inline-block;
  margin-right: 6px;
}
.img-chip img {
  width: 56px;
  height: 56px;
  border-radius: 6px;
  border: 1px solid #e4e7ed;
}
.img-chip .remove,
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
.doc-chips {
  padding: 4px 0;
}
.doc-chip {
  position: relative;
  display: inline-block;
  max-width: 240px;
  padding: 3px 22px 3px 8px;
  background: #fff;
  border: 1px solid #d9e2ec;
  border-radius: 6px;
  font-size: 12px;
  color: #374151;
  margin-right: 6px;
}
.doc-chip .doc-name {
  max-width: 140px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.doc-chip .doc-size {
  color: #9ca3af;
  margin-left: 5px;
}
</style>
