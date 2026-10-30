<script setup>
import { ref, onMounted } from "vue";
import { getRuntimeInfo } from "../utils";
import { developer, department } from "../about-info";
import { version } from "../../package.json";

const runtimeInfo = ref(getRuntimeInfo());
// Office.context.diagnostics 需等 Office.onReady 后才完整（纯浏览器环境无此回调，保持初值）
onMounted(() => {
  try {
    if (window.Office && Office.onReady) Office.onReady(() => { runtimeInfo.value = getRuntimeInfo(); });
  } catch (e) { /* 非 Office 宿主 */ }
});
</script>

<template>
  <div class="about-view">
    <el-scrollbar class="scroll">
      <el-card shadow="never" class="card">
        <div class="about-head">
          <span class="about-title">AI 助手（Office 加载项）</span>
          <el-tag size="small" effect="plain">v{{ version }}</el-tag>
        </div>
        <div class="about-sub" v-if="developer || department">
          <span v-if="developer">开发人员：<b>{{ developer }}</b></span>
          <span v-if="department">部门：{{ department }}</span>
        </div>
        <div class="about-sub">
          <span>运行环境：<b>{{ runtimeInfo }}</b></span>
        </div>
      </el-card>
    </el-scrollbar>
  </div>
</template>

<style scoped>
.about-view {
  height: 100%;
  padding-top: 8px;
}
.card {
  margin-bottom: 10px;
}
.about-head {
  display: flex;
  align-items: center;
  gap: 8px;
}
.about-title {
  font-weight: 600;
}
.about-sub {
  margin-top: 6px;
  font-size: 12px;
  color: #6b7280;
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}
.scroll {
  height: 100%;
}
</style>
