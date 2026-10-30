import { defineConfig } from "vite";
import vue2 from "@vitejs/plugin-vue2";
import legacy from "@vitejs/plugin-legacy";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * IE11 兼容版构建（Office 2019 任务窗格）
 * ---------------------------------------------------------------
 * - Vue 2.7 + Element UI（两者官方支持 IE11）
 * - @vitejs/plugin-legacy 产出 ES5 + core-js + regenerator 的 nomodule 包供 IE11 执行
 * - root 指向 src-legacy/：该目录自带 node_modules（vue@2.7.16 + element-ui），
 *   使 @vitejs/plugin-vue2 的 resolveCompiler(root) 与应用代码都解析到 vue2，
 *   与项目根的 vue@3（现代版）完全隔离
 * - 与主构建（vite.config.js，Vue3 + Element Plus）产物共存于 dist/：
 *   先跑主构建（清空 dist），再跑本构建（emptyOutDir: false），legacy 产物文件名带 legacy- 前缀
 * - 服务端按 UA（Trident）把 /taskpane.html 路由到 taskpane-legacy.html
 */
const ROOT = path.dirname(fileURLToPath(import.meta.url));
const APP_ROOT = path.join(ROOT, "src-legacy");

export default defineConfig({
  root: APP_ROOT,
  publicDir: false,
  plugins: [
    vue2(),
    legacy({
      targets: ["ie >= 11"],
      renderLegacyChunks: true,
      modernPolyfills: false,
    }),
  ],
  resolve: {
    alias: {
      "@shared": path.join(ROOT, "shared"),
    },
  },
  build: {
    outDir: path.join(ROOT, "dist"),
    emptyOutDir: false, // 主构建已清空 dist，本构建只追加 legacy 产物
    chunkSizeWarningLimit: 4000,
    rollupOptions: {
      input: {
        // 注意：不要用带连字符 "-legacy-" 的名字命名入口，否则与 plugin-legacy
        // 给 babel 产物追加的 "-legacy-" 中缀混淆，产物文件名无法区分 module/legacy 包
        "taskpaneLegacy": path.join(APP_ROOT, "taskpane-legacy.html"),
      },
      output: {
        entryFileNames: "assets/legacy-[name]-[hash].js",
        chunkFileNames: "assets/legacy-[name]-[hash].js",
        assetFileNames: "assets/legacy-[name]-[hash].[ext]",
      },
    },
  },
});
