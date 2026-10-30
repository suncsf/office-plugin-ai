#!/usr/bin/env node
/**
 * ES5 产物校验：用 acorn（ecmaVersion: 5）真正解析 dist 下的 legacy 包，
 * 存在任何 ES6+ 语法即失败。相比正则扫描，解析器不受字符串/正则/注释干扰，零误报。
 *
 * 校验范围：plugin-legacy 的 babel 产物（文件名带 "-legacy-" 中缀，即 IE11 实际执行的
 * nomodule 脚本）。module 包（现代浏览器执行）与 IE11 无关，不校验。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as acorn from "acorn";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ASSETS = path.join(ROOT, "dist", "assets");
const LEGACY_CHUNK = /-legacy-[A-Za-z0-9_-]+\.js$/;

let files = [];
if (fs.existsSync(ASSETS)) {
  files = fs.readdirSync(ASSETS).filter((f) => f.startsWith("legacy-") && f.endsWith(".js") && LEGACY_CHUNK.test(f));
}
if (!files.length) {
  console.error("未找到 dist/assets/legacy-*legacy-*.js（plugin-legacy 产物），请先执行 npm run build:legacy");
  process.exit(1);
}

let failed = false;
for (const file of files) {
  const full = path.join(ASSETS, file);
  const content = fs.readFileSync(full, "utf8");
  try {
    acorn.parse(content, { ecmaVersion: 5, sourceType: "script" });
    console.log(`✅ ${file}：ES5 语法校验通过（${(content.length / 1024 / 1024).toFixed(2)} MB）`);
  } catch (e) {
    failed = true;
    const loc = e.loc ? `行 ${e.loc.line} 列 ${e.loc.column}（偏移 ${e.pos}）` : "";
    const context = e.pos != null ? content.slice(Math.max(0, e.pos - 80), e.pos + 80).replace(/\s+/g, " ") : "";
    console.error(`❌ ${file}：非 ES5 语法 → ${e.message} ${loc}`);
    if (context) console.error(`   上下文: …${context}…`);
  }
}

process.exit(failed ? 1 : 0);
