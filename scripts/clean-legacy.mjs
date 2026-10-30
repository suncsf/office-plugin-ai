#!/usr/bin/env node
/**
 * 清理 legacy 构建产物：dist/assets/legacy-* 与 dist/taskpane-legacy.html。
 * 背景：legacy 构建 emptyOutDir:false（与主构建产物共存于 dist），旧 hash 产物会跨次累积，
 * 构建前先清掉本轮的 legacy 文件（不动主构建产物 index.html / assets/index-* / assets/icon-*）。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(ROOT, "dist");
const ASSETS = path.join(DIST, "assets");

let removed = 0;
if (fs.existsSync(ASSETS)) {
  for (const f of fs.readdirSync(ASSETS)) {
    if (f.startsWith("legacy-")) {
      fs.rmSync(path.join(ASSETS, f), { force: true });
      removed++;
    }
  }
}
const html = path.join(DIST, "taskpane-legacy.html");
if (fs.existsSync(html)) {
  fs.rmSync(html, { force: true });
  removed++;
}
console.log(`已清理 legacy 产物 ${removed} 个文件`);
