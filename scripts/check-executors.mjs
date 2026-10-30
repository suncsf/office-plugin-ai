/** 核对：tools-catalog 中每个客户端工具的 executor 是否都有对应实现（side:"server" 的在 Node 内直执行，无需客户端） */
import fs from "node:fs";
import { TOOLS, SERVER_TOOLS } from "../shared/tools-catalog.js";

const clientTools = TOOLS.filter((t) => !SERVER_TOOLS.has(t.name));
const reg = new Map();
for (const f of ["src/office/word.js", "src/office/excel.js", "src/office/ppt.js", "src/office/index.js"]) {
  const src = fs.readFileSync(f, "utf8");
  // registry["key"] = 形式（word/excel/ppt.js）与对象字面量键 "key": 形式（index.js）都算已实现
  for (const m of src.matchAll(/registry\[["']([^"']+)["']\]|^  "([^"]+)":/gm)) {
    reg.set(m[1] || m[2], f.split("/").pop());
  }
}
// ask_user 由任务窗格表单 UI 拦截（ChatView/MessageBubble），无 registry 实现，属预期
const missing = clientTools.filter((t) => !reg.has(t.executor) && t.executor !== "common.askUser");
console.log("目录工具数:", TOOLS.length, "（服务端直执行:", SERVER_TOOLS.size + "）", "| 已实现执行器:", reg.size, "| 缺失:", missing.length);
for (const t of missing) console.log("  ❌", t.name, "→", t.executor);
