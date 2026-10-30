/** 列出目录工具 → 执行器位置与函数体长度（找 stub/缺失实现；side:"server" 的工具在服务端直执行，跳过） */
import fs from "node:fs";
import { TOOLS, SERVER_TOOLS } from "../shared/tools-catalog.js";

const tools = TOOLS.filter((t) => !SERVER_TOOLS.has(t.name));
const files = ["src/office/word.js", "src/office/excel.js", "src/office/ppt.js", "src/office/index.js"];
const reg = new Map();
for (const f of files) {
  const src = fs.readFileSync(f, "utf8");
  const lines = src.split("\n");
  for (const m of src.matchAll(/registry\["([^"]+)"\]\s*=/g)) {
    const lineNo = src.slice(0, m.index).split("\n").length;
    // 估算函数体长度：到下一个 registry[ 或文件尾
    const rest = lines.slice(lineNo).join("\n");
    const next = rest.search(/registry\["[^"]+"\]\s*=|export \{ registry \}/);
    const bodyLen = next >= 0 ? next : rest.length;
    reg.set(m[1], { file: f.split("/").pop(), lineNo, bodyLen });
  }
}
for (const t of tools) {
  if (t.executor === "common.askUser") continue; // ask_user 由任务窗格表单 UI 拦截，无 registry 实现
  const r = reg.get(t.executor);
  const flag = !r ? "❌缺失" : r.bodyLen < 220 ? "⚠️短" : "  ";
  console.log(`${flag} ${t.name.padEnd(28)} ${r ? r.file + ":" + r.lineNo + " (" + Math.round(r.bodyLen / 1000) + "K chars)" : "无实现"}`);
}
