/** 核对：tools-catalog 中每个 executor 是否都有对应客户端实现 */
import fs from "node:fs";

const cat = fs.readFileSync("shared/tools-catalog.js", "utf8");
const tools = [...cat.matchAll(/\{\s*name:\s*"([^"]+)"[\s\S]*?executor:\s*"([^"]+)"/g)].map(
  (m) => ({ name: m[1], executor: m[2] })
);
const reg = new Map();
for (const f of ["src/office/word.js", "src/office/excel.js", "src/office/ppt.js", "src/office/index.js"]) {
  const src = fs.readFileSync(f, "utf8");
  for (const m of src.matchAll(/registry\["([^"]+)"\]/g)) reg.set(m[1], f.split("/").pop());
}
const missing = tools.filter((t) => !reg.has(t.executor));
console.log("目录工具数:", tools.length, "| 已实现执行器:", reg.size, "| 缺失:", missing.length);
for (const t of missing) console.log("  ❌", t.name, "→", t.executor);
