/**
 * 长期记忆 DB 链路回归测试（不依赖 Office / 外网）
 * 覆盖：memories 表 CRUD、按分类过滤、字段校验；
 *      buildMemoryPrompt 注入格式与截断（1500 字/条、12000 字总量上限）
 * 链路：insertMemory → listMemories(按 host) → buildSystemPrompt(memories)
 */
process.env.OFFICE_AI_DATA_DIR = new URL("./tmp-memory-db/", import.meta.url).pathname.replace(/^\/(\w:)/, "$1");
import fs from "node:fs";

fs.rmSync(process.env.OFFICE_AI_DATA_DIR, { recursive: true, force: true });

const {
  insertMemory, getMemory, listMemories, updateMemory, deleteMemory, db,
} = await import("../server/db.js");
const { buildSystemPrompt, buildMemoryPrompt } = await import("../server/agent.js");

let failed = false;
const check = (name, cond, detail = "") => {
  console.log(`${cond ? "✅" : "❌"} ${name}${cond || !detail ? "" : " " + detail}`);
  if (!cond) failed = true;
};
const j = (v) => JSON.stringify(v);

try {
  /* 1. 新增：四类 host 均可保存 */
  const word = insertMemory({ host: "word", title: "公司公文标准模板", content: "标题：二号方正小标宋；正文：三号仿宋_GB2312，行距 28 磅。" });
  const excel = insertMemory({ host: "excel", title: "提成计算公式", content: "=IF(销售额>100000,销售额*0.08,销售额*0.05)" });
  const ppt = insertMemory({ host: "ppt", title: "汇报页版式", content: "每页标题 24pt，正文 18pt，最多 5 个要点。" });
  const common = insertMemory({ host: "common", title: "用户偏好", content: "回复使用简体中文，正式场合用「您」。" });
  check("insertMemory 返回 id", [word, excel, ppt, common].every((m) => typeof m.id === "string" && m.id.length > 0));

  /* 2. 字段校验：非法 host / 空标题 / 空内容 */
  check("非法 host 被拒绝", (() => { try { insertMemory({ host: "pdf", title: "x", content: "y" }); return false; } catch { return true; } })());
  check("空标题被拒绝", (() => { try { insertMemory({ host: "word", title: "  ", content: "y" }); return false; } catch { return true; } })());
  check("空内容被拒绝", (() => { try { insertMemory({ host: "word", title: "x", content: "" }); return false; } catch { return true; } })());

  /* 3. 查询：按分类过滤 / 全量 */
  check("listMemories(word) 只返回 Word 分类", listMemories("word").length === 1 && listMemories("word")[0].id === word.id);
  check("listMemories(all) 返回全部 4 条", listMemories("all").length === 4);
  check("getMemory 带回驼峰字段", (() => {
    const m = getMemory(excel.id);
    return m.host === "excel" && m.title === "提成计算公式" && m.content.includes("IF") && !!m.createdAt && !!m.updatedAt;
  })());

  /* 4. 更新：部分字段更新，其余保持不变 */
  const updated = updateMemory(excel.id, { content: "=ROUND(销售额*0.08,2)" });
  check("updateMemory 只改传入字段", updated.content === "=ROUND(销售额*0.08,2)" && updated.title === "提成计算公式");
  check("updateMemory 更新 updated_at", updated.updatedAt >= excel.updatedAt);
  check("updateMemory 不存在的 id 返回 null", updateMemory("mem_not_exist", { title: "x" }) === null);

  /* 5. 删除 */
  check("deleteMemory 删除成功", deleteMemory(ppt.id) === true && listMemories("ppt").length === 0);
  check("deleteMemory 重复删除返回 false", deleteMemory(ppt.id) === false);

  /* 6. 注入：buildMemoryPrompt 输出格式 */
  const prompt = buildMemoryPrompt(listMemories("word").concat(listMemories("common")));
  check("包含分类标签与标题", prompt.includes("【Word】公司公文标准模板") && prompt.includes("【通用】用户偏好"), prompt);
  check("包含工具管理提示", prompt.includes("memory_list") && prompt.includes("memory_save"));

  /* 7. 截断：单条 1500 字符上限 */
  const longContent = "长".repeat(2000);
  const longPrompt = buildMemoryPrompt([{ host: "word", title: "超长记忆", content: longContent }]);
  check("单条内容截断到 1500+省略号", longPrompt.includes("…[内容过长已截断") && !longPrompt.includes("长".repeat(1600)));

  /* 8. 总量：12000 字符上限（单条先截断为 1500+30 提示文字，每条累计 1553；第 9 条累计 13977 触发 break → 注入 8 条） */
  const many = Array.from({ length: 10 }, (_, i) => ({ host: "common", title: `记忆${i}`, content: "字".repeat(2000) }));
  const manyPrompt = buildMemoryPrompt(many);
  const injected = (manyPrompt.match(/【通用】/g) || []).length;
  check("总量超 12000 字符时停止注入（注入 8 条）", injected === 8, `实际注入 ${injected} 条`);
  check("未注入的条目被丢弃", !manyPrompt.includes("记忆9"));

  /* 9. 空记忆不产生注入块 */
  check("无记忆时 buildMemoryPrompt 为空串", buildMemoryPrompt([]) === "");
  const plain = buildSystemPrompt("word", { systemPrompt: "操作 {APP}" });
  check("buildSystemPrompt 兼容不传 memories", plain.startsWith("操作 Word 文档") && !plain.includes("用户长期记忆"));

  /* 10. 带记忆的完整系统提示词 */
  const withMem = buildSystemPrompt("word", { systemPrompt: "操作 {APP}" }, listMemories("word"));
  check("buildSystemPrompt 注入记忆块", withMem.includes("操作 Word 文档") && withMem.includes("【Word】公司公文标准模板"));
} finally {
  db.close();
  fs.rmSync(process.env.OFFICE_AI_DATA_DIR, { recursive: true, force: true });
}

console.log(failed ? "\n❌ 用例失败" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
