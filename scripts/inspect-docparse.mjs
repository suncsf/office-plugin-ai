/** 真实 Word 文档解析诊断 */
import { readFileSync } from "node:fs";
import { parseDocument, attachmentRead, attachmentOverview } from "../server/docparse.js";

const buf = readFileSync("test/fixtures/real-word.docx");
try {
  const p = parseDocument("fixtures-real.docx", buf);
  console.log("类型:", p.type, "| 段落数:", p.paragraphCount, "| 表格数:", p.tables.length);
  for (const para of p.paragraphs) {
    console.log(`  [${para.index}] style=${para.style ?? "-"} align=${para.align ?? "-"} font=${para.font?.name ?? "-"} ${para.font?.sizePt ?? "-"}pt bold=${para.font?.bold ?? "-"} | ${String(para.text).slice(0, 30)}`);
  }
  console.log("表格:", JSON.stringify(p.tables));
  for (const mode of ["structure", "text", "format", "tables"]) {
    try {
      const r = attachmentRead(p, { mode });
      console.log(`mode=${mode}: OK`, JSON.stringify(r).slice(0, 120));
    } catch (e) {
      console.log(`mode=${mode}: ❌`, e.message);
    }
  }
  console.log("概览:", JSON.stringify(attachmentOverview(p)).slice(0, 200));
} catch (e) {
  console.log("❌ 解析崩溃:", e.message);
  console.log(e.stack);
}
