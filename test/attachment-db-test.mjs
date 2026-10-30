/**
 * 附件 DB 链路回归测试（不依赖 Office / 外网）
 * 回归背景：toClientAttachment 曾丢失 path 字段，导致 attachment_read 报
 * "The \"path\" argument must be of type string or an instance of Buffer or URL. Received undefined"
 * （安装版首条真实附件 9.4MB docx 即触发，见 2026-10-30 修复归档）
 * 链路：insertAttachment → getAttachment/listAttachments → readFileSync(att.path) → parseDocument
 */
process.env.OFFICE_AI_DATA_DIR = new URL("./tmp-attachment-db/", import.meta.url).pathname.replace(/^\/(\w:)/, "$1");
import fs from "node:fs";
import path from "node:path";
import { zipSync, strToU8 } from "fflate";

fs.rmSync(process.env.OFFICE_AI_DATA_DIR, { recursive: true, force: true });

const { insertAttachment, getAttachment, listAttachments, deleteConversation, db } = await import("../server/db.js");
const { parseDocument, attachmentOverview } = await import("../server/docparse.js");

let failed = false;
const check = (name, cond, detail = "") => {
  console.log(`${cond ? "✅" : "❌"} ${name}${cond || !detail ? "" : " " + detail}`);
  if (!cond) failed = true;
};
const j = (v) => JSON.stringify(v);

/* 构造最小 docx */
const docxXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 测试标题</w:t></w:r></w:p>
<w:p><w:r><w:t>这是正文段落。</w:t></w:r></w:p>
</w:body></w:document>`;
const docxBytes = zipSync({
  "[Content_Types].xml": strToU8("<Types/>"),
  "word/document.xml": strToU8(docxXml),
});

/* 模拟 /api/chat 的附件落库流程 */
const convId = "conv_test_attachments";
const attDir = path.join(process.env.OFFICE_AI_DATA_DIR, "conversations", convId, "attachments");
fs.mkdirSync(attDir, { recursive: true });
const id = "att_dbtest01";
const filePath = path.join(attDir, `${id}_测试.docx`);
fs.writeFileSync(filePath, docxBytes);
insertAttachment({ id, conversationId: convId, name: "测试.docx", ext: "docx", size: docxBytes.length, path: filePath });

try {
  /* 1. getAttachment 必须带回 path（getParsedDoc 依赖它读文件） */
  const att = getAttachment(id);
  check("getAttachment 返回 path", typeof att?.path === "string" && att.path === filePath, j(att));

  /* 2. listAttachments 同样带回 path（attachment_list 的 overview 依赖） */
  const listed = listAttachments(convId);
  check("listAttachments 返回 path", listed.length === 1 && listed[0].path === filePath);

  /* 3. 端到端：readFileSync(att.path) → 解析成功（复现原报错点） */
  const parsed = parseDocument(att.name, fs.readFileSync(att.path));
  check("按 path 读取并解析 docx", parsed?.type === "docx" && parsed.paragraphCount === 2, j(parsed?.paragraphCount));
  const ov = attachmentOverview(parsed);
  check("overview 正常", ov?.kind === "Word 文档", j(ov));
} finally {
  deleteConversation(convId);
}

db.close();
fs.rmSync(process.env.OFFICE_AI_DATA_DIR, { recursive: true, force: true });

console.log(failed ? "\n❌ 用例失败" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
