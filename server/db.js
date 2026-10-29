/**
 * SQLite 持久化（better-sqlite3）
 * 表：conversations 对话 | messages 消息 | settings 配置（模型列表、通用设置）
 */
import Database from "better-sqlite3";
import path from "node:path";
import fs from "node:fs";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// 数据目录：默认项目内 data/；安装版通过 OFFICE_AI_DATA_DIR 指向用户数据目录（卸载不丢失）
const dataDir = process.env.OFFICE_AI_DATA_DIR
  ? path.resolve(process.env.OFFICE_AI_DATA_DIR)
  : path.join(__dirname, "..", "data");
fs.mkdirSync(dataDir, { recursive: true });

export const db = new Database(path.join(dataDir, "app.db"));
db.pragma("journal_mode = WAL");

db.exec(`
CREATE TABLE IF NOT EXISTS conversations (
  id         TEXT PRIMARY KEY,
  title      TEXT NOT NULL DEFAULT '新对话',
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL,
  ctx_used   INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS messages (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  conversation_id TEXT NOT NULL,
  role            TEXT NOT NULL,           -- user | assistant | tool
  content         TEXT NOT NULL DEFAULT '',
  thinking        TEXT,
  tool_calls      TEXT,                    -- JSON [{id,name,args,status?}]
  tool_call_id    TEXT,
  name            TEXT,
  images          TEXT,                    -- JSON [dataURL]
  meta            TEXT,                    -- JSON {usage,error,...}
  created_at      INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_messages_conv ON messages(conversation_id, id);
CREATE TABLE IF NOT EXISTS settings (
  key   TEXT PRIMARY KEY,
  value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS attachments (
  id              TEXT PRIMARY KEY,
  conversation_id TEXT NOT NULL,
  name            TEXT NOT NULL,
  ext             TEXT,
  size            INTEGER NOT NULL,
  path            TEXT NOT NULL,
  created_at      INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_attachments_conv ON attachments(conversation_id, id);
`);

/* ---------------- 工作区（按对话 id 建文件夹，对话内产生的文件都存放在此） ---------------- */

export function conversationDir(convId) {
  const dir = path.join(dataDir, "conversations", String(convId));
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

export const DEFAULT_GENERAL = {
  systemPrompt: `你是一个嵌入在 Microsoft Office 中的 AI 助手，可以通过调用工具直接读写用户当前打开的 {APP}。

工作准则：
1. 修改文档前，先使用读取类工具了解现状（结构、内容、格式），再精准操作。
2. 破坏性操作（删除段落/形状、删除行/列/工作表/幻灯片、批量替换）前，必须用 ask_user 工具向用户确认。
3. 只调用与 {APP} 匹配的工具；跨应用工具会失败。复杂操作前可先用 think 工具列出计划。
4. 工具结果出现"[已截断]"时，说明内容超长，请改用更小范围的工具（如按段落/按区域读取）。
5. 完成操作后，用简洁的中文 Markdown 向用户汇报做了什么、结果如何；列出关键数据时使用表格或代码块。`,
  temperature: 0.7,
  maxToolTurns: 25,
  maxToolResultChars: 6000,
};

export const DEFAULT_SETTINGS = {
  models: [],
  general: { ...DEFAULT_GENERAL },
};

export function newId() {
  return crypto.randomUUID();
}

/* ---------------- settings ---------------- */

function safeParse(text, fallback = null) {
  try {
    return text == null ? fallback : JSON.parse(text);
  } catch {
    return fallback;
  }
}

export function getSettings() {
  const stored = {};
  for (const row of db.prepare("SELECT key, value FROM settings").all()) {
    stored[row.key] = safeParse(row.value);
  }
  const settings = {
    models: Array.isArray(stored.models) ? stored.models : [],
    general: { ...DEFAULT_GENERAL, ...(stored.general || {}) },
  };
  return settings;
}

export function saveSettings({ models, general }) {
  if (models !== undefined) {
    if (!Array.isArray(models)) throw new Error("models 必须是数组");
    for (const m of models) {
      if (!m || typeof m !== "object") throw new Error("模型配置必须是对象");
      if (!m.name || !m.model) throw new Error("模型配置缺少 name 或 model");
      if (!m.baseUrl) m.baseUrl = "https://api.openai.com/v1";
      if (m.id == null) m.id = newId();
      if (!m.contextTokens || Number(m.contextTokens) <= 0) m.contextTokens = 128000;
      m.supportsTools = m.supportsTools !== false;
      m.supportsImages = m.supportsImages === true;
      m.streamUsage = m.streamUsage !== false;
      if (m.temperature === "" || m.temperature == null) m.temperature = null;
      else m.temperature = Number(m.temperature);
    }
    db.prepare(
      "INSERT INTO settings (key, value) VALUES ('models', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value"
    ).run(JSON.stringify(models));
  }
  if (general !== undefined) {
    db.prepare(
      "INSERT INTO settings (key, value) VALUES ('general', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value"
    ).run(JSON.stringify({ ...DEFAULT_GENERAL, ...(general || {}) }));
  }
  return getSettings();
}

/* ---------------- conversations ---------------- */

const toClientConversation = (r) => ({
  id: r.id,
  title: r.title,
  createdAt: r.created_at,
  updatedAt: r.updated_at,
  ctxUsed: r.ctx_used,
  messageCount: r.message_count ?? undefined,
});

export function createConversation(title = "新对话") {
  const now = Date.now();
  const conv = { id: newId(), title, created_at: now, updated_at: now, ctx_used: 0 };
  db.prepare(
    "INSERT INTO conversations (id, title, created_at, updated_at, ctx_used) VALUES (@id, @title, @created_at, @updated_at, @ctx_used)"
  ).run(conv);
  return toClientConversation(conv);
}

export function listConversations() {
  return db
    .prepare(
      `SELECT c.*, (SELECT COUNT(*) FROM messages m WHERE m.conversation_id = c.id) AS message_count
       FROM conversations c ORDER BY c.updated_at DESC`
    )
    .all()
    .map(toClientConversation);
}

export function getConversation(id) {
  const row = db.prepare("SELECT * FROM conversations WHERE id = ?").get(id);
  return row ? toClientConversation(row) : null;
}

export function updateConversation(id, { title, ctxUsed } = {}) {
  const conv = db.prepare("SELECT * FROM conversations WHERE id = ?").get(id);
  if (!conv) return null;
  db.prepare("UPDATE conversations SET title = ?, updated_at = ?, ctx_used = ? WHERE id = ?").run(
    title ?? conv.title,
    Date.now(),
    ctxUsed ?? conv.ctx_used,
    id
  );
  return getConversation(id);
}

export function deleteConversation(id) {
  // 清理对话工作区内的附件文件
  for (const a of db.prepare("SELECT path FROM attachments WHERE conversation_id = ?").all(id)) {
    try {
      fs.unlinkSync(a.path);
    } catch { /* 文件可能已被删除 */ }
  }
  db.prepare("DELETE FROM attachments WHERE conversation_id = ?").run(id);
  db.prepare("DELETE FROM messages WHERE conversation_id = ?").run(id);
  db.prepare("DELETE FROM conversations WHERE id = ?").run(id);
  try {
    fs.rmSync(path.join(dataDir, "conversations", String(id)), { recursive: true, force: true });
  } catch { /* 尽力而为 */ }
}

/* ---------------- messages ---------------- */

const toClientMessage = (r) => ({
  id: r.id,
  role: r.role,
  content: r.content ?? "",
  thinking: r.thinking || "",
  toolCalls: safeParse(r.tool_calls, []) || [],
  toolCallId: r.tool_call_id || null,
  name: r.name || null,
  images: safeParse(r.images, []) || [],
  meta: safeParse(r.meta, {}) || {},
  createdAt: r.created_at,
});

export function insertMessage(convId, { role, content = "", thinking = "", toolCalls, toolCallId, name, images, meta }) {
  const info = db
    .prepare(
      `INSERT INTO messages (conversation_id, role, content, thinking, tool_calls, tool_call_id, name, images, meta, created_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .run(
      convId,
      role,
      content ?? "",
      thinking || null,
      toolCalls ? JSON.stringify(toolCalls) : null,
      toolCallId || null,
      name || null,
      images?.length ? JSON.stringify(images) : null,
      meta && Object.keys(meta).length ? JSON.stringify(meta) : null,
      Date.now()
    );
  return getMessageById(info.lastInsertRowid);
}

export function getMessageById(id) {
  const row = db.prepare("SELECT * FROM messages WHERE id = ?").get(id);
  return row ? toClientMessage(row) : null;
}

export function getMessages(convId) {
  return db
    .prepare("SELECT * FROM messages WHERE conversation_id = ? ORDER BY id ASC")
    .all(convId)
    .map(toClientMessage);
}

/* ---------------- attachments（对话附件） ---------------- */

const toClientAttachment = (r) => ({
  id: r.id,
  conversationId: r.conversation_id,
  name: r.name,
  ext: r.ext || null,
  size: r.size,
  createdAt: r.created_at,
});

export function insertAttachment({ id, conversationId, name, ext, size, path: filePath }) {
  db.prepare(
    "INSERT INTO attachments (id, conversation_id, name, ext, size, path, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)"
  ).run(id, conversationId, name, ext || null, size, filePath, Date.now());
  return getAttachment(id);
}

export function getAttachment(id) {
  const row = db.prepare("SELECT * FROM attachments WHERE id = ?").get(id);
  return row ? toClientAttachment(row) : null;
}

export function listAttachments(convId) {
  return db
    .prepare("SELECT * FROM attachments WHERE conversation_id = ? ORDER BY created_at ASC, id ASC")
    .all(convId)
    .map(toClientAttachment);
}

export { safeParse };
