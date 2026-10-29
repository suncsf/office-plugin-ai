/**
 * HTTP 服务
 * ---------------------------------------------------------------
 * - HTTPS（office-addin-dev-certs）供 Office 加载项加载任务窗格
 * - 托管 Vue 构建产物（dist/）
 * - REST API：设置/模型、对话历史（SQLite）
 * - POST /api/chat        SSE 流式对话（含工具调用中断协议）
 * - POST /api/chat/resume 任务窗格回传工具执行结果，继续 LangGraph 图
 */
import express from "express";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";
import { ChatOpenAI } from "@langchain/openai";

import {
  getSettings, saveSettings,
  createConversation, listConversations, getConversation, updateConversation, deleteConversation,
  insertMessage, getMessages,
  insertAttachment, getAttachment, listAttachments, conversationDir, newId,
} from "./db.js";
import { runAgentTurn } from "./agent.js";
import { TOOLS } from "../shared/tools-catalog.js";
import { parseDocument, attachmentOverview, attachmentRead, SUPPORTED_EXTS } from "./docparse.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const distDir = path.join(__dirname, "..", "dist");

const app = express();
app.use(express.json({ limit: "25mb" }));

const wrap = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/* ---------------- 健康检查 ---------------- */

app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    version: "1.0.0",
    tools: {
      total: TOOLS.length,
      word: TOOLS.filter((t) => t.host === "word").length,
      excel: TOOLS.filter((t) => t.host === "excel").length,
      ppt: TOOLS.filter((t) => t.host === "ppt").length,
      common: TOOLS.filter((t) => t.host === "common").length,
    },
  });
});

/* ---------------- 设置 / 模型 ---------------- */

app.get("/api/settings", (req, res) => res.json(getSettings()));

app.put("/api/settings", wrap(async (req, res) => {
  res.json(saveSettings(req.body || {}));
}));

app.post("/api/settings/test-model", wrap(async (req, res) => {
  const model = req.body || {};
  if (!model.model) return res.status(400).json({ ok: false, error: "缺少 model 名称" });
  const t0 = Date.now();
  try {
    const llm = new ChatOpenAI({
      model: model.model,
      apiKey: model.apiKey || "not-set",
      temperature: 0,
      maxRetries: 0,
      timeout: 30000,
      configuration: { baseURL: model.baseUrl || "https://api.openai.com/v1" },
    });
    const resp = await llm.invoke([{ role: "user", content: "测试连接，请只回复：pong" }]);
    res.json({ ok: true, ms: Date.now() - t0, sample: String(resp.content || "").slice(0, 100) });
  } catch (err) {
    res.json({ ok: false, ms: Date.now() - t0, error: err?.message || String(err) });
  }
}));

/* ---------------- 对话历史 ---------------- */

app.get("/api/conversations", (req, res) => res.json(listConversations()));

app.post("/api/conversations", wrap(async (req, res) => {
  res.json(createConversation(req.body?.title || "新对话"));
}));

app.get("/api/conversations/:id", wrap(async (req, res) => {
  const conv = getConversation(req.params.id);
  if (!conv) return res.status(404).json({ error: "对话不存在" });
  res.json({ conversation: conv, messages: getMessages(conv.id) });
}));

app.patch("/api/conversations/:id", wrap(async (req, res) => {
  const conv = updateConversation(req.params.id, { title: req.body?.title });
  if (!conv) return res.status(404).json({ error: "对话不存在" });
  res.json(conv);
}));

app.delete("/api/conversations/:id", wrap(async (req, res) => {
  deleteConversation(req.params.id);
  res.json({ ok: true });
}));

/* ---------------- AI 对话（SSE + 客户端工具中断协议） ---------------- */

const pendingResolves = new Map(); // turnId -> { toolCalls, resolve, timer }
const busyConversations = new Set();

/* ---------------- 附件（服务端工具） ---------------- */

const MAX_ATTACHMENT_MB = 15;
const MAX_ATTACHMENTS = 5;
const parsedCache = new Map(); // attachmentId -> 解析结果（LRU，容量 24）

function getParsedDoc(att) {
  if (parsedCache.has(att.id)) return parsedCache.get(att.id);
  const parsed = parseDocument(att.name, fs.readFileSync(att.path));
  if (parsedCache.size >= 24) parsedCache.delete(parsedCache.keys().next().value);
  parsedCache.set(att.id, parsed);
  return parsed;
}

function safeOverview(att) {
  try {
    return attachmentOverview(getParsedDoc(att));
  } catch (e) {
    return { error: e?.message || String(e) };
  }
}

/** 服务端工具执行器（绑定对话：附件按对话隔离） */
function attachmentExecutor(convId) {
  return async (name, args) => {
    if (name === "attachment_list") {
      const atts = listAttachments(convId);
      return {
        count: atts.length,
        attachments: atts.map((a) => ({
          id: a.id,
          name: a.name,
          size: a.size,
          ext: a.ext,
          overview: safeOverview(a),
        })),
      };
    }
    if (name === "attachment_read") {
      const att = getAttachment(String(args?.id || ""));
      if (!att || att.conversationId !== convId) throw new Error(`附件不存在或不属于当前对话：${args?.id}`);
      return attachmentRead(getParsedDoc(att), args || {});
    }
    throw new Error(`未知的服务端工具：${name}`);
  };
}

function normalizeResults(results) {
  const arr = Array.isArray(results) ? results : [];
  return arr.map((r) => ({
    content: typeof r?.content === "string" ? r.content : JSON.stringify(r?.content ?? ""),
    isError: !!r?.isError,
  }));
}

app.post("/api/chat", wrap(async (req, res) => {
  const body = req.body || {};
  const content = typeof body.content === "string" ? body.content : "";
  const images = Array.isArray(body.images) ? body.images : [];
  const appType = ["word", "excel", "ppt", "unknown"].includes(body.appType) ? body.appType : "unknown";
  const modelId = body.modelId || null;

  // SSE 头
  res.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-cache, no-transform",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });
  const send = (evt) => { try { res.write(`data: ${JSON.stringify(evt)}\n\n`); } catch { /* noop */ } };
  const heartbeat = setInterval(() => { try { res.write(": hb\n\n"); } catch { /* noop */ } }, 15000);

  const controller = new AbortController();
  let waitState = null;
  let finished = false;
  let conv = null;

  // 客户端提前断开（停止按钮 / 关闭面板）：ServerResponse close 且响应未正常结束
  res.on("close", () => {
    if (finished || res.writableEnded) return;
    controller.abort();
    if (waitState) {
      waitState.resolve((waitState.toolCalls || []).map(() => ({ content: "[用户中断了请求]", isError: true })));
    }
  });

  try {
    conv = body.conversationId ? getConversation(body.conversationId) : null;
    if (!conv) conv = createConversation();

    const settings = getSettings();
    const model = settings.models.find((m) => m.id === modelId) || settings.models[0];
    if (!model) {
      send({ type: "error", message: "尚未配置 AI 模型，请到「设置 → 模型管理」添加（需 OpenAI 兼容接口）" });
      send({ type: "done" });
      return;
    }
    if (busyConversations.has(conv.id)) {
      send({ type: "error", message: "当前对话正在处理中，请稍候" });
      send({ type: "done" });
      return;
    }
    busyConversations.add(conv.id);

    // 附件：保存到对话工作区 data/conversations/<对话id>/attachments/，并解析校验
    const attachmentsMeta = [];
    const rawAtts = Array.isArray(body.attachments) ? body.attachments.slice(0, MAX_ATTACHMENTS) : [];
    try {
      for (const a of rawAtts) {
        const name = String(a?.name || "attachment").split(/[\\/]/).pop().slice(0, 120) || "attachment";
        const dataUrl = String(a?.dataUrl || "");
        const b64 = dataUrl.includes(",") ? dataUrl.slice(dataUrl.indexOf(",") + 1) : "";
        if (!b64) continue;
        const buf = Buffer.from(b64, "base64");
        if (!buf.length) continue;
        const ext = name.includes(".") ? name.split(".").pop().toLowerCase() : "";
        if (!SUPPORTED_EXTS.includes(ext)) {
          throw new Error(`不支持的附件格式：${name}（支持 ${SUPPORTED_EXTS.join(" / ")}）`);
        }
        if (buf.length > MAX_ATTACHMENT_MB * 1024 * 1024) {
          throw new Error(`附件过大（>${MAX_ATTACHMENT_MB}MB）：${name}`);
        }
        const id = `att_${newId().slice(0, 8)}`;
        const attDir = path.join(conversationDir(conv.id), "attachments");
        fs.mkdirSync(attDir, { recursive: true });
        const filePath = path.join(attDir, `${id}_${name}`);
        fs.writeFileSync(filePath, buf);
        insertAttachment({ id, conversationId: conv.id, name, ext, size: buf.length, path: filePath });
        attachmentsMeta.push({ id, name, size: buf.length });
      }
    } catch (e) {
      send({ type: "error", message: `附件处理失败：${e?.message || e}` });
      send({ type: "done" });
      return;
    }

    const userMsg = insertMessage(conv.id, {
      role: "user",
      content,
      images,
      meta: attachmentsMeta.length ? { attachments: attachmentsMeta } : undefined,
    });
    // 自动用首条用户消息生成标题
    if (!conv.title || conv.title === "新对话") {
      const title = (content || "[图片]").replace(/\s+/g, " ").trim().slice(0, 30) || "新对话";
      updateConversation(conv.id, { title });
      conv = getConversation(conv.id);
    }

    const turnId = `${conv.id}_${Date.now().toString(36)}`;
    send({
      type: "start",
      conversationId: conv.id,
      turnId,
      userMessage: userMsg,
      model: { id: model.id, name: model.name, model: model.model, supportsTools: model.supportsTools !== false },
    });

    const historyRows = getMessages(conv.id);
    let ctxUsed = conv.ctxUsed || 0;

    await runAgentTurn({
      turnId,
      historyRows,
      model,
      general: settings.general,
      appType,
      signal: controller.signal,
      onDelta: (kind, text) => send({ type: kind === "thinking" ? "thinking_delta" : "content_delta", text }),
      onMessage: (m) => {
        const row = insertMessage(conv.id, m);
        send({ type: "message", message: row });
      },
      serverToolExecutor: attachmentExecutor(conv.id),
      onInterrupt: (toolCalls) => {
        send({ type: "interrupt", turnId, toolCalls });
        const st = { toolCalls, resolve: null, timer: null };
        st.promise = new Promise((r) => { st.resolve = r; });
        pendingResolves.set(turnId, st);
        waitState = st;
        // 15 分钟安全超时
        st.timer = setTimeout(() => {
          st.resolve((toolCalls || []).map(() => ({ content: "[等待工具执行超时]", isError: true })));
        }, 15 * 60 * 1000);
      },
      waitForResults: async () => {
        if (!waitState) return [];
        const results = await waitState.promise;
        clearTimeout(waitState.timer);
        pendingResolves.delete(turnId);
        waitState = null;
        return normalizeResults(results);
      },
      onUsage: (usage) => {
        send({ type: "usage", usage });
        if (!usage.estimated && usage.input) ctxUsed = usage.input;
        else if (usage.input) ctxUsed = Math.max(ctxUsed, usage.input);
      },
    });

    updateConversation(conv.id, { ctxUsed });
    send({ type: "done", conversationId: conv.id, ctxUsed });
  } catch (err) {
    send({ type: "error", message: err?.message || String(err) });
    send({ type: "done" });
  } finally {
    finished = true;
    if (conv?.id) busyConversations.delete(conv.id);
    clearInterval(heartbeat);
    res.end();
  }
}));

app.post("/api/chat/resume", wrap(async (req, res) => {
  const { turnId, results } = req.body || {};
  const st = turnId ? pendingResolves.get(turnId) : null;
  if (!st) {
    return res.status(410).json({ error: "没有正在等待的工具调用（可能已超时或服务已重启），请重新发送消息" });
  }
  pendingResolves.delete(turnId);
  clearTimeout(st.timer);
  st.resolve(normalizeResults(results));
  res.json({ ok: true });
}));

/* ---------------- 静态资源（Vue 构建产物） ---------------- */

app.get(["/", "/index.html", "/taskpane.html"], (req, res) => {
  const indexFile = path.join(distDir, "index.html");
  if (!fs.existsSync(indexFile)) {
    return res.status(503).send("任务窗格尚未构建：请先执行 npm run build");
  }
  res.sendFile(indexFile);
});
app.use(express.static(distDir));

/* ---------------- 错误处理 ---------------- */

// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  const status = err?.status || err?.statusCode || 500;
  const message = err?.type === "entity.parse.failed" ? "请求体 JSON 解析失败" : err?.message || "服务器内部错误";
  if (res.headersSent) { try { res.end(); } catch { /* noop */ } return; }
  res.status(status).json({ error: message });
});

/* ---------------- 启动 ---------------- */

function readDevCerts() {
  // office-addin-dev-certs 生成的固定路径（由 npm run setup 安装/生成）
  const dir = path.join(process.env.USERPROFILE || process.env.HOME || "", ".office-addin-dev-certs");
  const key = path.join(dir, "localhost.key");
  const cert = path.join(dir, "localhost.crt");
  const ca = path.join(dir, "ca.crt");
  if (fs.existsSync(key) && fs.existsSync(cert)) {
    const opts = { key: fs.readFileSync(key), cert: fs.readFileSync(cert) };
    if (fs.existsSync(ca)) opts.ca = fs.readFileSync(ca);
    return opts;
  }
  return null;
}

async function start() {
  let server;
  let usingHttps = false;
  const certOpts = readDevCerts();
  if (certOpts) {
    server = https.createServer(certOpts, app);
    usingHttps = true;
  } else {
    console.warn("[警告] 未找到 Office 开发证书（%USERPROFILE%\\.office-addin-dev-certs），回退 HTTP。");
    console.warn("       请执行: npm run setup   （manifest 要求 HTTPS，加载到 Office 前必须配置证书）");
    server = http.createServer(app);
  }

  server.listen(PORT, () => {
    const proto = usingHttps ? "https" : "http";
    console.log("");
    console.log("  ╭──────────────────────────────────────────────────╮");
    console.log("  │        Office AI 助手 · LangGraph 加载项服务      │");
    console.log("  ╰──────────────────────────────────────────────────╯");
    console.log(`  任务窗格: ${proto}://localhost:${PORT}/taskpane.html`);
    console.log(`  健康检查: ${proto}://localhost:${PORT}/api/health`);
    console.log(`  工具数量: Word ${TOOLS.filter((t) => t.host === "word").length} 个 | Excel ${TOOLS.filter((t) => t.host === "excel").length} 个 | PPT ${TOOLS.filter((t) => t.host === "ppt").length} 个 | 通用 ${TOOLS.filter((t) => t.host === "common").length} 个`);
    console.log(`  数据文件: ${path.join(__dirname, "..", "data", "app.db")}`);
    console.log("");
  });
}

start();
