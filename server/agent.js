/**
 * LangGraph Agent
 * ---------------------------------------------------------------
 * 架构：
 *   START → agent(LLM) ──有 tool_calls──→ tools(客户端工具，interrupt 等待任务窗格执行) → agent
 *                    └──无 tool_calls──→ END
 *
 * - Word/Excel/PPT 的所有工具都在任务窗格（浏览器端，Office.js）执行；
 *   agent 节点决定调用工具后，tools 节点通过 interrupt() 暂停图执行，
 *   由 HTTP 层把工具调用流式推给任务窗格，任务窗格执行完 POST /api/chat/resume 回传结果，
 *   图以 Command({ resume }) 继续，直到模型不再请求工具。
 * - 流式输出：streamMode ["updates","messages"]，token 级推送思考与正文。
 */
import {
  StateGraph,
  MessagesAnnotation,
  MemorySaver,
  interrupt,
  Command,
  START,
  END,
} from "@langchain/langgraph";
import { ChatOpenAI } from "@langchain/openai";
import { AIMessage, HumanMessage, SystemMessage, ToolMessage } from "@langchain/core/messages";
import { TOOLS, HOST_NAMES, SERVER_TOOLS } from "../shared/tools-catalog.js";

export const APP_TYPE_NAMES = {
  word: HOST_NAMES.word,
  excel: HOST_NAMES.excel,
  ppt: HOST_NAMES.ppt,
  unknown: "Office 文档",
};

/* ---------------- 提示词 ---------------- */

export function buildSystemPrompt(appType, general, memories = []) {
  const now = new Date().toLocaleString("zh-CN", { hour12: false });
  const base = String(general.systemPrompt || "")
    .replaceAll("{APP}", APP_TYPE_NAMES[appType] || APP_TYPE_NAMES.unknown)
    + `\n当前应用类型: ${appType}${appType === "unknown" ? "（未知，请先调用 get_office_context 确认）" : ""}\n当前时间: ${now}`;
  return base + buildMemoryPrompt(memories);
}

/**
 * 长期记忆 → 系统提示词附加块
 * 每条内容最多截取 1500 字符、整体最多 12000 字符，防止记忆过多撑爆上下文；
 * 超出部分提示模型用 memory_list / memory_read 工具按需查询。
 */
export function buildMemoryPrompt(memories) {
  if (!Array.isArray(memories) || !memories.length) return "";
  const HOST_TAGS = { word: "Word", excel: "Excel", ppt: "PPT", common: "通用" };
  let total = 0;
  const lines = [];
  for (const m of memories) {
    if (!m || !m.title || !m.content) continue;
    if (total >= 12000) break;
    let content = String(m.content);
    if (content.length > 1500) content = content.slice(0, 1500) + "…[内容过长已截断，可用 memory_list 查询完整内容]";
    total += content.length + m.title.length + 20;
    lines.push(`【${HOST_TAGS[m.host] || m.host}】${m.title}\n${content}`);
  }
  if (!lines.length) return "";
  return (
    `\n\n[用户长期记忆] 以下是用户之前让你记住的内容（跨对话持久），与当前任务相关时直接遵循，不必再询问：\n`
    + lines.join("\n---\n")
    + `\n管理方式：memory_list 查看 / memory_save 保存 / memory_update 更新 / memory_delete 删除。`
  );
}

/* ---------------- 工具绑定 ---------------- */

export function toolsForHost(appType, excludeNames) {
  // excludeNames：需要从 LLM 绑定中剔除的工具名（如 IE11 兼容版的 *_run_code——
  // 其执行依赖 async/await 语言特性，IE11 无法运行）
  const excluded = excludeNames && excludeNames.length ? new Set(excludeNames) : null;
  return TOOLS.filter(
    (t) => (t.host === "common" || t.host === appType) && (!excluded || !excluded.has(t.name))
  ).map((t) => ({
    type: "function",
    function: { name: t.name, description: t.description, parameters: t.parameters },
  }));
}

/* ---------------- 思考内容捕获（reasoning 的唯一采集来源） ----------------
 * 背景：@langchain/openai 0.6.x 解析 chat.completions 增量时会丢弃 delta.reasoning_content
 * （DeepSeek/GLM/Kimi 等兼容接口的思考字段），导致思考既不能实时展示也无法持久化。
 * 方案：向 OpenAI SDK 注入自定义 fetch，旁路（tee）解析 SSE 原始流中的思考增量：
 *   - delta.reasoning_content（DeepSeek/GLM/Kimi 等）
 *   - delta.reasoning（OpenRouter 等字符串形式）
 *   - response.reasoning_summary_text.delta（OpenAI Responses API）
 * 捕获文本在 callModel 结束后写回 res.additional_kwargs.reasoning_content，
 * 由 emitLcMessage 持久化入库；实时增量经 onDelta 直接推送前端。
 */
export function createReasoningCapture() {
  const state = { text: "", onDelta: null, done: Promise.resolve(), _buf: "" };

  const handleEvent = (data) => {
    if (!data || data === "[DONE]") return;
    let j;
    try {
      j = JSON.parse(data);
    } catch {
      return;
    }
    let t = "";
    const d = j?.choices?.[0]?.delta;
    if (d) {
      if (typeof d.reasoning_content === "string") t = d.reasoning_content;
      else if (typeof d.reasoning === "string") t = d.reasoning;
    }
    if (!t && j?.type === "response.reasoning_summary_text.delta" && typeof j.delta === "string") {
      t = j.delta;
    }
    if (t) {
      state.text += t;
      try {
        state.onDelta?.(t);
      } catch { /* 回调异常不影响流 */ }
    }
  };

  const pump = (stream) => {
    const dec = new TextDecoder();
    state.done = (async () => {
      try {
        for await (const piece of stream) {
          state._buf += dec.decode(piece, { stream: true });
          let m;
          while ((m = state._buf.match(/\r?\n\r?\n/))) {
            const raw = state._buf.slice(0, m.index);
            state._buf = state._buf.slice(m.index + m[0].length);
            for (const line of raw.split(/\r?\n/)) {
              if (line.startsWith("data:")) handleEvent(line.slice(5).trim());
            }
          }
        }
      } catch { /* 连接中断等，尽力而为 */ }
    })();
  };

  state.fetch = async (url, init) => {
    const res = await fetch(url, init);
    const ct = String(res.headers?.get?.("content-type") || "");
    if (!ct.includes("text/event-stream") || !res.body) return res;
    const [a, b] = res.body.tee();
    pump(b);
    return new Response(a, { status: res.status, statusText: res.statusText, headers: res.headers });
  };

  state.begin = (onDelta) => {
    state.text = "";
    state._buf = "";
    state.done = Promise.resolve();
    state.onDelta = onDelta || null;
  };
  state.end = async () => {
    try {
      await state.done; // 等旁路解析器消费完整个流，避免尾部增量丢失
    } catch { /* 忽略 */ }
    const t = state.text;
    state.text = "";
    state._buf = "";
    state.onDelta = null;
    return t;
  };
  return state;
}

/* ---------------- 模型 ---------------- */

export function createChatModel(modelCfg, general) {
  const capture = createReasoningCapture();
  const opts = {
    model: modelCfg.model,
    apiKey: modelCfg.apiKey || "not-set",
    streaming: true,
    streamUsage: modelCfg.streamUsage !== false,
    maxRetries: 2,
    timeout: 180000,
    configuration: { baseURL: modelCfg.baseUrl || "https://api.openai.com/v1", fetch: capture.fetch },
  };
  const temperature = modelCfg.temperature ?? general?.temperature;
  if (temperature !== null && temperature !== undefined && temperature !== "") {
    opts.temperature = Number(temperature);
  }
  const model = new ChatOpenAI(opts);
  model.__reasoningCapture = capture;
  return model;
}

/* ---------------- 历史消息 → LangChain 消息 ---------------- */

function textOf(content) {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map((p) => (typeof p === "string" ? p : p?.text || "")).join("");
  return content == null ? "" : String(content);
}

export function historyToLcMessages(rows, { supportsImages }) {
  const out = [];
  // 记录尚未收到结果的 tool_call，之后补一条合成的 ToolMessage，避免请求被 API 拒绝
  let pendingToolIds = [];

  const flushPending = () => {
    for (const { id, name } of pendingToolIds) {
      out.push(
        new ToolMessage({
          tool_call_id: id,
          name,
          content: "[会话中断，未记录到该工具的执行结果]",
          status: "error",
        })
      );
    }
    pendingToolIds = [];
  };

  for (const r of rows) {
    if (r.role === "user") {
      flushPending();
      // 附件标注：让模型知道用户上传了哪些文档、如何读取
      const atts = Array.isArray(r.meta?.attachments) ? r.meta.attachments : [];
      const attNote = atts.length
        ? `\n\n[本条消息带有附件：${atts.map((a, i) => `${i + 1}. ${a.name}（id=${a.id}）`).join("；")}。用 attachment_list 查看概览，attachment_read 读取内容与格式]`
        : "";
      // 贴图落盘标注：图片已存入工作区，可直接用 file 参数插入文档
      const imgFiles = Array.isArray(r.meta?.imageFiles) ? r.meta.imageFiles : [];
      const imgNote = imgFiles.length
        ? `\n\n[本条消息的图片已存入对话工作区：${imgFiles.map((f, i) => `${i + 1}. ${f.file}`).join("；")}。需要把图片插入文档时，把它作为插入图片工具的 file 参数；也可先用 image_process 裁剪/缩放]`
        : "";
      const imgs = supportsImages ? r.images || [] : [];
      if (imgs.length) {
        out.push(
          new HumanMessage({
            content: [
              { type: "text", text: (r.content || "（见图片）") + attNote + imgNote },
              ...imgs.map((url) => ({ type: "image_url", image_url: { url } })),
            ],
          })
        );
      } else {
        out.push(new HumanMessage((r.content || "") + attNote + imgNote));
      }
    } else if (r.role === "assistant") {
      if (r.meta?.error) continue; // 错误占位消息不回传给模型
      const toolCalls = (r.toolCalls || []).map((tc, i) => ({
        id: tc.id || `call_${i}`,
        name: tc.name,
        args: tc.args ?? {},
        type: "tool_call",
      }));
      out.push(new AIMessage({ content: r.content || "", tool_calls: toolCalls.length ? toolCalls : undefined }));
      if (toolCalls.length) pendingToolIds = toolCalls.map((tc) => ({ id: tc.id, name: tc.name }));
    } else if (r.role === "tool") {
      out.push(
        new ToolMessage({
          tool_call_id: r.toolCallId || "",
          name: r.name || undefined,
          content: r.content || "",
          status: r.meta?.isError ? "error" : "success",
        })
      );
      pendingToolIds = pendingToolIds.filter((p) => p.id !== r.toolCallId);
    }
  }
  flushPending();
  return out;
}

/* ---------------- 简单 token 估算（无 usage 时兜底） ---------------- */

export function estimateTokens(text) {
  if (!text) return 0;
  let cjk = 0;
  for (const ch of text) if (/[\u3000-\u9fff\uff00-\uffef]/.test(ch)) cjk++;
  return Math.ceil(cjk * 1 + (text.length - cjk) / 3.5);
}

/* ---------------- 图 ---------------- */

export function buildGraph({ model, tools, systemPrompt, signal, onThinkingDelta, serverToolExecutor }) {
  const capture = model.__reasoningCapture;
  const callModel = async (state) => {
    const messages = [new SystemMessage(systemPrompt), ...state.messages];
    const llm = tools.length ? model.bindTools(tools) : model;
    capture?.begin(onThinkingDelta);
    const res = await llm.invoke(messages, signal ? { signal } : undefined);
    const captured = capture ? await capture.end() : "";
    if (captured) {
      // 原始流捕获到的完整思考（聚合后的 AIMessage 会丢失 reasoning_content，此处回填）
      res.additional_kwargs = { ...(res.additional_kwargs || {}), reasoning_content: captured };
    } else if (typeof res.additional_kwargs?.reasoning === "string" && !res.additional_kwargs?.reasoning_content) {
      // 兜底：个别实现把思考放在 additional_kwargs.reasoning
      res.additional_kwargs.reasoning_content = res.additional_kwargs.reasoning;
    }
    return { messages: [res] };
  };

  const clientTools = async (state) => {
    const last = state.messages[state.messages.length - 1];
    const toolCalls = (last?.tool_calls ?? []).map((tc, i) => ({
      id: tc.id || `call_${i}`,
      name: tc.name,
      args: tc.args ?? {},
    }));

    // 分流：服务端工具（attachment_*）在 Node 内直接执行；客户端工具经 interrupt 交给任务窗格
    const slots = new Array(toolCalls.length).fill(null);
    const clientCalls = [];
    const clientSlotIdx = [];
    for (let i = 0; i < toolCalls.length; i++) {
      const tc = toolCalls[i];
      if (SERVER_TOOLS.has(tc.name) && serverToolExecutor) {
        let content;
        let isError = false;
        try {
          content = await serverToolExecutor(tc.name, tc.args);
          if (typeof content !== "string") content = JSON.stringify(content);
        } catch (e) {
          content = `工具执行失败：${e?.message || e}`;
          isError = true;
        }
        if (content.length > 24000) content = content.slice(0, 24000) + "…[结果过长已截断]";
        slots[i] = { content, isError };
      } else {
        clientCalls.push(tc);
        clientSlotIdx.push(i);
      }
    }
    if (clientCalls.length) {
      // 暂停图执行：HTTP 层把 toolCalls 发给任务窗格，执行结果经 resume 回来
      const results = await interrupt({ type: "client_tools", toolCalls: clientCalls });
      results.forEach((r, i) => {
        slots[clientSlotIdx[i]] = r;
      });
    }
    const msgs = toolCalls.map((tc, i) => {
      const r = slots[i] || { content: "[工具结果缺失]", isError: true };
      let content = typeof r?.content === "string" ? r.content : JSON.stringify(r?.content ?? "");
      if (content.length > 24000) content = content.slice(0, 24000) + "…[结果过长已截断]";
      return new ToolMessage({
        tool_call_id: tc.id || r?.toolCallId || `call_${i}`,
        name: tc.name || r?.name || "tool",
        content,
        status: r?.isError ? "error" : "success",
      });
    });
    return { messages: msgs };
  };

  const route = (state) => {
    const last = state.messages[state.messages.length - 1];
    return last?.tool_calls?.length ? "tools" : END;
  };

  return new StateGraph(MessagesAnnotation)
    .addNode("agent", callModel)
    .addNode("tools", clientTools)
    .addEdge(START, "agent")
    .addConditionalEdges("agent", route, { tools: "tools", [END]: END })
    .addEdge("tools", "agent")
    .compile({ checkpointer: new MemorySaver() });
}

/* ---------------- 一轮对话运行 ---------------- */

/**
 * @param {object} opts
 *  turnId          线程 id（interrupt 恢复依赖同一 checkpointer 线程）
 *  historyRows     SQLite 中的历史消息（含刚保存的用户消息），toClientMessage 形态
 *  model           模型配置
 *  general         通用设置
 *  appType         word | excel | ppt | unknown
 *  memories        注入系统提示词的长期记忆（[{host,title,content}]，可省略）
 *  signal          AbortSignal（用户停止/断开时中断 LLM 调用）
 *  onDelta(kind, text)     token 流：kind = "content" | "thinking"
 *  onMessage(clientMsg)    一条完整消息产生（assistant / tool），由 HTTP 层持久化并推送
 *  onInterrupt(toolCalls)  等待任务窗格执行工具
 *  waitForResults()        阻塞直到 resume（返回与 toolCalls 对齐的结果数组）
 *  onUsage(usage)          LLM 用量（usage = {input, output, total, estimated}）
 */
export async function runAgentTurn(opts) {
  const {
    turnId, historyRows, model, general, appType, signal, excludeToolNames = null,
    memories = [],
    onDelta = () => {}, onMessage = () => {}, onInterrupt = () => {},
    waitForResults = async () => [], onUsage = () => {}, serverToolExecutor = null,
  } = opts;

  const tools = model.supportsTools === false ? [] : toolsForHost(appType, excludeToolNames);
  const systemPrompt = buildSystemPrompt(appType, general, memories);
  const baseMessages = historyToLcMessages(historyRows, { supportsImages: model.supportsImages !== false });
  const chatModel = opts.modelInstance || createChatModel(model, general);
  const graph = buildGraph({
    model: chatModel,
    tools,
    systemPrompt,
    signal,
    onThinkingDelta: (t) => onDelta("thinking", t),
    serverToolExecutor,
  });
  const config = {
    configurable: { thread_id: turnId },
    recursionLimit: Math.max(8, (Number(general?.maxToolTurns) || 25) * 2 + 6),
  };
  if (signal) config.signal = signal;

  let usageLast = null;
  let usageEstimated = false;

  // 增量流只负责正文（thinking 由 buildGraph 的原始流捕获回调直接推送）。
  // 仅处理 AI 消息的 chunk：messages 模式也会把 ToolMessage 等作为 chunk 吐出，
  // 不加过滤会把工具结果 JSON 混进正文增量。
  const handleChunk = (chunk) => {
    if (!chunk) return;
    if (typeof chunk._getType === "function" && chunk._getType() !== "ai") return;
    const c = chunk.content;
    if (typeof c === "string" && c) onDelta("content", c);
    else if (Array.isArray(c)) {
      const t = c.map((p) => (typeof p === "string" ? p : p?.text || "")).join("");
      if (t) onDelta("content", t);
    }
  };

  const emitLcMessage = (m) => {
    if (m._getType?.() === "ai") {
      const usage = m.usage_metadata;
      if (usage) {
        usageLast = { input: usage.input_tokens, output: usage.output_tokens, total: usage.total_tokens };
        onUsage({ ...usageLast, estimated: false });
      }
      onMessage({
        role: "assistant",
        content: textOf(m.content),
        thinking: m.additional_kwargs?.reasoning_content || (typeof m.additional_kwargs?.reasoning === "string" ? m.additional_kwargs.reasoning : "") || "",
        toolCalls: (m.tool_calls ?? []).map((tc, i) => ({
          id: tc.id || `call_${i}`,
          name: tc.name,
          args: tc.args ?? {},
        })),
      });
    } else if (m._getType?.() === "tool") {
      onMessage({
        role: "tool",
        toolCallId: m.tool_call_id,
        name: m.name || m.additional_kwargs?.name || "tool",
        content: textOf(m.content),
        isError: m.status === "error",
      });
    }
  };

  const runStream = async (input) => {
    // LangGraph JS 要求输入为状态对象（裸消息数组需包进 messages 通道）；
    // streamMode 必须放在 options（第二个参数）里，多模式时输出为 [mode, payload] 元组
    const graphInput = Array.isArray(input) ? { messages: input } : input;
    const stream = await graph.stream(graphInput, { ...config, streamMode: ["updates", "messages"] });
    let interrupted = false;
    for await (const item of stream) {
      if (Array.isArray(item) && item.length === 2 && typeof item[0] === "string") {
        const [mode, payload] = item;
        if (mode === "messages") handleChunk(Array.isArray(payload) ? payload[0] : payload);
        else if (mode === "updates") {
          if (payload && typeof payload === "object" && payload.__interrupt__) {
            const intr = Array.isArray(payload.__interrupt__) ? payload.__interrupt__[0] : payload.__interrupt__;
            opts.__pendingCalls = intr?.value?.toolCalls ?? [];
            interrupted = true;
          } else {
            for (const patch of Object.values(payload || {})) {
              for (const m of patch?.messages ?? []) emitLcMessage(m);
            }
          }
        }
      } else if (item && typeof item === "object") {
        // 单 streamMode 兜底
        for (const patch of Object.values(item)) {
          for (const m of patch?.messages ?? []) emitLcMessage(m);
        }
      }
    }
    return interrupted;
  };

  let input = baseMessages;
  try {
    for (let iter = 0; ; iter++) {
      const interrupted = await runStream(input);
      if (!interrupted) break;
      const toolCalls = opts.__pendingCalls || [];
      onInterrupt(toolCalls);
      const results = await waitForResults(); // 与 toolCalls 顺序对齐
      input = new Command({ resume: results });
    }
  } catch (err) {
    const aborted = signal?.aborted;
    const msg = aborted ? "已停止本次生成" : err?.message || String(err);
    onMessage({
      role: "assistant",
      content: aborted ? "⏹ 已停止生成。" : `⚠️ 执行出错：${msg}`,
      meta: { error: true },
    });
    if (!aborted) throw err;
    return { usage: { input: 0, output: 0, total: 0, estimated: true, aborted: true } };
  }

  if (!usageLast) {
    // 供应商未返回用量时按字符估算
    let chars = 0;
    for (const r of historyRows) chars += (r.content || "").length + (r.toolCalls ? JSON.stringify(r.toolCalls).length : 0);
    usageLast = { input: estimateTokens(String(chars)), output: 0, total: estimateTokens(String(chars)) };
    usageEstimated = true;
    onUsage({ ...usageLast, estimated: true });
  }
  return { usage: { ...usageLast, estimated: usageEstimated } };
}
