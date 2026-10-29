/**
 * thinking 持久化链路实测（离线：本地假 OpenAI SSE 服务 + 真实 ChatOpenAI）
 * 复现生产路径：ChatOpenAI(streaming:true) + llm.invoke() + LangGraph streamMode ["updates","messages"]
 * 观测：① reasoning 增量是否到达 onDelta；② 最终 onMessage 的 assistant 消息是否带 thinking
 */
import http from "node:http";
import { runAgentTurn, createChatModel } from "../server/agent.js";

/** 构造一个 SSE chunk（OpenAI 格式） */
const chunk = (delta, finish = null, extra = {}) =>
  `data: ${JSON.stringify({
    id: "chatcmpl-test",
    object: "chat.completion.chunk",
    created: 1,
    model: "test-model",
    choices: [{ index: 0, delta, finish_reason: finish }],
    ...extra,
  })}\n\n`;

let callCount = 0;

const server = http.createServer((req, res) => {
  let body = "";
  req.on("data", (d) => (body += d));
  req.on("end", () => {
    callCount++;
    res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache" });
    if (callCount === 1) {
      // 第 1 次调用：思考 → 正文 → 工具调用
      for (const r of ["让我", "想想", "步骤"]) {
        res.write(chunk({ role: "assistant", reasoning_content: r }));
      }
      res.write(chunk({ role: "assistant", content: "我来查看" }));
      res.write(
        chunk({
          role: "assistant",
          tool_calls: [{ index: 0, id: "call_1", type: "function", function: { name: "word_get_document_info", arguments: "{}" } }],
        })
      );
      res.write(chunk({}, "tool_calls"));
    } else {
      // 第 2 次调用：思考 → 正文，无工具
      for (const r of ["结合", "结果", "总结"]) {
        res.write(chunk({ role: "assistant", reasoning_content: r }));
      }
      res.write(chunk({ role: "assistant", content: "文档正常" }));
      res.write(chunk({}, "stop", { usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 } }));
    }
    res.write("data: [DONE]\n\n");
    res.end();
  });
});

await new Promise((r) => server.listen(0, "127.0.0.1", r));
const port = server.address().port;

const events = [];
const deltas = { thinking: "", content: "" };
const messages = [];

await runAgentTurn({
  turnId: `test_${Date.now()}`,
  historyRows: [{ role: "user", content: "测试", images: [], toolCalls: [] }],
  model: {
    name: "fake", model: "test-model", baseUrl: `http://127.0.0.1:${port}/v1`, apiKey: "test",
    contextTokens: 1000, supportsTools: true, supportsImages: false, streamUsage: true,
  },
  general: { maxToolTurns: 5 },
  appType: "word",
  modelInstance: createChatModel(
    {
      name: "fake", model: "test-model", baseUrl: `http://127.0.0.1:${port}/v1`, apiKey: "test",
      contextTokens: 1000, supportsTools: true, supportsImages: false, streamUsage: true,
    },
    { maxToolTurns: 5 }
  ),
  onDelta: (kind, text) => (deltas[kind] += text),
  onMessage: (m) => messages.push(m),
  onInterrupt: () => {},
  waitForResults: async () => [{ content: '{"ok":true}' }],
  onUsage: () => {},
});

server.close();

console.log("=== 实测结果 ===");
console.log("thinking deltas 收到:", JSON.stringify(deltas.thinking));
console.log("content deltas 收到 :", JSON.stringify(deltas.content));
for (const m of messages) {
  console.log(
    `onMessage: role=${m.role} thinking=${JSON.stringify((m.thinking || "").slice(0, 40))} content=${JSON.stringify((m.content || "").slice(0, 30))} toolCalls=${(m.toolCalls || []).length}`
  );
}

let failed = false;
const check = (name, cond, detail = "") => {
  console.log(`${cond ? "✅" : "❌"} ${name}${cond ? "" : " " + detail}`);
  if (!cond) failed = true;
};

const assistants = messages.filter((m) => m.role === "assistant");
check("thinking 增量完整（两条模型的思考都推送到前端）", deltas.thinking === "让我想想步骤结合结果总结", JSON.stringify(deltas.thinking));
check("assistant 消息全部带完整 thinking（持久化内容）", assistants.length === 2 && assistants.every((m) => m.thinking === (m.content === "我来查看" ? "让我想想步骤" : "结合结果总结")),
  JSON.stringify(assistants.map((m) => m.thinking)));
check("正文增量不含工具结果 JSON（ToolMessage 已过滤）", deltas.content === "我来查看文档正常", JSON.stringify(deltas.content));
check("思考未混入正文", assistants.every((m) => !(m.content || "").includes("让我") && !(m.content || "").includes("结合")));
console.log(failed ? "\n❌ 存在失败项" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
