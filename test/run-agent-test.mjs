/**
 * Agent 离线自测（不需要真实 LLM / Office）
 *   node test/run-agent-test.mjs
 * 覆盖：流式元组解析、工具调用 interrupt/resume 循环、ask_user 表单、历史合成 ToolMessage。
 */
import { BaseChatModel } from "@langchain/core/language_models/chat_models";
import { AIMessage } from "@langchain/core/messages";
import { runAgentTurn, historyToLcMessages } from "../server/agent.js";
import { runShellCommand } from "../server/shell.js";
import os from "node:os";

/** 桩模型：按脚本依次返回响应（文本或 tool_calls） */
class StubModel extends BaseChatModel {
  constructor(script) {
    super({});
    this.script = [...script];
    this.callCount = 0;
  }
  _llmType() {
    return "stub";
  }
  bindTools() {
    return this;
  }
  async _generate(_messages, _options, _runManager) {
    this.callCount++;
    const next = this.script.shift();
    if (next == null) throw new Error("桩模型脚本耗尽（agent 循环异常）");
    let message;
    if (next.tool_calls) {
      message = new AIMessage({
        content: next.content || "",
        tool_calls: next.tool_calls.map((tc, i) => ({ id: tc.id || `stub_${this.callCount}_${i}`, ...tc, type: "tool_call" })),
      });
    } else {
      message = new AIMessage({ content: next });
    }
    return { generations: [{ text: typeof message.content === "string" ? message.content : "", message }] };
  }
}

let failed = false;
function check(name, cond, detail = "") {
  if (cond) console.log(`  ✅ ${name}`);
  else {
    failed = true;
    console.error(`  ❌ ${name} ${detail}`);
  }
}

/* ---------- 场景 1：工具调用循环（agent → tools → agent → END） ---------- */
console.log("\n[场景1] Word 工具调用 + interrupt/resume");
{
  const events = [];
  let queuedResults = [];
  const { usage } = await runAgentTurn({
    turnId: "test-1",
    historyRows: [{ role: "user", content: "帮我看看文档结构", toolCalls: [], images: [] }],
    model: { model: "stub", supportsTools: true },
    general: { systemPrompt: "测试提示词 {APP}", maxToolTurns: 10 },
    appType: "word",
    modelInstance: new StubModel([
      {
        tool_calls: [{ name: "word_get_document_info", args: {} }],
      },
      {
        content: "",
        tool_calls: [{ name: "word_get_paragraphs", args: { offset: 0, limit: 5 } }],
      },
      "文档共 3 段，结构如下……",
    ]),
    onDelta: (kind, text) => events.push(["delta", kind, text]),
    onMessage: (m) => events.push(["message", m.role, m.content?.slice(0, 30), m.toolCalls?.map((t) => t.name)]),
    onInterrupt: (toolCalls) => {
      events.push(["interrupt", toolCalls.map((t) => t.name)]);
      queuedResults = toolCalls.map((tc) => ({ content: JSON.stringify({ ok: true, tool: tc.name }), isError: false }));
    },
    waitForResults: async () => queuedResults,
  });

  const ints = events.filter((e) => e[0] === "interrupt");
  const msgs = events.filter((e) => e[0] === "message");
  check("发生 2 次工具中断", ints.length === 2, JSON.stringify(ints));
  check("工具名正确", ints.flat(2).includes("word_get_document_info") && ints.flat(2).includes("word_get_paragraphs"));
  check("产生 3 条 assistant + 2 条 tool 消息", msgs.filter((m) => m[1] === "assistant").length === 3 && msgs.filter((m) => m[1] === "tool").length === 2, JSON.stringify(msgs));
  check("最终回复在最后", msgs.at(-1)[1] === "assistant" && msgs.at(-1)[2].includes("结构"));
  check("用量兜底估算", usage && typeof usage.input === "number");
}

/* ---------- 场景 2：ask_user 表单提问 ---------- */
console.log("\n[场景2] ask_user 表单提问");
{
  let queuedResults = [];
  const interrupts = [];
  const messages = [];
  await runAgentTurn({
    turnId: "test-2",
    historyRows: [{ role: "user", content: "删除第二段", toolCalls: [], images: [] }],
    model: { model: "stub", supportsTools: true },
    general: { systemPrompt: "测试", maxToolTurns: 10 },
    appType: "word",
    modelInstance: new StubModel([
      {
        tool_calls: [
          {
            name: "ask_user",
            args: { question: "确认删除第二段？", form_type: "confirm", options: ["是", "否"] },
          },
        ],
      },
      "好的，已按你的选择处理。",
    ]),
    onMessage: (m) => messages.push(m),
    onInterrupt: (toolCalls) => {
      interrupts.push(toolCalls);
      queuedResults = toolCalls.map((tc) => ({ content: tc.name === "ask_user" ? "是" : "", isError: false }));
    },
    waitForResults: async () => queuedResults,
  });
  check("ask_user 中断 1 次", interrupts.length === 1 && interrupts[0][0].name === "ask_user");
  check("用户回答成为 ToolMessage", messages.some((m) => m.role === "tool" && m.content === "是"));
  check("最终回复生成", messages.at(-1).role === "assistant" && messages.at(-1).content.includes("按你的选择"));
}

/* ---------- 场景 3：历史消息 → LangChain 消息（悬空 tool_calls 合成结果） ---------- */
console.log("\n[场景3] 历史转换与悬空 tool_calls 修复");
{
  const lc = historyToLcMessages(
    [
      { role: "user", content: "问题", toolCalls: [], images: [] },
      { role: "assistant", content: "", toolCalls: [{ id: "call_a", name: "word_select_text", args: { query: "x" } }] },
      // 缺少 call_a 的 tool 结果
      { role: "user", content: "继续", toolCalls: [], images: [] },
    ],
    { supportsImages: false }
  );
  const kinds = lc.map((m) => m._getType());
  const hasSynthetic = lc.some((m) => m._getType() === "tool" && m.tool_call_id === "call_a");
  check("消息序列 human/ai/tool/human", JSON.stringify(kinds) === JSON.stringify(["human", "ai", "tool", "human"]), JSON.stringify(kinds));
  check("悬空 tool_call 被合成错误结果", hasSynthetic);
  // 图片消息
  const withImg = historyToLcMessages(
    [{ role: "user", content: "看图", toolCalls: [], images: ["data:image/png;base64,xxx"] }],
    { supportsImages: true }
  );
  check("图片转 content blocks", Array.isArray(withImg[0].content) && withImg[0].content[1].type === "image_url");
  const noImg = historyToLcMessages(
    [{ role: "user", content: "看图", toolCalls: [], images: ["data:image/png;base64,xxx"] }],
    { supportsImages: false }
  );
  check("不支持图片的模型剔除图片", typeof noImg[0].content === "string");
}

/* ---------- 场景 4：系统提示词占位符 ---------- */
console.log("\n[场景4] 系统提示词");
{
  const { buildSystemPrompt } = await import("../server/agent.js");
  const p = buildSystemPrompt("excel", { systemPrompt: "操作 {APP}" });
  check("{APP} 替换为 Excel", p.includes("Excel 工作簿"));
  // PPT 编写流程开关（设置 general.pptToolPrecheck：默认直接编写，勾选后先试工具可用性）
  const pptCaps = { host: "ppt", maxVersion: "1.10", capabilities: [], unsupported: [] };
  const off = buildSystemPrompt("ppt", { systemPrompt: "x" }, [], pptCaps);
  check("默认不勾选：注入「直接开始编写」流程", off.includes("直接开始编写") && !off.includes("先试试工具能不能用"), off.slice(-300));
  const on = buildSystemPrompt("ppt", { systemPrompt: "x", pptToolPrecheck: true }, [], pptCaps);
  check("勾选后：注入「先试试工具能不能用」流程", on.includes("先试试工具能不能用") && !on.includes("直接开始编写"), on.slice(-300));
  check("设计指南两档都注入", off.includes("PPT 绘制与美化设计指南") && on.includes("PPT 绘制与美化设计指南"));
  check("非 ppt 宿主不注入 PPT 流程", !buildSystemPrompt("word", { systemPrompt: "x", pptToolPrecheck: true }).includes("先试试工具能不能用"));
}

/* ---------- 场景 5：shell 工具（服务端直执行，不经任务窗格） ---------- */
console.log("\n[场景5] shell 工具服务端直执行");
{
  const interrupts = [];
  const messages = [];
  await runAgentTurn({
    turnId: "test-5",
    historyRows: [{ role: "user", content: "帮我算个东西", toolCalls: [], images: [] }],
    model: { model: "stub", supportsTools: true },
    general: { systemPrompt: "测试", maxToolTurns: 10 },
    appType: "word",
    modelInstance: new StubModel([
      {
        tool_calls: [
          { name: "shell", args: { command: "node -e \"console.log('中文输出- shell-ok-' + (6*7))\"" } },
        ],
      },
      "命令已执行完毕。",
    ]),
    onMessage: (m) => messages.push(m),
    onInterrupt: (tc) => interrupts.push(tc),
    waitForResults: async () => [],
    serverToolExecutor: (name, args) => runShellCommand(args),
  });
  check("shell 不产生任务窗格中断", interrupts.length === 0, JSON.stringify(interrupts));
  const toolMsg = messages.find((m) => m.role === "tool" && m.name === "shell");
  check("shell 结果回传且中文不乱码", toolMsg && toolMsg.content.includes("中文输出- shell-ok-42"), toolMsg?.content?.slice(0, 200));
  check("退出码 0", /"exitCode":0/.test(toolMsg?.content || ""), toolMsg?.content?.slice(0, 200));
}

console.log("\n[场景6] shell 执行细节（退出码/超时/开关）");
{
  const r = await runShellCommand({ command: "node -e \"process.exit(3)\"" });
  check("非零退出码透传", r.exitCode === 3, JSON.stringify(r));
  check("结果含 cwd 且存在", typeof r.cwd === "string" && r.stdout !== undefined);

  const t0 = Date.now();
  const slow = await runShellCommand({ command: "Start-Sleep -Seconds 30", timeout_ms: 1500 }, { defaultCwd: os.tmpdir() });
  check("超时终止进程树", slow.timedOut === true && slow.exitCode === null && Date.now() - t0 < 15000, JSON.stringify({ ...slow, stdout: slow.stdout.slice(0, 50) }));

  const err = await runShellCommand({ command: "echo hi" }, { enabled: false }).catch((e) => e);
  check("开关关闭时拒绝执行", err instanceof Error && err.message.includes("关闭"), String(err));

  const badCwd = await runShellCommand({ command: "echo hi", cwd: "Z:\\__no_such_dir__" }).catch((e) => e);
  check("工作目录不存在时报错", badCwd instanceof Error && badCwd.message.includes("工作目录"), String(badCwd));
}

console.log(failed ? "\n❌ 存在失败项" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
