/**
 * Office 执行器注册中心
 * - initOffice(): 等待 Office.onReady，识别宿主应用
 * - executeTool(name, args): 按工具目录调用对应执行器，统一包装结果与错误
 */
import { TOOL_MAP, HOST_NAMES } from "@shared/tools-catalog.js";
import { store } from "../store";
import { truncate } from "../utils";
import * as wordExec from "./word";
import * as excelExec from "./excel";
import * as pptExec from "./ppt";

const registry = {
  "common.getOfficeContext": async () => {
    let diag = null;
    try {
      const d = Office.context.diagnostics;
      diag = { host: d?.host, version: d?.version, platform: d?.platform };
    } catch { /* noop */ }
    return { appType: store.host, appName: HOST_NAMES[store.host] || "未知", diagnostics: diag };
  },
  "common.think": async (args) => `已记录思考：${String(args.thought || "").slice(0, 200)}`,
  ...wordExec.registry,
  ...excelExec.registry,
  ...pptExec.registry,
};

function detectHost() {
  if (typeof Office === "undefined" || !Office.context) return "unknown";
  const host = Office.context.host;
  if (host === Office.HostType?.Word || host === "Word") return "word";
  if (host === Office.HostType?.Excel || host === "Excel") return "excel";
  if (host === Office.HostType?.PowerPoint || host === "PowerPoint") return "ppt";
  return "unknown";
}

export function initOffice() {
  return new Promise((resolve) => {
    if (typeof Office === "undefined") {
      store.host = "unknown";
      resolve({ host: "unknown" });
      return;
    }
    Office.onReady(() => {
      store.host = detectHost();
      resolve({ host: store.host });
    });
  });
}

import { stripDataUrl } from "./dataurl";

/**
 * 执行一个工具调用
 * @returns {Promise<{content: string, isError: boolean}>}
 */
export async function executeTool(name, args = {}) {
  const def = TOOL_MAP[name];
  const maxChars = Number(store.settings.general?.maxToolResultChars) || 6000;

  if (!def) return { content: `未知工具：${name}`, isError: true };
  if (!registry[def.executor]) return { content: `工具 ${name} 未实现（executor: ${def.executor}）`, isError: true };
  if (def.host !== "common" && store.host !== "unknown" && def.host !== store.host) {
    return {
      content: `当前打开的是 ${HOST_NAMES[store.host]}，无法执行 ${HOST_NAMES[def.host]} 的工具「${name}」。`,
      isError: true,
    };
  }
  if (def.host !== "common" && store.host === "unknown") {
    return { content: "未检测到 Office 宿主环境，无法执行文档工具。请在 Word/Excel/PowerPoint 的任务窗格中使用。", isError: true };
  }

  try {
    // 注入对话上下文：工作区文件桥接（word_get_image / file 参数插图）需要 conversationId
    const out = await registry[def.executor](args, { conversationId: store.conversationId || null });
    const content = typeof out === "string" ? out : JSON.stringify(out, null, 0);
    return { content: truncate(content, maxChars), isError: false };
  } catch (err) {
    const msg = err?.message || err?.debugInfo?.message || String(err);
    const code = err?.code || (err?.name && err.name !== "Error" ? err.name : null);
    const loc = err?.debugInfo?.errorLocation || err?.innerError?.debugInfo?.errorLocation;
    const extra = [code, loc ? `位置: ${loc}` : null].filter(Boolean).join(" @ ");
    return { content: `工具执行失败：${extra ? `${msg}（${extra}）` : msg}`, isError: true };
  }
}

export { stripDataUrl };
