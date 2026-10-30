/**
 * Office 执行器注册中心（IE11 兼容版）
 * 与现代版 src/office/index.js 保持一致；区别仅在于：
 *  - store 指向 src-legacy/store.js（Vue 2.7 observable）
 *  - word/excel/ppt 执行器直接复用现代版模块（无 Vue 依赖，经 babel 转译后可在 IE11 运行）
 * 注意：word/excel/ppt 的 run_code 依赖 AsyncFunction（async/await 语言特性），
 * IE11 无法执行；服务端收到 caps.legacy 时已把这些工具从 LLM 绑定中剔除。
 */
import { TOOL_MAP, HOST_NAMES } from "@shared/tools-catalog.js";
import { store } from "../store";
import { truncate } from "../utils";
import { stripDataUrl } from "../../src/office/dataurl";
import * as wordExec from "../../src/office/word";
import * as excelExec from "../../src/office/excel";
import * as pptExec from "../../src/office/ppt";

const registry = {
  "common.getOfficeContext": async () => {
    let diag = null;
    try {
      const d = Office.context.diagnostics;
      diag = { host: d && d.host, version: d && d.version, platform: d && d.platform };
    } catch (e) {
      /* noop */
    }
    return { appType: store.host, appName: HOST_NAMES[store.host] || "未知", diagnostics: diag };
  },
  "common.think": async (args) => "已记录思考：" + String(args.thought || "").slice(0, 200),
  ...wordExec.registry,
  ...excelExec.registry,
  ...pptExec.registry,
};

function detectHost() {
  if (typeof Office === "undefined" || !Office.context) return "unknown";
  const host = Office.context.host;
  if (host === (Office.HostType && Office.HostType.Word) || host === "Word") return "word";
  if (host === (Office.HostType && Office.HostType.Excel) || host === "Excel") return "excel";
  if (host === (Office.HostType && Office.HostType.PowerPoint) || host === "PowerPoint") return "ppt";
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

/**
 * 执行一个工具调用
 * @returns {Promise<{content: string, isError: boolean}>}
 */
export async function executeTool(name, args = {}) {
  const def = TOOL_MAP[name];
  const maxChars = Number(store.settings.general && store.settings.general.maxToolResultChars) || 6000;

  if (!def) return { content: "未知工具：" + name, isError: true };
  if (!registry[def.executor])
    return { content: "工具 " + name + " 未实现（executor: " + def.executor + "）", isError: true };
  if (def.host !== "common" && store.host !== "unknown" && def.host !== store.host) {
    return {
      content:
        "当前打开的是 " + (HOST_NAMES[store.host] || store.host) + "，无法执行 " +
        (HOST_NAMES[def.host] || def.host) + " 的工具「" + name + "」。",
      isError: true,
    };
  }
  if (def.host !== "common" && store.host === "unknown") {
    return {
      content: "未检测到 Office 宿主环境，无法执行文档工具。请在 Word/Excel/PowerPoint 的任务窗格中使用。",
      isError: true,
    };
  }

  try {
    const out = await registry[def.executor](args);
    const content = typeof out === "string" ? out : JSON.stringify(out, null, 0);
    return { content: truncate(content, maxChars), isError: false };
  } catch (err) {
    const msg = (err && err.message) || (err && err.debugInfo && err.debugInfo.message) || String(err);
    const code = err && err.code ? err.code : err && err.name && err.name !== "Error" ? err.name : null;
    const loc =
      (err && err.debugInfo && err.debugInfo.errorLocation) ||
      (err && err.innerError && err.innerError.debugInfo && err.innerError.debugInfo.errorLocation);
    const extra = [code, loc ? "位置: " + loc : null].filter(Boolean).join(" @ ");
    return { content: "工具执行失败：" + (extra ? msg + "（" + extra + "）" : msg), isError: true };
  }
}

export { stripDataUrl };
