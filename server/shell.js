/**
 * Shell 工具（服务端直执行）
 * ---------------------------------------------------------------
 * 在运行 Node 服务的本机（Windows）执行命令，不经任务窗格：
 *  - powershell（默认）：powershell.exe -NoProfile -NonInteractive -Command <script>
 *  - cmd：cmd.exe /d /s /c <command>
 * 编码：输出统一按 UTF-8 解码——执行前把控制台代码页切到 65001
 *   （PowerShell 设 [Console]::OutputEncoding，cmd 用 chcp 65001），
 *   个别老旧中文程序按 GBK 输出时会出现乱码，属已知取舍。
 * 安全：
 *  - enabled=false（设置开关关闭）时拒绝执行；
 *  - 超时用 taskkill /T /F 杀整棵进程树；
 *  - stdout/stderr 各自截断，防止撑爆模型上下文（agent 层还会整体截断一次）。
 */
import { spawn, execFile } from "node:child_process";
import { StringDecoder } from "node:string_decoder";
import fs from "node:fs";

const DEFAULT_TIMEOUT_MS = 60000;
const MAX_TIMEOUT_MS = 600000;
const MAX_STREAM_CHARS = 12000; // stdout / stderr 各自上限

function capStream(text) {
  if (text.length <= MAX_STREAM_CHARS) return text;
  return text.slice(0, MAX_STREAM_CHARS) + `\n…[超长截断，仅保留前 ${MAX_STREAM_CHARS} 字符]`;
}

/** 杀掉进程树（Windows 下 child.kill() 只杀直接子进程，不足以终止孙进程） */
function killTree(pid) {
  try {
    execFile("taskkill", ["/pid", String(pid), "/T", "/F"], { windowsHide: true }, () => {});
  } catch { /* 忽略 */ }
}

function collectStream(stream, onText) {
  const dec = new StringDecoder("utf8");
  stream.on("data", (chunk) => onText(dec.write(chunk)));
  // flush：把可能残留的多字节尾字符并入累计器
  return () => onText(dec.end());
}

/**
 * 执行一条 shell 命令
 * @param {object} args { command, shell?("powershell"|"cmd"), cwd?, timeout_ms? }
 * @param {object} ctx  { defaultCwd?, enabled? }
 * @returns {Promise<object>} { command, shell, cwd, exitCode, timedOut, durationMs, stdout, stderr }
 */
export async function runShellCommand(args = {}, { defaultCwd = process.cwd(), enabled = true } = {}) {
  if (!enabled) throw new Error("Shell 工具已在「设置 → 通用设置」中关闭");

  const command = String(args.command ?? "").trim();
  if (!command) throw new Error("缺少 command 参数");

  const shell = args.shell === "cmd" ? "cmd" : "powershell";
  const timeoutMs = Math.min(
    Math.max(Number(args.timeout_ms) > 0 ? Number(args.timeout_ms) : DEFAULT_TIMEOUT_MS, 1000),
    MAX_TIMEOUT_MS
  );

  const cwd = String(args.cwd || "").trim() || defaultCwd;
  if (!fs.existsSync(cwd) || !fs.statSync(cwd).isDirectory()) {
    throw new Error(`工作目录不存在：${cwd}`);
  }

  const t0 = Date.now();
  let child;
  if (shell === "powershell") {
    // ① 输出编码切 UTF-8（含对原生程序 stdout 的解码）；② 末尾透传原生命令的退出码
    const script =
      "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;" +
      "$OutputEncoding=[System.Text.Encoding]::UTF8;" +
      command +
      '; if ($LASTEXITCODE -is [int]) { exit $LASTEXITCODE }';
    child = spawn(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
      { cwd, windowsHide: true }
    );
  } else {
    child = spawn("cmd.exe", ["/d", "/s", "/c", `chcp 65001>nul&${command}`], { cwd, windowsHide: true });
  }

  let stdout = "";
  let stderr = "";
  const endOut = collectStream(child.stdout, (t) => {
    if (stdout.length < MAX_STREAM_CHARS) stdout += t;
  });
  const endErr = collectStream(child.stderr, (t) => {
    if (stderr.length < MAX_STREAM_CHARS) stderr += t;
  });

  const done = new Promise((resolve) => {
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      killTree(child.pid);
      setTimeout(() => child.kill("SIGKILL"), 3000); // taskkill 失效时兜底
    }, timeoutMs);
    child.on("error", (e) => {
      clearTimeout(timer);
      resolve({ spawnError: e?.message || String(e), code: null, timedOut });
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      resolve({ code, timedOut });
    });
  });

  const { code, timedOut, spawnError } = await done;
  endOut();
  endErr();
  return {
    command,
    shell,
    cwd,
    exitCode: spawnError || timedOut ? null : code, // 超时被杀时退出码无意义
    timedOut,
    spawnError: spawnError || undefined,
    durationMs: Date.now() - t0,
    stdout: capStream(stdout),
    stderr: capStream(stderr),
  };
}
