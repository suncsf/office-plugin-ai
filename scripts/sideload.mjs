#!/usr/bin/env node
/**
 * 加载项旁加载注册（持久化，替代 office-addin-debugging 的临时文档流程）
 *
 * 原理（与 office-addin-dev-settings registerAddIn 完全一致）：
 *   在 HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer 下写入
 *   字符串值：值名 = manifest <Id>，数据 = manifest.xml 绝对路径。
 *   该注册持久有效，Word/Excel/PPT 启动时自动加载加载项，
 *   之后打开任何自己的文档，都可以在【开始】选项卡点击 "AI 助手" 打开面板。
 *
 * 用法：
 *   node scripts/sideload.mjs register     注册（一次性，持久有效）
 *   node scripts/sideload.mjs unregister   取消注册
 *   node scripts/sideload.mjs status       查看当前注册
 *   node scripts/sideload.mjs launch word|excel|ppt   启动对应 Office 应用
 */
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const MANIFEST = resolve(ROOT, "manifest.xml");
const REG_KEY = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer";

function manifestId() {
  const xml = readFileSync(MANIFEST, "utf8");
  const m = xml.match(/<Id>([^<]+)<\/Id>/i);
  if (!m) throw new Error("manifest.xml 中未找到 <Id>");
  return m[1].trim();
}

function reg(args) {
  return execFileSync("reg.exe", args, { encoding: "utf8" });
}

function register() {
  const id = manifestId();
  reg(["add", REG_KEY, "/v", id, "/t", "REG_SZ", "/d", MANIFEST, "/f"]);
  console.log("✅ 已注册（持久有效，重启 Office 后自动加载）");
  console.log(`   加载项 Id: ${id}`);
  console.log(`   manifest : ${MANIFEST}`);
  console.log("");
  console.log("接下来：");
  console.log("  1. 完全关闭所有 Word / Excel / PowerPoint 窗口");
  console.log("  2. 重新打开您的文档 → 【开始】选项卡 → 点击 “AI 助手” 按钮");
  console.log("     （服务需在运行中：npm start）");
}

function unregister() {
  const id = manifestId();
  try {
    reg(["delete", REG_KEY, "/v", id, "/f"]);
    console.log("✅ 已取消注册");
  } catch {
    console.log("ℹ️ 未找到注册项（本来就未注册）");
  }
}

function status() {
  try {
    const out = reg(["query", REG_KEY]);
    const id = manifestId();
    const lines = out.split(/\r?\n/).filter((l) => l.includes("REG_SZ"));
    const mine = lines.find((l) => l.includes(id));
    console.log(`注册表位置: ${REG_KEY}`);
    if (mine) {
      const data = mine.split(/\s{2,}/).pop();
      console.log(`✅ 本加载项已注册 → ${data}`);
      console.log(`   数据与当前 manifest ${data === MANIFEST ? "一致" : "不一致（建议重新 register）"}`);
    } else {
      console.log("❌ 本加载项未注册（运行 npm run register）");
    }
    console.log("");
    console.log("其它已注册的加载项:");
    console.log(lines.filter((l) => !l.includes(id)).join("\n") || "  （无）");
  } catch {
    console.log(`❌ 注册表键不存在（从未旁加载过任何加载项）: ${REG_KEY}`);
  }
}

function launch(app) {
  const exe = { word: "WINWORD.EXE", excel: "EXCEL.EXE", ppt: "POWERPNT.EXE" }[app];
  if (!exe) {
    console.error("用法: node scripts/sideload.mjs launch word|excel|ppt");
    process.exit(1);
  }
  execFileSync("cmd.exe", ["/c", "start", "", exe], { stdio: "ignore" });
  console.log(`✅ 已启动 ${exe}`);
}

const [cmd, arg] = process.argv.slice(2);
try {
  if (cmd === "register") register();
  else if (cmd === "unregister") unregister();
  else if (cmd === "status") status();
  else if (cmd === "launch") launch(arg);
  else {
    console.log(__doc__);
    process.exit(cmd ? 1 : 0);
  }
} catch (e) {
  console.error(`❌ 失败: ${e.message}`);
  process.exit(1);
}
