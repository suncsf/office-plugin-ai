#!/usr/bin/env node
/**
 * 组装安装包源目录：installer/staging/app
 *  - 复制运行所需文件（server/shared/dist/scripts/manifest/package.json）
 *  - 生产依赖安装（含 office-addin-dev-certs，供安装后生成开发证书）
 *  - 内置 Node 运行时（runtime/node.exe）
 *  - 编译可视化控制台（系统 csc → OfficeAIConsole.exe）
 */
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const STAGE = path.join(ROOT, "installer", "staging");
const APP = path.join(STAGE, "app");

const log = (m) => console.log("  " + m);
const rm = (p) => fs.rmSync(p, { recursive: true, force: true });

console.log("[1/5] 清理并创建 staging");
rm(STAGE);
fs.mkdirSync(APP, { recursive: true });

console.log("[2/5] 复制应用文件");
for (const item of ["server", "shared", "dist", "scripts", "manifest.xml", "package.json", "package-lock.json", "README.md"]) {
  fs.cpSync(path.join(ROOT, item), path.join(APP, item), { recursive: true, force: true });
  log(item);
}

console.log("[3/5] 生产依赖（npm install --omit=dev，含 office-addin-dev-certs）");
{
  const pkg = JSON.parse(fs.readFileSync(path.join(APP, "package.json"), "utf8"));
  pkg.dependencies = { ...(pkg.dependencies || {}), "office-addin-dev-certs": "^3.0.1" };
  delete pkg.devDependencies;
  delete pkg.sideload;
  fs.writeFileSync(path.join(APP, "package.json"), JSON.stringify(pkg, null, 2));
  execFileSync("cmd.exe", ["/c", "npm", "install", "--omit=dev", "--no-audit", "--no-fund", "--loglevel=error"], {
    cwd: APP, stdio: "inherit", shell: false,
  });
}

console.log("[4/5] 内置 Node 运行时");
{
  const nodeDir = path.dirname(process.execPath);
  fs.mkdirSync(path.join(APP, "runtime"), { recursive: true });
  fs.copyFileSync(path.join(nodeDir, "node.exe"), path.join(APP, "runtime", "node.exe"));
  log("node.exe (" + Math.round(fs.statSync(path.join(APP, "runtime", "node.exe")).size / 1048576) + " MB)");
}

console.log("[5/5] 编译可视化控制台（csc.exe）");
{
  const ico = path.join(ROOT, "launcher", "app.ico");
  if (!fs.existsSync(ico)) throw new Error("缺少 launcher/app.ico，请先运行 node scripts/gen-ico.mjs");
  fs.copyFileSync(ico, path.join(APP, "app.ico"));
  const csc = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\csc.exe";
  execFileSync(csc, [
    "/nologo", "/target:winexe", "/platform:anycpu",
    "/win32icon:" + ico,
    "/r:System.Windows.Forms.dll", "/r:System.Drawing.dll",
    `/out:${path.join(APP, "OfficeAIConsole.exe")}`,
    path.join(ROOT, "launcher", "OfficeAIConsole.cs"),
  ], { stdio: "inherit" });
  log("OfficeAIConsole.exe（含 app.ico）");
}

let total = 0;
const walk = (dir) => {
  for (const f of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, f.name);
    if (f.isDirectory()) walk(p);
    else total += fs.statSync(p).size;
  }
};
walk(APP);
console.log(`\n✅ staging 完成：${APP}`);
console.log(`   总大小约 ${Math.round(total / 1048576)} MB`);
