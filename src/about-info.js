// 本地隐私信息（开发人员/部门）：不入库。
// 复制同目录新增 about-info.local.js 填写真实信息（该文件名已被 .gitignore 排除）；
// 未配置时「关于」页自动隐藏对应行。
const localModules = import.meta.glob("./about-info.local.js", { eager: true });
const local = localModules["./about-info.local.js"] || {};
export const developer = local.developer || "";
export const department = local.department || "";
