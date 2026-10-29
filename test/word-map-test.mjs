/**
 * Word 样式/对齐映射离线单测（不依赖 Word）
 *   node test/word-map-test.mjs
 * 背景：中文版 Word 中 style 属性只接受本地化名，内置样式必须用 styleBuiltIn 枚举值（无空格）。
 */
import { resolveStyleRequest, normalizeAlignment, BUILT_IN_STYLES } from "../src/office/style-map.js";

let failed = false;
function check(name, cond, detail = "") {
  if (cond) console.log(`  ✅ ${name}`);
  else {
    failed = true;
    console.error(`  ❌ ${name} ${detail}`);
  }
}

console.log("\n[样式解析] 用户此前全部失败的输入");
// 之前：英文/中文/紧凑写法全部 InvalidArgument
const cases = [
  ["Title", "Title"],
  ["Heading 1", "Heading1"],
  ["Heading 2", "Heading2"],
  ["Normal", "Normal"],
  ["标题", "Title"],
  ["标题 1", "Heading1"],
  ["标题 2", "Heading2"],
  ["正文", "Normal"],
  ["Heading1", "Heading1"],
  ["标题1", "Heading1"],
  ["h1", "Heading1"],
  ["H2", "Heading2"],
  ["heading 6", "Heading6"],
  ["Quote", "Quote"],
  ["引用", "Quote"],
  ["List Paragraph", "ListParagraph"],
  ["列表段落", "ListParagraph"],
  ["No Spacing", "NoSpacing"],
  ["无间隔", "NoSpacing"],
];
for (const [input, expected] of cases) {
  const r = resolveStyleRequest(input);
  check(`"${input}" → ${expected}`, r?.builtin?.value === expected, JSON.stringify(r));
}

console.log("\n[样式解析] 自定义样式与边界");
{
  const r = resolveStyleRequest("我的自定义样式");
  check("未知名称走自定义路径", r?.custom === "我的自定义样式", JSON.stringify(r));
  check("空输入返回 null", resolveStyleRequest("") === null && resolveStyleRequest(null) === null);
  const r2 = resolveStyleRequest("  Heading  1 ");
  check("多余空白容忍", r2?.builtin?.value === "Heading1", JSON.stringify(r2));
}

console.log("\n[枚举值合法性] 与官方 BuiltInStyleName 枚举值逐一比对");
// 来源: https://learn.microsoft.com/javascript/api/word/word.builtinstylename
const OFFICIAL_VALUES = new Set([
  "Bibliography", "BookTitle", "Caption", "Emphasis", "EndnoteReference", "EndnoteText", "Footer",
  "FootnoteReference", "FootnoteText", "Header", "Heading1", "Heading2", "Heading3", "Heading4",
  "Heading5", "Heading6", "Heading7", "Heading8", "Heading9", "Hyperlink", "IntenseEmphasis",
  "IntenseQuote", "IntenseReference", "ListParagraph", "Normal", "NoSpacing", "Other", "Quote",
  "Strong", "Subtitle", "SubtleEmphasis", "SubtleReference", "Title", "TocHeading", "TableGrid",
]);
check("表中所有值都在官方枚举集合内", BUILT_IN_STYLES.every((s) => OFFICIAL_VALUES.has(s.value)),
  BUILT_IN_STYLES.filter((s) => !OFFICIAL_VALUES.has(s.value)).map((s) => s.value).join(","));

console.log("\n[对齐映射]");
for (const [input, expected] of [
  ["Left", "Left"], ["Center", "Centered"], ["Centered", "Centered"], ["居中", "Centered"],
  ["Right", "Right"], ["Justified", "Justified"], ["Justify", "Justified"], ["两端对齐", "Justified"],
  ["  center ", "Centered"],
]) {
  check(`"${input}" → ${expected}`, normalizeAlignment(input) === expected, String(normalizeAlignment(input)));
}
check("非法对齐返回 null", normalizeAlignment("diagonal") === null);

console.log("\n[工具目录完整性]");
{
  const { TOOLS } = await import("../shared/tools-catalog.js");
  const names = TOOLS.map((t) => t.name);
  check("工具名全局唯一", new Set(names).size === names.length, `共 ${names.length} 个`);
  check("每个工具都有 host/executor/description/parameters", TOOLS.every((t) => t.host && t.executor && t.description && t.parameters?.type === "object"));
  check("host 取值合法", TOOLS.every((t) => ["word", "excel", "ppt", "common"].includes(t.host)));
  check("包含 word_list_styles（全局样式库读取）", names.includes("word_list_styles"));
  check("包含 word_create_style（自定义样式创建）", names.includes("word_create_style"));
  const wordCount = TOOLS.filter((t) => t.host === "word").length;
  check("Word 工具数 ≥ 26", wordCount >= 26, String(wordCount));
}

console.log(failed ? "\n❌ 存在失败项" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
