/**
 * Word 内置样式与对齐方式的解析映射
 *
 * ⚠️ 重要（本模块的存在原因）：
 * 1. Word.BuiltInStyleName 枚举值【没有空格】：heading1 = "Heading1"（不是 "Heading 1"）。
 *    官方文档: https://learn.microsoft.com/javascript/api/word/word.builtinstylename
 * 2. `style` 属性只接受【本地化显示名 / 自定义样式名】——在中文版 Word 里
 *    赋 "Heading 1" 或 "Normal" 会抛 InvalidArgument（显示名分别是 "标题 1"、"正文"）。
 * 3. 跨语言可移植的内置样式必须用 `styleBuiltIn` 属性（赋枚举值，无空格）。
 *
 * 因此：applyStyle 的主路径是 styleBuiltIn = 枚举值；style 仅用于自定义样式与兜底。
 */

/* 官方枚举值 + 常见别名（匹配时忽略大小写与空白）+ 本地化显示名（仅兜底用） */
export const BUILT_IN_STYLES = [
  { value: "Normal", zh: "正文", en: "Normal", keys: ["normal", "正文", "默认", "default"] },
  { value: "Title", zh: "标题", en: "Title", keys: ["title", "标题"] },
  { value: "Subtitle", zh: "副标题", en: "Subtitle", keys: ["subtitle", "副标题"] },
  { value: "Heading1", zh: "标题 1", en: "Heading 1", keys: ["heading1", "h1", "标题1", "一级标题"] },
  { value: "Heading2", zh: "标题 2", en: "Heading 2", keys: ["heading2", "h2", "标题2", "二级标题"] },
  { value: "Heading3", zh: "标题 3", en: "Heading 3", keys: ["heading3", "h3", "标题3", "三级标题"] },
  { value: "Heading4", zh: "标题 4", en: "Heading 4", keys: ["heading4", "h4", "标题4"] },
  { value: "Heading5", zh: "标题 5", en: "Heading 5", keys: ["heading5", "h5", "标题5"] },
  { value: "Heading6", zh: "标题 6", en: "Heading 6", keys: ["heading6", "h6", "标题6"] },
  { value: "Heading7", zh: "标题 7", en: "Heading 7", keys: ["heading7", "h7", "标题7"] },
  { value: "Heading8", zh: "标题 8", en: "Heading 8", keys: ["heading8", "h8", "标题8"] },
  { value: "Heading9", zh: "标题 9", en: "Heading 9", keys: ["heading9", "h9", "标题9"] },
  { value: "Quote", zh: "引用", en: "Quote", keys: ["quote", "引用"] },
  { value: "IntenseQuote", zh: "明显引用", en: "Intense Quote", keys: ["intensequote", "明显引用"] },
  { value: "ListParagraph", zh: "列表段落", en: "List Paragraph", keys: ["listparagraph", "列表段落"] },
  { value: "Caption", zh: "题注", en: "Caption", keys: ["caption", "题注"] },
  { value: "NoSpacing", zh: "无间隔", en: "No Spacing", keys: ["nospacing", "无间隔"] },
  { value: "Strong", zh: "加强", en: "Strong", keys: ["strong", "加强"] },
  { value: "Emphasis", zh: "强调", en: "Emphasis", keys: ["emphasis", "强调"] },
  { value: "SubtleEmphasis", zh: "微妙强调", en: "Subtle Emphasis", keys: ["subtleemphasis", "微妙强调"] },
  { value: "IntenseEmphasis", zh: "明显强调", en: "Intense Emphasis", keys: ["intenseemphasis", "明显强调"] },
  { value: "SubtleReference", zh: "微妙参考", en: "Subtle Reference", keys: ["subtlereference"] },
  { value: "IntenseReference", zh: "明显参考", en: "Intense Reference", keys: ["intensereference", "明显参考"] },
  { value: "BookTitle", zh: "书籍标题", en: "Book Title", keys: ["booktitle", "书籍标题"] },
  { value: "Bibliography", zh: "书目", en: "Bibliography", keys: ["bibliography", "书目"] },
  { value: "TocHeading", zh: "目录标题", en: "TOC Heading", keys: ["tocheading", "目录标题"] },
  { value: "TableGrid", zh: "网格型", en: "Table Grid", keys: ["tablegrid", "网格型"] },
  { value: "Hyperlink", zh: "超链接", en: "Hyperlink", keys: ["hyperlink", "超链接"] },
];

const LOOKUP = new Map();
for (const s of BUILT_IN_STYLES) {
  LOOKUP.set(s.value.toLowerCase(), s);
  for (const k of s.keys) LOOKUP.set(k, s);
}

export function normalizeStyleKey(input) {
  return String(input || "").toLowerCase().replace(/\s+/g, "");
}

/**
 * 解析样式请求
 * @returns {{ builtin?: object, custom?: string } | null}
 *  builtin 存在时形如 { value, zh, en, keys }；否则 custom 为原样输入（自定义样式名）
 */
export function resolveStyleRequest(input) {
  const raw = String(input || "").trim();
  if (!raw) return null;
  const hit = LOOKUP.get(normalizeStyleKey(raw));
  if (hit) return { builtin: hit };
  return { custom: raw };
}

/** Word.Alignment 合法值: Left / Centered / Right / Justified（宽容映射常见变体） */
const ALIGN_MAP = {
  left: "Left", "左": "Left", "左对齐": "Left",
  center: "Centered", centered: "Centered", middle: "Centered", "居中": "Centered",
  right: "Right", "右": "Right", "右对齐": "Right",
  justify: "Justified", justified: "Justified", both: "Justified", "两端对齐": "Justified",
};

export function normalizeAlignment(input) {
  return ALIGN_MAP[normalizeStyleKey(input)] || null;
}
