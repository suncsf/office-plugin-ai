/**
 * 工具函数：直接复用现代版实现（markdown-it 渲染、格式化等均为纯 JS，
 * 兼容版构建经 babel 转译后可在 IE11 运行）。
 */
export { renderMarkdown, formatTokens, formatTime, truncate } from "../src/utils";
