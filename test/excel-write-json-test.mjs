/**
 * Excel JSON 写表离线自测（不需要真实 Office / LLM）
 *   node test/excel-write-json-test.mjs
 * 覆盖：jsonToGrid（对象数组自动表头/二维数组/headers/嵌套值/异常输入）、
 *       normalize2D 字符串 JSON 自动解析（修复"整段 JSON 塞进单元格"问题）。
 */
import { jsonToGrid, normalize2D } from "../src/office/excel.js";

let failed = false;
function check(name, cond, detail = "") {
  if (cond) console.log(`  ✅ ${name}`);
  else {
    failed = true;
    console.error(`  ❌ ${name} ${detail}`);
  }
}
function eq(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  check(name, a === e, `got ${a}, want ${e}`);
}
const throws = (fn) => {
  try { fn(); return false; } catch { return true; }
};

console.log("== jsonToGrid ==");

/* 对象数组 → 自动表头（键名，首次出现顺序） */
const sales = [
  { 日期: "2026-10-01", 产品: "笔记本电脑", 数量: 5, 单价: 5999 },
  { 日期: "2026-10-03", 产品: "无线鼠标", 数量: 30, 单价: 89 },
];
eq("对象数组 → 表头+数据行", jsonToGrid(sales), [
  ["日期", "产品", "数量", "单价"],
  ["2026-10-01", "笔记本电脑", 5, 5999],
  ["2026-10-03", "无线鼠标", 30, 89],
]);

/* 后续行多出的键 → 并集，缺失补 "" */
eq("键并集+缺失补空", jsonToGrid([{ a: 1 }, { b: 2 }]), [["a", "b"], [1, ""], ["", 2]]);

/* headers 指定列顺序/子集 */
eq("headers 重排列", jsonToGrid(sales, ["产品", "日期"]), [
  ["产品", "日期"],
  ["笔记本电脑", "2026-10-01"],
  ["无线鼠标", "2026-10-03"],
]);

/* headers 逗号分隔字符串 */
eq("headers 逗号字符串", jsonToGrid(sales, "产品,数量")[0], ["产品", "数量"]);

/* headers JSON 字符串 */
eq("headers JSON 字符串", jsonToGrid(sales, '["数量","产品"]')[0], ["数量", "产品"]);

/* 二维数组 → 原样；headers 补表头 */
eq("二维数组原样", jsonToGrid([[1, 2], [3, 4]]), [[1, 2], [3, 4]]);
eq("二维数组+表头", jsonToGrid([[1, 2], [3, 4]], "x,y"), [["x", "y"], [1, 2], [3, 4]]);
eq("二维数组+表头补齐列宽", jsonToGrid([[1, 2, 3]], "x,y"), [["x", "y", ""], [1, 2, 3]]);
eq("空字符串 headers 视为不补表头", jsonToGrid([[1]], ""), [[1]]);

/* 字符串形态 JSON 自动解析（模型把数组序列化成字符串传入） */
eq("字符串 JSON 解析", jsonToGrid('[{"名称":"A","数量":1},{"名称":"B","数量":2}]'), [
  ["名称", "数量"],
  ["A", 1],
  ["B", 2],
]);

/* 单个对象 → 包装成一行 */
eq("单对象包装", jsonToGrid({ a: 1, b: "x" }), [["a", "b"], [1, "x"]]);

/* 嵌套对象/数组 → JSON 字符串；null/undefined → "" */
eq("嵌套值序列化", jsonToGrid([{ a: { x: 1 }, b: [1, 2], c: null, d: undefined }]), [
  ["a", "b", "c", "d"],
  ['{"x":1}', "[1,2]", "", ""],
]);

/* 非法输入 */
check("空数组报错", throws(() => jsonToGrid([])));
check("非数组报错", throws(() => jsonToGrid("abc")));
check("对象/数组混用报错", throws(() => jsonToGrid([{ a: 1 }, [1, 2]])));
check("无效 headers 报错", throws(() => jsonToGrid([[1]], {})));
check("无效 JSON 字符串报错", throws(() => jsonToGrid("[[[")));

console.log("== normalize2D（字符串 JSON 自动解析）==");

/* 模拟截图 bug：模型把二维数组序列化成字符串传入 → 之前被当标量填满每格，现在应正确展开 */
const strJson = JSON.stringify([["日期", "产品"], ["2026-10-01", "笔记本"]]);
eq("字符串二维数组 → 正确展开", normalize2D(strJson, 2, 2), [["日期", "产品"], ["2026-10-01", "笔记本"]]);
check("字符串尺寸不符仍报错", throws(() => normalize2D(JSON.stringify([["a"]]), 2, 2)));
eq("非 JSON 字符串仍作标量填充", normalize2D("备注", 2, 2), [["备注", "备注"], ["备注", "备注"]]);
eq("真实二维数组不受影响", normalize2D([[1, 2]], 1, 2), [[1, 2]]);
eq("单值填充不受影响", normalize2D(5, 2, 2), [[5, 5], [5, 5]]);
check("null 仍报错", throws(() => normalize2D(null, 2, 2)));

console.log(failed ? "\n❌ 存在失败断言" : "\n✅ 全部通过");
process.exit(failed ? 1 : 0);
