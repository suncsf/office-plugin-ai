/**
 * Excel 执行器（Office.js Excel API）
 */
import { stripDataUrl } from "./dataurl";

/* ---------------- 内部工具函数 ---------------- */

function getSheet(ctx, name) {
  return name ? ctx.workbook.worksheets.getItem(String(name)) : ctx.workbook.worksheets.getActiveWorksheet();
}

/** "B3" → {row:2, col:1}（0 起始） */
export function parseCellRef(ref) {
  const m = /^([A-Za-z]{1,3})(\d+)$/.exec(String(ref).replace(/\$/g, ""));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10) - 1, col: col - 1 };
}

function rangeStartCell(address) {
  const a = String(address).split("!").pop();
  const first = a.split(":")[0];
  return parseCellRef(first);
}

function colName(c) {
  let s = "";
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function normalizeColor(c) {
  const s = String(c || "").trim();
  if (!s) return s;
  return s.startsWith("#") || s.startsWith("rgb") ? s : `#${s}`;
}

/** 把任意输入规整为 rows×cols 的二维数组 */
function normalize2D(values, rows, cols) {
  const need = `${rows} 行 × ${cols} 列`;
  if (values === null || values === undefined) throw new Error("values 不能为空");
  if (!Array.isArray(values)) {
    return new Array(rows).fill(null).map(() => new Array(cols).fill(values));
  }
  if (values.length && Array.isArray(values[0])) {
    if (values.length !== rows || values[0].length !== cols) {
      throw new Error(`values 是二维数组，但尺寸 (${values.length}×${values[0].length}) 与区域 (${need}) 不一致，请调整地址或数组`);
    }
    return values.map((r) => r.map((v) => (v === undefined || v === null ? "" : v)));
  }
  // 一维 → 单行
  if (rows === 1) {
    if (values.length !== cols) throw new Error(`一维数组长度 (${values.length}) 与区域列数 (${cols}) 不一致`);
    return [values.map((v) => (v === undefined || v === null ? "" : v))];
  }
  if (values.length === cols) {
    return new Array(rows).fill(null).map(() => values.map((v) => (v === undefined || v === null ? "" : v)));
  }
  throw new Error(`一维数组长度 (${values.length}) 与区域 (${need}) 不一致，请使用与区域等尺寸的二维数组`);
}

function slicePreview(values, maxRows = 30, maxCols = 30) {
  const rows = values.length;
  const cols = rows ? values[0].length : 0;
  const r = Math.min(rows, maxRows);
  const c = Math.min(cols, maxCols);
  const out = values.slice(0, r).map((row) => row.slice(0, c).map((v) => (v == null ? "" : v)));
  const note = rows > r || cols > c ? `（预览 ${r} 行 ${c} 列，实际 ${rows} 行 ${cols} 列）` : "";
  return { values: out, note };
}

function cellValueString(v) {
  if (v == null) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

/* ---------------- 执行器 ---------------- */

const registry = {};

registry["excel.getWorkbookInfo"] = async () => {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("name,position,visibility");
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    const sel = ctx.workbook.getSelectedRange();
    sel.load("address");
    const used = active.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,isNullObject");
    await ctx.sync();
    return {
      activeSheet: active.name,
      selectedRange: sel.address,
      usedRange: used.isNullObject ? null : { address: used.address, rows: used.rowCount, cols: used.columnCount },
      sheets: sheets.items
        .slice()
        .sort((a, b) => a.position - b.position)
        .map((s) => ({ name: s.name, position: s.position, visibility: s.visibility })),
    };
  });
};

registry["excel.activateSheet"] = async (args) => {
  const name = String(args?.name || "");
  if (!name) throw new Error("name 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(name);
    sheet.load("name");
    await ctx.sync();
    sheet.activate();
    await ctx.sync();
    return { ok: true, activeSheet: name };
  });
};

registry["excel.addSheet"] = async (args) => {
  return Excel.run(async (ctx) => {
    let name = args?.name ? String(args.name) : null;
    if (name) {
      const existing = ctx.workbook.worksheets.getItemOrNullObject(name);
      existing.load("isNullObject");
      await ctx.sync();
      let i = 2;
      while (!existing.isNullObject) {
        name = `${String(args.name)}${i}`;
        const again = ctx.workbook.worksheets.getItemOrNullObject(name);
        again.load("isNullObject");
        await ctx.sync();
        if (again.isNullObject) break;
        i++;
      }
    }
    const sheet = name ? ctx.workbook.worksheets.add(name) : ctx.workbook.worksheets.add();
    sheet.load("name");
    await ctx.sync();
    sheet.activate();
    await ctx.sync();
    return { ok: true, name: sheet.name };
  });
};

registry["excel.renameSheet"] = async (args) => {
  const oldName = String(args?.old_name || "");
  const newName = String(args?.new_name || "");
  if (!oldName || !newName) throw new Error("old_name / new_name 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(oldName);
    sheet.load("name");
    await ctx.sync();
    sheet.name = newName;
    await ctx.sync();
    return { ok: true, oldName, newName };
  });
};

registry["excel.deleteSheet"] = async (args) => {
  const name = String(args?.name || "");
  if (!name) throw new Error("name 不能为空");
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    if (sheets.items.length <= 1) throw new Error("工作簿至少要保留一个工作表");
    const sheet = ctx.workbook.worksheets.getItem(name);
    sheet.load("name");
    await ctx.sync();
    sheet.delete();
    await ctx.sync();
    return { ok: true, deleted: name };
  });
};

registry["excel.getUsedRange"] = async (args) => {
  const previewRows = Math.min(200, Math.max(1, Number(args?.preview_rows) || 20));
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,isNullObject,values");
    await ctx.sync();
    if (used.isNullObject) return { address: null, note: "该工作表为空" };
    const { values, note } = slicePreview(used.values, previewRows, 40);
    return {
      address: used.address,
      rows: used.rowCount,
      cols: used.columnCount,
      preview: values,
      note: `已截取前 ${previewRows} 行${note ? " " + note : ""}`,
    };
  });
};

registry["excel.getRange"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.load("address,rowCount,columnCount,values,numberFormat");
    await ctx.sync();
    const { values, note } = slicePreview(rng.values, 100, 40);
    return { address: rng.address, rows: rng.rowCount, cols: rng.columnCount, numberFormatSample: rng.numberFormat?.[0]?.[0], values, note };
  });
};

registry["excel.getSelectedRange"] = async () => {
  return Excel.run(async (ctx) => {
    const rng = ctx.workbook.getSelectedRange();
    rng.load("address,rowCount,columnCount,values");
    await ctx.sync();
    const { values } = slicePreview(rng.values, 30, 30);
    return { address: rng.address, rows: rng.rowCount, cols: rng.columnCount, preview: values };
  });
};

registry["excel.setRangeValues"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.load("rowCount,columnCount");
    await ctx.sync();
    const values = normalize2D(args?.values, rng.rowCount, rng.columnCount);
    rng.values = values;
    await ctx.sync();
    return { ok: true, address, rows: rng.rowCount, cols: rng.columnCount };
  });
};

registry["excel.setFormula"] = async (args) => {
  const address = String(args?.address || "");
  const formula = String(args?.formula || "");
  if (!address || !formula) throw new Error("address / formula 不能为空");
  if (!formula.startsWith("=")) throw new Error("公式必须以 = 开头");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.formula = formula;
    rng.load("values");
    await ctx.sync();
    return { ok: true, address, formula, calculatedValue: cellValueString(rng.values?.[0]?.[0]) };
  });
};

registry["excel.getFormula"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.load("formula,values");
    await ctx.sync();
    return { address, formula: rng.formula, value: cellValueString(rng.values?.[0]?.[0]) };
  });
};

registry["excel.insertRows"] = async (args) => {
  const index = Number(args?.index);
  const count = Math.max(1, Number(args?.count) || 1);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    sheet.getRangeByIndexes(index, 0, count, 16384).insert("Down");
    await ctx.sync();
    return { ok: true, insertedRows: count, atIndex: index };
  });
};

registry["excel.deleteRows"] = async (args) => {
  const index = Number(args?.index);
  const count = Math.max(1, Number(args?.count) || 1);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    sheet.getRangeByIndexes(index, 0, count, 16384).delete("Up");
    await ctx.sync();
    return { ok: true, deletedRows: count, fromIndex: index };
  });
};

registry["excel.insertColumns"] = async (args) => {
  const index = Number(args?.index);
  const count = Math.max(1, Number(args?.count) || 1);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    sheet.getRangeByIndexes(0, index, 1048576, count).insert("Right");
    await ctx.sync();
    return { ok: true, insertedCols: count, atIndex: index };
  });
};

registry["excel.deleteColumns"] = async (args) => {
  const index = Number(args?.index);
  const count = Math.max(1, Number(args?.count) || 1);
  if (!Number.isInteger(index) || index < 0) throw new Error("index 必须是非负整数");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    sheet.getRangeByIndexes(0, index, 1048576, count).delete("Left");
    await ctx.sync();
    return { ok: true, deletedCols: count, fromIndex: index };
  });
};

registry["excel.formatRange"] = async (args) => {
  const address = String(args?.address || "");
  let fmt = args?.format;
  if (typeof fmt === "string") {
    try {
      fmt = JSON.parse(fmt);
    } catch {
      throw new Error("format 必须是 JSON 对象");
    }
  }
  if (!address) throw new Error("address 不能为空");
  if (!fmt || typeof fmt !== "object") throw new Error("缺少 format");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.load("rowCount,columnCount");
    await ctx.sync();
    const f = rng.format;
    if (fmt.font_name) f.font.name = String(fmt.font_name);
    if (fmt.font_size != null) f.font.size = Number(fmt.font_size);
    if (fmt.bold != null) f.font.bold = !!fmt.bold;
    if (fmt.italic != null) f.font.italic = !!fmt.italic;
    if (fmt.underline != null) f.font.underline = fmt.underline ? "Single" : "None";
    if (fmt.font_color) f.font.color = normalizeColor(fmt.font_color);
    if (fmt.fill_color) {
      f.fill.color = normalizeColor(fmt.fill_color);
    }
    if (fmt.horizontal_alignment) f.horizontalAlignment = String(fmt.horizontal_alignment);
    if (fmt.vertical_alignment) f.verticalAlignment = String(fmt.vertical_alignment);
    if (fmt.wrap_text != null) f.wrapText = !!fmt.wrap_text;
    if (fmt.number_format) {
      const nf = String(fmt.number_format);
      rng.numberFormat = new Array(rng.rowCount).fill(null).map(() => new Array(rng.columnCount).fill(nf));
    }
    if (fmt.border_style || fmt.border_color) {
      const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
      for (const e of edges) {
        const b = f.borders.getItem(e);
        if (fmt.border_style) b.style = String(fmt.border_style);
        if (fmt.border_color) b.color = normalizeColor(fmt.border_color);
      }
    }
    await ctx.sync();
    return { ok: true, address, applied: Object.keys(fmt) };
  });
};

registry["excel.autoFit"] = async (args) => {
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = args?.address ? sheet.getRange(String(args.address)) : sheet.getUsedRangeOrNullObject(true);
    rng.load("isNullObject,address");
    await ctx.sync();
    if (rng.isNullObject) return { ok: true, note: "工作表为空" };
    rng.format.autofitColumns();
    rng.format.autofitRows();
    await ctx.sync();
    return { ok: true, address: rng.address };
  });
};

registry["excel.setColumnWidth"] = async (args) => {
  const width = Number(args?.width);
  if (!Number.isFinite(width)) throw new Error("width 必须是数字（磅）");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const colArg = String(args?.column ?? "").trim();
    let colIdx;
    if (/^\d+$/.test(colArg)) {
      colIdx = Number(colArg);
    } else {
      const ref = parseCellRef(/\d+$/.test(colArg) ? colArg : colArg + "1");
      colIdx = ref?.col;
    }
    if (colIdx == null || colIdx < 0) throw new Error("column 必须是 0 起始的数字或字母（如 A）");
    sheet.getRangeByIndexes(0, colIdx, 1, 1).format.columnWidth = width;
    await ctx.sync();
    return { ok: true, column: colName(colIdx), width };
  });
};

registry["excel.setRowHeight"] = async (args) => {
  const row = Number(args?.row);
  const height = Number(args?.height);
  if (!Number.isInteger(row) || row < 0) throw new Error("row 必须是非负整数");
  if (!Number.isFinite(height)) throw new Error("height 必须是数字（磅）");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    sheet.getRangeByIndexes(row, 0, 1, 1).format.rowHeight = height;
    await ctx.sync();
    return { ok: true, row, height };
  });
};

registry["excel.sortRange"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.sort.apply(
      [{ key: Math.max(0, Number(args?.by_column) || 0), ascending: args?.ascending !== false }],
      { matchCase: false, hasHeaders: !!args?.has_headers, orientation: "Rows" }
    );
    await ctx.sync();
    return { ok: true, address };
  });
};

registry["excel.createTable"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const table = sheet.tables.add(address, args?.has_headers !== false);
    if (args?.name) table.name = String(args.name);
    table.load("name,address");
    await ctx.sync();
    return { ok: true, name: table.name, address: table.address };
  });
};

registry["excel.findCells"] = async (args) => {
  const text = String(args?.text || "");
  if (!text) throw new Error("text 不能为空");
  const matchCase = !!args?.match_case;
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("isNullObject,address,rowCount,columnCount,values");
    await ctx.sync();
    if (used.isNullObject) return { matches: [], note: "工作表为空" };
    const start = rangeStartCell(used.address) || { row: 0, col: 0 };
    const needle = matchCase ? text : text.toLowerCase();
    const matches = [];
    for (let r = 0; r < used.rowCount && matches.length < 50; r++) {
      const row = used.values[r] || [];
      for (let c = 0; c < used.columnCount && matches.length < 50; c++) {
        const cellStr = cellValueString(row[c]);
        const hay = matchCase ? cellStr : cellStr.toLowerCase();
        if (hay.includes(needle)) {
          matches.push({ address: `${colName(start.col + c)}${start.row + r + 1}`, value: cellStr.slice(0, 100) });
        }
      }
    }
    return { searchedRange: used.address, matchCount: matches.length, matches, note: matches.length >= 50 ? "仅显示前 50 个" : undefined };
  });
};

registry["excel.clearRange"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  const mode = ["all", "contents", "formats"].includes(args?.mode) ? args.mode : "all";
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    rng.clear(mode);
    await ctx.sync();
    return { ok: true, address, mode };
  });
};

registry["excel.mergeCells"] = async (args) => {
  const address = String(args?.address || "");
  if (!address) throw new Error("address 不能为空");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const rng = sheet.getRange(address);
    if (args?.unmerge) rng.unmerge();
    else rng.merge(false);
    await ctx.sync();
    return { ok: true, address, unmerged: !!args?.unmerge };
  });
};

registry["excel.addChart"] = async (args) => {
  const dataAddress = String(args?.data_address || "");
  if (!dataAddress) throw new Error("data_address 不能为空");
  const chartType = String(args?.chart_type || "ColumnClustered");
  return Excel.run(async (ctx) => {
    const sheet = getSheet(ctx, args?.sheet);
    const chart = sheet.charts.add(chartType, sheet.getRange(dataAddress), "Auto");
    if (args?.title) chart.title.text = String(args.title);
    if (args?.target_cell) {
      try {
        chart.setPosition(String(args.target_cell));
      } catch { /* 旧版本不支持 */ }
    }
    chart.load("id,name");
    await ctx.sync();
    return { ok: true, chartId: chart.id, chartName: chart.name, chartType };
  });
};

/* ---------------- 全量 API：通用代码执行（excel_run_code） ---------------- */

/**
 * 在 Excel.run 回调内执行 AI 生成的任意 Excel.js 代码，开放完整 Excel 对象模型，
 * 作为内置 excel_* 工具未覆盖能力的兜底入口。
 * - 作用域注入 Excel 与 context；代码体支持 await / return
 * - 返回值在 run 结束前完成序列化（已 load 的代理对象经 toJSON 自动展开为普通数据）
 * - 带超时保护，避免 AI 代码死循环导致工具调用永久挂起
 */
registry["excel.runCode"] = async (args) => {
  const code = String(args?.code ?? "").trim();
  if (!code) throw new Error("缺少必填参数 code（在 Excel.run 回调内执行的 JavaScript 源码）");
  const timeoutMs = Math.min(Math.max(Number(args?.timeout_ms) || 30000, 1000), 120000);

  let serialized;
  await Excel.run(async (ctx) => {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    const fn = new AsyncFunction("Excel", "context", `"use strict";\n${code}`);
    let timer;
    const work = (async () => {
      const result = await fn(Excel, ctx);
      await ctx.sync(); // 收尾同步：确保队列中未提交的写入全部落地
      let out;
      try {
        out = JSON.stringify(result ?? null);
      } catch {
        out = JSON.stringify(String(result));
      }
      return out;
    })();
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(
        () => reject(new Error(`代码执行超时（${timeoutMs}ms）：请检查是否存在死循环、或读取属性前未 load/未 await context.sync()`)),
        timeoutMs
      );
    });
    try {
      serialized = await Promise.race([work, timeout]);
    } finally {
      clearTimeout(timer);
      work.catch(() => {}); // 超时后代码仍在后台收尾，吞掉迟到的 rejection
    }
  });

  try {
    return { ok: true, result: JSON.parse(serialized) };
  } catch {
    return { ok: true, result: serialized };
  }
};

export { registry };
