/**
 * 工具目录（Tool Catalog）
 * ---------------------------------------------------------------
 * 所有 Office 能力工具的唯一定义处：
 *  - 服务端（server/agent.js）用它向 LLM 绑定工具（OpenAI function 格式）
 *  - 任务窗格（src/office/*）用它把 executorId 映射到 Office.js 实现
 *
 * 字段说明：
 *  name        工具名（全局唯一，`word_`/`excel_`/`ppt_` 前缀）
 *  host        归属应用：word | excel | ppt | common
 *  executor    客户端执行器注册键（src/office/index.js 的 registry）
 *  description 给 LLM 看的中文说明（含索引约定、注意事项）
 *  parameters  JSON Schema
 */

const S = { type: "object", properties: {}, required: [], additionalProperties: false };
const schema = (properties = {}, required = []) => ({
  type: "object",
  properties,
  required,
  additionalProperties: false,
});
const str = (description, enumValues) =>
  enumValues ? { type: "string", description, enum: enumValues } : { type: "string", description };
const num = (description) => ({ type: "number", description });
const bool = (description) => ({ type: "boolean", description });
const arr = (description, items) => ({ type: "array", description, items });

export const TOOLS = [
  /* ============================ 通用工具 ============================ */
  {
    name: "get_office_context",
    host: "common",
    executor: "common.getOfficeContext",
    description:
      "获取当前打开的 Office 应用信息（应用类型 word/excel/ppt、版本、平台）。当你不确定当前是哪个应用时，先调用本工具。",
    parameters: schema(),
  },
  {
    name: "think",
    host: "common",
    executor: "common.think",
    description:
      "记录你的思考、计划或检查清单，不会对文档产生任何影响，用户可以看到这段思考。在执行复杂的多步文档操作之前，建议先用它列出步骤计划。",
    parameters: schema({ thought: str("思考内容、计划步骤") }, ["thought"]),
  },
  {
    name: "ask_user",
    host: "common",
    executor: "common.askUser",
    description:
      "以表单形式向用户提问并等待回答。form_type: text=文本输入；single_choice=单选；multiple_choice=多选；confirm=是/否确认。当操作有歧义、或要执行破坏性操作（删除/替换大量内容）前，用它向用户确认。single_choice/multiple_choice/confirm 必须提供 options。",
    parameters: schema(
      {
        question: str("要问用户的问题"),
        form_type: str("表单类型", ["text", "single_choice", "multiple_choice", "confirm"]),
        options: arr("选项列表（single_choice / multiple_choice / confirm 时必填）", str("选项")),
      },
      ["question", "form_type"]
    ),
  },

  /* ============================ Word ============================ */
  {
    name: "word_get_document_info",
    host: "word",
    executor: "word.getDocumentInfo",
    description: "获取 Word 文档概要信息：段落总数、字符数、字数（去空白）、节数。不返回正文内容。",
    parameters: schema(),
  },
  {
    name: "word_get_document_text",
    host: "word",
    executor: "word.getDocumentText",
    description: "获取整个 Word 文档的纯文本内容（按段落分隔）。文档过长时会被截断，需要局部内容请改用 word_get_paragraphs。",
    parameters: schema({ max_length: num("最多返回的字符数，默认 8000") }),
  },
  {
    name: "word_get_paragraphs",
    host: "word",
    executor: "word.getParagraphs",
    description:
      "分页获取文档段落列表。每段含 0 起始的 index、文本预览、样式名、是否列表项、大纲级别。用于了解文档结构后再精准操作。",
    parameters: schema({
      offset: num("起始段落 index（0 起始），默认 0"),
      limit: num("本次返回的段落数，默认 50"),
      include_style: bool("是否包含样式信息，默认 true"),
    }),
  },
  {
    name: "word_get_paragraph_text",
    host: "word",
    executor: "word.getParagraphText",
    description: "获取指定段落的完整文本。",
    parameters: schema({ index: num("段落 index（0 起始）") }, ["index"]),
  },
  {
    name: "word_get_selection",
    host: "word",
    executor: "word.getSelection",
    description: "获取当前选区：选中文本、是否为空、样式及字体格式。",
    parameters: schema(),
  },
  {
    name: "word_get_cursor_position",
    host: "word",
    executor: "word.getCursorPosition",
    description:
      "获取光标位置：所在段落 index（0 起始）、该段落文本预览、当前选区文本、段落总数。用于移动光标前的定位。",
    parameters: schema(),
  },
  {
    name: "word_move_cursor",
    host: "word",
    executor: "word.moveCursor",
    description:
      "移动光标到指定段落。position: start=段落开头，end=段落末尾；也可用 search_text 定位段落内第一个匹配文本处。",
    parameters: schema(
      {
        paragraph_index: num("目标段落 index（0 起始）"),
        position: str("光标落在段落的位置", ["start", "end"]),
        search_text: str("（可选）在目标段落内查找该文本，把光标放到其开头"),
      },
      ["paragraph_index"]
    ),
  },
  {
    name: "word_select_text",
    host: "word",
    executor: "word.selectText",
    description: "在全文查找文本并选中第 occurrence 处匹配（默认第 1 处）。返回匹配总数。",
    parameters: schema(
      {
        query: str("要查找的文本"),
        match_case: bool("区分大小写，默认 false"),
        match_whole_word: bool("全字匹配，默认 false"),
        occurrence: num("选中第几处匹配（1 起始），默认 1"),
      },
      ["query"]
    ),
  },
  {
    name: "word_select_paragraph",
    host: "word",
    executor: "word.selectParagraph",
    description: "选中指定段落（整段，含段尾换行）。",
    parameters: schema({ index: num("段落 index（0 起始）") }, ["index"]),
  },
  {
    name: "word_select_all",
    host: "word",
    executor: "word.selectAll",
    description: "全选文档内容（等同 Ctrl+A）。",
    parameters: schema(),
  },
  {
    name: "word_insert_text",
    host: "word",
    executor: "word.insertText",
    description:
      "插入文本。location: cursor=在光标处插入（默认）；end=文档末尾；start=文档开头。插入后该文本处于选中状态。",
    parameters: schema(
      {
        text: str("要插入的文本（可含换行，每个\\n为一段）"),
        location: str("插入位置", ["cursor", "end", "start"]),
      },
      ["text"]
    ),
  },
  {
    name: "word_insert_paragraph",
    host: "word",
    executor: "word.insertParagraph",
    description: "插入一个新段落（可带文字）。location: cursor=在光标/选区之后插入（默认）；end=文档末尾。",
    parameters: schema({
      text: str("新段落的文本，可为空"),
      location: str("插入位置", ["cursor", "end"]),
    }),
  },
  {
    name: "word_insert_page_break",
    host: "word",
    executor: "word.insertPageBreak",
    description: "在光标处插入分页符（新建一页）。",
    parameters: schema(),
  },
  {
    name: "word_insert_line_break",
    host: "word",
    executor: "word.insertLineBreak",
    description: "在光标处插入换行符（软回车，同 Shift+Enter，不产生新段落）。",
    parameters: schema(),
  },
  {
    name: "word_insert_symbol",
    host: "word",
    executor: "word.insertSymbol",
    description: "在光标处插入符号/特殊字符（如 ★、§、①、℃、√ 等），可指定重复次数。",
    parameters: schema(
      {
        symbol: str("要插入的符号（单个或多个字符）"),
        count: num("重复次数，默认 1"),
      },
      ["symbol"]
    ),
  },
  {
    name: "word_insert_hyperlink",
    host: "word",
    executor: "word.insertHyperlink",
    description: "插入超链接。若当前选区为空则先插入 text 再加链接；若选区非空则直接把选中文本变为链接。",
    parameters: schema({ text: str("链接文字"), url: str("链接地址") }, ["text", "url"]),
  },
  {
    name: "word_insert_image",
    host: "word",
    executor: "word.insertImage",
    description: "在光标处插入图片（base64 编码，支持 dataURL）。width/height 单位为磅，可省略保持原始尺寸。",
    parameters: schema({
      base64: str("图片 base64 或 dataURL"),
      width: num("宽度（磅）"),
      height: num("高度（磅）"),
    }, ["base64"]),
  },
  {
    name: "word_insert_table",
    host: "word",
    executor: "word.insertTable",
    description: "插入表格。data 为可选的二维数组（行优先），不足处补空；location: cursor=在光标/选区之后插入（默认）；end=文档末尾。",
    parameters: schema({
      rows: num("行数"),
      cols: num("列数"),
      data: arr("表格数据（二维数组）", arr("一行", str("单元格文本"))),
      location: str("插入位置", ["cursor", "end"]),
    }, ["rows", "cols"]),
  },
  {
    name: "word_replace_text",
    host: "word",
    executor: "word.replaceText",
    description: "查找并替换文本。replace_all=false 时只替换第一处。返回替换数量。",
    parameters: schema(
      {
        search: str("要查找的文本"),
        replace: str("替换为的文本"),
        match_case: bool("区分大小写，默认 false"),
        replace_all: bool("替换全部，默认 true"),
      },
      ["search", "replace"]
    ),
  },
  {
    name: "word_delete_selection",
    host: "word",
    executor: "word.deleteSelection",
    description: "删除当前选中的内容。",
    parameters: schema(),
  },
  {
    name: "word_delete_paragraph",
    host: "word",
    executor: "word.deleteParagraph",
    description: "删除指定段落（含段落标记）。破坏性操作，建议先确认。",
    parameters: schema({ index: num("段落 index（0 起始）") }, ["index"]),
  },
  {
    name: "word_get_formatting",
    host: "word",
    executor: "word.getFormatting",
    description: "读取格式：默认读取当前选区（target=selection），或读取指定段落（target=paragraph + index）。返回字体（名称/字号/加粗/斜体/下划线/颜色/高亮）、对齐、缩进、行距、段前段后间距、大纲级别、样式名，以及文档默认字体。",
    parameters: schema({
      target: str("读取目标", ["selection", "paragraph"]),
      index: num("target=paragraph 时的段落 index"),
    }),
  },
  {
    name: "word_list_styles",
    host: "word",
    executor: "word.listStyles",
    description:
      "枚举 Word 全局样式库（标题、正文、引用等内置样式与文档自定义样式），含每个样式的字体与段落格式定义，应用样式（word_apply_style）前建议先调用本工具了解可用样式。name 参数可查询单个样式详情；type 默认只列 Paragraph（段落）样式。",
    parameters: schema({
      name: str("（可选）按名称查询单个样式的完整定义（本地化名或自定义名）"),
      type: str("样式类型过滤", ["Paragraph", "Character", "Table", "List", "all"]),
      only_in_use: bool("仅列出文档中已使用的样式，默认 false"),
      limit: num("最多返回条数，默认 60"),
    }),
  },
  {
    name: "word_set_formatting",
    host: "word",
    executor: "word.setFormatting",
    description:
      "设置格式。默认作用于当前选区；给 paragraph_index 则作用于指定段落。仅传入需要修改的属性。alignment: Left/Centered/Right/Justified（也接受 Center/居中/两端对齐 等写法，自动转换）；underline: None/Single；颜色用 #RRGGBB，highlight_color 传 null 表示清除高亮。",
    parameters: schema({
      paragraph_index: num("（可选）目标段落 index，不传则作用于选区"),
      font_name: str("字体名称"),
      font_size: num("字号（磅）"),
      bold: bool("加粗"),
      italic: bool("斜体"),
      underline: str("下划线", ["None", "Single", "Double", "Dotted", "Thick"]),
      color: str("字体颜色 #RRGGBB"),
      highlight_color: str("高亮颜色（如 yellow 或 #FFFF00，传 null 清除）"),
      alignment: str("对齐方式", ["Left", "Centered", "Right", "Justified"]),
    }),
  },
  {
    name: "word_apply_style",
    host: "word",
    executor: "word.applyStyle",
    description:
      "为选区（默认）或指定段落应用样式。内置样式会自动转为跨语言枚举（如 \"Heading 1\"/\"标题 1\"/\"h1\" 均映射到 Heading1），在中文/英文版 Word 中均可生效，并回读验证实际样式；自定义样式按文档中的原样名称匹配。返回 appliedVia 与 actualStyle 字段。",
    parameters: schema({
      style: str("样式名，支持：Heading 1~9 / h1~h9 / 标题 1~9、Title/标题、Subtitle/副标题、Normal/正文、Quote/引用、ListParagraph/列表段落、Caption/题注、NoSpacing/无间隔，或文档中的自定义样式名"),
      paragraph_index: num("（可选）目标段落 index"),
    }, ["style"]),
  },

  {
    name: "word_create_style",
    host: "word",
    executor: "word.createStyle",
    description:
      "创建（或覆盖更新）一套自定义格式样式：可定义字体（name/size磅/bold/italic/underline/color）与段落格式（alignment/lineSpacing磅/spaceBefore磅/spaceAfter磅/leftIndent磅/rightIndent磅/firstLineIndent磅），可基于现有样式（base_style）、设置后续段落样式（next_paragraph_style）、是否进入快速样式库（quick_style，默认 true）。创建后用 word_apply_style 应用。font 与 paragraph_format 传 JSON 对象。",
    parameters: schema({
      name: str("样式名（文档内唯一，如“公司正文样式”）"),
      type: str("样式类型", ["Paragraph", "Character", "Table", "List"]),
      base_style: str("（可选）基于的现有样式名"),
      update_if_exists: bool("同名样式已存在时是否更新其格式，默认 false（报错）"),
      quick_style: bool("是否显示在快速样式库，默认 true"),
      next_paragraph_style: str("（可选）该样式段落之后新段落的默认样式名"),
      font: str("（可选）字体定义 JSON：{name, size, bold, italic, underline, color}"),
      paragraph_format: str("（可选）段落格式 JSON：{alignment, lineSpacing, spaceBefore, spaceAfter, leftIndent, rightIndent, firstLineIndent}（单位磅）"),
    }, ["name"]),
  },

  /* ============================ Excel ============================ */
  {
    name: "excel_get_workbook_info",
    host: "excel",
    executor: "excel.getWorkbookInfo",
    description: "获取工作簿信息：所有工作表（名称/位置/可见性）、活动工作表、当前选区地址、活动表已用区域。",
    parameters: schema(),
  },
  {
    name: "excel_activate_sheet",
    host: "excel",
    executor: "excel.activateSheet",
    description: "切换到指定工作表（使其成为活动工作表）。",
    parameters: schema({ name: str("工作表名") }, ["name"]),
  },
  {
    name: "excel_add_sheet",
    host: "excel",
    executor: "excel.addSheet",
    description: "新建工作表并激活。name 可省略（自动命名）；若重名将自动追加序号。",
    parameters: schema({ name: str("新工作表名") }),
  },
  {
    name: "excel_rename_sheet",
    host: "excel",
    executor: "excel.renameSheet",
    description: "重命名工作表。",
    parameters: schema({ old_name: str("原工作表名"), new_name: str("新工作表名") }, ["old_name", "new_name"]),
  },
  {
    name: "excel_delete_sheet",
    host: "excel",
    executor: "excel.deleteSheet",
    description: "删除工作表。破坏性操作，建议先确认。不能删除最后一个工作表。",
    parameters: schema({ name: str("工作表名") }, ["name"]),
  },
  {
    name: "excel_get_used_range",
    host: "excel",
    executor: "excel.getUsedRange",
    description: "获取工作表已使用区域：地址、行列数及前 preview_rows 行数据（默认 20 行，值会被截断）。",
    parameters: schema({
      sheet: str("工作表名，默认活动工作表"),
      preview_rows: num("预览行数，默认 20"),
    }),
  },
  {
    name: "excel_get_range",
    host: "excel",
    executor: "excel.getRange",
    description: "读取区域数据（二维数组）与数字格式。address 如 A1:C10 或 B2。",
    parameters: schema({
      address: str("区域地址，如 A1:C10"),
      sheet: str("工作表名，默认活动工作表"),
    }, ["address"]),
  },
  {
    name: "excel_get_selected_range",
    host: "excel",
    executor: "excel.getSelectedRange",
    description: "获取当前选中区域：地址、行列数与数据预览。",
    parameters: schema(),
  },
  {
    name: "excel_set_range_values",
    host: "excel",
    executor: "excel.setRangeValues",
    description:
      "写入数据。values 可为：单值（填充整个区域）、一维数组（单行，行数需为 1）、二维数组（尺寸必须与区域一致，否则报错）。地址数量不符时会返回错误提示。",
    parameters: schema({
      address: str("区域地址，如 A1:C3"),
      values: str("要写入的值（标量/一维/二维数组）"),
      sheet: str("工作表名，默认活动工作表"),
    }, ["address", "values"]),
  },
  {
    name: "excel_set_formula",
    host: "excel",
    executor: "excel.setFormula",
    description: "为单元格设置公式（以 = 开头，如 =SUM(A1:A10)），公式使用英文逗号分隔。",
    parameters: schema({ address: str("单元格地址"), formula: str("公式，以 = 开头"), sheet: str("工作表名") }, ["address", "formula"]),
  },
  {
    name: "excel_get_formula",
    host: "excel",
    executor: "excel.getFormula",
    description: "读取单元格公式与当前计算值。",
    parameters: schema({ address: str("单元格地址"), sheet: str("工作表名") }, ["address"]),
  },
  {
    name: "excel_insert_rows",
    host: "excel",
    executor: "excel.insertRows",
    description: "在第 index 行（0 起始）之前插入 count 行整行。",
    parameters: schema({ index: num("插入位置行号（0 起始）"), count: num("插入行数，默认 1"), sheet: str("工作表名") }, ["index"]),
  },
  {
    name: "excel_delete_rows",
    host: "excel",
    executor: "excel.deleteRows",
    description: "删除从 index 行（0 起始）开始的 count 行。破坏性操作，建议先确认。",
    parameters: schema({ index: num("起始行号（0 起始）"), count: num("删除行数，默认 1"), sheet: str("工作表名") }, ["index"]),
  },
  {
    name: "excel_insert_columns",
    host: "excel",
    executor: "excel.insertColumns",
    description: "在第 index 列（0 起始）之前插入 count 列整列。",
    parameters: schema({ index: num("插入位置列号（0 起始）"), count: num("插入列数，默认 1"), sheet: str("工作表名") }, ["index"]),
  },
  {
    name: "excel_delete_columns",
    host: "excel",
    executor: "excel.deleteColumns",
    description: "删除从 index 列（0 起始）开始的 count 列。破坏性操作，建议先确认。",
    parameters: schema({ index: num("起始列号（0 起始）"), count: num("删除列数，默认 1"), sheet: str("工作表名") }, ["index"]),
  },
  {
    name: "excel_format_range",
    host: "excel",
    executor: "excel.formatRange",
    description:
      "设置区域格式。format 可含：font_name、font_size、bold、italic、underline、font_color、fill_color、horizontal_alignment(Left/Center/Right)、vertical_alignment(Top/Center/Bottom)、number_format(如 0.00、0%)、wrap_text、border_style(Continuous/Dash/Dot/Double/None)、border_color。仅传入需要修改的项。",
    parameters: schema({
      address: str("区域地址"),
      sheet: str("工作表名"),
      format: str("格式键值对象（JSON）"),
    }, ["address", "format"]),
  },
  {
    name: "excel_auto_fit",
    host: "excel",
    executor: "excel.autoFit",
    description: "自动调整行高列宽。address 可省略（默认整个已用区域）。",
    parameters: schema({ address: str("区域地址"), sheet: str("工作表名") }),
  },
  {
    name: "excel_set_column_width",
    host: "excel",
    executor: "excel.setColumnWidth",
    description: "设置列宽（磅）。column 支持列号（0 起始）或字母（如 A）。",
    parameters: schema({ column: str("列号（0 起始）或字母，如 0 或 A"), width: num("宽度（磅，默认约 8.43）"), sheet: str("工作表名") }, ["column", "width"]),
  },
  {
    name: "excel_set_row_height",
    host: "excel",
    executor: "excel.setRowHeight",
    description: "设置行高（磅）。row 为 0 起始行号。",
    parameters: schema({ row: num("行号（0 起始）"), height: num("高度（磅，默认约 14.4）"), sheet: str("工作表名") }, ["row", "height"]),
  },
  {
    name: "excel_sort_range",
    host: "excel",
    executor: "excel.sortRange",
    description: "对区域排序。by_column 为区域内的相对列（0 起始）。",
    parameters: schema({
      address: str("要排序的区域"),
      by_column: num("排序依据的相对列（0 起始），默认 0"),
      ascending: bool("升序，默认 true"),
      has_headers: bool("首行为标题（不参与排序），默认 false"),
      sheet: str("工作表名"),
    }, ["address"]),
  },
  {
    name: "excel_create_table",
    host: "excel",
    executor: "excel.createTable",
    description: "把区域转为智能表格（ListObject），可带名称。",
    parameters: schema({
      address: str("区域地址"),
      name: str("表格名"),
      has_headers: bool("首行是否为标题，默认 true"),
      sheet: str("工作表名"),
    }, ["address"]),
  },
  {
    name: "excel_find_cells",
    host: "excel",
    executor: "excel.findCells",
    description: "在已用区域内查找包含指定文本的单元格（包含匹配），返回地址列表（最多 50 个）。",
    parameters: schema({
      text: str("要查找的文本"),
      match_case: bool("区分大小写，默认 false"),
      sheet: str("工作表名"),
    }, ["text"]),
  },
  {
    name: "excel_clear_range",
    host: "excel",
    executor: "excel.clearRange",
    description: "清除区域。mode: all=全部（默认）；contents=仅内容；formats=仅格式。",
    parameters: schema({
      address: str("区域地址"),
      mode: str("清除模式", ["all", "contents", "formats"]),
      sheet: str("工作表名"),
    }, ["address"]),
  },
  {
    name: "excel_merge_cells",
    host: "excel",
    executor: "excel.mergeCells",
    description: "合并单元格；unmerge=true 时取消合并。",
    parameters: schema({ address: str("区域地址"), unmerge: bool("取消合并，默认 false"), sheet: str("工作表名") }, ["address"]),
  },
  {
    name: "excel_add_chart",
    host: "excel",
    executor: "excel.addChart",
    description:
      "基于数据区域插入图表。chart_type 可选：ColumnClustered(柱形,默认)、BarClustered(条形)、Line(折线)、Pie(饼)、Area(面积)、XYScatter(散点)。",
    parameters: schema({
      data_address: str("数据区域地址（含标题行）"),
      chart_type: str("图表类型", ["ColumnClustered", "BarClustered", "Line", "Pie", "Area", "XYScatter"]),
      title: str("图表标题"),
      target_cell: str("图表左上角放置的单元格，如 E2"),
      sheet: str("工作表名"),
    }, ["data_address"]),
  },

  /* ============================ PowerPoint ============================ */
  {
    name: "ppt_get_presentation_info",
    host: "ppt",
    executor: "ppt.getPresentationInfo",
    description: "获取演示文稿信息：幻灯片总数、页面尺寸（磅，视 API 支持情况）。",
    parameters: schema(),
  },
  {
    name: "ppt_list_slides",
    host: "ppt",
    executor: "ppt.listSlides",
    description: "列出所有幻灯片：index（0 起始）、id、形状数量、文本预览（读取文本需要较新的 PowerPoint API，Office 2024 满足）。",
    parameters: schema({ include_text: bool("是否包含每页文本预览，默认 true") }),
  },
  {
    name: "ppt_get_slide_text",
    host: "ppt",
    executor: "ppt.getSlideText",
    description: "获取指定幻灯片上的全部文本（按形状分组，含形状位置与 id）。",
    parameters: schema({ slide_index: num("幻灯片 index（0 起始）") }, ["slide_index"]),
  },
  {
    name: "ppt_get_all_text",
    host: "ppt",
    executor: "ppt.getAllText",
    description: "获取整个演示文稿每一页的文本内容。",
    parameters: schema({ max_slides: num("最多读取页数，默认 50") }),
  },
  {
    name: "ppt_add_slide",
    host: "ppt",
    executor: "ppt.addSlide",
    description:
      "新建幻灯片。layout_index 为版式序号（0 起始，可先用本工具返回的 layouts 查看，缺省使用默认版式）；title_text 传入时会在页面顶部添加标题文本框。",
    parameters: schema({
      layout_index: num("版式序号（0 起始）"),
      title_text: str("标题文字"),
    }),
  },
  {
    name: "ppt_delete_slide",
    host: "ppt",
    executor: "ppt.deleteSlide",
    description: "删除指定幻灯片。破坏性操作，建议先确认。",
    parameters: schema({ slide_index: num("幻灯片 index（0 起始）") }, ["slide_index"]),
  },
  {
    name: "ppt_move_slide",
    host: "ppt",
    executor: "ppt.moveSlide",
    description: "把幻灯片从 from_index 移动到 to_index。若当前 PowerPoint JS API 不支持移动会返回明确错误。",
    parameters: schema({ from_index: num("原位置（0 起始）"), to_index: num("目标位置（0 起始）") }, ["from_index", "to_index"]),
  },
  {
    name: "ppt_insert_textbox",
    host: "ppt",
    executor: "ppt.insertTextbox",
    description:
      "在幻灯片上插入文本框。坐标与尺寸单位为磅（默认页面 960x540）。可设置 font_size/bold/color/align(Left/Center/Right)。",
    parameters: schema({
      slide_index: num("幻灯片 index，默认当前最后追加页？实际默认为第 0 页，建议显式传入"),
      text: str("文本内容（\\n 换行）"),
      left: num("左边距（磅），默认 60"),
      top: num("上边距（磅），默认 60"),
      width: num("宽度（磅），默认 300"),
      height: num("高度（磅），默认 80"),
      font_size: num("字号（磅），默认 18"),
      bold: bool("加粗"),
      color: str("文字颜色 #RRGGBB"),
      align: str("水平对齐", ["Left", "Center", "Right"]),
    }, ["text"]),
  },
  {
    name: "ppt_set_shape_text",
    host: "ppt",
    executor: "ppt.setShapeText",
    description:
      "修改幻灯片上已有形状/文本框的文字：用 shape_id 精确定位，或用 match_text 匹配第一个包含该文本的形状。",
    parameters: schema({
      slide_index: num("幻灯片 index（0 起始）"),
      shape_id: num("形状 id"),
      match_text: str("匹配文本（包含即命中）"),
      text: str("新的文字"),
    }, ["slide_index", "text"]),
  },
  {
    name: "ppt_insert_image",
    host: "ppt",
    executor: "ppt.insertImage",
    description: "在幻灯片上插入图片（base64/dataURL），可指定位置与尺寸（磅）。",
    parameters: schema({
      slide_index: num("幻灯片 index（0 起始）"),
      base64: str("图片 base64 或 dataURL"),
      left: num("左边距（磅），默认 60"),
      top: num("上边距（磅），默认 60"),
      width: num("宽度（磅）"),
      height: num("高度（磅）"),
    }, ["base64"]),
  },
  {
    name: "ppt_insert_table",
    host: "ppt",
    executor: "ppt.insertTable",
    description: "在幻灯片上插入表格，data 为可选二维数组（行优先）。",
    parameters: schema({
      slide_index: num("幻灯片 index（0 起始）"),
      rows: num("行数"),
      cols: num("列数"),
      data: arr("表格数据（二维数组）", arr("一行", str("单元格文本"))),
      left: num("左边距（磅），默认 60"),
      top: num("上边距（磅），默认 120"),
      width: num("宽度（磅），默认 500"),
    }, ["rows", "cols"]),
  },
  {
    name: "ppt_set_slide_notes",
    host: "ppt",
    executor: "ppt.setSlideNotes",
    description: "设置指定幻灯片的演讲者备注。",
    parameters: schema({ slide_index: num("幻灯片 index（0 起始）"), notes: str("备注内容") }, ["slide_index", "notes"]),
  },
  {
    name: "attachment_list",
    host: "common",
    side: "server",
    executor: "attachment.list",
    description:
      "列出当前对话中的所有附件及每个附件的概览（Word：段落数/表格数/标题大纲；Excel：工作表与行列范围；PPT：幻灯片与版式；文本：字符数）。用户上传文档后，先用本工具了解有哪些附件、该从哪里读起。",
    parameters: schema({}, []),
  },
  {
    name: "attachment_read",
    host: "common",
    side: "server",
    executor: "attachment.read",
    description:
      "读取附件文档的内容与格式。docx：mode=structure(标题大纲)/text(段落文本)/format(逐段字体、字号、加粗、颜色、对齐、缩进、行距)/tables(表格数据)；xlsx：sheet 选表，mode=structure/text/format(单元格值+数字格式+加粗字体+合并区域+列宽)；pptx：mode=structure/text/format(版式、字体、字号)；txt/md/csv/json 直接返回文本。读取附件格式后，可用 word_create_style/word_apply_style/word_setFormatting/word_setFormatting 等工具按附件样式编排当前文档。",
    parameters: schema({
      id: str("附件 id（来自 attachment_list 或用户消息的附件标注）"),
      mode: str("读取模式", ["structure", "text", "format", "tables"]),
      offset: num("（可选）起始位置：docx=段落 index，xlsx=行号(1 起)，pptx=幻灯片序号(1 起)，文本=字符偏移"),
      limit: num("（可选）返回数量上限"),
      sheet: str("（可选）Excel 工作表名或序号（1 起）"),
    }, ["id"]),
  },

];

export const TOOL_MAP = Object.fromEntries(TOOLS.map((t) => [t.name, t]));

/** 服务端工具：在 Node 服务内直接执行，不经过任务窗格中断 */
export const SERVER_TOOLS = new Set(TOOLS.filter((t) => t.side === "server").map((t) => t.name));

export const HOST_NAMES = {
  word: "Word 文档",
  excel: "Excel 工作簿",
  ppt: "PowerPoint 演示文稿",
  common: "通用",
  unknown: "未知应用",
};
