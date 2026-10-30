# Office AI 助手（Office 加载项：2024 / 365 / 2019）

> **开发人员**：见本地配置（`about-info.local.js`，隐私信息不入库）
> 📦 完整开发归档（架构/决策/修复史/验证）：[docs/ARCHIVE-2026-10-29.md](docs/ARCHIVE-2026-10-29.md) ·
> 增量归档（Office 2019 IE11 兼容版）：[docs/ARCHIVE-2026-10-30.md](docs/ARCHIVE-2026-10-30.md)

一个面向 **Office 2024 / Microsoft 365 / Office 2019（Word · Excel · PowerPoint）** 的任务窗格加载项：
内嵌 **AI Agent 面板（LangGraph.js）**，把 Office 的文档能力封装成 **109 个工具**开放给大模型，
让你用自然语言直接查看和修改当前打开的文档。

> **Office 2019 兼容**：2019 永久版的任务窗格 WebView 锁定 IE11（无法运行 Vue3），项目内置一套
> **IE11 兼容版前端**（`src-legacy/`，Vue2 + Element UI，构建产物纯 ES5），服务端按浏览器内核
> 自动分流——2019 与新版 Office 共用同一 manifest 与后端，无需分别安装。

![stack](https://img.shields.io/badge/LangGraph.js-Agent-blue) ![ui](https://img.shields.io/badge/Vue3-Element%20Plus-green) ![db](https://img.shields.io/badge/SQLite-持久化-lightgrey)

---

## ✨ 功能总览

| 模块 | 能力 |
| --- | --- |
| **Word 工具（41 个，含全 API 代码执行）** | 文档信息/全文/分段读取、字数统计、光标位置与移动、选区读取与选择、插入文本/段落/分页符（新建页面）/换行符/符号/超链接/图片/表格、查找替换、删除选区/段落、读取与设置格式（字体/颜色/对齐/缩进/行距/间距 + 文档默认字体）、**读取全局样式库（含每个样式的字体与段落格式定义）**、应用样式（标题/正文/引用等，跨语言）、**创建/更新自定义样式（字体+段落格式定义，可基于现有样式）**、**表格操作：读取/单元格读写/行列增删/合并/格式**、**图片操作：列出内嵌图片（尺寸）、导出图片到工作区、调整图片大小（等比/非等比，支持选区定位）**、**`word_run_code` 通用代码执行：直接运行任意 Word.js API 代码，开放全部 Word 对象模型（页眉页脚/节与页面设置/脚注尾注/批注/目录/域/书签/内容控件/文档属性等内置工具未覆盖的场景）** |
| **Excel 工具（30 个，含全 API 代码执行）** | 工作簿信息、工作表增删改切换、读写区域值、公式读写、行/列插入删除、区域格式化（字体/填充/对齐/数字格式/边框）、自适应行高列宽、排序、智能表格、查找单元格、清除、合并、插入图表、**`excel_write_json` JSON 写表：对象数组自动生成表头+数据行（免算地址/免对齐尺寸），写模拟数据与结构化数据首选**、**图片操作：插入浮动图片（锚定单元格/坐标）、列出图片、调整图片大小**、**`excel_run_code` 通用代码执行：直接运行任意 Excel.js API 代码，开放全部 Excel 对象模型（冻结窗格/工作表保护/条件格式/数据验证/命名区域/批注/自动筛选/分组大纲/透视表/形状图片等内置工具未覆盖的场景）** |
| **PPT 工具（27 个，含全 API 代码执行）** | 演示文稿信息、列出/读取每页文本、**读取形状清单**、新建/删除/移动幻灯片、读取版式、插入文本框/图片/表格/**几何形状/线条**、修改形状文字、**删除单个/多个形状、移动缩放形状、形状格式（填充/边框）、文字格式、层级调整（置顶/置底）、解散组合**、**文本框三件套：读取（含文字格式详情）/大小与位置调整/删除**、设置/读取备注（部分能力按 API 集合特性检测）、**图片导出：把页内图片形状保存为工作区图片文件（需 PowerPointApi 1.10+，按特性检测）**、**`ppt_run_code` 通用代码执行：直接运行任意 PowerPoint.js API 代码，开放全部 PowerPoint 对象模型（占位符/组合与层级/表格样式/主题与母版信息/幻灯片切换等内置工具未覆盖的场景）** |
| **通用工具** | `ask_user` 表单提问（文本/单选/多选/确认）、`think` 思考计划、`get_office_context`、`attachment_list`/`attachment_read` 附件读取（服务端直执行）、`shell` 本机命令执行（PowerShell/cmd，服务端直执行，可在设置中关闭）、`memory_save/list/update/delete` 长期记忆（跨对话持久）、`image_process` 图片处理（服务端 jimp：缩放/裁剪/旋转/翻转/转格式/压缩，作用于对话工作区文件） |
| **图片能力** | 插入：Word/PPT/Excel 均可插入图片（base64 或工作区文件，Excel 锚定单元格）；来源：用户贴图自动落盘（`images/`）、图片附件（png/jpg/gif/bmp/webp）、文档内提取（word_get_image / ppt_get_image）、shell 下载生成；裁剪：Office.js 不支持文档内裁剪，走「导出 → `image_process` 裁剪 → file 参数插回」流程；调大小：Word 内嵌图（等比/选区）、Excel/PPT 浮动图；工作区文件经受限桥接端点（`/api/workspace/file`，严格限制在对话工作区内）在任务窗格与服务端之间直传，不经过 LLM 上下文 |
| **附件** | 消息可携带文档或图片（docx/xlsx/pptx/txt/md/csv/json/log + png/jpg/jpeg/gif/bmp/webp/tif/tiff，≤15MB×5），存于对话工作区 `data/conversations/<对话id>/`；文档附件 AI 可读取内容与格式（段落样式/字体/对齐/行距/表格、单元格数字格式/合并、版式与字号），并可参照附件格式编排当前文档（create_style/apply_style/setFormatting）；图片附件不解析内容，可整份插入文档或先用 image_process 裁剪/缩放 |
| **Agent** | LangGraph.js 状态图（`agent → tools → agent` 循环），Office 工具在任务窗格执行（`interrupt/resume` 协议），附件与 shell 工具在服务端直执行；支持 thinking（reasoning_content）流式展示与持久化、表单提问、最大轮次控制 |
| **对话** | 多会话新建/历史/搜索/重命名/删除、Markdown 渲染、思考过程折叠、工具调用卡片（参数/结果/状态）、上下文用量进度条、对话内切换模型 |
| **设置** | 多模型管理（名称 / baseURL / apiKey / model / 上下文长度 / 是否支持工具 / 是否支持图片 / stream_usage）、连接测试、系统提示词、温度、工具轮次、结果截断长度、Shell 开关 |
| **持久化** | SQLite（`data/app.db`）：对话、消息（含工具调用与结果、思考过程）、模型配置、附件记录；对话工作区文件夹 `data/conversations/<对话id>/` 存放该对话产生的所有文件 |

---

## 🏗 架构

```
┌─────────────────────────── Office 2024 (Word/Excel/PPT) ───────────────────────────┐
│  任务窗格（WebView2）                                                               │
│  ┌─────────────────────────────────────────────────────┐                           │
│  │  Vue3 + Element Plus  对话 / 历史 / 设置              │                           │
│  │  Office.js 工具执行器（src/office/*）                 │─── Word.run / Excel.run / │
│  └────────────────────┬────────────────────────────────┘     PowerPoint.run          │
└───────────────────────┼──────────────────────────────────────────────────────────────┘
                        │ SSE 流 + POST /api/chat/resume（工具执行结果）
┌───────────────────────▼──────────────────────────────────────────────────────────────┐
│  Node 服务（Express, HTTPS :3000）                                                    │
│   ├─ LangGraph.js Agent：agent → tools(interrupt 等待客户端执行) → agent               │
│   ├─ LLM 代理：任意 OpenAI 兼容 API（DeepSeek / 通义 / Kimi / GLM / Ollama / OpenAI…） │
│   ├─ SQLite（better-sqlite3）：会话、消息、模型配置                                    │
│   └─ 静态托管 Vue 构建产物（manifest.xml 指向 /taskpane.html）                         │
└──────────────────────────────────────────────────────────────────────────────────────┘
```

**为什么工具在浏览器端执行？** Office.js 只能运行在 Office 的任务窗格进程里。
LangGraph 的 `tools` 节点通过 `interrupt()` 暂停图执行，把工具调用流式推给任务窗格；
任务窗格调用 Office.js 执行后 `POST /api/chat/resume` 回传结果，图以 `Command({ resume })` 继续。
这样服务端只做 LLM 调度与持久化，文档操作始终发生在 Office 内。

---

## 🚀 快速开始

> 要求：**Node.js ≥ 18.17（推荐 20/22 LTS）**、Windows 上的 Office 2019 / 2021 / 2024 或 Microsoft 365
> （Office 2019 自动使用 IE11 兼容版界面，功能见下方「已知限制」）

```bash
# 1. 安装依赖
npm install

# 2. 安装 Office 开发证书（manifest 要求 HTTPS）
npm run setup

# 3. 构建任务窗格并生成本地服务
npm run build

# 4. 启动服务（另开一个窗口也可以）
npm run start

# 5. 旁加载（sideload）加载项到 Office —— 在新窗口执行
npm run sideload            # 默认打开 Word
npm run sideload -- excel   # 指定 Excel
```

启动后在 Word/Excel/PPT 的 **「开始」→「加载项」** 里点击 **AI 助手** 打开面板。

### 首次使用

1. 打开面板 → **设置 → 模型管理 → 添加模型**：
   - DeepSeek：`https://api.deepseek.com/v1`，model `deepseek-chat`（或 `deepseek-reasoner`， thinking 会自动展示）
   - 通义千问：`https://dashscope.aliyuncs.com/compatible-mode/v1`，model `qwen-max`
   - Kimi：`https://api.moonshot.cn/v1`，model `moonshot-v1-8k`
   - GLM：`https://open.bigmodel.cn/api/paas/v4`，model `glm-4-plus`
   - 本地 Ollama：`http://localhost:11434/v1`，model `qwen2.5:14b`（key 随意填）
2. 点击 **测试连接** 确认可用 → 保存。
3. 回到 **对话** 页，右上角选择模型，开始对话。试试：

```
帮我看看这篇文档的结构，把所有标题列出来
把第一段改写得更正式，然后应用到文档里
在文档末尾新建一页，插入一个 3×4 的表格，写上项目排期
（Excel）把 A 列大于 100 的行标成红色，再按 B 列降序排序
（PPT）给每页加一个"项目名称"文本框在左上角
（Word）把文档里第 1 张图片裁成正方形，缩小到 200 磅宽，放到第二段后面
（Excel）把我刚发的截图插入到 E2 单元格，宽度缩到 300 磅
```

---

## 📦 打包安装程序（Win10+ / Office 2019+）

```bash
npm run package     # 组装 staging：应用文件 + 生产依赖 + 内置 Node + 编译控制台
npm run installer   # staging + Inno Setup → installer/output/AI助手-Setup-1.0.0.exe
```

- **免环境**：内置 `node.exe` 与全部生产依赖，装完即用（安装包约 36MB）
- **用户级安装**：无需管理员，默认安装到 `%LOCALAPPDATA%\Programs\OfficeAI`
- **可视化控制台**（OfficeAIConsole.exe）：启动/停止服务、安装证书、注册/取消旁加载、启动 Word/Excel/PPT、开机自启、日志查看；最小化隐藏到系统托盘（任务栏不留按钮，托盘左键恢复/右键退出，重复启动唤起已隐藏窗口）；关闭控制台自动停止服务（JobObject）
- **首次自动就绪**：控制台打开时自动安装开发证书 → 注册加载项 → 启动服务
- **数据与程序分离**：用户数据在 `%LOCALAPPDATA%\OfficeAI\data`，卸载/升级不丢失
- 安装任务可选：桌面快捷方式 / 开机自启 / 旁加载注册（`HKCU\...\Wef\Developer`）
- 构建前置：仅需 Inno Setup（`winget install JRSoftware.InnoSetup`），其余零依赖

---

## 📁 目录结构

```
office-plugin-ai/
├─ manifest.xml               # Office 加载项清单（Word/Excel/PPT 任务窗格）
├─ index.html                 # 任务窗格入口（引入 Office.js）
├─ vite.config.js
├─ shared/
│  └─ tools-catalog.js        # ★ 全部工具定义（服务端绑定 LLM / 客户端执行 共用）
├─ server/
│  ├─ index.js                # Express + HTTPS + SSE + REST API + 工作区文件桥接端点
│  ├─ agent.js                # LangGraph 图、interrupt/resume 协议、历史转换
│  ├─ image.js                # ★ 图片处理核心（jimp 缩放/裁剪/转格式 + 工作区路径安全）
│  ├─ shell.js                # 本机命令执行（PowerShell/cmd）
│  └─ db.js                   # SQLite（conversations / messages / settings / memories）
├─ src/
│  ├─ App.vue                 # 布局与 Tab
│  ├─ store.js / api.js       # 全局状态 / REST+SSE 客户端
│  ├─ office/
│  │  ├─ index.js             # 执行器注册中心、宿主识别
│  │  ├─ word.js / excel.js / ppt.js   # ★ Office.js 工具实现
│  │  └─ bridge.js            # 工作区图片文件桥接（插入图片的 file 参数 / 提取图落盘）
│  └─ components/
│     ├─ ChatView.vue         # 对话编排（SSE、interrupt 执行、上下文额度）
│     ├─ MessageBubble.vue    # Markdown / 思考 / 工具卡片 / 表单提问
│     ├─ ChatInput.vue        # 输入、图片附件
│     ├─ HistoryView.vue      # 历史会话
│     └─ SettingsView.vue     # 模型与通用设置
├─ scripts/gen-icons.mjs      # 生成清单所需 PNG 图标
└─ data/app.db                # SQLite（运行后生成）
```

---

## 🔧 常用命令

| 命令 | 说明 |
| --- | --- |
| `npm run build` | 构建任务窗格到 `dist/` |
| `npm run start` | 启动 HTTPS 服务（`https://localhost:3000`） |
| `npm run setup` | 安装 office-addin-dev-certs 信任证书 |
| `npm run sideload` | 启动服务并旁加载到 Office（`-- word/excel/powerpoint`） |
| `npm run validate` | 校验 manifest.xml |
| `npm run test` | Agent 离线自测（桩模型验证 interrupt/resume 循环，无需 LLM） |
| `npm run dev` | Vite 热更新开发（需同时 `npm run start`，自动代理 `/api`） |

---

## ❓ 故障排查

**加载项不出现 / 无法加载**
- 确认服务已启动且用的是 HTTPS：访问 `https://localhost:3000/api/health`。
- 证书未信任：运行 `npm run setup`，重启 Office。
- 旁加载后没显示：重启 Office；或使用「插入 → 获取加载项 → 上传我的加载项」手动选择 `manifest.xml`。

**SQLite 安装失败（better-sqlite3 编译错误）**
- 升级到 Node 20/22 LTS（有预编译二进制）；或安装 VS Build Tools 后重试。

**工具调用失败**
- 任务窗格右上角会显示当前宿主；跨应用工具会被拒绝（例如在 Excel 里调用 Word 工具）。
- PPT 的形状级操作（删除/移动/格式化单个形状）需要 PowerPointApi 1.3+/1.4+，文本框、图片、表格添加需要 PowerPointApi 1.8+（Office 2024 与新版 365 均支持）；旧版本会返回明确错误。

**Office 2019（IE11 兼容版）能力差异**
- Office 2019 任务窗格运行在 IE11 内核，前端自动切换为兼容版（Vue2 + Element UI），界面与新版略有差异、无热更新。
- `word_run_code` / `excel_run_code` / `ppt_run_code`（AI 直接执行代码）在 IE11 下**不可用**（依赖 async/await 语言特性，服务端已自动对兼容版隐藏这三个工具，其余工具不受影响）。
- PPT 工具集本身要求 PowerPointApi 1.8+，而 Office 2019 的 PowerPointApi 上限为 1.1，因此 PPT 工具在 2019 下基本不可用；Word/Excel 的基础读写工具可用（Word API 上限 1.3、Excel API 上限 1.7）。
- 更完整的体验建议 Office 2021 / 2024 / Microsoft 365（WebView2）。

**模型不支持工具**
- 在模型设置里关闭「允许调用工具」即可作为纯对话模型使用；工具类请求请切换到支持 Function Calling 的模型。

**服务重启后工具中断丢失**
- 每一轮对话是独立的 LangGraph 线程，服务重启会使等待中的工具调用失效，重新发送消息即可。

## 🔐 安全说明

- API Key 仅保存在本机 `data/app.db`，只在你本机的 Node 服务与模型服务商之间传输。
- 加载项权限为 `ReadWriteDocument`（任务窗格加载项的最高权限级别，Word.js 富 API 全部可用，无需再提升）。
- `word_run_code` 允许 AI 直接执行 Word.js 代码（完整对象模型）：代码仅在本机任务窗格（WebView2）内运行、只操作当前文档，带超时保护（默认 30s）；内置 word_* 工具优先、该工具仅作兜底。
- `excel_run_code` 允许 AI 直接执行 Excel.js 代码（完整对象模型）：与 `word_run_code` 同等安全边界——仅在本机任务窗格内运行、只操作当前工作簿，带超时保护（默认 30s）；内置 excel_* 工具优先、该工具仅作兜底。
- `shell` 允许 AI 在本机（Node 服务所在电脑）执行命令行：默认 PowerShell，可指定 cmd；工作目录默认为当前对话的工作区；超时杀整棵进程树（默认 60s、上限 10 分钟）；stdout/stderr 截断防撑爆上下文。**「设置 → 通用设置 → 允许 AI 执行 Shell 命令」默认开启，关闭后 AI 既看不到也执行不了该工具**（绑定与执行双重拦截）。命令以服务进程同权限真实作用于本机，AI 被提示在破坏性操作前先用 ask_user 确认；输出统一按 UTF-8 解码，个别老旧 GBK 中文程序输出可能乱码。
- `image_process` 与工作区文件端点（`/api/workspace/file`）只作用于**当前对话的工作区目录**（`data/conversations/<对话id>/`）：路径服务端二次校验（拒绝绝对路径与 `..` 越界）、仅接受图片格式（魔数嗅探）、单文件 ≤20MB、处理结果输出为新文件不覆盖原图；不同对话的工作区互相隔离。
- `word_get_image` / `ppt_get_image` 只读取当前打开文档内的图片并保存到对话工作区，不上传任何远端；图片数据在任务窗格与服务端之间经本机 HTTPS 直传，不进入 LLM 上下文（工具参数/结果只有文件名与尺寸）。
