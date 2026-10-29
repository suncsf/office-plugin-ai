# Office AI 助手（Office 2024 加载项）

> **开发人员**：suncsf · **GitHub**：https://github.com/suncsf

一个面向 **Office 2024 / Microsoft 365（Word · Excel · PowerPoint）** 的任务窗格加载项：
内嵌 **AI Agent 面板（LangGraph.js）**，把 Office 的文档能力封装成 **60+ 个工具**开放给大模型，
让你用自然语言直接查看和修改当前打开的文档。

![stack](https://img.shields.io/badge/LangGraph.js-Agent-blue) ![ui](https://img.shields.io/badge/Vue3-Element%20Plus-green) ![db](https://img.shields.io/badge/SQLite-持久化-lightgrey)

---

## ✨ 功能总览

| 模块 | 能力 |
| --- | --- |
| **Word 工具（26 个）** | 文档信息/全文/分段读取、字数统计、光标位置与移动、选区读取与选择、插入文本/段落/分页符（新建页面）/换行符/符号/超链接/图片/表格、查找替换、删除选区/段落、读取与设置格式（字体/颜色/对齐/缩进/行距/间距 + 文档默认字体）、**读取全局样式库（含每个样式的字体与段落格式定义）**、应用样式（标题/正文/引用等，跨语言）、**创建/更新自定义样式（字体+段落格式定义，可基于现有样式）** |
| **Excel 工具（25 个）** | 工作簿信息、工作表增删改切换、读写区域值、公式读写、行/列插入删除、区域格式化（字体/填充/对齐/数字格式/边框）、自适应行高列宽、排序、智能表格、查找单元格、清除、合并、插入图表 |
| **PPT 工具（11 个）** | 演示文稿信息、列出/读取每页文本、新建（可指定版式）/删除幻灯片、插入文本框/图片/表格、修改形状文字、设置备注（部分能力按 API 集合特性检测） |
| **通用工具** | `ask_user` 表单提问（文本/单选/多选/确认）、`think` 思考计划、`get_office_context`、`attachment_list`/`attachment_read` 附件读取（服务端直执行） |
| **附件** | 消息可携带文档（docx/xlsx/pptx/txt/md/csv/json，≤15MB×5），存于对话工作区 `data/conversations/<对话id>/`；AI 可读取附件内容与格式（段落样式/字体/对齐/行距/表格、单元格数字格式/合并、版式与字号），并可参照附件格式编排当前文档（create_style/apply_style/setFormatting） |
| **Agent** | LangGraph.js 状态图（`agent → tools → agent` 循环），Office 工具在任务窗格执行（`interrupt/resume` 协议），附件工具在服务端直执行；支持 thinking（reasoning_content）流式展示与持久化、表单提问、最大轮次控制 |
| **对话** | 多会话新建/历史/搜索/重命名/删除、Markdown 渲染、思考过程折叠、工具调用卡片（参数/结果/状态）、上下文用量进度条、对话内切换模型 |
| **设置** | 多模型管理（名称 / baseURL / apiKey / model / 上下文长度 / 是否支持工具 / 是否支持图片 / stream_usage）、连接测试、系统提示词、温度、工具轮次、结果截断长度 |
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

> 要求：**Node.js ≥ 18.17（推荐 20/22 LTS）**、Windows 上的 Office 2024 或 Microsoft 365

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
```

---

## 📦 打包安装程序（Win10+ / Office 2021+）

```bash
npm run package     # 组装 staging：应用文件 + 生产依赖 + 内置 Node + 编译控制台
npm run installer   # staging + Inno Setup → installer/output/AI助手-Setup-1.0.0.exe
```

- **免环境**：内置 `node.exe` 与全部生产依赖，装完即用（安装包约 36MB）
- **用户级安装**：无需管理员，默认安装到 `%LOCALAPPDATA%\Programs\OfficeAI`
- **可视化控制台**（OfficeAIConsole.exe）：启动/停止服务、安装证书、注册/取消旁加载、启动 Word/Excel/PPT、开机自启、日志查看；关闭控制台自动停止服务（JobObject）
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
│  ├─ index.js                # Express + HTTPS + SSE + REST API
│  ├─ agent.js                # LangGraph 图、interrupt/resume 协议、历史转换
│  └─ db.js                   # SQLite（conversations / messages / settings）
├─ src/
│  ├─ App.vue                 # 布局与 Tab
│  ├─ store.js / api.js       # 全局状态 / REST+SSE 客户端
│  ├─ office/
│  │  ├─ index.js             # 执行器注册中心、宿主识别
│  │  ├─ word.js / excel.js / ppt.js   # ★ Office.js 工具实现
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
- PPT 的文本/形状能力需要 PowerPointApi 1.8+（Office 2024 与新版 365 支持）；旧版本会返回明确错误。

**模型不支持工具**
- 在模型设置里关闭「允许调用工具」即可作为纯对话模型使用；工具类请求请切换到支持 Function Calling 的模型。

**服务重启后工具中断丢失**
- 每一轮对话是独立的 LangGraph 线程，服务重启会使等待中的工具调用失效，重新发送消息即可。

## 🔐 安全说明

- API Key 仅保存在本机 `data/app.db`，只在你本机的 Node 服务与模型服务商之间传输。
- 加载项权限为 `ReadWriteDocument`，所有文档修改都由你在面板里发起并可见。
