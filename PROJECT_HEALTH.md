# PROJECT_HEALTH.md — Excel VBA Assistant 项目体检报告

> 体检日期：2026-09-13。全程只读取证，本文件是本次体检唯一的写入产物。
> 体检范围：本地仓库 `D:\excel-vba-assistant`（git 46 条 commit，非浅克隆，工作区干净）。
> 外网限制：本机代理不可达，无法 fetch/访问 GitHub。凡涉及远端状态的结论均标注"因外网不通无法确认"。

---

## 一、技术栈识别

| 项 | 结论 | 证据 |
|---|---|---|
| 主语言 | **TypeScript**（唯一源码语言） | `git ls-files` 按扩展名统计：16 个 `.ts`、42 个 `.md`（skills 文档）、4 个 `.js`（tools 构建脚本）；无 go.mod / Cargo.toml / pom.xml / *.csproj / CMakeLists.txt |
| 运行时 | Node.js（VS Code 扩展宿主）+ 外部 exe 子进程 | `package.json:36` `"main": "./dist/extension.js"`；`package.json:28` `"engines": {"vscode": "^1.85.0"}` |
| 框架/平台 | **VS Code / Trae 编辑器扩展**（Webview 面板 + 命令 + 状态栏 + 图标主题） | `package.json:37-154` `contributes.viewsContainers / views / commands / iconThemes / colors` |
| 核心外部依赖 | `@modelcontextprotocol/sdk` ^1.0.0（实际安装 1.29.0） | `package.json:173`；`node_modules/@modelcontextprotocol/sdk/package.json` |
| 构建方式 | esbuild 打包（非 tsc 编译输出；tsc 仅做类型检查） | `package.json:156` `esbuild ./src/extension.ts --bundle --outfile=dist/extension.js --external:vscode ...` |
| 依赖管理 | npm + package-lock.json（lock 被 .gitignore 忽略，未入库） | `.gitignore:3`；本地存在 `package-lock.json`（178KB） |
| 打包 | vsce 生成 .vsix，内嵌 156MB 的 `mcp-excel.exe`（来自第三方 `sbroenne/mcp-server-excel`） | `package.json:160` `"package": "npm run build && npm run download:mcp-server && vsce package ..."`；`tools/download-mcp-server.js:11-13` |
| 交互技术 | PowerShell 5.1 + C# Add-Type 内联代码，通过 Excel COM 自动化操作 VBE | `src/runtime/powershell.ts`（231 行）；`src/client/vbaClient.ts` 头注释"通过 PowerShell 调用 Excel COM Automation" |
| MCP 架构 | 插件内嵌 McpClient（stdio）启动 `dist/mcp-excel.exe`，并向工作区 `.trae/mcp.json` 写入配置让 Trae AI 直连同一 exe | `src/mcp/serverManager.ts:48`；`src/extension.ts:1283-1314` writeMcpConfig |

**是否为编辑器扩展：是，且仅是。** 这是 Trae/VS Code 扩展（activityBar 视图 `excel-vba-panel`，`package.json:39-46`），同时它把一个第三方 MCP server（exe）作为子进程管理和分发。

---

## 二、项目意图还原

**项目想做什么**（README.md:1-5，确证）：
在 Trae/VS Code 里提供"Excel VBA 本地同步器 + AI 编程桥接器"——
1. 把 Excel VBE 里的宏代码双向同步到本地文件（`.bas/.cls/.frm/.wks/.wbk`），让 AI 能用普通文件编辑的方式改 VBA；
2. 内嵌 `sbroenne/mcp-server-excel`（230+ 操作），让 AI 直接通过 MCP 操作工作簿（透视表、Power Query、DAX 等）；
3. 插件是 Excel 的启动入口：用户选文件 → 插件经 MCP 打开 Excel → COM 负责同步、MCP 负责高级操作（`.trae/documents/mcp-managed-excel-refactor.md:20-36` 的架构图，确证）。

**入口与启动链路**（代码确证）：
- 激活：`onStartupFinished`（package.json:35）→ `src/extension.ts:68` activate() → 启动 mcp-excel.exe 子进程（extension.ts:100）→ 写 `.trae/mcp.json`（extension.ts:117）→ 同步 skills 到 `.trae/skills/`（extension.ts:127）。
- 用户选文件：`handleSelectWorkbook`（extension.ts:286）→ 冲突检测/接管（332-380）→ `excel.openWorkbook(filePath)`（390，MCP file(open) 自动打开 Excel）→ PID 探测（399）。
- 同步：全部走 `VbaClient`（COM）：`executeSync` → `client.syncVbeToLocal/syncLocalToVbe`（extension.ts:785-786）。
- 断开：`handleDisconnectExcel`（extension.ts:465）→ 关 MCP session → 等进程退出/强杀 → 清状态。

**README 与代码不符处以代码为准**（差异清单见第十一节问题 1）：代码会**自动打开 Excel**（extension.ts:390），README §6/§8/§18 仍写"必须手动先打开"；代码选择文件后**不再自动创建同步目录、不再自动开同步**（extension.ts:316 注释明示），README §6/§13 仍写会自动。

---

## 三、git 历史还原（46 条 commit 完整脉络）

### 3.1 两个 "Initial commit" 的关系
- `2f64b40`（2026-06-29 00:34，作者 `martin94zh <33682349+martin94zh@users.noreply.github.com>`）：只含 2 行 README——这是在 GitHub 网页上建仓库的初始化提交。
- `fed0289`（2026-06-29 00:37，作者 `Developer <developer@example.com>`）：**2f64b40 的直接子提交**（`git log --format="%h parents:[%p]"` 证实），一次性导入 7411 行：完整扩展 + **自建 MCP server（`server/` 目录，含 ExcelVbaServiceClient.ts 516 行、tools.ts 395 行、schemas.ts 359 行）**。
- 结论：不是分叉，是"网页初始化 + 本地批量导入"的正常序列。

### 3.2 06-29 单日阶段划分（43 条 commit，00:34 → 06:24，约 6 小时通宵）

| 阶段 | commit 范围 | 内容 | 结局 |
|---|---|---|---|
| P0 导入 | 2f64b40→fed0289 (00:34-00:37) | 全量代码入库，内部版本 0.1.0 | — |
| P1 发布准备 | 70d9fe4→2e23cd6 | LICENSE、skills 移到顶层、激活时写 MCP 配置；**版本号 0.1.0 直接跳到 0.5.0**（094dd87）；用 pkg 把自建 server 打成 exe（2e23cd6） | v0.5.2 |
| P2 文档对齐 | 09683a8→25bff61 | 3 条纯 README 提交 | — |
| P3 图标主题摇摆 | 35fd3a4→0185221 | 加图标主题(v0.5.3) → **删除**(v0.5.4) → **恢复并加自动应用**(v0.5.5) | 反复推翻自己 |
| P4 关闭检测之役① | 940a717→86654f7 (v0.5.6→v0.6.8，9条) | Excel 关闭检测：PowerShell 5.1 兼容、轮询 5s→2s、COM HWND、PID 类型修复、GetActiveObject 副作用 | 问题没解决 |
| P5 功能扩张 | 8bc4b17→4310050 (v0.7.0→v0.7.2) | Table/标签格式工具 +526 行、skills 自动同步 | — |
| P6 关闭检测之役② | 08c16b2→b12c41e (v0.7.3→v0.7.9，7条) | 弃 COM→窗口句柄→又回 COM workbook 检测；uninstall.ts 诞生 | 仍在绕圈 |
| P7 重写与回退 | 488dae4→9517c55 (v0.8.0→v0.8.2) | **v0.8.0 整体换成窗口标题检测 → v0.8.2 整体回退**（"remove fragile window-title detection"） | 一个通宵内完成一次完整的"另起炉灶+推倒重来" |
| P8 收尾 | a3d199a→c639870 (v0.8.3→v0.8.6，06:24) | 残留进程强杀、宏弹窗交互处理（vbaMacroRunner +387 行）、InputBox 处理 | 06:24 最后一条，通宵结束 |

**高频改动区**：`src/extension.ts`（43 条里约 30 条涉及）+ `src/client/vbaClient.ts` + `src/runtime/powershell.ts`。**核心痛点是"Excel 进程/窗口生命周期检测"**——v0.5.6 到 v0.8.4 约 19 条 commit 在这一件事上反复。

### 3.3 70 天停滞（06-30 → 09-07）的成因还原
间接证据链（仓库内无直接记录，以下为带证据的推断）：
1. `.trae/documents/mcp-managed-excel-refactor.md:476` 显示 **07-01 又一轮通宵**（02:40-12:00 的进度日志）：完成了 T1-T8 全部编码（MCP 托管 Excel、PID 发现、COM 按 PID attach、断开强杀、冲突接管），但**这些工作从未 commit**——git 历史里 06-29 与 09-08 之间为零提交。
2. 同文档 T9（:528-546）：构建/打包通过，但**11 项端到端验证全部"无法自动验证"**——需要真实 Windows + Excel + Trae 环境，AI 环境里没有 `code` CLI 和 Excel。
3. 即：07-01 时功能写完但没人（也没法）验证，工作区带着大量未提交改动被搁置。

### 3.4 09-08 复活（00:56 → 01:36，40 分钟 3 条 commit）
- `7b2dcef` **迁移到 mcp-server-excel**（12518 插入/3888 删除，实际是把 07-01 前后所有未提交工作一次性入库 + 架构切换）：
  - **删除**：整个自建 `server/` 目录（ExcelVbaServiceClient/index/schemas/tools，约 1665 行）——自建 MCP server 时代终结；
  - **新增**：`src/excel/excelClient.ts`（1771 行，166 个 MCP 工具的强类型封装）、`sessionStore.ts`、`vbaSyncEngine.ts`、`src/mcp/serverManager.ts + mcpClient.ts`、`src/native/downloader.ts + processUtils.ts`、`tools/`（下载/枚举/导出 schema）；skills 从 15 篇扩到 41 篇（新增 references/ 21 篇）；
  - **"privacy protection" 的实际内容**（commit message 自述 + diff 证实）：`.gitignore` 从忽略 `.trae/skills/` 改为忽略整个 `.trae/`（因 mcp.json 含本机用户名路径），README 加第三方来源声明。
- `0bf2abd`：README 加 GitHub Release 下载链接（开源分发准备）。
- `e0a610d` "code cleanup and polish"：把 `JrExcelPid/JrVbeWin32/Jr-OpenWorkbook` 等 **"Jr" 前缀符号改名**（vbaMacroRunner.ts:2 头注释"移植并简化自原项目 vbe-macro.ts"暴露了代码源自更早的个人项目），并把 `resources/ui-preview.html:286` 的 **"仅限内部使用，请勿外传"** 改为正常标语——**这一天的主题是开源化准备**。

---

## 四、目录与文件清单（活的 / 死的 / 不确定）

git 跟踪 77 个文件；本地另有未跟踪的 dist/（150MB）、node_modules/、66MB .vsix、.trae/、.vscode/。

| 路径 | 状态 | 说明 |
|---|---|---|
| `src/extension.ts` (1579行) | **活**（核心） | 入口、全部命令与生命周期 |
| `src/client/vbaClient.ts` (1320行) | **活，约 60% 死** | COM 客户端。被调用的只有 8 个方法（checkAccess/getResources/syncVbeToLocal/syncLocalToVbe/listMacros/vbaRun/getVbeCodeChecksum/setWindowTopMost）；工作表/单元格/区域编辑的十余个方法（约 :1072-1278"需求1"区块、openWorkbookInExcel、inspectWorkbook 等）无任何调用者——它们原本服务于已删除的 server/ |
| `src/client/vbaMacroRunner.ts` (919行) | **活** | 宏执行 + Win32 弹窗检测/自动点击（v0.8.5-0.8.6 打磨） |
| `src/excel/excelClient.ts` (1771行) | **活 1%，死 99%** | 166 个 async 封装方法，extension.ts 只调用了 `openWorkbook` 和 `closeWorkbook` 两个（README:83 自述"其余可通过通用 callTool 调用"——即封装层本来就是给 AI 文档用的参考面） |
| `src/excel/vbaSyncEngine.ts` (239行) | **死代码（零引用）** | 基于 MCP vba 工具的同步引擎——正是工程文档 §1.2 宣布"测试多次未成功，已废弃"的方向，却在 7b2dcef 作为新架构一部分入库 |
| `src/excel/sessionStore.ts` (78行) | **边缘存活** | 只被 excelClient 内部使用；extension.ts 从不直接触碰 |
| `src/mcp/serverManager.ts` (116行) | **活** | exe 启动/健康检查/自动重启，含激活时联网下载兜底 |
| `src/mcp/mcpClient.ts` (87行) | **活** | MCP stdio 客户端封装 |
| `src/native/downloader.ts` (145行) | **活（兜底路径）** | 运行时下载 mcp-excel.exe；正常安装包内已带 exe，不会触发 |
| `src/native/processUtils.ts` (338行) | **活** | PID 发现/冲突接管/等待退出/进程树终止 |
| `src/runtime/powershell.ts` (231行) | **活** | PowerShell 执行器 |
| `src/state.ts` (102行) | **活（含死字段）** | `mcpSessionId` 从未被写入真实值（disconnect 清空除外），SessionStore 在内部管 session |
| `src/status/statusBar.ts` (113行) | **活** | 状态栏 |
| `src/webview.ts` (560行) | **活** | 控制面板 UI（本次未逐行审读 UI 文案） |
| `src/uninstall.ts` (52行) | **活** | 卸载时清 MCP 配置 |
| `src/output/outputChannel.ts` (37行) | **活** | 日志 |
| `skills/`（41 个 .md） | **活（功能组成部分）** | AI 指引文档，激活时同步到 `.trae/skills/`；`references/` 21 篇是 Excel 自动化经验库 |
| `tools/download-mcp-server.js` | **活（含 bug）** | 构建时下载 exe；:104 引用未定义常量 `MCP_SERVER_EXE_NAME`（见问题 3） |
| `tools/dump-schemas.js / inspect-tool.js / list-tools.js / schemas.json` | **活（开发工具）** | schemas.json 2692 行是 mcp-server-excel 全部工具 schema 快照 |
| `resources/`（9 文件） | **活** | 图标 + file-icon-theme.json + ui-preview.html（静态设计稿，非运行时 UI） |
| `.vscode/launch.json + tasks.json` | **活（未跟踪）** | F5 调试与 build 任务，配置有效 |
| `.trae/documents/mcp-managed-excel-refactor.md` | **活（最重要的交接文档）** | 上一任 AI 的"唯一真相源"工程文档，含架构决策、废弃方向、T1-T9 进度日志 |
| `.trae/mcp.json` / `.trae/skills/` | **运行时生成（未跟踪）** | mcp.json 含本机用户名路径 `c:\Users\marti\...`（gitignore 已覆盖，不会入库） |
| `dist/`（150MB，未跟踪） | 构建产物 | extension.js 430KB + **mcp-excel.exe 156MB** + **README.md 是 06-29 旧版副本（陈旧）** |
| `excel-vba-assistant-0.8.6.vsix`（66MB，未跟踪） | 构建产物 | 09-08 打包 |
| `scripts/` | **空目录** | 无任何文件 |
| `package.json` 的 devDependency `pkg` | **死依赖** | 唯一使用者 `compile:server:exe` 脚本已随 7b2dcef 删除，依赖声明残留 |
| `server/` 目录 | **已删除** | 仅存在于 git 历史（fed0289→7b2dcef 之前） |

**多套并存版本**：无 v1/v2/old/backup/_legacy 目录，无 xxx_old/xxx_new 文件。历史上唯一"双实现并存"是 COM 同步（活）与 MCP 同步（vbaSyncEngine，死）。

---

## 五、当前可运行性（静态判断，未安装/未启动/未联网）

**能通过静态检查**：
- `npx tsc --noEmit` **零错误**（TypeScript 5.x + SDK 1.29.0，node_modules 已安装且与当前代码匹配，295 个包）。
- 构建工具链齐备：node_modules/.bin 含 tsc/esbuild/vsce。
- F5 调试配置有效（launch.json → preLaunchTask "npm: build"）。
- 打包（`npm run package`）：本地 dist/mcp-excel.exe 已存在时下载脚本会跳过联网（download-mcp-server.js:74-78），可离线打包；exe 不存在时则需要访问 GitHub。

**真实运行的前置条件**（README §3 + 代码）：
Windows 10/11 + 桌面版 Excel（开"信任对 VBA 项目对象模型的访问"）+ Trae/VS Code ≥1.85。非 Windows 直接警告退出（extension.ts:72-74）。

**堵在哪（最大不确定性）**：
- 端到端主流程（选文件 → MCP 自动打开 Excel → COM 双向同步 → 断开清理）**从未被完整验证过**——工程文档 T9 明确记录 11 项验证需要人工在真实环境做，而 T9 之后项目即搁置。
- 本仓库 Windows+Excel+Trae 的真实行为本次无法验证（禁止启动服务/连接 Excel）。
- 结论：**代码可编译、类型干净、打包链完整；"能不能用"取决于从未做过的人工端到端验证。**

---

## 六、半成品与死代码清单

1. **`src/excel/vbaSyncEngine.ts` 整文件死代码**（239 行）：零引用。是被工程文档 §1.2 明确废弃的"MCP 同步"方向的实现，7b2dcef 却连同新架构一起提交。位置：`src/excel/vbaSyncEngine.ts:43`（类定义）。
2. **`src/excel/excelClient.ts` 166 个方法中 164 个无调用者**（1771 行）：extension.ts 只用 openWorkbook/closeWorkbook。属"面向文档的封装面"，对开源者而言是巨大的维护噪声。
3. **`src/client/vbaClient.ts` 约 15 个方法无调用者**：`openWorkbookInExcel`(:128)、`inspectWorkbook`(:230)、`getAllComponentCode`(:296)、`getComponentCode`(:327)、`listSheets`(:977)、`readUsedRange`(:1004)、`readRange`(:1040)、`createSheet`(:1079)、`deleteSheet`(:1106)、`renameSheet`(:1128)、`setCellValue`(:1150)、`setRangeValues`(:1175)、`clearRange`(:1218)、`setCellFormat`(:1242)、`listDialogs/clickDialog/fillDialog`(:958-968)。它们服务的前端（server/tools.ts）已在 7b2dcef 删除，但服务端方法忘了删。
4. **`tools/download-mcp-server.js:104` 引用未定义常量 `MCP_SERVER_EXE_NAME`**：该常量实为 `MCP_SERVER_EXE_CANDIDATES`(:13)。解压后找不到 exe 时会抛 ReferenceError 而非预期错误信息。`src/native/downloader.ts` 是同一逻辑的正确版本（:80 直接返回 message）。
5. **`state.ts:27` `mcpSessionId` 字段**：仅被清空（extension.ts:517），从未被赋真实值——残留设计。
6. **`package.json` devDependency `pkg`**：唯一使用者已删，依赖残留。
7. **无 TODO/FIXME/DEPRECATED 标记、无注释掉的代码块**（全 src/tools 扫描为零）——e0a610d 清理过，表面干净，但上述结构性死代码不是注释能暴露的。
8. **大段注释掉的代码**：无。
9. **空函数/未实现接口**：未发现（tsc strict 通过）。
10. **迁移完成度评估**：7b2dcef 的迁移在"MCP 接通"意义上完成（serverManager/excelClient/PID/冲突接管/断开强杀俱全），在"收敛"意义上未完成——死代码、README 旧文案、鸡肋封装面都是迁移留下的尾巴。

---

## 七、开源就绪度评估（13 项，重点章节）

| # | 项 | 评级 | 结论与证据 |
|---|---|---|---|
| 1 | 敏感信息 | **✅ 已就绪** | 对**全部 46 条 commit** 的 `git grep` 扫描：password/secret/api_key/apikey/PRIVATE KEY/Bearer/mongodb/postgres/mysql/jdbc/192.168./127.0.0.1/localhost/C:\Users 全部无真实命中（"password"命中为 skills 文档讲解 Excel 连接串参数、schemas.json 的 save_password 参数定义；"10.0."命中为版本号文本，当前树为零）。历史上从未提交过 .env/credentials/本地路径文件（`git log --all --name-only` 全集核实，.trae 与 .vscode 从未入库）。作者邮箱：45 条为占位符 `Developer <developer@example.com>`（不泄露隐私，但不专业），1 条为 GitHub noreply（2f64b40，网页初始化）。**无任何"开源即泄露"项，无需重写历史来除敏。** 注意两点非敏感瑕疵：历史中 ui-preview.html 曾含"仅限内部使用，请勿外传"（e0a610d 才改掉，历史仍可见）；vbaMacroRunner 注释提到"原项目 vbe-macro.ts"（自引，无碍）。 |
| 2 | LICENSE | **⚠️ 有缺陷但可修** | 根目录有 MIT LICENSE（70d9fe4 引入）。问题 a：版权行"(c) 2025"，项目实际始于 2026；问题 b：**分发的 .vsix 内嵌第三方 156MB 的 mcp-excel.exe，其上游 sbroenne/mcp-server-excel 的许可证与再分发条款因外网不通无法确认**——这是开源发布前必须线上核实的一项。 |
| 3 | README 质量 | **⚠️ 有缺陷但可修** | 有：项目简介、安装（Release 下载/本地构建）、快速开始、19 个章节的详细用法、错误对照表、安全说明。缺/错：**多处与当前代码矛盾**（见问题 1）；无截图或演示动图（ui-preview.html 只是静态设计稿）；无贡献方式；工具类数自相矛盾（:5 说 16 类，:25/:61 说 17 类，SKILL.md 说 16 类）；正文为中文（目标受众决策，非缺陷）。 |
| 4 | .gitignore | **✅ 已就绪** | node_modules/dist/*.vsix/.trae/.env/.idea 等全忽略；历史核实 dist、node_modules、.vsix、.trae、.vscode **从未入库**；仓库仅 251KB。 |
| 5 | 构建产物 | **✅ 已就绪（git 内）** | `git ls-files dist` = 0。本地 dist/ 150MB 与 66MB vsix 均未跟踪。发布侧 vsix 内嵌 exe 属分发策略而非仓库污染。 |
| 6 | 依赖清单 | **⚠️ 有缺陷但可修** | dependencies 只有 @modelcontextprotocol/sdk（MIT）；devDependencies 齐全但 **pkg 已无用**。package-lock.json（Jun 30）早于 package.json 最后修改（Jul 1），且未入库——可复现性一般。依赖许可兼容性：MIT 系，与项目 MIT 兼容（未逐一在线核验，无风险信号）。 |
| 7 | CI/CD | **❌ 阻塞开源（级别：缺失）** | 无 .github/ 目录，无 workflows，无 lint/test/release 自动化。 |
| 8 | 开源配套 | **❌ 阻塞开源（级别：缺失）** | 无 CONTRIBUTING.md、CODE_OF_CONDUCT.md、Issue/PR 模板（.github/ 不存在）。 |
| 9 | 版本与 tag | **❌ 阻塞开源（级别：缺失）** | `git tag` 为空；版本仅存在于 package.json（0.1.0→0.5.0 跳号，0.2-0.4 从未存在）。README 引用 `releases/download/v0.8.6/...` 链接，**该 release 是否真的存在于 GitHub 因外网不通无法确认**——若不存在，README 的主安装路径是死链。 |
| 10 | commit 质量 | **✅ 已就绪** | 45/46 条为规范 conventional commits 且绝大多数带版本号，信息量大（如 "fix: move EnumWindows logic into C# method to avoid PowerShell callback scoping issue"）；无 "fix"/"update"/"1" 类垃圾提交；两个 Initial commit 可接受。可选优化：占位符作者身份。**无需为开源清理提交历史。** |
| 11 | 项目元信息 | **✅ 基本就绪** | name/displayName/description/repository/bugs/homepage/keywords/license 齐全（package.json:2-27）。瑕疵：publisher 为占位符（上 Marketplace 需注册真实 publisher）；无 author 字段；categories 只有 "Other"。 |
| 12 | 测试 | **❌ 阻塞开源** | 零测试：无 test script、无任何 .test./.spec. 文件。对 Excel COM 这种强环境依赖的项目，至少需要 CI 可跑的单元层（如 manifest 比较逻辑、同步规则纯函数化后的测试）。 |
| 13 | 历史可清理性 | **✅ 低难度** | 46 条、线性、单分支（main）。无敏感内容需清除；若想合并两个 Initial commit、规范化作者或压缩成更干净的叙事，一次 `git filter-repo` / squash 即可完成，半天内可控。唯一代价：若远端已有 release/clone，hash 会变（远端状态无法离线确认）。 |

**开源就绪度总评**：仓库本体干净（无敏感信息、无产物污染、提交历史体面），**可以直接公开而不出安全事故**；但作为"开源项目"的产品化配套（tag/release 确认、CI、CONTRIBUTING、测试、README 修正）缺 4 项 ❌ + 3 项 ⚠️，距离"体面开源"约几天的补课量。

---

## 八、失败原因分析（按证据强度排序）

1. **验证环境无法自动化 → 端到端从未闭环（最直接的停摆原因）**。项目本质是 Windows+Excel COM+VBE+Trae 的强环境耦合件（extension.ts:72-74 平台警告；README §3；工程文档 T9 列出 11 项"无法自动验证"）。AI 能写码能打包，但不能替人点 Excel；T9 之后工作失去推进抓手，带着未提交的改动搁置 68 天。
2. **核心难点（Excel 进程生命周期检测）被证明不可靠，最终放弃**。v0.5.6→v0.8.4 约 19 条 commit 在关闭检测上反复（P4/P6/P7 三个阶段），v0.8.0→v0.8.2 完成一整轮"重写→回退"；终局是工程文档 §4.7 明确宣布"本功能明确不做"，改为用户手动断开。这是最典型的"失败尝试痕迹"。
3. **单日 AI 爆发式开发 + 零测试**。43 commit/6 小时，每条 fix 都是"改一行→bump 版本→重打包"的循环；没有测试护栏，回归只能靠人肉点击。直接后果就是第 2 条的反复震荡（icon theme 在 30 分钟内加了删、删了加：35fd3a4→65b5de9→0185221）。
4. **技术方向摇摆，架构级返工一次**。同步通道在 COM 与 MCP 之间摇摆（工程文档 §1.2 记录 MCP 同步"测试多次未成功，已废弃"，但 vbaSyncEngine 仍被提交）；07-01 又把整个架构从"自建 MCP server"迁到"外部 mcp-server-excel"（删除 ~1665 行、新增 ~4800 行）。每次摇摆都产生一批死代码（第六节 1-3 条）。
5. **目标面过宽**。从"VBA 同步器"膨胀到"230+ 操作的 AI 桥接器"（README §2 的 17 类工具表）。单人+AI 无法验证这么大的面，166 个封装方法只用 2 个就是失衡的直接体现；同时文档跟不上代码（README 描述旧架构行为），用户可感知的"效果不理想"很大程度来自文档承诺与实际行为不一致。

---

## 九、可复用资产清单（即使重写也值得带走）

1. **`.trae/documents/mcp-managed-excel-refactor.md`**：完整架构决策记录（目标架构、状态机、废弃方向及原因、T1-T9 验证清单）——任何后续开发的第一输入。
2. **`skills/references/` 21 篇 + `skills/` 20 篇**：Excel 自动化领域知识库（anti-patterns、gotchas、M 语法、DMV、agent 工作流），与本项目代码解耦，可直接复用于任何 Excel+AI 项目。
3. **`tools/schemas.json`（2692 行）**：mcp-server-excel 全部工具的 schema 快照——接口定义资产。
4. **`src/client/vbaMacroRunner.ts`（919 行）**：Win32 级宏弹窗检测/点击/填表/续跑（EnumWindows+EnumChildWindows+SendMessage），移植自更早的个人项目，是全仓库最独特的技术积累。
5. **`src/client/vbaClient.ts` 的双向同步核心**（:404-895）：VBE↔本地目录映射（模块/类模块/窗体/Microsoft Excel 对象 + workbook.json 清单）、增量比较、删除保护规则（文档对象模块不可删、窗体不可本地新建）——领域规则已打磨清楚。
6. **`src/native/processUtils.ts`**：PID 探测（WMI CommandLine+窗口标题交叉验证）、冲突接管（COM 保存退出→强杀兜底→等待退出）。
7. **中文 README 的骨架**：功能清单、同步规则表、错误对照表、安全说明——改写后即是对外文档。
8. **同步目录约定**（中文目录名与 VBE 原生分类一致，README §1）：面向中文用户的产品差异点。

---

## 十、决策选项

### A. 继续：在现有基础上开发
- **前提条件**：先补一次人工端到端验证（真实 Windows+Excel+Trae 环境跑通 T9 的 11 项清单），否则继续开发是在未验证的地基上加盖。
- **第一步**：装上 vsix 跑通"选文件→自动打开→同步→断开"，把结果记录回工程文档进度日志；随后按第六节清单删死代码、按问题 1 修 README。
- **对开源的影响**：最快达到"可发布"——仓库本体已开源安全；补 tag v0.8.6+release、CONTRIBUTING、CI（至少 build+typecheck）、修 LICENSE 年份即可首发。上游 exe 许可需线上确认。

### B. 局部重写（推荐）
- **保留**：vbaClient 同步核心 + vbaMacroRunner + processUtils + statusBar/webview 骨架 + skills/ 全部 + tools/schemas.json + 工程文档 + README 素材。
- **重写/删除**：删除 vbaSyncEngine、sessionStore、excelClient 的 166 方法封装面（保留 openWorkbook/closeWorkbook 或干脆直连 mcpClient）、vbaClient 的 15 个死方法、pkg 依赖；重写 README 使与代码一致。
- **边界**：不动 COM 同步与弹窗处理这两块已打磨的逻辑，只做"减法 + 文档对齐"。
- **第一步**：一个纯删除+文档修正的 commit（预计删 2500+ 行），让"活代码=已验证代码"；再走 A 的验证与开源补课清单。
- **对开源的影响**：最好——开源者看到的是可信子集，维护面缩小一半以上，且历史保留（无损）。

### C. 彻底重写
- **带走**：第九节 1-8 全部资产（文档、skills、schema、领域规则、弹窗处理思路）。
- **放弃**：全部现有代码与 git 历史（46 条历史本身无需清理，放弃它只为了叙事干净）。
- **第一步**：新仓库，先立 CI+测试框架与最小架构（扩展壳 + MCP 子进程 + 同步引擎三个模块），再逐个移植资产。
- **对开源的影响**：最干净的开源起点，但代价最大——VBE 双向同步与 COM attach 是踩了 19 个 commit 才收敛的硬骨头，重写等于重趟；除非决定换技术路线（例如放弃 COM 双轨、全部走 mcp-server-excel 的 vba 工具——这正是文档记载已失败过的方向），否则不建议。

---

## 十一、问题清单（看不懂的 / 矛盾的 / 缺失的）

**矛盾类**
1. **README vs 代码行为矛盾（4 处）**，以代码为准：
   - README:146/190/442 说必须手动先打开 Excel、插件不会自动启动 Excel；代码会经 MCP 自动打开（extension.ts:390），skills/SKILL.md:33 也明确"do not tell the user to manually open Excel first"。
   - README:147-148/284 说选文件后自动创建同步目录并自动开同步；代码明确不做（extension.ts:316-318 注释"选择文件为独立操作，不再自动创建同步目录、不再自动同步"）。
   - README:442 的错误文案"未检测到 Excel 中已打开该文件"在当前代码中已不存在（全文 grep 无此字符串）。
   - 工具类数：README:5 说 16 类，README:25/:61 说 17 类，SKILL.md:17 说 16 类。
2. **工程文档 vs git 历史**：文档（07-01）声称 excelClient.ts 等当时已存在，但它们 09-08 才首次入库——说明 07-01 的工作长期未提交，git 历史不能当作开发时序的完整记录。
3. **vbaSyncEngine 的存在本身**：文档说该方向已废弃，代码却作为 7b2dcef"新架构"的一部分入库（见第六节 1）。

**缺陷类**
4. `tools/download-mcp-server.js:104` 引用未定义常量 `MCP_SERVER_EXE_NAME`（潜在 ReferenceError）。
5. `excelVba.disconnectExcel` 与 `excelVba.useExcelSameDirectory` 在 extension.ts:194-195 注册，但 **package.json contributes.commands 未声明**——命令面板搜不到，VS Code 会有未注册命令告警。
6. LICENSE 版权年份 2025（项目实际 2026）；package.json 无 author、publisher 为占位符。
7. 死依赖 pkg；package-lock.json（Jun 30）早于 package.json 最后修改（Jul 1），可复现性存疑（lock 未入库）。
8. `state.ts:27` mcpSessionId 从未被赋真实值。
9. `dist/README.md` 是 06-29 旧版副本（陈旧构建残留，未跟踪，无碍 git）。
10. c639870 commit message 提到的 "debug action" 在当前代码无踪迹（迁移时移除），读历史会被误导。

**看不懂/需要人工确认类**
11. T9 的 11 项端到端验证从未被确认完成——**插件实际好不好用，仓库内无任何记录**，这是全项目最大的未知数。
12. `.trae/mcp.json` 指向的安装路径是 `excel-vba-assistant.excel-vba-assistant-0.8.6`（Trae 扩展目录）——说明 0.8.6 曾被真实安装过，但安装后是否验证通过无记录。

**因外网不通无法确认类**
13. 远端 GitHub 仓库状态：是否存在 v0.8.6 release 及 66MB vsix 资产（README:112 的主下载链接是否为死链）、issues/PR/star、远端 main 与本地 `e0a610d` 是否一致（本地显示 up-to-date 基于既有跟踪信息）。
14. 上游 `sbroenne/mcp-server-excel` 的许可证与再分发条款（vsix 内嵌其 156MB exe，开源发布前必须核实）。
15. `tools/download-mcp-server.js` 依赖的 `ExcelMcp-MCP-Server*.zip` 资产命名在最新 release 中是否仍然存在（构建链对上游 release 格式有脆弱依赖）。

---

*报告完。等待决策，未做任何修改。*
