# Excel VBA Assistant

![Version](https://img.shields.io/github/v/release/martin94zh/excel-vba-assistant)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)
![Excel](https://img.shields.io/badge/Excel-2016%2B-green)
![License](https://img.shields.io/badge/license-MIT-yellow)

**在 Trae / VS Code 中让 AI 直接操作真实的 Excel：前台可见、326 种操作、VBA 双向同步、AI 宏测试闭环。**

你选择 Excel 文件，插件把它**前台可见地打开**；AI 通过命令行工具对同一个 Excel 增删改查（数据、表格、图表、透视表、Power Query、DAX、VBA……），你全程看得到每一步；插件同步引擎让 **VBA 代码在 Excel 与编辑器之间双向实时同步**——你在 Excel 里手改宏，编辑器里立刻出现；AI 改宏、跑宏、报错弹窗自动抓取，形成完整的 AI 宏开发闭环。

> 当前版本：**v0.9.2** ｜ [下载安装包](https://github.com/martin94zh/excel-vba-assistant/releases/latest)

---

## 目录

- [功能亮点](#功能亮点)
- [架构与工作原理](#架构与工作原理)
- [快速开始](#快速开始)
- [使用指南](#使用指南)
  - [选择 Excel 文件](#1-选择-excel-文件)
  - [VBE ↔ 本地双向同步](#2-vbe--本地双向同步)
  - [AI 操作 Excel](#3-ai-操作-excel)
  - [AI 宏开发循环](#4-ai-宏开发循环核心特性)
  - [宏运行（控制面板）](#5-宏运行控制面板)
  - [Excel 置顶](#6-excel-置顶)
  - [断开与清理](#7-断开与清理)
- [完整能力清单（326 操作）](#完整能力清单326-操作)
- [控制面板与命令参考](#控制面板与命令参考)
- [环境要求与安全设置](#环境要求与安全设置)
- [常见问题](#常见问题)
- [安全说明](#安全说明)
- [架构决策：为什么用 CLI 而不是 MCP Server](#架构决策为什么用-cli-而不是-mcp-server)
- [开发与测试](#开发与测试)
- [致谢](#致谢)

---

## 功能亮点

| | 功能 | 说明 |
|---|---|---|
| 🖥️ | **前台可见** | Excel 真实打开在桌面上，AI 的每一步写入、格式化、画图你实时看得见（不是黑盒后台操作） |
| 🤖 | **AI 全量操作** | 内置 [ExcelMcp 2.0.8](https://github.com/sbroenne/mcp-server-excel) CLI：**31 组命令、326 操作**，覆盖 VBA、工作表、单元格、格式、表格、透视表、图表、Power Query、DAX、切片器、截图、窗口管理 |
| 🔄 | **VBE ↔ 本地双向同步** | Excel 里手改宏 → 编辑器自动出现；编辑器里改代码 → 写回 Excel。支持自动（3 秒轮询 + 300ms 防抖）与手动两种模式，带删除同步与队列防冲突 |
| 🧪 | **AI 宏测试闭环** | AI 改完宏自动运行：**报错弹窗全文抓取**（运行时错误/编译错误）、**MsgBox/InputBox/确认框自动应答**、运行结果自动读取——AI 根据报错和结果自主修正，全程不卡死、不抢你的键盘 |
| 📌 | **Excel 置顶** | 一键保持 Excel 窗口在最前，边写代码边看运行结果 |
| 📚 | **Skills 自动注入** | 内置 AI 使用文档（含官方 31 组命令手册），激活时自动同步到工作区并注入实际工具路径，AI 零配置上手 |
| 🔒 | **单持有者架构** | 工作簿只被一个后台服务持有，插件、AI、你三方协作互不冲突（[为什么这样设计](#架构决策为什么用-cli-而不是-mcp-server)） |

---

## 架构与工作原理

```
┌──────────────────────────┐   macro-run.json    ┌───────────────────────────┐
│    Trae / VS Code 的 AI   │ ──────────────────▶ │   Excel VBA Assistant 插件  │
│                          │ ◀────────────────── │   · VBE ↔ 本地双向同步 (COM) │
│  通过 excelcli 命令        │  macro-run-result   │   · 宏运行桥（弹窗捕获/应答）  │
│  操作同一工作簿            │  (结果+报错弹窗全文)  │   · Excel 置顶 / 状态栏      │
└────────────┬─────────────┘                     └────────────┬──────────────┘
             │ excelcli -q ... --session <id>                 │ COM 附着
             ▼                                                ▼
┌─────────────────────────────────────────────────────────────────────────┐
│              excelcli daemon（共享后台服务，工作簿唯一持有者）                  │
│                          Excel —— 前台可见，人人可编辑                       │
└─────────────────────────────────────────────────────────────────────────┘
```

**三个角色，一个持有者：**

1. **excelcli daemon**（来自开源项目 [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) 2.0.8）：后台常驻服务，**独占持有工作簿**。所有 `excelcli` 进程都连接到它，会话跨进程持久——这是插件与 AI 能协作同一工作簿的关键。
2. **AI**：在你的 Trae / VS Code 对话框里提需求，AI 执行 `excelcli` 命令操作 daemon 里的同一会话——326 种操作随便用。
3. **插件**：选择文件时让 daemon 前台可见地打开 Excel；随后通过 **COM 附着**同一个 Excel 实例，提供 AI 做不到的事——VBE 双向同步、带弹窗处理的宏运行桥、窗口置顶。

> 完整的决策依据见[架构决策](#架构决策为什么用-cli-而不是-mcp-server)：为什么不使用 MCP Server 形态。

---

## 快速开始

### 安装

**方式 A：下载安装包（推荐）**

1. 下载 [excel-vba-assistant-0.9.2.vsix](https://github.com/martin94zh/excel-vba-assistant/releases/download/v0.9.2/excel-vba-assistant-0.9.2.vsix)（内置 excelcli.exe，**无需联网、无需 Node.js**）
2. Trae / VS Code 中按 `Ctrl+Shift+P` → **Extensions: Install from VSIX...** → 选择文件
3. 重新加载窗口

**方式 B：本地构建**

```powershell
git clone https://github.com/martin94zh/excel-vba-assistant.git
cd excel-vba-assistant
npm install
npm run package    # 生成 excel-vba-assistant-0.9.2.vsix
```

### 六步上手

1. 左侧活动栏出现 **Excel VBA** 图标 → 点击打开控制面板
2. 点击「**选择 Excel 文件**」，选择 `.xlsm` / `.xlsb` / `.xlam` 文件
3. 插件通过 excelcli **前台可见地打开 Excel**，状态栏变绿 ✅
4. 点击「**Excel相同目录**」创建同步目录（或手动选择），执行 **VBE → 本地 同步**
5. 打开「**自动同步**」开关，按需开启「**Excel 置顶**」
6. 在 AI 对话框里直接提需求，例如：
   - *"把 Sheet1 的 A1:D1 填入表头并做成表格"*
   - *"在 Module1 里写一个统计 B 列平均值的宏，运行它，把结果显示在 C1"*

   AI 会通过 excelcli 操作**同一个前台 Excel**——你看着它干活。

---

## 使用指南

### 1. 选择 Excel 文件

支持 `.xlsm` / `.xlsb` / `.xlam`。控制面板「选择文件」按钮或命令面板 `Excel VBA: 选择 Excel 文件`。

选择时插件自动检测文件状态：

| 状态 | 处理 |
|------|------|
| 未打开 | 插件通过 excelcli **前台可见地打开**，AI 可立即操作（推荐） |
| 已在其他 Excel 实例打开 | 弹窗选择：**「关闭并让 AI 重新打开」**（AI 可完整操作）或 **「保持打开」**（仅本地同步，AI 无法操作——独占访问限制） |

> **为什么必须独占？** Excel COM 层限制：一个工作簿同一时间只能被一个自动化持有者打开。这是上游开源项目的官方约束，无法绕过。

### 2. VBE ↔ 本地双向同步

同步目录结构（文件夹名与 VBE 资源管理器原生分类一致）：

```
syncDirectory/
├─ workbook.json                            清单
├─ Microsoft Excel 对象/ThisWorkbook.wbk    工作簿对象模块
├─ Microsoft Excel 对象/Sheet1.wks          工作表对象模块
├─ 模块/Module1.bas                         标准模块
├─ 类模块/Class1.cls                        类模块
└─ 窗体/UserForm1.frm                       窗体（含 .frx）
```

**两个方向：**

- **VBE → 本地（同步至目录）**：导出全部 VBA 组件为 `.bas` / `.cls` / `.frm` / `.wks` / `.wbk`。VBE 里删除了组件，本地文件同步删除。
- **本地 → VBE（写回Excel）**：本地代码覆盖写回 VBE（弹确认框）。本地删除文件 = VBE 删除组件；本地新建 `.bas` / `.cls` = VBE 新建模块。

**同步规则：**

| 目录 | 行为 |
|------|------|
| `模块/*.bas` | 存在则覆盖；不存在则新建；本地删除则 VBE 删除 |
| `类模块/*.cls` | 同上 |
| `Microsoft Excel 对象/*.wks` / `*.wbk` | **只更新事件代码**，不新建/删除工作表或工作簿（受保护） |
| `窗体/*.frm` | 只更新事件代码，不新建窗体（布局需在 VBE 中设计）；本地删除则 VBE 删除 |

**自动同步（双向）：**

- 本地文件保存 → 300ms 防抖 → 自动写回 VBE
- VBE 代码变更 → 每 3 秒 checksum 检测 → 自动导出到本地
- 变更进入队列按序处理，避免双向互相覆盖
- 内容比对**区分大小写**，未变化的文件不触发反向同步

> 建议配合编辑器自动保存使用（`files.autoSave: "afterDelay"`，延迟 500ms），停止输入约 1 秒即可完成同步。担心误改可保持关闭，使用手动同步。

### 3. AI 操作 Excel

**零配置**：插件激活时自动把两组 Skill 文档同步到工作区 `.trae/skills/`，并注入 excelcli 的实际路径——AI 打开对话即可工作：

- **excel-vba-assistant**：插件编排说明（会话模型、宏开发循环、同步边界）
- **excel-cli**：官方 31 组命令 / 326 操作完整手册（含工作流、反模式、批量模式）

AI 的典型操作（都作用在**你眼前的同一个 Excel** 上）：

```powershell
excelcli -q session list                                              # 找到插件已打开的会话
excelcli -q sheet list --session <id>                                 # 发现工作表
excelcli -q range set-values --session <id> --sheet Sheet1 --range A1:B1 --values '[["姓名","年龄"]]'
excelcli -q table create --session <id> --table-name Sales --range-address A1:B10
excelcli -q pivottable create --session <id> ...
excelcli -q chart create-from-range --session <id> ...
excelcli -q vba view --session <id> --module-name Module1             # 读宏
excelcli -q vba update --session <id> --module-name Module1 --vba-code-file xxx.bas
excelcli -q screenshot capture --session <id> ...                     # 截图验证
```

想在终端手动验证？控制面板的「**复制 excelcli 路径与常用命令**」一键复制完整路径模板。

### 4. AI 宏开发循环（核心特性）

AI 修改宏后需要**运行、看报错、读结果、再修正**。本插件为此提供**宏运行桥**——这是与"直接用 CLI 跑宏"的本质区别：

> 直接用 `excelcli vba run` 跑一个会报错的宏：VBA 错误弹窗会阻塞后台服务，超时后**会话被销毁、Excel 退出、AI 拿不到任何报错文本**。宏运行桥解决这一切。

**协议**：AI 向同步目录写 `macro-run.json`，插件用带弹窗处理引擎的 COM 通道执行，结果写回 `macro-run-result.json`。

```json
{
  "macro": "Module1.ProcessData",
  "timeoutMs": 30000,
  "dialogMode": "auto",
  "confirmButton": "是",
  "inputValue": "42"
}
```

| 字段 | 说明 |
|------|------|
| `macro` | 要运行的过程，格式 `模块名.过程名` |
| `timeoutMs` | 运行超时（默认 45000） |
| `dialogMode` | `auto`（默认）：自动处理全部弹窗；`errors`：仅自动结束报错弹窗 |
| `confirmButton` | 宏内"是/否"确认框点击的按钮（默认"取消"安全值） |
| `inputValue` | 宏内 InputBox 自动填入的文本 |

**弹窗自动处理矩阵：**

| 弹窗类型 | 自动动作 | AI 收到什么 |
|---------|---------|------------|
| 运行时错误（如 `除数为零`） | 自动点"结束" | 报错全文：`运行时错误 '11': 除数为零` |
| 编译错误（如调用不存在的过程） | 自动点"确定" + **自动重置 VBE 中断模式** | 报错全文：`编译错误: 子过程或函数未定义` |
| MsgBox（信息提示） | 自动点"确定"，**宏继续执行** | 弹窗全文（用于验证宏的提示行为） |
| InputBox | 自动填入 `inputValue` 并提交 | 已填入的值 |
| 是/否确认框 | 自动点 `confirmButton`（默认取消） | 弹窗全文 + 所点按钮 |
| 调用堆栈窗口（编译错误的伴生窗口） | 自动关闭 | — |
| "重新设置工程"确认框 | 自动点"确定" | — |

**结果文件**（`macro-run-result.json`）：

```json
{
  "finishedAt": "...",
  "macro": "Module1.ProcessData",
  "success": false,
  "message": "检测到弹窗：Microsoft Visual Basic: 运行时错误 '11': 除数为零",
  "dialogs": [
    { "kind": "vb_runtime_error", "title": "Microsoft Visual Basic",
      "text": "运行时错误 '11':\n\n除数为零", "autoHandled": true, "autoAction": "结束" }
  ]
}
```

AI 拿到报错全文 → 修正代码 → 重新运行 → `range get-values` 读取结果 → 循环直到通过。
另外，最近 20 条错误弹窗始终保留在同步目录 `excel-dialogs.json`，插件宿主也会弹 VS Code 通知。

> 全程**不使用键盘模拟**：按钮点击走 Win32 消息（BM_CLICK）与 UIAutomation，VBE 重置走 COM 菜单命令——**绝不干扰你正在前台的编辑操作**。

### 5. 宏运行（控制面板）

不经过 AI 也可以手动跑宏：

1. 命令面板 → `Excel VBA: 运行宏`
2. QuickPick 列出所有 `模块.过程名`，选择执行
3. 同样具备弹窗检测与自动处理能力，日志输出到输出通道

**自动执行 VBA**（可选）：开启后，AI 修改代码并同步后会自动提示运行指定宏（仍需确认）：

```json
{ "excelVba.autoRunVba": true, "excelVba.autoRunMacroName": "Module1.Main" }
```

### 6. Excel 置顶

控制面板「Excel 置顶」开关：Excel 主窗口始终保持在最前，方便边写代码边看运行结果。断开 Excel 或重置服务时自动取消置顶。

### 7. 断开与清理

控制面板「**断开 Excel**」：

- 通过 excelcli 保存并关闭工作簿（`session close --save`，COM 兜底强制结束）
- 取消置顶、从工作区移除同步目录、删除自动创建的目录

> 插件不会感知 Excel 被外部直接关闭。关闭 Excel 或切换工作簿前，请先点「断开 Excel」。会话若已失效，重新选择文件即可恢复。

---

## 完整能力清单（326 操作）

AI 可用的 31 组命令（另有 `session` / `batch` / `service` / `diag` 基础命令）：

| 命令组 | 主要能力 |
|---|---|
| `range` / `rangeedit` / `rangeformat` / `rangelink` | 读写值/公式、清除、复制、插入删除行列、查找替换、排序、超链接、数字格式、样式、数据验证、合并、锁定、自动调整 |
| `sheet` / `worksheetstyle` | 工作表增删改、复制/移动、标签颜色、可见性、样式 |
| `table` / `tablecolumn` | Excel 表格生命周期、样式、汇总、筛选、排序、列管理、DAX-backed 表 |
| `pivottable` / `pivottablefield` / `pivottablecalc` | 透视表创建/字段布局/值字段函数/计算字段与成员/布局/刷新 |
| `chart` / `chartconfig` | 图表创建/系列/类型/标题/坐标轴/数据标签/趋势线/定位 |
| `powerquery` / `querytable` | Power Query M 代码、刷新、加载配置；QueryTable 管理 |
| `datamodel` / `datamodelrelationship` | 数据模型表/度量值/关系、DAX 查询、DMV 元数据 |
| `vba` | VBA 组件 list/view/import/update/delete + 运行宏 |
| `namedrange` / `connection` / `slicer` / `conditionalformat` | 命名区域、OLEDB/ODBC 连接、切片器、条件格式规则 |
| `screenshot` / `window` / `drawing` / `xmlmap` | 截图、窗口管理、绘图对象与迷你图、XML 映射 |
| `calculationmode` / `analysis` / `workbook` / `pythoninexcel` | 计算模式、Goal Seek/方案分析、工作簿元数据、Python in Excel |

> 全部通过 `excelcli -q <命令组> <动作> --session <id> ...` 调用，完整语法见插件内置的官方 excel-cli Skill 文档；10+ 条命令可用 `batch --input commands.json` 批量执行。

---

## 控制面板与命令参考

**控制面板**（活动栏 Excel VBA 图标）：

| 区域 | 元素 | 说明 |
|------|------|------|
| 路径 | 选择文件 / 断开 Excel | 打开与断开（断开=保存关闭+清理） |
| 路径 | 选择目录 / Excel相同目录 | 设置同步目录（后者自动在 Excel 同目录建同名文件夹） |
| 同步操作 | 同步至目录 / 写回Excel | VBE → 本地 / 本地 → VBE（写回需确认） |
| 配置选项 | 自动同步 / 自动执行 VBA / Excel 置顶 | 三个开关 |

**命令面板**（`Ctrl+Shift+P` 搜 "Excel VBA"）：

`选择 Excel 文件`、`选择本地同步目录`、`使用当前工作区作为同步目录`、`使用 Excel 相同目录`、`断开 Excel`、`VBE → 本地 同步`、`本地 → VBE 同步`、`切换自动同步`、`切换自动执行 VBA`、`运行宏`、`刷新资源`、`打开输出日志`、`同步 Skill 到工作区`、`复制 excelcli 路径与常用命令`

---

## 环境要求与安全设置

- **操作系统**：Windows 10 / 11
- **Excel**：桌面版 Office 2016 及以上
- **Trae / VS Code**：1.85 及以上

VBA 项目对象模型默认禁止程序访问，必须开启：

1. Excel → **文件** → **选项** → **信任中心** → **信任中心设置**
2. **宏设置** → 勾选 **「信任对 VBA 项目对象模型的访问」**

未开启时，所有 VBA 读写操作都会报错。

---

## 常见问题

**Q：安装插件后 AI 就能操作 Excel 吗？**
可以，零配置。激活时 Skill 文档（含工具实际路径）自动同步到 `.trae/skills/`；选择 Excel 文件后 AI 即可操作。

**Q：没有安装 Node.js 能用吗？**
可以。excelcli 已打包为独立 exe，无需 Node.js。

**Q：AI 改宏时 VBE 弹报错对话框，会卡死吗？**
不会。报错弹窗被自动抓取全文并点击"结束/确定"，编译错误还会自动重置 VBE 中断模式；报错内容写进结果文件供 AI 修正。宏里的 MsgBox/InputBox 也会被自动应答（内容记录在案）。

**Q：AI 跑宏会影响我在前台打字吗？**
不会。弹窗处理全部使用 Win32 消息（BM_CLICK）与 UIAutomation，VBE 重置走 COM 菜单命令——全程零键盘模拟。

**Q：我手改的 VBA 会丢吗？**
不会。开启自动同步后，VBE 变更每 3 秒导出到本地；你也可以随时手动「同步至目录」。建议定期同步作为备份。

**Q：为什么 AI 说"无法打开文件，已被占用"？**
工作簿是独占持有的。文件已在 Excel 中打开（手动双击打开的）时 AI 无法接管——在插件提示中选择"关闭并让 AI 重新打开"，或手动关闭后让 AI 打开。

**Q：.vsix 里有什么？**
`dist/extension.js`（扩展主程序）、`dist/excelcli.exe`（Excel CLI，175MB 级自包含可执行文件）、`skills/`（AI 文档）、图标与资源。源码已被排除。

---

## 安全说明

以下操作会修改工作簿或执行代码，AI 侧遵循"先确认再执行"约定，插件侧提供多层保护：

- **本地 → VBE 覆盖同步**：弹确认对话框
- **自动执行 VBA**：每次执行前弹确认
- **文档对象模块保护**：`Sheet*.wks` / `ThisWorkbook.wbk` 永远不会被同步删除或新建
- **窗体保护**：窗体不会从本地新建（布局无法还原），只能更新已有窗体代码
- **危险命令清单**（`vba(update/import/delete/run)`、`worksheet(delete/rename)`、`range(set-values/clear-*)`、`table/pivottable/chart(create/delete)`、`powerquery`、`datamodel` 写操作）已写入 AI 文档，要求执行前确认

建议定期执行 VBE → 本地同步作为备份，配合 git 做版本管理。

---

## 架构决策：为什么用 CLI 而不是 MCP Server

上游 `sbroenne/mcp-server-excel` 同时提供 MCP Server 与 CLI 两种形态。本项目**实测验证**后选择 CLI：

| | MCP Server | CLI（本插件采用） |
|---|---|---|
| 会话归属 | 每个客户端进程私有（进程内服务） | **共享 daemon 单例，跨进程持久** |
| 插件与 AI 协作同一工作簿 | ❌ 数学上不可能（实测：B 进程 `session list` 看不到 A 的会话，打开被独占拒绝） | ✅ 天然支持 |
| 官方文档依据 | "hosts the ExcelMcp Service in-process" + "files must not be open in another instance" | "background daemon so workbook sessions persist across commands" |

`mcp-server-excel` 对工作簿要求**独占访问**（Excel COM 层限制），且 MCP over stdio 时每个 AI 客户端各起一个 server 进程——"插件打开 + AI 操作"必然变成两个持有者互相阻塞。CLI 的 daemon 模式天然单实例，完美匹配"人 + AI + 插件"三方协作。

本插件的宏运行桥则补齐了官方 CLI 的空白：官方 `vba run` 不处理任何弹窗（实测：运行报错宏会卡死后台服务导致会话销毁、Excel 退出）。

---

## 开发与测试

```powershell
npm install
npm run build        # tsc + esbuild 打包 dist/extension.js
npm run package      # 构建 + 下载 excelcli + 打包 vsix
```

**端到端测试**（真实 Excel + 真实代码路径，非模拟）：

```powershell
# 编译测试台
npx esbuild tools/e2e-harness.ts --bundle --platform=node --format=cjs --outfile=build/e2e-harness.js
npx esbuild tools/e2e-bridge.ts --bundle --platform=node --format=cjs --outfile=build/e2e-bridge.js

# 打开测试工作簿后运行（10 项：双向同步闭环/置顶/PID 定位/checksum）
node build/e2e-harness.js <workbook> <syncDir>
# 19 项：运行时错误/编译错误/正常宏/MsgBox/InputBox/确认框（全程 Excel 存活断言）
node build/e2e-bridge.js <workbook> <syncDir>
```

当前测试状态：**同步套件 10/10**、**桥接套件 19/19**。`tools/probe-dual-session.js` 与 `tools/probe-inputbox.ts` 是架构验证与问题定位探针。

**调试**：F5 启动扩展开发宿主；日志见命令面板 `Excel VBA: 打开输出日志`。

---

## 致谢

- [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) — 本插件全部 Excel 自动化能力的来源（ExcelMcp 2.0.8 CLI，31 组命令 / 326 操作），以及官方 excel-cli Skill 文档
- 上游采用 MIT 许可证，本项目同样以 [MIT](LICENSE) 开源

---

## 许可证

MIT © 2025-2026
