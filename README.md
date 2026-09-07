# Excel VBA Assistant

Excel VBA 本地同步器 + AI 编程桥接器：在 Trae / VS Code 中管理、同步、编辑、运行和 AI 辅助开发 Excel VBA 项目。

本插件基于 [`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel) MCP 服务器构建，所有 Excel/VBA 操作均通过 Model Context Protocol (MCP) 完成，保留原有 UI 与交互习惯的同时，完整继承 `mcp-server-excel` 的 **16 类工具、230+ 操作**，覆盖 VBA、工作表、单元格、格式、Excel 表格、数据透视表、图表、Power Query、DAX 数据模型、命名区域、数据连接、切片器、条件格式、截图、窗口置顶等全部能力。

> 当前版本：**v0.8.6**

---

## 1. 项目介绍

### MCP 核心来源

本插件的 Excel/VBA 自动化能力**完全基于 [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)** 构建。`mcp-server-excel` 是一个开源的 Model Context Protocol (MCP) 服务器，通过 COM 接口与 Excel 桌面版交互，提供了 16 类工具、230+ 操作。

本插件在其基础上增加了：
- VBE ↔ 本地文件双向同步
- 自动同步与文件监听
- 可视化控制面板与状态栏
- AI Skill 指引文档
- VBA 宏运行与管理
- Excel 窗口置顶

**核心 MCP 工具（17 类，230+ 操作）均来自 `mcp-server-excel`**，包括：`file`、`worksheet`、`range`、`table`、`pivottable`、`chart`、`powerquery`、`datamodel`、`vba`、`slicer`、`conditionalformat`、`screenshot`、`window` 等。

### 核心能力

本插件提供以下核心能力：

- **VBE ↔ 本地双向同步**：把 Excel VBE 中的模块/类模块/窗体/工作表对象代码导出为本地 `.bas` / `.cls` / `.frm` / `.wks` / `.wbk` 文件；反之把本地代码覆盖写回 VBE。
- **删除同步**：在 VBE 删除组件后同步到本地会删除对应文件；在本地删除文件后同步到 VBE 会删除对应组件（工作表/工作簿对象模块始终受保护）。
- **双向自动同步**：开启后，本地文件保存自动写回 VBE，VBE 代码变更自动导出到本地。变更按时间顺序排队处理，避免冲突覆盖。
- **本地文件监听**：本地 VBA 文件保存后防抖 300ms 自动同步到 VBE。
- **VBE 轮询检测**：每 3 秒检测一次 VBE 代码变更，自动导出到本地。
- **宏执行**：列出工作簿中的所有 Sub/Function，运行无参数宏。
- **完整 Excel 自动化**：通过 MCP 调用 `mcp-server-excel` 的 230+ 操作，几乎覆盖 Excel 桌面版的全部可编程对象。
- **状态栏**：在 VS Code 状态栏实时显示连接/同步/错误状态。
- **输出日志**：所有操作记录到 OutputChannel，便于排查问题。
- **MCP Server**：集成 `mcp-server-excel`，选择 Excel 文件并同步后自动向工作区 `.trae/mcp.json` 写入配置，Trae AI 可直接调用 Excel/VBA 工具。
- **AI 指引（Skills）**：扩展内置 `skills/` 文档，AI 调用 MCP 工具前会先读取对应操作说明。

### 本地同步目录结构

```
syncDirectory/
├─ workbook.json                            清单
├─ Microsoft Excel 对象/ThisWorkbook.wbk    工作簿对象模块
├─ Microsoft Excel 对象/Sheet1.wks          工作表对象模块
├─ 模块/Module1.bas                         标准模块
├─ 类模块/Class1.cls                        类模块
└─ 窗体/UserForm1.frm                       窗体（含 .frx）
```

> 文件夹名与 VBE 资源管理器原生分类保持一致。插件自带「Excel VBA Icons」文件图标主题，启用后 `.bas` / `.cls` / `.frm` / `.wks` / `.wbk` 和上述文件夹会显示不同图标。首次安装时若未设置图标主题，插件会自动启用；也可手动通过 `Ctrl+Shift+P` → `Preferences: File Icon Theme` → 选择 `Excel VBA Icons` 切换。

---

## 2. 完整能力清单（基于 mcp-server-excel）

插件通过 MCP 暴露了 `mcp-server-excel` 的全部 17 类工具，共 230+ 操作：

| 工具类别 | 主要能力 | 操作数 |
|---|---|---|
| `file` | 打开、创建、关闭、测试工作簿，管理会话 | 6 |
| `calculation_mode` | 获取/设置计算模式，触发重新计算 | 3 |
| `worksheet` | 工作表增删改、复制/移动、标签颜色、可见性 | 16 |
| `range` | 读写值/公式、清除、复制、插入删除行列、查找替换、排序、超链接、数字格式 | 46 |
| `range_format` | 单元格格式、样式、数据验证、合并单元格、锁定、自动调整 | — |
| `table` | Excel 表格创建/删除/重命名/调整大小、样式、汇总行、筛选、排序、列管理、DAX-backed 表 | 27 |
| `pivottable` | 数据透视表创建/删除、字段布局、值字段函数、计算字段/成员、布局格式、刷新 | 30 |
| `chart` | 图表创建/删除/移动、系列管理、类型、标题、坐标轴、数据标签、趋势线、定位 | 29 |
| `powerquery` | Power Query 查询创建/更新/重命名/删除、刷新、加载配置、M 代码执行 | 12 |
| `datamodel` | 数据模型表/列/度量值/关系管理、DAX 查询、DMV 元数据查询 | 19 |
| `namedrange` | 命名区域的增删改查 | 6 |
| `connection` | OLEDB/ODBC 数据连接的创建/测试/刷新/属性管理 | 9 |
| `vba` | VBA 组件的列出、查看、导入、更新、删除、运行宏 | 6 |
| `slicer` | 数据透视表切片器和表格切片器的创建/设置/删除 | 8 |
| `conditionalformat` | 条件格式规则（单元格值、表达式、色阶、数据条、图标集） | 2 |
| `screenshot` | 截取区域或工作表为 PNG 图片 | 2 |
| `window` | Excel 窗口显示/隐藏、置顶、状态、位置、排列、状态栏 | 9 |

> 以上操作全部通过 MCP 调用 `mcp-server-excel` 完成。常用操作在插件代码中有类型化封装，其余操作可通过通用 `callTool` 直接调用。

---

## 3. 环境要求

- **操作系统**：Windows 10 / 11
- **Excel**：已安装的桌面版 Excel（Office 2016 及以上推荐），并已打开目标工作簿
- **Trae / VS Code**：1.85 及以上

---

## 4. Excel 安全设置

VBA 项目对象模型默认不允许程序访问，必须先开启：

1. 打开 Excel → **文件** → **选项**
2. **信任中心** → **信任中心设置**
3. **宏设置** → 勾选 **「信任对 VBA 项目对象模型的访问」**
4. 点击确定保存

未开启时，所有读写 VBA 的操作都会返回错误。

---

## 5. 安装插件

### 方式 A：从 .vsix 安装（推荐）

1. 下载 `excel-vba-assistant-0.8.6.vsix`
2. 在 Trae / VS Code 中执行 `Extensions: Install from VSIX...`
3. 选择该文件安装

> 安装包已内置 `mcp-excel.exe`（来自 `sbroenne/mcp-server-excel`），**无需联网下载**，也无需用户本机安装 Node.js。

### 方式 B：本地构建

```powershell
cd excel-vba-assistant
npm install
npm run package
```

构建产物：

- `dist/extension.js` — VS Code 扩展主入口
- `dist/uninstall.js` — 卸载清理脚本
- `dist/mcp-excel.exe` — 已嵌入的 MCP Server

---

## 6. 快速开始

1. 安装插件后，左侧活动栏会出现 **Excel VBA** 图标
2. **先在 Excel 中手动打开目标工作簿**
3. 点击图标打开控制面板
4. 点击「选择 Excel 文件」，选择已打开的 `.xlsm` / `.xlsb` / `.xlam` / `.xls` 文件
5. 插件会确认文件已打开，自动创建同步目录并执行 **VBE → 本地 同步**
6. 同步完成后，自动开启自动同步，并自动写入 MCP 配置
7. 在 Trae 的 AI 对话中即可直接调用 `file`、`range`、`vba`、`table` 等 MCP 工具

> 插件不会自动打开 Excel，也不会自动检测 Excel 是否关闭。关闭 Excel 后请手动点击控制面板「断开 Excel」清理状态。

---

## 7. 控制面板说明

控制面板位于活动栏「Excel VBA」视图，包含以下区域：

### 路径

- **Excel 文件**：当前连接的工作簿路径
- **VBA 同步目录**：本地同步目录路径
- **选择文件**：选择 Excel 工作簿
- **断开 Excel**：强制重置服务状态，取消置顶，移除工作区目录，清理 MCP 配置
- **选择目录**：手动选择本地同步目录
- **Excel相同目录**：使用 Excel 文件同路径下的同名文件夹作为同步目录

### 同步操作

- **同步至目录**：把 VBE 代码导出到本地同步目录
- **写回Excel**：把本地代码覆盖写回 VBE（会弹确认对话框）

### 配置选项

- **自动同步**：本地保存后自动写回 VBE；VBE 变更后自动导出到本地
- **自动执行 VBA**：AI 修改代码后自动提示执行指定宏
- **Excel 置顶**：保持 Excel 窗口始终在最前

---

## 8. 如何选择 xlsm 文件

支持格式：`.xlsm` / `.xlsb` / `.xlam` / `.xls`

两种方式：

1. **控制面板按钮**：点击活动栏 Excel VBA 图标 → 控制面板中点击「选择 Excel 文件」
2. **命令面板**：`Ctrl+Shift+P` → `Excel VBA: 选择 Excel 文件`

选择后会弹出文件选择对话框。**请确保选中的文件已在 Excel 中手动打开**。如果未检测到该文件在 Excel 中打开，插件会提示你手动打开后重试，不会自动启动 Excel。

---

## 9. 如何选择同步目录

三种方式：

1. **控制面板按钮**：「选择本地同步目录」或「Excel相同目录」
2. **命令面板**：
   - `Excel VBA: 选择本地同步目录` — 弹出目录选择对话框
   - `Excel VBA: 使用当前工作区作为同步目录` — 直接使用当前打开的工作区

建议选择一个空目录或专用的 VBA 代码目录，避免与其他项目混淆。

---

## 10. 如何执行 同步至目录（VBE → 本地）

把 Excel VBE 中的所有 VBA 组件导出到本地同步目录：

1. 确保目标工作簿已在 Excel 中打开
2. 控制面板点击「同步至目录」按钮，或命令面板执行 `Excel VBA: VBE → 本地 同步`
3. 等待状态栏显示同步成功
4. 在同步目录中查看生成的 `模块/`、`类模块/`、`窗体/`、`Microsoft Excel 对象/` 和 `workbook.json`

**删除同步**：如果 VBE 中删除了某个标准模块/类模块/窗体，执行 VBE → 本地 同步后，本地对应文件也会被删除。

---

## 11. 如何执行 写回Excel（本地 → VBE）

把本地代码覆盖写回 VBE：

1. 确保目标工作簿已在 Excel 中打开
2. 控制面板点击「写回Excel」按钮，或命令面板执行 `Excel VBA: 本地 → VBE 同步`
3. **会弹出确认对话框**（此操作会覆盖 VBE 中的代码，需用户确认）
4. 确认后执行同步

**同步规则**：

| 目录 | 行为 |
|------|------|
| `模块/*.bas` | VBE 中存在则覆盖更新；不存在则新建标准模块；本地删除则 VBE 中删除 |
| `类模块/*.cls` | VBE 中存在则覆盖更新；不存在则新建类模块；本地删除则 VBE 中删除 |
| `Microsoft Excel 对象/*.wks` / `*.wbk` | 只更新已存在的 Excel 对象模块事件代码，不会新建/删除工作表或工作簿对象 |
| `窗体/*.frm` | 只更新已存在窗体的事件代码部分，不会新建窗体；本地删除则 VBE 中删除；窗体布局需在 VBE 中操作 |

---

## 12. 如何开启自动同步

开启后：

- 本地 `*.bas` / `*.cls` / `*.frm` / `*.wks` / `*.wbk` 文件**保存到磁盘后**自动写回 VBE
- VBE 代码变更后每 3 秒检测一次，自动导出到本地
- 所有变更进入队列，按时间先后顺序处理，避免双边反复覆盖

开启方式：

1. **控制面板开关**：点击「自动同步」开关
2. **命令面板**：`Excel VBA: 切换自动同步`

也可在 `settings.json` 中配置：

```json
{
  "excelVba.autoSync": true
}
```

### 让 Trae 自动保存

文件监听只能感知磁盘文件，因此编辑后**必须保存**才会触发同步。建议开启 VS Code/Trae 的自动保存：

按 `Ctrl+Shift+P` → `Preferences: Open User Settings (JSON)`，添加：

```json
{
  "files.autoSave": "afterDelay",
  "files.autoSaveDelay": 500
}
```

这样停止输入约 500ms 后文件自动保存，再经过 300ms 防抖即可同步到 VBE，整体延迟约 1 秒。

**注意**：自动同步使用静默模式，不会弹确认对话框；如果担心误改，建议保持关闭，改用手动同步。

---

## 13. 同步目录自动管理

### 选择 Excel 文件时自动创建同步目录

选择 Excel 文件后，如果尚未设置同步目录，插件会自动在 Excel 文件同路径下创建一个与文件名同名的文件夹作为同步目录：

- 例如选择 `D:\work\book.xlsm` → 自动创建 `D:\work\book\`
- 同步目录会被自动添加到当前工作区，并在资源管理器中打开

### 修改同步目录时自动清理

如果同步目录是插件自动创建的，当用户**手动选择新的同步目录**时，旧目录会被自动删除，避免多地冗余。

如果旧目录中有文件，迁移到新目录后再删除旧目录。

### 手动断开 Excel

插件不再自动检测 Excel 是否关闭。关闭 Excel 或切换工作簿前，请点击控制面板中的「断开 Excel」：

- 关闭 MCP session
- 取消 Excel 窗口置顶
- 从工作区移除同步目录
- 删除自动创建的同步目录
- 清理 MCP 配置中的环境变量（保留 server 条目）

如果你希望长期保留解析文件，请手动选择同步目录（而非使用自动创建的目录），或点击「断开 Excel」前手动复制目录。

---

## 14. 常见问题

### 为什么 .vsix 里没有源码？

`.vsix` 只包含运行所需的文件：

- `dist/extension.js` — 扩展主程序
- `dist/uninstall.js` — 卸载清理脚本
- `dist/mcp-excel.exe` — 已嵌入的 MCP Server
- `resources/` — 图标和 Webview 资源
- `skills/` — AI 指引文档
- `package.json` / `README.md` / `LICENSE`

源码（`src/`、TypeScript 配置等）已被 `.vscodeignore` 排除。

### 没有安装 Node.js 能用吗？

**可以。**

MCP Server 已打包为 `dist/mcp-excel.exe`，用户无需安装 Node.js 即可使用全部功能。

### 安装插件后 MCP 就可用吗？

插件激活时会自动向当前工作区的 `.trae/mcp.json` 写入基础 MCP 配置。选择 Excel 文件并设置同步目录后，AI 即可调用工具。

---

## 15. 如何运行宏

1. 命令面板执行 `Excel VBA: 运行宏`
2. 在弹出的 QuickPick 中选择要运行的宏（列出所有 `module.procedure` 形式的宏名）
3. 宏在已打开的 Excel 中执行，结果输出到 OutputChannel

**宏名建议**：优先使用完整限定名，例如 `VBAProject.Module1.Main` 或 `Sheet1.Worksheet_Activate`，避免重名冲突。

### 自动执行 VBA（可选）

开启后，AI 修改或新增代码后会自动提示执行指定宏：

```json
{
  "excelVba.autoRunVba": true,
  "excelVba.autoRunMacroName": "Module1.Main"
}
```

执行前仍会弹出确认对话框，避免误运行。

---

## 16. Excel 置顶

在控制面板开启「Excel 置顶」后，Excel 主窗口会始终保持在最前面，方便边写代码边查看运行结果。

断开 Excel 或服务重置时会自动取消置顶。

---

## 17. 如何使用 MCP Server

MCP Server 让 AI 通过 Model Context Protocol 直接调用 Excel/VBA 工具。

### 自动配置

本插件已内置 `mcp-server-excel`。**选择 Excel 文件并完成同步后，插件会自动向当前工作区的 `.trae/mcp.json` 写入 MCP 配置**，无需手动编辑。

示例配置：

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "C:/Users/<user>/.vscode/extensions/excel-vba-assistant-0.8.6/dist/mcp-excel.exe",
      "args": []
    }
  }
}
```

### 使用流程

1. 调用 `file(open, filePath)` 打开工作簿，获得 `sessionId`
2. 将 `sessionId` 传给后续所有工具调用
3. 按需选择以下工具类别：
   - **VBA**：`vba(list/view/import/update/delete/run)`
   - **工作表**：`worksheet(list/create/delete/rename/copy/move)`
   - **单元格**：`range(get-values/set-values/get-formulas/set-formulas/clear-contents)`
   - **格式**：`range_format(set-format/format-ranges/merge-cells/add-validation)`
   - **Excel 表格**：`table(create/delete/get-data/append-rows/add-to-datamodel)`
   - **数据透视表**：`pivottable(create/add-field/set-field-function/refresh)`
   - **图表**：`chart(create/set-title/set-data-source/add-trendline)`
   - **Power Query**：`powerquery(create/update/evaluate/refresh/load-to)`
   - **DAX 数据模型**：`datamodel(create-measure/update-measure/evaluate/create-relationship)`
   - **命名区域**：`namedrange(read/write/create/update/delete)`
   - **数据连接**：`connection(create/test/refresh/set-properties)`
   - **切片器**：`slicer(create/set-selection/delete)`
   - **条件格式**：`conditionalformat(add-rule/clear-rules)`
   - **截图**：`screenshot(capture/capture-sheet)`
   - **窗口**：`window(show/hide/set-topmost/set-position/arrange)`
   - **计算**：`calculation_mode(set-mode/calculate)`
4. 完成后调用 `file(close, sessionId, save=true)` 保存并释放会话

### AI 指引（Skills）

插件内置 `skills/` 目录，包含各功能域的操作指引。Trae AI 在调用 MCP 工具前会先读取 `skills/SKILL.md` 和相关文档，遵循以下约定：

- 先读取指引，再调用工具
- 危险操作需要确认
- 工作表/工作簿对象模块只更新，不新建/删除
- 批量读取优先使用 `vba(list)` 遍历
- 大数据量写入前使用 `calculation_mode(set-mode, manual)` 提升性能

### 推荐工作流

1. AI 调用 `vba(list)` 了解 VBA 项目结构
2. AI 调用 `vba(view)` 读取要修改的组件
3. AI 调用 `vba(update)` 写入新代码
4. AI 调用 `vba(run)` 执行宏
5. AI 调用 `range(get-values)` 或 `screenshot` 验证结果

---

## 18. 常见错误处理

| 错误现象 | 原因 | 解决方法 |
|---------|------|---------|
| 未检测到正在运行的 Excel | Excel 未启动或目标工作簿未打开 | 打开 Excel 并加载目标文件后重试 |
| 无法访问 VBAProject | 未开启「信任对 VBA 项目对象模型的访问」 | 见 [4. Excel 安全设置](#4-excel-安全设置) |
| 与 Excel 的连接已断开 | Excel 被关闭或崩溃 | 重新打开 Excel 和目标文件 |
| 操作超时 | Excel 正忙 | 检查 Excel 窗口是否有弹窗需要处理 |
| 文件正被其他程序占用 | 文件被其他进程锁定 | 关闭占用该文件的程序 |
| 宏执行失败 | 宏名错误或工作簿中不存在该宏 | 用 `vba(list/view)` 确认宏名 |
| 未找到目标工作表 | 工作表名称错误 | 用 `worksheet(list)` 确认名称 |
| 选择文件后提示「未检测到 Excel 中已打开该文件」| 工作簿未在 Excel 中打开 | 先在 Excel 中手动打开目标工作簿 |
| 关闭 Excel 后状态未自动清理 | 插件不再自动检测关闭 | 点击「断开 Excel」手动重置 |

---

## 19. 安全说明

### 危险操作

以下操作会修改 Excel 工作簿或执行代码，请谨慎使用：

1. **修改 VBA 代码**（`vba(update/import)`、本地 → VBE 同步）
2. **删除组件**（`vba(delete)`）
3. **运行宏**（`vba(run)`、自动执行 VBA）
4. **本地 → VBE 覆盖同步**
5. **工作表操作**（`worksheet(create/delete/rename)`）
6. **单元格写入/格式**（`range(set-values/clear-contents)`、`range_format(set-format)`）
7. **数据结构操作**（`table(create/delete)`、`pivottable(create/delete)`、`chart(create/delete)`）
8. **Power Query / DAX 修改**（`powerquery(create/update/delete)`、`datamodel(create-measure/delete-measure)`）

### 保护机制

- **手动同步确认**：通过命令面板执行「本地 → VBE 同步」时会弹确认对话框。
- **自动执行确认**：开启自动执行 VBA 后，每次执行前仍会弹确认对话框。
- **文档对象模块保护**：`Microsoft Excel 对象` 下的 `.wks` / `.wbk` 对象模块不允许通过同步删除，避免误删工作表。
- **窗体保护**：窗体不会从本地新建，必须先在 VBE 中创建，避免布局丢失。
- **删除同步保护**：工作表/工作簿对象模块永远不会被同步删除。

### 数据备份

建议定期执行 VBE → 本地同步，把 VBA 代码备份到本地目录，便于版本管理和回滚。

---

## 20. 开发与调试

### 在 Trae / VS Code 中调试

1. 用 Trae / VS Code 打开 `excel-vba-assistant` 目录
2. 按 `F5` 选择「扩展开发宿主」
3. 会启动一个新的 Extension Development Host 窗口
4. 左侧活动栏会出现 Excel VBA 图标，点击即可看到控制面板

### 打包

```powershell
npm run package
# 生成 excel-vba-assistant-0.8.6.vsix
```

### 查看日志

命令面板 → `Excel VBA: 打开输出日志`，会打开 `Excel VBA Assistant` 输出通道，显示同步、宏执行、错误等详情。

### Skill 同步

插件内置的 Skill 文档默认不会自动进入 Trae 的 Skill 系统。插件激活时，会自动把内置 Skill 同步到当前工作区的 `.trae/skills/excel-vba-assistant/` 目录；你也可以手动执行命令面板 → `Excel VBA: 同步 Skill 到工作区`。

---

## 许可证

MIT
