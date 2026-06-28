# Excel VBA Assistant

Excel VBA 本地同步器 + AI 编程桥接器：在 Trae / VS Code 中管理、同步、编辑、运行和 AI 辅助开发 Excel VBA 项目。

通过 PowerShell 调用 Excel COM Automation 操作已打开的工作簿，实现 VBAProject 的读写、双向同步、宏执行与单元格读取，并可选启用 MCP Server 让 AI 直接调用 Excel/VBA 工具。

---

## 1. 项目介绍

本插件提供以下核心能力：

- **VBE ↔ 本地双向同步**：把 Excel VBE 中的模块/类模块/窗体/工作表对象代码导出为本地 `.bas` / `.cls` / `.frm` 文件；反之把本地代码覆盖写回 VBE。
- **本地文件监听**：开启自动同步后，本地文件变动会自动写回 VBE（防抖 1000ms）。
- **宏执行**：列出工作簿中的所有 Sub/Function，并运行无参数宏。
- **单元格读取**：读取工作表的 UsedRange 或指定 Range，供 AI 理解数据。
- **状态栏**：在 VS Code 状态栏实时显示连接/同步/错误状态。
- **输出日志**：所有操作记录到 OutputChannel，便于排查问题。
- **MCP Server（可选）**：暴露 14 个 `excel_*` 工具，让 Trae AI 通过 MCP 协议直接读取 Excel、修改 VBA、运行宏。

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

> 文件夹名与 VBE 资源管理器原生分类保持一致。插件自带「Excel VBA Icons」文件图标主题，启用后 `.bas` / `.cls` / `.frm` / `.wks` / `.wbk` 和上述文件夹会显示不同图标。开启方式：`Ctrl+Shift+P` → `Preferences: File Icon Theme` → 选择 `Excel VBA Icons`。

---

## 2. 环境要求

- **操作系统**：Windows 10 / 11（依赖 PowerShell + Excel COM）
- **Excel**：已安装的桌面版 Excel（Office 2016 及以上推荐），并已打开目标工作簿
- **PowerShell**：系统自带 Windows PowerShell 5.1 或 PowerShell 7
- **Node.js**：18 及以上（用于构建和运行 MCP Server）
- **Trae / VS Code**：1.85 及以上

---

## 3. Excel 安全设置

VBA 项目对象模型默认不允许程序访问，必须先开启：

1. 打开 Excel → **文件** → **选项**
2. **信任中心** → **信任中心设置**
3. **宏设置** → 勾选 **「信任对 VBA 项目对象模型的访问」**
4. 点击确定保存

未开启时，所有读写 VBA 的操作都会返回错误：
> 无法访问 VBAProject。请在 Excel 中依次打开：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」。

---

## 4. 如何启动本地服务

本插件不依赖独立 HTTP 服务，所有 Excel 操作通过 PowerShell 直接调用 COM，开箱即用。

**首次构建**：

```powershell
cd d:\BaiduNetdiskDownload\jr-vbe-helper-1.0.8\assistance\excel-vba-assistant
npm install
npm run build
```

构建产物：

- `dist/extension.js` — VS Code 扩展主入口
- `dist/mcp-server.js` — MCP Server（可选）

---

## 5. 如何调试插件

### 方式 A：在 Trae / VS Code 中调试运行

1. 用 Trae / VS Code 打开 `excel-vba-assistant` 目录
2. 按 `F5` 选择「扩展开发宿主」（或运行命令 `Extensions: Debug Extension`）
3. 会启动一个新的 Extension Development Host 窗口
4. 左侧活动栏会出现 Excel VBA 图标，点击即可看到控制面板

### 方式 B：打包成 .vsix 安装

```powershell
npm run package
# 生成 excel-vba-assistant-0.1.0.vsix
```

在 Trae / VS Code 中执行 `Extensions: Install from VSIX...` 选择该文件安装。

### 查看 OutputChannel 日志

命令面板 → `Excel VBA: 打开输出日志`，会打开 `Excel VBA Assistant` 输出通道，显示同步、宏执行、错误等详情。

---

## 6. 如何选择 xlsm 文件

支持格式：`.xlsm` / `.xlsb` / `.xlam` / `.xls`

两种方式：

1. **控制面板按钮**：点击活动栏 Excel VBA 图标 → 控制面板中点击「选择 Excel 文件」
2. **命令面板**：`Ctrl+Shift+P` → `Excel VBA: 选择 Excel 文件`

选择后会弹出文件选择对话框。**请确保选中的文件已在 Excel 中打开**，否则后续操作会提示「未检测到正在运行的 Excel」。

---

## 7. 如何选择同步目录

两种方式：

1. **控制面板按钮**：「选择本地同步目录」
2. **命令面板**：
   - `Excel VBA: 选择本地同步目录` — 弹出目录选择对话框
   - `Excel VBA: 使用当前工作区作为同步目录` — 直接使用当前打开的工作区

建议选择一个空目录或专用的 VBA 代码目录，避免与其他项目混淆。

---

## 8. 如何执行 VBE → 本地

把 Excel VBE 中的所有 VBA 组件导出到本地同步目录：

1. 确保目标工作簿已在 Excel 中打开
2. 控制面板点击「VBE → 本地 同步」按钮，或命令面板执行 `Excel VBA: VBE → 本地 同步`
3. 等待状态栏显示同步成功
4. 在同步目录中查看生成的 `模块/`、`类模块/`、`窗体/`、`Microsoft Excel 对象/` 和 `workbook.json`

**编码处理**：VBE Export 产生 GBK 编码，本插件会转换为 UTF-8 存储，方便 AI 编辑和版本管理。

---

## 9. 如何执行 本地 → VBE

把本地代码覆盖写回 VBE：

1. 确保目标工作簿已在 Excel 中打开
2. 控制面板点击「本地 → VBE 同步」按钮，或命令面板执行 `Excel VBA: 本地 → VBE 同步`
3. **会弹出确认对话框**（此操作会覆盖 VBE 中的代码，需用户确认）
4. 确认后执行同步

**同步规则**：

| 目录 | 行为 |
|------|------|
| `模块/*.bas` | VBE 中存在则覆盖更新；不存在则新建标准模块 |
| `类模块/*.cls` | VBE 中存在则覆盖更新；不存在则新建类模块 |
| `Microsoft Excel 对象/*.wks` / `*.wbk` | 只更新已存在的 Excel 对象模块事件代码，不会新建工作表/工作簿对象 |
| `窗体/*.frm` | 只更新已存在窗体的事件代码部分，不会新建窗体；窗体布局需在 VBE 中操作 |

**代码清洗**：写入前会剥离 `VERSION` / `Attribute` / `Begin` / `End` 行，仅注入干净的过程代码。

---

## 10. 如何开启自动同步

开启后，本地 `*.bas` / `*.cls` / `*.frm` 文件**保存到磁盘后**会自动写回 VBE，无需手动触发。默认防抖 300ms（从保存到触发同步）。

两种方式：

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

### 缩短同步延迟

如果觉得同步慢，主要是 PowerShell 启动 + Excel COM 调用耗时。本插件已将防抖从 1000ms 降到 300ms。若仍想更快，可：

- 开启 `files.autoSave: afterDelay` 并减小 `autoSaveDelay`
- 关闭其他占用 Excel 的宏或弹窗
- 避免一次保存大量文件（批量改动建议先手动同步一次）

**注意**：自动同步使用静默模式，不会弹确认对话框；如果担心误改，建议保持关闭，改用手动同步。

---

## 11. 同步目录自动管理

### 选择 Excel 文件时自动创建同步目录

选择 Excel 文件后，如果尚未设置同步目录，插件会自动在 Excel 文件同路径下创建一个与文件名同名的文件夹作为同步目录：

- 例如选择 `D:\work\book.xlsm` → 自动创建 `D:\work\book\`
- 同步目录会被自动添加到当前工作区，并在资源管理器中打开

### 修改同步目录时自动清理

如果同步目录是插件自动创建的，当用户**手动选择新的同步目录**时，旧目录会被自动删除，避免多地冗余。

### 关闭 Excel 时自动清理

插件会每 5 秒检测一次 Excel 运行状态。当检测到 Excel 已关闭时：

- 如果同步目录是自动创建的，会自动删除该目录
- 清空插件中保存的 Excel 文件和同步目录状态
- 停止文件监听

这样设计的目的是：同步目录是从 Excel 解析出来的「衍生文件」，Excel 关闭后本地不再需要保留。

> 如果你希望长期保留解析文件，请手动选择同步目录（而非使用自动创建的目录），或关闭 Excel 前手动复制目录。

---

## 12. 如何运行宏

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

## 13. 如何使用 MCP Server

MCP Server 让 AI 通过 Model Context Protocol 直接调用 Excel/VBA 工具，无需用户手动触发命令。

### 配置 MCP Server

在 Trae 的 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "excel-vba-assistant": {
      "command": "node",
      "args": [
        "D:\\BaiduNetdiskDownload\\jr-vbe-helper-1.0.8\\assistance\\excel-vba-assistant\\dist\\mcp-server.js"
      ],
      "env": {
        "VBE_FILE_PATH": "D:\\work\\my_addin.xlam",
        "VBE_LOCAL_DIR": "D:\\work\\vba_sync",
        "VBE_READ_ONLY": "false"
      }
    }
  }
}
```

### 环境变量

| 变量 | 必填 | 说明 |
|------|------|------|
| `VBE_FILE_PATH` | 否 | 默认 Excel 文件路径，启动时自动注册为默认工作簿；不填则需通过 `excel_open_workbook` 工具注册 |
| `VBE_LOCAL_DIR` | 否 | 默认本地同步目录；不填则同步工具需在参数中显式指定 `localDir` |
| `VBE_READ_ONLY` | 否 | `true` 时启用只读模式，禁止所有危险操作（修改/创建/删除/运行宏/本地→VBE 同步） |

### 暴露的 21 个工具

| 工具名 | 类型 | 说明 |
|--------|------|------|
| `excel_list_workbooks` | 只读 | 列出已注册的工作簿 |
| `excel_open_workbook` | 注册 | 注册新的工作簿 |
| `excel_list_resources` | 只读 | 列出 VBAProject 所有组件 |
| `excel_get_vba_code` | 只读 | 读取指定组件代码 |
| `excel_update_vba_code` | **危险** | 覆盖性写入组件代码 |
| `excel_create_vba_component` | **危险** | 新建标准模块/类模块/窗体 |
| `excel_delete_vba_component` | **危险** | 删除组件（文档对象模块不可删） |
| `excel_list_macros` | 只读 | 列出所有 Sub/Function |
| `excel_run_macro` | **危险** | 运行无参数宏 |
| `excel_list_sheets` | 只读 | 列出工作表 |
| `excel_read_range` | 只读 | 读取指定 Range 数据 |
| `excel_read_used_range` | 只读 | 读取 UsedRange 数据 |
| `excel_sync_vbe_to_local` | 只读方向 | VBE 导出到本地 |
| `excel_sync_local_to_vbe` | **危险** | 本地覆盖写回 VBE |
| `excel_create_sheet` | **危险** | 新建工作表 |
| `excel_delete_sheet` | **危险** | 删除工作表 |
| `excel_rename_sheet` | **危险** | 重命名工作表 |
| `excel_set_cell_value` | **危险** | 设置单个单元格值 |
| `excel_set_range_values` | **危险** | 批量写入区域 |
| `excel_clear_range` | **危险** | 清空区域内容 |
| `excel_set_cell_format` | **危险** | 设置单元格格式 |

### 推荐工作流

1. AI 调用 `excel_list_resources` 了解 VBA 项目结构
2. AI 调用 `excel_get_vba_code` 读取要修改的组件
3. AI 调用 `excel_update_vba_code` 写入新代码
4. AI 调用 `excel_list_macros` 确认宏存在
5. AI 调用 `excel_run_macro` 执行宏
6. AI 调用 `excel_read_range` 验证结果

---

## 14. 常见错误处理

| 错误现象 | 原因 | 解决方法 |
|---------|------|---------|
| 未检测到正在运行的 Excel | Excel 未启动或目标工作簿未打开 | 打开 Excel 并加载目标文件后重试 |
| 无法访问 VBAProject | 未开启「信任对 VBA 项目对象模型的访问」 | 见 [3. Excel 安全设置](#3-excel-安全设置) |
| 与 Excel 的连接已断开 | Excel 被关闭或崩溃 | 重新打开 Excel 和目标文件 |
| 操作超时 | Excel 正忙或弹出了对话框 | 检查 Excel 窗口是否有弹窗需要处理 |
| 文件正被其他程序占用 | 文件被其他进程锁定 | 关闭占用该文件的程序 |
| 宏执行失败 | 宏名错误或工作簿中不存在该宏 | 用 `excel_list_macros` 确认宏名 |
| 未找到目标工作表 | 工作表名称错误（区分大小写） | 用 `excel_list_sheets` 确认名称 |
| 窗体同步失败 (0x800a004b) | VBE 对同名窗体旧状态未释放 | 关闭窗体设计器后重试 |

---

## 15. 安全说明

### 危险操作

以下操作会修改 Excel 工作簿或执行代码，请谨慎使用：

1. **修改 VBA 代码**（`excel_update_vba_code`、本地 → VBE 同步）
2. **删除组件**（`excel_delete_vba_component`）
3. **运行宏**（`excel_run_macro`、自动执行 VBA）
4. **本地 → VBE 覆盖同步**（`excel_sync_local_to_vbe`）
5. **工作表操作**（`excel_create_sheet` / `excel_delete_sheet` / `excel_rename_sheet`）
6. **单元格写入/格式**（`excel_set_cell_value` / `excel_set_range_values` / `excel_clear_range` / `excel_set_cell_format`）

### 保护机制

- **MCP Server 只读模式**：设置环境变量 `VBE_READ_ONLY=true` 后，所有危险操作会被拒绝，仅允许读取。
- **手动同步确认**：通过命令面板执行「本地 → VBE 同步」时会弹确认对话框。
- **自动执行确认**：开启自动执行 VBA 后，每次执行前仍会弹确认对话框。
- **文档对象模块保护**：`Microsoft Excel 对象` 下的 `.wks` / `.wbk` 对象模块不允许通过 MCP 工具删除，避免误删工作表。
- **窗体保护**：窗体不会从本地新建，必须先在 VBE 中创建，避免布局丢失。

### 数据备份

建议定期执行 VBE → 本地同步，把 VBA 代码备份到本地目录，便于版本管理和回滚。

---

## 许可证

MIT
