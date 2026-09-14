---
name: excel-vba-assistant
description: >
  Excel VBA Assistant plugin orchestrator: how the plugin, the excelcli daemon,
  and VBE sync work together. The user picks a workbook in the plugin; the plugin
  opens it VISIBLE via `excelcli session open --show` (daemon-owned session);
  the AI then operates the SAME session with excelcli commands (ranges, tables,
  pivots, charts, Power Query, DAX, VBA, screenshots — 31 command groups / 326
  operations), and the plugin keeps VBE code bidirectionally synced with local
  files. Use for ANY Excel task in this workspace.
  Triggers: Excel, spreadsheet, workbook, xlsm, VBA, macro, excelcli, session.
---

# Excel VBA Assistant — 插件编排说明

本插件 = **excelcli（31 组命令 / 326 操作）+ VBE 双向同步（COM）+ 前台可视化**。

## 核心模型（必读）

1. **会话由 daemon 持有**：用户在插件中"选择 Excel 文件"后，插件执行
   `excelcli session open <文件> --show`，Excel **前台可见**地打开，会话归共享 daemon 所有。
2. **AI 用同一会话操作**：用 `session list` 找到目标文件的 `sessionId`（通常已存在，
   不要重复 open），随后所有命令携带 `--session <id>`。任意新起的 excelcli 进程都能
   操作同一工作簿——这就是插件与 AI 协作的通道。
3. **excelcli 的完整路径**（未加入 PATH，必须使用完整路径）：
   **`{{EXCELCLI_PATH}}`**
4. **VBE 双向同步由插件负责**：AI 修改 VBA 会实时出现在前台 Excel 的 VBE 中；插件的
   自动同步会把 VBE 变更导出到本地文件（用户在编辑器中可见），也会把本地修改写回 VBE。
   **AI 不要用 `vba` 命令做批量同步**，只做单模块/单过程操作（详见
   [vba-sync.md](vba-sync.md)）。

## AI 工作流

```
1. & '{{EXCELCLI_PATH}}' -q session list                → 找到插件已打开会话的 sessionId
2. & '{{EXCELCLI_PATH}}' -q sheet list --session <id>   → 发现工作表
3. 按需调用 31 组命令（range/table/pivottable/chart/powerquery/datamodel/vba/screenshot…）
   → 全部携带 --session <id>；Excel 前台可见，用户实时看到你的每一步
4. 结束时若会话是你自己打开的：session close --session <id> --save
   （插件打开的会话通常由插件断开时保存，一般无需你关闭）
```

完整命令参考：[../excel-cli/SKILL.md](../excel-cli/SKILL.md)（31 组命令与 326 操作的官方文档）。

## 宏开发循环（改宏 → 跑宏 → 抓报错 → 读结果 → 再改）

> **通道路由（必读）**：如果环境中存在旧版 `excel-mcp` MCP 工具（`file`/`range` 等），**不要使用**——
> 它们自建独占会话，与插件的 daemon 会话互斥：插件已打开的文件用 MCP open 必然报
> "already open"，反过来 MCP 打开的文件插件也无法附着。遇到 "already open" 的正确解法：
> `excelcli -q session list` 找插件会话，走本节流程。所有 Excel 操作一律走 excelcli。

AI 修改 VBA 后测试宏，**必须走插件的宏运行桥**（能抓取报错/MsgBox/InputBox 弹窗全文、自动交互、Excel 不会被杀）。
说明：官方 excel-cli 文档的 `vba run` 不处理任何弹窗——含 MsgBox/InputBox 的真实宏用它运行必然卡死，
本插件的宏运行桥专为此设计，两者不重复：**桥=测试运行，CLI vba run=仅限已验证无弹窗的宏**。

1. `vba update`（excelcli）写入/修改模块代码
2. 用文件工具向**同步目录**写 `macro-run.json`（UTF-8）。PowerShell 写法（务必无 BOM）：
   ```powershell
   $req = @'
   {"macro": "Module1.YourProc", "timeoutMs": 90000, "dialogMode": "auto",
    "confirmButton": "是", "inputValue": "42", "saveAfterRun": true}
   '@
   [IO.File]::WriteAllText("<同步目录>\macro-run.json", $req, [Text.UTF8Encoding]::new($false))
   ```
   | 字段 | 说明 |
   |------|------|
   | `macro` | 过程名，格式 `模块名.过程名` |
   | `timeoutMs` | 默认 45000。**首次运行新宏建议 90000**（VBA 编译 + 首次执行较慢） |
   | `dialogMode` | `auto`（默认）自动处理全部弹窗；`errors` 仅自动结束报错弹窗 |
   | `confirmButton` | 宏内 `MsgBox vbYesNo` 确认框点哪个按钮（默认"取消"，可指定"是"） |
   | `inputValue` | 宏内 `InputBox` 自动填入的文本 |
   | `saveAfterRun` | `true` = 运行成功后自动保存工作簿（CLI 没有 save 动作，这是推荐的保存方式） |
3. 轮询读取同目录 `macro-run-result.json`（文件出现即为运行完成）：
   - `success=false` 时 `dialogs[]` 含报错弹窗的标题与**全文**（如"运行时错误 '11': 除数为零"、
     "编译错误: 子过程或函数未定义"），弹窗已被自动关闭，按报错内容修正代码
   - `success=true` 时 `dialogs[]` 记录宏触发的 MsgBox/InputBox 内容与自动操作
     （如 MsgBox 文本"处理完成：10 行"已被点确定）——**用它验证宏的交互行为是否符合预期**
   - **超时 ≠ 失败**：结果含 `macroStillRunning: true` 表示宏还在跑，等待后重查；
     即使超时，宏也可能已实际执行完毕——先 `range get-values` 检查副作用，**确认后再决定是否重跑**，
     避免非幂等宏重复执行
   - `saved: true` 表示工作簿已自动保存
4. `range get-values` 读取运行结果单元格，依据结果继续修改，重复 1-4

保存工作簿的三种方式：`saveAfterRun: true`（推荐，随桥运行自动保存）；桥接运行一个
`ThisWorkbook.Save` 宏；`session close --save`（会关闭会话，仅在彻底结束时用）。

注意事项：

- **不要对可能有错的宏直接用 `excelcli vba run`**：VBA 错误弹窗会阻塞 daemon，超时后会话被销毁、
  Excel 进程退出且拿不到报错文本。确认无风险的宏（已通过桥接跑通）才可直接 `vba run`。
- 若会话因历史操作失效：`session list` 检查 → `session open` 重新打开即可恢复。
- 所有错误弹窗的最近 20 条记录同时保存在同步目录 `excel-dialogs.json`，插件宿主也会弹出 VS Code 通知。

## 插件侧能力（AI 无需介入）

| 能力 | 实现方式 |
|------|----------|
| 选择文件并前台打开 | `excelcli session open --show`（daemon 持有） |
| VBE ↔ 本地双向同步 | 插件内部 COM（自动 3 秒轮询 + 文件监听 + 手动命令） |
| 宏运行（含弹窗处理） | 插件内部 COM `runMacro`（弹窗自动检测、抓取文本、点击"结束"，编译错误自动重置 VBE） |
| 宏运行桥（AI 专用） | 同步目录 `macro-run.json` → `macro-run-result.json` |
| Excel 置顶 | 插件内部 COM（keepExcelOnTop 开关） |
| 断开并保存 | `excelcli session close --save`，COM 兜底 |

## 关键规则

- **同一时间只有一个持有者**：若文件已被用户手动在 Excel 中打开，`session open` 会因
  独占访问失败——请让用户先关闭它，或在插件提示中选择"关闭并让插件重新打开"。
- **保持前台可见**：不要调用 `window hide`；用户要求"看到处理过程"是本插件的默认预期。
- **批量写入**用 `range set-values` 一次一行的二维数组；10+ 条命令改用
  `excelcli batch --input commands.json`（见官方 SKILL Rule 8）。
- **VBA 文档模块**（Sheet1/ThisWorkbook）只能更新代码，不能增删模块本身。
