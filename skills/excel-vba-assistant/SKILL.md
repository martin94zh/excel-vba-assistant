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

## 插件侧能力（AI 无需介入）

| 能力 | 实现方式 |
|------|----------|
| 选择文件并前台打开 | `excelcli session open --show`（daemon 持有） |
| VBE ↔ 本地双向同步 | 插件内部 COM（自动 3 秒轮询 + 文件监听 + 手动命令） |
| 宏运行（含弹窗处理） | 插件内部 COM `runMacro`（弹窗自动检测交互） |
| Excel 置顶 | 插件内部 COM（keepExcelOnTop 开关） |
| 断开并保存 | `excelcli session close --save`，COM 兜底 |

## 关键规则

- **同一时间只有一个持有者**：若文件已被用户手动在 Excel 中打开，`session open` 会因
  独占访问失败——请让用户先关闭它，或在插件提示中选择"关闭并让插件重新打开"。
- **保持前台可见**：不要调用 `window hide`；用户要求"看到处理过程"是本插件的默认预期。
- **批量写入**用 `range set-values` 一次一行的二维数组；10+ 条命令改用
  `excelcli batch --input commands.json`（见官方 SKILL Rule 8）。
- **VBA 文档模块**（Sheet1/ThisWorkbook）只能更新代码，不能增删模块本身。
