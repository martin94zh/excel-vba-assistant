# excel-vba-cli（ZCode / Codex agent 插件）

把 Excel 自动化能力以 **Agent Skill** 形式提供给编码 agent：安装后，ZCode、Codex 等
支持 `~/.agents/skills/` 或插件机制的客户端即可用自然语言驱动真实 Excel。

## 内容

- `skills/excel-cli/`：excel-cli 技能（SKILL.md + 26 篇领域指南 + 34 个命令组参数表），
  与 VSIX 扩展内置版本同源，参数表由 ExcelMcp 2.2.0 CLI 的 `--help` 实测生成。

## 前置条件

- Windows + Microsoft Excel 2016+
- excelcli 运行时（二选一）：
  - `npm install -g @sbroenne/excelcli`（推荐，命令上 PATH）
  - 或使用 Excel VBA Assistant 扩展（VSIX）内置的 `dist/excelcli.exe` 完整路径

## 与 VSIX 扩展的关系

两者经**同一个单实例 daemon** 共享会话：Trae/VS Code 里插件打开的前台工作簿，
agent 用 `excelcli -q session list` 即可接管；反之亦然。VBE 双向同步、宏运行桥
仅在 VSIX 扩展中可用。

## 本地测试

本目录附带 `plugins/marketplace.json`（dev 市场）。在 ZCode 中：
Plugin Marketplace → Add → Add Plugin Marketplace → 选择本仓库的 `plugins/` 目录 →
安装 excel-vba-cli。

## 公开分发

仓库根目录的 `marketplace.json` 使本仓库可直接作为 ZCode 插件市场添加：
Plugin Marketplace → Add → Add Plugin Marketplace → 输入 `martin94zh/excel-vba-assistant`。
发布新版 = 更新 `.zcode-plugin/plugin.json` 的版本号并推送，用户在市场刷新后即可更新。
