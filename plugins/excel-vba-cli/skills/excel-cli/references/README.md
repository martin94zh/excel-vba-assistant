# Excel CLI References

本目录是 excel-cli Skill 的参考文档，供 AI 按需加载。内容分两层：

- **叙事指南**（本目录下的 .md）：Excel 领域的工作流、约束、易错点与各功能领域指南，
  取自官方文档站 [excelmcpserver.dev/reference](https://excelmcpserver.dev/reference/)
  （ExcelMcp 2.2.0，MIT License，抓取于 2026-10-05；各文件头部有来源注释）。
  上游 2.2.0 起不再随包发布这些文档，统一维护在官网。
- **命令参考**（[commands/](commands/) 目录）：30 个命令组的动作与参数表，
  由脚本直接从插件内置 `excelcli.exe`（2.2.0）的 `--help` 输出生成，
  与实际二进制严格一致。`service` / `batch` / `session` / `diag` 四个基础命令
  的文档经人工核对无变化，保留原版。

命令组索引见 [cli-commands.md](cli-commands.md)（含 Common Pitfalls）。

> **注意**：遇到任何命令报"Unknown action/option"，以 `excelcli <命令组> --help`
> 的实时输出为准——文档可能滞后于内置 CLI 的版本。
