# FEASIBILITY.md — 技术可行性侦察报告

> 日期：2026-09-13。针对用户实测确认的三个核心痛点做只读取证：①同步慢/偏差高 ②AI 找不到工具 ③可视化（连接已打开的 Excel）。全程未修改任何代码，本文件为唯一新增产物。所有结论附文件与行号证据。

---

## 一、同步机制诊断

### 1.1 当前完整流程

**VBE → 本地（自动轮询）**
1. 每 3 秒轮询（`AUTO_VBE_TO_LOCAL_INTERVAL_MS = 3000`，extension.ts:42）：`detectVbeChangesAndQueue`（extension.ts:1147-1173）调用 `getVbeCodeChecksum()`。
2. checksum 变化 → 事件入队（extension.ts:1167-1171）→ `processPendingChanges`（:1175-1203）按事件源顺序处理 → `syncVbeToLocalQuiet`（:1077-1132）。入口保护：若本地有文件 mtime 晚于上次同步，直接跳过（:1084-1092）。
3. `syncVbeToLocal`（vbaClient.ts:404-632）：一条大 PowerShell 脚本内遍历**全部**组件 → 每个组件 `comp.Export()` 到临时文件 → 字节按 **GBK 解码**（:428, :441, :481）→ 与本地 UTF-8 内容做**全文本比较**（:448, :489）→ 变了才写（UTF-8 无 BOM）→ 清理本地孤儿文件（:553-573）→ 写 workbook.json。

**本地 → VBE（文件 watcher）**
1. watcher 监听 `*.bas,cls,frm,wks,wbk`，300ms 防抖（extension.ts:39, :953-981）→ 队列 → `syncLocalToVbeQuiet`（:996-1047）。
2. `ensureExcelRunning` 预检 = **独立的一次 PowerShell 进程**（vbaClient.ts:642 → powershell.ts:202-230）。
3. `syncLocalToVbe`（vbaClient.ts:639-894）：逐文件读取 → `Get-CleanCode` 清洗（剥 `VERSION..END` 块、`Attribute` 行、`Begin/End`，:657-672）→ `Find-Component` 线性扫全部组件找同名（O(n²)，:679-684）→ `Update-ComponentCode` 比较后整模块 `DeleteLines + AddFromString` 重写（:686-695）→ **无条件 `$wb.Save()`**（:856）→ 脚本内写 workbook.json（:875-882），TS 侧再写一次（:891, :897-907）。

### 1.2 慢的具体环节（调用点清单）

| # | 慢点 | 证据 |
|---|---|---|
| 1 | **每次同步 2-3 次 PowerShell 进程冷启动**：预检一次 + 主脚本一次；自动同步完成后还要 `captureVbeChecksum` 再一次（extension.ts:1028）。每次都写临时 .ps1 再 `chcp 65001 & powershell -NoProfile -File`（powershell.ts:90-103） | powershell.ts:90-123 |
| 2 | **PID attach 模式下每次进程内 `Add-Type` 现场编译 C#**（EnumWindows/AccessibleObjectFromWindow）。因为是"每任务一个新 powershell.exe"，编译缓存完全失效，每个脚本平白多出编译开销 | powershell.ts:135-164 |
| 3 | **轮询即全量读**：checksum 每 3 秒把所有组件的**每一行**代码读出拼接再 SHA256（:381-392），组件多时 3 秒周期本身被吃满 | vbaClient.ts:380-392 |
| 4 | **VBE→本地是全量导出**：每个组件都 Export 到临时文件、读回、比较，"未变"也要付全量导出的钱（:439, :479-482） | vbaClient.ts:432-497 |
| 5 | **本地→VBE 整模块重写**：`DeleteLines(全部) + AddFromString`，每次触发 VBE 重编译；且**每次都 `$wb.Save()`**——大工作簿保存秒级、可能弹兼容性对话框 | vbaClient.ts:692-693, :856 |
| 6 | **串行 + 全局互斥**：`isSyncing` 互斥（extension.ts:756, :997, :1148），一个慢同步阻塞一切；VBE→本地完成后还强制 500ms 缓冲（:1122-1130） | extension.ts |
| 7 | workbook.json 双写（脚本内 :875-882 + TS 侧 :891/:897-907）；每次操作新建 VbaClient 无复用（extension.ts:741-749） | — |

### 1.3 getVbeCodeChecksum 的真实现状

- 实现：vbaClient.ts:369-397，按名排序遍历组件，取 `CodeModule` 全部文本 + 组件名，拼串算 SHA256。
- 调用者：仅 `captureVbeChecksum`（extension.ts:1135-1144）与 `detectVbeChangesAndQueue`（:1154）。
- **它不是"先校验再读"的优化——校验本身就是全量读**。真正的同步（syncVbeToLocal）不消费它，两边各自全量扫描一遍。
- 额外缺陷：checksum 只含 `CodeModule` 文本，**不含窗体 .frx 二进制**（:382-388）→ 窗体控件布局改动不会触发 VBE→本地同步（漏变化）。

### 1.4 偏差（伪变化/漏变化）的来源

| # | 问题 | 证据 | 后果 |
|---|---|---|---|
| 1 | **PowerShell `-eq` 是大小写不敏感比较**：VBE→本地内容比较（:448, :489）与本地→VBE 的 `Update-ComponentCode`（:691）都用 `-eq` | vbaClient.ts | 仅大小写不同的改动（注释、字符串字面量——VBE 只自动规范化标识符大小写）被判"未变"，**双向漏同步** |
| 2 | checksum 不含 .frx | :382-388 | 窗体布局改动不触发同步 |
| 3 | 编码链不对称：导出按 GBK 解码→写 UTF-8；写回按 UTF-8 读。GBK 缺失时回退 `Encoding.Default`（:428-429） | vbaClient.ts | 中文 Windows 闭环成立；**非中文系统或用户把文件存成 ANSI/GBK → 乱码写进 VBE** |
| 4 | `Attribute`/`VERSION` 行不对称：本地→VBE 剥掉这些行（:657-672, :690），但 VBE→本地导出的 .bas/.cls 是**带 Attribute 头的全量内容**直接落盘（:479-492） | vbaClient.ts | 用户对 Attribute/VERSION 行的修改被静默忽略；两侧文件表示不一致，格式差异会造成伪变化 |
| 5 | 全文本精确比较含行尾符：`$existingContent -eq $content`（:448/:489）；编辑器把 CRLF 存成 LF → 判"变化"→ 写回 CRLF → 触发 watcher 反向同步（有 500ms 抑制 + mtime 保护，但链路存在） | vbaClient.ts / extension.ts:1122-1130 | 伪变化与循环风险 |
| 6 | **删除同步竞态（最危险）**：本地→VBE 会删除"本地没有"的模块（:810-854）。用户在 VBE 新建模块 → checksum 事件入队；若队列里先前已有 local 事件，它先执行时会把**尚未来得及导出的新模块直接删掉**。窗口期 ≈ 3 秒轮询 + 队列积压 | vbaClient.ts / extension.ts:1175-1203 | 用户代码丢失 |
| 7 | 冲突处理 = "后处理的事件源整体覆盖"，无内容级合并、无权威方向；唯一保护是 VBE→本地的 mtime 检查 | extension.ts:1175-1203, :1084-1092 | 双边同时改动时静默丢一边 |

### 1.5 结论与可改进点

**是双向同步**（两个方向 + 双向删除同步），但**没有冲突合并，没有权威方向**——权威就是"队列里后处理的事件"。可改进点（不动架构）：
1. 常驻单 PowerShell 进程/Runspace，消灭每 3 秒冷启动与重复 Add-Type（预计消除大部分延迟）；
2. 比较全部改大小写敏感（`-ceq` 或 `[String]::Equals(…, Ordinal)`）——一行级修复；
3. checksum 改为分组件哈希表，并纳入 .frx 哈希，实现增量检测；
4. 删除同步默认关闭，或改为"超过 N 个删除需确认"；
5. Save 策略改为"VBE 确有变更才 Save"。

---

## 二、工具发现问题诊断

### 2.1 工具总数与组织方式

- `tools/schemas.json` 是从 `dist/mcp-excel.exe` 直接 `tools/list` 落盘的快照（dump-schemas.js:16-19）。**共 25 个工具**：`calculation_mode, chart, chart_config, conditionalformat, connection, datamodel, datamodel_relationship, file, namedrange, powerquery, pivottable, pivottable_calc, pivottable_field, range, range_edit, range_format, range_link, screenshot, slicer, table, table_column, vba, window, worksheet, worksheet_style`。
- 每个工具 = `name/title/description/inputSchema`，操作靠 `action` 枚举（如 vba 的 list/view/import/update/run/delete，schemas.json:10-21）。README:61 说"17 类工具 230+ 操作"、SKILL.md:17 说"16 类"——**都落后于快照**（25 个里有 8 个是后来拆分的新工具：range_edit/range_link/chart_config/pivottable_calc/pivottable_field/table_column/datamodel_relationship/worksheet_style）。
- 无独立 category 字段，分类靠命名前缀 + description 开头的路由句（如 range 的描述明确说"insert/delete/find/sort 用 range_edit，styling 用 range_format"）。

### 2.2 工具如何暴露给 AI

- **插件完全不代理 AI 的工具调用**。Trae 按 `.trae/mcp.json`（extension.ts:1297-1307 写入）拉起**自己的** mcp-excel.exe，一次 `tools/list` 拿到**全部 25 个工具**，每个带数百字的 description（如 file 工具的描述是一大段工作流墙文，schemas.json:1454）。**全量暴露，无检索、无过滤、无分层。** 插件自己的 MCP 子进程（serverManager.ts）只服务于 openWorkbook/closeWorkbook/PID 探测，与 AI 的连接互不相干。
- 每次同步 skills 到 `.trae/skills/excel-vba-assistant/`（extension.ts:1523-1578，按 mtime 增量覆盖），触发机制是 Trae 对 SKILL.md frontmatter description 关键词的静态匹配（SKILL.md:1-11 "Triggers: Excel, spreadsheet..."）；skill 与工具的"绑定"是纯文本约定（cell-read.md:5 `## Tool` 写 `range` + JSON 示例），无运行时校验。
- **文档已与 exe 脱节**：skill 体系按 16-17 类编写，exe 实际 25 个工具；`range` 名字还在（get-values 等仍在 range 下），但拆分出的 8 个新工具没有任何 skill 覆盖。

### 2.3 根源与缺失层

AI 面对的是 25 个工具 × 数百个 action 枚举 × 大描述（README 宣称 230+ 操作），而同步+验证的核心场景真正需要的只有约 6 个：`file / vba / range / screenshot / calculation_mode / window`。**缺的是"工具检索/分层路由"这一层**：没有"按任务找工具"的入口，没有子集会话，AI 必须在 25 份长描述里自行大海捞针——这就是"找不到工具"的机制性根源。

改进方向（按代价升序）：
1. **对齐文档**：把 skills 更新到 25 工具版本（消除教错工具名）；
2. **Proxy MCP server**：插件写 `.trae/mcp.json` 时指向一个小代理（只暴露精选子集 + 一个 `search_tools(query)` 元工具，按需转发给真 exe）——插件事实上控制着 Trae 的连接配置，这是它唯一能下手的点；
3. **上游开关**：查证 mcp-server-excel 是否支持工具子集配置（离线无法确认）。

---

## 三、可视化问题诊断

### 3.1 schemas.json 里的证据（本地能力快照）

- `file` 工具的 `open` 动作有 **`show` 参数**："Whether to make Excel window visible. **Default: false (hidden automation)**"（schemas.json:1488）。**插件已经用上**：`sessionStore.ts:31` `show: options?.show ?? false`，`openWorkbook` 显式传 `show: true`（excelClient.ts:59-62）。→ **"插件打开的 Excel 用户看得见"已经实现**，不是待办。
- `file` 描述（:1454）："SESSION REUSE: Call 'list' first to check for existing sessions. **If file is already open, reuse existing sessionId** instead of opening again."——字面上支持复用已打开文件；同段还警告"操作超时会触发激进清理，可能让 Excel 处于不一致状态"。
- `window` 工具（:2282）可 show/hide/bring-to-front/arrange——对 server 管的实例有完整窗口控制。

### 3.2 与工程文档的矛盾（关键分歧点）

`.trae/documents/mcp-managed-excel-refactor.md:191`（§4.4，07-01 实测记录）："MCP `file(open)` 会失败，因为 **'Excel files must not be open in another Excel instance'**"。而 schemas.json 快照说"already open → reuse session"。两种解读：'already open' 指 server 自己管理的会话（则用户双击打开的文件仍被拒绝），或指任何已开文件（则支持 attach）。**本地证据冲突，倾向前者**（文档作者的报错信息明确说"another Excel instance"），但快照对应的 exe 可能比 07-01 实测时新，**必须实测裁定**。

### 3.3 COM 侧的历史结论（T3，已验证成功）

按 PID attach 在**本项目的 COM 层已经做成并在线上路径使用**：`buildExcelAttachScript`（powershell.ts:130-195）= EnumWindows 找目标 PID 的 XLMAIN 窗口 → `AccessibleObjectFromWindow(0xFFFFFFF0)` 取 Application → 失败回退 `GetActiveObject`。工程文档 T3 日志（:497-499）确认所有 VbaClient 方法都走这条 attach。即：**对"用户已打开的 Excel"做 COM 操作（含同步、读单元格、建表）今天就能工作**。

### 3.4 mcp-server-excel 的版本与来源

- 不是 npm 依赖：构建/运行时从 `sbroenne/mcp-server-excel` 的 **releases/latest** 下载（download-mcp-server.js:11、downloader.ts:15），**版本未固定**；dist/mcp-excel.exe（156MB）是 6-29 下载的某个 latest。
- 本地唯一的能力文档就是 tools/schemas.json 快照；dist/README.md 是本项目旧自述，无上游能力说明。上游 README/文档因外网不通无法查阅。

### 3.5 结论

- **MCP(上游) 能否 attach 到用户已打开的前台 Excel：本地证据冲突，倾向不支持，需实测裁定**（一行命令的实验：双击打开 xlsm → 让 AI `file(open)` 看报错）。
- 若确认不支持，可视化的可行路径有两条，且都不需要新发明：
  1. **"插件开、插件管"（已实现大半）**：一切从插件选择文件开始，`show:true` 已保证可见；用户放弃双击打开文件的习惯。
  2. **复活 COM 桥（代码已在仓库）**：vbaClient.ts:977-1278 的 listSheets/readRange/createSheet/setCellValue/setRangeValues/clearRange/setCellFormat 等死方法 + 已验证的 PID attach，正是"对用户已开实例做受限操作"的完整实现——它们当初被闲置只是因为自建 MCP server（server/）被删了。

---

## 四、三个问题的可解性评级

| 问题 | 评级 | 理由 | 代价 |
|---|---|---|---|
| 同步慢/偏差高 | **高可解** | 根因全部是实现层：进程模型（每 3 秒冷启动+Add-Type）、比较语义（`-eq` 大小写不敏感）、全量扫描、删除竞态。每一项都有明确、局部、低风险的修法（见 1.5），不动整体架构 | 重写 vbaClient 的执行与比较层 + 队列仲裁，是"局部重写"的主战场；估计占重写工作量的一半，但收益最直接 |
| AI 找不到工具 | **中可解** | 插件控制 `.trae/mcp.json`（Trae 的连接配置）和 skills 文档，这是抓手；但 AI 的会话由 Trae 自己管理，proxy 方案要新写一个组件，且 Trae 如何消化 25 份大 schema 只能实测；上游若有子集开关则代价骤降 | 小：对齐 skills 文档；中：写 proxy server（几百行）；未知项取决于上游与 Trae 行为 |
| 可视化 | **中可解** | "插件打开的文件可见"已实现（show:true）；"attach 用户已开实例"在 MCP 层证据冲突需一行实测；即便不支持，COM 桥的完整实现在仓库里现成放着（含已被生产验证的 PID attach） | 低：实测 + 文档引导；中：若走 COM 桥，把死方法包一层 MCP/命令暴露即可 |

---

## 五、路线建议

**插件与 MCP 的分工（基于以上证据的自然边界）**：
- **mcp-excel.exe（MCP 层）**：只服务"插件启动的、可见的（show:true）"Excel 实例——高级操作（透视表/Power Query/DAX/图表…）+ AI 桥（AI 经 `.trae/mcp.json` 连自己的 exe 实例，靠 session 复用操作同一工作簿）。
- **COM / vbaClient（插件层）**：VBE 双向同步（修复后）、宏执行与弹窗处理、窗口置顶、以及**对用户已打开实例的兼容 attach**（COM 按 PID attach 已验证可行）。

**应保留**：同步核心（修复比较语义与进程模型后）、vbaMacroRunner（弹窗处理，最独特资产）、PID attach（processUtils + buildExcelAttachScript）、skills（对齐 25 工具后）、工程文档、workbook.json 同步目录约定。
**应砍掉**：excelClient 的 166 方法封装面（留 openWorkbook/closeWorkbook/callTool）、vbaSyncEngine（死代码）、"230+ 操作全部类型化封装"这个目标、无条件删除同步（改为默认关+确认）、releases/latest 不固定版本（改为 pin 版本号，避免上游 breaking change 打穿构建）。

**如果三选一先做**：先修同步（问题一）——它可解性最高、是产品核心承诺、且修复手段全部是本地代码层；工具发现与可视化都依赖"实测上游行为"这个外部输入，可以并行侦察但不应阻塞。

---

## 六、仍需人工确认的事项

1. **mcp-server-excel 新版 `file(open)` 对"用户已打开文件"的真实行为**——schemas.json:1454 说复用会话，工程文档:191 记录实测拒绝。一行实验可裁定：双击打开 xlsm → AI 执行 `file(open)` 看报错。这是可视化路线的分岔点。
2. **上游是否支持工具子集/过滤配置**（启动参数/env）——决定"AI 找不到工具"是写 proxy 还是改配置。需在线查 sbroenne/mcp-server-excel 文档。
3. **Trae 如何把 25 个工具的 schema 注入 AI 会话**（全量塞 context 还是有检索）——决定 proxy 方案的实际收益。
4. **schemas.json 快照对应的 exe 版本**——dump 时间未记录，dist 的 exe 是 6-29 的 latest；快照能力（show 参数、session 复用）是否等于当前上游 latest 未知。
5. **非中文 Windows 上 GBK 回退的真实表现**（.NET CP936 是否可用、`Encoding.Default` 是什么）——决定编码链是否是用户实际遇到的偏差来源。
6. **`$wb.Save()` 在受保护视图/IRM 文件上的行为**（schemas.json:1454 提到 IRM 文件会被强制只读）——可能造成同步成功但保存失败。

---

*报告完。全程只读，未修改任何代码。*
