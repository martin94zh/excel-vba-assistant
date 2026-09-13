/**
 * Excel VBA Assistant - 扩展入口
 *
 * 职责：
 * 1. 注册 Webview 面板、命令、状态栏、输出通道
 * 2. 处理 Webview 消息（选择文件/目录、同步、开关）
 * 3. 自动同步：监听本地 *.bas/*.cls/*.frm 变化，防抖后写回 VBE
 * 4. 自动执行 VBA：本地 → VBE 同步后，确认运行宏
 *
 * 工作簿归属模型（CLI daemon 单持有者）：
 * ExcelMcp 2.x 的 excelcli.exe 通过共享后台 daemon（命名管道服务）持有 Excel，
 * 所有 excelcli 进程连接同一 daemon、会话跨进程持久。因此：
 * - 用户在插件中选择文件 → 插件执行 `excelcli session open <file> --show`，
 *   Excel 前台可见，会话归 daemon 所有；
 * - AI（Trae/VS Code）用任意新起的 excelcli 进程携带 --session <id> 即可操作
 *   同一工作簿（31 组命令 / 326 操作，含 range/vba/powerquery/datamodel 等）；
 * - 插件通过 COM 附着同一 Excel 实例完成 VBE 双向同步、宏运行（含弹窗处理）
 *   与窗口置顶。
 * （不使用 MCP server：stdio 模式下每个 MCP 客户端各自持有进程内服务，
 *   会话互不可见且独占互斥，无法满足"插件打开 + AI 操作"同会话需求。）
 */
import * as vscode from "vscode";
import * as path from "path";
import * as os from "os";
import * as fs from "fs";
import { mkdir, readdir, rm } from "fs/promises";

import { StateManager } from "./state";
import { OutputManager } from "./output/outputChannel";
import { StatusBarManager } from "./status/statusBar";
import { ExcelVbaPanelProvider, type WebviewMessage, type StatePayload } from "./webview";
import { VbaClient } from "./client/vbaClient";
import { findLocalExcelCli, runExcelCli } from "./native/cliRunner";
import {
  findExcelPidForWorkbook,
  waitForProcessExit,
  terminateProcessTree,
  closeWorkbookAndQuit,
} from "./native/processUtils";

const SUPPORTED_EXCEL_EXTENSIONS = ["xlsm", "xlsb", "xlam", "xls"];
const VBA_FILE_EXTENSIONS = [".bas", ".cls", ".frm", ".wks", ".wbk"];

// 自动同步防抖毫秒数。太短容易连续触发，太长体感延迟高；300ms 在敲击停止后几乎无感知。
const AUTO_SYNC_DEBOUNCE_MS = 300;

// Excel → 本地 自动同步轮询间隔。Excel/VBE 没有直接文件系统事件，只能轮询检测变更。
const AUTO_VBE_TO_LOCAL_INTERVAL_MS = 3000;

// 等待工作簿被打开的轮询间隔与超时（AI 的 MCP server 打开，或用户手动打开）
const OPEN_WAITER_INTERVAL_MS = 3000;
const OPEN_WAITER_TIMEOUT_MS = 10 * 60 * 1000;

interface PendingChange {
  source: "local" | "vbe";
  timestamp: number;
}

let extensionContext: vscode.ExtensionContext;
let stateManager: StateManager;
let output: OutputManager;
let statusBar: StatusBarManager;
let viewProvider: ExcelVbaPanelProvider;
let fileWatcher: vscode.FileSystemWatcher | undefined;
let syncDebounceTimer: NodeJS.Timeout | undefined;
let vbeToLocalCheckTimer: NodeJS.Timeout | undefined;
let openWaiterTimer: NodeJS.Timeout | undefined;
let openWaiterDeadline = 0;
let isSyncing = false;
let isSyncingVbeToLocal = false;
let lastExcelAvailable = false;

// 自动同步变更队列：记录本地和 VBE 的修改事件，按时间顺序处理
let pendingChanges: PendingChange[] = [];
// 上次 VBE 代码 checksum，用于检测 VBE 是否有新变更
let lastVbeChecksum = "";

export function activate(context: vscode.ExtensionContext): void {
  extensionContext = context;

  // 平台检查：仅支持 Windows
  if (process.platform !== "win32") {
    vscode.window.showWarningMessage("当前插件第一版仅支持 Windows + Microsoft Excel 桌面版。");
  }

  stateManager = new StateManager(context);
  output = new OutputManager();
  statusBar = new StatusBarManager(stateManager);
  context.subscriptions.push(output, statusBar);

  viewProvider = new ExcelVbaPanelProvider(context.extensionUri, stateManager, (msg) => {
    void handleWebviewMessage(msg);
  });
  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider(ExcelVbaPanelProvider.viewType, viewProvider)
  );

  registerCommands(context);
  stateManager.onChange(() => pushStateToWebview());

  output.info("Excel VBA Assistant 已激活");

  // 自动设置文件图标主题为 Excel VBA Icons（若用户已选择其他主题则不覆盖）
  void setExcelVbaIconTheme();

  // 立即推送一次已持久化的状态，避免 webview 启动时显示空白/丢失
  pushStateToWebview();

  // 工作簿归属模型：插件通过共享的 excelcli daemon 打开工作簿（前台可见），
  // AI 用任意 excelcli 进程操作同一会话；插件本身经 COM 伴随同步。
  // 激活时只做两件事：恢复连接状态、按需启动自动同步。
  void restoreExcelConnectionStatus();

  if (stateManager.get("autoSync") && stateManager.get("syncDirectory")) {
    startFileWatcher();
    startVbeToLocalWatcher();
  }

  // 将内置 Skill 同步到当前工作区的 .trae/skills，使 Trae AI 能按需加载
  void syncBuiltinSkillsToWorkspace();
}

async function restoreExcelConnectionStatus(): Promise<boolean> {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) {
    stateManager.setRuntime({ serviceStatus: "disconnected" });
    pushStateToWebview();
    return false;
  }

  // 服务恢复前，确保 Excel 置顶处于关闭状态，避免恢复后窗口异常置顶
  if (stateManager.get("keepExcelOnTop")) {
    await restoreExcelWindowState().catch(() => {});
    await stateManager.set("keepExcelOnTop", false);
  }

  // 优先校验 excelcli daemon 会话是否仍然存活（daemon 重启后会话会失效）
  const cliPath = getExcelCliPath();
  const sid = stateManager.get("cliSessionId");
  if (cliPath && sid) {
    const listed = await runExcelCli(cliPath, ["-q", "session", "list"], 30000);
    const sessions = (listed.data?.sessions as Array<{ sessionId?: string }> | undefined) ?? [];
    if (listed.success && sessions.some((s) => s.sessionId === sid)) {
      output.info(`excelcli 会话仍存活（${sid.slice(0, 8)}…），恢复连接`);
      return await attachAndConnect(workbookPath);
    }
    output.warn("excelcli 会话已失效，清除后尝试按 PID 恢复");
    await stateManager.set("cliSessionId", "");
  }

  // 定位当前持有工作簿的 Excel 实例（daemon 打开或用户手动打开，均可附着）
  const pid = await findExcelPidForWorkbook(workbookPath);
  if (!pid) {
    // 工作簿尚未打开：进入等待状态，由 open-waiter 检测其被打开
    output.info("Excel 尚未打开所选工作簿，等待其被打开（让 AI 执行 session open 或手动打开）…");
    stateManager.setRuntime({ serviceStatus: "starting" });
    startOpenWaiter(workbookPath);
    pushStateToWebview();
    return false;
  }

  return await attachAndConnect(workbookPath, pid);
}

/** 校验 VBA 访问能力并置为已连接（附着指定实例或自动定位） */
async function attachAndConnect(workbookPath: string, knownPid?: number): Promise<boolean> {
  const pid = knownPid ?? (await findExcelPidForWorkbook(workbookPath));
  if (pid) {
    await stateManager.set("excelProcessId", pid);
  }

  // 通过 COM 验证该实例的 VBA 访问能力（信任设置等）
  const accessError = await new VbaClient(workbookPath, pid).checkAccess();
  if (accessError) {
    output.warn(`恢复连接状态时 Excel 未就绪：${accessError}`);
    stateManager.setRuntime({ lastError: accessError, serviceStatus: "error" });
    // 服务未连接时，强制关闭自动同步和自动执行 VBA
    await stateManager.set("autoSync", false);
    await stateManager.set("autoRunVba", false);
    stopFileWatcher();
    pushStateToWebview();
    return false;
  }

  stateManager.setRuntime({ serviceStatus: "connected", lastError: undefined });
  lastExcelAvailable = true;
  output.info(`恢复连接状态：Excel 已连接${pid ? `（PID ${pid}）` : ""}`);
  pushStateToWebview();
  return true;
}

export async function deactivate(): Promise<void> {
  stopFileWatcher();
  stopVbeToLocalWatcher();
  stopOpenWaiter();
  // 扩展被禁用/重载时清理旧版 MCP 配置残留
  await removeLegacyMcpConfig();
}

// ============================================================
// 命令注册
// ============================================================

function registerCommands(context: vscode.ExtensionContext): void {
  const register = (id: string, handler: () => void) =>
    context.subscriptions.push(vscode.commands.registerCommand(id, handler));

  register("excelVba.openSettingsPanel", () => {
    void vscode.commands.executeCommand("excelVbaControlPanel.focus");
  });

  register("excelVba.selectWorkbook", () => void handleSelectWorkbook());
  register("excelVba.selectSyncDirectory", () => void handleSelectSyncDirectory());
  register("excelVba.useCurrentWorkspace", () => void handleUseCurrentWorkspace());
  register("excelVba.useExcelSameDirectory", () => void handleUseExcelSameDirectory());
  register("excelVba.disconnectExcel", () => void handleDisconnectExcel());
  register("excelVba.syncVbeToLocal", () => void executeSync("vbe-to-local"));
  register("excelVba.syncLocalToVbe", () => void executeSync("local-to-vbe"));

  register("excelVba.toggleAutoSync", async () => {
    const next = !stateManager.get("autoSync");
    await stateManager.set("autoSync", next);
    if (next) {
      startFileWatcher();
      startVbeToLocalWatcher();
    } else {
      stopFileWatcher();
      stopVbeToLocalWatcher();
    }
    output.info(`自动同步已${next ? "开启" : "关闭"}`);
  });

  register("excelVba.toggleAutoRunVba", async () => {
    const next = !stateManager.get("autoRunVba");
    await stateManager.set("autoRunVba", next);
    output.info(`自动执行 VBA 已${next ? "开启" : "关闭"}`);
  });

  register("excelVba.runMacro", () => void handleRunMacro());
  register("excelVba.openOutput", () => output.show(false));
  register("excelVba.refreshResources", () => void handleRefreshResources());
  register("excelVba.syncSkills", () => void syncBuiltinSkillsToWorkspace(true));
  register("excelVba.copyCliPath", () => void copyCliPath());
}

// ============================================================
// Webview 消息处理
// ============================================================

async function handleWebviewMessage(msg: WebviewMessage): Promise<void> {
  switch (msg.type) {
    case "selectWorkbook":
      await handleSelectWorkbook();
      break;
    case "selectSyncDirectory":
      await handleSelectSyncDirectory();
      break;
    case "useCurrentWorkspace":
      await handleUseCurrentWorkspace();
      break;
    case "useExcelSameDirectory":
      await handleUseExcelSameDirectory();
      break;
    case "disconnectExcel":
      await handleDisconnectExcel();
      break;
    case "toggleAutoRun":
      await stateManager.set("autoRunVba", msg.value);
      output.info(`自动执行 VBA 已${msg.value ? "开启" : "关闭"}`);
      break;
    case "toggleAutoSync":
      await stateManager.set("autoSync", msg.value);
      if (msg.value) {
        startFileWatcher();
        startVbeToLocalWatcher();
      } else {
        stopFileWatcher();
        stopVbeToLocalWatcher();
      }
      output.info(`自动同步已${msg.value ? "开启" : "关闭"}`);
      break;
    case "toggleKeepExcelOnTop":
      await stateManager.set("keepExcelOnTop", msg.value);
      output.info(`Excel 置顶已${msg.value ? "开启" : "关闭"}`);
      if (!msg.value) {
        await restoreExcelWindowState();
      } else {
        await applyExcelOnTopIfNeeded();
      }
      break;
    case "syncVbeToLocal":
      await executeSync("vbe-to-local");
      break;
    case "syncLocalToVbe":
      await executeSync("local-to-vbe");
      break;
    case "copyCliPath":
      await copyCliPath();
      break;
  }
}

// ============================================================
// 选择文件 / 目录
// ============================================================

async function handleSelectWorkbook(): Promise<void> {
  const uris = await vscode.window.showOpenDialog({
    canSelectFiles: true,
    canSelectFolders: false,
    canSelectMany: false,
    filters: { "Excel Macro Files": SUPPORTED_EXCEL_EXTENSIONS },
    title: "Excel VBA Assistant - 选择 Excel 文件",
  });
  if (!uris || uris.length === 0) return;
  const filePath = uris[0].fsPath;
  const ext = path.extname(filePath).replace(".", "").toLowerCase();
  if (!SUPPORTED_EXCEL_EXTENSIONS.includes(ext)) {
    vscode.window.showWarningMessage("当前文件格式不支持保存 VBA。请使用 .xlsm、.xlsb 或 .xlam 文件。");
    return;
  }
  const previousWorkbookPath = stateManager.get("workbookPath");
  const previousSyncDir = stateManager.get("syncDirectory");
  const previousAutoCreated = stateManager.get("syncDirectoryAutoCreated");

  await stateManager.set("workbookPath", filePath);
  output.info(`已选择 Excel 文件：${filePath}`);

  // 如果切换到了不同的 Excel 文件，且旧同步目录是自动创建的，先清理旧目录
  const isSwitchingWorkbook = previousWorkbookPath && previousWorkbookPath !== filePath;
  if (isSwitchingWorkbook && previousAutoCreated && previousSyncDir) {
    await removeAutoCreatedSyncDir(previousSyncDir);
    await stateManager.set("syncDirectory", "");
    await stateManager.set("syncDirectoryAutoCreated", false);
  }

  // 选择文件为独立操作，不再自动创建同步目录、不再自动同步
  viewProvider.refresh();
  pushStateToWebview();

  // 检测目标文件是否已被其他 Excel 实例打开（独占模型下 AI 的 server 将无法 open）。
  // 用 CommandLine + 窗口标题双信号定位，覆盖双击打开与 Excel 内部打开两种情形。
  const conflictingOwnerPid = await findExcelPidForWorkbook(filePath);
  const currentManagedPid = stateManager.get("excelProcessId");
  if (conflictingOwnerPid && conflictingOwnerPid !== currentManagedPid) {
    const fileName = path.basename(filePath);
    const choice = await vscode.window.showWarningMessage(
      `检测到 ${fileName} 已在另一个 Excel 实例中打开。AI 操作工作簿需要独占访问：` +
      "关闭它并让 AI 重新打开后，AI 才能操作该文件；保持打开则仅支持本地 VBA 同步。",
      { modal: true },
      "关闭并让 AI 重新打开",
      "保持打开（仅本地同步）"
    );
    if (choice === "保持打开（仅本地同步）") {
      // 直接附着现有实例：本地同步可用，但该实例非 daemon 持有，
      // AI 无法用 excelcli 打开同一文件（独占限制）
      output.info(`用户选择保持现有实例（PID ${conflictingOwnerPid}），仅启用本地同步`);
      await stateManager.set("excelProcessId", conflictingOwnerPid);
      stateManager.setRuntime({ serviceStatus: "connected", lastError: undefined });
      lastExcelAvailable = true;
      if (stateManager.get("autoSync") && stateManager.get("syncDirectory")) {
        startFileWatcher();
        startVbeToLocalWatcher();
      }
      pushStateToWebview();
      return;
    }
    if (choice !== "关闭并让 AI 重新打开") {
      output.info("用户取消接管已打开的 Excel 文件");
      await stateManager.set("workbookPath", "");
      pushStateToWebview();
      return;
    }

    output.info(`用户选择关闭由 PID ${conflictingOwnerPid} 打开的 Excel 实例`);
    stateManager.setRuntime({ serviceStatus: "starting" });
    pushStateToWebview();

    const closed = await closeWorkbookAndQuit(conflictingOwnerPid, filePath);
    if (!closed) {
      output.warn(`COM 方式关闭 PID ${conflictingOwnerPid} 的 Excel 失败，尝试强制终止`);
      try {
        await terminateProcessTree(conflictingOwnerPid);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        output.error(`强制终止 Excel 进程 ${conflictingOwnerPid} 失败：${msg}`);
        stateManager.setRuntime({ lastError: `无法关闭已打开的 Excel：${msg}`, serviceStatus: "error" });
        vscode.window.showErrorMessage(`无法关闭已打开的 Excel：${msg}`);
        pushStateToWebview();
        return;
      }
    }

    output.info(`等待 Excel 进程 ${conflictingOwnerPid} 退出（最多 5 秒）`);
    const exited = await waitForProcessExit(conflictingOwnerPid, 5000);
    if (!exited) {
      output.warn(`Excel 进程 ${conflictingOwnerPid} 未在 5 秒内退出，执行强制终止兜底`);
      try {
        await terminateProcessTree(conflictingOwnerPid);
        await waitForProcessExit(conflictingOwnerPid, 2000);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        output.warn(`强制终止兜底失败：${msg}`);
      }
    }
  }

  // 通过共享 daemon 打开工作簿：--show 使 Excel 前台可见，
  // 会话由 daemon 持有，AI 之后用任意 excelcli 进程 --session <id> 操作同一工作簿
  const cliPath = getExcelCliPath();
  if (!cliPath) {
    const msg = "未找到 excelcli.exe（插件安装不完整），无法打开 Excel";
    output.error(msg);
    stateManager.setRuntime({ lastError: msg, serviceStatus: "error" });
    vscode.window.showErrorMessage(msg);
    pushStateToWebview();
    return;
  }

  stateManager.setRuntime({ serviceStatus: "starting" });
  pushStateToWebview();

  output.info(`正在通过 excelcli 打开工作簿（前台可见）：${filePath}`);
  const opened = await runExcelCli(cliPath, ["-q", "session", "open", filePath, "--show"], 120000);
  if (!opened.success) {
    const msg = opened.error || "excelcli 打开工作簿失败";
    output.error(`excelcli 打开工作簿失败：${msg}`);
    stateManager.setRuntime({ lastError: msg, serviceStatus: "error" });
    vscode.window.showErrorMessage(`无法打开 Excel 文件：${msg}`);
    pushStateToWebview();
    return;
  }

  const sessionId = String((opened.data?.sessionId as string) || "");
  await stateManager.set("cliSessionId", sessionId);
  output.info(`excelcli 会话已建立：${sessionId.slice(0, 8)}…（Excel 前台可见）`);

  // 定位 daemon 启动的 Excel 进程并附着（COM 同步用）
  const pid = await findExcelPidForWorkbook(filePath);
  if (pid) {
    await stateManager.set("excelProcessId", pid);
    output.info(`已记录 Excel 进程 PID：${pid}`);
  } else {
    // Excel 窗口尚未就绪：交给 open-waiter 检测
    startOpenWaiter(filePath);
  }

  stateManager.setRuntime({ serviceStatus: "connected", lastError: undefined });
  lastExcelAvailable = true;
  if (stateManager.get("autoSync") && stateManager.get("syncDirectory")) {
    startFileWatcher();
    startVbeToLocalWatcher();
  }

  vscode.window.showInformationMessage(
    "Excel 已打开（前台可见）。直接在 AI 对话中提出要求即可，AI 会通过 excelcli 操作这个工作簿。"
  );
  pushStateToWebview();
}

async function handleSelectSyncDirectory(): Promise<void> {
  const uris = await vscode.window.showOpenDialog({
    canSelectFiles: false,
    canSelectFolders: true,
    canSelectMany: false,
    title: "Excel VBA Assistant - 选择本地同步目录",
  });
  if (!uris || uris.length === 0) return;
  await setSyncDirectory(uris[0].fsPath, false);

  // 选择目录为独立操作，不再自动触发同步
}

async function handleUseCurrentWorkspace(): Promise<void> {
  const folders = vscode.workspace.workspaceFolders;
  if (!folders || folders.length === 0) {
    vscode.window.showWarningMessage("当前没有打开工作区，请先打开一个文件夹，或手动选择同步目录。");
    return;
  }
  const wsPath = folders[0].uri.fsPath;
  await setSyncDirectory(wsPath, false);
}

async function handleUseExcelSameDirectory(): Promise<void> {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) {
    vscode.window.showWarningMessage("请先选择 Excel 文件，再使用 Excel 相同目录。");
    return;
  }
  // 在 Excel 所在目录创建同名子目录作为同步目录
  const defaultDir = deriveDefaultSyncDir(workbookPath);

  try {
    await mkdir(defaultDir, { recursive: true });
    await setSyncDirectory(defaultDir, true);
    output.info(`已在 Excel 相同目录创建同步目录：${defaultDir}`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.error(`设置 Excel 相同目录失败：${msg}`);
    vscode.window.showErrorMessage(`设置 Excel 相同目录失败：${msg}`);
  }
}

/** 手动断开 Excel 连接：关闭 MCP session、等待/强制结束 Excel 进程并清理状态 */
async function handleDisconnectExcel(): Promise<void> {
  output.info("用户手动断开 Excel 连接，开始重置服务状态");

  // 先读取旧值，用于后续清理
  const syncDir = stateManager.get("syncDirectory");
  const autoCreated = stateManager.get("syncDirectoryAutoCreated");
  const keepOnTop = stateManager.get("keepExcelOnTop");
  const workbookPath = stateManager.get("workbookPath");
  const excelProcessId = stateManager.get("excelProcessId");

  // 第一步：立即停止所有 watcher、轮询和等待器，避免清理过程中触发同步
  stopFileWatcher();
  stopVbeToLocalWatcher();
  stopOpenWaiter();

  // 第二步：优先通过 excelcli 关闭会话（daemon 保存并关闭工作簿），
  // 失败时回退 COM 直接保存关闭并退出 Excel
  if (workbookPath && excelProcessId > 0) {
    let closed = false;
    const cliPath = getExcelCliPath();
    const sid = stateManager.get("cliSessionId");
    if (cliPath && sid) {
      const r = await runExcelCli(cliPath, ["-q", "session", "close", "--session", sid, "--save"], 60000);
      closed = r.success;
      if (closed) {
        output.info("excelcli 会话已关闭并保存工作簿");
      } else {
        output.warn(`excelcli 关闭会话失败：${r.error ?? "未知原因"}，回退 COM 关闭`);
      }
    }
    if (!closed) {
      closed = await closeWorkbookAndQuit(excelProcessId, workbookPath);
      if (closed) {
        output.info("已通过 COM 保存并关闭工作簿");
      } else {
        output.warn(`COM 方式关闭 PID ${excelProcessId} 的 Excel 失败，稍后尝试强制终止`);
      }
    }
  }

  // 第三步：等待 Excel 进程退出，最多 5 秒；若未退出则强制终止
  if (excelProcessId > 0) {
    output.info(`等待 Excel 进程 ${excelProcessId} 退出（最多 5 秒）`);
    const exited = await waitForProcessExit(excelProcessId, 5000);
    if (exited) {
      output.info("Excel 进程已正常退出");
    } else {
      output.warn(`Excel 进程 ${excelProcessId} 未在 5 秒内退出，执行强制终止`);
      try {
        await terminateProcessTree(excelProcessId);
        output.info("已强制终止 Excel 进程");
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        output.warn(`强制终止 Excel 进程失败：${msg}`);
      }
    }
  }

  // 清空 MCP 环境变量（保留 server 条目，避免下次显示 no tools）
  await removeLegacyMcpConfig().catch(() => {});

  // 第四步：清空所有持久化和运行期状态
  await stateManager.set("workbookPath", "");
  await stateManager.set("syncDirectory", "");
  await stateManager.set("syncDirectoryAutoCreated", false);
  await stateManager.set("previousSyncDirectory", "");
  await stateManager.set("excelProcessId", 0);
  await stateManager.set("cliSessionId", "");
  await stateManager.set("autoSync", false);
  await stateManager.set("autoRunVba", false);
  await stateManager.set("keepExcelOnTop", false);

  pendingChanges = [];
  lastVbeChecksum = "";
  lastExcelAvailable = false;
  stateManager.setRuntime({
    serviceStatus: "disconnected",
    lastSyncDirection: undefined,
    lastSyncAt: undefined,
    lastError: undefined,
    warningCount: 0,
    isSyncing: false,
  });

  // 状态清空后立即刷新面板，让用户看到重置效果
  pushStateToWebview();

  // 第五步：执行可能耗时的 UI/文件清理
  // 取消 Excel 窗口置顶（使用 COM）
  if (keepOnTop && workbookPath) {
    try {
      const client = new VbaClient(workbookPath, excelProcessId || undefined);
      const result = await client.setWindowTopMost(false);
      if (result.success) {
        output.info("已取消 Excel 置顶");
      } else {
        output.warn(`取消 Excel 置顶失败：${result.message}`);
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      output.warn(`取消 Excel 置顶失败：${msg}`);
    }
  }

  // 从工作区中移除同步目录根文件夹（如果它是作为根目录添加的）
  const folders = vscode.workspace.workspaceFolders;
  if (folders && syncDir) {
    const index = folders.findIndex((f) => f.uri.fsPath === syncDir);
    if (index !== -1) {
      vscode.workspace.updateWorkspaceFolders(index, 1);
      output.info(`已从工作区移除同步目录：${syncDir}`);
    }
  }

  // 尝试删除自动创建的同步目录，删除失败也不阻塞后续清理
  if (syncDir && autoCreated) {
    const removed = await removeAutoCreatedSyncDir(syncDir);
    if (!removed) {
      output.warn("手动断开：自动同步目录删除失败，将保留目录但清空插件状态");
    }
  }

  output.info("服务状态已重置");
  viewProvider.refresh();
  vscode.window.showInformationMessage("已断开与 Excel 的连接");
}

// ============================================================
// 等待工作簿被打开（open-waiter）
// ============================================================

/**
 * 工作簿归属模型下，插件不代开工作簿：选择文件后进入等待状态，
 * 轮询检测工作簿被打开（AI 的 MCP server 以 show:true 打开，或用户手动打开），
 * 一旦检测到即附着 COM 并转入 connected。
 */
function startOpenWaiter(filePath: string): void {
  stopOpenWaiter();
  openWaiterDeadline = Date.now() + OPEN_WAITER_TIMEOUT_MS;
  output.info(`开始等待工作簿被打开：${filePath}`);
  openWaiterTimer = setInterval(() => {
    void checkWorkbookOpened(filePath);
  }, OPEN_WAITER_INTERVAL_MS);
  void checkWorkbookOpened(filePath);
}

function stopOpenWaiter(): void {
  if (openWaiterTimer) {
    clearInterval(openWaiterTimer);
    openWaiterTimer = undefined;
    output.info("已停止等待工作簿打开的检测");
  }
}

async function checkWorkbookOpened(filePath: string): Promise<void> {
  let pid: number | undefined;
  try {
    pid = await findExcelPidForWorkbook(filePath);
  } catch {
    return; // 单次探测失败不影响下一轮
  }

  if (pid) {
    stopOpenWaiter();
    await stateManager.set("excelProcessId", pid);
    stateManager.setRuntime({ serviceStatus: "connected", lastError: undefined });
    lastExcelAvailable = true;
    output.info(`检测到工作簿已打开（PID ${pid}），Excel 已连接`);
    if (stateManager.get("autoSync") && stateManager.get("syncDirectory")) {
      startFileWatcher();
      startVbeToLocalWatcher();
    }
    pushStateToWebview();
    return;
  }

  if (Date.now() > openWaiterDeadline) {
    stopOpenWaiter();
    stateManager.setRuntime({ serviceStatus: "disconnected", lastError: undefined });
    output.warn("等待工作簿打开超时（10 分钟），已停止检测。可重新选择文件或让 AI 打开。");
    pushStateToWebview();
  }
}

// ============================================================
// 同步目录管理
// ============================================================

/** 根据 Excel 文件路径推导默认同步目录：<excelDir>/<fileNameWithoutExt> */
function deriveDefaultSyncDir(workbookPath: string): string {
  const dir = path.dirname(workbookPath);
  const base = path.basename(workbookPath, path.extname(workbookPath));
  return path.join(dir, base);
}

/** 设置同步目录，处理旧目录清理、文件监听启停、资源管理器切换 */
async function setSyncDirectory(newDir: string, autoCreated: boolean): Promise<void> {
  const currentDir = stateManager.get("syncDirectory");
  const currentAutoCreated = stateManager.get("syncDirectoryAutoCreated");

  // 记录旧目录
  if (currentDir && currentDir !== newDir) {
    await stateManager.set("previousSyncDirectory", currentDir);
  }

  await stateManager.set("syncDirectory", newDir);
  await stateManager.set("syncDirectoryAutoCreated", autoCreated);
  output.info(`已设置同步目录：${newDir}${autoCreated ? "（自动创建）" : ""}`);

  viewProvider.refresh();
  pushStateToWebview();

  if (stateManager.get("autoSync")) {
    stopFileWatcher();
    startFileWatcher();
  }

  // 在资源管理器中打开同步目录
  await openSyncDirInWorkspace(newDir);

  // 如果旧目录是自动创建的且与新目录不同，先迁移文件再删除旧目录
  if (currentAutoCreated && currentDir && currentDir !== newDir) {
    await migrateAutoCreatedSyncDir(currentDir, newDir);
    await removeAutoCreatedSyncDir(currentDir);
  }
}

/** 判断 child 是否是 parent 的子目录（或相同） */
function isSubPath(parent: string, child: string): boolean {
  const rel = path.relative(parent, child);
  return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel));
}

/** 在资源管理器中显示同步目录。优先尝试加入当前工作区，否则提示用户手动打开 */
async function openSyncDirInWorkspace(syncDir: string): Promise<void> {
  const uri = vscode.Uri.file(syncDir);
  const folders = vscode.workspace.workspaceFolders;

  try {
    // 如果当前没有工作区，直接用 openFolder 打开（会重载窗口）
    if (!folders || folders.length === 0) {
      await vscode.commands.executeCommand("vscode.openFolder", uri);
      return;
    }

    // 如果已经作为工作区根目录存在，只聚焦
    const existing = folders.find((f) => f.uri.fsPath === syncDir);
    if (existing) {
      await vscode.commands.executeCommand("revealInExplorer", uri);
      return;
    }

    // 如果同步目录已经在某个现有工作区根目录内部，直接 reveal，避免触发信任对话框且更快
    const parentFolder = folders.find((f) => isSubPath(f.uri.fsPath, syncDir));
    if (parentFolder) {
      await vscode.commands.executeCommand("revealInExplorer", uri);
      output.info(`同步目录位于已信任工作区 ${parentFolder.uri.fsPath} 内，直接显示`);
      return;
    }

    // 否则添加为新工作区根目录，并等待工作区变化事件完成后聚焦
    const added = vscode.workspace.updateWorkspaceFolders(folders.length, 0, {
      uri,
      name: path.basename(syncDir),
    });
    if (!added) {
      output.warn("无法将同步目录加入当前工作区");
      return;
    }

    // 使用一次性订阅等待工作区添加完成，通常 100ms 内触发
    await new Promise<void>((resolve) => {
      const disposable = vscode.workspace.onDidChangeWorkspaceFolders(() => {
        const nowFolders = vscode.workspace.workspaceFolders || [];
        if (nowFolders.some((f) => f.uri.fsPath === syncDir)) {
          disposable.dispose();
          resolve();
        }
      });
      // 兜底：最多等 1 秒
      setTimeout(() => {
        disposable.dispose();
        resolve();
      }, 1000);
    });
    await vscode.commands.executeCommand("revealInExplorer", uri);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`无法在资源管理器中打开同步目录：${msg}`);
  }
}

/** 删除自动创建的同步目录 */
async function removeAutoCreatedSyncDir(dir: string): Promise<boolean> {
  try {
    if (!fs.existsSync(dir)) return true;
    await rm(dir, { recursive: true, force: true });
    output.info(`已删除原自动同步目录：${dir}`);
    return true;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`删除原同步目录失败：${msg}`);
    return false;
  }
}

/** 将自动创建的旧同步目录内容迁移到新目录 */
async function migrateAutoCreatedSyncDir(oldDir: string, newDir: string): Promise<void> {
  try {
    if (!fs.existsSync(oldDir) || oldDir === newDir) return;
    await mkdir(newDir, { recursive: true });

    const entries = await readdir(oldDir, { withFileTypes: true });
    for (const entry of entries) {
      const oldPath = path.join(oldDir, entry.name);
      const newPath = path.join(newDir, entry.name);
      if (entry.isDirectory()) {
        await migrateAutoCreatedSyncDir(oldPath, newPath);
      } else if (!fs.existsSync(newPath)) {
        fs.renameSync(oldPath, newPath);
      }
    }

    output.info(`已将原自动同步目录内容迁移：${oldDir} → ${newDir}`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`迁移原同步目录失败：${msg}`);
  }
}

// ============================================================
// Excel Client / Sync Engine 工厂
// ============================================================

function createVbaClient(): VbaClient | null {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) {
    vscode.window.showWarningMessage("请先选择 xlsm/xlsb/xlam 文件。");
    return null;
  }
  const excelProcessId = stateManager.get("excelProcessId") || undefined;
  return new VbaClient(workbookPath, excelProcessId);
}

// ============================================================
// 同步执行
// ============================================================

async function executeSync(direction: "vbe-to-local" | "local-to-vbe"): Promise<void> {
  if (isSyncing) {
    vscode.window.showWarningMessage("正在同步中，请稍候。");
    return;
  }
  const client = createVbaClient();
  if (!client) return;
  const syncDir = stateManager.get("syncDirectory");
  if (!syncDir) {
    vscode.window.showWarningMessage("请先选择本地同步目录。");
    return;
  }

  // 本地 → VBE 属于覆盖性写入，必须弹窗确认
  if (direction === "local-to-vbe") {
    const confirm = await vscode.window.showWarningMessage(
      "该操作会用本地 VBA 文件覆盖 Excel VBE 中的代码，是否继续？",
      { modal: true },
      "继续",
      "取消"
    );
    if (confirm !== "继续") return;
  }

  isSyncing = true;
  stateManager.setRuntime({ isSyncing: true, serviceStatus: "syncing" });
  output.info(`开始同步：${direction === "vbe-to-local" ? "VBE → 本地" : "本地 → VBE"}`);

  try {
    const result = direction === "vbe-to-local"
      ? await client.syncVbeToLocal(syncDir)
      : await client.syncLocalToVbe(syncDir);

    if (result.success) {
      // 同步过程中若连接已被重置（如 Excel 关闭或手动断开），不再更新为 synced
      if (!stateManager.get("workbookPath")) {
        output.warn("同步完成时检测到已断开连接，忽略同步结果");
        return;
      }
      const summary = (result.output || "同步完成").split("\n").pop() || "同步完成";
      stateManager.setRuntime({
        lastSyncDirection: direction,
        lastSyncAt: new Date().toISOString(),
        lastError: undefined,
        serviceStatus: "synced",
      });
      output.info(summary);
      if (result.output) output.log(result.output);

      // 自动执行 VBA：本地 → VBE 成功后使用 MCP
      if (direction === "local-to-vbe" && stateManager.get("autoRunVba")) {
        await tryAutoRunMacro();
      }

      // Excel 置顶：同步完成后生效
      if (direction === "vbe-to-local") {
        await applyExcelOnTopIfNeeded();
      }
    } else {
      stateManager.setRuntime({ lastError: result.message, serviceStatus: "error" });
      output.error(result.message);
      if (result.output) output.log(result.output);
      output.show(true);
      vscode.window.showErrorMessage(result.message);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    stateManager.setRuntime({ lastError: `同步异常：${msg}`, serviceStatus: "error" });
    output.error(`同步异常：${msg}`);
    vscode.window.showErrorMessage(`同步失败：${msg}`);
  } finally {
    isSyncing = false;
    stateManager.setRuntime({ isSyncing: false });
  }
}

// ============================================================
// 自动执行 VBA
// ============================================================

interface MacroInfo {
  name: string;
  module: string;
  procedure: string;
}

/** 通过 COM 列出宏。listMacros 输出为 JSON：{macros: [{name, module, procedure}]} */
async function listMacrosViaCom(client: VbaClient): Promise<MacroInfo[] | null> {
  try {
    const result = await client.listMacros();
    if (!result.success || !result.output) {
      output.warn(`获取宏列表失败：${result.message}`);
      vscode.window.showWarningMessage(`获取宏列表失败：${result.message}`);
      return null;
    }
    const parsed = JSON.parse(result.output) as { macros?: MacroInfo[] };
    return parsed.macros ?? [];
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`获取宏列表失败：${msg}`);
    vscode.window.showWarningMessage(`获取宏列表失败：${msg}`);
    return null;
  }
}

async function tryAutoRunMacro(): Promise<void> {
  let macroName = stateManager.get("autoRunMacroName");
  const client = createVbaClient();
  if (!client) return;
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) return;

  // 未设置宏名时弹出宏选择列表
  if (!macroName) {
    const macros = await listMacrosViaCom(client);
    if (!macros) return;
    if (macros.length === 0) {
      vscode.window.showWarningMessage("工作簿中未找到任何宏。");
      return;
    }
    const picked = await vscode.window.showQuickPick(
      macros.map((m) => ({ label: m.name, description: m.module, detail: m.procedure })),
      { title: "选择要运行的宏", placeHolder: "Module1.Main" }
    );
    if (!picked) return;
    macroName = picked.label;
    await stateManager.set("autoRunMacroName", macroName);
  }

  // 弹窗确认，禁止无确认自动执行
  const confirm = await vscode.window.showInformationMessage(
    `代码已同步到 VBE，是否运行宏：${macroName}？`,
    "运行",
    "取消"
  );
  if (confirm !== "运行") return;

  // COM 运行宏：runMacroWithDialogHandling 自带 Excel 弹窗检测与交互处理
  const result = await client.runMacro(macroName);
  if (result.success) {
    output.info(`宏 ${macroName} 执行成功`);
    if (result.message) output.log(result.message);
    vscode.window.showInformationMessage(`宏 ${macroName} 已执行`);
  } else {
    output.error(`宏 ${macroName} 执行失败：${result.message}`);
    output.show(true);
    vscode.window.showErrorMessage(`宏执行失败：${result.message}`);
  }
}

// ============================================================
// 运行宏 / 刷新资源
// ============================================================

async function handleRunMacro(): Promise<void> {
  const client = createVbaClient();
  if (!client) return;

  const macros = await listMacrosViaCom(client);
  if (!macros) return;
  if (macros.length === 0) {
    vscode.window.showWarningMessage("工作簿中未找到任何宏。");
    return;
  }
  const picked = await vscode.window.showQuickPick(
    macros.map((m) => ({ label: m.name, description: m.module, detail: m.procedure })),
    { title: "选择要运行的宏" }
  );
  if (!picked) return;
  output.info(`开始运行宏：${picked.label}`);
  // COM 运行宏：runMacroWithDialogHandling 自带 Excel 弹窗检测与交互处理
  const result = await client.runMacro(picked.label);
  if (result.success) {
    output.info(`宏 ${picked.label} 执行成功`);
    if (result.message) output.log(result.message);
    vscode.window.showInformationMessage(`宏 ${picked.label} 已执行`);
  } else {
    output.error(`宏 ${picked.label} 执行失败：${result.message}`);
    output.show(true);
    vscode.window.showErrorMessage(`宏执行失败：${result.message}`);
  }
}

async function handleRefreshResources(): Promise<void> {
  const client = createVbaClient();
  if (!client) return;

  output.info("刷新 VBA 资源列表");
  try {
    const result = await client.getResources();
    if (!result.success || !result.output) {
      output.warn(`刷新资源失败：${result.message}`);
      return;
    }
    const parsed = JSON.parse(result.output);
    const items = parsed.items ?? [];
    output.info(`共 ${items.length} 个 VBA 组件`);
    for (const item of items) {
      output.info(`  ${item.name} (${item.type})`);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.error(`刷新资源失败：${msg}`);
    vscode.window.showErrorMessage(`刷新资源失败：${msg}`);
  }
}

// ============================================================
// 自动同步：文件监听
// ============================================================

function startFileWatcher(): void {
  stopFileWatcher();
  const syncDir = stateManager.get("syncDirectory");
  if (!syncDir) return;

  const pattern = new vscode.RelativePattern(syncDir, "**/*.{bas,cls,frm,wks,wbk}");
  fileWatcher = vscode.workspace.createFileSystemWatcher(pattern);

  const debouncedSync = (triggerUri?: vscode.Uri, eventLabel?: string): void => {
    // 忽略 Excel → 本地 同步过程中产生的文件修改事件，避免双向循环
    if (isSyncingVbeToLocal) return;
    if (syncDebounceTimer) clearTimeout(syncDebounceTimer);
    if (triggerUri && eventLabel) {
      output.info(`检测到${eventLabel}：${path.basename(triggerUri.fsPath)}，${AUTO_SYNC_DEBOUNCE_MS}ms 后入队`);
    }
    syncDebounceTimer = setTimeout(() => {
      pendingChanges.push({ source: "local", timestamp: Date.now() });
      void processPendingChanges();
    }, AUTO_SYNC_DEBOUNCE_MS);
  };

  fileWatcher.onDidChange((uri) => debouncedSync(uri, "文件修改"));
  fileWatcher.onDidCreate((uri) => debouncedSync(uri, "文件创建"));
  fileWatcher.onDidDelete((uri) => {
    output.warn(`本地文件被删除：${uri.fsPath}（Sheet 模块删除不会影响 VBE 对象模块）`);
    debouncedSync(uri, "文件删除");
  });
  output.info(`已启动文件监听：${syncDir}`);
}

function stopFileWatcher(): void {
  if (fileWatcher) {
    fileWatcher.dispose();
    fileWatcher = undefined;
    output.info("已停止文件监听");
  }
  if (syncDebounceTimer) {
    clearTimeout(syncDebounceTimer);
    syncDebounceTimer = undefined;
  }
}

/** 自动同步：静默执行本地 → VBE，仅记录日志与状态栏 */
async function syncLocalToVbeQuiet(skipQueue = false): Promise<void> {
  if (isSyncing) return;
  const client = createVbaClient();
  if (!client) return;
  const syncDir = stateManager.get("syncDirectory");
  if (!syncDir) return;

  // 如果当前没有活动编辑器或文件不在同步目录，跳过（避免误触发）
  const activeEditor = vscode.window.activeTextEditor;
  if (activeEditor) {
    const docPath = activeEditor.document.uri.fsPath;
    if (!docPath.startsWith(syncDir)) return;
  }

  isSyncing = true;
  stateManager.setRuntime({ isSyncing: true, serviceStatus: "syncing" });
  output.info("自动同步：本地 → VBE");

  try {
    const result = await client.syncLocalToVbe(syncDir);
    if (result.success) {
      const summary = (result.output || "自动同步完成").split("\n").pop() || "自动同步完成";
      stateManager.setRuntime({
        lastSyncDirection: "local-to-vbe",
        lastSyncAt: new Date().toISOString(),
        lastError: undefined,
        serviceStatus: "synced",
      });
      output.info(summary);
      if (result.output) output.log(result.output);

      // 同步完成后 VBE 与本地一致，更新 checksum
      await captureVbeChecksum();

      // 注意：自动同步不触发运行宏弹窗，避免打断用户编辑
    } else {
      stateManager.setRuntime({ lastError: `自动同步失败：${result.message}`, serviceStatus: "error" });
      output.error(`自动同步失败：${result.message}`);
      if (result.output) output.log(result.output);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    stateManager.setRuntime({ lastError: `自动同步异常：${msg}`, serviceStatus: "error" });
    output.error(`自动同步异常：${msg}`);
  } finally {
    isSyncing = false;
    stateManager.setRuntime({ isSyncing: false });
    if (!skipQueue) {
      void processPendingChanges();
    }
  }
}

// ============================================================
// 自动同步：Excel → 本地（新增，不改动现有手工/本地→VBE 同步）
// ============================================================

function startVbeToLocalWatcher(): void {
  stopVbeToLocalWatcher();
  const syncDir = stateManager.get("syncDirectory");
  const workbookPath = stateManager.get("workbookPath");
  if (!stateManager.get("autoSync") || !syncDir || !workbookPath) return;

  // 启动轮询前先记录一次当前 VBE checksum，避免首次轮询误判
  void captureVbeChecksum();

  vbeToLocalCheckTimer = setInterval(() => {
    void detectVbeChangesAndQueue();
  }, AUTO_VBE_TO_LOCAL_INTERVAL_MS);
  output.info("已启动 Excel → 本地 自动同步轮询");
}

function stopVbeToLocalWatcher(): void {
  if (vbeToLocalCheckTimer) {
    clearInterval(vbeToLocalCheckTimer);
    vbeToLocalCheckTimer = undefined;
    output.info("已停止 Excel → 本地 自动同步轮询");
  }
}

/** 自动同步：静默执行 Excel → 本地，复用手工同步 syncVbeToLocal */
async function syncVbeToLocalQuiet(skipQueue = false): Promise<void> {
  if (isSyncing) return;
  const client = createVbaClient();
  if (!client) return;
  const syncDir = stateManager.get("syncDirectory");
  if (!syncDir) return;

  // 保护本地未保存修改：若本地 VBA 文件比上次同步时间新，则跳过，避免覆盖用户正在编辑的代码
  const lastSyncAt = stateManager.getAll().lastSyncAt;
  if (lastSyncAt) {
    const localHasChanges = await checkLocalVbaFilesNewerThan(syncDir, new Date(lastSyncAt));
    if (localHasChanges) {
      output.info("自动同步：本地 VBA 文件有更新，跳过 Excel → 本地");
      return;
    }
  }

  isSyncing = true;
  isSyncingVbeToLocal = true;
  stateManager.setRuntime({ isSyncing: true, serviceStatus: "syncing" });
  output.info("自动同步：Excel → 本地");

  try {
    const result = await client.syncVbeToLocal(syncDir);
    if (result.success) {
      const summary = (result.output || "自动同步完成").split("\n").pop() || "自动同步完成";
      stateManager.setRuntime({
        lastSyncDirection: "vbe-to-local",
        lastSyncAt: new Date().toISOString(),
        lastError: undefined,
        serviceStatus: "synced",
      });
      output.info(summary);
      if (result.output) output.log(result.output);
      await applyExcelOnTopIfNeeded();
    } else {
      stateManager.setRuntime({ lastError: `自动同步失败：${result.message}`, serviceStatus: "error" });
      output.error(`自动同步失败：${result.message}`);
      if (result.output) output.log(result.output);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    stateManager.setRuntime({ lastError: `自动同步异常：${msg}`, serviceStatus: "error" });
    output.error(`自动同步异常：${msg}`);
  } finally {
    // 延迟释放标志，给文件系统 watcher 一段缓冲期，避免本次写入触发本地 → VBE 同步
    setTimeout(() => {
      isSyncingVbeToLocal = false;
      isSyncing = false;
      stateManager.setRuntime({ isSyncing: false });
      if (!skipQueue) {
        void processPendingChanges();
      }
    }, 500);
  }
}

/** 捕获当前 VBE 代码 checksum */
async function captureVbeChecksum(): Promise<void> {
  const client = createVbaClient();
  if (!client) return;
  try {
    const result = await client.getVbeCodeChecksum();
    if (result.success && result.output) {
      lastVbeChecksum = result.output.trim();
    }
  } catch { /* ignore */ }
}

/** 检测 VBE 是否有新变更，有则加入队列 */
async function detectVbeChangesAndQueue(): Promise<void> {
  if (isSyncing) return;
  const client = createVbaClient();
  if (!client) return;

  let currentChecksum = "";
  try {
    const result = await client.getVbeCodeChecksum();
    if (result.success && result.output) {
      currentChecksum = result.output.trim();
    }
  } catch {
    return;
  }

  if (!lastVbeChecksum) {
    lastVbeChecksum = currentChecksum;
    return;
  }

  if (currentChecksum !== lastVbeChecksum) {
    output.info("检测到 VBE 有变更，加入同步队列");
    lastVbeChecksum = currentChecksum;
    pendingChanges.push({ source: "vbe", timestamp: Date.now() });
    void processPendingChanges();
  }
}

/** 按时间顺序处理待同步变更队列 */
async function processPendingChanges(): Promise<void> {
  if (isSyncing || pendingChanges.length === 0) return;

  // 合并相邻同源事件，保留最新时间戳
  const merged: PendingChange[] = [];
  for (const change of pendingChanges) {
    const last = merged[merged.length - 1];
    if (last && last.source === change.source) {
      last.timestamp = Math.max(last.timestamp, change.timestamp);
    } else {
      merged.push({ ...change });
    }
  }
  pendingChanges = merged;

  const next = pendingChanges.shift();
  if (!next) return;

  output.info(`处理同步队列：${next.source === "local" ? "本地 → VBE" : "Excel → 本地"}`);
  if (next.source === "local") {
    await syncLocalToVbeQuiet(true);
  } else {
    await syncVbeToLocalQuiet(true);
  }

  // 继续处理队列中的下一项
  void processPendingChanges();
}

/** 检查同步目录下是否有 VBA 文件比指定时间更新 */
async function checkLocalVbaFilesNewerThan(syncDir: string, time: Date): Promise<boolean> {
  try {
    const files = await vscode.workspace.findFiles(new vscode.RelativePattern(syncDir, "**/*.{bas,cls,frm,wks,wbk}"));
    for (const file of files) {
      const stat = await vscode.workspace.fs.stat(file);
      if (stat.mtime > time.getTime()) {
        return true;
      }
    }
  } catch {
    // 忽略统计错误
  }
  return false;
}

// ============================================================
// 推送状态到 Webview
// ============================================================

function pushStateToWebview(): void {
  const state = stateManager.getAll();
  const payload: StatePayload = {
    workbookPath: state.workbookPath,
    syncDirectory: state.syncDirectory,
    autoRunVba: state.autoRunVba,
    autoSync: state.autoSync,
    keepExcelOnTop: state.keepExcelOnTop,
    serviceStatus: state.serviceStatus,
    lastError: state.lastError,
    warningCount: state.warningCount,
  };
  viewProvider.postState(payload);
}

// ============================================================
// Excel 窗口置顶
// ============================================================

/** 若用户开启了置顶选项且 Excel 已连接，将 Excel 主窗口置顶（使用 COM） */
async function applyExcelOnTopIfNeeded(): Promise<void> {
  if (!stateManager.get("keepExcelOnTop")) return;
  const client = createVbaClient();
  if (!client) return;
  try {
    const result = await client.setWindowTopMost(true);
    if (result.success) {
      output.info("Excel 已置顶");
    } else {
      output.warn(`Excel 置顶失败：${result.message}`);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`Excel 置顶失败：${msg}`);
  }
}

/** 恢复 Excel 主窗口为正常状态（取消置顶，使用 COM） */
async function restoreExcelWindowState(): Promise<void> {
  const client = createVbaClient();
  if (!client) return;
  try {
    const result = await client.setWindowTopMost(false);
    if (result.success) {
      output.info("Excel 置顶已取消");
    } else {
      output.warn(`取消 Excel 置顶失败：${result.message}`);
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`取消 Excel 置顶失败：${msg}`);
  }
}

// ============================================================
// excelcli 工具 & 旧版 MCP 配置清理
// ============================================================

/** 定位插件内置的 excelcli.exe（CLI daemon 模式的唯一 Excel 持有者入口） */
function getExcelCliPath(): string | undefined {
  if (!extensionContext) return undefined;
  const p = findLocalExcelCli(path.join(extensionContext.extensionPath, "dist"));
  if (!p) output.warn("未找到 excelcli.exe（插件安装不完整，或未执行 npm run download:cli）");
  return p;
}

/** 复制 excelcli 完整路径与常用命令到剪贴板，方便在终端验证 / 交给 AI 使用 */
async function copyCliPath(): Promise<void> {
  if (!extensionContext) {
    vscode.window.showErrorMessage("插件尚未完全激活，请稍后再试。");
    return;
  }
  const cliPath = getExcelCliPath();
  if (!cliPath) {
    vscode.window.showErrorMessage("未找到 excelcli.exe。请确认插件已正确安装。");
    return;
  }
  const text = [
    "excelcli 完整路径（在终端或让 AI 使用此路径操作 Excel）：",
    cliPath,
    "",
    "常用命令：",
    `& '${cliPath}' -q session list`,
    `& '${cliPath}' -q session open "D:\\path\\book.xlsm" --show`,
    `& '${cliPath}' -q sheet list --session <sessionId>`,
    `& '${cliPath}' -q range set-values --session <sessionId> --sheet Sheet1 --range A1:B1 --values '[["姓名","年龄"]]'`,
  ].join("\n");
  try {
    await vscode.env.clipboard.writeText(text);
    output.info("已复制 excelcli 路径到剪贴板");
    vscode.window.showInformationMessage("excelcli 路径已复制到剪贴板。");
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    vscode.window.showErrorMessage(`复制到剪贴板失败：${msg}`);
  }
}

/**
 * 清理旧版本（v0.8.x MCP 架构）遗留的 .trae/mcp.json 中 excel-mcp 条目。
 * CLI 架构不再需要任何 MCP 配置。
 */
async function removeLegacyMcpConfig(): Promise<void> {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) return;
  for (const wf of workspaceFolders) {
    try {
      await removeMcpServerEntry(vscode.Uri.joinPath(wf.uri, ".trae", "mcp.json"), "excel-mcp");
    } catch { /* ignore */ }
  }
}

async function removeMcpServerEntry(uri: vscode.Uri, name: string): Promise<void> {
  let json: { mcpServers?: Record<string, unknown> } = {};
  try {
    const data = await vscode.workspace.fs.readFile(uri);
    json = JSON.parse(Buffer.from(data).toString("utf-8"));
  } catch {
    return;
  }
  if (!json.mcpServers) return;
  delete json.mcpServers[name];
  const content = JSON.stringify(json, null, 2);
  await vscode.workspace.fs.writeFile(uri, Buffer.from(content, "utf-8"));
}

async function setExcelVbaIconTheme(): Promise<void> {
  try {
    const config = vscode.workspace.getConfiguration("workbench");
    const current = config.get<string>("iconTheme");
    // 仅在用户未设置任何图标主题时自动应用，避免覆盖用户已有偏好
    if (!current || current === "vs-seti" || current === "vs-minimal") {
      await config.update("iconTheme", "excelVbaIcons", true);
      output.info("已自动设置文件图标主题为 Excel VBA Icons");
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`自动设置图标主题失败：${msg}`);
  }
}

/**
 * 将插件内置的 skills 目录递归同步到当前工作区的 .trae/skills/，
 * 让 Trae AI 能按需加载（excel-vba-assistant 编排文档 + 官方 excel-cli 31 组命令文档）。
 * 复制文本文件时替换 {{EXCELCLI_PATH}} 占位符为内置 excelcli.exe 的真实路径。
 */
async function syncBuiltinSkillsToWorkspace(showMessage = false): Promise<void> {
  if (!extensionContext) return;
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    if (showMessage) {
      void vscode.window.showWarningMessage("未打开工作区，无法同步 Skill。");
    }
    return;
  }

  const sourceDir = path.join(extensionContext.extensionPath, "skills");
  if (!fs.existsSync(sourceDir)) {
    if (showMessage) {
      void vscode.window.showWarningMessage("插件内置 Skill 目录不存在。");
    }
    return;
  }

  const cliPath = findLocalExcelCli(path.join(extensionContext.extensionPath, "dist"))
    || "excelcli.exe 缺失（插件安装不完整，请重新安装）";

  let totalCopied = 0;
  const copyDir = async (src: string, dest: string): Promise<void> => {
    await mkdir(dest, { recursive: true });
    const entries = await readdir(src, { withFileTypes: true });
    for (const entry of entries) {
      const srcPath = path.join(src, entry.name);
      const destPath = path.join(dest, entry.name);
      if (entry.isDirectory()) {
        await copyDir(srcPath, destPath);
        continue;
      }
      if (!entry.isFile()) continue;
      let shouldCopy = true;
      if (fs.existsSync(destPath)) {
        shouldCopy = fs.statSync(srcPath).mtime > fs.statSync(destPath).mtime;
      }
      if (shouldCopy) {
        if (/\.(md|txt|json)$/i.test(entry.name)) {
          const content = fs.readFileSync(srcPath, "utf-8").split("{{EXCELCLI_PATH}}").join(cliPath);
          fs.writeFileSync(destPath, content, "utf-8");
        } else {
          fs.copyFileSync(srcPath, destPath);
        }
        totalCopied++;
      }
    }
  };

  for (const wf of workspaceFolders) {
    const targetDir = path.join(wf.uri.fsPath, ".trae", "skills");
    try {
      await copyDir(sourceDir, targetDir);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      output.warn(`同步 Skill 到工作区失败：${msg}`);
      if (showMessage) {
        void vscode.window.showErrorMessage(`同步 Skill 失败：${msg}`);
      }
      return;
    }
  }

  output.info(`已同步 ${totalCopied} 个 Skill 文件到 .trae/skills/`);
  if (showMessage) {
    void vscode.window.showInformationMessage(`已同步 ${totalCopied} 个 Skill 文件到当前工作区。`);
  }
}

