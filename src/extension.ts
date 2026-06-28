/**
 * Excel VBA Assistant - 扩展入口
 *
 * 职责：
 * 1. 注册 Webview 面板、命令、状态栏、输出通道
 * 2. 处理 Webview 消息（选择文件/目录、同步、开关）
 * 3. 自动同步：监听本地 *.bas/*.cls/*.frm 变化，防抖后写回 VBE
 * 4. 自动执行 VBA：本地 → VBE 同步后，确认运行宏
 * 5. 写入 .trae/mcp.json，让 Trae AI 通过 MCP 调用 Excel/VBA 工具
 */
import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs";
import { mkdir, readdir, rm } from "fs/promises";

import { StateManager } from "./state";
import { OutputManager } from "./output/outputChannel";
import { StatusBarManager } from "./status/statusBar";
import { ExcelVbaPanelProvider, type WebviewMessage, type StatePayload } from "./webview";
import { VbaClient } from "./client/vbaClient";
import { escapePowerShellSingleQuoted, runPowerShell } from "./runtime/powershell";

const SUPPORTED_EXCEL_EXTENSIONS = ["xlsm", "xlsb", "xlam", "xls"];
const VBA_FILE_EXTENSIONS = [".bas", ".cls", ".frm", ".wks", ".wbk"];
const MCP_SERVER_NAME = "excel-vba-mcp-server";

// 自动同步防抖毫秒数。太短容易连续触发，太长体感延迟高；300ms 在敲击停止后几乎无感知。
const AUTO_SYNC_DEBOUNCE_MS = 300;

// Excel → 本地 自动同步轮询间隔。Excel/VBE 没有直接文件系统事件，只能轮询检测变更。
const AUTO_VBE_TO_LOCAL_INTERVAL_MS = 3000;

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
let excelCheckTimer: NodeJS.Timeout | undefined;
let vbeToLocalCheckTimer: NodeJS.Timeout | undefined;
let isSyncing = false;
let isSyncingVbeToLocal = false;
let lastExcelAvailable = false;
let excelUnavailableCount = 0;
const EXCEL_UNAVAILABLE_THRESHOLD = 2;
let lastExcelProcessAlive = false;
let lastExcelWindowVisible = false;

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

  // 自动同步开关已开启时，启动双向自动同步
  if (stateManager.get("autoSync") && stateManager.get("syncDirectory")) {
    startFileWatcher();
    startVbeToLocalWatcher();
  }

  // 立即推送一次已持久化的状态，避免 webview 启动时显示空白/丢失
  pushStateToWebview();

  // 插件激活时立即写入基础 MCP 配置；选择 Excel 文件/目录后会再次更新
  void writeMcpConfig();

  // 将内置 Skill 同步到当前工作区的 .trae/skills，使 Trae AI 能按需加载
  void syncBuiltinSkillsToWorkspace();

  // 恢复时若已保存 workbookPath，自动检测 Excel 连接状态，避免面板显示「未知」
  // 检测器在恢复成功后启动，防止恢复完成前误触发清理
  void restoreExcelConnectionStatus().then(() => {
    startExcelCloseWatcher();
  });
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

  const client = new VbaClient(workbookPath);
  let accessError = await client.checkAccess();
  let excelPid = 0;
  if (accessError) {
    output.warn(`恢复连接状态时 Excel 未就绪，尝试自动打开：${accessError}`);
    const openResult = await client.openWorkbookInExcel();
    if (openResult.success) {
      output.info(openResult.message);
      excelPid = (openResult.details?.pid as number) || 0;
      accessError = await client.checkAccess();
    } else {
      output.error(`恢复连接状态时自动打开 Excel 失败：${openResult.message}`);
    }
  }
  // 如果当前没有记录 PID，尝试获取已运行 Excel 的 PID
  if (excelPid === 0) {
    excelPid = await getExcelProcessId(workbookPath);
  }
  if (accessError) {
    stateManager.setRuntime({ lastError: accessError, serviceStatus: "error" });
    output.warn(`恢复连接状态时检测到错误：${accessError}`);
    // 服务未连接时，强制关闭自动同步和自动执行 VBA
    await stateManager.set("autoSync", false);
    await stateManager.set("autoRunVba", false);
    stopFileWatcher();
    pushStateToWebview();
    return false;
  }
  stateManager.setRuntime({ serviceStatus: "connected" });
  output.info("恢复连接状态：Excel 已连接");
  if (excelPid > 0) {
    await stateManager.set("excelProcessId", excelPid);
  }
  lastExcelAvailable = true;

  // 恢复连接后，自动创建同步目录并执行 VBE → 本地 同步（与选择文件流程一致）
  let syncDir = stateManager.get("syncDirectory");
  if (!syncDir) {
    const defaultDir = deriveDefaultSyncDir(workbookPath);
    try {
      await mkdir(defaultDir, { recursive: true });
      await setSyncDirectory(defaultDir, true);
      syncDir = defaultDir;
      output.info(`恢复连接时自动创建同步目录：${defaultDir}`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      output.error(`恢复连接时自动创建同步目录失败：${msg}`);
    }
  }

  if (syncDir) {
    output.info("恢复连接后开始自动执行 VBE → 本地 同步");
    await executeSync("vbe-to-local");

    // 同步成功后自动开启自动同步和自动执行 VBA（除非用户之后手工关闭）
    await stateManager.set("autoSync", true);
    await stateManager.set("autoRunVba", true);
    startFileWatcher();
    startVbeToLocalWatcher();
    output.info("已自动开启自动同步和自动执行 VBA");

    // 恢复连接后写入 MCP 配置
    await writeMcpConfig();
  }

  pushStateToWebview();
  return true;
}

export function deactivate(): void {
  stopFileWatcher();
  stopVbeToLocalWatcher();
  stopExcelCloseWatcher();
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
}

// ============================================================
// Webview 消息处理（任务文档第五节）
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
  }
}

// ============================================================
// 选择文件 / 目录（任务文档第五节）
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

  // 需求 2：若未设置同步目录，自动在 Excel 同路径下创建同名文件夹作为同步目录
  // 如果切换到了不同的 Excel 文件，且旧同步目录是自动创建的，先清理旧目录
  let syncDir = stateManager.get("syncDirectory");
  const isSwitchingWorkbook = previousWorkbookPath && previousWorkbookPath !== filePath;
  if (isSwitchingWorkbook && previousAutoCreated && previousSyncDir) {
    await removeAutoCreatedSyncDir(previousSyncDir);
    syncDir = "";
    await stateManager.set("syncDirectory", "");
    await stateManager.set("syncDirectoryAutoCreated", false);
  }

  if (!syncDir) {
    const defaultDir = deriveDefaultSyncDir(filePath);
    try {
      await mkdir(defaultDir, { recursive: true });
      await setSyncDirectory(defaultDir, true);
      syncDir = defaultDir;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      output.error(`自动创建同步目录失败：${msg}`);
      vscode.window.showErrorMessage(`自动创建同步目录失败：${msg}`);
    }
  }

  viewProvider.refresh();
  pushStateToWebview();

  // 检查 Excel 访问；若 Excel 未打开或文件未加载，尝试自动打开
  const client = createVbaClient();
  let excelPid = 0;
  if (client) {
    let accessError = await client.checkAccess();
    if (accessError) {
      output.warn(`Excel 未就绪，尝试自动打开：${accessError}`);
      const openResult = await client.openWorkbookInExcel();
      if (openResult.success) {
        output.info(openResult.message);
        excelPid = (openResult.details?.pid as number) || 0;
        // 打开后再次检查 VBAProject 访问
        accessError = await client.checkAccess();
      } else {
        output.error(`自动打开 Excel 失败：${openResult.message}`);
        vscode.window.showErrorMessage(`无法自动打开 Excel：${openResult.message}`);
      }
    }
    // 如果当前没有记录 PID，尝试获取已运行 Excel 的 PID
    if (excelPid === 0) {
      excelPid = await getExcelProcessId(filePath);
    }

    if (accessError) {
      output.error(accessError);
      stateManager.setRuntime({ lastError: accessError, serviceStatus: "error" });
      vscode.window.showErrorMessage(accessError);
    } else {
      stateManager.setRuntime({ serviceStatus: "connected" });
      output.info("Excel VBAProject 可访问");
      if (excelPid > 0) {
        await stateManager.set("excelProcessId", excelPid);
      }
      // Excel 连接成功后启动关闭检测器（如果尚未启动）
      startExcelCloseWatcher();
      lastExcelAvailable = true;
    }
  }

  // 已设置同步目录时直接执行 VBE → 本地 同步（不再询问）
  if (syncDir) {
    output.info("已选择 Excel 文件，开始自动执行 VBE → 本地 同步");
    await executeSync("vbe-to-local");

    // 同步成功后自动开启自动同步和自动执行 VBA（除非用户之后手工关闭）
    await stateManager.set("autoSync", true);
    await stateManager.set("autoRunVba", true);
    startFileWatcher();
    startVbeToLocalWatcher();
    output.info("已自动开启自动同步和自动执行 VBA");

    // 写入 MCP 配置，让 AI 可以直接调用 Excel/VBA 工具
    await writeMcpConfig();
  }
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

  // 已设置同步目录且已选择 Excel 文件时，写入 MCP 配置
  if (stateManager.get("workbookPath")) {
    await writeMcpConfig();
  }

  // 已选择 Excel 文件时询问是否立即同步
  const workbookPath = stateManager.get("workbookPath");
  if (workbookPath) {
    const choice = await vscode.window.showInformationMessage(
      "已设置同步目录，是否立即执行 VBE → 本地 同步？",
      "是",
      "否"
    );
    if (choice === "是") {
      await executeSync("vbe-to-local");
    }
  }
}

async function handleUseCurrentWorkspace(): Promise<void> {
  const folders = vscode.workspace.workspaceFolders;
  if (!folders || folders.length === 0) {
    vscode.window.showWarningMessage("当前没有打开工作区，请先打开一个文件夹，或手动选择同步目录。");
    return;
  }
  const wsPath = folders[0].uri.fsPath;
  await setSyncDirectory(wsPath, false);

  if (stateManager.get("workbookPath")) {
    await writeMcpConfig();
  }
}

async function handleUseExcelSameDirectory(): Promise<void> {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) {
    vscode.window.showWarningMessage("请先选择 Excel 文件，再使用 Excel 相同目录。");
    return;
  }
  const defaultDir = deriveDefaultSyncDir(workbookPath);

  try {
    await mkdir(defaultDir, { recursive: true });
    await setSyncDirectory(defaultDir, true);

    if (stateManager.get("workbookPath")) {
      await writeMcpConfig();
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.error(`设置 Excel 相同目录失败：${msg}`);
    vscode.window.showErrorMessage(`设置 Excel 相同目录失败：${msg}`);
  }
}

/** 手动断开 Excel 连接：强制重置整个服务状态（用于 Excel 异常关闭后无法自动清理的场景） */
async function handleDisconnectExcel(): Promise<void> {
  output.info("用户手动断开 Excel 连接，开始重置服务状态");

  // 先读取旧值，用于后续清理
  const syncDir = stateManager.get("syncDirectory");
  const autoCreated = stateManager.get("syncDirectoryAutoCreated");
  const keepOnTop = stateManager.get("keepExcelOnTop");
  const workbookPath = stateManager.get("workbookPath");

  // 第一步：立即停止所有 watcher 和轮询，避免清理过程中触发同步
  stopExcelCloseWatcher();
  stopFileWatcher();
  stopVbeToLocalWatcher();

  // 清理 MCP 配置
  await removeMcpConfig();

  // 第二步：立即清空所有持久化和运行期状态，确保即使 F5 中断也不会残留
  await stateManager.set("workbookPath", "");
  await stateManager.set("syncDirectory", "");
  await stateManager.set("syncDirectoryAutoCreated", false);
  await stateManager.set("previousSyncDirectory", "");
  await stateManager.set("excelProcessId", 0);
  await stateManager.set("autoSync", false);
  await stateManager.set("autoRunVba", false);
  await stateManager.set("keepExcelOnTop", false);

  pendingChanges = [];
  lastVbeChecksum = "";
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

  // 第三步：执行可能耗时的 UI/文件清理
  // 取消 Excel 窗口置顶
  if (keepOnTop && workbookPath) {
    const client = new VbaClient(workbookPath);
    const result = await client.setWindowTopMost(false);
    if (result.success) {
      output.info(result.message);
    } else {
      output.warn(`取消 Excel 置顶失败：${result.message}`);
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
// 同步目录管理（需求 2 / 3 / 4）
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

  // 在资源管理器中打开同步目录（需求 4）
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
// 同步执行（任务文档第十二、十三节）
// ============================================================

function createVbaClient(): VbaClient | null {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) {
    vscode.window.showWarningMessage("请先选择 xlsm/xlsb/xlam 文件。");
    return null;
  }
  return new VbaClient(workbookPath);
}

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

  // 本地 → VBE 属于覆盖性写入，必须弹窗确认（任务文档第十三节）
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
      const summary = (result.output || "同步完成").split("\n").pop() || "同步完成";
      stateManager.setRuntime({
        lastSyncDirection: direction,
        lastSyncAt: new Date().toISOString(),
        lastError: undefined,
        serviceStatus: "synced",
      });
      output.info(summary);
      if (result.output) output.log(result.output);

      // 自动执行 VBA：本地 → VBE 成功后（任务文档第十五节）
      if (direction === "local-to-vbe" && stateManager.get("autoRunVba")) {
        await tryAutoRunMacro();
      }

      // Excel 置顶：用户选择文件后自动同步完成时生效
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
// 自动执行 VBA（任务文档第十五节）
// ============================================================

async function tryAutoRunMacro(): Promise<void> {
  let macroName = stateManager.get("autoRunMacroName");
  const client = createVbaClient();
  if (!client) return;

  // 未设置宏名时弹出宏选择列表
  if (!macroName) {
    const macrosResult = await client.listMacros();
    if (!macrosResult.success || !macrosResult.output) {
      vscode.window.showWarningMessage("自动执行 VBA：未找到可用宏。");
      return;
    }
    let macros: Array<{ name: string; module: string; procedure: string }> = [];
    try {
      const parsed = JSON.parse(macrosResult.output);
      macros = parsed.macros || [];
    } catch {
      // ignore
    }
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

  // 弹窗确认（任务文档第十五节：禁止无确认自动执行）
  const confirm = await vscode.window.showInformationMessage(
    `代码已同步到 VBE，是否运行宏：${macroName}？`,
    "运行",
    "取消"
  );
  if (confirm !== "运行") return;

  const result = await client.runMacro(macroName);
  if (result.success) {
    output.info(`宏 ${macroName} 执行成功`);
    if (result.output) output.log(result.output);
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
  const macrosResult = await client.listMacros();
  if (!macrosResult.success || !macrosResult.output) {
    vscode.window.showWarningMessage(macrosResult.message || "未找到可用宏");
    return;
  }
  let macros: Array<{ name: string; module: string; procedure: string }> = [];
  try {
    const parsed = JSON.parse(macrosResult.output);
    macros = parsed.macros || [];
  } catch {
    // ignore
  }
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
  const result = await client.runMacro(picked.label);
  if (result.success) {
    output.info(`宏 ${picked.label} 执行成功`);
    if (result.output) output.log(result.output);
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
  const result = await client.getResources();
  if (result.success && result.output) {
    output.info(result.output);
  } else {
    output.error(result.message);
    vscode.window.showErrorMessage(result.message);
  }
}

// ============================================================
// 自动同步：文件监听（任务文档第十四节）
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

// ============================================================
// Excel 关闭检测（需求 3）
// ============================================================

const EXCEL_CHECK_INTERVAL_MS = 2000;

function startExcelCloseWatcher(): void {
  stopExcelCloseWatcher();
  excelUnavailableCount = 0;
  lastExcelProcessAlive = false;
  lastExcelWindowVisible = false;
  excelCheckTimer = setInterval(async () => {
    const workbookPath = stateManager.get("workbookPath");
    if (!workbookPath) {
      lastExcelAvailable = false;
      excelUnavailableCount = 0;
      return;
    }
    let expectedPid = stateManager.get("excelProcessId");

    // 如果未记录 PID，先尝试通过 COM 获取一次，并持久化；成功后就走 PID 检测路径
    if (!expectedPid || expectedPid <= 0) {
      const detectedPid = await getExcelProcessId(workbookPath);
      if (detectedPid > 0) {
        await stateManager.set("excelProcessId", detectedPid);
        expectedPid = detectedPid;
        output.info(`检测到 Excel 进程 PID：${detectedPid}，后续优先使用 PID 检测`);
      }
    }

    // 如果记录了 PID，优先只检查进程是否还存在，避免频繁调用 COM 导致 Excel 进程无法自然退出
    if (expectedPid && expectedPid > 0) {
      const alive = await isExcelProcessAlive(expectedPid);
      if (alive !== lastExcelProcessAlive) {
        lastExcelProcessAlive = alive;
        output.info(`Excel 进程 ${expectedPid} 状态变化：${alive ? "ALIVE" : "DEAD"}`);
      }
      if (!alive) {
        await handleExcelClosed();
        return;
      }
      // 进程还在时，进一步检查是否仍有可见窗口；部分场景 Excel 窗口已关闭但进程残留
      const hasWindow = await hasExcelVisibleWindow(expectedPid);
      if (hasWindow !== lastExcelWindowVisible) {
        lastExcelWindowVisible = hasWindow;
        output.info(`Excel 进程 ${expectedPid} 窗口可见性变化：${hasWindow ? "VISIBLE" : "HIDDEN"}`);
      }
      if (!hasWindow) {
        excelUnavailableCount++;
        output.warn(
          `Excel 进程 ${expectedPid} 仍在运行，但已无可视窗口（连续 ${excelUnavailableCount}/${EXCEL_UNAVAILABLE_THRESHOLD} 次）`
        );
        if (excelUnavailableCount >= EXCEL_UNAVAILABLE_THRESHOLD) {
          output.info("Excel 窗口已关闭且进程无可见窗口，执行关闭清理");
          await handleExcelClosed();
          excelUnavailableCount = 0;
        }
        lastExcelAvailable = false;
        return;
      }
      if (excelUnavailableCount > 0) {
        output.info("Excel 恢复可见，取消关闭计数");
      }
      excelUnavailableCount = 0;
      lastExcelAvailable = true;
      return;
    }

    // 实在拿不到 PID 时，才退回到 COM 检测
    output.warn("未记录 Excel PID，退回到 COM 可用性检测");
    const check = await isExcelWithWorkbookRunning(workbookPath);
    let available = check.running;

    // 如果工作簿仍在运行，但所在进程与记录的不一致，说明工作簿已切换到其他 Excel 实例，视为不稳定状态
    if (available && expectedPid && expectedPid > 0 && check.actualPid > 0 && check.actualPid !== expectedPid) {
      output.warn(
        `工作簿所在 Excel 实例发生变化：记录 PID=${expectedPid}，实际 PID=${check.actualPid}`
      );
      available = false;
    }

    if (!available) {
      excelUnavailableCount++;
      output.warn(`检测到 Excel 可能已关闭（连续 ${excelUnavailableCount}/${EXCEL_UNAVAILABLE_THRESHOLD} 次）`);
      if (excelUnavailableCount >= EXCEL_UNAVAILABLE_THRESHOLD) {
        output.info("连续检测不到 Excel，执行关闭清理");
        await handleExcelClosed();
        excelUnavailableCount = 0;
      }
    } else {
      if (excelUnavailableCount > 0) {
        output.info("Excel 重新变为可用，取消关闭计数");
      }
      excelUnavailableCount = 0;
    }
    lastExcelAvailable = available;
  }, EXCEL_CHECK_INTERVAL_MS);
}

function stopExcelCloseWatcher(): void {
  if (excelCheckTimer) {
    clearInterval(excelCheckTimer);
    excelCheckTimer = undefined;
  }
}

/** 检查指定 Excel 进程是否仍然存在 */
async function isExcelProcessAlive(pid: number): Promise<boolean> {
  try {
    const result = await runPowerShell(`
$ErrorActionPreference = "Stop"
try {
    $p = Get-Process -Id ${pid} -ErrorAction Stop
    Write-Output "ALIVE"
} catch {
    Write-Output "DEAD"
}
`);
    return result.success && (result.output || "").trim() === "ALIVE";
  } catch {
    return false;
  }
}

/** 检查指定 PID 的 Excel 进程是否仍有可见窗口 */
async function hasExcelVisibleWindow(pid: number): Promise<boolean> {
  try {
    const result = await runPowerShell(`
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class JrWindowChecker {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd);
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}
"@
$targetPid = ${pid}
$found = $false
$excelClassPattern = [regex]::new('^(XLMAIN|EXCEL|bosa_sdm_)', 'IgnoreCase')
[JrWindowChecker]::EnumWindows({
  param($hWnd, $lParam)
  $isVisible = [JrWindowChecker]::IsWindowVisible($hWnd)
  $isIconic = [JrWindowChecker]::IsIconic($hWnd)
  if (-not $isVisible -and -not $isIconic) { return $true }
  $winPid = [uint32]0
  [void][JrWindowChecker]::GetWindowThreadProcessId($hWnd, [ref]$winPid)
  if ($winPid -ne $targetPid) { return $true }
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrWindowChecker]::GetClassName($hWnd, $sb, $sb.Capacity)
  $className = $sb.ToString()
  if ($excelClassPattern.IsMatch($className)) {
    $found = $true
    return $false
  }
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($found) { Write-Output "VISIBLE" } else { Write-Output "HIDDEN" }
`);
    return result.success && (result.output || "").trim() === "VISIBLE";
  } catch {
    return false;
  }
}

/** 检查指定工作簿是否仍在 Excel 中打开，同时返回该工作簿所在 Excel 实例的 PID */
async function isExcelWithWorkbookRunning(
  workbookPath: string
): Promise<{ running: boolean; actualPid: number; reason: string }> {
  const defaultResult = { running: false, actualPid: 0, reason: "exception" };
  try {
    const wbName = path.basename(workbookPath);
    const script = `
$ErrorActionPreference = "Stop"
try {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32Check {
        [DllImport(\"user32.dll\")]
        public static extern bool IsWindow(IntPtr hWnd);
        [DllImport(\"user32.dll\")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }
"@
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $found = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $found = $w; break }
    }
    if ($found -eq $null) {
        $payload = @{ running = $false; actualPid = 0; reason = "no-workbook" }
        Write-Output (ConvertTo-Json $payload -Compress)
        return
    }
    $hwnd = [IntPtr]::new([long]$excel.Hwnd)
    if (-not [Win32Check]::IsWindow($hwnd)) {
        $payload = @{ running = $false; actualPid = 0; reason = "invalid-hwnd" }
        Write-Output (ConvertTo-Json $payload -Compress)
        return
    }
    $pidValue = [uint32]0
    [void][Win32Check]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
    $runningPayload = @{ running = $true; actualPid = [int]$pidValue; reason = "ok" }
    Write-Output (ConvertTo-Json $runningPayload -Compress)
} catch {
    $errPayload = @{ running = $false; actualPid = 0; reason = "exception" }
    Write-Output (ConvertTo-Json $errPayload -Compress)
}
`;
    const result = await runPowerShell(script);
    if (result.success && result.output) {
      try {
        const parsed = JSON.parse(result.output.trim());
        output.info(
          `Excel 可用性检测：running=${parsed.running}, pid=${parsed.actualPid}, reason=${parsed.reason}`
        );
        return {
          running: !!parsed.running,
          actualPid: Number(parsed.actualPid) || 0,
          reason: String(parsed.reason || "unknown"),
        };
      } catch {
        output.warn(`Excel 可用性检测返回解析失败：${result.output}`);
      }
    }
    return defaultResult;
  } catch {
    return defaultResult;
  }
}

/** 获取当前包含目标工作簿的 Excel 进程 ID */
async function getExcelProcessId(workbookPath: string): Promise<number> {
  try {
    const wbName = path.basename(workbookPath);
    const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class JrExcelPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
function Jr-GetExcelPid {
  try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $found = $false
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $found = $true; break }
    }
    if (-not $found) { Write-Output "0"; return }
    $pidValue = [uint32]0
    $hwnd = $excel.Hwnd
    $hwndPtr = [IntPtr]::new([long]$hwnd)
    [void][JrExcelPid]::GetWindowThreadProcessId($hwndPtr, [ref]$pidValue)
    Write-Output $pidValue
  } catch {
    Write-Output "0"
  }
}
Jr-GetExcelPid
`;
    const result = await runPowerShell(script);
    if (result.success && result.output) {
      const pid = parseInt(result.output.trim(), 10);
      return isNaN(pid) ? 0 : pid;
    }
  } catch {
    // ignore
  }
  return 0;
}

/** Excel 关闭后的清理：删除自动创建的同步目录，清空相关状态 */
async function handleExcelClosed(): Promise<void> {
  // 第一步：立即停止检测器，避免清理过程中再次触发检测
  stopExcelCloseWatcher();

  // 取消 Excel 窗口置顶，恢复正常状态
  if (stateManager.get("keepExcelOnTop")) {
    await restoreExcelWindowState();
  }

  const syncDir = stateManager.get("syncDirectory");
  const autoCreated = stateManager.get("syncDirectoryAutoCreated");
  let removed = true;
  if (syncDir && autoCreated) {
    removed = await removeAutoCreatedSyncDir(syncDir);
  }
  // 如果目录删除失败（通常是因为仍被占用），保留状态，避免误清空
  if (!removed) {
    output.warn("Excel 关闭后同步目录删除失败，保留当前状态等待用户处理");
    stateManager.setRuntime({ lastError: "同步目录删除失败，保留当前状态等待用户处理", serviceStatus: "error" });
    pushStateToWebview();
    return;
  }

  // Excel 关闭且目录清理成功后，移除 MCP 配置
  await removeMcpConfig();

  await stateManager.set("workbookPath", "");
  await stateManager.set("syncDirectory", "");
  await stateManager.set("syncDirectoryAutoCreated", false);
  await stateManager.set("previousSyncDirectory", "");
  // Excel 关闭后，强制关闭自动同步和自动执行 VBA
  await stateManager.set("autoSync", false);
  await stateManager.set("autoRunVba", false);
  await stateManager.set("excelProcessId", 0);
  stopExcelCloseWatcher();
  stopFileWatcher();
  stopVbeToLocalWatcher();
  stateManager.setRuntime({ serviceStatus: "disconnected" });
  viewProvider.refresh();
  pushStateToWebview();
  output.info("已清理 Excel 关闭状态，同步目录已删除");
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
  const result = await client.getVbeCodeChecksum();
  if (result.success && result.output) {
    lastVbeChecksum = result.output.trim();
  }
}

/** 检测 VBE 是否有新变更，有则加入队列 */
async function detectVbeChangesAndQueue(): Promise<void> {
  if (isSyncing) return;
  const client = createVbaClient();
  if (!client) return;

  const result = await client.getVbeCodeChecksum();
  if (!result.success || !result.output) return;
  const currentChecksum = result.output.trim();

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
// 推送状态到 Webview（任务文档第六节）
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

/** 若用户开启了置顶选项且 Excel 已连接，将 Excel 主窗口置顶 */
async function applyExcelOnTopIfNeeded(): Promise<void> {
  if (!stateManager.get("keepExcelOnTop")) return;
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) return;
  const client = new VbaClient(workbookPath);
  const result = await client.setWindowTopMost(true);
  if (result.success) {
    output.info(result.message);
  } else {
    output.warn(`Excel 置顶失败：${result.message}`);
  }
}

/** 恢复 Excel 主窗口为正常状态（取消置顶） */
async function restoreExcelWindowState(): Promise<void> {
  const workbookPath = stateManager.get("workbookPath");
  if (!workbookPath) return;
  const client = new VbaClient(workbookPath);
  const result = await client.setWindowTopMost(false);
  if (result.success) {
    output.info(result.message);
  } else {
    output.warn(`取消 Excel 置顶失败：${result.message}`);
  }
}

// ============================================================
// MCP 配置自动写入
// ============================================================

async function writeMcpConfig(): Promise<void> {
  if (!extensionContext) return;
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) return;

  const workbookPath = stateManager.get("workbookPath");
  const syncDir = stateManager.get("syncDirectory");

  const syncDirNorm = syncDir ? syncDir.replace(/\\/g, "/") : "";
  // 固定使用第一个工作区写入 MCP 配置，避免多工作区/目录切换时产生重复配置
  const targetWsFolder = workspaceFolders[0];

  const serverExePath = path.join(extensionContext.extensionPath, "dist", "mcp-server.exe");
  const envVars: Record<string, string> = {};
  if (workbookPath) envVars.VBE_FILE_PATH = workbookPath.replace(/\\/g, "/");
  if (syncDirNorm) envVars.VBE_LOCAL_DIR = syncDirNorm;

  const config = {
    command: serverExePath.replace(/\\/g, "/"),
    args: [],
    env: envVars,
  };

  const wsRoot = targetWsFolder.uri;
  try {
    const dirUri = vscode.Uri.joinPath(wsRoot, ".trae");
    try { await vscode.workspace.fs.createDirectory(dirUri); } catch { /* exists */ }
    const mcpJsonUri = vscode.Uri.joinPath(dirUri, "mcp.json");
    await upsertMcpServerConfig(mcpJsonUri, MCP_SERVER_NAME, config);
    output.info(`已写入 MCP 配置：${mcpJsonUri.fsPath}`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    output.warn(`写入 MCP 配置失败：${msg}`);
  }
}

async function removeMcpConfig(): Promise<void> {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) return;

  // 清理所有工作区中的同名 MCP 配置，防止重复或残留
  for (const wf of workspaceFolders) {
    try {
      await removeMcpServerConfig(vscode.Uri.joinPath(wf.uri, ".trae", "mcp.json"), MCP_SERVER_NAME);
    } catch { /* ignore */ }
  }
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

async function upsertMcpServerConfig(uri: vscode.Uri, name: string, config: unknown): Promise<void> {
  let json: { mcpServers?: Record<string, unknown> } = {};
  try {
    const data = await vscode.workspace.fs.readFile(uri);
    json = JSON.parse(Buffer.from(data).toString("utf-8"));
  } catch {
    json = {};
  }
  json.mcpServers = json.mcpServers || {};
  json.mcpServers[name] = config;
  const content = JSON.stringify(json, null, 2);
  await vscode.workspace.fs.writeFile(uri, Buffer.from(content, "utf-8"));
}

async function removeMcpServerConfig(uri: vscode.Uri, name: string): Promise<void> {
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

/** 将插件内置的 skills 目录同步到当前工作区的 .trae/skills/excel-vba-assistant/，让 Trae AI 按需加载 */
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

  let totalCopied = 0;
  for (const wf of workspaceFolders) {
    const targetDir = path.join(wf.uri.fsPath, ".trae", "skills", "excel-vba-assistant");
    try {
      await mkdir(targetDir, { recursive: true });
      const files = await readdir(sourceDir);
      for (const file of files) {
        const src = path.join(sourceDir, file);
        const stat = fs.statSync(src);
        if (!stat.isFile()) continue;
        const dest = path.join(targetDir, file);
        let shouldCopy = true;
        if (fs.existsSync(dest)) {
          const destStat = fs.statSync(dest);
          // 仅当内置文件更新时才覆盖，避免覆盖用户自定义内容
          shouldCopy = stat.mtime > destStat.mtime;
        }
        if (shouldCopy) {
          fs.copyFileSync(src, dest);
          totalCopied++;
        }
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      output.warn(`同步 Skill 到工作区失败：${msg}`);
      if (showMessage) {
        void vscode.window.showErrorMessage(`同步 Skill 失败：${msg}`);
      }
      return;
    }
  }

  output.info(`已同步 ${totalCopied} 个 Skill 文件到 .trae/skills/excel-vba-assistant/`);
  if (showMessage) {
    void vscode.window.showInformationMessage(`已同步 ${totalCopied} 个 Skill 文件到当前工作区。`);
  }
}
