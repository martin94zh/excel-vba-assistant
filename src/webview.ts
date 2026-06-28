/**
 * Excel VBA Assistant - Webview 侧边栏面板
 *
 * 界面布局与样式遵循任务文档第七、八节要求，
 * 使用 VS Code 主题变量适配 Trae 深色/浅色主题。
 */
import * as vscode from "vscode";
import { StateManager } from "./state";

/** Webview 发送给插件的消息类型 */
export type WebviewMessage =
  | { type: "selectWorkbook" }
  | { type: "selectSyncDirectory" }
  | { type: "useCurrentWorkspace" }
  | { type: "useExcelSameDirectory" }
  | { type: "toggleAutoRun"; value: boolean }
  | { type: "toggleAutoSync"; value: boolean }
  | { type: "toggleKeepExcelOnTop"; value: boolean }
  | { type: "syncVbeToLocal" }
  | { type: "syncLocalToVbe" }
  | { type: "disconnectExcel" };

/** 插件发送给 Webview 的状态消息 */
export interface StatePayload {
  workbookPath: string;
  syncDirectory: string;
  autoRunVba: boolean;
  autoSync: boolean;
  keepExcelOnTop: boolean;
  serviceStatus: string;
  lastError?: string;
  warningCount: number;
}

export class ExcelVbaPanelProvider implements vscode.WebviewViewProvider {
  public static readonly viewType = "excelVbaControlPanel";

  private view?: vscode.WebviewView;
  private onCommand: (msg: WebviewMessage) => void;
  private version = "0.6.6";

  constructor(
    private readonly extensionUri: vscode.Uri,
    private readonly stateManager: StateManager,
    onCommand: (msg: WebviewMessage) => void
  ) {
    this.onCommand = onCommand;
    stateManager.onChange(() => this.refresh());
    void this.loadVersion();
  }

  private async loadVersion(): Promise<void> {
    try {
      const pkgUri = vscode.Uri.joinPath(this.extensionUri, "package.json");
      const data = await vscode.workspace.fs.readFile(pkgUri);
      const pkg = JSON.parse(Buffer.from(data).toString("utf-8"));
      if (pkg.version) {
        this.version = pkg.version;
        this.refresh();
      }
    } catch { /* ignore */ }
  }

  resolveWebviewView(webviewView: vscode.WebviewView): void {
    this.view = webviewView;
    webviewView.webview.options = {
      enableScripts: true,
      localResourceRoots: [this.extensionUri],
    };
    webviewView.webview.html = this.getHtml();

    webviewView.webview.onDidReceiveMessage((msg: WebviewMessage) => {
      this.onCommand(msg);
    });
  }

  /** 重新渲染面板 */
  refresh(): void {
    if (this.view) {
      this.view.webview.html = this.getHtml();
    }
  }

  /** 向 Webview 推送状态变更（任务文档第六节） */
  postState(payload: StatePayload): void {
    if (this.view) {
      this.view.webview.postMessage({ type: "stateChanged", state: payload });
    }
  }

  private escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  private getHtml(): string {
    const state = this.stateManager.getAll();
    const workbookPath = state.workbookPath || "";
    const syncDirectory = state.syncDirectory || "";
    const autoRunVba = state.autoRunVba;
    const autoSync = state.autoSync;
    const keepExcelOnTop = state.keepExcelOnTop;
    const serviceStatus = state.serviceStatus;
    const lastError = state.lastError;
    const warningCount = state.warningCount;

    const statusLabel: Record<string, string> = {
      unknown: "未知",
      starting: "启动中",
      connected: "已连接（未同步）",
      syncing: "同步中",
      synced: "已同步",
      disconnected: "未连接",
      error: "错误",
    };
    const statusText = statusLabel[serviceStatus] || "未知";
    const statusBadgeClass =
      serviceStatus === "synced" ? "badge-ok" :
      serviceStatus === "connected" ? "badge-warn" :
      serviceStatus === "error" ? "badge-err" :
      serviceStatus === "disconnected" ? "badge-warn" :
      serviceStatus === "syncing" ? "badge-info" : "badge-info";

    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  :root {
    --excel-bg: #f6f9fc;
    --excel-surface: #ffffff;
    --excel-input: #f8fafc;
    --excel-border: #d9e5f2;
    --excel-border-strong: #bfd3ec;
    --excel-text: #172033;
    --excel-muted: #5f6f85;
    --excel-primary: #2563eb;
    --excel-primary-hover: #1d4ed8;
    --excel-secondary: #64748b;
    --excel-secondary-hover: #475569;
    --excel-switch-off: #dbe3ed;
    --excel-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
  }
  body.vscode-dark {
    --excel-bg: #000000;
    --excel-surface: #000000;
    --excel-input: #0a0a0a;
    --excel-border: #262626;
    --excel-border-strong: #404040;
    --excel-text: #e5edf8;
    --excel-muted: #a8b6ca;
    --excel-primary: #3b82f6;
    --excel-primary-hover: #60a5fa;
    --excel-secondary: #475569;
    --excel-secondary-hover: #64748b;
    --excel-switch-off: #334155;
    --excel-shadow: none;
  }
  body {
    padding: 14px;
    color: var(--excel-text);
    background: var(--excel-bg);
    font-family: var(--vscode-font-family);
    font-size: 12px;
    line-height: 1.55;
  }

  .header {
    padding: 4px 0 14px;
    margin-bottom: 14px;
    border-bottom: 1px solid var(--excel-border);
  }
  .header-title {
    font-size: 22px;
    line-height: 1.15;
    font-weight: 800;
    color: var(--excel-text);
  }
  .version-line,
  .notice-line {
    margin-top: 5px;
    color: var(--excel-muted);
    font-size: 12px;
  }
  .notice-line {
    color: var(--excel-text);
    font-weight: 700;
  }
  .status-row {
    margin-top: 12px;
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 12px;
  }

  .card {
    border: 1px solid var(--excel-border);
    border-radius: 10px;
    padding: 16px;
    margin-bottom: 14px;
    background: var(--excel-surface);
    box-shadow: var(--excel-shadow);
  }

  h2 {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 15px;
    font-weight: 800;
    margin: 0 0 16px 0;
    color: var(--excel-text);
  }
  h2::before {
    content: "";
    width: 4px;
    height: 16px;
    border-radius: 999px;
    background: var(--excel-primary);
  }

  label {
    display: block;
    margin: 14px 0 7px;
    color: var(--excel-muted);
    font-size: 12px;
    font-weight: 700;
  }

  .input-display {
    min-height: 42px;
    padding: 10px 12px;
    border-radius: 8px;
    background: var(--excel-input);
    color: var(--excel-text);
    word-break: break-all;
    font-family: var(--vscode-font-family);
    font-size: 12px;
    border: 1px solid var(--excel-border);
    line-height: 22px;
  }
  .input-display.placeholder {
    color: var(--excel-muted);
  }
  .input-display.path-display {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  button {
    width: 100%;
    min-height: 42px;
    margin-top: 12px;
    padding: 9px 12px;
    border: none;
    border-radius: 8px;
    color: #ffffff;
    background: var(--excel-primary);
    cursor: pointer;
    font-family: var(--vscode-font-family);
    font-size: 12px;
    font-weight: 800;
    transition: background 0.16s ease, box-shadow 0.16s ease, transform 0.16s ease, opacity 0.12s;
  }
  button:hover {
    background: var(--excel-primary-hover);
    box-shadow: inset 0 0 0 999px rgba(255,255,255,0.08), 0 8px 18px rgba(37,99,235,0.22);
    transform: translateY(-1px);
  }
  button:active { opacity: 0.82; transform: translateY(0); }
  button:disabled {
    opacity: 0.4;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
  }
  button.secondary {
    background: var(--excel-secondary);
    color: #ffffff;
  }
  button.secondary:hover {
    background: var(--excel-secondary-hover);
    box-shadow: inset 0 0 0 999px rgba(255,255,255,0.08), 0 8px 18px rgba(71,85,105,0.18);
  }
  button.danger {
    background: #dc2626;
    color: #ffffff;
  }
  button.danger:hover {
    background: #b91d1d;
    box-shadow: inset 0 0 0 999px rgba(255,255,255,0.08), 0 8px 18px rgba(220,38,38,0.22);
  }

  .button-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
  }
  .button-row button {
    min-width: 0;
  }

  .switch-row {
    display: grid;
    grid-template-columns: minmax(0, 1fr) auto;
    align-items: center;
    gap: 16px;
    padding: 14px 0;
    border-bottom: 1px solid var(--excel-border);
  }
  .switch-row:first-of-type {
    padding-top: 0;
  }
  .switch-row:last-of-type {
    border-bottom: 0;
    padding-bottom: 0;
  }
  .switch-row .label-block {
    min-width: 0;
  }
  .title {
    font-size: 14px;
    font-weight: 800;
    color: var(--excel-text);
  }
  .desc {
    margin-top: 6px;
    line-height: 1.65;
    color: var(--excel-muted);
    font-size: 12px;
  }
  .hint {
    margin-top: 12px;
    line-height: 1.6;
    color: var(--excel-muted);
    font-size: 12px;
  }

  .switch {
    position: relative;
    display: inline-block;
    width: 44px;
    height: 24px;
    flex: 0 0 44px;
  }
  .switch input {
    position: absolute;
    opacity: 0;
    width: 0;
    height: 0;
  }
  .slider {
    position: absolute;
    inset: 0;
    background: var(--excel-switch-off);
    border: 1px solid var(--excel-border-strong);
    border-radius: 999px;
    transition: background 0.18s ease, border-color 0.18s ease;
  }
  .slider:before {
    content: "";
    position: absolute;
    width: 18px;
    height: 18px;
    left: 3px;
    top: 50%;
    transform: translateY(-50%);
    background: #ffffff;
    border-radius: 50%;
    box-shadow: 0 2px 6px rgba(15,23,42,0.25);
    transition: transform 0.18s ease;
  }
  .switch input:checked + .slider {
    background: var(--excel-primary);
    border-color: var(--excel-primary);
  }
  .switch input:checked + .slider:before {
    transform: translate(20px, -50%);
  }

  .badge {
    display: inline-flex;
    align-items: center;
    min-height: 24px;
    padding: 2px 10px;
    border-radius: 6px;
    font-size: 11px;
    font-weight: 700;
  }
  .badge-ok { background: rgba(46,160,67,0.15); color: #3fb950; }
  .badge-warn { background: rgba(249,115,22,0.12); color: #c2410c; }
  .badge-err { background: rgba(248,81,73,0.15); color: #f85149; }
  .badge-info { color: var(--excel-muted); }
</style>
</head>
<body>

<div class="header">
  <div class="header-title">VBA助手</div>
  <div class="version-line">版本号 v${this.version}</div>
  <div class="notice-line">仅限内部使用，请勿外传</div>
  <div class="status-row">
    <span>服务状态：</span>
    <span id="serviceStatusBadge" class="badge ${statusBadgeClass}">${statusText}</span>
    <span id="errorBadge" class="badge badge-err" style="${lastError ? '' : 'display:none;'}" title="${(lastError || '').replace(/"/g, '&quot;')}">${lastError ? (lastError.length > 10 ? lastError.slice(0, 10) + "…" : lastError) : ""}</span>
    ${warningCount > 0 ? `<span class="badge badge-warn">警告 ${warningCount}</span>` : ""}
  </div>
</div>

<div class="container">
  <section class="card">
    <h2>路径</h2>

    <label>Excel 文件</label>
    <div class="input-display path-display ${workbookPath ? "" : "placeholder"}" id="workbookPath" title="${this.escapeHtml(workbookPath)}">${this.escapeHtml(workbookPath) || "未选择"}</div>
    <div class="button-row">
      <button id="selectWorkbook" onclick="send('selectWorkbook')">选择文件</button>
      <button class="danger" id="disconnectExcel" onclick="send('disconnectExcel')" ${serviceStatus === "disconnected" || serviceStatus === "unknown" ? "disabled" : ""}>断开Excel</button>
    </div>

    <label>VBA 同步目录</label>
    <div class="input-display path-display ${syncDirectory ? "" : "placeholder"}" id="syncDirectory" title="${this.escapeHtml(syncDirectory)}">${this.escapeHtml(syncDirectory) || "未设置"}</div>

    <div class="button-row">
      <button class="secondary" id="selectDirectory" onclick="send('selectSyncDirectory')">选择目录</button>
      <button class="secondary" id="useExcelSameDirectory" onclick="send('useExcelSameDirectory')">Excel相同目录</button>
    </div>
  </section>

  <section class="card">
    <h2>配置选项</h2>

    <div class="switch-row">
      <div class="label-block">
        <div class="title">自动同步</div>
        <div class="desc">开启后，VBA代码自动写回至Excel</div>
      </div>
      <label class="switch">
        <input type="checkbox" id="autoSync" ${autoSync ? "checked" : ""} onchange="send('toggleAutoSync', {value: this.checked})" />
        <span class="slider"></span>
      </label>
    </div>

    <div class="switch-row">
      <div class="label-block">
        <div class="title">自动执行 VBA</div>
        <div class="desc">开启后，AI可自动执行测试</div>
      </div>
      <label class="switch">
        <input type="checkbox" id="autoRunVba" ${autoRunVba ? "checked" : ""} onchange="send('toggleAutoRun', {value: this.checked})" />
        <span class="slider"></span>
      </label>
    </div>

    <div class="switch-row">
      <div class="label-block">
        <div class="title">Excel 置顶</div>
        <div class="desc">保持Excel窗口始终在最前</div>
      </div>
      <label class="switch">
        <input type="checkbox" id="keepExcelOnTop" ${keepExcelOnTop ? "checked" : ""} onchange="send('toggleKeepExcelOnTop', {value: this.checked})" />
        <span class="slider"></span>
      </label>
    </div>
  </section>

  <section class="card">
    <h2>手工同步</h2>
    <div class="button-row">
      <button id="syncVbeToLocal" onclick="send('syncVbeToLocal')" ${workbookPath ? "" : "disabled"}>同步至目录</button>
      <button id="syncLocalToVbe" onclick="send('syncLocalToVbe')" ${syncDirectory ? "" : "disabled"}>写回Excel</button>
    </div>

    <p class="hint">如遇错误，请手工执行上述功能</p>
  </section>
</div>

<script>
  const vscode = acquireVsCodeApi();
  function send(type, data) {
    const msg = { type, ...(data || {}) };
    vscode.postMessage(msg);
  }
  // 接收插件推送的状态变更
  window.addEventListener('message', (event) => {
    const message = event.data;
    if (message && message.type === 'stateChanged' && message.state) {
      const s = message.state;
      const wb = document.getElementById('workbookPath');
      const sd = document.getElementById('syncDirectory');
      if (wb) {
        wb.textContent = s.workbookPath || '未选择';
        wb.title = s.workbookPath || '';
        wb.classList.toggle('placeholder', !s.workbookPath);
      }
      if (sd) {
        sd.textContent = s.syncDirectory || '未设置';
        sd.title = s.syncDirectory || '';
        sd.classList.toggle('placeholder', !s.syncDirectory);
      }
      const autoRun = document.getElementById('autoRunVba');
      const autoSync = document.getElementById('autoSync');
      const keepOnTop = document.getElementById('keepExcelOnTop');
      if (autoRun) autoRun.checked = !!s.autoRunVba;
      if (autoSync) autoSync.checked = !!s.autoSync;
      if (keepOnTop) keepOnTop.checked = !!s.keepExcelOnTop;

      // 更新服务状态与错误提示
      const statusLabel = {
        unknown: "未知",
        starting: "启动中",
        connected: "已连接（未同步）",
        syncing: "同步中",
        synced: "已同步",
        disconnected: "未连接",
        error: "错误",
      };
      const statusText = statusLabel[s.serviceStatus] || "未知";
      const statusBadgeClass =
        s.serviceStatus === "synced" ? "badge-ok" :
        s.serviceStatus === "connected" ? "badge-warn" :
        s.serviceStatus === "error" ? "badge-err" :
        s.serviceStatus === "disconnected" ? "badge-warn" :
        s.serviceStatus === "syncing" ? "badge-info" : "badge-info";
      const statusBadge = document.getElementById('serviceStatusBadge');
      if (statusBadge) {
        statusBadge.textContent = statusText;
        statusBadge.className = "badge " + statusBadgeClass;
      }
      const errorBadge = document.getElementById('errorBadge');
      if (errorBadge) {
        if (s.lastError) {
          errorBadge.textContent = s.lastError.length > 10 ? s.lastError.slice(0, 10) + "…" : s.lastError;
          errorBadge.title = s.lastError;
          errorBadge.className = "badge badge-err";
          errorBadge.style.display = "";
        } else {
          errorBadge.style.display = "none";
        }
      }
    }
  });
</script>
</body>
</html>`;
  }
}
