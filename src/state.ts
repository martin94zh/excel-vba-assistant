/**
 * Excel VBA Assistant - 状态管理
 *
 * 所有用户配置通过 globalState 存储，不走 JSON 文件，
 * 与任务文档中的 ExcelVbaPluginState 模型对齐。
 */
import * as vscode from "vscode";

export type ServiceStatus = "unknown" | "starting" | "connected" | "syncing" | "synced" | "disconnected" | "error";
export type SyncDirection = "vbe-to-local" | "local-to-vbe";

export interface ExcelVbaPluginState {
  workbookPath: string;
  syncDirectory: string;
  syncDirectoryAutoCreated: boolean;
  previousSyncDirectory: string;
  autoRunVba: boolean;
  autoRunMacroName: string;
  autoSync: boolean;
  keepExcelOnTop: boolean;
  excelProcessId: number;
  serviceStatus: ServiceStatus;
  lastSyncDirection?: SyncDirection;
  lastSyncAt?: string;
  errorCount: number;
  warningCount: number;
  isSyncing: boolean;
}

const DEFAULTS: ExcelVbaPluginState = {
  workbookPath: "",
  syncDirectory: "",
  syncDirectoryAutoCreated: false,
  previousSyncDirectory: "",
  autoRunVba: false,
  autoRunMacroName: "",
  autoSync: false,
  keepExcelOnTop: false,
  excelProcessId: 0,
  serviceStatus: "unknown",
  errorCount: 0,
  warningCount: 0,
  isSyncing: false,
};

const PREFIX = "excelVba-";

export class StateManager {
  private context: vscode.ExtensionContext;
  private _onChange = new vscode.EventEmitter<Partial<ExcelVbaPluginState>>();
  readonly onChange = this._onChange.event;

  constructor(context: vscode.ExtensionContext) {
    this.context = context;
  }

  get<K extends keyof ExcelVbaPluginState>(key: K): ExcelVbaPluginState[K] {
    return this.context.globalState.get<ExcelVbaPluginState[K]>(`${PREFIX}${key}`, DEFAULTS[key]);
  }

  async set<K extends keyof ExcelVbaPluginState>(key: K, value: ExcelVbaPluginState[K]): Promise<void> {
    await this.context.globalState.update(`${PREFIX}${key}`, value);
    this._onChange.fire({ [key]: value } as Partial<ExcelVbaPluginState>);
  }

  /** 运行期状态（不持久化） */
  private runtime: Partial<ExcelVbaPluginState> = {};

  setRuntime(patch: Partial<ExcelVbaPluginState>): void {
    this.runtime = { ...this.runtime, ...patch };
    this._onChange.fire(patch);
  }

  getRuntime<K extends keyof ExcelVbaPluginState>(key: K): ExcelVbaPluginState[K] | undefined {
    return this.runtime[key];
  }

  getAll(): ExcelVbaPluginState {
    return {
      ...DEFAULTS,
      workbookPath: this.get("workbookPath"),
      syncDirectory: this.get("syncDirectory"),
      syncDirectoryAutoCreated: this.get("syncDirectoryAutoCreated"),
      previousSyncDirectory: this.get("previousSyncDirectory"),
      autoRunVba: this.get("autoRunVba"),
      autoRunMacroName: this.get("autoRunMacroName"),
      autoSync: this.get("autoSync"),
      keepExcelOnTop: this.get("keepExcelOnTop"),
      excelProcessId: this.get("excelProcessId"),
      serviceStatus: this.getRuntime("serviceStatus") ?? DEFAULTS.serviceStatus,
      lastSyncDirection: this.getRuntime("lastSyncDirection"),
      lastSyncAt: this.getRuntime("lastSyncAt"),
      errorCount: this.getRuntime("errorCount") ?? DEFAULTS.errorCount,
      warningCount: this.getRuntime("warningCount") ?? DEFAULTS.warningCount,
      isSyncing: this.getRuntime("isSyncing") ?? DEFAULTS.isSyncing,
    };
  }
}
