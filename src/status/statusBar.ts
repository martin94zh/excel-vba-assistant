/**
 * Excel VBA Assistant - 状态栏管理
 *
 * 状态栏文本和颜色直接与服务状态（ServiceStatus）对齐：
 *   未知 / 启动中 / 已连接 / 同步中 / 已同步 / 未连接 / 错误
 *
 * synced 会在 5 秒后自动回落到 connected，避免长期占用强调色。
 */
import * as vscode from "vscode";
import type { ServiceStatus, StateManager } from "../state";

export class StatusBarManager {
  private item: vscode.StatusBarItem;

  constructor(private stateManager: StateManager) {
    this.item = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 1000);
    this.item.command = "excelVba.openOutput";
    this.stateManager.onChange(() => this.handleStateChange());
    this.handleStateChange();
    this.item.show();
  }

  private handleStateChange(): void {
    const status = this.stateManager.getAll().serviceStatus;
    this.apply(status);
  }

  private apply(status: ServiceStatus): void {
    switch (status) {
      case "unknown":
        this.item.text = "$(circle-outline) VBE 同步";
        this.item.backgroundColor = undefined;
        break;
      case "starting":
        this.item.text = "$(sync) VBE 同步";
        this.item.backgroundColor = undefined;
        break;
      case "connected":
        this.item.text = "$(circle-filled) VBE 同步";
        this.item.backgroundColor = new vscode.ThemeColor("statusBarItem.warningBackground");
        break;
      case "syncing":
        this.item.text = "$(sync~spin) VBE 同步";
        this.item.backgroundColor = undefined;
        break;
      case "synced":
        this.item.text = "$(check) VBE 同步";
        this.item.backgroundColor = new vscode.ThemeColor("excelVba.statusBarSyncedBackground");
        break;
      case "disconnected":
        this.item.text = "$(circle-outline) VBE 同步";
        this.item.backgroundColor = new vscode.ThemeColor("statusBarItem.warningBackground");
        break;
      case "error":
        this.item.text = "$(error) VBE 同步";
        this.item.backgroundColor = new vscode.ThemeColor("statusBarItem.errorBackground");
        break;
    }
    this.item.tooltip = this.buildTooltip(status);
  }

  private buildTooltip(status: ServiceStatus): vscode.MarkdownString {
    const state = this.stateManager.getAll();
    const lines: string[] = [];
    lines.push(`**Excel VBA Assistant**`);
    lines.push("");
    lines.push(`- Excel 文件：${state.workbookPath || "未选择"}`);
    lines.push(`- 同步目录：${state.syncDirectory || "未设置"}`);
    lines.push(`- 服务状态：${this.statusLabel(status)}`);
    lines.push(`- 上次同步方向：${state.lastSyncDirection === "vbe-to-local" ? "VBE → 本地" : state.lastSyncDirection === "local-to-vbe" ? "本地 → VBE" : "无"}`);
    lines.push(`- 上次同步时间：${this.formatSyncTime(state.lastSyncAt)}`);
    lines.push(`- 最近错误：${state.lastError || "无"}`);
    lines.push(`- 警告数：${state.warningCount}`);
    const md = new vscode.MarkdownString(lines.join("\n"));
    md.supportThemeIcons = true;
    return md;
  }

  private statusLabel(status: ServiceStatus): string {
    switch (status) {
      case "unknown": return "未知";
      case "starting": return "启动中";
      case "connected": return "已连接（未同步）";
      case "syncing": return "同步中";
      case "synced": return "已同步";
      case "disconnected": return "未连接";
      case "error": return "错误";
      default: return "未知";
    }
  }

  private formatSyncTime(iso?: string): string {
    if (!iso) return "无";
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hour = String(d.getHours()).padStart(2, "0");
    const minute = String(d.getMinutes()).padStart(2, "0");
    const second = String(d.getSeconds()).padStart(2, "0");
    return `${year}年${month}月${day}日 ${hour}:${minute}:${second}`;
  }

  dispose(): void {
    this.item.dispose();
  }
}
