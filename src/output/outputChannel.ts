/**
 * Excel VBA Assistant - 输出日志通道
 */
import * as vscode from "vscode";

export class OutputManager {
  private channel: vscode.OutputChannel;

  constructor() {
    this.channel = vscode.window.createOutputChannel("Excel VBA Assistant");
  }

  log(message: string): void {
    const time = new Date().toLocaleTimeString("zh-CN", { hour12: false });
    this.channel.appendLine(`[${time}] ${message}`);
  }

  info(message: string): void {
    this.log(`[INFO] ${message}`);
  }

  warn(message: string): void {
    this.log(`[WARN] ${message}`);
  }

  error(message: string): void {
    this.log(`[ERROR] ${message}`);
  }

  show(preserveFocus = true): void {
    this.channel.show(preserveFocus);
  }

  dispose(): void {
    this.channel.dispose();
  }
}
