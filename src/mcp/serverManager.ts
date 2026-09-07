/**
 * mcp-server-excel 进程管理器
 *
 * 负责启动、监控、重启和停止 mcp-server-excel.exe 子进程。
 */
import { existsSync } from "fs";
import { join } from "path";
import { EventEmitter } from "events";
import { McpClientWrapper } from "./mcpClient";
import { downloadLatestMcpServer, findLocalMcpServer } from "../native/downloader";

const HEALTH_CHECK_INTERVAL_MS = 30000;
const RESTART_DELAY_MS = 3000;

export class McpServerManager {
  private client: McpClientWrapper | undefined;
  private healthTimer: NodeJS.Timeout | undefined;
  private isShuttingDown = false;
  private _onCrash = new EventEmitter();

  constructor(private extensionPath: string) {}

  /** 确保 mcp-server-excel 已启动并连接 */
  async start(): Promise<McpClientWrapper> {
    if (this.client) {
      const healthy = await this.healthCheck();
      if (healthy) return this.client;
      await this.stop();
    }

    const distDir = join(this.extensionPath, "dist");
    let exePath = findLocalMcpServer(distDir);

    if (!exePath) {
      console.log("[McpServerManager] 本地未找到 mcp-server-excel，尝试下载...");
      const result = await downloadLatestMcpServer(distDir);
      if (!result.success) {
        throw new Error(`下载 mcp-server-excel 失败: ${result.message}`);
      }
      exePath = result.exePath;
    }

    if (!existsSync(exePath)) {
      throw new Error(`mcp-server-excel 可执行文件不存在: ${exePath}`);
    }

    console.log(`[McpServerManager] 启动 ${exePath}`);
    this.client = new McpClientWrapper({ command: exePath, args: [], stderr: "pipe" });
    await this.client.initialize();
    console.log("[McpServerManager] MCP client 初始化成功");

    const transport = this.client.getTransport();
    if (transport) {
      transport.stderr?.on("data", (data: Buffer) => {
        const msg = data.toString("utf-8").trim();
        if (msg) console.error(`[mcp-server-excel stderr] ${msg}`);
      });
      transport.onclose = () => {
        console.log("[McpServerManager] mcp-server-excel 退出");
        if (!this.isShuttingDown) {
          this._onCrash.emit("crash");
        }
      };
    }

    this.startHealthCheck();
    return this.client;
  }

  async stop(): Promise<void> {
    this.isShuttingDown = true;
    this.stopHealthCheck();

    if (this.client) {
      try {
        await this.client.close();
      } catch { /* ignore */ }
      this.client = undefined;
    }

    this.isShuttingDown = false;
  }

  getClient(): McpClientWrapper | undefined {
    return this.client;
  }

  private async healthCheck(): Promise<boolean> {
    if (!this.client) return false;
    try {
      const tools = await this.client.listTools();
      return tools.length > 0;
    } catch {
      return false;
    }
  }

  private startHealthCheck(): void {
    this.stopHealthCheck();
    this.healthTimer = setInterval(async () => {
      const ok = await this.healthCheck();
      if (!ok && !this.isShuttingDown) {
        console.warn("[McpServerManager] 健康检查失败，准备重启");
        await this.stop();
        setTimeout(() => this.start().catch((e) => console.error("[McpServerManager] 自动重启失败:", e)), RESTART_DELAY_MS);
      }
    }, HEALTH_CHECK_INTERVAL_MS);
  }

  private stopHealthCheck(): void {
    if (this.healthTimer) {
      clearInterval(this.healthTimer);
      this.healthTimer = undefined;
    }
  }
}
