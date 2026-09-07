/**
 * MCP Client 封装
 *
 * 基于 @modelcontextprotocol/sdk 的 Client，通过 stdio 与 mcp-server-excel 通信。
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { CallToolResultSchema } from "@modelcontextprotocol/sdk/types.js";
import type { TextContent, CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import type { StdioServerParameters } from "@modelcontextprotocol/sdk/client/stdio.js";

export interface McpToolCallOptions {
  name: string;
  arguments?: Record<string, unknown>;
}

export class McpClientWrapper {
  private client: Client | undefined;
  private transport: StdioClientTransport | undefined;

  constructor(private serverParams: StdioServerParameters) {}

  async initialize(): Promise<void> {
    this.transport = new StdioClientTransport(this.serverParams);

    this.client = new Client(
      { name: "excel-vba-assistant", version: "0.8.6" },
      { capabilities: {} }
    );

    await this.client.connect(this.transport);
  }

  async callTool(options: McpToolCallOptions): Promise<CallToolResult> {
    if (!this.client) {
      throw new Error("MCP client 未初始化");
    }
    const result = await this.client.callTool(
      { name: options.name, arguments: options.arguments ?? {} },
      CallToolResultSchema
    );
    return result as CallToolResult;
  }

  async listTools(): Promise<Array<{ name: string; description?: string }>> {
    if (!this.client) {
      throw new Error("MCP client 未初始化");
    }
    const result = await this.client.listTools();
    return result.tools.map((t) => ({ name: t.name, description: t.description }));
  }

  getTransport(): StdioClientTransport | undefined {
    return this.transport;
  }

  async close(): Promise<void> {
    if (this.client) {
      await this.client.close();
      this.client = undefined;
    }
    if (this.transport) {
      await this.transport.close();
      this.transport = undefined;
    }
  }
}

/** 从 CallToolResult 中提取文本内容 */
export function extractText(result: CallToolResult): string {
  if (!result.content || result.content.length === 0) return "";
  return result.content
    .filter((c): c is TextContent => c.type === "text")
    .map((c) => c.text)
    .join("\n");
}

/** 尝试将 CallToolResult 解析为 JSON */
export function extractJson<T = unknown>(result: CallToolResult): T | undefined {
  const text = extractText(result);
  if (!text) return undefined;
  try {
    return JSON.parse(text) as T;
  } catch {
    return undefined;
  }
}
