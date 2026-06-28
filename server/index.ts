/**
 * Excel VBA Assistant - MCP Server 入口
 *
 * 通过 stdio 与 MCP 客户端（如 Trae / Claude Desktop）通信。
 *
 * 环境变量：
 *   VBE_FILE_PATH   默认 Excel 文件路径（启动时自动注册为默认工作簿）
 *   VBE_LOCAL_DIR   默认本地同步目录
 *   VBE_READ_ONLY   true 时启用只读模式，禁止所有危险操作（修改/创建/删除/运行宏/本地→VBE 同步）
 *
 * 用法（在 Trae 的 mcp 配置中）：
 *   {
 *     "mcpServers": {
 *       "excel-vba-assistant": {
 *         "command": "node",
 *         "args": ["<插件目录>/dist/mcp-server.js"],
 *         "env": {
 *           "VBE_FILE_PATH": "D:\\work\\my_addin.xlam",
 *           "VBE_LOCAL_DIR": "D:\\work\\vba_sync",
 *           "VBE_READ_ONLY": "false"
 *         }
 *       }
 *     }
 *   }
 *
 * 对齐任务文档第二十节的 14 个 excel_* 工具。
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";

import { ExcelVbaServiceClient } from "./client/ExcelVbaServiceClient";
import { buildToolDefinitions, dispatchToolCall } from "./tools";

const client = new ExcelVbaServiceClient();
const readOnly = client.isReadOnly();
const defaultId = client.getDefaultWorkbookId();
const defaultDir = client.getDefaultLocalDir();

const server = new Server(
  { name: "excel-vba-assistant", version: "0.1.0" },
  { capabilities: { tools: {} } }
);

// 列出所有可用工具
server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: buildToolDefinitions(),
}));

// 分发工具调用
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const result = await dispatchToolCall(name, args || {}, { client });
  // SDK 1.29.0 的 CallToolResult 期望 content 元素携带可选的 annotations/_meta，
  // 这里做一次结构对齐，避免 TS 联合类型校验失败。
  return {
    content: result.content.map((item) => ({ type: "text" as const, text: item.text })),
    isError: result.isError,
  } as any;
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  const banner: string[] = [
    `Excel VBA Assistant MCP Server v0.1.0 已启动`,
  ];
  if (readOnly) {
    banner.push(`[只读模式] VBE_READ_ONLY=true，所有危险操作将被拒绝`);
  }
  if (defaultId) {
    banner.push(`默认工作簿: ${defaultId}`);
  } else {
    banner.push(`未配置默认工作簿（VBE_FILE_PATH 未设置），请通过 excel_open_workbook 注册`);
  }
  if (defaultDir) {
    banner.push(`默认本地同步目录: ${defaultDir}`);
  } else {
    banner.push(`未配置默认本地同步目录（VBE_LOCAL_DIR 未设置），同步工具需在参数中显式指定 localDir`);
  }
  // 启动信息输出到 stderr，避免污染 stdout 的 MCP 协议数据
  console.error(banner.join("\n"));
}

main().catch((error) => {
  console.error("Excel VBA Assistant MCP Server 启动失败:", error);
  process.exit(1);
});
