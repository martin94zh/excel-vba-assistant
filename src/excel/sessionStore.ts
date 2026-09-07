/**
 * Session 存储管理
 *
 * mcp-server-excel 使用 session-based 架构：
 * 1. file(action: 'open', path) -> 返回 session_id
 * 2. 后续调用携带 session_id
 * 3. file(action: 'close', session_id, save=true) -> 保存并释放
 */
import { McpClientWrapper, extractJson } from "../mcp/mcpClient";

interface SessionEntry {
  sessionId: string;
  workbookPath: string;
}

export class SessionStore {
  private sessions = new Map<string, SessionEntry>();

  constructor(private client: McpClientWrapper) {}

  async ensureSession(workbookPath: string, options?: { show?: boolean }): Promise<string> {
    const existing = this.sessions.get(workbookPath);
    if (existing) {
      const valid = await this.checkSession(existing.sessionId);
      if (valid) return existing.sessionId;
      this.sessions.delete(workbookPath);
    }

    const result = await this.client.callTool({
      name: "file",
      arguments: { action: "open", path: workbookPath, show: options?.show ?? false },
    });

    const data = extractJson<{ session_id?: string; error?: string }>(result);
    if (!data?.session_id) {
      const text = extractJson<{ message?: string }>(result);
      throw new Error(text?.message || data?.error || "打开工作簿失败，无法获取 session_id");
    }

    this.sessions.set(workbookPath, { sessionId: data.session_id, workbookPath });
    return data.session_id;
  }

  async closeSession(workbookPath: string, save = true): Promise<void> {
    const entry = this.sessions.get(workbookPath);
    if (!entry) return;

    try {
      await this.client.callTool({
        name: "file",
        arguments: { action: "close", session_id: entry.sessionId, save },
      });
    } catch (err) {
      console.warn(`[SessionStore] 关闭 session 失败: ${err}`);
    }
    this.sessions.delete(workbookPath);
  }

  async closeAllSessions(save = true): Promise<void> {
    const paths = Array.from(this.sessions.keys());
    for (const path of paths) {
      await this.closeSession(path, save);
    }
  }

  private async checkSession(sessionId: string): Promise<boolean> {
    try {
      const result = await this.client.callTool({
        name: "file",
        arguments: { action: "list" },
      });
      const data = extractJson<{ sessions?: Array<{ session_id?: string }> }>(result);
      return data?.sessions?.some((s) => s.session_id === sessionId) ?? false;
    } catch {
      return false;
    }
  }
}
