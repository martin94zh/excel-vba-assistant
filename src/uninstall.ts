import * as fs from "fs";
import * as path from "path";

const MCP_SERVER_NAME = "excel-mcp";
const TRACK_FILE = "mcp-workspaces.json";
const APP_DIR = "excel-vba-assistant";

function getTrackFilePath(): string {
  const base = process.env.APPDATA || process.env.USERPROFILE || process.env.HOME;
  if (!base) return "";
  return path.join(base, APP_DIR, TRACK_FILE);
}

function removeMcpServerFromWorkspace(wsPath: string): void {
  const mcpPath = path.join(wsPath, ".trae", "mcp.json");
  if (!fs.existsSync(mcpPath)) return;
  try {
    const content = fs.readFileSync(mcpPath, "utf-8");
    const json = JSON.parse(content);
    if (json.mcpServers && json.mcpServers[MCP_SERVER_NAME]) {
      delete json.mcpServers[MCP_SERVER_NAME];
      if (Object.keys(json.mcpServers).length === 0) {
        delete json.mcpServers;
      }
      fs.writeFileSync(mcpPath, JSON.stringify(json, null, 2), "utf-8");
      console.log(`Removed MCP server config from ${mcpPath}`);
    }
  } catch (err) {
    console.error(`Failed to clean ${mcpPath}:`, err);
  }
}

function main(): void {
  const trackFile = getTrackFilePath();
  if (!trackFile || !fs.existsSync(trackFile)) {
    console.log("No tracked workspaces to clean.");
    return;
  }
  try {
    const workspaces: string[] = JSON.parse(fs.readFileSync(trackFile, "utf-8"));
    for (const ws of workspaces) {
      if (fs.existsSync(ws)) {
        removeMcpServerFromWorkspace(ws);
      }
    }
    fs.unlinkSync(trackFile);
  } catch (err) {
    console.error("Failed to clean MCP configs:", err);
  }
}

main();
