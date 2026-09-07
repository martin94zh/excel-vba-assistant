/**
 * Build-time helper: download mcp-server-excel from GitHub Releases into dist/.
 * This ensures the .exe is embedded in the .vsix so users don't need to download at runtime.
 */
const https = require("https");
const fs = require("fs");
const path = require("path");
const { execSync, execFileSync } = require("child_process");

const GITHUB_API = "api.github.com";
const REPO_PATH = "/repos/sbroenne/mcp-server-excel/releases/latest";
const MCP_SERVER_ASSET_PREFIX = "ExcelMcp-MCP-Server";
const MCP_SERVER_EXE_CANDIDATES = ["mcp-excel.exe", "Sbroenne.ExcelMcp.McpServer.exe"];

function findLocalMcpServer(distDir) {
  for (const name of MCP_SERVER_EXE_CANDIDATES) {
    const candidate = path.join(distDir, name);
    if (fs.existsSync(candidate)) return candidate;
  }
  try {
    for (const f of fs.readdirSync(distDir)) {
      const lower = f.toLowerCase();
      if (lower.endsWith(".exe") && (lower.includes("excelmcp") || lower === "mcp-excel.exe")) {
        return path.join(distDir, f);
      }
    }
  } catch { /* ignore */ }
  return undefined;
}

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers: { "User-Agent": "excel-vba-assistant" } }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        const loc = res.headers.location;
        if (loc) return fetchJson(loc).then(resolve).catch(reject);
      }
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error(`JSON parse error: ${e}`));
        }
      });
    });
    req.on("error", reject);
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error("Request timeout"));
    });
  });
}

function downloadFile(url, dest) {
  // Use PowerShell Invoke-WebRequest for reliable TLS/redirect handling on Windows.
  console.log(`[download-mcp-server] 通过 PowerShell 下载 ...`);
  execFileSync(
    "powershell",
    [
      "-NoProfile",
      "-Command",
      `Invoke-WebRequest -Uri '${url.replace(/'/g, "''")}' -OutFile '${dest.replace(/'/g, "''")}' -TimeoutSec 300 -UseBasicParsing`,
    ],
    { stdio: "inherit", timeout: 300000 }
  );
}

async function main() {
  const distDir = path.join(process.cwd(), "dist");
  fs.mkdirSync(distDir, { recursive: true });

  const existing = findLocalMcpServer(distDir);
  if (existing) {
    console.log(`[download-mcp-server] 本地已存在 ${existing}，跳过下载`);
    return;
  }

  console.log("[download-mcp-server] 查询 GitHub Releases ...");
  const releaseInfo = await fetchJson(`https://${GITHUB_API}${REPO_PATH}`);
  const version = releaseInfo.tag_name;
  const asset = releaseInfo.assets.find((a) =>
    a.name.startsWith(MCP_SERVER_ASSET_PREFIX) && a.name.endsWith(".zip")
  );
  if (!asset) {
    throw new Error(`Release ${version} 中未找到 MCP Server zip 资产`);
  }

  const zipPath = path.join(distDir, asset.name);
  console.log(`[download-mcp-server] 下载 ${asset.name} ...`);
  await downloadFile(asset.browser_download_url, zipPath);

  console.log(`[download-mcp-server] 解压到 ${distDir} ...`);
  execSync(
    `powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${distDir}' -Force"`,
    { timeout: 60000, stdio: "inherit" }
  );

  fs.unlinkSync(zipPath);

  const exePath = findLocalMcpServer(distDir);
  if (!exePath) {
    throw new Error(`解压后未找到 ${MCP_SERVER_EXE_NAME}`);
  }
  console.log(`[download-mcp-server] 已下载 mcp-server-excel ${version}: ${exePath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
