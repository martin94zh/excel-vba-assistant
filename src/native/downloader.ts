/**
 * mcp-server-excel 下载器
 *
 * 从 GitHub Releases 下载最新版 ExcelMcp-MCP-Server-windows.zip 并解压。
 * 构建时和激活时均可调用。
 */
import { createWriteStream, existsSync, mkdirSync } from "fs";
import { writeFile, unlink, readdir } from "fs/promises";
import { get as httpsGet } from "https";
import { join } from "path";
import { createUnzip } from "zlib";
import { pipeline } from "stream/promises";

const GITHUB_API = "api.github.com";
const REPO_PATH = "/repos/sbroenne/mcp-server-excel/releases/latest";
const MCP_SERVER_ASSET_PREFIX = "ExcelMcp-MCP-Server";
const MCP_SERVER_EXE_CANDIDATES = ["mcp-excel.exe", "Sbroenne.ExcelMcp.McpServer.exe"];

export interface DownloadResult {
  success: boolean;
  exePath: string;
  version: string;
  message: string;
}

/** 检查本地是否已有 mcp-server-excel exe */
export function findLocalMcpServer(distDir: string): string | undefined {
  for (const name of MCP_SERVER_EXE_CANDIDATES) {
    const candidate = join(distDir, name);
    if (existsSync(candidate)) return candidate;
  }
  // 也尝试查找任何疑似 exe
  try {
    const files = readdir(distDir);
    for (const f of files as unknown as string[]) {
      const lower = f.toLowerCase();
      if (lower.endsWith(".exe") && (lower.includes("excelmcp") || lower === "mcp-excel.exe")) {
        return join(distDir, f);
      }
    }
  } catch { /* ignore */ }
  return undefined;
}

/** 从 GitHub Releases 下载最新 MCP Server */
export async function downloadLatestMcpServer(distDir: string): Promise<DownloadResult> {
  mkdirSync(distDir, { recursive: true });

  // 1. 获取最新 release 信息
  const releaseInfo = await fetchJson<{ tag_name: string; assets: Array<{ name: string; browser_download_url: string }> }>(
    `https://${GITHUB_API}${REPO_PATH}`
  );
  const version = releaseInfo.tag_name;
  const asset = releaseInfo.assets.find((a) =>
    a.name.startsWith(MCP_SERVER_ASSET_PREFIX) && a.name.endsWith(".zip")
  );
  if (!asset) {
    return { success: false, exePath: "", version, message: `Release ${version} 中未找到 MCP Server zip 资产` };
  }

  const zipPath = join(distDir, asset.name);

  // 2. 下载 zip
  console.log(`[downloader] 下载 ${asset.name} ...`);
  await downloadFile(asset.browser_download_url, zipPath);

  // 3. 解压（使用 PowerShell Expand-Archive，跨 Windows 版本最可靠）
  console.log(`[downloader] 解压到 ${distDir} ...`);
  const { exec } = await import("child_process");
  const { promisify } = await import("util");
  const execAsync = promisify(exec);
  await execAsync(`powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${distDir}' -Force"`, { timeout: 60000 });

  // 4. 删除 zip
  await unlink(zipPath).catch(() => {});

  // 5. 确认 exe 存在
  const exePath = findLocalMcpServer(distDir);
  if (!exePath) {
    return { success: false, exePath: "", version, message: `解压后未找到 mcp-server-excel 可执行文件` };
  }

  return { success: true, exePath, version, message: `已下载 mcp-server-excel ${version}` };
}

function fetchJson<T>(url: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const req = httpsGet(url, { headers: { "User-Agent": "excel-vba-assistant" } }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        const loc = res.headers.location;
        if (loc) {
          fetchJson<T>(loc).then(resolve).catch(reject);
          return;
        }
      }
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          resolve(JSON.parse(data) as T);
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

function downloadFile(url: string, dest: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const file = createWriteStream(dest);
    const req = httpsGet(url, { headers: { "User-Agent": "excel-vba-assistant" } }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        const loc = res.headers.location;
        if (loc) {
          downloadFile(loc, dest).then(resolve).catch(reject);
          return;
        }
      }
      if (res.statusCode !== 200) {
        reject(new Error(`Download failed: ${res.statusCode}`));
        return;
      }
      res.pipe(file);
      file.on("finish", () => {
        file.close();
        resolve();
      });
    });
    req.on("error", (err) => {
      file.close();
      reject(err);
    });
    req.setTimeout(120000, () => {
      req.destroy();
      file.close();
      reject(new Error("Download timeout"));
    });
  });
}
