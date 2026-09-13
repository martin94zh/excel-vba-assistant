/**
 * 下载 ExcelMcp CLI（excelcli.exe）到 dist/
 *
 * 从 GitHub Releases 获取最新 ExcelMcp-CLI-*-windows.zip 并解压。
 * 说明：release 资产直连可能超时，脚本会自动尝试 HTTPS_PROXY/HTTP_PROXY 代理。
 */
const { execFileSync, execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const GITHUB_API = "api.github.com";
const REPO_PATH = "/repos/sbroenne/mcp-server-excel/releases/latest";
const CLI_ASSET_PREFIX = "ExcelMcp-CLI";

function fetchJson(url) {
  const result = execFileSync("powershell", [
    "-NoProfile",
    "-Command",
    `
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$proxy = $env:HTTPS_PROXY; if (-not $proxy) { $proxy = $env:HTTP_PROXY }
$props = @{ Uri = '${url}'; UseBasicParsing = $true; TimeoutSec = 60; Headers = @{ 'User-Agent' = 'excel-vba-assistant' } }
if ($proxy) { $props['Proxy'] = $proxy }
(Invoke-WebRequest @props).Content
`,
  ], { encoding: "utf-8", timeout: 90000, stdio: ["ignore", "pipe", "inherit"] });
  return JSON.parse(result);
}

function downloadFile(url, dest) {
  console.log("[download-excel-cli] 通过 PowerShell 下载 ...");
  execFileSync(
    "powershell",
    [
      "-NoProfile",
      "-Command",
      `
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$proxy = $env:HTTPS_PROXY; if (-not $proxy) { $proxy = $env:HTTP_PROXY }
$props = @{ Uri = '${url.replace(/'/g, "''")}'; OutFile = '${dest.replace(/'/g, "''")}'; TimeoutSec = 580; UseBasicParsing = $true }
if ($proxy) { $props['Proxy'] = $proxy }
Invoke-WebRequest @props
`,
    ],
    { stdio: "inherit", timeout: 600000 }
  );
}

async function main() {
  const distDir = path.join(process.cwd(), "dist");
  fs.mkdirSync(distDir, { recursive: true });

  const exePath = path.join(distDir, "excelcli.exe");
  if (fs.existsSync(exePath)) {
    console.log(`[download-excel-cli] 本地已存在 ${exePath}，跳过下载`);
    return;
  }

  console.log("[download-excel-cli] 查询 GitHub Releases ...");
  const releaseInfo = fetchJson(`https://${GITHUB_API}${REPO_PATH}`);
  const version = releaseInfo.tag_name;
  const asset = (releaseInfo.assets || []).find(
    (a) => a.name.startsWith(CLI_ASSET_PREFIX) && a.name.endsWith(".zip")
  );
  if (!asset) {
    throw new Error(`Release ${version} 中未找到 CLI zip 资产`);
  }

  const zipPath = path.join(distDir, asset.name);
  console.log(`[download-excel-cli] 下载 ${asset.name} ...`);
  downloadFile(asset.browser_download_url, zipPath);

  console.log(`[download-excel-cli] 解压到 ${distDir} ...`);
  execSync(
    `powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${distDir}' -Force"`,
    { timeout: 120000, stdio: "inherit" }
  );

  fs.unlinkSync(zipPath);

  if (!fs.existsSync(exePath)) {
    throw new Error("解压后未找到 excelcli.exe");
  }
  console.log(`[download-excel-cli] 已下载 ExcelMcp CLI ${version}: ${exePath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
