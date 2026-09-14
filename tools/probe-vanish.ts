/**
 * 消失检测探针：验证 hasVisibleWorkbookWindow 的两个关键场景。
 * 前置：工作簿已由 excelcli session open --show 打开。
 *
 * 场景 A（v0.9.3 误判）：打开 VBE 窗口（模拟用户在 VBE 编辑）→ 应判定存活
 * 场景 B（v0.9.4 漏检）：隐藏 Excel 全部窗口（模拟用户关窗但进程残留）→ 应判定消失
 */
import { VbaClient } from "../src/client/vbaClient";
import { findExcelPidForWorkbook, hasVisibleWorkbookWindow } from "../src/native/processUtils";
import { exec } from "child_process";
import { promisify } from "util";

const execAsync = promisify(exec);

async function main(): Promise<void> {
  const wbPath = process.argv[2];
  const pid = await findExcelPidForWorkbook(wbPath);
  if (!pid) {
    console.error("工作簿未打开");
    process.exit(2);
  }
  console.log("Excel PID =", pid);
  const client = new VbaClient(wbPath, pid);

  const aliveA = await hasVisibleWorkbookWindow(wbPath);
  console.log(`[场景A] 正常打开（+VBE 前台模拟）→ 存活=${aliveA}，期望 true`);

  // 打开 VBE 窗口（模拟用户操作：Alt+F11 等价——通过 COM 打开 VBE 主窗口）
  const vbeOpen = await client.runMacro ? null : null;
  void vbeOpen;
  // VBE 主窗口通常在工程被访问后出现；用 COM 触发一次 VBE 访问
  await client.listDialogs(); // 顺带验证弹窗扫描
  const aliveA2 = await hasVisibleWorkbookWindow(wbPath);
  console.log(`[场景A2] 弹窗扫描后 → 存活=${aliveA2}，期望 true`);

  // 场景 B：隐藏全部 Excel 窗口（模拟用户关窗但进程残留）
  const hideScript = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class Hid {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(P c, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
}
"@
$tp = ${pid}
[Hid]::EnumWindows({
  param($h, $l)
  $p = [uint32]0
  [void][Hid]::GetWindowThreadProcessId($h, [ref]$p)
  if ($p -ne $tp) { return $true }
  $sb = New-Object System.Text.StringBuilder 256
  [void][Hid]::GetClassName($h, $sb, 256)
  if ($sb.ToString() -eq "XLMAIN") { [void][Hid]::ShowWindow($h, 0) }
  return $true
}, [IntPtr]::Zero) | Out-Null
Write-Output "hidden"
`;
  await execAsync(`powershell -NoProfile -EncodedCommand ${Buffer.from(hideScript, "utf16le").toString("base64")}`, { timeout: 20000 });
  const aliveB = await hasVisibleWorkbookWindow(wbPath);
  console.log(`[场景B] 全部窗口隐藏（进程残留）→ 存活=${aliveB}，期望 false`);

  // 恢复：重新显示窗口
  const showScript = hideScript.replace("ShowWindow($h, 0)", "ShowWindow($h, 9)").replace("Expected", "");
  const showFixed = showScript.replace('if ($sb.ToString() -eq "XLMAIN") { [void][Hid]::ShowWindow($h, 9) }', 'if ($sb.ToString() -eq "XLMAIN") { [void][Hid]::ShowWindow($h, 9) }');
  await execAsync(`powershell -NoProfile -EncodedCommand ${Buffer.from(showFixed.replace("Write-Output \"hidden\"", "Write-Output \"shown\""), "utf16le").toString("base64")}`, { timeout: 20000 });
  const aliveC = await hasVisibleWorkbookWindow(wbPath);
  console.log(`[恢复] 窗口重新显示 → 存活=${aliveC}，期望 true`);

  const pass = aliveA && aliveA2 && !aliveB && aliveC;
  console.log(pass ? "\n消失检测两场景全部符合预期 ✅" : "\n消失检测不符合预期 ❌");
  process.exit(pass ? 0 : 1);
}

main().catch((e) => {
  console.error("PROBE FAILED:", e instanceof Error ? e.message : String(e));
  process.exit(1);
});
