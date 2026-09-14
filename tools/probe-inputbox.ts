/**
 * InputBox 填入链路探针：runMacro(带 InputBox 的宏) → 3 秒后 listDialogs 查看 hasEdit
 * → 直接 fillDialogInput(handle, "42", true) → 读 B4 验证。
 * 前置：工作簿已由 excelcli session open --show 打开，且模块含 E2E_InputBox。
 */
import { VbaClient } from "../src/client/vbaClient";
import { findExcelPidForWorkbook } from "../src/native/processUtils";
import { execSync } from "child_process";
import { writeFileSync } from "fs";
import { join } from "path";

async function main(): Promise<void> {
  const wbPath = process.argv[2];
  const sid = process.env.E2E_SID || "";
  const exeAbs = require("path").resolve("dist", "excelcli.exe");

  // 写入 InputBox 宏
  const code = [
    'Public Function E2E_Base() As String',
    '    E2E_Base = "base"',
    'End Function',
    '',
    'Sub E2E_InputBox()',
    '    Dim v As String',
    '    v = InputBox("请输入数量")',
    '    Range("B4").Value = "N=" & v',
    'End Sub',
  ].join("\r\n") + "\r\n";
  const p = join(process.cwd(), "build", "mod_inputbox.bas");
  writeFileSync(p, code, "utf-8");
  try {
    execSync(`"${exeAbs}" -q vba update --session ${sid} --module-name Module1 --vba-code-file "${p.replace(/\//g, "\\")}"`, { stdio: "pipe", timeout: 30000 });
  } catch (e: any) {
    console.log("update fail:", String(e.message).slice(0, 150));
  }

  const pid = await findExcelPidForWorkbook(wbPath);
  console.log("PID =", pid);
  const client = new VbaClient(wbPath, pid || undefined);

  // 直接用 COM 起宏（不经过 runMacro 的弹窗循环，避免干扰），3 秒后探测
  const { spawn } = await import("child_process");
  void spawn;

  // 用 runMacro 但 policy=errors-only（不自动处理），让 InputBox 弹着，手工探测
  const runPromise = client.runMacro("Module1.E2E_InputBox", { timeoutMs: 40000, dialogPolicy: { mode: "errors" } });
  await new Promise((r) => setTimeout(r, 3000));

  const list = await client.listDialogs();
  console.log("listDialogs:", list.output);
  const dialogs = JSON.parse(list.output || "[]") as Array<{ handle: string; hasEdit?: boolean; title: string; text: string }>;
  const inputDialog = dialogs.find((d) => d.hasEdit);
  if (!inputDialog) {
    console.log("FAIL: 没有探测到含输入框的弹窗");
    process.exit(1);
  }
  console.log("找到 InputBox 弹窗 handle =", inputDialog.handle);

  const fill = await client.fillDialog(inputDialog.handle, "42", true);
  console.log("fillDialogInput:", fill.success, fill.message, JSON.stringify(fill.details || {}));

  const r = await runPromise;
  console.log("runMacro result:", r.success, r.message);

  const b4 = execSync(`"${exeAbs}" -q range get-values --session ${sid} --sheet Sheet1 --range B4`, { encoding: "utf-8", timeout: 30000 });
  console.log("B4 =", b4.trim());
  process.exit(b4.includes("N=42") ? 0 : 1);
}

main().catch((e) => {
  console.error("PROBE FAILED:", e instanceof Error ? e.message : String(e));
  process.exit(1);
});
