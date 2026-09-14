/**
 * E2E：宏运行桥 + 弹窗捕获测试。
 *
 * 前置：工作簿已由 excelcli session open --show 打开。
 * 用法：node build/e2e-bridge.js <workbookPath> <syncDir>
 */
import { existsSync, writeFileSync, readFileSync } from "fs";
import { join } from "path";
import { unlink } from "fs/promises";
import { VbaClient } from "../src/client/vbaClient";
import { startMacroBridge } from "../src/client/macroBridge";
import { findExcelPidForWorkbook } from "../src/native/processUtils";

const wbPath = process.argv[2];
const syncDir = process.argv[3];
if (!wbPath || !syncDir) {
  console.error("用法: node build/e2e-bridge.js <workbookPath> <syncDir>");
  process.exit(2);
}

let pass = 0;
let total = 0;
function assert(name: string, ok: boolean, detail: string): void {
  total++;
  if (ok) pass++;
  console.log(`[${ok ? "PASS" : "FAIL"}] ${name} ${detail}`);
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitResult(timeoutMs: number): Promise<any | null> {
  const resultPath = join(syncDir, "macro-run-result.json");
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (existsSync(resultPath)) {
      const raw = readFileSync(resultPath, "utf-8");
      await unlink(resultPath).catch(() => {});
      try {
        return JSON.parse(raw);
      } catch {
        return null;
      }
    }
    await sleep(300);
  }
  return null;
}

// 允许编译错误等复杂弹窗流程耗时更长
const RESULT_WAIT_MS = 150000;

async function excelAlive(): Promise<boolean> {
  const pid = await findExcelPidForWorkbook(wbPath);
  return !!pid;
}

async function main(): Promise<void> {
  const pid = await findExcelPidForWorkbook(wbPath);
  if (!pid) {
    console.error("工作簿未打开（先 excelcli session open --show）");
    process.exit(2);
  }
  console.log(`Excel PID=${pid}`);
  const client = new VbaClient(wbPath, pid);

  const notifications: string[] = [];
  const bridge = startMacroBridge({
    getClient: () => client,
    getSyncDir: () => syncDir,
    intervalMs: 1000,
    log: (m) => console.log(`  [bridge] ${m}`),
    notify: (m) => {
      notifications.push(m);
      console.log(`  [notify] ${m}`);
    },
  });

  // ---- 前置：写入运行时错误宏（自包含，不依赖上次状态）----
  const { execSync: execSyncEarly } = await import("child_process");
  const exeAbsEarly = require("path").resolve("dist", "excelcli.exe");
  const rtCode = 'Public Function E2E_Base() As String\r\n    E2E_Base = "base"\r\nEnd Function\r\n\r\nSub E2E_RuntimeErr()\r\n    Dim x As Long\r\n    x = 1 / 0\r\nEnd Sub\r\n';
  const rtPath = join(syncDir, "mod_runtime_err.bas");
  writeFileSync(rtPath, rtCode, "utf-8");
  try {
    execSyncEarly(`"${exeAbsEarly}" -q vba update --session ${process.env.E2E_SID} --module-name Module1 --vba-code-file "${rtPath.replace(/\//g, "\\")}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      timeout: 90000,
    });
  } catch (e: any) {
    console.log("  [warn] 前置模块写入失败:", String(e.message).slice(0, 200));
  }

  // ---- 场景 1：运行时错误宏 ----
  for (const f of ["macro-run-result.json", "excel-dialogs.json"]) {
    await unlink(join(syncDir, f)).catch(() => {});
  }
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_RuntimeErr", timeoutMs: 30000 }), "utf-8");
  const r1 = await waitResult(RESULT_WAIT_MS);
  assert("场景1 运行时错误：结果文件生成", !!r1, r1 ? `success=${r1.success} msg=${(r1.message || "").slice(0, 100)}` : "超时未生成");
  const allText1 = JSON.stringify(r1?.dialogs || []);
  assert("场景1 抓到运行时错误弹窗文本", /运行时错误|除数为零|Divide by zero|1004|11/i.test(allText1), allText1.slice(0, 200));
  assert("场景1 弹窗已被自动结束", (r1?.dialogs || []).some((d: any) => d.autoHandled), "");
  assert("场景1 Excel 仍存活", await excelAlive(), "");

  // ---- 场景 2：编译错误宏（先写坏代码，再桥接运行）----
  // 调用不存在的过程 → 运行时必弹"编译错误: 子过程未定义"对话框
  const badCode = 'Public Function E2E_Base() As String\r\n    E2E_Base = "base"\r\nEnd Function\r\n\r\nSub E2E_CompileErr2()\r\n    Call ThisSubDoesNotExistXYZ99\r\nEnd Sub\r\n';
  console.log("  [step] 用 CLI vba update 写入编译错误代码...");
  const { execSync } = await import("child_process");
  const exeAbs = require("path").resolve("dist", "excelcli.exe");
  const badPath = join(syncDir, "mod_compile_err.bas");
  writeFileSync(badPath, badCode, "utf-8");
  try {
    execSync(`"${exeAbs}" -q vba update --session ${process.env.E2E_SID} --module-name Module1 --vba-code-file "${badPath.replace(/\//g, "\\")}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      timeout: 90000,
    });
  } catch (e: any) {
    console.log("  [warn] vba update 失败(可能语法过不了):", String(e.message).slice(0, 200));
  }
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_CompileErr2", timeoutMs: 30000 }), "utf-8");
  const r2 = await waitResult(RESULT_WAIT_MS);
  assert("场景2 编译错误：结果文件生成", !!r2, r2 ? `success=${r2.success}` : "超时未生成");
  const allText2 = JSON.stringify({ dialogs: r2?.dialogs, message: r2?.message });
  assert("场景2 抓到编译错误信息", /编译错误|子过程|未定义|Compile|undefined|无法/i.test(allText2), allText2.slice(0, 240));
  assert("场景2 Excel 仍存活", await excelAlive(), "");

  // ---- 场景 3：正常宏桥接运行 + 读取运行结果 ----
  const okCode = 'Public Function E2E_Base() As String\r\n    E2E_Base = "base"\r\nEnd Function\r\n\r\nSub E2E_Result2()\r\n    Range("B2").Value = "BRIDGE_OK"\r\nEnd Sub\r\n';
  const okPath = join(syncDir, "mod_ok2.bas");
  writeFileSync(okPath, okCode, "utf-8");
  try {
    execSync(`"${exeAbs}" -q vba update --session ${process.env.E2E_SID} --module-name Module1 --vba-code-file "${okPath.replace(/\//g, "\\")}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      timeout: 90000,
    });
  } catch (e: any) {
    console.log("  [warn] vba update 失败:", String(e.message).slice(0, 200));
  }
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_Result2", timeoutMs: 30000, saveAfterRun: true }), "utf-8");
  const r3 = await waitResult(RESULT_WAIT_MS);
  assert("场景3 正常宏：success=true", r3?.success === true, r3?.message || "");
  assert("场景3 Excel 仍存活", await excelAlive(), "");
  const b2 = execSync(`"${exeAbs}" -q range get-values --session ${process.env.E2E_SID} --sheet Sheet1 --range B2`, { cwd: process.cwd(), encoding: "utf-8", timeout: 30000 });
  assert("场景3 运行结果可读取", b2.includes("BRIDGE_OK"), b2.slice(0, 120));

  // ---- 场景 4-6：交互弹窗宏（MsgBox / InputBox / 是否确认）—— dialogMode: auto ----
  const interactiveCode = [
    'Public Function E2E_Base() As String',
    '    E2E_Base = "base"',
    'End Function',
    '',
    'Sub E2E_MsgBox()',
    '    MsgBox "处理完成：10 行"',
    '    Range("B3").Value = "AFTER_MSGBOX"',
    'End Sub',
    '',
    'Sub E2E_InputBox()',
    '    Dim v As String',
    '    v = InputBox("请输入数量")',
    '    Range("B4").Value = "N=" & v',
    'End Sub',
    '',
    'Sub E2E_Confirm()',
    '    Dim a As VbMsgBoxResult',
    '    a = MsgBox("继续处理吗？", vbYesNo)',
    '    Range("B5").Value = IIf(a = vbYes, "YES", "NO")',
    'End Sub',
  ].join("\r\n") + "\r\n";
  const interPath = join(syncDir, "mod_interactive.bas");
  writeFileSync(interPath, interactiveCode, "utf-8");
  try {
    execSync(`"${exeAbs}" -q vba update --session ${process.env.E2E_SID} --module-name Module1 --vba-code-file "${interPath.replace(/\//g, "\\")}"`, {
      cwd: process.cwd(),
      stdio: "pipe",
      timeout: 90000,
    });
  } catch (e: any) {
    console.log("  [warn] 交互宏写入失败:", String(e.message).slice(0, 200));
  }

  // 场景 4：MsgBox 自动点确定，宏继续执行
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_MsgBox", timeoutMs: 30000, dialogMode: "auto" }), "utf-8");
  const r4 = await waitResult(RESULT_WAIT_MS);
  assert("场景4 MsgBox：结果文件生成", !!r4, r4 ? `success=${r4.success}` : "超时未生成");
  const allText4 = JSON.stringify(r4?.dialogs || []);
  assert("场景4 抓到 MsgBox 文本", allText4.includes("处理完成：10 行"), allText4.slice(0, 200));
  assert("场景4 弹窗已自动点确定", (r4?.dialogs || []).some((d: any) => d.autoHandled), "");
  const b3 = execSync(`"${exeAbs}" -q range get-values --session ${process.env.E2E_SID} --sheet Sheet1 --range B3`, { cwd: process.cwd(), encoding: "utf-8", timeout: 30000 });
  assert("场景4 MsgBox 后宏继续执行", b3.includes("AFTER_MSGBOX"), b3.slice(0, 120));

  // 场景 5：InputBox 自动填入
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_InputBox", timeoutMs: 30000, dialogMode: "auto", inputValue: "42" }), "utf-8");
  const r5 = await waitResult(RESULT_WAIT_MS);
  assert("场景5 InputBox：结果文件生成", !!r5, r5 ? `success=${r5.success}` : "超时未生成");
  console.log("  [debug] r5 =", JSON.stringify(r5, null, 1));
  const b4 = execSync(`"${exeAbs}" -q range get-values --session ${process.env.E2E_SID} --sheet Sheet1 --range B4`, { cwd: process.cwd(), encoding: "utf-8", timeout: 30000 });
  assert("场景5 InputBox 输入已填入(N=42)", b4.includes("N=42"), b4.slice(0, 120));

  // 场景 6：是否确认框，指定点"是"
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_Confirm", timeoutMs: 30000, dialogMode: "auto", confirmButton: "是" }), "utf-8");
  const r6 = await waitResult(RESULT_WAIT_MS);
  assert("场景6 确认框：结果文件生成", !!r6, r6 ? `success=${r6.success}` : "超时未生成");
  const b5 = execSync(`"${exeAbs}" -q range get-values --session ${process.env.E2E_SID} --sheet Sheet1 --range B5`, { cwd: process.cwd(), encoding: "utf-8", timeout: 30000 });
  assert("场景6 确认框点了'是'(YES)", b5.includes("YES"), b5.slice(0, 120));
  assert("场景6 Excel 仍存活", await excelAlive(), "");

  bridge.stop();
  console.log(`\n===== 桥接测试: ${pass}/${total} 通过 =====`);
  process.exit(pass === total ? 0 : 1);
}

main().catch((e) => {
  console.error("E2E-BRIDGE FAILED:", e instanceof Error ? e.message : String(e));
  process.exit(1);
});
