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
      timeout: 30000,
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
      timeout: 30000,
    });
  } catch (e: any) {
    console.log("  [warn] vba update 失败:", String(e.message).slice(0, 200));
  }
  writeFileSync(join(syncDir, "macro-run.json"), JSON.stringify({ macro: "Module1.E2E_Result2", timeoutMs: 30000 }), "utf-8");
  const r3 = await waitResult(RESULT_WAIT_MS);
  assert("场景3 正常宏：success=true", r3?.success === true, r3?.message || "");
  assert("场景3 Excel 仍存活", await excelAlive(), "");
  const b2 = execSync(`"${exeAbs}" -q range get-values --session ${process.env.E2E_SID} --sheet Sheet1 --range B2`, { cwd: process.cwd(), encoding: "utf-8", timeout: 30000 });
  assert("场景3 运行结果可读取", b2.includes("BRIDGE_OK"), b2.slice(0, 120));

  bridge.stop();
  console.log(`\n===== 桥接测试: ${pass}/${total} 通过 =====`);
  process.exit(pass === total ? 0 : 1);
}

main().catch((e) => {
  console.error("E2E-BRIDGE FAILED:", e instanceof Error ? e.message : String(e));
  process.exit(1);
});
