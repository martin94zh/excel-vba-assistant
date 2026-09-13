/**
 * E2E 测试台：在 VS Code 宿主外直接驱动插件的 VbaClient（真实 COM 同步代码路径）。
 *
 * 用法：node build/e2e-harness.js <workbookPath> <syncDir>
 * 前置：工作簿已由 excelcli session open --show 打开（前台可见）。
 */
import { readdir, readFile, writeFile } from "fs/promises";
import { join } from "path";
import { VbaClient } from "../src/client/vbaClient";
import { findExcelPidForWorkbook } from "../src/native/processUtils";

interface StepResult {
  step: string;
  ok: boolean;
  ms: number;
  detail: string;
}

const results: StepResult[] = [];

function record(step: string, ok: boolean, ms: number, detail: string): void {
  results.push({ step, ok, ms, detail });
  console.log(`[${ok ? "PASS" : "FAIL"}] ${step} (${ms}ms) ${detail}`);
}

async function timed<T>(step: string, fn: () => Promise<T>, judge: (r: T) => { ok: boolean; detail: string }): Promise<T> {
  const t0 = Date.now();
  try {
    const r = await fn();
    const { ok, detail } = judge(r);
    record(step, ok, Date.now() - t0, detail);
    return r;
  } catch (err) {
    record(step, false, Date.now() - t0, err instanceof Error ? err.message : String(err));
    throw err;
  }
}

async function main(): Promise<void> {
  const wbPath = process.argv[2];
  const syncDir = process.argv[3];
  if (!wbPath || !syncDir) {
    console.error("用法: node build/e2e-harness.js <workbookPath> <syncDir>");
    process.exit(2);
  }

  // 0. PID 定位（窗口标题法，验证前台窗口存在）
  const pid = await timed("PID 定位（前台窗口）", () => findExcelPidForWorkbook(wbPath), (p) => ({
    ok: !!p,
    detail: `PID=${p ?? "未找到"}`,
  }));

  const client = new VbaClient(wbPath, pid || undefined);

  // 1. COM 访问（VBA 信任）
  await timed("COM/VBA 访问校验", () => client.checkAccess(), (e) => ({
    ok: !e,
    detail: e ? String(e) : "VBAProject 可访问",
  }));

  // 2. VBE → 本地（导出）
  await timed("VBE → 本地 同步", () => client.syncVbeToLocal(syncDir), (r) => ({
    ok: r.success,
    detail: (r.output || r.message || "").split("\n").slice(0, 3).join(" | ").slice(0, 160),
  }));

  // 3. 本地修改 → VBE（导入）：向第一个 .bas 追加标记函数
  const modulesDir = join(syncDir, "模块");
  const files = await readdir(modulesDir).catch(() => [] as string[]);
  const basFile = files.find((f) => f.endsWith(".bas"));
  if (!basFile) {
    record("本地 → VBE 同步", false, 0, `同步目录无标准模块（模块/ 下: ${files.join(",") || "空"}）`);
  } else {
    const basPath = join(modulesDir, basFile);
    const moduleName = basFile.replace(/\.bas$/i, "");
    const original = await readFile(basPath, "utf-8");
    const marker = `E2E_MARK_${Date.now()}`;
    const markerCode = `Public Function ${marker}() As String\r\n    ${marker} = "ok"\r\nEnd Function\r\n`;
    await writeFile(basPath, original.replace(/\r?\n$/, "") + "\r\n\r\n" + markerCode, "utf-8");

    await timed("本地 → VBE 同步（写入标记函数）", () => client.syncLocalToVbe(syncDir), (r) => ({
      ok: r.success,
      detail: (r.output || r.message || "").split("\n").slice(0, 3).join(" | ").slice(0, 160),
    }));

    // 4. 回读验证：VBE → 本地 再导出，检查标记存在
    const r2 = await timed("回读验证（再次 VBE → 本地）", () => client.syncVbeToLocal(syncDir), (r) => ({
      ok: r.success,
      detail: "二次导出完成",
    }));
    void r2;
    const after = await readFile(basPath, "utf-8");
    record("双向闭环：标记函数已在 VBE 中", after.includes(marker), 0, marker);

    // 5. VBE checksum 速度（自动同步轮询的代价）
    const t0 = Date.now();
    const cs = await client.getVbeCodeChecksum();
    record("VBE checksum（3 秒轮询单次代价）", cs.success, Date.now() - t0, (cs.output || "").slice(0, 16));
  }

  // 6. 宏列表
  await timed("宏列表（listMacros）", () => client.listMacros(), (r) => ({
    ok: r.success && !!r.output,
    detail: (r.output || r.message || "").slice(0, 120),
  }));

  // 7. 置顶开/关
  await timed("Excel 置顶 ON", () => client.setWindowTopMost(true), (r) => ({ ok: r.success, detail: r.message }));
  await timed("Excel 置顶 OFF", () => client.setWindowTopMost(false), (r) => ({ ok: r.success, detail: r.message }));

  // 汇总
  const pass = results.filter((r) => r.ok).length;
  console.log(`\n===== 汇总: ${pass}/${results.length} 通过 =====`);
  const totalSync = results.filter((r) => r.step.includes("同步")).reduce((a, b) => a + b.ms, 0);
  console.log(`同步相关步骤累计耗时: ${totalSync}ms`);
  if (pass !== results.length) process.exit(1);
}

main().catch(() => process.exit(1));
