/**
 * 宏运行桥 + Excel 弹窗监视器
 *
 * 背景：AI 通过 excelcli 直接 `vba run` 一个会抛运行时错误/编译错误的宏时，
 * VBA 错误弹窗会阻塞 daemon 的 COM 线程，超时后 daemon 会取消并销毁会话
 * （Excel 进程退出），AI 拿不到任何报错文本。
 *
 * 本模块提供两条补救能力（在插件宿主或测试台中以后台轮询运行）：
 * 1. 宏运行桥：AI 往同步目录写 macro-run.json（{"macro":"Module1.X","timeoutMs":30000}），
 *    由插件经 COM（runMacroWithDialogHandling）执行——弹窗被自动检测、抓取文本并点击"结束"，
 *    完整结果写回 macro-run-result.json，Excel 保持存活。
 * 2. 弹窗监视：任何来源（含 AI 直接 vba run）触发的错误弹窗都会被抓取文本、
 *    自动点击"结束/确定"并记录到 excel-dialogs.json，避免 daemon/Excel 被卡死。
 *
 * 本模块不依赖 vscode，可在宿主外测试。
 */
import { existsSync } from "fs";
import { mkdir, readFile, rename, unlink, writeFile } from "fs/promises";
import { join } from "path";
import { VbaClient } from "./vbaClient";
import type { MacroDialogPolicy } from "./vbaMacroRunner";

export interface MacroRunRequest {
  macro: string;
  timeoutMs?: number;
  /** 交互弹窗处理：auto = MsgBox 自动点确定/确认框点 confirmButton/InputBox 填 inputValue（默认，AI 测试场景） */
  dialogMode?: "auto" | "errors";
  /** 确认（是/否）对话框点击的按钮，默认"取消" */
  confirmButton?: string;
  /** InputBox 自动填入的文本 */
  inputValue?: string;
}

export interface CapturedDialog {
  time: string;
  kind: string;
  title: string;
  text: string;
  buttons: string[];
  autoHandled: boolean;
  autoAction?: string;
  source?: string;
}

export interface MacroRunResultFile {
  finishedAt: string;
  macro: string;
  success: boolean;
  message: string;
  output?: string;
  dialogs: CapturedDialog[];
  error?: string;
}

export interface MacroBridgeOptions {
  /** 返回当前连接的 VbaClient；未连接时返回 null（轮询空转） */
  getClient: () => VbaClient | null;
  /** 返回当前同步目录；未设置时返回 null（桥接文件无处安放，轮询空转） */
  getSyncDir: () => string | null;
  intervalMs?: number;
  log?: (message: string) => void;
  /** 检测到新的错误弹窗时回调（VS Code 宿主中显示警告通知） */
  notify?: (message: string) => void;
}

const DIALOGS_FILE = "excel-dialogs.json";
const RUN_REQUEST_FILE = "macro-run.json";
const RUN_RESULT_FILE = "macro-run-result.json";
const MAX_CAPTURED = 20;
const FINGERPRINT_TTL_MS = 10 * 60 * 1000;

function fingerprintOf(d: { title: string; text: string; buttons?: string[] }): string {
  return `${(d.title || "").trim()}|${(d.text || "").replace(/\s+/g, " ").trim()}|${(d.buttons || []).join(",")}`;
}

async function writeAtomic(filePath: string, content: string): Promise<void> {
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await writeFile(tmp, content, "utf-8");
  await rename(tmp, filePath).catch(async () => {
    await writeFile(filePath, content, "utf-8");
    await unlink(tmp).catch(() => {});
  });
}

export function startMacroBridge(options: MacroBridgeOptions): { stop: () => void } {
  const intervalMs = options.intervalMs ?? 2000;
  const log = options.log ?? (() => {});
  const captured: CapturedDialog[] = [];
  const seenFingerprints = new Map<string, number>();
  const dismissedHandles = new Set<string>();
  let busy = false;

  async function recordDialogs(
    client: VbaClient,
    syncDir: string,
    rawDialogs: Array<{ kind?: string; title?: string; text?: string; buttons?: string[]; handle?: string; autoHandled?: boolean; autoAction?: string }>,
    source: string
  ): Promise<void> {
    const now = new Date().toISOString();
    for (const raw of rawDialogs) {
      const kind = raw.kind || "unknown";
      const fp = fingerprintOf({ title: raw.title || "", text: raw.text || "", buttons: raw.buttons });
      if (kind !== "vb_runtime_error" && kind !== "error") continue;
      if (seenFingerprints.has(fp)) continue;
      seenFingerprints.set(fp, Date.now());
      captured.unshift({
        time: now,
        kind,
        title: raw.title || "",
        text: raw.text || "",
        buttons: raw.buttons || [],
        autoHandled: !!raw.autoHandled,
        autoAction: raw.autoAction,
        source,
      });
      while (captured.length > MAX_CAPTURED) captured.pop();
      log(`检测到错误弹窗（${source}）：${raw.title} — ${(raw.text || "").replace(/\s+/g, " ").slice(0, 160)}`);
      options.notify?.(`检测到 Excel 报错弹窗（已记录到 ${DIALOGS_FILE}）：${raw.title}: ${(raw.text || "").slice(0, 80)}`);
    }

    // 自动结束错误弹窗（对 XLMAIN/编译错误点"结束/确定"），解除 COM 阻塞
    for (const raw of rawDialogs) {
      const kind = raw.kind || "unknown";
      if (kind !== "vb_runtime_error" && kind !== "error") continue;
      if (!raw.handle || dismissedHandles.has(raw.handle)) continue;
      const clickResult = await client.clickDialog(raw.handle, "auto");
      if (clickResult.success) {
        dismissedHandles.add(raw.handle);
        const entry = captured.find((c) => fingerprintOf(c) === fingerprintOf({ title: raw.title || "", text: raw.text || "", buttons: raw.buttons }));
        if (entry) {
          entry.autoHandled = true;
          entry.autoAction = typeof clickResult.message === "string" ? clickResult.message : "auto";
        }
        log(`已自动关闭错误弹窗：${raw.title}`);
      }
    }

    // 清理过期指纹
    const cutoff = Date.now() - FINGERPRINT_TTL_MS;
    for (const [fp, ts] of seenFingerprints) {
      if (ts < cutoff) seenFingerprints.delete(fp);
    }

    if (captured.length > 0) {
      await mkdir(syncDir, { recursive: true });
      await writeAtomic(join(syncDir, DIALOGS_FILE), JSON.stringify({ updatedAt: now, dialogs: captured.slice(0, MAX_CAPTURED) }, null, 2));
    }
  }

  async function runBridgeMacro(client: VbaClient, syncDir: string, raw: string): Promise<void> {
    let req: MacroRunRequest | null = null;
    try {
      req = JSON.parse(raw) as MacroRunRequest;
    } catch {
      req = null;
    }
    if (!req || !req.macro || typeof req.macro !== "string") {
      await writeAtomic(join(syncDir, RUN_RESULT_FILE), JSON.stringify({
        finishedAt: new Date().toISOString(),
        macro: "",
        success: false,
        message: "macro-run.json 解析失败：需要 {\"macro\":\"Module1.X\",\"timeoutMs\":30000}",
        dialogs: [],
        error: "invalid_request",
      } satisfies MacroRunResultFile, null, 2));
      return;
    }
    log(`宏运行桥：执行 ${req.macro}`);
    const policy: MacroDialogPolicy = {
      mode: req.dialogMode ?? "auto",
      confirmButton: req.confirmButton,
      inputValue: req.inputValue,
    };
    const result = await client.runMacro(req.macro, { timeoutMs: req.timeoutMs ?? 45000, dialogPolicy: policy });
    const dialogs = ((result.details?.dialogs as Array<{ kind?: string; title?: string; text?: string; buttons?: string[]; autoHandled?: boolean; autoAction?: string }>) || []).map((d) => ({
      time: new Date().toISOString(),
      kind: d.kind || "unknown",
      title: d.title || "",
      text: d.text || "",
      buttons: d.buttons || [],
      autoHandled: !!d.autoHandled,
      autoAction: d.autoAction,
      source: "macro-bridge",
    }));
    await recordDialogs(client, syncDir, dialogs, "macro-bridge");
    const payload: MacroRunResultFile = {
      finishedAt: new Date().toISOString(),
      macro: req.macro,
      success: result.success,
      message: result.message,
      output: result.output,
      dialogs,
    };
    if (!result.success) payload.error = "macro_failed";
    await writeAtomic(join(syncDir, RUN_RESULT_FILE), JSON.stringify(payload, null, 2));
    log(`宏运行桥：${req.macro} → ${result.success ? "成功" : "失败"}（${dialogs.length} 个弹窗记录）`);
  }

  async function tick(): Promise<void> {
    if (busy) return;
    const client = options.getClient();
    if (!client) return;
    const syncDir = options.getSyncDir();
    if (!syncDir) return;
    busy = true;
    try {
      const reqPath = join(syncDir, RUN_REQUEST_FILE);
      if (existsSync(reqPath)) {
        let raw = "";
        try {
          raw = await readFile(reqPath, "utf-8");
        } catch (e) {
          raw = "";
        }
        await unlink(reqPath).catch(() => {});
        await runBridgeMacro(client, syncDir, raw);
        return;
      }

      // 弹窗巡检（抓取文本 + 自动结束错误弹窗）
      const inspection = await client.listDialogs();
      if (inspection.success && inspection.output) {
        try {
          const parsed = JSON.parse(inspection.output) as Array<{ kind?: string; title?: string; text?: string; buttons?: string[]; handle?: string }>;
          const dialogs = Array.isArray(parsed) ? parsed : [parsed];
          await recordDialogs(client, syncDir, dialogs, "watcher");
        } catch {
          // 解析失败忽略，下一轮再试
        }
      }
    } catch (e) {
      log(`宏运行桥巡检异常：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      busy = false;
    }
  }

  const timer = setInterval(() => {
    void tick();
  }, intervalMs);

  return {
    stop: () => {
      clearInterval(timer);
    },
  };
}
