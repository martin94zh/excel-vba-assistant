/**
 * Excel VBA Assistant - Excel 窗口辅助（excelcli window 命令组）
 *
 * 仅封装状态栏进度文本（set-status-bar / clear-status-bar）：宏运行 / 同步期间
 * 在前台 Excel 状态栏显示"插件正在做什么"，结束自动恢复。
 *
 * 注意：窗口分屏（arrange）不在此处——excelcli 的 arrange 需要把 Excel 带到前台，
 * ExcelMcp 2.2.0 起受 Windows 前台锁限制在后台 daemon 中确定性失败，
 * 已改为 VbaClient.setWindowArrange（Win32 SetWindowPos，不抢焦点）。
 *
 * 约定：全部 fire-and-forget——未连接（无 cliSessionId）或命令失败时静默跳过，
 * 只写日志，绝不阻塞或打断同步 / 宏运行主流程。
 */
import { runExcelCli } from "./cliRunner";

export type ExcelWindowPreset =
  | "left-half"
  | "right-half"
  | "top-half"
  | "bottom-half"
  | "center"
  | "full-screen";

export const WINDOW_PRESETS: Array<{ id: ExcelWindowPreset; label: string; description: string }> = [
  { id: "left-half", label: "左半屏", description: "Excel 靠左半屏，编辑器在右" },
  { id: "right-half", label: "右半屏", description: "Excel 靠右半屏，编辑器在左" },
  { id: "top-half", label: "上半屏", description: "Excel 靠上半屏" },
  { id: "bottom-half", label: "下半屏", description: "Excel 靠下半屏" },
  { id: "center", label: "居中", description: "Excel 居中窗口" },
  { id: "full-screen", label: "全屏", description: "Excel 最大化占满工作区" },
];

/** 单条 window 命令的统一封装：失败返回 false，不抛出 */
async function runWindowCommand(
  cliPath: string,
  sessionId: string,
  args: string[],
  label: string
): Promise<boolean> {
  try {
    const result = await runExcelCli(cliPath, ["-q", "window", ...args, "--session", sessionId], 30000);
    if (!result.success) {
      console.warn(`[excelWindow] ${label} 失败：${result.error || result.raw}`);
      return false;
    }
    return true;
  } catch (err) {
    console.warn(`[excelWindow] ${label} 异常：${err instanceof Error ? err.message : String(err)}`);
    return false;
  }
}

/** 在 Excel 状态栏显示进度文本（如"正在运行宏 Module1.Main..."） */
export async function setExcelStatusText(
  cliPath: string | undefined,
  sessionId: string | undefined,
  text: string
): Promise<void> {
  if (!cliPath || !sessionId) return;
  await runWindowCommand(cliPath, sessionId, ["set-status-bar", "--text", text], "设置状态栏文本");
}

/** 恢复 Excel 状态栏为默认 */
export async function clearExcelStatusText(
  cliPath: string | undefined,
  sessionId: string | undefined
): Promise<void> {
  if (!cliPath || !sessionId) return;
  await runWindowCommand(cliPath, sessionId, ["clear-status-bar"], "清除状态栏文本");
}
