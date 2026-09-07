/**
 * VBA 同步引擎
 *
 * 通过组合 mcp-server-excel 的 vba list/view/import/update/delete 工具，
 * 实现 VBE ↔ 本地文件系统的双向同步。
 */
import { mkdir, writeFile, readFile, readdir, rm } from "fs/promises";
import { join, resolve, basename } from "path";
import { existsSync } from "fs";
import { ExcelClient } from "./excelClient";

export interface WorkbookComponentEntry {
  name: string;
  type: ComponentType;
  file: string;
}

export type ComponentType = "document" | "worksheet" | "standardModule" | "classModule" | "userForm";

export interface WorkbookManifest {
  workbookName: string;
  workbookPath: string;
  lastSyncAt: string;
  components: WorkbookComponentEntry[];
}

const TYPE_TO_FOLDER: Record<string, { folder: string; ext: string; typeName: ComponentType; canCreate: boolean }> = {
  "standardModule": { folder: "模块", ext: ".bas", typeName: "standardModule", canCreate: true },
  "classModule": { folder: "类模块", ext: ".cls", typeName: "classModule", canCreate: true },
  "userForm": { folder: "窗体", ext: ".frm", typeName: "userForm", canCreate: true },
  "worksheet": { folder: "Microsoft Excel 对象", ext: ".wks", typeName: "worksheet", canCreate: false },
  "document": { folder: "Microsoft Excel 对象", ext: ".wbk", typeName: "document", canCreate: false },
};

const EXT_TO_TYPE: Record<string, ComponentType> = {
  ".bas": "standardModule",
  ".cls": "classModule",
  ".frm": "userForm",
  ".wks": "worksheet",
  ".wbk": "document",
};

export class VbaSyncEngine {
  constructor(private client: ExcelClient) {}

  /** VBE → 本地：导出所有 VBA 组件到本地目录 */
  async syncVbeToLocal(workbookPath: string, localDir: string): Promise<{ success: boolean; message: string; output: string }> {
    const absDir = resolve(localDir);
    const wbName = basename(workbookPath);

    try {
      // 1. 列出所有模块
      const modules = await this.client.vbaList(workbookPath);

      // 2. 确保目录结构
      await mkdir(join(absDir, "Microsoft Excel 对象"), { recursive: true });
      await mkdir(join(absDir, "模块"), { recursive: true });
      await mkdir(join(absDir, "类模块"), { recursive: true });
      await mkdir(join(absDir, "窗体"), { recursive: true });

      const components: WorkbookComponentEntry[] = [];
      const logLines: string[] = [];

      // 3. 逐个获取代码并写入
      for (const mod of modules) {
        const mapping = TYPE_TO_FOLDER[mod.type];
        if (!mapping) continue;

        const folderPath = join(absDir, mapping.folder);
        await mkdir(folderPath, { recursive: true });

        const fileName = `${mod.name}${mapping.ext}`;
        const filePath = join(folderPath, fileName);
        const relFile = `${mapping.folder}/${fileName}`;

        const code = await this.client.vbaView(workbookPath, mod.name);

        let changed = true;
        if (existsSync(filePath)) {
          const existing = await readFile(filePath, "utf-8");
          if (existing === code) changed = false;
        }

        if (changed) {
          await writeFile(filePath, code, "utf-8");
          logLines.push(`同步: ${relFile}`);
        } else {
          logLines.push(`未变: ${relFile}`);
        }

        components.push({ name: mod.name, type: mapping.typeName, file: relFile });
      }

      // 4. 清理本地孤儿文件
      const expectedFiles = new Set(components.map((c) => c.file.toLowerCase()));
      const deleted = await this.removeOrphanFiles(absDir, expectedFiles);
      if (deleted > 0) {
        logLines.push(`删除 ${deleted} 个本地孤儿文件`);
      }

      // 5. 写入 workbook.json
      const manifest: WorkbookManifest = {
        workbookName: wbName,
        workbookPath,
        lastSyncAt: new Date().toISOString(),
        components,
      };
      await writeFile(join(absDir, "workbook.json"), JSON.stringify(manifest, null, 2), "utf-8");

      logLines.push(`同步完成: 共 ${components.length} 个组件`);
      return { success: true, message: logLines[logLines.length - 1], output: logLines.join("\n") };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return { success: false, message: `同步失败: ${msg}`, output: msg };
    }
  }

  /** 本地 → VBE：将本地代码覆盖性写回 */
  async syncLocalToVbe(workbookPath: string, localDir: string): Promise<{ success: boolean; message: string; output: string }> {
    const absDir = resolve(localDir);

    try {
      // 1. 获取当前 VBE 模块列表
      const vbeModules = await this.client.vbaList(workbookPath);
      const existingNames = new Map(vbeModules.map((m) => [m.name, m.type]));

      // 2. 收集本地文件
      const localFiles = await this.collectLocalVbaFiles(absDir);

      let added = 0, updated = 0, unchanged = 0, skipped = 0, deleted = 0;
      const logLines: string[] = [];

      // 3. 遍历本地文件：更新或导入
      for (const file of localFiles) {
        const code = await readFile(file.path, "utf-8");
        const vbeType = existingNames.get(file.moduleName);

        if (vbeType !== undefined) {
          // 更新现有模块
          const existingCode = await this.client.vbaView(workbookPath, file.moduleName);
          const cleanExisting = this.normalizeCode(existingCode);
          const cleanLocal = this.normalizeCode(code);

          if (cleanExisting === cleanLocal) {
            unchanged++;
            logLines.push(`未变: ${file.relPath}`);
          } else {
            await this.client.vbaUpdate(workbookPath, file.moduleName, code);
            updated++;
            logLines.push(`更新: ${file.relPath}`);
          }
        } else {
          // 新建模块（仅标准模块/类模块/窗体）
          if (file.type === "document" || file.type === "worksheet") {
            skipped++;
            logLines.push(`跳过: ${file.relPath} (Excel 对象模块不能从本地新建)`);
          } else {
            await this.client.vbaImport(workbookPath, file.moduleName, code, file.type);
            added++;
            logLines.push(`新增: ${file.relPath}`);
          }
        }
      }

      // 4. 删除同步：VBE 中存在但本地不存在，且非 document 类型
      for (const [name, type] of existingNames) {
        if (!localFiles.find((f) => f.moduleName === name) && type !== "document" && type !== "worksheet") {
          await this.client.vbaDelete(workbookPath, name);
          deleted++;
          logLines.push(`删除: ${name}`);
        }
      }

      // 5. 关闭 session 保存
      await this.client.getSessionStore().closeSession(workbookPath, true);

      const summary = `同步完成: 新增 ${added} / 更新 ${updated} / 跳过 ${skipped} / 未变 ${unchanged} / 删除 ${deleted}`;
      logLines.push("", summary);
      return { success: true, message: summary, output: logLines.join("\n") };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return { success: false, message: `同步失败: ${msg}`, output: msg };
    }
  }

  // ============================================================
  // 私有辅助
  // ============================================================

  private async collectLocalVbaFiles(absDir: string): Promise<Array<{ path: string; relPath: string; moduleName: string; type: ComponentType }>> {
    const result: Array<{ path: string; relPath: string; moduleName: string; type: ComponentType }> = [];

    for (const [ext, type] of Object.entries(EXT_TO_TYPE)) {
      const folder = TYPE_TO_FOLDER[type].folder;
      const dirPath = join(absDir, folder);
      if (!existsSync(dirPath)) continue;

      const files = await readdir(dirPath);
      for (const f of files) {
        if (f.toLowerCase().endsWith(ext.toLowerCase())) {
          const moduleName = f.slice(0, -ext.length);
          result.push({
            path: join(dirPath, f),
            relPath: `${folder}/${f}`,
            moduleName,
            type,
          });
        }
      }
    }

    return result;
  }

  private async removeOrphanFiles(absDir: string, expectedFiles: Set<string>): Promise<number> {
    let deleted = 0;
    for (const [ext, type] of Object.entries(EXT_TO_TYPE)) {
      const folder = TYPE_TO_FOLDER[type].folder;
      const dirPath = join(absDir, folder);
      if (!existsSync(dirPath)) continue;

      const files = await readdir(dirPath);
      for (const f of files) {
        if (!f.toLowerCase().endsWith(ext.toLowerCase())) continue;
        const rel = `${folder}/${f}`.toLowerCase();
        if (!expectedFiles.has(rel)) {
          await rm(join(dirPath, f), { force: true });
          deleted++;
        }
      }
    }
    return deleted;
  }

  /** 规范化代码用于比较（去除首尾空白和一致换行） */
  private normalizeCode(code: string): string {
    return code.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
  }
}
