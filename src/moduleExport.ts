import * as fs from 'fs';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import {
    normalizeExportMode,
    type ExportMode,
} from './workbookSettings';
import {
    effectiveWorkbookModuleSyncSettings,
    updateWorkbookModuleSyncSettings,
} from './workbookModuleSyncSettings';
import { measurePerformance } from './performanceTrace';

interface ModuleInfo {
    name: string;
    type: string;
    documentType?: string;
}

interface ExportModulesParams {
    filePath: string;
    exportFolder?: string;
    exportMode?: ExportMode;
}

interface ExportModulesResult {
    filePath: string;
    exportFolder: string;
    exportMode: ExportMode;
    writtenCount: number;
    removedCount: number;
    writtenFiles: string[];
    removedFiles: string[];
    totalModules: number;
    configPath: string;
}

interface ExportModuleParams {
    filePath: string;
    moduleName: string;
    exportFolder?: string;
    exportMode?: ExportMode;
}

interface ExportModuleResult {
    filePath: string;
    moduleName: string;
    moduleType: string;
    exportFolder: string;
    exportMode: ExportMode;
    relativeName: string;
    written: boolean;
    writtenFiles: string[];
    configPath: string;
}

function extensionForModuleType(moduleType: string): string {
    return moduleType === 'standard' ? 'bas' : 'cls';
}

function sanitizeFileName(name: string): string {
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/g, '');
}

function isRootVbaModuleFileName(value: string): boolean {
    return !/[\\/]/.test(value) && /\.(bas|cls)$/i.test(value);
}

async function listRootVbaModuleFiles(folder: string): Promise<string[]> {
    let entries: fs.Dirent[];
    try {
        entries = await fs.promises.readdir(folder, { withFileTypes: true });
    } catch {
        return [];
    }
    return entries
        .filter((entry) => entry.isFile() && isRootVbaModuleFileName(entry.name))
        .map((entry) => entry.name)
        .sort((left, right) => left.localeCompare(right));
}

function isPathInside(baseDir: string, targetPath: string): boolean {
    const base = path.resolve(baseDir);
    const target = path.resolve(targetPath);
    return target === base || target.startsWith(base + path.sep);
}

async function fileExists(filePath: string): Promise<boolean> {
    try {
        await fs.promises.access(filePath, fs.constants.F_OK);
        return true;
    } catch {
        return false;
    }
}

function relativeNameForModule(mod: ModuleInfo): string {
    const safeName = sanitizeFileName(mod.name) || mod.name;
    const ext = extensionForModuleType(mod.type);
    return `${safeName}.${ext}`;
}

async function exportModuleFile(
    bridge: PythonBridge,
    filePath: string,
    mod: ModuleInfo,
    exportFolder: string,
): Promise<{ relativeName: string; written: boolean }> {
    const relativeName = relativeNameForModule(mod);
    const outPath = path.join(exportFolder, relativeName);

    const sourceResult = await bridge.call<{ source: string }>('readModule', {
        path: filePath,
        module: mod.name,
        full: true,   // include VBA attribute headers so exported files round-trip cleanly
    });
    await fs.promises.writeFile(outPath, sourceResult.source, 'utf8');
    return { relativeName, written: true };
}

async function exportWorkbookModule(
    bridge: PythonBridge,
    params: ExportModuleParams,
): Promise<ExportModuleResult> {
    return measurePerformance('moduleExport.single', params.moduleName, async () => {
    const existingSettings = await effectiveWorkbookModuleSyncSettings(params.filePath);
    const exportFolder = params.exportFolder ?? existingSettings.folderPath;
    if (!exportFolder) {
        throw new Error('No export folder configured. Choose a folder first or provide exportFolder.');
    }

    const exportMode = normalizeExportMode(params.exportMode ?? existingSettings.exportMode);
    await fs.promises.mkdir(exportFolder, { recursive: true });

    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: params.filePath });
    const mod = modules.find(
        (candidate) => candidate.name.toLowerCase() === params.moduleName.toLowerCase(),
    );
    if (!mod) {
        throw new Error(`Module "${params.moduleName}" was not found in the workbook.`);
    }

    const exported = await exportModuleFile(bridge, params.filePath, mod, exportFolder);

    const updatedSettings = await updateWorkbookModuleSyncSettings(params.filePath, {
        folderPath: exportFolder,
        exportMode,
    });

    return {
        filePath: params.filePath,
        moduleName: mod.name,
        moduleType: mod.type,
        exportFolder,
        exportMode,
        relativeName: exported.relativeName,
        written: exported.written,
        writtenFiles: exported.written ? [exported.relativeName] : [],
        configPath: updatedSettings.settingsPath,
    };
    });
}

async function exportWorkbookModules(
    bridge: PythonBridge,
    params: ExportModulesParams,
): Promise<ExportModulesResult> {
    return measurePerformance('moduleExport.workbook', path.basename(params.filePath), async () => {
    const existingSettings = await effectiveWorkbookModuleSyncSettings(params.filePath);
    const exportFolder = params.exportFolder ?? existingSettings.folderPath;
    if (!exportFolder) {
        throw new Error('No export folder configured. Choose a folder first or provide exportFolder.');
    }

    const exportMode = normalizeExportMode(params.exportMode ?? existingSettings.exportMode);
    await fs.promises.mkdir(exportFolder, { recursive: true });

    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: params.filePath });
    const liveRelativeNames = new Set<string>();
    const writtenFiles: string[] = [];
    const removedFiles: string[] = [];

    for (const mod of modules) {
        const exported = await exportModuleFile(
            bridge,
            params.filePath,
            mod,
            exportFolder,
        );
        liveRelativeNames.add(exported.relativeName);
        writtenFiles.push(exported.relativeName);
    }

    if (exportMode === 'trueUp') {
        for (const relPath of await listRootVbaModuleFiles(exportFolder)) {
            if (liveRelativeNames.has(relPath)) {
                continue;
            }

            const stalePath = path.join(exportFolder, relPath);
            if (!isPathInside(exportFolder, stalePath)) {
                continue;
            }
            if (!(await fileExists(stalePath))) {
                continue;
            }

            await fs.promises.unlink(stalePath);
            removedFiles.push(relPath);
        }
    }

    const updatedSettings = await updateWorkbookModuleSyncSettings(params.filePath, {
        folderPath: exportFolder,
        exportMode,
    });

    return {
        filePath: params.filePath,
        exportFolder,
        exportMode,
        writtenCount: writtenFiles.length,
        removedCount: removedFiles.length,
        writtenFiles,
        removedFiles,
        totalModules: modules.length,
        configPath: updatedSettings.settingsPath,
    };
    });
}

export {
    type ModuleInfo,
    type ExportModulesParams,
    type ExportModulesResult,
    type ExportModuleParams,
    type ExportModuleResult,
    extensionForModuleType,
    listRootVbaModuleFiles,
    relativeNameForModule,
    sanitizeFileName,
    exportWorkbookModule,
    exportWorkbookModules,
};
