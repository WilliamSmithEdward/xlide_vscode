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
import { isReadModulesUnavailable } from './pythonBridgeErrors';
import { fileExists, isPathInside } from './util/fs';
import { createKeyedAsyncLock } from './util/keyedAsyncLock';

const exportFolderLock = createKeyedAsyncLock();

function exportFolderKey(folder: string): string {
    const resolved = path.resolve(folder);
    return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

/**
 * Serializes mutating operations on a given export folder so a single export, a
 * whole-workbook export, a sync-plan apply's deletes, and an import apply's reads
 * cannot interleave their file writes/deletes/reads (which would risk partial
 * files or deleting freshly-written content). Keyed per folder, so different
 * folders run concurrently.
 */
export function withExportFolderLock<T>(folder: string, action: () => Promise<T>): Promise<T> {
    return exportFolderLock(exportFolderKey(folder), action);
}

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

// Stale repo files for trueUp export: root .bas/.cls files with no live
// workbook module, guarded against escaping the export folder. Shared by the
// sync-plan preview and the export action so both agree on what gets removed.
async function computeStaleExportFiles(
    exportFolder: string,
    liveRelativeNames: ReadonlySet<string>,
): Promise<string[]> {
    const stale: string[] = [];
    for (const relPath of await listRootVbaModuleFiles(exportFolder)) {
        if (liveRelativeNames.has(relPath)) {
            continue;
        }
        const stalePath = path.join(exportFolder, relPath);
        if (!isPathInside(exportFolder, stalePath) || !(await fileExists(stalePath))) {
            continue;
        }
        stale.push(relPath);
    }
    return stale;
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
    source?: string,
): Promise<{ relativeName: string; written: boolean }> {
    const relativeName = relativeNameForModule(mod);
    const outPath = path.join(exportFolder, relativeName);

    const moduleSource = source ?? await readFullModuleSource(bridge, filePath, mod.name);
    await fs.promises.writeFile(outPath, moduleSource, 'utf8');
    return { relativeName, written: true };
}

async function readFullModuleSource(
    bridge: PythonBridge,
    filePath: string,
    moduleName: string,
): Promise<string> {
    const result = await bridge.call<{ source: string }>('readModule', {
        path: filePath,
        module: moduleName,
        full: true,   // include VBA attribute headers so exported files round-trip cleanly
    });
    return result.source;
}

interface WorkbookModulesWithSources {
    modules: ModuleInfo[];
    sourceFor: (moduleName: string) => Promise<string>;
}

// Full-source batch read in a single workbook open; falls back to listModules
// plus one readModule per module for backends without readModules.
async function loadWorkbookModulesWithSources(
    bridge: PythonBridge,
    filePath: string,
): Promise<WorkbookModulesWithSources> {
    try {
        const modules = await bridge.call<Array<ModuleInfo & { source?: string }>>(
            'readModules',
            { path: filePath, full: true },
        );
        const sources = new Map<string, string>();
        for (const mod of modules) {
            if (typeof mod.source === 'string') {
                sources.set(mod.name.toLowerCase(), mod.source);
            }
        }
        return {
            modules: modules.map(({ name, type, documentType }) => ({ name, type, documentType })),
            sourceFor: async (moduleName) =>
                sources.get(moduleName.toLowerCase()) ??
                    readFullModuleSource(bridge, filePath, moduleName),
        };
    } catch (err) {
        if (!isReadModulesUnavailable(err)) {
            throw err;
        }
    }
    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: filePath });
    return {
        modules,
        sourceFor: (moduleName) => readFullModuleSource(bridge, filePath, moduleName),
    };
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
    return withExportFolderLock(exportFolder, async () => {
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
    return withExportFolderLock(exportFolder, async () => {
    await fs.promises.mkdir(exportFolder, { recursive: true });

    const { modules, sourceFor } = await loadWorkbookModulesWithSources(bridge, params.filePath);
    const liveRelativeNames = new Set<string>();
    const writtenFiles: string[] = [];
    const removedFiles: string[] = [];

    for (const mod of modules) {
        const exported = await exportModuleFile(
            bridge,
            params.filePath,
            mod,
            exportFolder,
            await sourceFor(mod.name),
        );
        liveRelativeNames.add(exported.relativeName);
        writtenFiles.push(exported.relativeName);
    }

    if (exportMode === 'trueUp') {
        for (const relPath of await computeStaleExportFiles(exportFolder, liveRelativeNames)) {
            await fs.promises.unlink(path.join(exportFolder, relPath));
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
    });
}

export {
    type ModuleInfo,
    type ExportModulesParams,
    type ExportModulesResult,
    type ExportModuleParams,
    type ExportModuleResult,
    type WorkbookModulesWithSources,
    computeStaleExportFiles,
    extensionForModuleType,
    loadWorkbookModulesWithSources,
    relativeNameForModule,
    sanitizeFileName,
    exportWorkbookModule,
    exportWorkbookModules,
};
