import * as fs from 'fs';
import * as path from 'path';
import { WorkbookEngine } from './workbookEngine';
import {
    normalizeExportMode,
    type ExportMode,
} from './workbookSettings';
import {
    effectiveWorkbookModuleSyncSettings,
    updateWorkbookModuleSyncSettings,
} from './workbookModuleSyncSettings';
import { measurePerformance } from './performanceTrace';
import { fileExists, isPathInside } from './util/fs';
import { createKeyedAsyncLock } from './util/keyedAsyncLock';

const exportFolderLock = createKeyedAsyncLock();

function exportFolderKey(folder: string): string {
    const resolved = path.resolve(folder);
    return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

// Windows and (default) macOS filesystems are case-insensitive, so a live module
// and its on-disk export file can legitimately differ only by case. Compare
// export filenames case-insensitively on those platforms; otherwise trueUp would
// classify the just-written export as "stale" and delete it (data loss).
function caseNormalizedRelName(name: string): string {
    return process.platform === 'win32' || process.platform === 'darwin'
        ? name.toLowerCase()
        : name;
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
    // A form is a .frm: that is what the VBE's own exporter writes, and the
    // only name under which its designer sidecar (.frx) can travel with it.
    if (moduleType === 'userform') {
        return 'frm';
    }
    return moduleType === 'standard' ? 'bas' : 'cls';
}

function sanitizeFileName(name: string): string {
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/g, '');
}

function isRootVbaModuleFileName(value: string): boolean {
    // A .frx is NOT a module file: it is the binary designer sidecar the VBE's
    // importer reads beside a .frm, so it is neither listed nor ever removed
    // as a stale module.
    return !/[\\/]/.test(value) && /\.(bas|cls|frm)$/i.test(value);
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

// Stale repo files for trueUp export: root .bas/.cls/.frm files with no live
// workbook module, guarded against escaping the export folder. Shared by the
// sync-plan preview and the export action so both agree on what gets removed.
async function computeStaleExportFiles(
    exportFolder: string,
    liveRelativeNames: ReadonlySet<string>,
): Promise<string[]> {
    const stale: string[] = [];
    const liveNormalized = new Set([...liveRelativeNames].map(caseNormalizedRelName));
    for (const relPath of await listRootVbaModuleFiles(exportFolder)) {
        if (liveNormalized.has(caseNormalizedRelName(relPath))) {
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
    bridge: WorkbookEngine,
    filePath: string,
    mod: ModuleInfo,
    exportFolder: string,
    source?: string,
): Promise<{ relativeName: string; sidecarRelativeName?: string; written: boolean }> {
    const relativeName = relativeNameForModule(mod);
    const outPath = path.join(exportFolder, relativeName);

    // A form exports as a pair: the composed .frm and the .frx sidecar the
    // designer travels in. When the designer cannot be read, the .frm is the
    // module text alone and no sidecar is written - never a guessed one.
    if (mod.type === 'userform') {
        try {
            const pair = await bridge.call<{ frm: string; frx: { data?: number[]; type?: string } | Buffer }>(
                'readFormExport',
                { path: filePath, module: mod.name },
            );
            const frxBytes = Buffer.isBuffer(pair.frx)
                ? pair.frx
                : Buffer.from((pair.frx as { data?: number[] }).data ?? []);
            const sidecarRelativeName = relativeName.replace(/\.frm$/i, '.frx');
            await fs.promises.writeFile(outPath, pair.frm, 'utf8');
            await fs.promises.writeFile(path.join(exportFolder, sidecarRelativeName), frxBytes);
            return { relativeName, sidecarRelativeName, written: true };
        } catch {
            // Fall through to the plain module text.
        }
    }

    const moduleSource = source ?? await readFullModuleSource(bridge, filePath, mod.name);
    await fs.promises.writeFile(outPath, moduleSource, 'utf8');
    return { relativeName, written: true };
}

async function readFullModuleSource(
    bridge: WorkbookEngine,
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

// Full-source batch read in a single workbook open.
async function loadWorkbookModulesWithSources(
    bridge: WorkbookEngine,
    filePath: string,
): Promise<WorkbookModulesWithSources> {
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
    // A form's file is the composed .frm - designer block plus code - so the
    // sync plan compares, and the exporter writes, the same text. A form whose
    // designer cannot be read keeps its raw module text, as before.
    for (const mod of modules) {
        if (mod.type !== 'userform') {
            continue;
        }
        try {
            const pair = await bridge.call<{ frm: string }>(
                'readFormExport',
                { path: filePath, module: mod.name },
            );
            sources.set(mod.name.toLowerCase(), pair.frm);
        } catch {
            // No designer storage: the raw source already in the map stands.
        }
    }
    return {
        modules: modules.map(({ name, type, documentType }) => ({ name, type, documentType })),
        sourceFor: async (moduleName) =>
            sources.get(moduleName.toLowerCase()) ??
                readFullModuleSource(bridge, filePath, moduleName),
    };
}

async function exportWorkbookModule(
    bridge: WorkbookEngine,
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
    bridge: WorkbookEngine,
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
        if (exported.sidecarRelativeName) {
            writtenFiles.push(exported.sidecarRelativeName);
        }
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
