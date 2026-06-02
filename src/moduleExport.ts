import * as fs from 'fs';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import type { ImportMode } from './moduleSyncPlan';

interface ModuleInfo {
    name: string;
    type: string;
    documentType?: string;
}

type ExportMode = 'exportAll' | 'trueUp';
type LegacyExportMode = 'replaceExistingOnly';

interface WorkbookRepoConfig {
    exportFolder?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
}

type WorkbookRepoConfigInput = Omit<WorkbookRepoConfig, 'exportMode'> & {
    exportMode?: ExportMode | LegacyExportMode;
};

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

function configPathForWorkbook(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.repo.json`);
}

function legacyConfigPathForWorkbook(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.extension.repo.json`);
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

function normalizeExportMode(mode: ExportMode | LegacyExportMode | unknown): ExportMode {
    return mode === 'trueUp' ? 'trueUp' : 'exportAll';
}

function normalizeImportModeValue(mode: unknown): ImportMode | undefined {
    return mode === 'updateOnly' || mode === 'trueUpStandardClass' ? mode : undefined;
}

function normalizeWorkbookRepoConfig(config: {
    exportFolder?: unknown;
    exportMode?: unknown;
    importMode?: unknown;
}): WorkbookRepoConfig {
    const normalized: WorkbookRepoConfig = {
        exportMode: normalizeExportMode(config.exportMode),
    };
    if (typeof config.exportFolder === 'string') {
        normalized.exportFolder = config.exportFolder;
    }
    const importMode = normalizeImportModeValue(config.importMode);
    if (importMode) {
        normalized.importMode = importMode;
    }
    return normalized;
}

async function readWorkbookRepoConfig(filePath: string): Promise<WorkbookRepoConfig> {
    for (const configPath of [configPathForWorkbook(filePath), legacyConfigPathForWorkbook(filePath)]) {
        try {
            const raw = await fs.promises.readFile(configPath, 'utf8');
            const parsed = JSON.parse(raw);
            return parsed && typeof parsed === 'object' ? normalizeWorkbookRepoConfig(parsed) : {};
        } catch {
            // Try the next supported sidecar path.
        }
    }
    return {};
}

async function writeWorkbookRepoConfig(
    filePath: string,
    config: WorkbookRepoConfigInput,
): Promise<void> {
    const configPath = configPathForWorkbook(filePath);
    await fs.promises.writeFile(
        configPath,
        `${JSON.stringify(normalizeWorkbookRepoConfig(config), null, 2)}\n`,
        'utf8',
    );
}

async function setWorkbookExportMode(filePath: string, mode: ExportMode): Promise<WorkbookRepoConfig> {
    const existing = await readWorkbookRepoConfig(filePath);
    const updated: WorkbookRepoConfig = {
        exportFolder: existing.exportFolder,
        exportMode: normalizeExportMode(mode),
        importMode: existing.importMode,
    };
    await writeWorkbookRepoConfig(filePath, updated);
    return updated;
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
    const existingConfig = await readWorkbookRepoConfig(params.filePath);
    const exportFolder = params.exportFolder ?? existingConfig.exportFolder;
    if (!exportFolder) {
        throw new Error('No export folder configured. Choose a folder first or provide exportFolder.');
    }

    const exportMode = normalizeExportMode(params.exportMode ?? existingConfig.exportMode);
    await fs.promises.mkdir(exportFolder, { recursive: true });

    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: params.filePath });
    const mod = modules.find(
        (candidate) => candidate.name.toLowerCase() === params.moduleName.toLowerCase(),
    );
    if (!mod) {
        throw new Error(`Module "${params.moduleName}" was not found in the workbook.`);
    }

    const exported = await exportModuleFile(bridge, params.filePath, mod, exportFolder);

    await writeWorkbookRepoConfig(params.filePath, {
        exportFolder,
        exportMode,
        importMode: existingConfig.importMode,
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
        configPath: configPathForWorkbook(params.filePath),
    };
}

async function exportWorkbookModules(
    bridge: PythonBridge,
    params: ExportModulesParams,
): Promise<ExportModulesResult> {
    const existingConfig = await readWorkbookRepoConfig(params.filePath);
    const exportFolder = params.exportFolder ?? existingConfig.exportFolder;
    if (!exportFolder) {
        throw new Error('No export folder configured. Choose a folder first or provide exportFolder.');
    }

    const exportMode = normalizeExportMode(params.exportMode ?? existingConfig.exportMode);
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

    await writeWorkbookRepoConfig(params.filePath, {
        exportFolder,
        exportMode,
        importMode: existingConfig.importMode,
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
        configPath: configPathForWorkbook(params.filePath),
    };
}

export {
    type ExportMode,
    type ModuleInfo,
    type WorkbookRepoConfig,
    type ExportModulesParams,
    type ExportModulesResult,
    type ExportModuleParams,
    type ExportModuleResult,
    configPathForWorkbook,
    extensionForModuleType,
    legacyConfigPathForWorkbook,
    listRootVbaModuleFiles,
    normalizeExportMode,
    relativeNameForModule,
    readWorkbookRepoConfig,
    sanitizeFileName,
    writeWorkbookRepoConfig,
    setWorkbookExportMode,
    exportWorkbookModule,
    exportWorkbookModules,
};
