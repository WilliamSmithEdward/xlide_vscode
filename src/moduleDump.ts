import * as fs from 'fs';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';

interface ModuleInfo {
    name: string;
    type: string;
}

type ExportMode = 'trueUp' | 'replaceExistingOnly';

interface WorkbookRepoConfig {
    // Preferred keys
    exportFolder?: string;
    exportMode?: ExportMode;
    // Legacy keys kept for backward compatibility
    dumpFolder?: string;
    dumpMode?: ExportMode;
    managedFiles?: string[];
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
    skippedNewCount: number;
    removedCount: number;
    totalModules: number;
    configPath: string;
}

function configPathForWorkbook(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.extension.repo.json`);
}

function extensionForModuleType(moduleType: string): string {
    return moduleType === 'standard' ? 'bas' : 'cls';
}

function sanitizeFileName(name: string): string {
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/g, '');
}

function getManagedFiles(config: WorkbookRepoConfig): string[] {
    if (!Array.isArray(config.managedFiles)) {
        return [];
    }
    return config.managedFiles.filter((v) => typeof v === 'string');
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

function normalizeExportMode(mode: ExportMode | undefined): ExportMode {
    return mode === 'replaceExistingOnly' ? 'replaceExistingOnly' : 'trueUp';
}

function getConfiguredFolder(config: WorkbookRepoConfig): string | undefined {
    return config.exportFolder ?? config.dumpFolder;
}

function getConfiguredMode(config: WorkbookRepoConfig): ExportMode {
    return normalizeExportMode(config.exportMode ?? config.dumpMode);
}

async function readWorkbookRepoConfig(filePath: string): Promise<WorkbookRepoConfig> {
    const configPath = configPathForWorkbook(filePath);
    try {
        const raw = await fs.promises.readFile(configPath, 'utf8');
        const parsed = JSON.parse(raw) as WorkbookRepoConfig;
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch {
        return {};
    }
}

async function writeWorkbookRepoConfig(filePath: string, config: WorkbookRepoConfig): Promise<void> {
    const configPath = configPathForWorkbook(filePath);
    await fs.promises.writeFile(configPath, `${JSON.stringify(config, null, 2)}\n`, 'utf8');
}

async function setWorkbookExportMode(filePath: string, mode: ExportMode): Promise<WorkbookRepoConfig> {
    const existing = await readWorkbookRepoConfig(filePath);
    const updated: WorkbookRepoConfig = {
        exportFolder: getConfiguredFolder(existing),
        exportMode: normalizeExportMode(mode),
        managedFiles: getManagedFiles(existing),
    };
    await writeWorkbookRepoConfig(filePath, updated);
    return updated;
}

async function exportWorkbookModules(
    bridge: PythonBridge,
    params: ExportModulesParams,
): Promise<ExportModulesResult> {
    const existingConfig = await readWorkbookRepoConfig(params.filePath);
    const exportFolder = params.exportFolder ?? getConfiguredFolder(existingConfig);
    if (!exportFolder) {
        throw new Error('No export folder configured. Choose a folder first or provide exportFolder.');
    }

    const exportMode = normalizeExportMode(params.exportMode ?? getConfiguredMode(existingConfig));
    await fs.promises.mkdir(exportFolder, { recursive: true });

    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: params.filePath });
    const managedNow = new Set<string>();
    let writtenCount = 0;
    let skippedNewCount = 0;

    for (const mod of modules) {
        const sourceResult = await bridge.call<{ source: string }>('readModule', {
            path: params.filePath,
            module: mod.name,
            full: true,   // include VBA attribute headers so exported files round-trip cleanly
        });

        const safeName = sanitizeFileName(mod.name) || mod.name;
        const ext = extensionForModuleType(mod.type);
        const relativeName = `${safeName}.${ext}`;
        const outPath = path.join(exportFolder, relativeName);

        if (exportMode === 'replaceExistingOnly' && !(await fileExists(outPath))) {
            skippedNewCount++;
            continue;
        }

        await fs.promises.writeFile(outPath, sourceResult.source, 'utf8');
        managedNow.add(relativeName);
        writtenCount++;
    }

    let removedCount = 0;
    if (exportMode === 'trueUp') {
        const previouslyManaged = getManagedFiles(existingConfig);
        for (const relPath of previouslyManaged) {
            if (managedNow.has(relPath)) {
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
            removedCount++;
        }
    }

    await writeWorkbookRepoConfig(params.filePath, {
        exportFolder,
        exportMode,
        managedFiles: Array.from(managedNow).sort(),
    });

    return {
        filePath: params.filePath,
        exportFolder,
        exportMode,
        writtenCount,
        skippedNewCount,
        removedCount,
        totalModules: modules.length,
        configPath: configPathForWorkbook(params.filePath),
    };
}

export {
    type ExportMode,
    type WorkbookRepoConfig,
    type ExportModulesParams,
    type ExportModulesResult,
    configPathForWorkbook,
    normalizeExportMode,
    readWorkbookRepoConfig,
    writeWorkbookRepoConfig,
    setWorkbookExportMode,
    exportWorkbookModules,
};
