import * as fs from 'fs';
import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import {
    type ExportMode,
    type ModuleInfo,
    normalizeExportMode,
    readWorkbookRepoConfig,
    relativeNameForModule,
    sanitizeFileName,
} from './moduleExport';

export type ModuleSyncDirection = 'export' | 'import';
export type ModuleSyncItemStatus =
    | 'will-write'
    | 'unchanged'
    | 'will-create'
    | 'will-update'
    | 'will-remove'
    | 'skipping-export'
    | 'skipping-import'
    | 'read-error';

export interface ModuleSyncDiffLine {
    leftNumber?: number;
    rightNumber?: number;
    left: string;
    right: string;
    kind: 'equal' | 'changed' | 'added' | 'removed';
}

export interface ModuleSyncPlanItem {
    id: string;
    moduleName: string;
    moduleType: string;
    documentType?: string;
    relativeName: string;
    sourcePath?: string;
    targetPath?: string;
    status: ModuleSyncItemStatus;
    checked: boolean;
    selectable: boolean;
    warning?: string;
    detail?: string;
    existsInWorkbook: boolean;
    existsInRepo: boolean;
    unsupportedDirectCreation: boolean;
    leftTitle: string;
    rightTitle: string;
    diff: ModuleSyncDiffLine[];
}

export interface ModuleSyncPlan {
    direction: ModuleSyncDirection;
    workbookPath: string;
    folderPath: string;
    exportMode?: ExportMode;
    title: string;
    items: ModuleSyncPlanItem[];
    warnings: string[];
}

interface RepoModuleFile {
    file: string;
    moduleName: string;
    inferredType: string;
    subtype: 'standard' | 'class' | 'document' | 'userform';
    source: string;
    sourcePath: string;
}

const GUID_RE = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g;
const DOCUMENT_CLSIDS = new Set([
    '{00020819-0000-0000-C000-000000000046}',
    '{00020820-0000-0000-C000-000000000046}',
    '{00020821-0000-0000-C000-000000000046}',
]);

export async function buildExportModuleSyncPlan(
    bridge: PythonBridge,
    params: {
        workbookPath: string;
        exportFolder: string;
        exportMode?: ExportMode;
    },
): Promise<ModuleSyncPlan> {
    const exportMode = normalizeExportMode(params.exportMode);
    await fs.promises.mkdir(params.exportFolder, { recursive: true });
    const modules = await bridge.call<ModuleInfo[]>('listModules', { path: params.workbookPath });
    const liveRelativeNames = new Set(modules.map(relativeNameForModule));
    const items = await Promise.all(modules.map(async (mod): Promise<ModuleSyncPlanItem> => {
        const relativeName = relativeNameForModule(mod);
        const targetPath = path.join(params.exportFolder, relativeName);
        const existsInRepo = await fileExists(targetPath);
        const liveSource = await readWorkbookModuleSource(bridge, params.workbookPath, mod.name);
        const repoSource = existsInRepo
            ? await fs.promises.readFile(targetPath, 'utf8')
            : '';
        const equal = existsInRepo && normalizeText(liveSource) === normalizeText(repoSource);
        const skipNew = exportMode === 'replaceExistingOnly' && !existsInRepo;
        const status: ModuleSyncItemStatus = skipNew
            ? 'skipping-export'
            : equal
                ? 'unchanged'
                : existsInRepo
                    ? 'will-write'
                    : 'will-create';

        return {
            id: `export:${mod.name}`,
            moduleName: mod.name,
            moduleType: mod.type,
            documentType: mod.documentType,
            relativeName,
            targetPath,
            status,
            checked: status === 'will-write' || status === 'will-create',
            selectable: !skipNew,
            warning: skipNew
                ? 'Skipping export because mode is replaceExistingOnly and the file does not exist.'
                : undefined,
            detail: statusLabel(status),
            existsInWorkbook: true,
            existsInRepo,
            unsupportedDirectCreation: false,
            leftTitle: `Workbook: ${mod.name}`,
            rightTitle: `Repo: ${relativeName}`,
            diff: buildSideBySideDiff(liveSource, repoSource),
        };
    }));
    const staleItems: ModuleSyncPlanItem[] = [];

    if (exportMode === 'trueUp') {
        const config = await readWorkbookRepoConfig(params.workbookPath);
        for (const relPath of managedFilesFromConfig(config)) {
            if (liveRelativeNames.has(relPath)) {
                continue;
            }

            const stalePath = path.join(params.exportFolder, relPath);
            if (!isPathInside(params.exportFolder, stalePath) || !(await fileExists(stalePath))) {
                continue;
            }

            const repoSource = await fs.promises.readFile(stalePath, 'utf8');
            const moduleName = path.basename(relPath, path.extname(relPath));
            staleItems.push({
                id: `export-stale:${relPath}`,
                moduleName,
                moduleType: 'stale',
                relativeName: relPath,
                targetPath: stalePath,
                status: 'will-remove',
                checked: true,
                selectable: true,
                warning: 'This managed repo file no longer exists as a workbook module and will be removed during true-up.',
                detail: statusLabel('will-remove'),
                existsInWorkbook: false,
                existsInRepo: true,
                unsupportedDirectCreation: false,
                leftTitle: `Repo: ${relPath}`,
                rightTitle: 'Workbook: missing module',
                diff: buildSideBySideDiff(repoSource, ''),
            });
        }
    }

    return {
        direction: 'export',
        workbookPath: params.workbookPath,
        folderPath: params.exportFolder,
        exportMode,
        title: `Export modules: ${path.basename(params.workbookPath)}`,
        items: [...items, ...staleItems].sort(compareSyncItems),
        warnings: [],
    };
}

export async function buildImportModuleSyncPlan(
    bridge: PythonBridge,
    params: {
        workbookPath: string;
        importFolder: string;
    },
): Promise<ModuleSyncPlan> {
    const liveModules = await bridge.call<ModuleInfo[]>('listModules', { path: params.workbookPath });
    const liveByName = new Map(liveModules.map((mod) => [mod.name.toLowerCase(), mod]));
    const entries = (await fs.promises.readdir(params.importFolder))
        .filter((entry) => /\.(bas|cls|frm)$/i.test(entry))
        .sort((a, b) => a.localeCompare(b));
    const repoFiles = await Promise.all(entries.map((entry) =>
        readRepoModuleFile(params.importFolder, entry),
    ));

    const items = await Promise.all(repoFiles.map(async (repo): Promise<ModuleSyncPlanItem> => {
        const live = liveByName.get(repo.moduleName.toLowerCase());
        const existsInWorkbook = Boolean(live);
        const moduleType = live?.type ?? repo.inferredType;
        const unsupportedDirectCreation =
            !existsInWorkbook && (repo.subtype === 'document' || repo.subtype === 'userform');
        const workbookSource = existsInWorkbook
            ? await readWorkbookModuleSource(bridge, params.workbookPath, repo.moduleName)
            : '';
        const equal = existsInWorkbook && normalizeText(repo.source) === normalizeText(workbookSource);
        const status: ModuleSyncItemStatus = unsupportedDirectCreation
            ? 'skipping-import'
            : equal
                ? 'unchanged'
                : existsInWorkbook
                    ? 'will-update'
                    : 'will-create';
        const warning = unsupportedDirectCreation
            ? `${moduleKindLabel(repo.subtype)} modules cannot be created directly. XLIDE will skip this import unless the module already exists in the workbook.`
            : documentLike(moduleType)
                ? `${moduleKindLabel(moduleType)} code can be updated because the module already exists in the workbook.`
                : undefined;

        return {
            id: `import:${repo.file}`,
            moduleName: repo.moduleName,
            moduleType,
            documentType: live?.documentType,
            relativeName: repo.file,
            sourcePath: repo.sourcePath,
            status,
            checked: status === 'will-update' || status === 'will-create',
            selectable: true,
            warning,
            detail: statusLabel(status),
            existsInWorkbook,
            existsInRepo: true,
            unsupportedDirectCreation,
            leftTitle: `Repo: ${repo.file}`,
            rightTitle: `Workbook: ${repo.moduleName}`,
            diff: buildSideBySideDiff(repo.source, workbookSource),
        };
    }));

    return {
        direction: 'import',
        workbookPath: params.workbookPath,
        folderPath: params.importFolder,
        title: `Import modules: ${path.basename(params.workbookPath)}`,
        items: items.sort(compareSyncItems),
        warnings: items
            .filter((item) => item.unsupportedDirectCreation)
            .map((item) => `${item.moduleName}: skipping import unless the module already exists in the workbook.`),
    };
}

export function buildSideBySideDiff(leftText: string, rightText: string): ModuleSyncDiffLine[] {
    const left = splitLines(leftText);
    const right = splitLines(rightText);
    const table = lcsTable(left, right);
    const out: ModuleSyncDiffLine[] = [];
    let i = 0;
    let j = 0;
    while (i < left.length && j < right.length) {
        if (left[i] === right[j]) {
            out.push({
                leftNumber: i + 1,
                rightNumber: j + 1,
                left: left[i],
                right: right[j],
                kind: 'equal',
            });
            i++;
            j++;
        } else if (table[i + 1][j] >= table[i][j + 1]) {
            if (j < right.length && table[i + 1][j] === table[i][j + 1]) {
                out.push({
                    leftNumber: i + 1,
                    rightNumber: j + 1,
                    left: left[i],
                    right: right[j],
                    kind: 'changed',
                });
                i++;
                j++;
            } else {
                out.push({
                    leftNumber: i + 1,
                    left: left[i],
                    right: '',
                    kind: 'removed',
                });
                i++;
            }
        } else {
            out.push({
                rightNumber: j + 1,
                left: '',
                right: right[j],
                kind: 'added',
            });
            j++;
        }
    }
    while (i < left.length) {
        out.push({ leftNumber: i + 1, left: left[i], right: '', kind: 'removed' });
        i++;
    }
    while (j < right.length) {
        out.push({ rightNumber: j + 1, left: '', right: right[j], kind: 'added' });
        j++;
    }
    return out.length > 0 ? out : [{ left: '', right: '', kind: 'equal' }];
}

export function statusLabel(status: ModuleSyncItemStatus): string {
    switch (status) {
        case 'will-write':
            return 'Will overwrite repo file';
        case 'will-create':
            return 'Will create';
        case 'will-update':
            return 'Will update workbook module';
        case 'will-remove':
            return 'Will remove stale repo file';
        case 'unchanged':
            return 'Unchanged';
        case 'skipping-export':
            return 'Skipping export';
        case 'skipping-import':
            return 'Skipping import';
        case 'read-error':
            return 'Read error';
    }
}

async function readWorkbookModuleSource(
    bridge: PythonBridge,
    workbookPath: string,
    moduleName: string,
): Promise<string> {
    const result = await bridge.call<{ source: string }>('readModule', {
        path: workbookPath,
        module: moduleName,
        full: true,
    });
    return result.source;
}

async function readRepoModuleFile(folder: string, file: string): Promise<RepoModuleFile> {
    const sourcePath = path.join(folder, file);
    const source = await fs.promises.readFile(sourcePath, 'utf8');
    const ext = path.extname(file).toLowerCase();
    const moduleName = sanitizeFileName(path.basename(file, ext)) || path.basename(file, ext);
    const subtype = ext === '.bas'
        ? 'standard'
        : ext === '.frm'
            ? 'userform'
            : detectClsSubtype(source);
    return {
        file,
        moduleName,
        inferredType: subtype === 'document' ? 'document' : subtype === 'userform' ? 'userform' : subtype,
        subtype,
        source,
        sourcePath,
    };
}

function detectClsSubtype(source: string): 'class' | 'document' | 'userform' {
    const head = source.slice(0, 2000);
    const vbBaseMatch = head.match(/Attribute\s+VB_Base\s*=\s*"([^"]*)"/i);
    if (vbBaseMatch) {
        const guids = vbBaseMatch[1].match(GUID_RE) ?? [];
        if (guids.length >= 2) {
            return 'userform';
        }
        if (guids.some((guid) => DOCUMENT_CLSIDS.has(guid.toUpperCase()))) {
            return 'document';
        }
        return 'class';
    }
    if (/Attribute\s+VB_PredeclaredId\s*=\s*True/i.test(head)) {
        return 'document';
    }
    return 'class';
}

function splitLines(text: string): string[] {
    if (text.length === 0) {
        return [];
    }
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
}

function normalizeText(text: string): string {
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

async function fileExists(filePath: string): Promise<boolean> {
    try {
        await fs.promises.access(filePath, fs.constants.F_OK);
        return true;
    } catch {
        return false;
    }
}

function isPathInside(baseDir: string, targetPath: string): boolean {
    const base = path.resolve(baseDir);
    const target = path.resolve(targetPath);
    return target === base || target.startsWith(base + path.sep);
}

function managedFilesFromConfig(config: { managedFiles?: unknown }): string[] {
    return Array.isArray(config.managedFiles)
        ? config.managedFiles.filter((item): item is string => typeof item === 'string')
        : [];
}

function lcsTable(left: readonly string[], right: readonly string[]): number[][] {
    const table = Array.from({ length: left.length + 1 }, () => Array(right.length + 1).fill(0));
    for (let i = left.length - 1; i >= 0; i--) {
        for (let j = right.length - 1; j >= 0; j--) {
            table[i][j] = left[i] === right[j]
                ? table[i + 1][j + 1] + 1
                : Math.max(table[i + 1][j], table[i][j + 1]);
        }
    }
    return table;
}

function compareSyncItems(a: ModuleSyncPlanItem, b: ModuleSyncPlanItem): number {
    if (a.status !== b.status) {
        return statusSort(a.status) - statusSort(b.status);
    }
    return a.moduleName.localeCompare(b.moduleName);
}

function statusSort(status: ModuleSyncItemStatus): number {
    switch (status) {
        case 'will-write':
        case 'will-update':
        case 'will-create':
        case 'will-remove':
            return 0;
        case 'skipping-import':
        case 'skipping-export':
            return 1;
        case 'read-error':
            return 2;
        case 'unchanged':
            return 3;
    }
}

function documentLike(moduleType: string): boolean {
    return moduleType === 'document' || moduleType === 'userform';
}

function moduleKindLabel(kind: string): string {
    switch (kind) {
        case 'document':
            return 'Worksheet/ThisWorkbook';
        case 'userform':
            return 'UserForm';
        default:
            return kind;
    }
}
