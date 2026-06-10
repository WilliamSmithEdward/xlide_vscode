import * as fs from 'fs';
import * as path from 'path';
import type { PythonBridge } from './pythonBridge';
import {
    type ModuleInfo,
    listRootVbaModuleFiles,
    loadWorkbookModulesWithSources,
    relativeNameForModule,
    sanitizeFileName,
} from './moduleExport';
import {
    normalizeExportMode,
    normalizeImportMode,
    type ExportMode,
    type WorkbookSettingSource,
} from './workbookSettings';
import { measurePerformance } from './performanceTrace';
import { isVbaAttributeLine, normalizeEol } from './vbaStructuralAnalysis';

export type ModuleSyncDirection = 'export' | 'import';
export type ImportMode = 'updateOnly' | 'trueUpStandardClass';
export type ModuleSyncFolderSource = 'workbook' | 'session' | 'missing';
export type ModuleSyncModeSource = WorkbookSettingSource | 'session';
export type ModuleSyncItemStatus =
    | 'will-write'
    | 'unchanged'
    | 'will-create'
    | 'will-update'
    | 'will-remove'
    | 'skipping-import'
    | 'read-error';

export interface ModuleSyncDiffLine {
    leftNumber?: number;
    rightNumber?: number;
    left: string;
    right: string;
    kind: 'equal' | 'changed' | 'added' | 'removed';
}

interface SideBySideDiffOptions {
    leftOnlyKind?: ModuleSyncDiffLine['kind'];
    rightOnlyKind?: ModuleSyncDiffLine['kind'];
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
    leftCode: string;
    rightCode: string;
    leftRawCode: string;
    rightRawCode: string;
    diff: ModuleSyncDiffLine[];
    diffWithHeaders: ModuleSyncDiffLine[];
}

export interface ModuleSyncPlan {
    direction: ModuleSyncDirection;
    workbookPath: string;
    folderPath: string;
    folderPathSource?: ModuleSyncFolderSource;
    exportMode?: ExportMode;
    exportModeSource?: ModuleSyncModeSource;
    importMode?: ImportMode;
    importModeSource?: ModuleSyncModeSource;
    settingsPath?: string;
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
const VB_BASE_RE = /^\s*Attribute\s+VB_Base\s*=\s*"([^"]*)"/im;
const DOCUMENT_MODULE_NAME_RE = /^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha)\d*$/i;

export async function buildExportModuleSyncPlan(
    bridge: PythonBridge,
    params: {
        workbookPath: string;
        exportFolder: string;
        exportMode?: ExportMode;
        folderPathSource?: ModuleSyncFolderSource;
        exportModeSource?: ModuleSyncModeSource;
        settingsPath?: string;
    },
): Promise<ModuleSyncPlan> {
    return measurePerformance('moduleSync.buildExportPlan', path.basename(params.workbookPath), async () => {
    const exportMode = normalizeExportMode(params.exportMode);
    await fs.promises.mkdir(params.exportFolder, { recursive: true });
    const { modules, sourceFor } = await loadWorkbookModulesWithSources(bridge, params.workbookPath);
    const liveRelativeNames = new Set(modules.map(relativeNameForModule));
    const items = await Promise.all(modules.map(async (mod): Promise<ModuleSyncPlanItem> => {
        const relativeName = relativeNameForModule(mod);
        const targetPath = path.join(params.exportFolder, relativeName);
        const existsInRepo = await fileExists(targetPath);
        const liveSource = await sourceFor(mod.name);
        const repoSource = existsInRepo
            ? await fs.promises.readFile(targetPath, 'utf8')
            : '';
        const liveDisplaySource = editorPreviewSource(liveSource);
        const repoDisplaySource = editorPreviewSource(repoSource);
        const equal = existsInRepo && normalizeEol(liveSource) === normalizeEol(repoSource);
        const status: ModuleSyncItemStatus = equal
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
            selectable: true,
            warning: undefined,
            detail: statusLabel(status),
            existsInWorkbook: true,
            existsInRepo,
            unsupportedDirectCreation: false,
            leftTitle: `Workbook: ${mod.name}`,
            rightTitle: exportRepoTitle(relativeName, status),
            leftCode: liveDisplaySource,
            rightCode: repoDisplaySource,
            leftRawCode: liveSource,
            rightRawCode: repoSource,
            diff: buildSideBySideDiff(liveDisplaySource, repoDisplaySource, writeDiffTones()),
            diffWithHeaders: buildSideBySideDiff(liveSource, repoSource, writeDiffTones()),
        };
    }));
    const staleItems: ModuleSyncPlanItem[] = [];

    if (exportMode === 'trueUp') {
        for (const relPath of await listRootVbaModuleFiles(params.exportFolder)) {
            if (liveRelativeNames.has(relPath)) {
                continue;
            }

            const stalePath = path.join(params.exportFolder, relPath);
            if (!isPathInside(params.exportFolder, stalePath) || !(await fileExists(stalePath))) {
                continue;
            }

            const repoSource = await fs.promises.readFile(stalePath, 'utf8');
            const repoDisplaySource = editorPreviewSource(repoSource);
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
                warning: 'This stale .bas/.cls repo module file no longer exists as a workbook module and will be removed during mirror export.',
                detail: statusLabel('will-remove'),
                existsInWorkbook: false,
                existsInRepo: true,
                unsupportedDirectCreation: false,
                leftTitle: `Repo: ${relPath} (will remove)`,
                rightTitle: 'Workbook: missing module',
                leftCode: repoDisplaySource,
                rightCode: '',
                leftRawCode: repoSource,
                rightRawCode: '',
                diff: buildSideBySideDiff(repoDisplaySource, ''),
                diffWithHeaders: buildSideBySideDiff(repoSource, ''),
            });
        }
    }

    return {
        direction: 'export',
        workbookPath: params.workbookPath,
        folderPath: params.exportFolder,
        folderPathSource: params.folderPathSource,
        exportMode,
        exportModeSource: params.exportModeSource,
        settingsPath: params.settingsPath,
        title: `Export modules: ${path.basename(params.workbookPath)}`,
        items: [...items, ...staleItems].sort(compareSyncItems),
        warnings: [],
    };
    });
}

export async function buildImportModuleSyncPlan(
    bridge: PythonBridge,
    params: {
        workbookPath: string;
        importFolder: string;
        importMode?: ImportMode;
        folderPathSource?: ModuleSyncFolderSource;
        importModeSource?: ModuleSyncModeSource;
        settingsPath?: string;
    },
): Promise<ModuleSyncPlan> {
    return measurePerformance('moduleSync.buildImportPlan', path.basename(params.workbookPath), async () => {
    const importMode = normalizeImportMode(params.importMode);
    const { modules: liveModules, sourceFor } = await loadWorkbookModulesWithSources(bridge, params.workbookPath);
    const liveByName = new Map(liveModules.map((mod) => [mod.name.toLowerCase(), mod]));
    const entries = (await fs.promises.readdir(params.importFolder))
        .filter((entry) => /\.(bas|cls)$/i.test(entry))
        .sort((a, b) => a.localeCompare(b));
    const repoFiles = await Promise.all(entries.map((entry) =>
        readRepoModuleFile(params.importFolder, entry),
    ));
    const repoModuleNames = new Set(repoFiles.map((repo) => repo.moduleName.toLowerCase()));

    const items = await Promise.all(repoFiles.map(async (repo): Promise<ModuleSyncPlanItem> => {
        const live = liveByName.get(repo.moduleName.toLowerCase());
        const existsInWorkbook = Boolean(live);
        const moduleType = live?.type ?? repo.inferredType;
        const unsupportedDirectCreation =
            !existsInWorkbook && (repo.subtype === 'document' || repo.subtype === 'userform');
        const workbookSource = existsInWorkbook
            ? await sourceFor(repo.moduleName)
            : '';
        const repoDisplaySource = editorPreviewSource(repo.source);
        const workbookDisplaySource = editorPreviewSource(workbookSource);
        const equal = existsInWorkbook && normalizeEol(repo.source) === normalizeEol(workbookSource);
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
            rightTitle: importWorkbookTitle(repo.moduleName, status),
            leftCode: repoDisplaySource,
            rightCode: workbookDisplaySource,
            leftRawCode: repo.source,
            rightRawCode: workbookSource,
            diff: buildSideBySideDiff(repoDisplaySource, workbookDisplaySource, writeDiffTones()),
            diffWithHeaders: buildSideBySideDiff(repo.source, workbookSource, writeDiffTones()),
        };
    }));
    const workbookOnlyItems: ModuleSyncPlanItem[] = [];

    if (importMode === 'trueUpStandardClass') {
        for (const mod of liveModules) {
            if (repoModuleNames.has(mod.name.toLowerCase()) || !importTrueUpCanRemove(mod)) {
                continue;
            }
            const relativeName = relativeNameForModule(mod);
            const workbookSource = await sourceFor(mod.name);
            const workbookDisplaySource = editorPreviewSource(workbookSource);
            workbookOnlyItems.push({
                id: `import-stale:${mod.name}`,
                moduleName: mod.name,
                moduleType: mod.type,
                documentType: mod.documentType,
                relativeName,
                status: 'will-remove',
                checked: true,
                selectable: true,
                warning: 'This standard/class workbook module is not present in the import folder and will be deleted during import true-up.',
                detail: 'Will delete workbook module',
                existsInWorkbook: true,
                existsInRepo: false,
                unsupportedDirectCreation: false,
                leftTitle: 'Repo: missing file',
                rightTitle: `Workbook: ${mod.name} (will delete)`,
                leftCode: '',
                rightCode: workbookDisplaySource,
                leftRawCode: '',
                rightRawCode: workbookSource,
                diff: buildSideBySideDiff('', workbookDisplaySource, deleteDiffTones()),
                diffWithHeaders: buildSideBySideDiff('', workbookSource, deleteDiffTones()),
            });
        }
    }

    return {
        direction: 'import',
        workbookPath: params.workbookPath,
        folderPath: params.importFolder,
        folderPathSource: params.folderPathSource,
        importMode,
        importModeSource: params.importModeSource,
        settingsPath: params.settingsPath,
        title: `Import modules: ${path.basename(params.workbookPath)}`,
        items: [...items, ...workbookOnlyItems].sort(compareSyncItems),
        warnings: items
            .filter((item) => item.unsupportedDirectCreation)
            .map((item) => `${item.moduleName}: skipping import unless the module already exists in the workbook.`),
    };
    });
}

export function buildSideBySideDiff(
    leftText: string,
    rightText: string,
    options: SideBySideDiffOptions = {},
): ModuleSyncDiffLine[] {
    const left = splitLines(leftText);
    const right = splitLines(rightText);
    const leftOnlyKind = options.leftOnlyKind ?? 'removed';
    const rightOnlyKind = options.rightOnlyKind ?? 'added';
    const out: ModuleSyncDiffLine[] = [];
    let i = 0;
    let j = 0;
    // Common prefix/suffix lines need no LCS table, so identical texts (the
    // "unchanged" plan items) skip the O(n*m) table allocation entirely.
    let leftEnd = left.length;
    let rightEnd = right.length;
    while (i < leftEnd && j < rightEnd && left[i] === right[j]) {
        out.push({
            leftNumber: i + 1,
            rightNumber: j + 1,
            left: left[i],
            right: right[j],
            kind: 'equal',
        });
        i++;
        j++;
    }
    let suffixLength = 0;
    while (leftEnd > i && rightEnd > j && left[leftEnd - 1] === right[rightEnd - 1]) {
        leftEnd--;
        rightEnd--;
        suffixLength++;
    }
    const midLeft = left.slice(i, leftEnd);
    const midRight = right.slice(j, rightEnd);
    const table = lcsTable(midLeft, midRight);
    let mi = 0;
    let mj = 0;
    while (mi < midLeft.length && mj < midRight.length) {
        if (midLeft[mi] === midRight[mj]) {
            out.push({
                leftNumber: i + mi + 1,
                rightNumber: j + mj + 1,
                left: midLeft[mi],
                right: midRight[mj],
                kind: 'equal',
            });
            mi++;
            mj++;
        } else if (table[mi + 1][mj] >= table[mi][mj + 1]) {
            if (table[mi + 1][mj] === table[mi][mj + 1]) {
                out.push({
                    leftNumber: i + mi + 1,
                    rightNumber: j + mj + 1,
                    left: midLeft[mi],
                    right: midRight[mj],
                    kind: 'changed',
                });
                mi++;
                mj++;
            } else {
                out.push({
                    leftNumber: i + mi + 1,
                    left: midLeft[mi],
                    right: '',
                    kind: leftOnlyKind,
                });
                mi++;
            }
        } else {
            out.push({
                rightNumber: j + mj + 1,
                left: '',
                right: midRight[mj],
                kind: rightOnlyKind,
            });
            mj++;
        }
    }
    while (mi < midLeft.length) {
        out.push({ leftNumber: i + mi + 1, left: midLeft[mi], right: '', kind: leftOnlyKind });
        mi++;
    }
    while (mj < midRight.length) {
        out.push({ rightNumber: j + mj + 1, left: '', right: midRight[mj], kind: rightOnlyKind });
        mj++;
    }
    for (let k = 0; k < suffixLength; k++) {
        out.push({
            leftNumber: leftEnd + k + 1,
            rightNumber: rightEnd + k + 1,
            left: left[leftEnd + k],
            right: right[rightEnd + k],
            kind: 'equal',
        });
    }
    return out.length > 0 ? out : [{ left: '', right: '', kind: 'equal' }];
}

function writeDiffTones(): SideBySideDiffOptions {
    return { leftOnlyKind: 'added', rightOnlyKind: 'removed' };
}

function deleteDiffTones(): SideBySideDiffOptions {
    return { leftOnlyKind: 'removed', rightOnlyKind: 'removed' };
}

function exportRepoTitle(relativeName: string, status: ModuleSyncItemStatus): string {
    switch (status) {
        case 'will-create':
            return `Repo: ${relativeName} (will create)`;
        case 'will-write':
            return `Repo: ${relativeName} (will overwrite)`;
        default:
            return `Repo: ${relativeName}`;
    }
}

function importWorkbookTitle(moduleName: string, status: ModuleSyncItemStatus): string {
    switch (status) {
        case 'will-create':
            return `Workbook: ${moduleName} (will create)`;
        case 'will-update':
            return `Workbook: ${moduleName} (will update)`;
        case 'skipping-import':
            return `Workbook: ${moduleName} (cannot create)`;
        default:
            return `Workbook: ${moduleName}`;
    }
}

export function editorPreviewSource(source: string): string {
    const lines = normalizeEol(source)
        .split('\n')
        .filter((line) => !isVbaAttributeLine(line));
    while (lines.length > 0 && lines[0].trim() === '') {
        lines.shift();
    }
    return lines.join('\n');
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
        case 'skipping-import':
            return 'Skipping import';
        case 'read-error':
            return 'Read error';
    }
}

async function readRepoModuleFile(folder: string, file: string): Promise<RepoModuleFile> {
    const sourcePath = path.join(folder, file);
    const source = await fs.promises.readFile(sourcePath, 'utf8');
    const ext = path.extname(file).toLowerCase();
    const moduleName = sanitizeFileName(path.basename(file, ext)) || path.basename(file, ext);
    const subtype = ext === '.bas' ? 'standard' : detectClsSubtype(moduleName, source);
    return {
        file,
        moduleName,
        inferredType: subtype === 'document' ? 'document' : subtype === 'userform' ? 'userform' : subtype,
        subtype,
        source,
        sourcePath,
    };
}

/**
 * Infer module type from source content and name.
 *
 * Mirrors _module_type in python/xlide/vba_io.py — the shared classification
 * table tests on both sides pin the two implementations together.
 */
export function classifyModuleType(name: string, source: string): 'standard' | 'document' | 'userform' {
    const vbBaseMatch = source.match(VB_BASE_RE);
    const vbBase = vbBaseMatch ? vbBaseMatch[1] : '';
    if (vbBase) {
        // UserForms always have TWO GUIDs in VB_Base (type-lib + instance).
        // Class and document modules each have exactly one.
        const guids = vbBase.match(GUID_RE) ?? [];
        if (guids.length >= 2) {
            return 'userform';
        }
        if (guids.some((guid) => DOCUMENT_CLSIDS.has(guid.toUpperCase()))) {
            return 'document';
        }
    }
    // VB_PredeclaredId=True is shared by Excel document modules and predeclared
    // class modules. Treating it as document-only misclassifies singleton-style
    // classes such as stdVBA's stdArray/stdLambda modules.
    // Well-known document-module names across common Excel locales.
    if (name === 'ThisWorkbook' || DOCUMENT_MODULE_NAME_RE.test(name)) {
        return 'document';
    }
    return 'standard';
}

function detectClsSubtype(name: string, source: string): 'class' | 'document' | 'userform' {
    const moduleType = classifyModuleType(name, source);
    // A .cls file is never a standard module — mirror the VBAModuleKind
    // upgrade in vba_io._module_entries.
    return moduleType === 'standard' ? 'class' : moduleType;
}

function splitLines(text: string): string[] {
    if (text.length === 0) {
        return [];
    }
    return normalizeEol(text).split('\n');
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

function importTrueUpCanRemove(mod: ModuleInfo): boolean {
    return mod.type === 'standard' || mod.type === 'class';
}

function moduleKindLabel(kind: string): string {
    switch (kind) {
        case 'document':
            return 'Worksheet/ThisWorkbook';
        case 'userform':
            return 'UserForm .cls code-behind';
        default:
            return kind;
    }
}
