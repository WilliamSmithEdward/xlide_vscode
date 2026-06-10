import * as fs from 'fs';
import * as path from 'path';
import type { ImportMode } from './moduleSyncPlan';
import {
    ANALYSIS_SEVERITIES,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisRuleSeverityOverrides,
    normalizeAnalysisVisibleSeverities,
    validateAnalysisRuleSeverityOverrideEntries,
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import type { XlideGlobalSettingSource } from './globalSettings';
import { workbookIdentityKey } from './workbookIdentity';
import { errorMessage } from './util/errors';

type ExportMode = 'exportAll' | 'trueUp';
type WorkbookSettingSource = 'workbook' | XlideGlobalSettingSource;

interface ResolvedWorkbookSetting<T> {
    value: T;
    source: WorkbookSettingSource;
}

interface WorkbookSettingsConfig {
    exportFolder?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
    analysis?: WorkbookAnalysisSettingsConfig;
    tests?: WorkbookTestSettingsConfig;
}

interface WorkbookAnalysisSettingsConfig {
    visibleSeverities?: AnalysisSeverityFilter[];
    untrackedRules?: string[];
    ruleSeverityOverrides?: AnalysisRuleSeverityOverrides;
}

interface WorkbookTestSettingsConfig {
    artifactFolder?: string;
    artifactRetention?: number;
}

type WorkbookSettingsConfigInput = Omit<WorkbookSettingsConfig, 'exportMode'> & {
    exportMode?: ExportMode;
};

class WorkbookSettingsError extends Error {
    constructor(
        public readonly settingsPath: string,
        message: string,
    ) {
        super(`XLIDE workbook settings file is invalid: ${settingsPath}. ${message}`);
        this.name = 'WorkbookSettingsError';
    }
}

const workbookSettingsWriteQueues = new Map<string, Promise<unknown>>();

function settingsPathForWorkbook(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.xlide_settings.json`);
}

function normalizeExportMode(mode: ExportMode | unknown): ExportMode {
    return mode === 'trueUp' ? 'trueUp' : 'exportAll';
}

function normalizeImportMode(mode: ImportMode | unknown): ImportMode {
    return mode === 'trueUpStandardClass' ? 'trueUpStandardClass' : 'updateOnly';
}

function resolveWorkbookSetting<T>(
    workbookValue: T | undefined,
    fallback: { value: T; source: XlideGlobalSettingSource },
): ResolvedWorkbookSetting<T> {
    return workbookValue === undefined
        ? { value: fallback.value, source: fallback.source }
        : { value: workbookValue, source: 'workbook' };
}

function normalizeImportModeValue(mode: unknown): ImportMode | undefined {
    return mode === 'updateOnly' || mode === 'trueUpStandardClass' ? mode : undefined;
}

function normalizeWorkbookAnalysisSettingsConfig(value: unknown): WorkbookAnalysisSettingsConfig | undefined {
    if (!value || typeof value !== 'object') {
        return undefined;
    }
    const source = value as {
        visibleSeverities?: unknown;
        untrackedRules?: unknown;
        ruleSeverityOverrides?: unknown;
    };
    const normalized: WorkbookAnalysisSettingsConfig = {};
    if (Array.isArray(source.visibleSeverities)) {
        normalized.visibleSeverities = normalizeAnalysisVisibleSeverities(source.visibleSeverities);
    }
    if (Array.isArray(source.untrackedRules)) {
        normalized.untrackedRules = normalizeAnalysisRuleCodes(source.untrackedRules);
    }
    if (source.ruleSeverityOverrides !== undefined) {
        normalized.ruleSeverityOverrides = normalizeAnalysisRuleSeverityOverrides(source.ruleSeverityOverrides);
    }
    return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function normalizeWorkbookTestSettingsConfig(value: unknown): WorkbookTestSettingsConfig | undefined {
    if (!value || typeof value !== 'object') {
        return undefined;
    }
    const source = value as {
        artifactFolder?: unknown;
        artifactRetention?: unknown;
    };
    const normalized: WorkbookTestSettingsConfig = {};
    if (typeof source.artifactFolder === 'string') {
        normalized.artifactFolder = source.artifactFolder;
    }
    if (
        typeof source.artifactRetention === 'number' &&
        Number.isInteger(source.artifactRetention) &&
        source.artifactRetention > 0
    ) {
        normalized.artifactRetention = source.artifactRetention;
    }
    return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function normalizeWorkbookSettingsConfig(config: {
    exportFolder?: unknown;
    exportMode?: unknown;
    importMode?: unknown;
    analysis?: unknown;
    tests?: unknown;
}): WorkbookSettingsConfig {
    const normalized: WorkbookSettingsConfig = {};
    if (typeof config.exportFolder === 'string') {
        normalized.exportFolder = config.exportFolder;
    }
    if (config.exportMode !== undefined) {
        normalized.exportMode = normalizeExportMode(config.exportMode);
    }
    const importMode = normalizeImportModeValue(config.importMode);
    if (importMode) {
        normalized.importMode = importMode;
    }
    const analysis = normalizeWorkbookAnalysisSettingsConfig(config.analysis);
    if (analysis) {
        normalized.analysis = analysis;
    }
    const tests = normalizeWorkbookTestSettingsConfig(config.tests);
    if (tests) {
        normalized.tests = tests;
    }
    return normalized;
}

function isWorkbookSettingsError(value: unknown): value is WorkbookSettingsError {
    return value instanceof WorkbookSettingsError;
}

function parseWorkbookSettingsConfig(value: unknown, configPath: string): WorkbookSettingsConfig {
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, 'Expected the root value to be a JSON object.');
    }
    assertKnownKeys(value, configPath, 'root', ['exportFolder', 'exportMode', 'importMode', 'analysis', 'tests']);

    const parsed: WorkbookSettingsConfig = {};
    if ('exportFolder' in value) {
        parsed.exportFolder = expectOptionalString(value.exportFolder, configPath, 'exportFolder');
    }
    if ('exportMode' in value) {
        parsed.exportMode = expectExportMode(value.exportMode, configPath, 'exportMode');
    }
    if ('importMode' in value) {
        parsed.importMode = expectImportMode(value.importMode, configPath, 'importMode');
    }
    if ('analysis' in value) {
        parsed.analysis = expectAnalysisSettings(value.analysis, configPath, 'analysis');
    }
    if ('tests' in value) {
        parsed.tests = expectTestSettings(value.tests, configPath, 'tests');
    }
    return parsed;
}

function expectAnalysisSettings(
    value: unknown,
    configPath: string,
    fieldPath: string,
): WorkbookAnalysisSettingsConfig | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be a JSON object.`);
    }
    assertKnownKeys(value, configPath, fieldPath, ['visibleSeverities', 'untrackedRules', 'ruleSeverityOverrides']);

    const parsed: WorkbookAnalysisSettingsConfig = {};
    if ('visibleSeverities' in value) {
        parsed.visibleSeverities = expectSeverityList(value.visibleSeverities, configPath, `${fieldPath}.visibleSeverities`);
    }
    if ('untrackedRules' in value) {
        parsed.untrackedRules = expectStringList(value.untrackedRules, configPath, `${fieldPath}.untrackedRules`);
    }
    if ('ruleSeverityOverrides' in value) {
        parsed.ruleSeverityOverrides = expectRuleSeverityOverrides(
            value.ruleSeverityOverrides,
            configPath,
            `${fieldPath}.ruleSeverityOverrides`,
        );
    }
    return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function expectTestSettings(
    value: unknown,
    configPath: string,
    fieldPath: string,
): WorkbookTestSettingsConfig | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be a JSON object.`);
    }
    assertKnownKeys(value, configPath, fieldPath, ['artifactFolder', 'artifactRetention']);

    const parsed: WorkbookTestSettingsConfig = {};
    if ('artifactFolder' in value) {
        parsed.artifactFolder = expectOptionalString(value.artifactFolder, configPath, `${fieldPath}.artifactFolder`);
    }
    if ('artifactRetention' in value) {
        parsed.artifactRetention = expectPositiveInteger(value.artifactRetention, configPath, `${fieldPath}.artifactRetention`);
    }
    return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function expectSeverityList(value: unknown, configPath: string, fieldPath: string): AnalysisSeverityFilter[] {
    if (!Array.isArray(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be an array.`);
    }
    const allowed = new Set<string>(ANALYSIS_SEVERITIES);
    const parsed: AnalysisSeverityFilter[] = [];
    for (const entry of value) {
        if (typeof entry !== 'string' || !allowed.has(entry)) {
            throw new WorkbookSettingsError(
                configPath,
                `Expected "${fieldPath}" entries to be one of: ${ANALYSIS_SEVERITIES.join(', ')}.`,
            );
        }
        parsed.push(entry as AnalysisSeverityFilter);
    }
    return normalizeAnalysisVisibleSeverities(parsed);
}

function expectStringList(value: unknown, configPath: string, fieldPath: string): string[] {
    if (!Array.isArray(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be an array.`);
    }
    if (!value.every((entry) => typeof entry === 'string')) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" entries to be strings.`);
    }
    return normalizeAnalysisRuleCodes(value);
}

function expectRuleSeverityOverrides(
    value: unknown,
    configPath: string,
    fieldPath: string,
): AnalysisRuleSeverityOverrides | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be a JSON object.`);
    }
    const parsed = validateAnalysisRuleSeverityOverrideEntries(value, (rawCode, requirement) => {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}.${rawCode}" ${requirement}`);
    });
    return Object.keys(parsed).length > 0 ? parsed : undefined;
}


function expectPositiveInteger(value: unknown, configPath: string, fieldPath: string): number | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be a positive integer.`);
    }
    return value;
}

function expectOptionalString(value: unknown, configPath: string, fieldPath: string): string | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== 'string') {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be a string.`);
    }
    return value;
}

function expectExportMode(value: unknown, configPath: string, fieldPath: string): ExportMode | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (value === 'exportAll' || value === 'trueUp') {
        return value;
    }
    throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be "exportAll" or "trueUp".`);
}

function expectImportMode(value: unknown, configPath: string, fieldPath: string): ImportMode | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (value === 'updateOnly' || value === 'trueUpStandardClass') {
        return value;
    }
    throw new WorkbookSettingsError(
        configPath,
        `Expected "${fieldPath}" to be "updateOnly" or "trueUpStandardClass".`,
    );
}

function assertKnownKeys(
    value: Record<string, unknown>,
    configPath: string,
    fieldPath: string,
    knownKeys: readonly string[],
): void {
    const known = new Set(knownKeys);
    const unknown = Object.keys(value).filter((key) => !known.has(key));
    if (unknown.length > 0) {
        throw new WorkbookSettingsError(
            configPath,
            `Unknown setting "${fieldPath === 'root' ? unknown[0] : `${fieldPath}.${unknown[0]}`}".`,
        );
    }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

async function readWorkbookSettings(filePath: string): Promise<WorkbookSettingsConfig> {
    const configPath = settingsPathForWorkbook(filePath);
    let raw: string;
    try {
        raw = await fs.promises.readFile(configPath, 'utf8');
    } catch (err) {
        if (isNodeError(err) && err.code === 'ENOENT') {
            return {};
        }
        throw new WorkbookSettingsError(
            configPath,
            `Unable to read settings: ${errorMessage(err)}`,
        );
    }

    let parsed: unknown;
    try {
        parsed = JSON.parse(raw);
    } catch (err) {
        parsed = recoverWorkbookSettingsJson(raw);
        if (parsed === undefined) {
            throw new WorkbookSettingsError(
                configPath,
                `Expected valid JSON: ${errorMessage(err)}`,
            );
        }
    }
    return parseWorkbookSettingsConfig(parsed, configPath);
}

async function updateWorkbookSettings(
    filePath: string,
    update: (existing: WorkbookSettingsConfig) => WorkbookSettingsConfig | undefined,
): Promise<WorkbookSettingsConfig> {
    return withWorkbookSettingsWriteLock(filePath, async () => {
        const existing = await readWorkbookSettings(filePath);
        const updatedInput = update(existing);
        if (!updatedInput) {
            return existing;
        }
        const updated = normalizeWorkbookSettingsConfig(updatedInput);
        await writeWorkbookSettingsUnlocked(filePath, updated);
        return updated;
    });
}

function isNodeError(value: unknown): value is NodeJS.ErrnoException {
    return value !== null && typeof value === 'object' && 'code' in value;
}

async function writeWorkbookSettings(
    filePath: string,
    config: WorkbookSettingsConfigInput,
): Promise<void> {
    await withWorkbookSettingsWriteLock(filePath, () => writeWorkbookSettingsUnlocked(filePath, config));
}

async function writeWorkbookSettingsUnlocked(
    filePath: string,
    config: WorkbookSettingsConfigInput,
): Promise<void> {
    const configPath = settingsPathForWorkbook(filePath);
    await fs.promises.writeFile(
        configPath,
        `${JSON.stringify(normalizeWorkbookSettingsConfig(config), null, 2)}\n`,
        'utf8',
    );
}

async function withWorkbookSettingsWriteLock<T>(
    filePath: string,
    action: () => Promise<T>,
): Promise<T> {
    const key = workbookIdentityKey(settingsPathForWorkbook(filePath));
    const previous = workbookSettingsWriteQueues.get(key) ?? Promise.resolve();
    let release: () => void = () => undefined;
    const current = new Promise<void>((resolve) => {
        release = resolve;
    });
    const queued = previous.catch(() => undefined).then(() => current);
    workbookSettingsWriteQueues.set(key, queued);
    await previous.catch(() => undefined);
    try {
        return await action();
    } finally {
        release();
        if (workbookSettingsWriteQueues.get(key) === queued) {
            workbookSettingsWriteQueues.delete(key);
        }
    }
}

function recoverWorkbookSettingsJson(raw: string): unknown | undefined {
    let offset = skipJsonWhitespace(raw, 0);
    let recovered: unknown;
    let recoveredAny = false;
    while (offset < raw.length) {
        if (raw[offset] !== '{') {
            return recoveredAny ? recovered : undefined;
        }
        const end = findJsonRootEnd(raw, offset);
        if (end === undefined) {
            return recoveredAny ? recovered : undefined;
        }
        try {
            const parsed = JSON.parse(raw.slice(offset, end));
            if (!isPlainObject(parsed)) {
                return recoveredAny ? recovered : undefined;
            }
            recovered = parsed;
            recoveredAny = true;
        } catch {
            return recoveredAny ? recovered : undefined;
        }
        offset = skipJsonWhitespace(raw, end);
    }
    return recoveredAny ? recovered : undefined;
}

function skipJsonWhitespace(raw: string, start: number): number {
    let i = start;
    while (i < raw.length && /\s/.test(raw[i])) {
        i += 1;
    }
    return i;
}

function findJsonRootEnd(raw: string, start: number): number | undefined {
    const stack: string[] = [];
    let inString = false;
    let escaping = false;
    for (let i = start; i < raw.length; i += 1) {
        const ch = raw[i];
        if (inString) {
            if (escaping) {
                escaping = false;
            } else if (ch === '\\') {
                escaping = true;
            } else if (ch === '"') {
                inString = false;
            }
            continue;
        }
        if (ch === '"') {
            inString = true;
            continue;
        }
        if (ch === '{') {
            stack.push('}');
            continue;
        }
        if (ch === '[') {
            stack.push(']');
            continue;
        }
        if (ch === '}' || ch === ']') {
            if (stack.pop() !== ch) {
                return undefined;
            }
            if (stack.length === 0) {
                return i + 1;
            }
        }
    }
    return undefined;
}

export {
    type ExportMode,
    type ResolvedWorkbookSetting,
    type WorkbookAnalysisSettingsConfig,
    type WorkbookSettingSource,
    type WorkbookSettingsConfig,
    type WorkbookTestSettingsConfig,
    WorkbookSettingsError,
    isWorkbookSettingsError,
    normalizeExportMode,
    normalizeImportMode,
    readWorkbookSettings,
    resolveWorkbookSetting,
    settingsPathForWorkbook,
    updateWorkbookSettings,
    writeWorkbookSettings,
};
