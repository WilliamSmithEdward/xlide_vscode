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

// One codec serves both sidecar passes: strict parsing passes a `reject`
// callback that throws WorkbookSettingsError (and rejects unknown keys),
// while the lenient normalizer passes undefined and coerces or drops
// invalid values instead.
type WorkbookSettingsReject = ((message: string) => never) | undefined;

function normalizeWorkbookSettingsConfig(config: {
    exportFolder?: unknown;
    exportMode?: unknown;
    importMode?: unknown;
    analysis?: unknown;
    tests?: unknown;
}): WorkbookSettingsConfig {
    return codecWorkbookSettingsConfig(config as Record<string, unknown>, undefined);
}

function isWorkbookSettingsError(value: unknown): value is WorkbookSettingsError {
    return value instanceof WorkbookSettingsError;
}

function parseWorkbookSettingsConfig(value: unknown, configPath: string): WorkbookSettingsConfig {
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, 'Expected the root value to be a JSON object.');
    }
    return codecWorkbookSettingsConfig(value, (message: string): never => {
        throw new WorkbookSettingsError(configPath, message);
    });
}

function codecWorkbookSettingsConfig(
    value: Record<string, unknown>,
    reject: WorkbookSettingsReject,
): WorkbookSettingsConfig {
    if (reject) {
        assertKnownKeys(value, 'root', ['exportFolder', 'exportMode', 'importMode', 'analysis', 'tests'], reject);
    }
    const config: WorkbookSettingsConfig = {};
    const exportFolder = codecOptionalString(value.exportFolder, 'exportFolder', reject);
    if (exportFolder !== undefined) {
        config.exportFolder = exportFolder;
    }
    const exportMode = codecExportMode(value.exportMode, 'exportMode', reject);
    if (exportMode !== undefined) {
        config.exportMode = exportMode;
    }
    const importMode = codecImportMode(value.importMode, 'importMode', reject);
    if (importMode !== undefined) {
        config.importMode = importMode;
    }
    const analysis = codecAnalysisSettings(value.analysis, 'analysis', reject);
    if (analysis) {
        config.analysis = analysis;
    }
    const tests = codecTestSettings(value.tests, 'tests', reject);
    if (tests) {
        config.tests = tests;
    }
    return config;
}

function codecAnalysisSettings(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): WorkbookAnalysisSettingsConfig | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        return reject?.(`Expected "${fieldPath}" to be a JSON object.`);
    }
    if (reject) {
        assertKnownKeys(value, fieldPath, ['visibleSeverities', 'untrackedRules', 'ruleSeverityOverrides'], reject);
    }
    const analysis: WorkbookAnalysisSettingsConfig = {};
    const visibleSeverities = codecSeverityList(value.visibleSeverities, `${fieldPath}.visibleSeverities`, reject);
    if (visibleSeverities !== undefined) {
        analysis.visibleSeverities = visibleSeverities;
    }
    const untrackedRules = codecRuleCodeList(value.untrackedRules, `${fieldPath}.untrackedRules`, reject);
    if (untrackedRules !== undefined) {
        analysis.untrackedRules = untrackedRules;
    }
    const ruleSeverityOverrides = codecRuleSeverityOverrides(
        value.ruleSeverityOverrides,
        `${fieldPath}.ruleSeverityOverrides`,
        reject,
    );
    if (ruleSeverityOverrides !== undefined) {
        analysis.ruleSeverityOverrides = ruleSeverityOverrides;
    }
    return Object.keys(analysis).length > 0 ? analysis : undefined;
}

function codecTestSettings(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): WorkbookTestSettingsConfig | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        return reject?.(`Expected "${fieldPath}" to be a JSON object.`);
    }
    if (reject) {
        assertKnownKeys(value, fieldPath, ['artifactFolder', 'artifactRetention'], reject);
    }
    const tests: WorkbookTestSettingsConfig = {};
    const artifactFolder = codecOptionalString(value.artifactFolder, `${fieldPath}.artifactFolder`, reject);
    if (artifactFolder !== undefined) {
        tests.artifactFolder = artifactFolder;
    }
    const artifactRetention = codecPositiveInteger(value.artifactRetention, `${fieldPath}.artifactRetention`, reject);
    if (artifactRetention !== undefined) {
        tests.artifactRetention = artifactRetention;
    }
    return Object.keys(tests).length > 0 ? tests : undefined;
}

function codecSeverityList(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): AnalysisSeverityFilter[] | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!Array.isArray(value)) {
        return reject?.(`Expected "${fieldPath}" to be an array.`);
    }
    if (reject) {
        const allowed = new Set<string>(ANALYSIS_SEVERITIES);
        for (const entry of value) {
            if (typeof entry !== 'string' || !allowed.has(entry)) {
                reject(`Expected "${fieldPath}" entries to be one of: ${ANALYSIS_SEVERITIES.join(', ')}.`);
            }
        }
    }
    return normalizeAnalysisVisibleSeverities(value);
}

function codecRuleCodeList(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): string[] | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!Array.isArray(value)) {
        return reject?.(`Expected "${fieldPath}" to be an array.`);
    }
    if (reject && !value.every((entry) => typeof entry === 'string')) {
        reject(`Expected "${fieldPath}" entries to be strings.`);
    }
    return normalizeAnalysisRuleCodes(value);
}

function codecRuleSeverityOverrides(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): AnalysisRuleSeverityOverrides | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!reject) {
        return normalizeAnalysisRuleSeverityOverrides(value);
    }
    if (!isPlainObject(value)) {
        return reject(`Expected "${fieldPath}" to be a JSON object.`);
    }
    const parsed = validateAnalysisRuleSeverityOverrideEntries(value, (rawCode, requirement) => {
        reject(`Expected "${fieldPath}.${rawCode}" ${requirement}`);
    });
    return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function codecPositiveInteger(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): number | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
        return value;
    }
    return reject?.(`Expected "${fieldPath}" to be a positive integer.`);
}

function codecOptionalString(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): string | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (typeof value === 'string') {
        return value;
    }
    return reject?.(`Expected "${fieldPath}" to be a string.`);
}

function codecExportMode(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): ExportMode | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (reject && value !== 'exportAll' && value !== 'trueUp') {
        reject(`Expected "${fieldPath}" to be "exportAll" or "trueUp".`);
    }
    return normalizeExportMode(value);
}

function codecImportMode(
    value: unknown,
    fieldPath: string,
    reject: WorkbookSettingsReject,
): ImportMode | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (value === 'updateOnly' || value === 'trueUpStandardClass') {
        return value;
    }
    return reject?.(`Expected "${fieldPath}" to be "updateOnly" or "trueUpStandardClass".`);
}

function assertKnownKeys(
    value: Record<string, unknown>,
    fieldPath: string,
    knownKeys: readonly string[],
    reject: (message: string) => never,
): void {
    const known = new Set(knownKeys);
    const unknown = Object.keys(value).filter((key) => !known.has(key));
    if (unknown.length > 0) {
        reject(`Unknown setting "${fieldPath === 'root' ? unknown[0] : `${fieldPath}.${unknown[0]}`}".`);
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
