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
import { projectIdentityKey } from './projectIdentity';
import { errorMessage } from './util/errors';

type ExportMode = 'exportAll' | 'trueUp';
type ProjectSettingSource = 'project' | XlideGlobalSettingSource;

interface ResolvedProjectSetting<T> {
    value: T;
    source: ProjectSettingSource;
}

interface ProjectSettingsConfig {
    exportFolder?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
    analysis?: ProjectAnalysisSettingsConfig;
    tests?: WorkbookTestSettingsConfig;
}

interface ProjectAnalysisSettingsConfig {
    visibleSeverities?: AnalysisSeverityFilter[];
    untrackedRules?: string[];
    ruleSeverityOverrides?: AnalysisRuleSeverityOverrides;
}

interface WorkbookTestSettingsConfig {
    artifactFolder?: string;
    artifactRetention?: number;
}

type ProjectSettingsConfigInput = Omit<ProjectSettingsConfig, 'exportMode'> & {
    exportMode?: ExportMode;
};

class ProjectSettingsError extends Error {
    constructor(
        public readonly settingsPath: string,
        message: string,
    ) {
        super(`XLIDE project settings file is invalid: ${settingsPath}. ${message}`);
        this.name = 'ProjectSettingsError';
    }
}

const projectSettingsWriteQueues = new Map<string, Promise<unknown>>();

function settingsPathForProject(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.xlide_settings.json`);
}

function normalizeExportMode(mode: ExportMode | unknown): ExportMode {
    return mode === 'trueUp' ? 'trueUp' : 'exportAll';
}

function normalizeImportMode(mode: ImportMode | unknown): ImportMode {
    return mode === 'trueUpStandardClass' ? 'trueUpStandardClass' : 'updateOnly';
}

function resolveProjectSetting<T>(
    projectValue: T | undefined,
    fallback: { value: T; source: XlideGlobalSettingSource },
): ResolvedProjectSetting<T> {
    return projectValue === undefined
        ? { value: fallback.value, source: fallback.source }
        : { value: projectValue, source: 'project' };
}

// One codec serves both sidecar passes: strict parsing passes a `reject`
// callback that throws ProjectSettingsError (and rejects unknown keys),
// while the lenient normalizer passes undefined and coerces or drops
// invalid values instead.
type ProjectSettingsReject = ((message: string) => never) | undefined;

function normalizeProjectSettingsConfig(config: {
    exportFolder?: unknown;
    exportMode?: unknown;
    importMode?: unknown;
    analysis?: unknown;
    tests?: unknown;
}): ProjectSettingsConfig {
    return codecProjectSettingsConfig(config as Record<string, unknown>, undefined);
}

function isProjectSettingsError(value: unknown): value is ProjectSettingsError {
    return value instanceof ProjectSettingsError;
}

function parseProjectSettingsConfig(value: unknown, configPath: string): ProjectSettingsConfig {
    if (!isPlainObject(value)) {
        throw new ProjectSettingsError(configPath, 'Expected the root value to be a JSON object.');
    }
    return codecProjectSettingsConfig(value, (message: string): never => {
        throw new ProjectSettingsError(configPath, message);
    });
}

function codecProjectSettingsConfig(
    value: Record<string, unknown>,
    reject: ProjectSettingsReject,
): ProjectSettingsConfig {
    if (reject) {
        assertKnownKeys(value, 'root', ['exportFolder', 'exportMode', 'importMode', 'analysis', 'tests'], reject);
    }
    const config: ProjectSettingsConfig = {};
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
    reject: ProjectSettingsReject,
): ProjectAnalysisSettingsConfig | undefined {
    if (value === undefined) {
        return undefined;
    }
    if (!isPlainObject(value)) {
        return reject?.(`Expected "${fieldPath}" to be a JSON object.`);
    }
    if (reject) {
        assertKnownKeys(value, fieldPath, ['visibleSeverities', 'untrackedRules', 'ruleSeverityOverrides'], reject);
    }
    const analysis: ProjectAnalysisSettingsConfig = {};
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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
    reject: ProjectSettingsReject,
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

async function readProjectSettings(
    filePath: string,
    opts: { lenient?: boolean } = {},
): Promise<ProjectSettingsConfig> {
    const configPath = settingsPathForProject(filePath);
    let raw: string;
    try {
        raw = await fs.promises.readFile(configPath, 'utf8');
    } catch (err) {
        if (isNodeError(err) && err.code === 'ENOENT') {
            return {};
        }
        // The lenient (apply) read powers per-keystroke diagnostics, so it must
        // never throw on a stale/unreadable sidecar - otherwise a single bad file
        // blasts an error across every module. Fall back to no project settings.
        if (opts.lenient) {
            return {};
        }
        throw new ProjectSettingsError(
            configPath,
            `Unable to read settings: ${errorMessage(err)}`,
        );
    }

    let parsed: unknown;
    try {
        parsed = JSON.parse(raw);
    } catch (err) {
        parsed = recoverProjectSettingsJson(raw);
        if (parsed === undefined) {
            if (opts.lenient) {
                return {};
            }
            throw new ProjectSettingsError(
                configPath,
                `Expected valid JSON: ${errorMessage(err)}`,
            );
        }
    }
    // Lenient: drop unknown/renamed keys and codes (forward/back-compat across
    // versions) and keep the still-valid subset, so version skew never surfaces
    // as a diagnostic. Strict: reject unknown keys so the settings editor can
    // surface a genuine user typo.
    if (opts.lenient) {
        return normalizeProjectSettingsConfig(isPlainObject(parsed) ? parsed : {});
    }
    return parseProjectSettingsConfig(parsed, configPath);
}

async function updateProjectSettings(
    filePath: string,
    update: (existing: ProjectSettingsConfig) => ProjectSettingsConfig | undefined,
): Promise<ProjectSettingsConfig> {
    return withProjectSettingsWriteLock(filePath, async () => {
        const existing = await readProjectSettings(filePath);
        const updatedInput = update(existing);
        if (!updatedInput) {
            return existing;
        }
        const updated = normalizeProjectSettingsConfig(updatedInput);
        await writeProjectSettingsUnlocked(filePath, updated);
        return updated;
    });
}

function isNodeError(value: unknown): value is NodeJS.ErrnoException {
    return value !== null && typeof value === 'object' && 'code' in value;
}

async function writeProjectSettings(
    filePath: string,
    config: ProjectSettingsConfigInput,
): Promise<void> {
    await withProjectSettingsWriteLock(filePath, () => writeProjectSettingsUnlocked(filePath, config));
}

async function writeProjectSettingsUnlocked(
    filePath: string,
    config: ProjectSettingsConfigInput,
): Promise<void> {
    const configPath = settingsPathForProject(filePath);
    await fs.promises.writeFile(
        configPath,
        `${JSON.stringify(normalizeProjectSettingsConfig(config), null, 2)}\n`,
        'utf8',
    );
}

async function withProjectSettingsWriteLock<T>(
    filePath: string,
    action: () => Promise<T>,
): Promise<T> {
    const key = projectIdentityKey(settingsPathForProject(filePath));
    const previous = projectSettingsWriteQueues.get(key) ?? Promise.resolve();
    let release: () => void = () => undefined;
    const current = new Promise<void>((resolve) => {
        release = resolve;
    });
    const queued = previous.catch(() => undefined).then(() => current);
    projectSettingsWriteQueues.set(key, queued);
    await previous.catch(() => undefined);
    try {
        return await action();
    } finally {
        release();
        if (projectSettingsWriteQueues.get(key) === queued) {
            projectSettingsWriteQueues.delete(key);
        }
    }
}

function recoverProjectSettingsJson(raw: string): unknown | undefined {
    // Recovery is only safe when the ENTIRE input parses as a sequence of complete
    // top-level objects (the intended "trailing duplicate, last wins" case). If any
    // trailing content is truncated or garbage, we must NOT silently fall back to an
    // earlier object and overwrite the newer (intended) one - return undefined so the
    // caller surfaces a ProjectSettingsError instead of destroying newer settings.
    let offset = skipJsonWhitespace(raw, 0);
    let recovered: unknown;
    let recoveredAny = false;
    while (offset < raw.length) {
        if (raw[offset] !== '{') {
            return undefined;
        }
        const end = findJsonRootEnd(raw, offset);
        if (end === undefined) {
            return undefined;
        }
        try {
            const parsed = JSON.parse(raw.slice(offset, end));
            if (!isPlainObject(parsed)) {
                return undefined;
            }
            recovered = parsed;
            recoveredAny = true;
        } catch {
            return undefined;
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
    type ResolvedProjectSetting,
    type ProjectAnalysisSettingsConfig,
    type ProjectSettingSource,
    type ProjectSettingsConfig,
    type WorkbookTestSettingsConfig,
    ProjectSettingsError,
    isProjectSettingsError,
    normalizeExportMode,
    normalizeImportMode,
    readProjectSettings,
    resolveProjectSetting,
    settingsPathForProject,
    updateProjectSettings,
    writeProjectSettings,
};
