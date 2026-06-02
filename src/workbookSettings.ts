import * as fs from 'fs';
import * as path from 'path';
import type { ImportMode } from './moduleSyncPlan';
import {
    normalizeAnalysisRuleCodes,
    normalizeAnalysisVisibleSeverities,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';

type ExportMode = 'exportAll' | 'trueUp';

interface WorkbookSettingsConfig {
    exportFolder?: string;
    exportMode?: ExportMode;
    importMode?: ImportMode;
    analysis?: WorkbookAnalysisSettingsConfig;
}

interface WorkbookAnalysisSettingsConfig {
    visibleSeverities?: AnalysisSeverityFilter[];
    untrackedRules?: string[];
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

function settingsPathForWorkbook(filePath: string): string {
    return path.join(path.dirname(filePath), `${path.basename(filePath)}.xlide_settings.json`);
}

function normalizeExportMode(mode: ExportMode | unknown): ExportMode {
    return mode === 'trueUp' ? 'trueUp' : 'exportAll';
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
    };
    const normalized: WorkbookAnalysisSettingsConfig = {};
    if (Array.isArray(source.visibleSeverities)) {
        normalized.visibleSeverities = normalizeAnalysisVisibleSeverities(source.visibleSeverities);
    }
    if (Array.isArray(source.untrackedRules)) {
        normalized.untrackedRules = normalizeAnalysisRuleCodes(source.untrackedRules);
    }
    return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function normalizeWorkbookSettingsConfig(config: {
    exportFolder?: unknown;
    exportMode?: unknown;
    importMode?: unknown;
    analysis?: unknown;
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
    return normalized;
}

function isWorkbookSettingsError(value: unknown): value is WorkbookSettingsError {
    return value instanceof WorkbookSettingsError;
}

function parseWorkbookSettingsConfig(value: unknown, configPath: string): WorkbookSettingsConfig {
    if (!isPlainObject(value)) {
        throw new WorkbookSettingsError(configPath, 'Expected the root value to be a JSON object.');
    }
    assertKnownKeys(value, configPath, 'root', ['exportFolder', 'exportMode', 'importMode', 'analysis']);

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
    assertKnownKeys(value, configPath, fieldPath, ['visibleSeverities', 'untrackedRules']);

    const parsed: WorkbookAnalysisSettingsConfig = {};
    if ('visibleSeverities' in value) {
        parsed.visibleSeverities = expectSeverityList(value.visibleSeverities, configPath, `${fieldPath}.visibleSeverities`);
    }
    if ('untrackedRules' in value) {
        parsed.untrackedRules = expectStringList(value.untrackedRules, configPath, `${fieldPath}.untrackedRules`);
    }
    return Object.keys(parsed).length > 0 ? parsed : undefined;
}

function expectSeverityList(value: unknown, configPath: string, fieldPath: string): AnalysisSeverityFilter[] {
    if (!Array.isArray(value)) {
        throw new WorkbookSettingsError(configPath, `Expected "${fieldPath}" to be an array.`);
    }
    const allowed = new Set<string>(['error', 'warning', 'information']);
    const parsed: AnalysisSeverityFilter[] = [];
    for (const entry of value) {
        if (typeof entry !== 'string' || !allowed.has(entry)) {
            throw new WorkbookSettingsError(
                configPath,
                `Expected "${fieldPath}" entries to be "error", "warning", or "information".`,
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
            `Unable to read settings: ${err instanceof Error ? err.message : String(err)}`,
        );
    }

    let parsed: unknown;
    try {
        parsed = JSON.parse(raw);
    } catch (err) {
        throw new WorkbookSettingsError(
            configPath,
            `Expected valid JSON: ${err instanceof Error ? err.message : String(err)}`,
        );
    }
    return parseWorkbookSettingsConfig(parsed, configPath);
}

async function setWorkbookExportMode(filePath: string, mode: ExportMode): Promise<WorkbookSettingsConfig> {
    const existing = await readWorkbookSettings(filePath);
    const updated: WorkbookSettingsConfig = {
        ...existing,
        exportMode: normalizeExportMode(mode),
    };
    await writeWorkbookSettings(filePath, updated);
    return updated;
}

function isNodeError(value: unknown): value is NodeJS.ErrnoException {
    return value !== null && typeof value === 'object' && 'code' in value;
}

async function writeWorkbookSettings(
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

export {
    type ExportMode,
    type WorkbookAnalysisSettingsConfig,
    type WorkbookSettingsConfig,
    WorkbookSettingsError,
    isWorkbookSettingsError,
    normalizeExportMode,
    readWorkbookSettings,
    setWorkbookExportMode,
    settingsPathForWorkbook,
    writeWorkbookSettings,
};
