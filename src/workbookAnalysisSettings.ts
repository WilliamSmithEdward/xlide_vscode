import {
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleSeverityOverrides,
    normalizeAnalysisVisibleSeverities,
    setAnalysisRuleTrackedInList,
    type AnalysisRuleTrackingUpdate,
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
    readWorkbookSettings,
    resolveWorkbookSetting,
    updateWorkbookSettings,
    type WorkbookAnalysisSettingsConfig,
    type WorkbookSettingSource,
    type WorkbookSettingsConfig,
} from './workbookSettings';
import {
    ruleSeverityOverridesSettingFromConfig,
    untrackedAnalysisRulesSettingFromConfig,
    visibleAnalysisSeveritiesSettingFromConfig,
} from './analysisOptions';

export type WorkbookAnalysisSettingsSource = WorkbookSettingSource;

export interface EffectiveWorkbookAnalysisSettings {
    visibleSeverities: AnalysisSeverityFilter[];
    visibleSeveritiesSource: WorkbookAnalysisSettingsSource;
    untrackedRules: string[];
    untrackedRulesSource: WorkbookAnalysisSettingsSource;
    ruleSeverityOverrides: AnalysisRuleSeverityOverrides;
    ruleSeverityOverridesSource: WorkbookAnalysisSettingsSource;
}

export async function effectiveWorkbookAnalysisSettings(
    workbookPath: string | undefined,
): Promise<EffectiveWorkbookAnalysisSettings> {
    if (!workbookPath) {
        return effectiveWorkbookAnalysisSettingsFromConfig(undefined);
    }
    return effectiveWorkbookAnalysisSettingsFromConfig(
        workbookPath,
        await readWorkbookSettings(workbookPath),
    );
}

export function effectiveWorkbookAnalysisSettingsFromConfig(
    workbookPath: string | undefined,
    config: WorkbookSettingsConfig = {},
): EffectiveWorkbookAnalysisSettings {
    const globalVisibleSeverities = visibleAnalysisSeveritiesSettingFromConfig();
    const globalUntrackedRules = untrackedAnalysisRulesSettingFromConfig();
    const globalRuleSeverityOverrides = ruleSeverityOverridesSettingFromConfig();
    if (!workbookPath) {
        return {
            visibleSeverities: globalVisibleSeverities.value,
            visibleSeveritiesSource: globalVisibleSeverities.source,
            untrackedRules: globalUntrackedRules.value,
            untrackedRulesSource: globalUntrackedRules.source,
            ruleSeverityOverrides: globalRuleSeverityOverrides.value,
            ruleSeverityOverridesSource: globalRuleSeverityOverrides.source,
        };
    }

    const { analysis } = config;
    const visibleSeverities = resolveWorkbookSetting(analysis?.visibleSeverities, globalVisibleSeverities);
    const untrackedRules = resolveWorkbookSetting(analysis?.untrackedRules, globalUntrackedRules);
    const ruleSeverityOverrides = resolveWorkbookSetting(
        analysis?.ruleSeverityOverrides,
        globalRuleSeverityOverrides,
    );
    return {
        visibleSeverities: visibleSeverities.value,
        visibleSeveritiesSource: visibleSeverities.source,
        untrackedRules: untrackedRules.value,
        untrackedRulesSource: untrackedRules.source,
        ruleSeverityOverrides: ruleSeverityOverrides.value,
        ruleSeverityOverridesSource: ruleSeverityOverrides.source,
    };
}

export async function setWorkbookAnalysisVisibleSeverities(
    workbookPath: string,
    severities: unknown,
): Promise<EffectiveWorkbookAnalysisSettings> {
    await updateWorkbookSettings(workbookPath, (existing) => withAnalysisSettings(existing, {
        ...existing.analysis,
        visibleSeverities: normalizeAnalysisVisibleSeverities(severities),
    }));
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function setWorkbookAnalysisRuleTracked(
    workbookPath: string,
    code: string | undefined,
    tracked: boolean,
): Promise<AnalysisRuleTrackingUpdate> {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        const settings = await effectiveWorkbookAnalysisSettings(workbookPath);
        return {
            tracked,
            changed: false,
            untrackedRules: settings.untrackedRules,
        };
    }

    let result: AnalysisRuleTrackingUpdate = {
        code: normalized,
        tracked,
        changed: false,
        untrackedRules: [],
    };
    await updateWorkbookSettings(workbookPath, (existing) => {
        const current = existing.analysis?.untrackedRules ?? untrackedAnalysisRulesSettingFromConfig().value;
        const next = setAnalysisRuleTrackedInList(current, normalized, tracked);
        const changed = current.length !== next.length || current.some((entry, index) => entry !== next[index]);
        result = {
            code: normalized,
            tracked,
            changed,
            untrackedRules: next,
        };
        return changed || !existing.analysis?.untrackedRules
            ? withAnalysisSettings(existing, {
                ...existing.analysis,
                untrackedRules: next,
            })
            : undefined;
    });
    return result;
}

export async function setWorkbookAnalysisRuleSeverityOverride(
    workbookPath: string,
    code: string | undefined,
    severity: unknown,
): Promise<EffectiveWorkbookAnalysisSettings> {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        return effectiveWorkbookAnalysisSettings(workbookPath);
    }
    await updateWorkbookSettings(workbookPath, (existing) => {
        const current = existing.analysis?.ruleSeverityOverrides ?? ruleSeverityOverridesSettingFromConfig().value;
        const next = normalizeAnalysisRuleSeverityOverrides({
            ...current,
            [normalized]: severity,
        });
        return withAnalysisSettings(existing, {
            ...existing.analysis,
            ruleSeverityOverrides: next,
        });
    });
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function clearWorkbookAnalysisRuleSeverityOverride(
    workbookPath: string,
    code: string | undefined,
): Promise<EffectiveWorkbookAnalysisSettings> {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        return effectiveWorkbookAnalysisSettings(workbookPath);
    }
    await updateWorkbookSettings(workbookPath, (existing) => {
        const current = existing.analysis?.ruleSeverityOverrides;
        if (!current || !(normalized in current)) {
            return undefined;
        }
        const next = { ...current };
        delete next[normalized];
        return withAnalysisSettings(existing, {
            ...existing.analysis,
            ruleSeverityOverrides: normalizeAnalysisRuleSeverityOverrides(next),
        });
    });
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function resetWorkbookAnalysisVisibleSeverities(
    workbookPath: string,
): Promise<EffectiveWorkbookAnalysisSettings> {
    await updateWorkbookSettings(workbookPath, (existing) => {
        const analysis = { ...(existing.analysis ?? {}) };
        delete analysis.visibleSeverities;
        return withAnalysisSettings(existing, analysis);
    });
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function resetWorkbookAnalysisRuleTracking(
    workbookPath: string,
): Promise<EffectiveWorkbookAnalysisSettings> {
    await updateWorkbookSettings(workbookPath, (existing) => {
        const analysis = { ...(existing.analysis ?? {}) };
        delete analysis.untrackedRules;
        return withAnalysisSettings(existing, analysis);
    });
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function resetWorkbookAnalysisRuleSeverities(
    workbookPath: string,
): Promise<EffectiveWorkbookAnalysisSettings> {
    await updateWorkbookSettings(workbookPath, (existing) => {
        const analysis = { ...(existing.analysis ?? {}) };
        delete analysis.ruleSeverityOverrides;
        return withAnalysisSettings(existing, analysis);
    });
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

export async function resetWorkbookAnalysisSettings(
    workbookPath: string,
): Promise<EffectiveWorkbookAnalysisSettings> {
    await updateWorkbookSettings(workbookPath, withoutAnalysisSettings);
    return effectiveWorkbookAnalysisSettings(workbookPath);
}

function withAnalysisSettings(
    config: WorkbookSettingsConfig,
    analysis: WorkbookAnalysisSettingsConfig,
): WorkbookSettingsConfig {
    const normalizedAnalysis = compactAnalysisSettings(analysis);
    return {
        exportFolder: config.exportFolder,
        exportMode: config.exportMode,
        importMode: config.importMode,
        tests: config.tests,
        ...(normalizedAnalysis ? { analysis: normalizedAnalysis } : {}),
    };
}

function withoutAnalysisSettings(config: WorkbookSettingsConfig): WorkbookSettingsConfig {
    return {
        exportFolder: config.exportFolder,
        exportMode: config.exportMode,
        importMode: config.importMode,
        tests: config.tests,
    };
}

function compactAnalysisSettings(
    analysis: WorkbookAnalysisSettingsConfig,
): WorkbookAnalysisSettingsConfig | undefined {
    const compacted: WorkbookAnalysisSettingsConfig = {};
    if (analysis.visibleSeverities) {
        compacted.visibleSeverities = analysis.visibleSeverities;
    }
    if (analysis.untrackedRules) {
        compacted.untrackedRules = analysis.untrackedRules;
    }
    if (analysis.ruleSeverityOverrides && Object.keys(analysis.ruleSeverityOverrides).length > 0) {
        compacted.ruleSeverityOverrides = analysis.ruleSeverityOverrides;
    }
    return Object.keys(compacted).length > 0 ? compacted : undefined;
}
