import {
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleCodes,
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
    workbookUntrackedRules: string[];
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
            workbookUntrackedRules: [],
            ruleSeverityOverrides: globalRuleSeverityOverrides.value,
            ruleSeverityOverridesSource: globalRuleSeverityOverrides.source,
        };
    }

    const { analysis } = config;
    const visibleSeverities = resolveWorkbookSetting(analysis?.visibleSeverities, globalVisibleSeverities);
    const workbookUntrackedRules = normalizeAnalysisRuleCodes(analysis?.untrackedRules ?? []);
    const untrackedRules = analysis?.untrackedRules === undefined
        ? {
            value: globalUntrackedRules.value,
            source: globalUntrackedRules.source,
        }
        : {
            value: normalizeAnalysisRuleCodes([...globalUntrackedRules.value, ...workbookUntrackedRules]),
            source: 'workbook' as const,
        };
    const ruleSeverityOverrides = resolveWorkbookSetting(
        analysis?.ruleSeverityOverrides,
        globalRuleSeverityOverrides,
    );
    return {
        visibleSeverities: visibleSeverities.value,
        visibleSeveritiesSource: visibleSeverities.source,
        untrackedRules: untrackedRules.value,
        untrackedRulesSource: untrackedRules.source,
        workbookUntrackedRules,
        ruleSeverityOverrides: ruleSeverityOverrides.value,
        ruleSeverityOverridesSource: ruleSeverityOverrides.source,
    };
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
        const current = existing.analysis?.untrackedRules ?? [];
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

export async function resetWorkbookAnalysisRuleTracking(
    workbookPath: string,
): Promise<EffectiveWorkbookAnalysisSettings> {
    const updated = await updateWorkbookSettings(workbookPath, (existing) => {
        const analysis = { ...(existing.analysis ?? {}) };
        delete analysis.untrackedRules;
        return withAnalysisSettings(existing, analysis);
    });
    return effectiveWorkbookAnalysisSettingsFromConfig(workbookPath, updated);
}

function withAnalysisSettings(
    config: WorkbookSettingsConfig,
    analysis: WorkbookAnalysisSettingsConfig,
): WorkbookSettingsConfig {
    // Spread so future top-level sidecar keys survive analysis-settings writes,
    // matching workbookSettingsWithModuleSyncPatch/workbookSettingsWithTestPatch.
    const normalizedAnalysis = compactAnalysisSettings(analysis);
    const next = { ...config };
    if (normalizedAnalysis) {
        next.analysis = normalizedAnalysis;
    } else {
        delete next.analysis;
    }
    return next;
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
