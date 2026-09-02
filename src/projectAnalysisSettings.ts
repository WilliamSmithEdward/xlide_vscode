import {
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleCodes,
    normalizeKnownAnalysisRuleCodes,
    planAnalysisRuleTrackingUpdate,
    type AnalysisRuleTrackingUpdate,
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
    readProjectSettings,
    resolveProjectSetting,
    updateProjectSettings,
    type ProjectAnalysisSettingsConfig,
    type ProjectSettingSource,
    type ProjectSettingsConfig,
} from './projectSettings';
import {
    ruleSeverityOverridesSettingFromConfig,
    untrackedAnalysisRulesSettingFromConfig,
    visibleAnalysisSeveritiesSettingFromConfig,
} from './analysisOptions';

export type ProjectAnalysisSettingsSource = ProjectSettingSource;

export interface EffectiveProjectAnalysisSettings {
    visibleSeverities: AnalysisSeverityFilter[];
    visibleSeveritiesSource: ProjectAnalysisSettingsSource;
    untrackedRules: string[];
    untrackedRulesSource: ProjectAnalysisSettingsSource;
    projectUntrackedRules: string[];
    ruleSeverityOverrides: AnalysisRuleSeverityOverrides;
    ruleSeverityOverridesSource: ProjectAnalysisSettingsSource;
}

export async function effectiveProjectAnalysisSettings(
    projectPath: string | undefined,
): Promise<EffectiveProjectAnalysisSettings> {
    if (!projectPath) {
        return effectiveProjectAnalysisSettingsFromConfig(undefined);
    }
    return effectiveProjectAnalysisSettingsFromConfig(
        projectPath,
        await readProjectSettings(projectPath, { lenient: true }),
    );
}

export function effectiveProjectAnalysisSettingsFromConfig(
    projectPath: string | undefined,
    config: ProjectSettingsConfig = {},
): EffectiveProjectAnalysisSettings {
    const globalVisibleSeverities = visibleAnalysisSeveritiesSettingFromConfig();
    const globalUntrackedRules = untrackedAnalysisRulesSettingFromConfig();
    const globalRuleSeverityOverrides = ruleSeverityOverridesSettingFromConfig();
    if (!projectPath) {
        return {
            visibleSeverities: globalVisibleSeverities.value,
            visibleSeveritiesSource: globalVisibleSeverities.source,
            untrackedRules: globalUntrackedRules.value,
            untrackedRulesSource: globalUntrackedRules.source,
            projectUntrackedRules: [],
            ruleSeverityOverrides: globalRuleSeverityOverrides.value,
            ruleSeverityOverridesSource: globalRuleSeverityOverrides.source,
        };
    }

    const { analysis } = config;
    const visibleSeverities = resolveProjectSetting(analysis?.visibleSeverities, globalVisibleSeverities);
    const projectUntrackedRules = normalizeAnalysisRuleCodes(analysis?.untrackedRules ?? []);
    const untrackedRules = analysis?.untrackedRules === undefined || projectUntrackedRules.length === 0
        ? {
            value: globalUntrackedRules.value,
            source: globalUntrackedRules.source,
        }
        : {
            value: normalizeAnalysisRuleCodes([...globalUntrackedRules.value, ...projectUntrackedRules]),
            source: 'project' as const,
        };
    const ruleSeverityOverrides = resolveProjectSetting(
        analysis?.ruleSeverityOverrides,
        globalRuleSeverityOverrides,
    );
    return {
        visibleSeverities: visibleSeverities.value,
        visibleSeveritiesSource: visibleSeverities.source,
        untrackedRules: untrackedRules.value,
        untrackedRulesSource: untrackedRules.source,
        projectUntrackedRules,
        ruleSeverityOverrides: ruleSeverityOverrides.value,
        ruleSeverityOverridesSource: ruleSeverityOverrides.source,
    };
}

export async function setProjectAnalysisRuleTracked(
    projectPath: string,
    code: string | undefined,
    tracked: boolean,
): Promise<AnalysisRuleTrackingUpdate> {
    const normalized = normalizeAnalysisRuleCode(code);
    // Mirror the global guard: refuse to persist codes that are not known
    // diagnostic rules, so a stale/renamed code cannot linger in the sidecar.
    if (!normalized || normalizeKnownAnalysisRuleCodes([normalized]).length === 0) {
        const settings = await effectiveProjectAnalysisSettings(projectPath);
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
    await updateProjectSettings(projectPath, (existing) => {
        const update = planAnalysisRuleTrackingUpdate(existing.analysis?.untrackedRules ?? [], normalized, tracked);
        result = update;
        return update.changed || !existing.analysis?.untrackedRules
            ? withAnalysisSettings(existing, {
                ...existing.analysis,
                untrackedRules: update.untrackedRules,
            })
            : undefined;
    });
    return result;
}

export async function resetProjectAnalysisRuleTracking(
    projectPath: string,
): Promise<EffectiveProjectAnalysisSettings> {
    const updated = await updateProjectSettings(projectPath, (existing) => {
        const analysis = { ...(existing.analysis ?? {}) };
        delete analysis.untrackedRules;
        return withAnalysisSettings(existing, analysis);
    });
    return effectiveProjectAnalysisSettingsFromConfig(projectPath, updated);
}

function withAnalysisSettings(
    config: ProjectSettingsConfig,
    analysis: ProjectAnalysisSettingsConfig,
): ProjectSettingsConfig {
    // Spread the existing config so other known top-level keys (exportFolder,
    // tests, etc.) survive analysis-settings writes. Unknown top-level keys are
    // intentionally dropped on write by normalizeProjectSettingsConfig, so this
    // is not forward-compatible persistence for unrecognized keys.
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
    analysis: ProjectAnalysisSettingsConfig,
): ProjectAnalysisSettingsConfig | undefined {
    const compacted: ProjectAnalysisSettingsConfig = {};
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
