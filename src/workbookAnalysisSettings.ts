import {
    normalizeAnalysisRuleCode,
    normalizeAnalysisVisibleSeverities,
    setAnalysisRuleTrackedInList,
    type AnalysisRuleTrackingUpdate,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
    readWorkbookSettings,
    writeWorkbookSettings,
    type WorkbookAnalysisSettingsConfig,
    type WorkbookSettingsConfig,
} from './moduleExport';
import {
    untrackedAnalysisRulesFromConfig,
    visibleAnalysisSeveritiesFromConfig,
} from './analysisOptions';

export type WorkbookAnalysisSettingsSource = 'workbook' | 'global';

export interface EffectiveWorkbookAnalysisSettings {
    visibleSeverities: AnalysisSeverityFilter[];
    visibleSeveritiesSource: WorkbookAnalysisSettingsSource;
    untrackedRules: string[];
    untrackedRulesSource: WorkbookAnalysisSettingsSource;
}

export async function effectiveWorkbookAnalysisSettings(
    workbookPath: string | undefined,
): Promise<EffectiveWorkbookAnalysisSettings> {
    const globalVisibleSeverities = visibleAnalysisSeveritiesFromConfig();
    const globalUntrackedRules = untrackedAnalysisRulesFromConfig();
    if (!workbookPath) {
        return {
            visibleSeverities: globalVisibleSeverities,
            visibleSeveritiesSource: 'global',
            untrackedRules: globalUntrackedRules,
            untrackedRulesSource: 'global',
        };
    }

    const config = await readWorkbookSettings(workbookPath);
    const analysis = config.analysis;
    return {
        visibleSeverities: analysis?.visibleSeverities ?? globalVisibleSeverities,
        visibleSeveritiesSource: analysis?.visibleSeverities ? 'workbook' : 'global',
        untrackedRules: analysis?.untrackedRules ?? globalUntrackedRules,
        untrackedRulesSource: analysis?.untrackedRules ? 'workbook' : 'global',
    };
}

export async function setWorkbookAnalysisVisibleSeverities(
    workbookPath: string,
    severities: unknown,
): Promise<EffectiveWorkbookAnalysisSettings> {
    const existing = await readWorkbookSettings(workbookPath);
    await writeWorkbookSettings(workbookPath, withAnalysisSettings(existing, {
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

    const existing = await readWorkbookSettings(workbookPath);
    const current = existing.analysis?.untrackedRules ?? untrackedAnalysisRulesFromConfig();
    const next = setAnalysisRuleTrackedInList(current, normalized, tracked);
    const changed = current.length !== next.length || current.some((entry, index) => entry !== next[index]);
    if (changed || !existing.analysis?.untrackedRules) {
        await writeWorkbookSettings(workbookPath, withAnalysisSettings(existing, {
            ...existing.analysis,
            untrackedRules: next,
        }));
    }
    return {
        code: normalized,
        tracked,
        changed,
        untrackedRules: next,
    };
}

function withAnalysisSettings(
    config: WorkbookSettingsConfig,
    analysis: WorkbookAnalysisSettingsConfig,
): WorkbookSettingsConfig {
    return {
        exportFolder: config.exportFolder,
        exportMode: config.exportMode,
        importMode: config.importMode,
        analysis,
    };
}
