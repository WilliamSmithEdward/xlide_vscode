import * as vscode from 'vscode';
import {
    setXlideGlobalAnalysisRuleTracked,
    xlideAnalysisRuleSeveritiesFromConfig,
    xlideAnalysisUntrackedRulesFromConfig,
    xlideAnalysisVisibleSeveritiesFromConfig,
} from './globalSettings';
export {
    ANALYSIS_SEVERITIES,
    isAnalysisRuleTracked,
    normalizeAnalysisRuleCode,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisVisibleSeverities,
    setAnalysisRuleTrackedInList,
} from './analysisSettingsCore';
export type {
    AnalysisRuleTrackingUpdate,
    AnalysisSeverityFilter,
} from './analysisSettingsCore';

export function visibleAnalysisSeveritiesSettingFromConfig() {
    return xlideAnalysisVisibleSeveritiesFromConfig(vscode.workspace.getConfiguration('xlide'));
}

export function untrackedAnalysisRulesSettingFromConfig() {
    return xlideAnalysisUntrackedRulesFromConfig(vscode.workspace.getConfiguration('xlide'));
}

export function ruleSeverityOverridesSettingFromConfig() {
    return xlideAnalysisRuleSeveritiesFromConfig(vscode.workspace.getConfiguration('xlide'));
}

export function setGlobalAnalysisRuleTracked(code: string | undefined, tracked: boolean) {
    return setXlideGlobalAnalysisRuleTracked(vscode.workspace.getConfiguration('xlide'), code, tracked);
}
