import * as vscode from 'vscode';
import {
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import {
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

export function visibleAnalysisSeveritiesFromConfig(): AnalysisSeverityFilter[] {
    return visibleAnalysisSeveritiesSettingFromConfig().value;
}

export function untrackedAnalysisRulesFromConfig(): string[] {
    return untrackedAnalysisRulesSettingFromConfig().value;
}

export function ruleSeverityOverridesFromConfig(): AnalysisRuleSeverityOverrides {
    return ruleSeverityOverridesSettingFromConfig().value;
}

export function visibleAnalysisSeveritiesSettingFromConfig() {
    return xlideAnalysisVisibleSeveritiesFromConfig(vscode.workspace.getConfiguration('xlide'));
}

export function untrackedAnalysisRulesSettingFromConfig() {
    return xlideAnalysisUntrackedRulesFromConfig(vscode.workspace.getConfiguration('xlide'));
}

export function ruleSeverityOverridesSettingFromConfig() {
    return xlideAnalysisRuleSeveritiesFromConfig(vscode.workspace.getConfiguration('xlide'));
}
