import * as vscode from 'vscode';
import {
    ANALYSIS_SEVERITIES,
    normalizeAnalysisRuleCodes,
    normalizeAnalysisVisibleSeverities,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
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
    const configured = vscode.workspace
        .getConfiguration('xlide')
        .get<string[]>('analysis.visibleSeverities', [...ANALYSIS_SEVERITIES]);
    return normalizeAnalysisVisibleSeverities(configured);
}

export function untrackedAnalysisRulesFromConfig(): string[] {
    const configured = vscode.workspace
        .getConfiguration('xlide')
        .get<string[]>('analysis.untrackedRules', []);
    return normalizeAnalysisRuleCodes(configured);
}
