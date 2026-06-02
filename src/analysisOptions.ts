import * as vscode from 'vscode';

export const ANALYSIS_SEVERITIES = ['error', 'warning', 'information'] as const;
export type AnalysisSeverityFilter = typeof ANALYSIS_SEVERITIES[number];

export function visibleAnalysisSeveritiesFromConfig(): AnalysisSeverityFilter[] {
    const configured = vscode.workspace
        .getConfiguration('xlide')
        .get<string[]>('analysis.visibleSeverities', [...ANALYSIS_SEVERITIES]);
    return normalizeAnalysisVisibleSeverities(configured);
}

export function normalizeAnalysisVisibleSeverities(value: unknown): AnalysisSeverityFilter[] {
    if (!Array.isArray(value)) {
        return [...ANALYSIS_SEVERITIES];
    }
    const allowed = new Set<string>(ANALYSIS_SEVERITIES);
    return value.filter((entry): entry is AnalysisSeverityFilter =>
        typeof entry === 'string' && allowed.has(entry),
    );
}

export function untrackedAnalysisRulesFromConfig(): string[] {
    const configured = vscode.workspace
        .getConfiguration('xlide')
        .get<string[]>('analysis.untrackedRules', []);
    return normalizeAnalysisRuleCodes(configured);
}

export interface AnalysisRuleTrackingUpdate {
    code?: string;
    tracked: boolean;
    changed: boolean;
    untrackedRules: string[];
}

export async function setAnalysisRuleTrackedInConfig(
    code: string | undefined,
    tracked: boolean,
): Promise<AnalysisRuleTrackingUpdate> {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        return {
            tracked,
            changed: false,
            untrackedRules: untrackedAnalysisRulesFromConfig(),
        };
    }
    const current = untrackedAnalysisRulesFromConfig();
    const next = setAnalysisRuleTrackedInList(current, normalized, tracked);
    const changed = current.length !== next.length || current.some((entry, index) => entry !== next[index]);
    if (changed) {
        await vscode.workspace
            .getConfiguration('xlide')
            .update('analysis.untrackedRules', next, vscode.ConfigurationTarget.Global);
    }
    return {
        code: normalized,
        tracked,
        changed,
        untrackedRules: next,
    };
}

export function setAnalysisRuleTrackedInList(
    untrackedRules: readonly string[],
    code: string | undefined,
    tracked: boolean,
): string[] {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        return normalizeAnalysisRuleCodes(untrackedRules);
    }
    return tracked
        ? normalizeAnalysisRuleCodes(untrackedRules.filter((entry) => normalizeAnalysisRuleCode(entry) !== normalized))
        : normalizeAnalysisRuleCodes([...untrackedRules, normalized]);
}

export function isAnalysisRuleTracked(
    code: string | undefined,
    untrackedRules: readonly string[] | ReadonlySet<string>,
): boolean {
    const normalized = normalizeAnalysisRuleCode(code);
    if (!normalized) {
        return true;
    }
    const rules = new Set(
        Array.from(untrackedRules)
            .map(normalizeAnalysisRuleCode)
            .filter((entry): entry is string => Boolean(entry)),
    );
    return !rules.has(normalized);
}

export function normalizeAnalysisRuleCodes(value: unknown): string[] {
    const incoming = Array.isArray(value) ? value : [];
    return [...new Set(incoming
        .map(normalizeAnalysisRuleCode)
        .filter((entry): entry is string => Boolean(entry)))]
        .sort();
}

function normalizeAnalysisRuleCode(code: unknown): string | undefined {
    return typeof code === 'string' ? code.trim().toLowerCase() || undefined : undefined;
}
