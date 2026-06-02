export const ANALYSIS_SEVERITIES = ['error', 'warning', 'information'] as const;
export type AnalysisSeverityFilter = typeof ANALYSIS_SEVERITIES[number];

export interface AnalysisRuleTrackingUpdate {
    code?: string;
    tracked: boolean;
    changed: boolean;
    untrackedRules: string[];
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

export function normalizeAnalysisRuleCode(code: unknown): string | undefined {
    return typeof code === 'string' ? code.trim().toLowerCase() || undefined : undefined;
}
