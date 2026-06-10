import {
    allowedDiagnosticSeverityOverridesForCode,
    diagnosticMetadataForCode,
    normalizeDiagnosticSeverityOverrides,
    type DiagnosticSeverityOverride,
} from './analyzer/diagnostics/ruleMetadata';

export const ANALYSIS_SEVERITIES = ['error', 'warning', 'information'] as const;
export type AnalysisSeverityFilter = typeof ANALYSIS_SEVERITIES[number];
export type AnalysisRuleSeverityOverride = DiagnosticSeverityOverride;
export type AnalysisRuleSeverityOverrides = Record<string, AnalysisRuleSeverityOverride>;

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

/**
 * Computes the shared rule-tracking mutation step for a backing store:
 * the next untracked-rule list plus whether persisting it would change
 * anything. `code` must already be normalized; store-specific bail rules
 * (unknown-code filtering, effective fallbacks) stay with the caller.
 */
export function planAnalysisRuleTrackingUpdate(
    untrackedRules: readonly string[],
    code: string,
    tracked: boolean,
): AnalysisRuleTrackingUpdate {
    const next = setAnalysisRuleTrackedInList(untrackedRules, code, tracked);
    const changed = untrackedRules.length !== next.length
        || untrackedRules.some((entry, index) => entry !== next[index]);
    return {
        code,
        tracked,
        changed,
        untrackedRules: next,
    };
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

export function normalizeKnownAnalysisRuleCodes(value: unknown): string[] {
    return normalizeAnalysisRuleCodes(value)
        .filter((code) => diagnosticMetadataForCode(code) !== undefined);
}

export function normalizeAnalysisRuleCode(code: unknown): string | undefined {
    return typeof code === 'string' ? code.trim().toLowerCase() || undefined : undefined;
}

export function normalizeAnalysisRuleSeverityOverrides(value: unknown): AnalysisRuleSeverityOverrides {
    return normalizeDiagnosticSeverityOverrides(value);
}

export function allowedAnalysisRuleSeverityOverrides(
    code: string | undefined,
): readonly AnalysisRuleSeverityOverride[] {
    return allowedDiagnosticSeverityOverridesForCode(code);
}

/**
 * Validates rule-severity-override entries against the known rule set,
 * reporting each invalid entry as `Expected "<prefix>.<rawCode>" <requirement>`
 * via the caller-supplied sink (throwing or collecting). Returns the
 * normalized valid entries.
 */
export function validateAnalysisRuleSeverityOverrideEntries(
    value: Record<string, unknown>,
    reportEntry: (rawCode: string, requirement: string) => void,
): AnalysisRuleSeverityOverrides {
    const parsed: Record<string, AnalysisRuleSeverityOverride> = {};
    for (const [rawCode, rawSeverity] of Object.entries(value)) {
        const code = normalizeAnalysisRuleCode(rawCode);
        const allowed = allowedAnalysisRuleSeverityOverrides(code);
        if (!code || allowed.length === 0) {
            reportEntry(rawCode, 'to target a known analysis rule that permits severity overrides.');
            continue;
        }
        if (typeof rawSeverity !== 'string' || !allowed.includes(rawSeverity as AnalysisRuleSeverityOverride)) {
            reportEntry(rawCode, `to be one of: ${allowed.join(', ')}.`);
            continue;
        }
        parsed[code] = rawSeverity as AnalysisRuleSeverityOverride;
    }
    return normalizeAnalysisRuleSeverityOverrides(parsed);
}
