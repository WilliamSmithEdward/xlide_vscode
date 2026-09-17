import {
    diagnosticSuppressionScopesForCode,
    type DiagnosticSuppressionScope,
} from './analyzer/diagnostics/ruleMetadata';
import type { ModuleMember } from './analyzer/parser/nodes';
import { parseModule } from './analyzer/parser/parseModule';
import {
    closestContainingBlock,
    containingSuppressibleMember,
    type SuppressibleMember,
} from './vbaAnalysisSuppression';

export type AnalysisSuppressionScope = DiagnosticSuppressionScope;

export type AnalysisSuppressionScopeResolver = (
    code: string | undefined,
    offset: number,
) => AnalysisSuppressionScope[];

/**
 * Returns a resolver that parses `source` lazily and at most once, so scope
 * lookups for many diagnostics of the same module share a single parse.
 */
export function analysisSuppressionScopeResolver(source: string): AnalysisSuppressionScopeResolver {
    let members: readonly ModuleMember[] | undefined;
    const parsedMembers = (): readonly ModuleMember[] => {
        if (!members) {
            try {
                members = parseModule(source).members;
            } catch {
                members = [];
            }
        }
        return members;
    };

    return (code, offset) => {
        const allowed = new Set(diagnosticSuppressionScopesForCode(code));
        const scopes: AnalysisSuppressionScope[] = [];
        let member: SuppressibleMember | undefined;

        if (allowed.has('member') || allowed.has('block')) {
            member = containingSuppressibleMember(parsedMembers(), offset);
        }

        if (allowed.has('block') && member?.kind === 'Procedure' && closestContainingBlock(member.body, offset)) {
            scopes.push('block');
        }
        if (allowed.has('member') && member) {
            scopes.push('member');
        }
        if (allowed.has('module')) {
            scopes.push('module');
        }
        return scopes;
    };
}

export function validAnalysisSuppressionScopesForDiagnostic(
    source: string,
    code: string | undefined,
    offset: number,
): AnalysisSuppressionScope[] {
    return analysisSuppressionScopeResolver(source)(code, offset);
}
