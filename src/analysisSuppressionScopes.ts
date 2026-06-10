import {
    diagnosticSuppressionScopesForCode,
    type DiagnosticSuppressionScope,
} from './analyzer/diagnostics/ruleMetadata';
import type { BodyNode, ModuleMember, Span } from './analyzer/parser/nodes';
import { parseModule } from './analyzer/parser/parseModule';

type SuppressibleMember = Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }>;
type BlockBodyNode = Extract<BodyNode, { body: BodyNode[] }>;

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

function containingSuppressibleMember(
    members: readonly ModuleMember[],
    offset: number,
): SuppressibleMember | undefined {
    return members
        .filter((member): member is SuppressibleMember =>
            member.kind === 'Procedure' || member.kind === 'Type' || member.kind === 'Enum',
        )
        .filter((member) => spanContainsOffset(member.span, offset))
        .sort((left, right) => spanLength(left.span) - spanLength(right.span))[0];
}

function closestContainingBlock(nodes: readonly BodyNode[], offset: number): BlockBodyNode | undefined {
    let best: BlockBodyNode | undefined;
    for (const node of nodes) {
        if (!isBlockBodyNode(node) || !spanContainsOffset(node.span, offset)) {
            continue;
        }
        const nested = closestContainingBlock(node.body, offset);
        const candidate = nested ?? node;
        if (!best || spanLength(candidate.span) < spanLength(best.span)) {
            best = candidate;
        }
    }
    return best;
}

function isBlockBodyNode(node: BodyNode): node is BlockBodyNode {
    return node.kind === 'IfBlock' ||
        node.kind === 'ForBlock' ||
        node.kind === 'DoBlock' ||
        node.kind === 'WhileBlock' ||
        node.kind === 'WithBlock' ||
        node.kind === 'SelectBlock';
}

function spanContainsOffset(span: Span, offset: number): boolean {
    return offset >= span.start && offset < Math.max(span.end, span.start + 1);
}

function spanLength(span: Span): number {
    return Math.max(1, span.end - span.start);
}
