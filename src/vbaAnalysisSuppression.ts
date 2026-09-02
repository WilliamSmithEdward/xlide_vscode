import { parseModule } from './analyzer/parser/parseModule';
import type { BodyNode, ModuleMember, Span } from './analyzer/parser/nodes';
import type { ProjectAnalysisSuppressScope } from './projectAnalysisWebview';

/**
 * Pure helpers that locate where an analysis-suppression directive should be
 * inserted for a problem at a given offset. They operate only on parsed VBA
 * source (no vscode runtime dependencies), so they can be unit-tested in
 * isolation from the command handlers that apply the edits.
 */

export type AnalysisSuppressionInsertionTarget =
    | { kind: 'module'; startLine: number }
    | { kind: 'member'; startLine: number }
    | { kind: 'block'; startLine: number; endLine: number };

type SuppressibleMember = Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }>;
type BlockBodyNode = Extract<BodyNode, { body: BodyNode[] }>;

export function suppressionTargetForProblem(
    source: string,
    starts: readonly number[],
    problemOffset: number,
    scope: ProjectAnalysisSuppressScope,
): AnalysisSuppressionInsertionTarget {
    if (scope === 'module') {
        return { kind: 'module', startLine: moduleSuppressionInsertLine(source) };
    }

    const parsed = parseModule(source);
    const member = containingSuppressibleMember(parsed.members, problemOffset);
    if (!member) {
        throw new Error(`No containing Sub, Function, Property, Type, or Enum was found for this analysis finding.`);
    }

    if (scope === 'member') {
        return {
            kind: 'member',
            startLine: lineForOffset(starts, member.span.start),
        };
    }

    if (member.kind !== 'Procedure') {
        throw new Error('No containing executable block was found for this analysis finding.');
    }
    const block = closestContainingBlock(member.body, problemOffset);
    if (!block) {
        throw new Error('No containing executable block was found for this analysis finding.');
    }

    return {
        kind: 'block',
        startLine: lineForOffset(starts, block.span.start),
        endLine: lineForOffset(starts, Math.max(block.span.start, block.span.end - 1)),
    };
}

function moduleSuppressionInsertLine(source: string): number {
    const lines = source.split(/\r\n|\r|\n/);
    let line = 0;
    while (line < lines.length && /^\s*Attribute\b/i.test(lines[line])) {
        line++;
    }
    return line;
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

export function spanContainsOffset(span: Span, offset: number): boolean {
    return offset >= span.start && offset < Math.max(span.end, span.start + 1);
}

export function spanLength(span: Span): number {
    return Math.max(1, span.end - span.start);
}

function lineForOffset(starts: readonly number[], offset: number): number {
    let lo = 0;
    let hi = starts.length - 1;
    while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (starts[mid] <= offset) {
            lo = mid;
        } else {
            hi = mid - 1;
        }
    }
    return lo;
}
