import { tokenize } from '../lexer/tokenize';
import type { ModuleMember, Span } from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import type { VbaDiagnostic } from './analyzeModule';
import { diagnosticMetadataForCode, DIAGNOSTIC_RULES } from './ruleMetadata';
import { lineStartOffsets } from '../../vbaSourceScan';

export const ANALYSIS_SUPPRESSION_DIRECTIVE_CODE = DIAGNOSTIC_RULES.analysisSuppressionDirective.code;

export interface AnalysisSuppressionAnalysis {
	diagnostics: readonly VbaDiagnostic[];
	isSuppressed(code: string | undefined, zeroBasedLine: number): boolean;
	isDiagnosticSuppressed(code: string | undefined, span: Span): boolean;
}

export interface AnalysisSuppressionFilterResult<T> {
	diagnostics: T[];
	directiveDiagnostics: readonly VbaDiagnostic[];
	suppressedCount: number;
	analysis: AnalysisSuppressionAnalysis;
}

type SuppressionTarget = 'all' | ReadonlySet<string>;

interface MutableSuppressions {
	file?: SuppressionTarget;
	linesAll: Set<number>;
	linesByCode: Map<number, Set<string>>;
	members: SuppressionRange[];
	blocks: SuppressionRange[];
	diagnostics: VbaDiagnostic[];
}

interface ParsedDirective {
	action:
		| 'disable-file'
		| 'disable-line'
		| 'disable-next-line'
		| 'disable-next-member'
		| 'disable-block'
		| 'enable-block';
	target: SuppressionTarget | undefined;
}

interface SuppressionRange {
	span: Span;
	target: SuppressionTarget;
}

interface OpenBlockSuppression {
	start: number;
	target: SuppressionTarget;
	directiveSpan: Span;
}

interface LineTokenSummary {
	line: number;
	hasSourceToken: boolean;
	isAttributeLine: boolean;
}

/**
 * Parses XLIDE analysis suppression directives:
 * apostrophe comments only, no doc comments, no Rem comments, and physical-line
 * line directives. Member and block directives are lexical source ranges. The
 * resulting predicate is intentionally shared by live diagnostics, workbook
 * analysis, and tests so suppression semantics cannot drift by surface.
 */
export function scanAnalysisSuppressions(source: string): AnalysisSuppressionAnalysis {
	const tokens = tokenize(source);
	const lineStarts = lineStartOffsets(source);
	const firstSourceLine = firstNonCommentNonAttributeLine(tokens);
	const members = suppressibleMembers(source);
	const openBlocks: OpenBlockSuppression[] = [];
	const state: MutableSuppressions = {
		linesAll: new Set<number>(),
		linesByCode: new Map<number, Set<string>>(),
		members: [],
		blocks: [],
		diagnostics: [],
	};

	for (const token of tokens) {
		if (token.kind !== 'comment' || !isApostropheDirectiveCandidate(token.rawText)) {
			continue;
		}
		const parsed = parseDirectiveComment(token.rawText, { start: token.start, end: token.end });
		if (!parsed) {
			continue;
		}
		state.diagnostics.push(...parsed.diagnostics);
		if (!parsed.directive || !parsed.directive.target) {
			continue;
		}
		if (parsed.directive.action === 'disable-file') {
			if (firstSourceLine !== undefined && token.line >= firstSourceLine) {
				state.diagnostics.push(directiveDiagnostic(
					'@xlide-analysis-disable-file must appear before the first non-comment, non-attribute source line.',
					{ start: token.start, end: token.end },
				));
				continue;
			}
			state.file = mergeTargets(state.file, parsed.directive.target);
			continue;
		}
		if (parsed.directive.action === 'disable-next-member') {
			const member = nextDirectSuppressibleMember(tokens, members, token.end);
			if (!member) {
				state.diagnostics.push(directiveDiagnostic(
					'@xlide-analysis-disable-next-member must appear directly before a Sub, Function, Property, Type, or Enum declaration.',
					{ start: token.start, end: token.end },
				));
				continue;
			}
			state.members.push({ span: member.span, target: parsed.directive.target });
			continue;
		}
		if (parsed.directive.action === 'disable-block') {
			openBlocks.push({
				start: token.end,
				target: parsed.directive.target,
				directiveSpan: { start: token.start, end: token.end },
			});
			continue;
		}
		if (parsed.directive.action === 'enable-block') {
			const open = openBlocks[openBlocks.length - 1];
			if (!open) {
				state.diagnostics.push(directiveDiagnostic(
					'@xlide-analysis-enable-block has no matching @xlide-analysis-disable-block.',
					{ start: token.start, end: token.end },
				));
				continue;
			}
			if (!targetsEqual(open.target, parsed.directive.target)) {
				state.diagnostics.push(directiveDiagnostic(
					'@xlide-analysis-enable-block code list must match the innermost open @xlide-analysis-disable-block.',
					{ start: token.start, end: token.end },
				));
				continue;
			}
			openBlocks.pop();
			state.blocks.push({
				span: { start: open.start, end: token.start },
				target: open.target,
			});
			continue;
		}
		const targetLine = parsed.directive.action === 'disable-line'
			? token.line
			: token.line + 1;
		addLineSuppression(state, targetLine, parsed.directive.target);
	}

	for (const open of openBlocks) {
		state.diagnostics.push(directiveDiagnostic(
			'@xlide-analysis-disable-block is missing a matching @xlide-analysis-enable-block.',
			open.directiveSpan,
		));
	}

	return {
		diagnostics: state.diagnostics,
		isSuppressed(code, zeroBasedLine) {
			const lineStart = lineStarts[zeroBasedLine];
			if (lineStart === undefined) {
				return false;
			}
			return diagnosticSuppressedAt(state, code, zeroBasedLine, {
				start: lineStart,
				end: lineStart,
			});
		},
		isDiagnosticSuppressed(code, span) {
			const line = lineForOffset(lineStarts, span.start);
			return diagnosticSuppressedAt(state, code, line, span);
		},
	};
}

function diagnosticSuppressedAt(
	state: MutableSuppressions,
	code: string | undefined,
	zeroBasedLine: number,
	span: Span,
): boolean {
	if (zeroBasedLine < 0 || code === ANALYSIS_SUPPRESSION_DIRECTIVE_CODE) {
		return false;
	}
	const normalized = normalizeCode(code);
	if (targetSuppresses(state.file, normalized)) {
		return true;
	}
	if (state.linesAll.has(zeroBasedLine)) {
		return true;
	}
	const lineCodes = state.linesByCode.get(zeroBasedLine);
	if (lineCodes && normalized && lineCodes.has(normalized)) {
		return true;
	}
	if (state.members.some((range) => rangeSuppresses(range, normalized, span))) {
		return true;
	}
	if (state.blocks.some((range) => rangeSuppresses(range, normalized, span))) {
		return true;
	}
	return false;
}

export function filterDiagnosticsWithSuppressions<T extends { code?: string; span: Span }>(
	source: string,
	diagnostics: readonly T[],
): AnalysisSuppressionFilterResult<T> {
	const analysis = scanAnalysisSuppressions(source);
	const filtered: T[] = [];
	let suppressedCount = 0;

	for (const diagnostic of diagnostics) {
		if (analysis.isDiagnosticSuppressed(diagnostic.code, diagnostic.span)) {
			suppressedCount++;
			continue;
		}
		filtered.push(diagnostic);
	}

	return {
		diagnostics: filtered,
		directiveDiagnostics: analysis.diagnostics,
		suppressedCount,
		analysis,
	};
}

function isApostropheDirectiveCandidate(rawText: string): boolean {
	return rawText.startsWith("'") && !rawText.startsWith("'''");
}

function parseDirectiveComment(
	rawText: string,
	span: Span,
): { directive?: ParsedDirective; diagnostics: VbaDiagnostic[] } | undefined {
	const body = rawText.slice(1).trimStart();
	if (!/^@xlide-analysis-/i.test(body)) {
		return undefined;
	}

	const match = /^@xlide-analysis-(disable-file|disable-line|disable-next-line|disable-next-member|disable-block|enable-block)\b(.*)$/i.exec(body);
	if (!match) {
		return {
			diagnostics: [directiveDiagnostic(
				'Unknown XLIDE analysis suppression directive. Supported directives are disable-file, disable-line, disable-next-line, disable-next-member, disable-block, and enable-block.',
				span,
			)],
		};
	}

	const action = match[1].toLowerCase() as ParsedDirective['action'];
	const target = parseTargetList(match[2], span);
	return {
		directive: { action, target: target.target },
		diagnostics: target.diagnostics,
	};
}

function parseTargetList(
	rawText: string,
	span: Span,
): { target: SuppressionTarget | undefined; diagnostics: VbaDiagnostic[] } {
	const diagnostics: VbaDiagnostic[] = [];
	const text = rawText.split('--', 1)[0].trim();
	if (!text) {
		return { target: 'all', diagnostics };
	}

	const parts = text.split(',').map((part) => part.trim()).filter((part) => part.length > 0);
	if (parts.length === 0 || parts.length !== text.split(',').length) {
		return {
			target: undefined,
			diagnostics: [directiveDiagnostic(
				'Malformed XLIDE analysis suppression code list. Use "all" or comma-separated diagnostic codes.',
				span,
			)],
		};
	}

	const normalized = parts.map(normalizeCode);
	if (normalized.some((part) => !part || /\s/.test(part))) {
		return {
			target: undefined,
			diagnostics: [directiveDiagnostic(
				'Malformed XLIDE analysis suppression code list. Use comma-separated diagnostic codes with no embedded whitespace.',
				span,
			)],
		};
	}

	const includesAll = normalized.includes('all');
	if (includesAll && normalized.length > 1) {
		return {
			target: undefined,
			diagnostics: [directiveDiagnostic(
				'Malformed XLIDE analysis suppression code list. Use "all" by itself, or list specific diagnostic codes.',
				span,
			)],
		};
	}
	if (includesAll) {
		return { target: 'all', diagnostics };
	}

	const knownCodes = new Set<string>();
	for (const code of normalized) {
		if (!code) {
			continue;
		}
		if (!diagnosticMetadataForCode(code)) {
			diagnostics.push(directiveDiagnostic(
				`Unknown XLIDE analysis diagnostic code '${code}'.`,
				span,
			));
			continue;
		}
		knownCodes.add(code);
	}

	return {
		target: knownCodes.size > 0 ? knownCodes : undefined,
		diagnostics,
	};
}

function addLineSuppression(
	state: MutableSuppressions,
	line: number,
	target: SuppressionTarget,
): void {
	if (target === 'all') {
		state.linesAll.add(line);
		state.linesByCode.delete(line);
		return;
	}
	if (state.linesAll.has(line)) {
		return;
	}
	let codes = state.linesByCode.get(line);
	if (!codes) {
		codes = new Set<string>();
		state.linesByCode.set(line, codes);
	}
	for (const code of target) {
		codes.add(code);
	}
}

function mergeTargets(
	current: SuppressionTarget | undefined,
	next: SuppressionTarget,
): SuppressionTarget {
	if (current === 'all' || next === 'all') {
		return 'all';
	}
	const merged = new Set<string>(current);
	for (const code of next) {
		merged.add(code);
	}
	return merged;
}

function targetSuppresses(target: SuppressionTarget | undefined, code: string | undefined): boolean {
	if (!target) {
		return false;
	}
	if (target === 'all') {
		return true;
	}
	return !!code && target.has(code);
}

function rangeSuppresses(range: SuppressionRange, code: string | undefined, span: Span): boolean {
	return span.start >= range.span.start &&
		span.start < range.span.end &&
		targetSuppresses(range.target, code);
}

function targetsEqual(left: SuppressionTarget, right: SuppressionTarget): boolean {
	if (left === 'all' || right === 'all') {
		return left === right;
	}
	if (left.size !== right.size) {
		return false;
	}
	for (const code of left) {
		if (!right.has(code)) {
			return false;
		}
	}
	return true;
}

function normalizeCode(code: string | undefined): string | undefined {
	return code?.trim().toLowerCase();
}

function directiveDiagnostic(message: string, span: Span): VbaDiagnostic {
	const meta = DIAGNOSTIC_RULES.analysisSuppressionDirective;
	return {
		code: meta.code,
		message,
		severity: meta.defaultSeverity,
		span,
		specReference: meta.specReference,
	};
}

function firstNonCommentNonAttributeLine(
	tokens: ReturnType<typeof tokenize>,
): number | undefined {
	const summaries = new Map<number, LineTokenSummary>();
	for (const token of tokens) {
		if (token.kind === 'newline' || token.kind === 'comment') {
			continue;
		}
		let summary = summaries.get(token.line);
		if (!summary) {
			summary = {
				line: token.line,
				hasSourceToken: false,
				isAttributeLine: false,
			};
			summaries.set(token.line, summary);
		}
		if (!summary.hasSourceToken) {
			summary.isAttributeLine = tokenText(token) === 'attribute';
		}
		summary.hasSourceToken = true;
	}

	return [...summaries.values()]
		.sort((a, b) => a.line - b.line)
		.find((line) => line.hasSourceToken && !line.isAttributeLine)
		?.line;
}

function suppressibleMembers(source: string): Array<Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }>> {
	return parseModule(source).members
		.filter((member): member is Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }> =>
			member.kind === 'Procedure' || member.kind === 'Type' || member.kind === 'Enum',
		)
		.sort((a, b) => a.span.start - b.span.start);
}

function nextDirectSuppressibleMember(
	tokens: ReturnType<typeof tokenize>,
	members: readonly Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }>[],
	afterOffset: number,
): Extract<ModuleMember, { kind: 'Procedure' | 'Type' | 'Enum' }> | undefined {
	const nextToken = tokens.find((token) =>
		token.start >= afterOffset &&
		token.kind !== 'comment' &&
		token.kind !== 'newline',
	);
	if (!nextToken) {
		return undefined;
	}
	return members.find((member) => member.span.start === nextToken.start);
}

function tokenText(token: ReturnType<typeof tokenize>[number]): string {
	return (token.canonicalText ?? token.rawText).toLowerCase();
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
