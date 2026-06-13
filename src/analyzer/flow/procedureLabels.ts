import type { VbaToken } from '../lexer/tokenKinds';
import {
	splitTopLevelTokenGroups,
	statementTokens,
	tokenName,
	tokensWithoutLeadingLineNumber,
	tokenWord,
} from '../lexer/tokenHelpers';
import { parseModule } from '../parser/parseModule';
import type {
	BodyNode,
	LeafStatementNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';

export interface VbaProcedureLabel {
	key: string;
	text: string;
	span: Span;
	kind: 'name' | 'line';
}

export interface VbaProcedureLabelReference extends VbaProcedureLabel {
	statementKind: 'goto' | 'gosub' | 'resume' | 'on-error-goto' | 'on-goto' | 'on-gosub';
}

export interface VbaProcedureLabelCompletion {
	label: string;
	kind: VbaProcedureLabel['kind'];
	detail: string;
}

export interface VbaProcedureLabelDefinition {
	procedure: ProcedureNode;
	reference: VbaProcedureLabelReference;
	label: VbaProcedureLabel;
}

export function collectProcedureLabels(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): Map<string, VbaProcedureLabel> {
	const labels = new Map<string, VbaProcedureLabel>();
	for (const label of collectProcedureLabelDeclarations(source, procedure, activity)) {
		if (!labels.has(label.key)) {
			labels.set(label.key, label);
		}
	}
	return labels;
}

export function collectProcedureLabelDeclarations(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): VbaProcedureLabel[] {
	const labels: VbaProcedureLabel[] = [];
	forEachProcedureStatement(procedure.body, (stmt) => {
		const label = statementLabelDeclaration(source, stmt.span);
		if (label) {
			labels.push(label);
		}
	}, activity);
	return labels;
}

export function collectProcedureLabelReferences(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): VbaProcedureLabelReference[] {
	const refs: VbaProcedureLabelReference[] = [];
	forEachProcedureStatement(procedure.body, (stmt) => {
		refs.push(...statementLabelReferences(source, stmt.span));
	}, activity);
	return refs;
}

export function resolveProcedureLabelCompletions(
	source: string,
	offset: number,
): VbaProcedureLabelCompletion[] {
	const module = parseModule(source);
	const procedure = procedureAtOffset(module, offset);
	if (!procedure || !isProcedureLabelCompletionContext(source, procedure, offset)) {
		return [];
	}
	return [...collectProcedureLabels(source, procedure).values()]
		.sort(compareLabels)
		.map((label) => ({
			label: label.text,
			kind: label.kind,
			detail: label.kind === 'line' ? 'Procedure line label' : 'Procedure label',
		}));
}

export function resolveProcedureLabelDefinitionAt(
	source: string,
	offset: number,
): VbaProcedureLabelDefinition | undefined {
	const module = parseModule(source);
	const procedure = procedureAtOffset(module, offset);
	if (!procedure) {
		return undefined;
	}
	const reference = collectProcedureLabelReferences(source, procedure)
		.find((ref) => offsetInSpan(offset, ref.span));
	if (!reference) {
		return undefined;
	}
	const label = collectProcedureLabels(source, procedure).get(reference.key);
	return label ? { procedure, reference, label } : undefined;
}

export function statementLabelReferences(
	source: string,
	span: Span,
): VbaProcedureLabelReference[] {
	const toks = tokensWithoutLeadingLineNumber(statementTokens(source, span));
	if (toks.length === 0) {
		return [];
	}
	if (tokenWord(toks[0]) === 'on') {
		return onStatementLabelReferences(toks, span);
	}
	const refs: VbaProcedureLabelReference[] = [];
	for (let i = 0; i < toks.length; i++) {
		const word = tokenWord(toks[i]);
		if (word === 'goto' || word === 'gosub') {
			if (word === 'goto' && isOnErrorGotoDisableAt(toks, i)) {
				continue;
			}
			const ref = labelReferenceAfter(toks, span, i + 1, word);
			if (ref) {
				refs.push(ref);
			}
			continue;
		}
		if (word === 'resume') {
			const nextWord = tokenWord(toks[i + 1]);
			if (!toks[i + 1] || nextWord === 'next') {
				continue;
			}
			const ref = labelReferenceAfter(toks, span, i + 1, 'resume');
			if (ref) {
				refs.push(ref);
			}
		}
	}
	return refs;
}

function isProcedureLabelCompletionContext(
	source: string,
	procedure: ProcedureNode,
	offset: number,
): boolean {
	const statement = statementAtOffset(procedure.body, offset);
	const span = statement?.span ?? physicalLineSpanAtOffset(source, offset);
	const end = Math.max(span.start, Math.min(offset, span.end));
	const localPrefix = source.slice(span.start, end);
	const rawTokens = statementTokens(localPrefix, { start: 0, end: localPrefix.length });
	return isLabelTargetPrefix(rawTokens, localPrefix.length);
}

function isLabelTargetPrefix(tokens: readonly VbaToken[], prefixLength: number): boolean {
	let toks = tokensWithoutLeadingLineNumber(
		tokens.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline'),
	);
	if (toks.length === 0) {
		return false;
	}
	const partial = removablePartialToken(toks, prefixLength);
	if (partial) {
		if (isOnErrorGotoDisablePartial(toks, partial)) {
			return false;
		}
		if (tokenWord(partial) === 'next' && tokenWord(toks[toks.length - 2]) === 'resume') {
			return false;
		}
		toks = toks.slice(0, -1);
	}
	const lastColon = toks.map((tok) => tok.rawText).lastIndexOf(':');
	if (lastColon >= 0) {
		toks = toks.slice(lastColon + 1);
	}
	if (toks.length === 0) {
		return false;
	}
	const lastWord = tokenWord(toks[toks.length - 1]);
	if (lastWord === 'goto' || lastWord === 'gosub' || lastWord === 'resume') {
		return true;
	}
	if (lastWord === ',') {
		return onFlowIndex(toks) >= 0;
	}
	if (tokenWord(toks[0]) === 'on' && tokenWord(toks[1]) === 'error') {
		return toks.length === 3 && tokenWord(toks[2]) === 'goto';
	}
	return false;
}

function removablePartialToken(tokens: readonly VbaToken[], prefixLength: number): VbaToken | undefined {
	const last = tokens[tokens.length - 1];
	if (!last) {
		return undefined;
	}
	if (last.end !== prefixLength) {
		return undefined;
	}
	if (tokenName(last) || last.kind === 'integerLiteral') {
		return last;
	}
	return undefined;
}

function isOnErrorGotoDisablePartial(
	toks: readonly VbaToken[],
	partial: VbaToken,
): boolean {
	if (
		tokenWord(toks[0]) !== 'on' ||
		tokenWord(toks[1]) !== 'error' ||
		tokenWord(toks[2]) !== 'goto'
	) {
		return false;
	}
	if (toks.length === 4 && partial.kind === 'integerLiteral') {
		return normalizedDecimalLabel(partial.rawText) === '0';
	}
	return toks.length === 5 &&
		toks[3].rawText === '-' &&
		partial.kind === 'integerLiteral' &&
		normalizedDecimalLabel(partial.rawText) === '1';
}

function statementAtOffset(body: BodyNode[], offset: number): LeafStatementNode | undefined {
	for (const node of body) {
		if (isLeafStatement(node) && offset >= node.span.start && offset <= node.span.end) {
			return node;
		}
		if ('body' in node && Array.isArray(node.body)) {
			const nested = statementAtOffset(node.body, offset);
			if (nested) {
				return nested;
			}
		}
	}
	return undefined;
}

function procedureAtOffset(module: ModuleNode, offset: number): ProcedureNode | undefined {
	return module.members.find(
		(member): member is ProcedureNode =>
			member.kind === 'Procedure' &&
			offset >= member.span.start &&
			offset <= member.span.end,
	);
}

function statementLabelDeclaration(source: string, span: Span): VbaProcedureLabel | undefined {
	const toks = statementTokens(source, span);
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const label = labelFromToken(first, span);
	if (!label) {
		return undefined;
	}
	if (first.kind === 'integerLiteral' && toks.length > 1) {
		return label;
	}
	if (toks.length >= 2 && toks[1].rawText === ':') {
		return label;
	}
	if (toks.length === 1 && hasSourceColonAfterToken(source, span, first)) {
		return label;
	}
	return undefined;
}

function onStatementLabelReferences(
	toks: readonly VbaToken[],
	span: Span,
): VbaProcedureLabelReference[] {
	if (tokenWord(toks[1]) === 'error') {
		if (tokenWord(toks[2]) === 'resume' && tokenWord(toks[3]) === 'next') {
			return [];
		}
		if (tokenWord(toks[2]) !== 'goto') {
			return [];
		}
		const target = toks[3];
		if (!target || onErrorGotoDisableTarget(toks, 3)) {
			return [];
		}
		const ref = labelReferenceAfter(toks, span, 3, 'on-error-goto');
		return ref ? [ref] : [];
	}

	const flowIndex = onFlowIndex(toks);
	if (flowIndex < 0) {
		return [];
	}
	const statementKind = tokenWord(toks[flowIndex]) === 'gosub' ? 'on-gosub' : 'on-goto';
	const refs: VbaProcedureLabelReference[] = [];
	for (const group of splitTopLevelTokenGroups(toks, flowIndex + 1, ',')) {
		const ref = labelReferenceGroup(group, span, statementKind);
		if (ref) {
			refs.push(ref);
		}
	}
	return refs;
}

function onFlowIndex(toks: readonly VbaToken[]): number {
	return toks.findIndex((tok, i) =>
		i > 0 && (tokenWord(tok) === 'goto' || tokenWord(tok) === 'gosub')
	);
}

function onErrorGotoDisableTarget(toks: readonly VbaToken[], index: number): boolean {
	const target = toks[index];
	if (!target) {
		return false;
	}
	if (target.kind === 'integerLiteral' && normalizedDecimalLabel(target.rawText) === '0') {
		return true;
	}
	return target.rawText === '-' &&
		toks[index + 1]?.kind === 'integerLiteral' &&
		normalizedDecimalLabel(toks[index + 1].rawText) === '1';
}

function isOnErrorGotoDisableAt(toks: readonly VbaToken[], gotoIndex: number): boolean {
	return tokenWord(toks[gotoIndex - 2]) === 'on' &&
		tokenWord(toks[gotoIndex - 1]) === 'error' &&
		onErrorGotoDisableTarget(toks, gotoIndex + 1);
}

function labelReferenceAfter(
	toks: readonly VbaToken[],
	base: Span,
	index: number,
	statementKind: VbaProcedureLabelReference['statementKind'],
): VbaProcedureLabelReference | undefined {
	const group = toks.slice(index);
	const end = group.findIndex((tok) => tok.rawText === ',' || tokenWord(tok) === 'else');
	return labelReferenceGroup(end >= 0 ? group.slice(0, end) : group, base, statementKind);
}

function labelReferenceGroup(
	group: readonly VbaToken[],
	base: Span,
	statementKind: VbaProcedureLabelReference['statementKind'],
): VbaProcedureLabelReference | undefined {
	const content = group.filter((tok) => tok.kind !== 'comment');
	if (content.length !== 1) {
		return undefined;
	}
	const label = labelFromToken(content[0], base);
	return label ? { ...label, statementKind } : undefined;
}

function labelFromToken(tok: VbaToken, base: Span): VbaProcedureLabel | undefined {
	const name = tokenName(tok);
	if (name) {
		return {
			key: `name:${name.toLowerCase()}`,
			text: name,
			span: absoluteSpan(base, tok),
			kind: 'name',
		};
	}
	if (tok.kind === 'integerLiteral') {
		const normalized = normalizedDecimalLabel(tok.rawText);
		if (normalized !== undefined) {
			return {
				key: `line:${normalized}`,
				text: tok.rawText,
				span: absoluteSpan(base, tok),
				kind: 'line',
			};
		}
	}
	return undefined;
}

function normalizedDecimalLabel(raw: string): string | undefined {
	if (!/^\d+$/.test(raw)) {
		return undefined;
	}
	return raw.replace(/^0+/, '') || '0';
}

function hasSourceColonAfterToken(source: string, span: Span, tok: VbaToken): boolean {
	let i = span.start + tok.end;
	while (i < source.length && (source[i] === ' ' || source[i] === '\t')) {
		i++;
	}
	return source[i] === ':';
}

function forEachProcedureStatement(
	body: BodyNode[],
	visit: (stmt: LeafStatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (isLeafStatement(node)) {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachProcedureStatement(node.body, visit, activity);
		}
	}
}

function absoluteSpan(base: Span, token: VbaToken): Span {
	return { start: base.start + token.start, end: base.start + token.end };
}

function offsetInSpan(offset: number, span: Span): boolean {
	return offset >= span.start && offset <= span.end;
}

function physicalLineSpanAtOffset(source: string, offset: number): Span {
	const safe = Math.max(0, Math.min(offset, source.length));
	const before = source.lastIndexOf('\n', Math.max(0, safe - 1));
	const start = before < 0 ? 0 : before + 1;
	const after = source.indexOf('\n', safe);
	let end = after < 0 ? source.length : after;
	if (end > start && source[end - 1] === '\r') {
		end--;
	}
	return { start, end };
}

function compareLabels(a: VbaProcedureLabel, b: VbaProcedureLabel): number {
	if (a.kind !== b.kind) {
		return a.kind === 'name' ? -1 : 1;
	}
	return a.text.localeCompare(b.text, undefined, { numeric: true, sensitivity: 'base' });
}
