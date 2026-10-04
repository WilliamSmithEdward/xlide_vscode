import type { VbaToken } from '../lexer/tokenKinds';
import { isReservedIdentifier } from '../lexer/keywordTable';
import {
	absoluteSpan,
	splitTopLevelTokenGroups,
	startsPhysicalLine,
	statementTokens,
	statementTokensCached,
	tokenName,
	tokensWithoutLeadingLineNumber,
	tokenWord,
} from '../lexer/tokenHelpers';
import { forEachStatement } from '../parser/statementWalk';
import { parseModule } from '../parser/parseModule';
import type {
	BodyNode,
	LeafStatementNode,
	ProcedureNode,
	Span,
} from '../parser/nodes';
import { isLeafStatement, procedureAtOffset } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import { physicalLineSpanAtOffset } from '../../vbaSourceScan';

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

interface ProcedureLabelFacts {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	declarations?: readonly VbaProcedureLabel[];
	references?: readonly VbaProcedureLabelReference[];
}

const PROCEDURE_LABEL_FACTS = new WeakMap<ProcedureNode, ProcedureLabelFacts>();

function procedureLabelFacts(source: string, procedure: ProcedureNode, activity: ConditionalActivityTracker | undefined): ProcedureLabelFacts {
	const cached = PROCEDURE_LABEL_FACTS.get(procedure);
	if (cached && cached.source === source && cached.activity === activity) {
		cached.source = source;
		return cached;
	}
	const facts = { source, activity };
	PROCEDURE_LABEL_FACTS.set(procedure, facts);
	return facts;
}

/** Public collections and their label/span objects remain caller-owned. */
function copyLabels<T extends VbaProcedureLabel>(labels: readonly T[]): T[] {
	return labels.map(label => ({ ...label, span: { ...label.span } }));
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
	const facts = procedureLabelFacts(source, procedure, activity);
	if (!facts.declarations) {
		const labels: VbaProcedureLabel[] = [];
		forEachStatement(procedure.body, (stmt) => {
			labels.push(...statementLabelDeclarations(source, stmt.span));
		}, activity);
		facts.declarations = labels;
	}
	return copyLabels(facts.declarations);
}

export function collectProcedureLabelReferences(
	source: string,
	procedure: ProcedureNode,
	activity?: ConditionalActivityTracker,
): VbaProcedureLabelReference[] {
	const facts = procedureLabelFacts(source, procedure, activity);
	if (!facts.references) {
		const refs: VbaProcedureLabelReference[] = [];
		forEachStatement(procedure.body, (stmt) => {
			refs.push(...statementLabelReferences(source, stmt.span));
		}, activity);
		facts.references = refs;
	}
	return copyLabels(facts.references);
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
	return labelReferencesIn(tokensWithoutLeadingLineNumber(statementTokensCached(source, span)), span);
}

function labelReferencesIn(tokens: readonly VbaToken[], span: Span, firstWord = tokenWord(tokens[0])): VbaProcedureLabelReference[] {
	let toks = tokens;
	if (toks.length === 0) {
		return [];
	}
	if (firstWord === 'on') {
		// `On Local Error GoTo 0` is `On Error GoTo 0` (issue #98): drop the
		// Local so the target reads the same way.
		if (tokenWord(toks[1]) === 'local' && tokenWord(toks[2]) === 'error') {
			toks = [toks[0], ...toks.slice(2)];
		}
		return onStatementLabelReferences(toks, span);
	}
	const refs: VbaProcedureLabelReference[] = [];
	for (let i = 0; i < toks.length; i++) {
		const word = i === 0 ? firstWord : tokenWord(toks[i]);
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
	// Raw on purpose: localPrefix is a fresh derived string per call, so the
	// by-source cache would thrash instead of hitting.
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

/** The statement's first label, when it declares any: the line number of `10 L1:`. */
export function statementLabelDeclaration(source: string, span: Span): VbaProcedureLabel | undefined {
	return statementLabelDeclarations(source, span)[0];
}

/**
 * The statement's first label that a GoTo, GoSub, Resume or On ... GoTo
 * anywhere in the module names, or undefined. Control reaches any other
 * label only by falling into it, so a label nothing names, and a line
 * number written for Erl, carry what the statements before them knew
 * (issue #321, measured in Excel 16.0).
 */
export function jumpTargetLabelDeclaration(source: string, span: Span): VbaProcedureLabel | undefined {
	const labels = statementLabelDeclarations(source, span);
	if (labels.length === 0) {
		return undefined;
	}
	const targets = moduleLabelTargets(source);
	return labels.find((label) => targets.has(label.key));
}

/** The label keys each recent module's statements jump to, by source text. */
const MODULE_TARGETS = new Map<string, ReadonlySet<string>>();

function moduleLabelTargets(source: string): ReadonlySet<string> {
	let targets = MODULE_TARGETS.get(source);
	if (!targets) {
		const keys = new Set<string>();
		for (const member of parseModule(source).members) {
			if (member.kind === 'Procedure') {
				for (const ref of collectProcedureLabelReferences(source, member)) {
					keys.add(ref.key);
				}
			}
		}
		if (MODULE_TARGETS.size >= 8) {
			MODULE_TARGETS.clear();
		}
		MODULE_TARGETS.set(source, keys);
		targets = keys;
	}
	return targets;
}

/**
 * Every label the statement declares. A line can carry a line number and a
 * name both, `10 L1: x = 1`, and each is a target: GoTo 10 and Erl see the
 * number, GoTo L1 and Resume L1 the name (issue #230, measured in Excel
 * 16.0). A name is a label only at the start of its physical line, after
 * the line number if there is one: in `10: L1:` and `10 L1: L2:` the VBE
 * reads the second word as a call.
 */
export function statementLabelDeclarations(source: string, span: Span): VbaProcedureLabel[] {
	return labelDeclarationsIn(source, span, statementTokensCached(source, span));
}

function labelDeclarationsIn(source: string, span: Span, toks: readonly VbaToken[]): VbaProcedureLabel[] {
	const first = toks[0];
	if (!first) {
		return [];
	}
	if (first.kind !== 'integerLiteral' && toks.length >= 2 && toks[1].rawText !== ':') {
		return [];
	}
	const label = labelFromToken(first, span);
	if (!label) {
		return [];
	}
	if (first.kind === 'integerLiteral') {
		// A leading decimal integer is a line-label declaration whether or not a
		// statement follows it on the same line: a bare `100` on its own line is a
		// valid line label (e.g. an `On n GoTo 100` / `GoTo 100` target). Requiring
		// a trailing statement here previously left bare numeric labels uncollected,
		// so references to them falsely fired `undefined-label`. `labelFromToken`
		// already gated non-decimal forms (hex/octal) to undefined above.
		const named = toks.length === 2 && toks[1].kind !== 'integerLiteral' && hasSourceColonAfterToken(source, span, toks[1])
			? labelFromToken(toks[1], span)
			: undefined;
		return named ? [label, named] : [label];
	}
	if (!startsPhysicalLine(source, span.start)) {
		return [];
	}
	if (toks.length >= 2 && toks[1].rawText === ':') {
		return [label];
	}
	if (toks.length === 1 && hasSourceColonAfterToken(source, span, first)) {
		return [label];
	}
	return [];
}

/** One statement's existing label/error predicates, sharing one token lookup. */
export function statementHasUnstructuredFlow(source: string, span: Span): boolean {
	const tokens = statementTokensCached(source, span);
	const significant = tokensWithoutLeadingLineNumber(tokens);
	const first = tokenWord(significant[0]);
	return first === 'resume' || (first === 'on' && tokenWord(significant[1]) === 'error') ||
		labelDeclarationsIn(source, span, tokens).length > 0 ||
		labelReferencesIn(significant, span, first).length > 0;
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
	// A reserved word is never a label: `Else:` is the Else of its If with a
	// colon after it (issue #129, measured in Excel 16.0: two of them in one
	// procedure compile), the way `Next:` and `End If:` already read.
	if (name && isReservedIdentifier(name)) {
		return undefined;
	}
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


function offsetInSpan(offset: number, span: Span): boolean {
	return offset >= span.start && offset <= span.end;
}

function compareLabels(a: VbaProcedureLabel, b: VbaProcedureLabel): number {
	if (a.kind !== b.kind) {
		return a.kind === 'name' ? -1 : 1;
	}
	return a.text.localeCompare(b.text, undefined, { numeric: true, sensitivity: 'base' });
}
