import type { Span } from '../parser/nodes';
import { tokenize } from '../lexer/tokenize';

export interface VbaDiagnosticCodeActionInput {
	code: string;
	message?: string;
	span: Span;
}

export interface VbaTextEdit {
	span: Span;
	newText: string;
}

export interface VbaDiagnosticCodeAction {
	title: string;
	kind: 'quickfix';
	isPreferred?: boolean;
	edits: readonly VbaTextEdit[];
}

export function normalizeDiagnosticCode(code: unknown): string | undefined {
	if (typeof code === 'string') {
		return code;
	}
	if (typeof code === 'number') {
		return String(code);
	}
	if (code && typeof code === 'object' && 'value' in code) {
		const value = (code as { value?: unknown }).value;
		if (typeof value === 'string') {
			return value;
		}
		if (typeof value === 'number') {
			return String(value);
		}
	}
	return undefined;
}

export function resolveDiagnosticCodeActions(
	source: string,
	diagnostic: VbaDiagnosticCodeActionInput,
): VbaDiagnosticCodeAction[] {
	switch (diagnostic.code) {
		case 'call-requires-parens':
			return callRequiresParensActions(source, diagnostic.span);
		case 'call-statement-forbids-parens':
			return callStatementForbidsParensActions(source, diagnostic.span);
		case 'expression-call-requires-parens':
			return expressionCallRequiresParensActions(source, diagnostic.span);
		case 'invalid-explicit-call-target':
			return invalidExplicitCallTargetActions(source, diagnostic.span);
		case 'option-explicit-missing':
			return optionExplicitMissingActions(source);
		case 'set-required':
			return setRequiredActions(source, diagnostic.span);
		case 'set-requires-object':
			return setRequiresObjectActions(source, diagnostic.span);
		default:
			return [];
	}
}

function callRequiresParensActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const toks = tokenize(source.slice(line.start, line.end));
	if (toks.length === 0 || toks[0].rawText.toLowerCase() !== 'call') {
		return [];
	}
	const argIdx = toks.findIndex((tok) => line.start + tok.start === span.start);
	if (argIdx <= 1) {
		return [];
	}
	const calleeEnd = line.start + toks[argIdx - 1].end;
	const argStart = line.start + toks[argIdx].start;
	if (!/^[ \t]+$/.test(source.slice(calleeEnd, argStart))) {
		return [];
	}
	const argEnd = callArgumentListEnd(source, line, toks, argIdx);
	if (argEnd === undefined || argEnd <= argStart) {
		return [];
	}
	return [{
		title: 'Add parentheses to Call argument list',
		kind: 'quickfix',
		isPreferred: true,
		edits: [
			{ span: { start: calleeEnd, end: argStart }, newText: '(' },
			{ span: { start: argEnd, end: argEnd }, newText: ')' },
		],
	}];
}

function expressionCallRequiresParensActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const toks = tokenize(source.slice(line.start, line.end));
	const calleeIdx = tokenIndexAtAbsoluteStart(toks, line, span.start);
	if (calleeIdx < 1 || toks[calleeIdx - 1].rawText !== '=') {
		return [];
	}
	const arg = toks[calleeIdx + 1];
	if (!arg || arg.kind === 'comment') {
		return [];
	}
	const calleeEnd = line.start + toks[calleeIdx].end;
	const argStart = line.start + arg.start;
	if (!/^[ \t]+$/.test(source.slice(calleeEnd, argStart))) {
		return [];
	}
	const argEnd = callArgumentListEnd(source, line, toks, calleeIdx + 1);
	if (argEnd === undefined || argEnd <= argStart) {
		return [];
	}
	return [{
		title: 'Add parentheses to function call arguments',
		kind: 'quickfix',
		isPreferred: true,
		edits: [
			{ span: { start: calleeEnd, end: argStart }, newText: '(' },
			{ span: { start: argEnd, end: argEnd }, newText: ')' },
		],
	}];
}

function callStatementForbidsParensActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const parens = trailingEmptyParens(source, span);
	if (!parens) {
		return [];
	}
	return [{
		title: 'Remove empty parentheses',
		kind: 'quickfix',
		isPreferred: true,
		edits: [{ span: parens, newText: '' }],
	}];
}

function invalidExplicitCallTargetActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const before = source.slice(line.start, span.start);
	const call = /\bCall[ \t]*$/i.exec(before);
	if (!call) {
		return [];
	}

	const after = source.slice(span.end, line.end);
	const emptyParens = /^[ \t]*\([ \t]*\)/.exec(after);
	const afterCall = emptyParens ? after.slice(emptyParens[0].length) : after;
	if (!isBlankOrCommentTail(afterCall)) {
		return [];
	}

	const edits: VbaTextEdit[] = [{
		span: { start: line.start + call.index, end: span.start },
		newText: '',
	}];
	if (emptyParens) {
		edits.push({
			span: { start: span.end, end: span.end + emptyParens[0].length },
			newText: '',
		});
	}

	return [{
		title: 'Use bare runtime call syntax',
		kind: 'quickfix',
		isPreferred: true,
		edits,
	}];
}

function setRequiredActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const statementStart = firstNonWhitespaceOffset(source, line);
	if (statementStart === undefined || source.slice(statementStart, span.start).includes(':')) {
		return [];
	}
	const letPrefix = /^Let[ \t]+/i.exec(source.slice(statementStart, span.start));
	const edit = letPrefix
		? {
			span: { start: statementStart, end: statementStart + letPrefix[0].length },
			newText: 'Set ',
		}
		: {
			span: { start: statementStart, end: statementStart },
			newText: 'Set ',
		};
	return [{
		title: letPrefix ? 'Replace Let with Set' : 'Add Set to object assignment',
		kind: 'quickfix',
		isPreferred: true,
		edits: [edit],
	}];
}

function setRequiresObjectActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const statementStart = firstNonWhitespaceOffset(source, line);
	if (statementStart === undefined || source.slice(statementStart, span.start).includes(':')) {
		return [];
	}
	const setPrefix = /^Set[ \t]+/i.exec(source.slice(statementStart, span.start));
	if (!setPrefix) {
		return [];
	}
	return [{
		title: 'Remove Set from scalar assignment',
		kind: 'quickfix',
		isPreferred: true,
		edits: [{
			span: { start: statementStart, end: statementStart + setPrefix[0].length },
			newText: '',
		}],
	}];
}

function optionExplicitMissingActions(source: string): VbaDiagnosticCodeAction[] {
	if (/\bOption\s+Explicit\b/i.test(source)) {
		return [];
	}
	const insertAt = optionExplicitInsertOffset(source);
	const eol = detectEol(source);
	return [{
		title: 'Add Option Explicit',
		kind: 'quickfix',
		isPreferred: true,
		edits: [{
			span: { start: insertAt, end: insertAt },
			newText: `Option Explicit${eol}`,
		}],
	}];
}

function trailingEmptyParens(source: string, span: Span): Span | undefined {
	const slice = source.slice(span.start, span.end);
	const hit = /\([ \t]*\)[ \t]*$/.exec(slice);
	if (!hit) {
		return undefined;
	}
	return {
		start: span.start + hit.index,
		end: span.start + hit.index + hit[0].length,
	};
}

function callArgumentListEnd(
	source: string,
	line: Span,
	toks: ReturnType<typeof tokenize>,
	argIdx: number,
): number | undefined {
	let end = line.end;
	for (let i = argIdx; i < toks.length; i++) {
		if (toks[i].kind === 'comment') {
			end = line.start + toks[i].start;
			break;
		}
		if (toks[i].rawText === ':') {
			return undefined;
		}
	}
	while (end > line.start && (source[end - 1] === ' ' || source[end - 1] === '\t')) {
		end--;
	}
	return end;
}

function tokenIndexAtAbsoluteStart(
	toks: ReturnType<typeof tokenize>,
	line: Span,
	start: number,
): number {
	return toks.findIndex((tok) => line.start + tok.start === start);
}

function optionExplicitInsertOffset(source: string): number {
	let offset = source.charCodeAt(0) === 0xfeff ? 1 : 0;
	while (offset < source.length) {
		const line = readPhysicalLine(source, offset);
		if (!/^\s*Attribute\s+\w+\s*=/i.test(line.text)) {
			break;
		}
		offset = line.next;
	}
	return offset;
}

function physicalLineSpan(source: string, offset: number): Span {
	let start = Math.max(0, Math.min(offset, source.length));
	while (start > 0 && source[start - 1] !== '\n' && source[start - 1] !== '\r') {
		start--;
	}
	let end = Math.max(0, Math.min(offset, source.length));
	while (end < source.length && source[end] !== '\n' && source[end] !== '\r') {
		end++;
	}
	return { start, end };
}

function firstNonWhitespaceOffset(source: string, line: Span): number | undefined {
	let offset = line.start;
	while (offset < line.end && (source[offset] === ' ' || source[offset] === '\t')) {
		offset++;
	}
	return offset < line.end ? offset : undefined;
}

function readPhysicalLine(
	source: string,
	offset: number,
): { text: string; next: number } {
	let end = offset;
	while (end < source.length && source[end] !== '\n' && source[end] !== '\r') {
		end++;
	}
	let next = end;
	if (source[next] === '\r') {
		next += source[next + 1] === '\n' ? 2 : 1;
	} else if (source[next] === '\n') {
		next++;
	}
	return {
		text: source.slice(offset, end),
		next,
	};
}

function isBlankOrCommentTail(text: string): boolean {
	const trimmed = text.trimStart();
	return trimmed.length === 0 || trimmed.startsWith("'");
}

function detectEol(source: string): string {
	return source.includes('\r\n') ? '\r\n' : '\n';
}
