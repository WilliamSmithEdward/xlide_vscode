import type { Span } from '../parser/nodes';
import { tokenize } from '../lexer/tokenize';
import {
	explicitCallStatementArgumentListWithoutParens,
	explicitCallStatementBareRuntimeRewrite,
	standaloneEmptyParenthesizedCallStatement,
} from '../call/callContext';
import { detectEol, leadingWhitespace } from '../../vbaStructuralAnalysis';
import { ANALYSIS_SUPPRESSION_DIRECTIVE_CODE } from '../diagnostics/analysisSuppressions';
import type { VbaDiagnosticData } from '../diagnostics/analyzeModule';

export interface VbaDiagnosticCodeActionInput {
	code: string;
	message?: string;
	span: Span;
	expectedClose?: string;
	insertLine?: number;
	expectedCloseReplacementSpan?: Span;
	expectedCloseReplacementText?: string;
	includeSuppressionAction?: boolean;
	data?: VbaDiagnosticData;
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
	const actions = ruleSpecificDiagnosticCodeActions(source, diagnostic);
	if (diagnostic.includeSuppressionAction) {
		const suppression = suppressNextLineAction(source, diagnostic);
		if (suppression) {
			actions.push(suppression);
		}
	}
	return actions;
}

function ruleSpecificDiagnosticCodeActions(
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
		case 'missing-block-closer':
			return missingBlockCloserActions(source, diagnostic);
		case 'dim-initializer':
			return dimInitializerActions(source, diagnostic.span);
		case 'option-after-declaration':
			return optionAfterDeclarationActions(source, diagnostic.span);
		case 'option-explicit-missing':
			return optionExplicitMissingActions(source);
		case 'set-required':
			return setRequiredActions(source, diagnostic.span);
		case 'set-requires-object':
			return setRequiresObjectActions(source, diagnostic.span);
		case 'argument-count':
			return missingRequiredArgumentPlaceholderActions(diagnostic);
		case 'unknown-call':
			return createProcedureStubActions(diagnostic);
		default:
			return [];
	}
}

function missingRequiredArgumentPlaceholderActions(
	diagnostic: VbaDiagnosticCodeActionInput,
): VbaDiagnosticCodeAction[] {
	const data = diagnostic.data?.missingRequiredArgumentPlaceholder;
	if (!data) {
		return [];
	}
	return [{
		title: `Insert placeholder for missing argument '${data.parameterName}'`,
		kind: 'quickfix',
		isPreferred: false,
		edits: [data.edit],
	}];
}

function createProcedureStubActions(
	diagnostic: VbaDiagnosticCodeActionInput,
): VbaDiagnosticCodeAction[] {
	const data = diagnostic.data?.createProcedureStub;
	if (!data) {
		return [];
	}
	return [{
		title: `Create Private Sub '${data.procedureName}' in this module`,
		kind: 'quickfix',
		isPreferred: false,
		edits: [data.edit],
	}];
}

function suppressNextLineAction(
	source: string,
	diagnostic: VbaDiagnosticCodeActionInput,
): VbaDiagnosticCodeAction | undefined {
	if (!diagnostic.code || diagnostic.code === ANALYSIS_SUPPRESSION_DIRECTIVE_CODE) {
		return undefined;
	}
	const line = physicalLineSpan(source, diagnostic.span.start);
	const indent = leadingWhitespace(source.slice(line.start, line.end));
	const eol = detectEol(source);
	return {
		title: `Suppress '${diagnostic.code}' on next line`,
		kind: 'quickfix',
		isPreferred: false,
		edits: [{
			span: { start: line.start, end: line.start },
			newText: `${indent}' @xlide-analysis-disable-next-line ${diagnostic.code}${eol}`,
		}],
	};
}

function callRequiresParensActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const args = explicitCallStatementArgumentListWithoutParens(source, line);
	if (!args || args.firstArgumentSpan.start !== span.start) {
		return [];
	}
	return [{
		title: 'Add parentheses to Call argument list',
		kind: 'quickfix',
		isPreferred: true,
		edits: [
			{
				span: { start: args.calleeEndOffset, end: args.argumentSpan.start },
				newText: '(',
			},
			{
				span: { start: args.argumentSpan.end, end: args.argumentSpan.end },
				newText: ')',
			},
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
	const call = standaloneEmptyParenthesizedCallStatement(source, span);
	if (!call) {
		return [];
	}
	return [{
		title: 'Remove empty parentheses',
		kind: 'quickfix',
		isPreferred: true,
		edits: [{ span: call.emptyParensSpan, newText: '' }],
	}];
}

function invalidExplicitCallTargetActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const rewrite = explicitCallStatementBareRuntimeRewrite(source, line);
	if (!rewrite || rewrite.targetSpan.start !== span.start || rewrite.targetSpan.end !== span.end) {
		return [];
	}

	const edits: VbaTextEdit[] = [{
		span: rewrite.callPrefixSpan,
		newText: '',
	}];
	if (rewrite.emptyParensSpan) {
		edits.push({
			span: rewrite.emptyParensSpan,
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

function missingBlockCloserActions(
	source: string,
	diagnostic: VbaDiagnosticCodeActionInput,
): VbaDiagnosticCodeAction[] {
	const expectedClose = diagnostic.expectedClose;
	if (!expectedClose || !SAFE_BLOCK_CLOSERS.has(expectedClose)) {
		return [];
	}
	if (
		diagnostic.expectedCloseReplacementSpan &&
		diagnostic.expectedCloseReplacementText === expectedClose &&
		SAFE_BLOCK_CLOSERS.has(diagnostic.expectedCloseReplacementText)
	) {
		const current = source.slice(
			diagnostic.expectedCloseReplacementSpan.start,
			diagnostic.expectedCloseReplacementSpan.end,
		);
		return [{
			title: `Replace '${current}' with '${expectedClose}'`,
			kind: 'quickfix',
			isPreferred: true,
			edits: [{
				span: diagnostic.expectedCloseReplacementSpan,
				newText: expectedClose,
			}],
		}];
	}

	const openerLine = physicalLineSpan(source, diagnostic.span.start);
	const indent = leadingWhitespace(source.slice(openerLine.start, openerLine.end));
	const eol = detectEol(source);
	const insert = missingCloserInsert(source, diagnostic.insertLine, `${indent}${expectedClose}`, eol);
	if (!insert) {
		return [];
	}

	return [{
		title: `Insert '${expectedClose}'`,
		kind: 'quickfix',
		isPreferred: true,
		edits: [insert],
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

function optionAfterDeclarationActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	const physicalLine = readPhysicalLine(source, line.start);
	const optionText = physicalLine.text.trim();
	if (!/^Option\b/i.test(optionText)) {
		return [];
	}

	const insertAt = optionExplicitInsertOffset(source);
	if (line.start <= insertAt || insertAt >= physicalLine.next) {
		return [];
	}

	const eol = detectEol(source);
	return [{
		title: 'Move Option statement before declarations',
		kind: 'quickfix',
		isPreferred: true,
		edits: [
			{
				span: { start: line.start, end: physicalLine.next },
				newText: '',
			},
			{
				span: { start: insertAt, end: insertAt },
				newText: `${optionText}${eol}`,
			},
		],
	}];
}

function dimInitializerActions(
	source: string,
	span: Span,
): VbaDiagnosticCodeAction[] {
	const line = physicalLineSpan(source, span.start);
	if (!isInsideProcedureBefore(source, line.start)) {
		return [];
	}

	const lineText = source.slice(line.start, line.end);
	const toks = tokenize(lineText);
	if (toks.length < 4 || toks[0].rawText.toLowerCase() !== 'dim') {
		return [];
	}
	if (toks.some((tok) => tok.kind === 'comment')) {
		return [];
	}

	const eqIdx = toks.findIndex((tok) => line.start + tok.start === span.start && tok.rawText === '=');
	if (eqIdx < 2) {
		return [];
	}
	if (hasUnsafeDeclarationDelimiter(toks, eqIdx)) {
		return [];
	}

	const name = declarationAssignmentTarget(toks[1]);
	const rhs = source.slice(line.start + toks[eqIdx].end, line.end).trim();
	if (!name || !rhs) {
		return [];
	}

	let replaceStart = span.start;
	while (replaceStart > line.start && (source[replaceStart - 1] === ' ' || source[replaceStart - 1] === '\t')) {
		replaceStart--;
	}

	const eol = detectEol(source);
	const indent = leadingWhitespace(lineText);
	return [{
		title: 'Split declaration initializer',
		kind: 'quickfix',
		isPreferred: true,
		edits: [{
			span: { start: replaceStart, end: line.end },
			newText: `${eol}${indent}${name} = ${rhs}`,
		}],
	}];
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

const SAFE_BLOCK_CLOSERS = new Set([
	'End Sub',
	'End Function',
	'End Property',
	'End If',
	'End With',
	'End Select',
	'End Type',
	'End Enum',
	'Next',
	'Loop',
	'Wend',
	'#End If',
]);

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

function missingCloserInsert(
	source: string,
	insertLine: number | undefined,
	closeLine: string,
	eol: string,
): VbaTextEdit | undefined {
	const lineStart = lineStartOffset(source, insertLine);
	if (lineStart === undefined || lineStart === source.length) {
		const prefix = source.length > 0 && !source.endsWith('\n') && !source.endsWith('\r') ? eol : '';
		return {
			span: { start: source.length, end: source.length },
			newText: `${prefix}${closeLine}${eol}`,
		};
	}
	return {
		span: { start: lineStart, end: lineStart },
		newText: `${closeLine}${eol}`,
	};
}

function lineStartOffset(source: string, line: number | undefined): number | undefined {
	if (line === undefined || line < 0) {
		return undefined;
	}
	let offset = 0;
	for (let current = 0; current < line; current++) {
		if (offset >= source.length) {
			return undefined;
		}
		const next = readPhysicalLine(source, offset).next;
		if (next <= offset) {
			return undefined;
		}
		offset = next;
	}
	return offset <= source.length ? offset : undefined;
}

function declarationAssignmentTarget(tok: ReturnType<typeof tokenize>[number] | undefined): string | undefined {
	if (!tok || (tok.kind !== 'identifier' && tok.kind !== 'bracketedIdentifier')) {
		return undefined;
	}
	return tok.rawText;
}

function hasUnsafeDeclarationDelimiter(toks: ReturnType<typeof tokenize>, eqIdx: number): boolean {
	let parenDepth = 0;
	for (let i = 0; i < toks.length; i++) {
		const text = toks[i].rawText;
		if (text === '(') {
			parenDepth++;
		} else if (text === ')') {
			parenDepth = Math.max(0, parenDepth - 1);
		} else if (text === ':' || (text === ',' && (i < eqIdx || parenDepth === 0))) {
			return true;
		}
	}
	return false;
}

function isInsideProcedureBefore(source: string, offset: number): boolean {
	let inside = false;
	let current = source.charCodeAt(0) === 0xfeff ? 1 : 0;
	while (current < offset) {
		const line = readPhysicalLine(source, current);
		const text = line.text.trim();
		if (/^(?:Public|Private|Friend|Static)?\s*(?:Sub|Function|Property\s+(?:Get|Let|Set))\b/i.test(text)) {
			inside = true;
		} else if (/^End\s+(?:Sub|Function|Property)\b/i.test(text)) {
			inside = false;
		}
		if (line.next <= current) {
			break;
		}
		current = line.next;
	}
	return inside;
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

