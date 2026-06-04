// Keyword and block-snippet completion resolver.
//
// Provides canonical VBA keyword/snippet completions in statement contexts
// without replacing symbol completion. Narrow grammar contexts such as
// `Option `, `End `, and `On Error ` are exclusive because arbitrary identifiers
// are invalid there; broad statement starts are additive.

import { tokenize } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import {
	openSmartBlockClosersBefore,
	VBA_BLOCK_INDENT_UNIT,
	vbaSmartBlockSnippetsFor,
	type VbaSmartBlockLayout,
	type VbaSmartBlockSnippetSpec,
} from '../../vbaStructuralAnalysis';

export type KeywordCompletionKind = 'keyword' | 'snippet';

export interface KeywordCompletion {
	label: string;
	kind: KeywordCompletionKind;
	detail: string;
	insertText: string;
	filterText?: string;
	documentation?: string;
	sortText?: string;
}

export interface KeywordCompletionResult {
	items: KeywordCompletion[];
	/** True when the grammar position should not show ordinary symbols. */
	exclusive: boolean;
}

export interface KeywordCompletionOptions {
	blockLayout?: VbaSmartBlockLayout;
}

interface CompletionContext {
	partial: string;
	prefix: VbaToken[];
	atStatementStart: boolean;
	statementStart: number;
}

interface KeywordSpec {
	label: string;
	insertText?: string;
	filterText?: string;
	matchText?: readonly string[];
	detail: string;
	documentation?: string;
	sortText?: string;
	kind?: KeywordCompletionKind;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const I = VBA_BLOCK_INDENT_UNIT;

const STATEMENT_SNIPPETS: readonly KeywordSpec[] = [
	snippet('Option Explicit', 'Option Explicit', 'Option statement'),
	snippet('Dim', 'Dim ${1:name} As ${2:Variant}', 'Variable declaration'),
	snippet('Set', 'Set ${1:object} = ${2:value}', 'Object assignment statement'),
	snippet('ReDim', 'ReDim ${1:array}(${2:upperBound})', 'Dynamic array allocation'),
	snippet(
		'ReDim Preserve',
		'ReDim Preserve ${1:array}(${2:upperBound})',
		'Resize a dynamic array while preserving values',
		undefined,
		['redimpreserve'],
	),
	snippet('Exit Sub', 'Exit Sub', 'Procedure exit statement'),
	snippet('Exit Function', 'Exit Function', 'Function exit statement'),
	snippet('Exit Property', 'Exit Property', 'Property exit statement'),
	snippet('Exit For', 'Exit For', 'For loop exit statement'),
	snippet('Exit Do', 'Exit Do', 'Do loop exit statement'),
	snippet('On Error Resume Next', 'On Error Resume Next', 'Error-handling statement'),
	snippet('On Error GoTo 0', 'On Error GoTo 0', 'Error-handling statement'),
	snippet(
		'On Error GoTo Handler',
		blockText(
			'On Error GoTo ${1:ErrHandler}',
			I + '$0',
			I + 'Exit ${2|Sub,Function,Property|}',
			'${1:ErrHandler}:',
			I + 'MsgBox "Error " & Err.Number & ": " & Err.Description',
		),
		'Error handler block',
		undefined,
		['onerror'],
	),
	snippet('Debug.Print', 'Debug.Print ${1:value}$0', 'Debug.Print statement', undefined, ['dp']),
];

const MODIFIER_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Const', 'Constant declaration'),
	keyword('Dim', 'Variable declaration'),
	snippet(
		'Declare PtrSafe Sub',
		'Declare PtrSafe Sub ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr})',
		'64-bit-safe external Sub declaration',
		undefined,
		['declareptrsafesub', 'ptrsafesub'],
	),
	snippet(
		'Declare PtrSafe Function',
		'Declare PtrSafe Function ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr}) As ${5:LongPtr}',
		'64-bit-safe external Function declaration',
		undefined,
		['declareptrsafefunction', 'ptrsafefunction'],
	),
	snippet(
		'Declare Sub',
		'Declare Sub ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:Long})',
		'External Sub declaration',
		undefined,
		['declaresub'],
	),
	snippet(
		'Declare Function',
		'Declare Function ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:Long}) As ${5:Long}',
		'External Function declaration',
		undefined,
		['declarefunction'],
	),
];

const OPTION_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Explicit', 'Option Explicit'),
	snippet('Base 0', 'Base 0', 'Option Base 0'),
	snippet('Base 1', 'Base 1', 'Option Base 1'),
	snippet('Compare Binary', 'Compare Binary', 'Option Compare Binary'),
	snippet('Compare Text', 'Compare Text', 'Option Compare Text'),
];

const OPTION_COMPARE_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Binary', 'Option Compare Binary'),
	keyword('Text', 'Option Compare Text'),
];

const ON_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Error', 'On Error statement'),
];

const ON_ERROR_SNIPPETS: readonly KeywordSpec[] = [
	snippet('GoTo 0', 'GoTo 0', 'Disable active error handler'),
	snippet('GoTo -1', 'GoTo -1', 'Clear current error state'),
	snippet('GoTo label', 'GoTo ${1:label}', 'Branch to an error handler label'),
	snippet('Resume Next', 'Resume Next', 'Continue after the statement that raised an error'),
];

const EXIT_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Sub', 'Exit Sub'),
	keyword('Function', 'Exit Function'),
	keyword('Property', 'Exit Property'),
	keyword('For', 'Exit For'),
	keyword('Do', 'Exit Do'),
];

const DO_SNIPPETS: readonly KeywordSpec[] = [
	snippet('While', 'While ${1:condition}', 'Do While condition'),
	snippet('Until', 'Until ${1:condition}', 'Do Until condition'),
];

const LOOP_SNIPPETS: readonly KeywordSpec[] = [
	snippet('While', 'While ${1:condition}', 'Loop While condition'),
	snippet('Until', 'Until ${1:condition}', 'Loop Until condition'),
];

const SELECT_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Case', 'Select Case'),
];

const CASE_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Else', 'Case Else'),
	snippet('Is', 'Is ${1:operator} ${2:value}', 'Case Is comparison'),
];

const FOR_EACH_ITEM_SNIPPETS: readonly KeywordSpec[] = [
	keyword('In', 'For Each item In collection'),
];

const DIRECTIVE_SNIPPETS: readonly KeywordSpec[] = [
	snippet('#Const', '#Const ${1:name} = ${2:value}', 'Conditional compilation constant'),
	snippet(
		'#If VBA7 Then',
		blockText('#If VBA7 Then', I + '$0', '#End If'),
		'VBA7 conditional compilation block',
		undefined,
		['vba7', 'ifvba7'],
	),
	snippet(
		'#If Win64 Then',
		blockText('#If Win64 Then', I + '$0', '#End If'),
		'64-bit Office conditional compilation block',
		undefined,
		['win64', 'ifwin64'],
	),
	// Microsoft PtrSafe guidance recommends #If VBA7 for compatibility with VBA6
	// and PtrSafe + LongPtr for pointer/handle-sized values in VBA7.
	snippet(
		'#If VBA7 Declare Sub',
		blockText(
			'#If VBA7 Then',
			'Public Declare PtrSafe Sub ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr})',
			'#Else',
			'Public Declare Sub ${1/(.*)/$1/} Lib "${2/(.*)/$1/}" (ByVal ${3/(.*)/$1/} As ${5:Long})',
			'#End If',
		),
		'VBA7/VBA6-compatible external Sub declaration',
		undefined,
		['vba7declare', 'vba7declaresub', 'ptrsafedeclare'],
	),
	snippet(
		'#If VBA7 Declare Function',
		blockText(
			'#If VBA7 Then',
			'Public Declare PtrSafe Function ${1:ProcedureName} Lib "${2:library}" (ByVal ${3:argument} As ${4:LongPtr}) As ${5:LongPtr}',
			'#Else',
			'Public Declare Function ${1/(.*)/$1/} Lib "${2/(.*)/$1/}" (ByVal ${3/(.*)/$1/} As ${6:Long}) As ${7:Long}',
			'#End If',
		),
		'VBA7/VBA6-compatible external Function declaration',
		undefined,
		['vba7declarefunction', 'ptrsafefunctiondeclare'],
	),
	keyword('#ElseIf', 'Conditional compilation branch'),
	keyword('#Else', 'Conditional compilation branch'),
	keyword('#End If', 'End conditional compilation block'),
];

export function resolveKeywordCompletions(
	source: string,
	offset: number,
	options: KeywordCompletionOptions = {},
): KeywordCompletionResult {
	const ctx = completionContext(source, offset);
	if (!ctx) {
		return { items: [], exclusive: false };
	}
	const blockLayout = options.blockLayout;

	const first = word(ctx.prefix[0]);
	const second = word(ctx.prefix[1]);
	if (first === 'option') {
		if (second === 'compare') {
			return complete(OPTION_COMPARE_SNIPPETS, ctx.partial, true);
		}
		return complete(OPTION_SNIPPETS, ctx.partial, true);
	}
	if (first === 'end') {
		return complete(endCompletions(source, ctx.statementStart), ctx.partial, true);
	}
	if (first === 'on') {
		if (second === 'error') {
			return complete(ON_ERROR_SNIPPETS, ctx.partial, true);
		}
		return complete(ON_SNIPPETS, ctx.partial, true);
	}
	if (first === 'exit' && ctx.prefix.length === 1) {
		return complete(EXIT_SNIPPETS, ctx.partial, true);
	}
	if (first === 'do' && ctx.prefix.length === 1) {
		return complete(DO_SNIPPETS, ctx.partial, true);
	}
	if (first === 'loop' && ctx.prefix.length === 1) {
		return complete(LOOP_SNIPPETS, ctx.partial, true);
	}
	if (first === 'select' && ctx.prefix.length === 1) {
		return complete(SELECT_SNIPPETS, ctx.partial, true);
	}
	if (first === 'case' && ctx.prefix.length === 1) {
		return complete(CASE_SNIPPETS, ctx.partial, false);
	}
	if (first === 'for' && second === 'each' && ctx.prefix.length === 3) {
		return complete(FOR_EACH_ITEM_SNIPPETS, ctx.partial, true);
	}
	if (first === '#') {
		return complete(directiveSnippets(blockLayout), ctx.partial, true);
	}
	if (ctx.prefix.length === 1 && isAccessModifier(first)) {
		return complete(modifierSnippets(blockLayout), ctx.partial, true);
	}
	if (ctx.atStatementStart) {
		const close = closingCompletion(source, ctx.statementStart);
		const snippets = statementSnippets(blockLayout);
		return complete(close ? [close, ...snippets] : snippets, ctx.partial, false);
	}
	return { items: [], exclusive: false };
}

export function materializeKeywordSnippet(
	insertText: string,
	baseIndent: string,
): string {
	if (!baseIndent) {
		return insertText;
	}
	return insertText.replace(/\n(?!\n|$)/g, `\n${baseIndent}`);
}

function completionContext(source: string, offset: number): CompletionContext | undefined {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	const tokens = tokenize(source.slice(0, safeOffset));
	const last = tokens[tokens.length - 1];
	if (last?.kind === 'comment' || last?.kind === 'stringLiteral') {
		return undefined;
	}

	const statementStart = currentStatementStart(tokens, safeOffset);
	const statementTokens = tokens
		.filter((token) => token.start >= statementStart && token.kind !== 'comment' && token.kind !== 'newline' && token.kind !== 'colon');
	let partial = '';
	if (statementTokens.length > 0) {
		const maybePartial = statementTokens[statementTokens.length - 1];
		if (isIdentLike(maybePartial) && maybePartial.end === safeOffset) {
			partial = maybePartial.rawText;
			statementTokens.pop();
		}
	}

	const prefix = statementTokens;
	const atStatementStart = prefix.length === 0;
	if (!atStatementStart && !isKeywordPrefix(prefix)) {
		return undefined;
	}
	return { partial, prefix, atStatementStart, statementStart };
}

function currentStatementStart(tokens: readonly VbaToken[], fallback: number): number {
	for (let i = tokens.length - 1; i >= 0; i -= 1) {
		if (tokens[i].kind === 'newline' || tokens[i].kind === 'colon') {
			return tokens[i].end;
		}
	}
	return 0;
}

function isKeywordPrefix(tokens: readonly VbaToken[]): boolean {
	return tokens.every((token, idx) => {
		if (idx === 0 && token.rawText === '#') {
			return true;
		}
		return isIdentLike(token);
	});
}

function statementSnippets(layout: VbaSmartBlockLayout | undefined): readonly KeywordSpec[] {
	return [
		...vbaSmartBlockSnippetsFor('statement', layout).map(keywordSpecFromSmartBlock),
		...STATEMENT_SNIPPETS,
	];
}

function modifierSnippets(layout: VbaSmartBlockLayout | undefined): readonly KeywordSpec[] {
	return [
		...vbaSmartBlockSnippetsFor('modifier', layout).map(keywordSpecFromSmartBlock),
		...MODIFIER_SNIPPETS,
	];
}

function directiveSnippets(layout: VbaSmartBlockLayout | undefined): readonly KeywordSpec[] {
	return [
		...vbaSmartBlockSnippetsFor('directive', layout).map(keywordSpecFromSmartBlock),
		...DIRECTIVE_SNIPPETS,
	];
}

function complete(
	specs: readonly KeywordSpec[],
	partial: string,
	exclusive: boolean,
): KeywordCompletionResult {
	const lower = partial.toLowerCase();
	const explicitAliasMatches = specs.filter((spec) => exactAliasMatch(spec, lower));
	const candidates = explicitAliasMatches.length > 0
		? explicitAliasMatches
		: specs.filter((spec) => matchesPartial(spec, lower));
	const items = candidates
		.map((spec, index) => ({
			label: spec.label,
			kind: spec.kind ?? 'keyword',
			detail: spec.detail,
			insertText: spec.insertText ?? spec.label,
			filterText: spec.filterText ?? spec.label,
			documentation: spec.documentation,
			sortText: spec.sortText ?? `${index.toString().padStart(3, '0')}:${spec.label}`,
		}));
	return { items, exclusive };
}

function exactAliasMatch(spec: KeywordSpec, lowerPartial: string): boolean {
	return (spec.matchText ?? []).some((alias) => alias.toLowerCase() === lowerPartial);
}

function matchesPartial(spec: KeywordSpec, lowerPartial: string): boolean {
	if (!lowerPartial) {
		return true;
	}
	const compactPartial = compactKeywordText(lowerPartial);
	const candidates = [
		spec.label,
		spec.filterText,
		spec.insertText,
		...(spec.matchText ?? []),
	].filter((candidate): candidate is string => Boolean(candidate));
	return candidates.some((candidate) => {
		const lower = candidate.toLowerCase();
		return lower.startsWith(lowerPartial) || compactKeywordText(lower).startsWith(compactPartial);
	});
}

function compactKeywordText(text: string): string {
	return text.replace(/\s+/g, '');
}

function endCompletions(source: string, statementStart: number): KeywordSpec[] {
	const stack = openBlockClosers(source, statementStart);
	const top = stack[stack.length - 1];
	if (top?.startsWith('End ')) {
		return [keyword(top.slice(4), top, '0:0')];
	}
	if (top) {
		return [];
	}
	return [
		keyword('If', 'End If'),
		keyword('Sub', 'End Sub'),
		keyword('Function', 'End Function'),
		keyword('Property', 'End Property'),
		keyword('Select', 'End Select'),
		keyword('With', 'End With'),
		keyword('Type', 'End Type'),
		keyword('Enum', 'End Enum'),
	];
}

function closingCompletion(source: string, statementStart: number): KeywordSpec | undefined {
	const closers = openBlockClosers(source, statementStart);
	const closer = closers[closers.length - 1];
	return closer ? keyword(closer, `Close ${closer}`, '000:close') : undefined;
}

function openBlockClosers(source: string, statementStart: number): string[] {
	return openSmartBlockClosersBefore(source, statementStart);
}

function word(token: VbaToken | undefined): string {
	return token ? (token.canonicalText ?? token.rawText).toLowerCase() : '';
}

function isAccessModifier(wordText: string): boolean {
	return wordText === 'private' || wordText === 'public' || wordText === 'friend' || wordText === 'static';
}

function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

function keyword(label: string, detail: string, sortText?: string): KeywordSpec {
	return { label, insertText: label, detail, sortText, kind: 'keyword' };
}

function keywordSpecFromSmartBlock(spec: VbaSmartBlockSnippetSpec): KeywordSpec {
	return snippet(
		spec.label,
		spec.insertText,
		spec.detail,
		undefined,
		spec.matchText,
	);
}

function blockText(...lines: string[]): string {
	return lines.join('\n');
}

function snippet(
	label: string,
	insertText: string,
	detail: string,
	sortText?: string,
	matchText?: readonly string[],
): KeywordSpec {
	return {
		label,
		insertText,
		filterText: label,
		matchText,
		detail,
		sortText,
		kind: 'snippet',
	};
}
