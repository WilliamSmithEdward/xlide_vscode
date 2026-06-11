// Keyword and block-snippet completion resolver.
//
// Provides canonical VBA keyword/snippet completions in statement contexts
// without replacing symbol completion. Narrow grammar contexts such as
// `Option `, `End `, and `On Error ` are exclusive because arbitrary identifiers
// are invalid there; broad statement starts are additive.

import { VbaToken } from '../lexer/tokenKinds';
import { isIdentLike } from '../lexer/tokenHelpers';
import { completionCursorContext } from './cursorContext';
import {
	openSmartBlockClosersBefore,
	VBA_BLOCK_INDENT_UNIT,
	type VbaSmartBlockLayout,
} from '../../vbaSmartEnter';
import {
	vbaSmartBlockSnippetsFor,
	type VbaSmartBlockSnippetSpec,
} from '../../vbaSmartBlockSnippets';

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

const I = VBA_BLOCK_INDENT_UNIT;

const STATEMENT_SNIPPETS: readonly KeywordSpec[] = [
	snippet('Option Explicit', 'Option Explicit', 'Option statement'),
	snippet('Dim', 'Dim ${1:name} As ${2:Variant}', 'Variable declaration'),
	snippet('Set', 'Set ${1:object} = ${2:value}', 'Object assignment statement'),
	snippet('Let', 'Let ${1:variable} = ${2:value}', 'Explicit value assignment statement'),
	snippet('LSet', 'LSet ${1:target} = ${2:value}', 'Left-align string assignment'),
	snippet('RSet', 'RSet ${1:target} = ${2:value}', 'Right-align string assignment'),
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
	snippet('Call', 'Call ${1:ProcedureName}($0)', 'Explicit procedure call statement'),
	snippet('GoTo', 'GoTo ${1:label}', 'Branch to a procedure label'),
	snippet('GoSub', 'GoSub ${1:label}', 'Branch to a procedure subroutine label'),
	snippet('Resume', 'Resume ${1:label}', 'Resume at an error-handling label'),
	snippet('Resume Next', 'Resume Next', 'Resume after the statement that raised an error'),
	snippet('Erase', 'Erase ${1:array}', 'Erase array contents'),
	keyword('Stop', 'Break execution'),
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
	keyword('WithEvents', 'Object variable event declaration'),
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

const ON_BRANCH_SNIPPETS: readonly KeywordSpec[] = [
	keyword('GoTo', 'On expression GoTo label list'),
	keyword('GoSub', 'On expression GoSub label list'),
];

const ON_ERROR_SNIPPETS: readonly KeywordSpec[] = [
	snippet('GoTo 0', 'GoTo 0', 'Disable active error handler'),
	snippet('GoTo -1', 'GoTo -1', 'Clear current error state'),
	snippet('GoTo label', 'GoTo ${1:label}', 'Branch to an error handler label'),
	snippet('Resume Next', 'Resume Next', 'Continue after the statement that raised an error'),
];

const AS_SNIPPETS: readonly KeywordSpec[] = [
	keyword('As', 'Declaration type clause'),
];

const PARAMETER_MODIFIER_SNIPPETS: readonly KeywordSpec[] = [
	keyword('ByVal', 'Pass parameter by value'),
	keyword('ByRef', 'Pass parameter by reference'),
	keyword('Optional', 'Optional parameter'),
	keyword('ParamArray', 'Variable-length parameter list'),
];

const PARAMETER_PASSING_SNIPPETS: readonly KeywordSpec[] = [
	keyword('ByVal', 'Pass parameter by value'),
	keyword('ByRef', 'Pass parameter by reference'),
];

const THEN_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Then', 'If condition terminator'),
];

const FOR_TO_SNIPPETS: readonly KeywordSpec[] = [
	keyword('To', 'For loop upper bound'),
];

const FOR_STEP_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Step', 'For loop increment'),
];

const REDIM_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Preserve', 'Preserve dynamic array values'),
];

const SET_ASSIGNMENT_SNIPPETS: readonly KeywordSpec[] = [
	snippet('New', 'New ${1:ClassName}', 'Object creation expression'),
];

const RESUME_SNIPPETS: readonly KeywordSpec[] = [
	keyword('Next', 'Resume Next'),
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

const IF_BRANCH_SNIPPETS: readonly KeywordSpec[] = [
	snippet('ElseIf', 'ElseIf ${1:condition} Then', 'ElseIf branch'),
	keyword('Else', 'Else branch'),
];

const SELECT_BRANCH_SNIPPETS: readonly KeywordSpec[] = [
	snippet('Case', 'Case ${1:value}', 'Select Case branch'),
	snippet('Case Is', 'Case Is ${1:operator} ${2:value}', 'Select Case comparison branch'),
	keyword('Case Else', 'Select Case fallback branch'),
];

const DO_CLOSE_SNIPPETS: readonly KeywordSpec[] = [
	snippet('Loop Until', 'Loop Until ${1:condition}', 'Close Do block with an Until condition'),
	snippet('Loop While', 'Loop While ${1:condition}', 'Close Do block with a While condition'),
];

const DIRECTIVE_END_SNIPPETS: readonly KeywordSpec[] = [
	keyword('If', '#End If'),
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
		if (ctx.prefix.length === 1) {
			return complete(ON_SNIPPETS, ctx.partial, true);
		}
		if (isOnExpressionBranchContext(ctx)) {
			return completeNarrow(ON_BRANCH_SNIPPETS, ctx.partial);
		}
		return { items: [], exclusive: false };
	}
	if (isThenContext(ctx)) {
		return completeNarrow(THEN_SNIPPETS, ctx.partial);
	}
	if (first === 'resume' && ctx.prefix.length === 1) {
		return complete(RESUME_SNIPPETS, ctx.partial, false);
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
	const forCountingCompletion = forCountingCompletionKind(ctx);
	if (forCountingCompletion === 'to') {
		return completeNarrow(FOR_TO_SNIPPETS, ctx.partial);
	}
	if (forCountingCompletion === 'step') {
		return completeNarrow(FOR_STEP_SNIPPETS, ctx.partial);
	}
	if (first === 'redim' && ctx.prefix.length === 1) {
		return complete(REDIM_SNIPPETS, ctx.partial, false);
	}
	const parameterModifierSnippets = parameterModifierSnippetsForContext(ctx);
	if (parameterModifierSnippets) {
		return complete(parameterModifierSnippets, ctx.partial, false);
	}
	if (isDeclarationAsContext(ctx) || isParameterAsContext(ctx) || isReturnAsContext(ctx)) {
		return complete(AS_SNIPPETS, ctx.partial, false);
	}
	if (isSetNewContext(ctx)) {
		return complete(SET_ASSIGNMENT_SNIPPETS, ctx.partial, false);
	}
	if (first === '#') {
		if (second === 'end') {
			return complete(DIRECTIVE_END_SNIPPETS, ctx.partial, true);
		}
		if (ctx.prefix.length > 1) {
			return { items: [], exclusive: false };
		}
		return complete(directiveSnippets(blockLayout), ctx.partial, true);
	}
	if (ctx.prefix.length === 1 && isAccessModifier(first)) {
		return complete(modifierSnippets(blockLayout), ctx.partial, true);
	}
	if (ctx.atStatementStart) {
		const close = closingCompletion(source, ctx.statementStart);
		const branchSnippets = branchSnippetsBefore(source, ctx.statementStart);
		const snippets = statementSnippets(blockLayout);
		return complete(close ? [close, ...branchSnippets, ...snippets] : [...branchSnippets, ...snippets], ctx.partial, false);
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
	const cursor = completionCursorContext(source, offset);
	if (cursor.inComment || cursor.inString) {
		return undefined;
	}

	const statementStart = cursor.statementStart;
	const statementTokens = cursor.tokens
		.filter((token) => token.start >= statementStart && token.kind !== 'comment' && token.kind !== 'newline' && token.kind !== 'colon');
	let partial = '';
	if (cursor.partialToken && statementTokens[statementTokens.length - 1] === cursor.partialToken) {
		partial = cursor.partialToken.rawText;
		statementTokens.pop();
	}

	const prefix = statementTokens;
	const atStatementStart = prefix.length === 0;
	return { partial, prefix, atStatementStart, statementStart };
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

function completeNarrow(
	specs: readonly KeywordSpec[],
	partial: string,
): KeywordCompletionResult {
	const result = complete(specs, partial, true);
	return result.items.length > 0
		? result
		: { items: [], exclusive: false };
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

function isThenContext(ctx: CompletionContext): boolean {
	const first = word(ctx.prefix[0]);
	const second = word(ctx.prefix[1]);
	if ((first === 'if' || first === 'elseif') && ctx.prefix.length > 1) {
		return !containsWord(ctx.prefix, 'then') && canEndExpression(ctx.prefix[ctx.prefix.length - 1]);
	}
	if (first === '#' && (second === 'if' || second === 'elseif') && ctx.prefix.length > 2) {
		return !containsWord(ctx.prefix, 'then') && canEndExpression(ctx.prefix[ctx.prefix.length - 1]);
	}
	return false;
}

function isOnExpressionBranchContext(ctx: CompletionContext): boolean {
	if (word(ctx.prefix[0]) !== 'on') {
		return false;
	}
	if (containsWord(ctx.prefix, 'goto') || containsWord(ctx.prefix, 'gosub')) {
		return false;
	}
	return ctx.prefix.length > 1 && canEndExpression(ctx.prefix[ctx.prefix.length - 1]);
}

function forCountingCompletionKind(ctx: CompletionContext): 'to' | 'step' | undefined {
	if (word(ctx.prefix[0]) !== 'for' || word(ctx.prefix[1]) === 'each') {
		return undefined;
	}
	const eqIndex = ctx.prefix.findIndex((token) => token.rawText === '=');
	if (eqIndex < 2) {
		return undefined;
	}
	const toIndex = findWordIndex(ctx.prefix, 'to', eqIndex + 1);
	if (toIndex < 0) {
		return hasExpressionAfter(ctx.prefix, eqIndex) ? 'to' : undefined;
	}
	if (findWordIndex(ctx.prefix, 'step', toIndex + 1) >= 0) {
		return undefined;
	}
	return hasExpressionAfter(ctx.prefix, toIndex) ? 'step' : undefined;
}

function hasExpressionAfter(tokens: readonly VbaToken[], index: number): boolean {
	return tokens.length > index + 1 && canEndExpression(tokens[tokens.length - 1]);
}

function isDeclarationAsContext(ctx: CompletionContext): boolean {
	const nameStart = declarationNameStart(ctx.prefix);
	if (nameStart === undefined) {
		return false;
	}
	return declarationSegmentNeedsAs(ctx.prefix, nameStart);
}

function declarationNameStart(tokens: readonly VbaToken[]): number | undefined {
	let i = 0;
	while (isDeclarationPrefixModifier(word(tokens[i]))) {
		i += 1;
	}
	const head = word(tokens[i]);
	if (head === 'dim' || head === 'const' || head === 'static') {
		return i + 1;
	}
	if (i > 0 && isNameLike(tokens[i])) {
		return i;
	}
	return undefined;
}

function declarationSegmentNeedsAs(tokens: readonly VbaToken[], nameStart: number): boolean {
	const comma = lastRawIndex(tokens, ',', nameStart);
	const segmentStart = comma >= 0 ? comma + 1 : nameStart;
	const segment = tokens.slice(segmentStart);
	if (segment.length === 0 || containsWord(segment, 'as')) {
		return false;
	}
	return segment.some(isNameLike) && canEndDeclarator(segment[segment.length - 1]);
}

function parameterModifierSnippetsForContext(ctx: CompletionContext): readonly KeywordSpec[] | undefined {
	const segment = currentParameterSegment(ctx.prefix);
	if (!segment || containsWord(segment, 'as')) {
		return undefined;
	}
	const words = segment.map(word).filter(Boolean);
	if (words.length === 0) {
		return PARAMETER_MODIFIER_SNIPPETS;
	}
	if (words.length === 1 && words[0] === 'optional') {
		return PARAMETER_PASSING_SNIPPETS;
	}
	return undefined;
}

function isParameterAsContext(ctx: CompletionContext): boolean {
	const segment = currentParameterSegment(ctx.prefix);
	if (!segment || segment.length === 0 || containsWord(segment, 'as')) {
		return false;
	}
	const nameTokens = segment.filter((token) => !isParameterModifier(word(token)));
	return nameTokens.some(isNameLike) && canEndDeclarator(segment[segment.length - 1]);
}

function currentParameterSegment(tokens: readonly VbaToken[]): readonly VbaToken[] | undefined {
	const open = lastUnclosedParenIndex(tokens);
	if (open < 0 || !hasProcedureHeaderBefore(tokens, open)) {
		return undefined;
	}
	const comma = lastRawIndex(tokens, ',', open + 1);
	const start = comma >= 0 ? comma + 1 : open + 1;
	return tokens.slice(start);
}

function isReturnAsContext(ctx: CompletionContext): boolean {
	const close = lastRawIndex(ctx.prefix, ')');
	if (close < 0 || close !== ctx.prefix.length - 1) {
		return false;
	}
	const beforeClose = ctx.prefix.slice(0, close);
	if (!hasReturnableProcedureHeader(beforeClose)) {
		return false;
	}
	return !containsWord(ctx.prefix.slice(close + 1), 'as');
}

function isSetNewContext(ctx: CompletionContext): boolean {
	return word(ctx.prefix[0]) === 'set' && ctx.prefix[ctx.prefix.length - 1]?.rawText === '=';
}

function hasProcedureHeaderBefore(tokens: readonly VbaToken[], openIndex: number): boolean {
	const beforeOpen = tokens.slice(0, openIndex);
	return (
		containsWord(beforeOpen, 'sub') ||
		containsWord(beforeOpen, 'function') ||
		containsWord(beforeOpen, 'property') ||
		containsWord(beforeOpen, 'declare')
	);
}

function hasReturnableProcedureHeader(tokens: readonly VbaToken[]): boolean {
	if (containsWord(tokens, 'function')) {
		return true;
	}
	const propertyIndex = findWordIndex(tokens, 'property');
	return propertyIndex >= 0 && word(tokens[propertyIndex + 1]) === 'get';
}

function lastUnclosedParenIndex(tokens: readonly VbaToken[]): number {
	const stack: number[] = [];
	for (let i = 0; i < tokens.length; i += 1) {
		const raw = tokens[i].rawText;
		if (raw === '(') {
			stack.push(i);
		} else if (raw === ')') {
			stack.pop();
		}
	}
	return stack[stack.length - 1] ?? -1;
}

function lastRawIndex(tokens: readonly VbaToken[], rawText: string, start = 0): number {
	for (let i = tokens.length - 1; i >= start; i -= 1) {
		if (tokens[i].rawText === rawText) {
			return i;
		}
	}
	return -1;
}

function findWordIndex(tokens: readonly VbaToken[], target: string, start = 0): number {
	for (let i = start; i < tokens.length; i += 1) {
		if (word(tokens[i]) === target) {
			return i;
		}
	}
	return -1;
}

function containsWord(tokens: readonly VbaToken[], target: string): boolean {
	return findWordIndex(tokens, target) >= 0;
}

function canEndExpression(token: VbaToken | undefined): boolean {
	if (!token) {
		return false;
	}
	return (
		isNameLike(token) ||
		isExpressionEndingKeyword(word(token)) ||
		token.kind === 'integerLiteral' ||
		token.kind === 'floatLiteral' ||
		token.kind === 'dateLiteral' ||
		token.kind === 'stringLiteral' ||
		token.rawText === ')'
	);
}

function isExpressionEndingKeyword(wordText: string): boolean {
	return (
		wordText === 'empty' ||
		wordText === 'false' ||
		wordText === 'me' ||
		wordText === 'nothing' ||
		wordText === 'null' ||
		wordText === 'true'
	);
}

function canEndDeclarator(token: VbaToken | undefined): boolean {
	return canEndExpression(token);
}

function isDeclarationPrefixModifier(wordText: string): boolean {
	return wordText === 'private' || wordText === 'public' || wordText === 'friend' || wordText === 'global';
}

function isParameterModifier(wordText: string): boolean {
	return (
		wordText === 'byval' ||
		wordText === 'byref' ||
		wordText === 'optional' ||
		wordText === 'paramarray'
	);
}

function isNameLike(token: VbaToken | undefined): boolean {
	return Boolean(token && (
		token.kind === 'identifier' ||
		token.kind === 'bracketedIdentifier'
	));
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

function branchSnippetsBefore(source: string, statementStart: number): readonly KeywordSpec[] {
	const stack = openBlockClosers(source, statementStart);
	const top = stack[stack.length - 1];
	if (top === 'End If') {
		return IF_BRANCH_SNIPPETS;
	}
	if (top === 'End Select') {
		return SELECT_BRANCH_SNIPPETS;
	}
	if (top === 'Loop') {
		return DO_CLOSE_SNIPPETS;
	}
	return [];
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
	return (
		wordText === 'private' ||
		wordText === 'public' ||
		wordText === 'friend' ||
		wordText === 'global' ||
		wordText === 'static'
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
