// Keyword and block-snippet completion resolver.
//
// Provides canonical VBA keyword/snippet completions in statement contexts
// without replacing symbol completion. Narrow grammar contexts such as
// `Option `, `End `, and `On Error ` are exclusive because arbitrary identifiers
// are invalid there; broad statement starts are additive.

import { tokenize } from '../lexer/tokenize';
import { VbaToken } from '../lexer/tokenKinds';
import { splitLogicalStatements, codeTokens, tokenWord } from '../parser/parserState';

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

const STATEMENT_SNIPPETS: readonly KeywordSpec[] = [
	snippet('If', 'If ${1:condition} Then\n    $0\nEnd If', 'If...Then block'),
	snippet('If Else', 'If ${1:condition} Then\n    $2\nElse\n    $0\nEnd If', 'If...Else block', undefined, ['ifelse']),
	snippet('With', 'With ${1:object}\n    .$0\nEnd With', 'With...End With block'),
	snippet('For', 'For ${1:i} = ${2:1} To ${3:10}\n    $0\nNext ${1:i}', 'For...Next block'),
	snippet('For Each', 'For Each ${1:item} In ${2:collection}\n    $0\nNext ${1:item}', 'For Each...Next block'),
	snippet('Do While', 'Do While ${1:condition}\n    $0\nLoop', 'Do While...Loop block'),
	snippet('Do Until', 'Do Until ${1:condition}\n    $0\nLoop', 'Do Until...Loop block'),
	snippet('Do Loop Until', 'Do\n    $0\nLoop Until ${1:condition}', 'Do...Loop Until block', undefined, ['dountil']),
	snippet('While', 'While ${1:condition}\n    $0\nWend', 'While...Wend block'),
	snippet('Select Case', 'Select Case ${1:expression}\n    Case ${2:value}\n        $0\nEnd Select', 'Select Case block'),
	snippet('Sub', 'Sub ${1:Name}()\n    $0\nEnd Sub', 'Procedure block'),
	snippet('Function', 'Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function', 'Function block', undefined, ['func']),
	snippet('Property Get', 'Property Get ${1:Name}() As ${2:Variant}\n    $0\nEnd Property', 'Property Get block', undefined, ['propget']),
	snippet('Property Let', 'Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})\n    $0\nEnd Property', 'Property Let block', undefined, ['proplet']),
	snippet('Property Set', 'Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})\n    $0\nEnd Property', 'Property Set block', undefined, ['propset']),
	snippet('Type', 'Type ${1:Name}\n    ${2:Field} As ${3:Variant}\nEnd Type', 'User-defined type block'),
	snippet('Enum', 'Enum ${1:Name}\n    ${2:Value1} = ${3:0}\nEnd Enum', 'Enum block'),
	snippet('Private Sub', 'Private Sub ${1:Name}()\n    $0\nEnd Sub', 'Private procedure block'),
	snippet('Public Sub', 'Public Sub ${1:Name}()\n    $0\nEnd Sub', 'Public procedure block'),
	snippet('Private Function', 'Private Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function', 'Private function block'),
	snippet('Public Function', 'Public Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function', 'Public function block'),
	snippet('Option Explicit', 'Option Explicit', 'Option statement'),
	snippet('On Error Resume Next', 'On Error Resume Next', 'Error-handling statement'),
	snippet('On Error GoTo 0', 'On Error GoTo 0', 'Error-handling statement'),
	snippet(
		'On Error GoTo Handler',
		'On Error GoTo ${1:ErrHandler}\n    $0\n    Exit ${2|Sub,Function,Property|}\n${1:ErrHandler}:\n    MsgBox "Error " & Err.Number & ": " & Err.Description',
		'Error handler block',
		undefined,
		['onerror'],
	),
	snippet('Debug.Print', 'Debug.Print ${1:value}$0', 'Debug.Print statement', undefined, ['dp']),
];

const MODIFIER_SNIPPETS: readonly KeywordSpec[] = [
	snippet('Sub', 'Sub ${1:Name}()\n    $0\nEnd Sub', 'Procedure block'),
	snippet('Function', 'Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function', 'Function block', undefined, ['func']),
	snippet('Property Get', 'Property Get ${1:Name}() As ${2:Variant}\n    $0\nEnd Property', 'Property Get block', undefined, ['propget']),
	snippet('Property Let', 'Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})\n    $0\nEnd Property', 'Property Let block', undefined, ['proplet']),
	snippet('Property Set', 'Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})\n    $0\nEnd Property', 'Property Set block', undefined, ['propset']),
	keyword('Const', 'Constant declaration'),
	keyword('Dim', 'Variable declaration'),
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

const DIRECTIVE_SNIPPETS: readonly KeywordSpec[] = [
	snippet('#If', '#If ${1:condition} Then\n    $0\n#End If', 'Conditional compilation block'),
	snippet('#Const', '#Const ${1:name} = ${2:value}', 'Conditional compilation constant'),
	keyword('#ElseIf', 'Conditional compilation branch'),
	keyword('#Else', 'Conditional compilation branch'),
	keyword('#End If', 'End conditional compilation block'),
];

export function resolveKeywordCompletions(
	source: string,
	offset: number,
): KeywordCompletionResult {
	const ctx = completionContext(source, offset);
	if (!ctx) {
		return { items: [], exclusive: false };
	}

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
	if (first === '#') {
		return complete(DIRECTIVE_SNIPPETS, ctx.partial, true);
	}
	if (ctx.prefix.length === 1 && isAccessModifier(first)) {
		return complete(MODIFIER_SNIPPETS, ctx.partial, true);
	}
	if (ctx.atStatementStart) {
		const close = closingCompletion(source, ctx.statementStart);
		return complete(close ? [close, ...STATEMENT_SNIPPETS] : STATEMENT_SNIPPETS, ctx.partial, false);
	}
	return { items: [], exclusive: false };
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
	const tokens = tokenize(source.slice(0, Math.max(0, statementStart))).filter(
		(token) => token.kind !== 'comment',
	);
	const statements = splitLogicalStatements(tokens);
	const stack: string[] = [];
	for (const statement of statements) {
		const code = codeTokens(statement);
		if (code.length === 0) {
			continue;
		}
		const closer = closerForStatement(code);
		if (closer) {
			popCloser(stack, closer);
			continue;
		}
		const opener = openerForStatement(code);
		if (opener) {
			stack.push(opener);
		}
	}
	return stack;
}

function openerForStatement(tokens: readonly VbaToken[]): string | undefined {
	const w0 = tokenWord(tokens[0]);
	const w1 = tokenWord(tokens[1]);
	if (w0 === 'sub') {
		return 'End Sub';
	}
	if (w0 === 'function') {
		return 'End Function';
	}
	if (w0 === 'property') {
		return 'End Property';
	}
	if (w0 === 'type') {
		return 'End Type';
	}
	if (w0 === 'enum') {
		return 'End Enum';
	}
	if (w0 === 'if' && tokenWord(tokens[tokens.length - 1]) === 'then') {
		return 'End If';
	}
	if (w0 === 'for') {
		return 'Next';
	}
	if (w0 === 'do') {
		return 'Loop';
	}
	if (w0 === 'while') {
		return 'Wend';
	}
	if (w0 === 'with') {
		return 'End With';
	}
	if (w0 === 'select' && w1 === 'case') {
		return 'End Select';
	}
	return undefined;
}

function closerForStatement(tokens: readonly VbaToken[]): string | undefined {
	const w0 = tokenWord(tokens[0]);
	const w1 = tokenWord(tokens[1]);
	if (w0 === 'end') {
		switch (w1) {
			case 'sub':
				return 'End Sub';
			case 'function':
				return 'End Function';
			case 'property':
				return 'End Property';
			case 'if':
				return 'End If';
			case 'select':
				return 'End Select';
			case 'with':
				return 'End With';
			case 'type':
				return 'End Type';
			case 'enum':
				return 'End Enum';
			default:
				return undefined;
		}
	}
	if (w0 === 'next') {
		return 'Next';
	}
	if (w0 === 'loop') {
		return 'Loop';
	}
	if (w0 === 'wend') {
		return 'Wend';
	}
	return undefined;
}

function popCloser(stack: string[], closer: string): void {
	for (let i = stack.length - 1; i >= 0; i -= 1) {
		if (stack[i] === closer) {
			stack.splice(i, 1);
			return;
		}
	}
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
