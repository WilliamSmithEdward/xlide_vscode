// VBA module formatter.
//
// Does what the VBE does to a line when you leave it, plus the one thing the
// VBE never does: it re-indents the whole module by block structure. Three
// passes over the token stream, none of which can change what the code means:
//
//   1. Indentation. Every logical line is placed at the depth its enclosing
//      blocks give it: procedure and Type/Enum bodies one level in, If/For/
//      Do/While/With bodies one level in, Select Case arms one level in and
//      their statements one more, `#If` arms one level in. Labels and line
//      numbers sit at column 1 the way the VBE keeps them. Continuation
//      lines keep their extra indentation relative to the line they continue
//      when they have any, and get one level otherwise.
//   2. Keyword casing. Every keyword token takes the canonical spelling the
//      lexer already knows (`end sub` -> `End Sub`). An optional resolver
//      does the same for identifiers from their declarations.
//   3. Spacing. The spaces the VBE inserts are inserted: around `=` and the
//      comparison operators, after `,`, `;` and a statement-separating `:`,
//      and around binary operators. Nothing is ever removed between tokens,
//      so aligned comments and aligned `Const` blocks survive.
//
// The formatter never touches strings, comments, `Attribute` lines, or the
// text of any token. It refuses to return a result whose token stream differs
// from the input's (see {@link tokenStreamDifference}), so the worst it can do
// is nothing.
//
// Pure: no `vscode` dependency. The editor layer turns the result into edits.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { lineIndexOf, lineStartOffsets } from '../../vbaSourceScan';
import {
	MARKER_KEYWORDS,
	OPERATOR_IDENTIFIERS,
	STATEMENT_KEYWORDS,
} from '../lexer/keywordTable';
import { tokenWord } from '../lexer/tokenHelpers';

export interface VbaFormatOptions {
	/** Columns per indent level, and the width a tab counts for in existing indents. */
	tabSize: number;
	/** Indent with spaces (true) or with tabs (false). */
	insertSpaces: boolean;
	/**
	 * The canonical spelling of an identifier, or undefined to leave it as
	 * written. Called for every identifier that is not a member name (nothing
	 * after `.` or `!`) and not a named-argument name, with the token's
	 * absolute offset so a resolver can scope locals to their procedure.
	 */
	identifierCase?: (name: string, offset: number) => string | undefined;
}

export interface VbaFormatResult {
	/** The formatted module, or undefined when the formatter declined. */
	text: string | undefined;
	/** Why it declined: the output would not lex to the input's tokens. */
	refusal?: string;
}

type BlockKind =
	| 'Procedure'
	| 'Type'
	| 'Enum'
	| 'If'
	| 'For'
	| 'Do'
	| 'While'
	| 'With'
	| 'Select'
	| 'PreprocessorIf';

interface OpenBlock {
	kind: BlockKind;
	/** Indent level of the opener line; the body sits one deeper. */
	level: number;
	/** True for a `#If` opened outside every procedure, which a procedure header keeps. */
	moduleLevel: boolean;
}

type StatementClass =
	| { type: 'plain' }
	| { type: 'label' }
	| { type: 'attribute' }
	| { type: 'lineNumber'; rest: VbaToken[] }
	| { type: 'procedure' }
	| { type: 'opener'; kind: BlockKind }
	| { type: 'closer'; kind: BlockKind; count: number }
	| { type: 'mid'; kind: 'If' | 'Select' | 'PreprocessorIf' };

interface Statement {
	tokens: VbaToken[];
	/** True when a `:` separator followed this statement on its logical line. */
	endedByColon: boolean;
}

const PROCEDURE_MODIFIERS = new Set(['public', 'private', 'friend', 'global', 'static']);
const TYPE_MODIFIERS = new Set(['public', 'private', 'global']);
const END_CLOSERS: Readonly<Record<string, BlockKind>> = {
	sub: 'Procedure',
	function: 'Procedure',
	property: 'Procedure',
	if: 'If',
	with: 'With',
	select: 'Select',
	type: 'Type',
	enum: 'Enum',
};

/** Operators that take an operand on both sides. */
const BINARY_OPERATOR_SYMBOLS = new Set(['+', '-', '*', '/', '\\', '^', '&']);
const BINARY_OPERATOR_WORDS = new Set(['mod', 'and', 'or', 'xor', 'eqv', 'imp', 'like', 'is']);
const COMPARISON_OR_ASSIGN = new Set(['=', '<', '>', '<=', '>=', '<>']);
/** Type-declaration characters, plus the `!` bang: glued to a name they are part of it. */
const NAME_SUFFIX_CHARS = new Set(['&', '%', '#', '!', '@', '$']);

/**
 * Keywords that cannot end an operand, so a `-` or `+` after one is unary.
 * Everything else the lexer calls a keyword (`Me`, `True`, `Date`, `Len`, the
 * type names, the contextual words that double as member names) can.
 */
const NON_OPERAND_KEYWORDS: ReadonlySet<string> = new Set([
	...STATEMENT_KEYWORDS,
	...MARKER_KEYWORDS,
	...OPERATOR_IDENTIFIERS,
	'Step',
	'Property',
	'Rem',
].map((word) => word.toLowerCase()));

function isNameToken(token: VbaToken | undefined): boolean {
	return !!token && (token.kind === 'identifier' || token.kind === 'keyword' || token.kind === 'bracketedIdentifier');
}

/** Formats a whole VBA module. See the file comment for what changes and what never does. */
export function formatVbaModule(source: string, options: VbaFormatOptions): VbaFormatResult {
	const tabSize = Math.max(1, Math.floor(options.tabSize) || 4);
	const lines = source.split(/\r\n|\r|\n/);
	const eols = source.match(/\r\n|\r|\n/g) ?? [];
	const tokens = tokenize(source);
	const lineStarts = lineStartOffsets(source);

	// Physical line -> its tokens (comments included, newlines left out).
	const tokensByLine: VbaToken[][] = lines.map(() => []);
	// A physical line that continues the one above it (that line ended in ` _`).
	const continuation: boolean[] = lines.map(() => false);
	const markContinuation = (trivia: { kind: string; start: number }): void => {
		if (trivia.kind !== 'lineContinuation') {
			return;
		}
		const underscoreLine = lineIndexOf(lineStarts, trivia.start);
		if (underscoreLine + 1 < lines.length) {
			continuation[underscoreLine + 1] = true;
		}
	};
	for (const token of tokens) {
		for (const trivia of token.leadingTrivia ?? []) {
			markContinuation(trivia);
		}
		for (const trivia of token.trailingTrivia ?? []) {
			markContinuation(trivia);
		}
		if (token.kind !== 'newline' && token.line < lines.length) {
			tokensByLine[token.line].push(token);
		}
	}

	const output: string[] = lines.map((line) => line);
	const stack: OpenBlock[] = [];
	let level = 0;

	const find = (kind: BlockKind): number => {
		for (let i = stack.length - 1; i >= 0; i--) {
			if (stack[i].kind === kind) {
				return i;
			}
		}
		return -1;
	};
	const inProcedure = (): boolean => stack.some((block) => block.kind === 'Procedure');

	let physical = 0;
	while (physical < lines.length) {
		const first = physical;
		let last = physical;
		while (last + 1 < lines.length && continuation[last + 1]) {
			last++;
		}
		physical = last + 1;

		const significant: VbaToken[] = [];
		for (let i = first; i <= last; i++) {
			for (const token of tokensByLine[i]) {
				if (token.kind !== 'comment') {
					significant.push(token);
				}
			}
		}
		const statements = splitStatements(significant);

			if (statements.length > 0 && classify(statements[0], true, statements.length === 1, stack).type === 'attribute') {
			// Attribute lines are the VBE's own metadata, written its way.
			continue;
		}

		// The first statement decides where the line sits; every statement
		// on the line moves the stack for the lines after it.
		let lineIndent = level;
		statements.forEach((statement, index) => {
			const isFirst = index === 0;
			const isLast = index === statements.length - 1;
			const cls = classify(statement, isFirst, isLast, stack);
			lineIndent = apply(cls, isFirst, lineIndent);
		});

		const firstWidth = lineIndent * tabSize;
		const originalFirstWidth = indentWidth(lines[first], tabSize);
		for (let i = first; i <= last; i++) {
			let width = firstWidth;
			if (i > first) {
				const delta = indentWidth(lines[i], tabSize) - originalFirstWidth;
				width = firstWidth + (delta > 0 ? delta : tabSize);
			}
			const body = rewriteLine(lines[i], tokensByLine[i], options);
			output[i] = indentString(width, tabSize, options.insertSpaces) + body;
		}
	}

	function apply(cls: StatementClass, isFirst: boolean, lineIndent: number): number {
		switch (cls.type) {
			case 'plain':
			case 'attribute':
				return lineIndent;
			case 'label':
				return isFirst ? 0 : lineIndent;
			case 'lineNumber': {
				const rest = classify({ tokens: cls.rest, endedByColon: false }, false, true, stack);
				apply(rest, false, lineIndent);
				return isFirst ? 0 : lineIndent;
			}
			case 'closer': {
				let indent = lineIndent;
				for (let n = 0; n < cls.count; n++) {
					const index = find(cls.kind);
					if (index < 0) {
						if (cls.kind === 'Procedure') {
							// A stray End Sub still ends whatever was open: back
							// to module level, where the next line belongs.
							stack.length = 0;
							level = 0;
							indent = 0;
						}
						break;
					}
					const target = stack[index];
					stack.length = index;
					level = target.level;
					indent = target.level;
				}
				return isFirst ? indent : lineIndent;
			}
			case 'mid': {
				const index = find(cls.kind);
				if (index < 0) {
					return lineIndent;
				}
				const target = stack[index];
				stack.length = index + 1;
				const own = cls.kind === 'Select' ? target.level + 1 : target.level;
				level = own + 1;
				return isFirst ? own : lineIndent;
			}
			case 'procedure': {
				while (stack.length > 0) {
					const top = stack[stack.length - 1];
					if (top.kind === 'PreprocessorIf' && top.moduleLevel) {
						break;
					}
					stack.pop();
				}
				const at = stack.length > 0 ? stack[stack.length - 1].level + 1 : 0;
				stack.push({ kind: 'Procedure', level: at, moduleLevel: false });
				level = at + 1;
				return isFirst ? at : lineIndent;
			}
			case 'opener': {
				const at = isFirst ? lineIndent : level;
				stack.push({ kind: cls.kind, level: at, moduleLevel: !inProcedure() });
				level = at + 1;
				return lineIndent;
			}
		}
	}

	const text = output.map((line, i) => line + (eols[i] ?? '')).join('');
	const difference = tokenStreamDifference(source, text);
	if (difference) {
		return { text: undefined, refusal: difference };
	}
	return { text };
}

/** Splits a logical line's significant tokens into `:`-separated statements. */
function splitStatements(tokens: readonly VbaToken[]): Statement[] {
	const out: Statement[] = [];
	let current: VbaToken[] = [];
	for (const token of tokens) {
		if (token.kind === 'colon') {
			if (current.length > 0) {
				out.push({ tokens: current, endedByColon: true });
			}
			current = [];
			continue;
		}
		current.push(token);
	}
	if (current.length > 0) {
		out.push({ tokens: current, endedByColon: false });
	}
	// A `:` after the last statement only shows on that statement.
	return out;
}

function classify(
	statement: Statement,
	isFirst: boolean,
	isLast: boolean,
	stack: readonly OpenBlock[],
): StatementClass {
	const tokens = statement.tokens;
	if (tokens.length === 0) {
		return { type: 'plain' };
	}
	const head = tokens[0];
	if (isFirst && statement.endedByColon && tokens.length === 1 && head.kind === 'identifier') {
		// `Retry:` - a line label, which VBA reads before a call by that name.
		return { type: 'label' };
	}
	if (head.kind === 'directive') {
		const word = tokenWord(tokens[1]);
		if (word === 'if') {
			return { type: 'opener', kind: 'PreprocessorIf' };
		}
		if (word === 'elseif' || word === 'else') {
			return { type: 'mid', kind: 'PreprocessorIf' };
		}
		if (word === 'endif' || (word === 'end' && tokenWord(tokens[2]) === 'if')) {
			return { type: 'closer', kind: 'PreprocessorIf', count: 1 };
		}
		return { type: 'plain' };
	}
	if (isFirst && head.kind === 'integerLiteral') {
		return { type: 'lineNumber', rest: tokens.slice(1) };
	}
	if (isFirst && head.kind === 'identifier' && head.rawText.toLowerCase() === 'attribute') {
		return { type: 'attribute' };
	}
	const first = tokenWord(head);
	if (first === 'end') {
		const kind = END_CLOSERS[tokenWord(tokens[1])];
		return kind ? { type: 'closer', kind, count: 1 } : { type: 'plain' };
	}
	if (first === 'loop') {
		return { type: 'closer', kind: 'Do', count: 1 };
	}
	if (first === 'wend') {
		return { type: 'closer', kind: 'While', count: 1 };
	}
	if (first === 'next') {
		// `Next i, j` closes two loops.
		let count = 1;
		for (const token of tokens.slice(1)) {
			if (token.kind === 'punctuation' && token.rawText === ',') {
				count++;
			}
		}
		return { type: 'closer', kind: 'For', count };
	}
	if (first === 'else' || first === 'elseif') {
		return { type: 'mid', kind: 'If' };
	}
	if (first === 'case') {
		return { type: 'mid', kind: 'Select' };
	}
	if (first === 'select' && tokenWord(tokens[1]) === 'case') {
		return { type: 'opener', kind: 'Select' };
	}
	if (first === 'if') {
		// A block If ends its logical line with Then; a single-line If carries
		// its statements after Then, or after a `:` (MS-VBAL 5.4.2.1).
		return isLast && tokenWord(tokens[tokens.length - 1]) === 'then' && tokens.length > 1
			? { type: 'opener', kind: 'If' }
			: { type: 'plain' };
	}
	if (first === 'for') {
		return { type: 'opener', kind: 'For' };
	}
	if (first === 'do') {
		return { type: 'opener', kind: 'Do' };
	}
	if (first === 'while') {
		return { type: 'opener', kind: 'While' };
	}
	if (first === 'with') {
		return { type: 'opener', kind: 'With' };
	}

	// Declarations: modifiers, then the word that says what is declared.
	let index = 0;
	while (index < tokens.length && PROCEDURE_MODIFIERS.has(tokenWord(tokens[index]))) {
		index++;
	}
	const declared = tokenWord(tokens[index]);
	if (declared === 'sub' || declared === 'function') {
		return { type: 'procedure' };
	}
	if (declared === 'property') {
		const accessor = tokenWord(tokens[index + 1]);
		if (accessor === 'get' || accessor === 'let' || accessor === 'set') {
			return { type: 'procedure' };
		}
		return { type: 'plain' };
	}
	const top = stack[stack.length - 1];
	const insideRecord = top !== undefined && (top.kind === 'Type' || top.kind === 'Enum');
	if ((declared === 'type' || declared === 'enum') && !insideRecord) {
		let modifiers = 0;
		while (modifiers < index && TYPE_MODIFIERS.has(tokenWord(tokens[modifiers]))) {
			modifiers++;
		}
		const name = tokens[index + 1];
		if (modifiers === index && isNameToken(name) && tokenWord(name) !== 'as') {
			return { type: 'opener', kind: declared === 'type' ? 'Type' : 'Enum' };
		}
	}
	return { type: 'plain' };
}

/**
 * Rewrites one physical line's content: tokens in order with their original
 * gaps, keyword and identifier casing applied, the VBE's spaces inserted where
 * two tokens touch, trailing whitespace dropped (a continuation's ` _` stays).
 * The leading indentation is left to the caller.
 */
function rewriteLine(
	line: string,
	lineTokens: readonly VbaToken[],
	options: VbaFormatOptions,
): string {
	if (lineTokens.length === 0) {
		return trailing(line, 0);
	}
	let out = '';
	let prevEndCol = lineTokens[0].character;
	let prev: VbaToken | undefined;
	let prevBinary = false;
	let prevSuffix = false;
	let prevAfterDot = false;
	// True while the statement so far is a name chain (`Foo`, `obj.Method`):
	// an operator here begins the first argument of a call, so `Foo -1`
	// keeps its unary minus.
	let statementHead = true;
	for (let i = 0; i < lineTokens.length; i++) {
		const token = lineTokens[i];
		let gap = line.slice(prevEndCol, token.character);
		prevEndCol = token.character + token.rawText.length;
		if (token.kind === 'comment') {
			out += gap + token.rawText;
			continue;
		}
		const spacing = classifySpacing(prev, token, prevBinary, prevSuffix, prevAfterDot, statementHead);
		if (gap === '' && spacing.spaceBefore) {
			gap = ' ';
		}
		const afterDot = !!prev && ((prev.kind === 'punctuation' && prev.rawText === '.')
			|| (prev.kind === 'operator' && prev.rawText === '!'));
		const next = lineTokens[i + 1];
		const namedArgument = !!next && next.kind === 'operator' && next.rawText === ':=';
		out += gap + tokenText(token, options, afterDot, namedArgument);
		const isDot = (token.kind === 'punctuation' && token.rawText === '.')
			|| (token.kind === 'operator' && token.rawText === '!');
		if (token.kind === 'colon') {
			statementHead = true;
		} else if (statementHead) {
			// The chain continues through dots and the names after them; a
			// second name with no dot before it is the first argument.
			const chainName = isNameToken(token) && (i === 0 || afterDot || prev?.kind === 'colon');
			if (!isDot && !chainName) {
				statementHead = false;
			}
		}
		prev = token;
		prevBinary = spacing.binary;
		prevSuffix = spacing.suffix;
		prevAfterDot = afterDot;
	}
	if (prev && prev.kind === 'unknown' && prev.rawText === '_') {
		// A stray `_` with whitespace after it is not a continuation; dropping
		// that whitespace would make it one and join the next line.
		return out + line.slice(prevEndCol);
	}
	return out + trailing(line, prevEndCol);
}

/** The text after the last token: nothing, or the ` _` of a continuation. */
function trailing(line: string, from: number): string {
	const rest = line.slice(from).replace(/[ \t　]+$/, '');
	if (rest.endsWith('_')) {
		return from === 0 ? '_' : rest;
	}
	return '';
}

function tokenText(
	token: VbaToken,
	options: VbaFormatOptions,
	memberName: boolean,
	namedArgument: boolean,
): string {
	if (token.kind === 'keyword') {
		return token.canonicalText ?? token.rawText;
	}
	if (token.kind === 'identifier' && options.identifierCase && !memberName && !namedArgument) {
		const canonical = options.identifierCase(token.rawText, token.start);
		if (canonical && canonical.toLowerCase() === token.rawText.toLowerCase()) {
			return canonical;
		}
	}
	return token.rawText;
}

interface Spacing {
	spaceBefore: boolean;
	/** The current token is a binary operator: the token after it wants a space too. */
	binary: boolean;
	/** The current token is a type-declaration character or bang glued to a name. */
	suffix: boolean;
}

function classifySpacing(
	prev: VbaToken | undefined,
	cur: VbaToken,
	prevBinary: boolean,
	prevSuffix: boolean,
	prevAfterDot: boolean,
	statementHead: boolean,
): Spacing {
	if (!prev) {
		return { spaceBefore: false, binary: false, suffix: false };
	}
	const curRaw = cur.rawText;
	if (statementHead && isNameToken(prev) && cur.kind === 'operator' && (curRaw === '-' || curRaw === '+')) {
		// `Foo -1` / `Debug.Print -1`: the sign opens the call's first argument.
		return { spaceBefore: true, binary: false, suffix: false };
	}
	if ((cur.kind === 'operator' || cur.kind === 'unknown') && NAME_SUFFIX_CHARS.has(curRaw) && isNameToken(prev)) {
		// `x&`, `s$`, `rs!Field`: the character belongs to the name. Whether a
		// glued `&` is a suffix or an operator is decided by VBA the same way,
		// so the gap is left exactly as written.
		return { spaceBefore: false, binary: false, suffix: true };
	}
	if (cur.kind === 'unknown') {
		// A character the lexer has no place for, such as the `_` of `&_`.
		// A space before it could make it a line continuation.
		return { spaceBefore: false, binary: false, suffix: false };
	}
	if (prev.kind === 'colon' || (prev.kind === 'punctuation' && (prev.rawText === ',' || prev.rawText === ';'))) {
		return { spaceBefore: true, binary: false, suffix: false };
	}
	if (prev.kind === 'keyword' && !prevAfterDot && NON_OPERAND_KEYWORDS.has(tokenWord(prev))
		&& (cur.kind === 'operator' || cur.kind === 'integerLiteral' || cur.kind === 'floatLiteral'
			|| cur.kind === 'stringLiteral' || cur.kind === 'dateLiteral')) {
		// `Step-1`, `Print#1`, `Case-1`: a word and what follows it.
		return { spaceBefore: true, binary: false, suffix: false };
	}
	if (cur.kind === 'operator' && COMPARISON_OR_ASSIGN.has(curRaw)) {
		return { spaceBefore: true, binary: false, suffix: false };
	}
	if (prev.kind === 'operator' && COMPARISON_OR_ASSIGN.has(prev.rawText)) {
		return { spaceBefore: true, binary: false, suffix: false };
	}
	const isBinaryOperator = (cur.kind === 'operator' && BINARY_OPERATOR_SYMBOLS.has(curRaw))
		|| (cur.kind === 'keyword' && BINARY_OPERATOR_WORDS.has(tokenWord(cur)));
	if (isBinaryOperator) {
		const binary = prevSuffix || operandEnds(prev, prevAfterDot);
		return { spaceBefore: binary || prevBinary, binary, suffix: false };
	}
	return { spaceBefore: prevBinary, binary: false, suffix: false };
}

/** True when `token` can be the last token of an operand, so an operator after it is binary. */
function operandEnds(token: VbaToken, afterDot: boolean): boolean {
	switch (token.kind) {
		case 'identifier':
		case 'bracketedIdentifier':
		case 'integerLiteral':
		case 'floatLiteral':
		case 'stringLiteral':
		case 'dateLiteral':
			return true;
		case 'punctuation':
			return token.rawText === ')';
		case 'keyword':
			return afterDot || !NON_OPERAND_KEYWORDS.has(tokenWord(token));
		default:
			return false;
	}
}

/** Width in columns of a line's leading whitespace, tabs advancing to the next stop. */
function indentWidth(line: string, tabSize: number): number {
	let width = 0;
	for (const ch of line) {
		if (ch === '\t') {
			width += tabSize - (width % tabSize);
		} else if (ch === ' ' || ch === '\u3000' || ch === '\u0019') {
			width += 1;
		} else {
			break;
		}
	}
	return width;
}

function indentString(width: number, tabSize: number, insertSpaces: boolean): string {
	if (width <= 0) {
		return '';
	}
	if (insertSpaces) {
		return ' '.repeat(width);
	}
	return '\t'.repeat(Math.floor(width / tabSize)) + ' '.repeat(width % tabSize);
}

/**
 * The first way two sources lex differently, or undefined when they lex to the
 * same tokens: the same kinds in the same order, the same text for every token
 * apart from letter case in keywords and identifiers, and the same number of
 * line continuations. Formatting that fails this check is thrown away.
 */
export function tokenStreamDifference(before: string, after: string): string | undefined {
	const a = tokenize(before);
	const b = tokenize(after);
	const continuations = (tokens: readonly VbaToken[]): number => {
		let n = 0;
		for (const token of tokens) {
			for (const trivia of token.leadingTrivia ?? []) {
				if (trivia.kind === 'lineContinuation') {
					n++;
				}
			}
			for (const trivia of token.trailingTrivia ?? []) {
				if (trivia.kind === 'lineContinuation') {
					n++;
				}
			}
		}
		return n;
	};
	if (a.length !== b.length) {
		return `token count changed from ${a.length} to ${b.length}`;
	}
	for (let i = 0; i < a.length; i++) {
		const x = a[i];
		const y = b[i];
		if (x.kind !== y.kind) {
			return `token ${i} changed kind from ${x.kind} to ${y.kind} at line ${x.line + 1}`;
		}
		const caseInsensitive = x.kind === 'keyword' || x.kind === 'identifier';
		const same = caseInsensitive
			? x.rawText.toLowerCase() === y.rawText.toLowerCase()
			: x.rawText === y.rawText;
		if (!same) {
			return `token ${i} changed from ${JSON.stringify(x.rawText)} to ${JSON.stringify(y.rawText)} at line ${x.line + 1}`;
		}
	}
	const ca = continuations(a);
	const cb = continuations(b);
	if (ca !== cb) {
		return `line continuation count changed from ${ca} to ${cb}`;
	}
	return undefined;
}
