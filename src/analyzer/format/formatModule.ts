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
	/** On a `#If`: an earlier arm ended inside a procedure header it opened. */
	armLeftProcedureOpen?: boolean;
	/** On a `#If`: a `#Else` or `#ElseIf` has been seen. */
	hasMidArm?: boolean;
}

type StatementClass =
	| { type: 'plain' }
	| { type: 'label' }
	| { type: 'attribute' }
	| { type: 'lineNumber'; rest: StatementClass }
	/** A one-line If, or the Else part of one; `tail` is the block its last part opens. */
	| { type: 'singleLineIf'; tail: StatementClass }
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
	// A physical line a comment ran on to through ` _` (MS-VBAL 3.3.1): all
	// comment text, kept exactly as written.
	const commentText: boolean[] = lines.map(() => false);
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
		if (token.kind === 'comment') {
			const lastLine = Math.min(lineIndexOf(lineStarts, token.end), lines.length - 1);
			for (let line = token.line + 1; line <= lastLine; line++) {
				continuation[line] = true;
				commentText[line] = true;
			}
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
	// A new procedure removes the previous one; record its only stack position.
	let procedureIndex = -1;
	const truncateStack = (length: number): void => {
		stack.length = length;
		if (procedureIndex >= length) {
			procedureIndex = -1;
		}
	};

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
		// A one-line If runs to the end of its logical line, so an `Else`
		// after a `:` there is its Else, not the enclosing block If's.
		let inSingleLineIf = false;
		statements.forEach((statement, index) => {
			const isFirst = index === 0;
			const isLast = index === statements.length - 1;
			const cls = classify(statement, isFirst, isLast, stack, inSingleLineIf);
			if (cls.type === 'singleLineIf' || (cls.type === 'lineNumber' && cls.rest.type === 'singleLineIf')) {
				inSingleLineIf = true;
			}
			lineIndent = apply(cls, isFirst, lineIndent);
		});

		const firstWidth = lineIndent * tabSize;
		const originalFirstWidth = indentWidth(lines[first], tabSize);
		for (let i = first; i <= last; i++) {
			if (commentText[i]) {
				continue;
			}
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
			case 'lineNumber':
				apply(cls.rest, false, lineIndent);
				return isFirst ? 0 : lineIndent;
			case 'singleLineIf':
				apply(cls.tail, false, lineIndent);
				return lineIndent;
			case 'closer': {
				if (cls.kind === 'PreprocessorIf') {
					return closePreprocessorIf(isFirst, lineIndent);
				}
				let indent = lineIndent;
				for (let n = 0; n < cls.count; n++) {
					const index = find(cls.kind);
					if (index < 0) {
						if (cls.kind === 'Procedure') {
							// A stray End Sub still ends whatever was open: back
							// to module level, where the next line belongs.
							truncateStack(0);
							level = 0;
							indent = 0;
						}
						break;
					}
					const target = stack[index];
					truncateStack(index);
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
				if (cls.kind === 'PreprocessorIf') {
					target.hasMidArm = true;
					if (procedureIndex > index) {
						target.armLeftProcedureOpen = true;
					}
				}
				truncateStack(index + 1);
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
				procedureIndex = stack.length;
				stack.push({ kind: 'Procedure', level: at, moduleLevel: false });
				level = at + 1;
				return isFirst ? at : lineIndent;
			}
			case 'opener': {
				const at = isFirst ? lineIndent : level;
				stack.push({ kind: cls.kind, level: at, moduleLevel: procedureIndex < 0 });
				level = at + 1;
				return lineIndent;
			}
		}
	}

	// `#If VBA7 Then` / `Sub Foo(p As LongPtr)` / `#Else` / `Sub Foo(p As
	// Long)` / `#End If`: whichever arm compiles, a procedure is open after
	// `#End If`, and its body follows (issue #1496). That holds once the `#If`
	// has more than one arm and one of them ended inside a header it opened.
	// A one-arm `#If` closes what it opened, as it always has.
	function closePreprocessorIf(isFirst: boolean, lineIndent: number): number {
		const index = find('PreprocessorIf');
		if (index < 0) {
			return lineIndent;
		}
		const target = stack[index];
		const keepProcedure = target.hasMidArm === true && (target.armLeftProcedureOpen === true || procedureIndex > index);
		truncateStack(index);
		level = target.level;
		if (keepProcedure) {
			procedureIndex = stack.length;
			stack.push({ kind: 'Procedure', level: target.level, moduleLevel: false });
			level = target.level + 1;
		}
		return isFirst ? target.level : lineIndent;
	}

	const text = output.map((line, i) => line + (eols[i] ?? '')).join('');
	const difference = tokenStreamsDifference(tokens, tokenize(text));
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
	inSingleLineIf = false,
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
		// `10 If x Then: y = 1`: the statement after the number keeps its `:`.
		const rest = { tokens: tokens.slice(1), endedByColon: statement.endedByColon };
		return { type: 'lineNumber', rest: classify(rest, false, isLast, stack, inSingleLineIf) };
	}
	if (isFirst && head.kind === 'identifier' && head.rawText.toLowerCase() === 'attribute') {
		return { type: 'attribute' };
	}
	const first = tokenWord(head);
	if (insideRecord(stack)) {
		// A Type or Enum holds only members, and `For As Long`, `Sub As Long`
		// and `Next = 1` are members whatever their names spell (issue #1496).
		const after = tokenWord(tokens[1]);
		if (after === 'as' || after === '=' || tokens[1]?.rawText === '(') {
			return { type: 'plain' };
		}
	}
	if (first === 'else' && inSingleLineIf) {
		// `If a Then b = 1: Else With e`: the one-line If's own Else.
		return { type: 'singleLineIf', tail: classifySingleLineIfTail(tokens.slice(1), stack) };
	}
	if (first === 'end') {
		const kind = END_CLOSERS[tokenWord(tokens[1])];
		return kind ? { type: 'closer', kind, count: 1 } : { type: 'plain' };
	}
	if (first === 'endif') {
		// The one-word closer MS-VBAL 5.4.2.8 allows for `End If`.
		return { type: 'closer', kind: 'If', count: 1 };
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
		// A block If ends its logical line with Then (MS-VBAL 5.4.2.8); a
		// single-line If carries its statements after Then, or after a `:`,
		// and `If x Then:` with nothing after the colon is one too (MS-VBAL
		// 5.4.2.9, issue #84). `d.Then` is a member, not the keyword.
		const lastIndex = tokens.length - 1;
		if (isLast && !statement.endedByColon && lastIndex > 0 && isThenOrElseKeyword(tokens, lastIndex) === 'then') {
			return { type: 'opener', kind: 'If' };
		}
		// `If a Then With d: .Add 1: End With` opens a block in its tail,
		// after the last Then or Else (MS-VBAL 5.4.2.9).
		let split = -1;
		for (let i = 1, depth = 0; i < tokens.length; i++) {
			const raw = tokens[i].rawText;
			if (raw === '(') {
				depth++;
			} else if (raw === ')') {
				depth--;
			} else if (depth === 0 && isThenOrElseKeyword(tokens, i)) {
				split = i;
			}
		}
		const tail = split < 0 ? [] : tokens.slice(split + 1);
		return { type: 'singleLineIf', tail: classifySingleLineIfTail(tail, stack) };
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
	if ((declared === 'type' || declared === 'enum') && !insideRecord(stack)) {
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

/** True inside a Type or Enum, conditional-compilation blocks aside. */
function insideRecord(stack: readonly OpenBlock[]): boolean {
	for (let i = stack.length - 1; i >= 0; i--) {
		if (stack[i].kind !== 'PreprocessorIf') {
			return stack[i].kind === 'Type' || stack[i].kind === 'Enum';
		}
	}
	return false;
}

/**
 * The keyword Then or Else at `index`, or undefined. After a `.`, or a `!`
 * touching it, the word is a member name: `d.Then`, `d!Else`. A `!` with a
 * space after it is a type suffix, so `x! Then` is the keyword.
 */
function isThenOrElseKeyword(tokens: readonly VbaToken[], index: number): 'then' | 'else' | undefined {
	const word = tokenWord(tokens[index]);
	if (word !== 'then' && word !== 'else') {
		return undefined;
	}
	const previous = tokens[index - 1];
	if (previous && (previous.rawText === '.' || (previous.rawText === '!' && previous.end === tokens[index].start))) {
		return undefined;
	}
	return word;
}

/**
 * The block a one-line If's tail opens. Only the loops, With and Select
 * count, with a nested one-line If's own tail; nothing else in a tail opens.
 */
function classifySingleLineIfTail(tail: VbaToken[], stack: readonly OpenBlock[]): StatementClass {
	const cls = classify({ tokens: tail, endedByColon: false }, false, false, stack);
	if (cls.type === 'singleLineIf') {
		return cls;
	}
	if (cls.type === 'opener' && cls.kind !== 'Type' && cls.kind !== 'Enum' && cls.kind !== 'If' && cls.kind !== 'PreprocessorIf') {
		return cls;
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
		if (token.kind === 'comment') {
			// Only this line's part: a comment continued with ` _` keeps its
			// later lines as they are.
			const text = token.rawText.split(/\r\n|\r|\n/)[0];
			prevEndCol = token.character + text.length;
			out += gap + text;
			continue;
		}
		prevEndCol = token.character + token.rawText.length;
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
	return out + trailing(line, prevEndCol);
}

/** The text after the last token: nothing, or the ` _` of a continuation. */
function trailing(line: string, from: number): string {
	const rest = line.slice(from).replace(/[ \t\u3000]+$/, '');
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
	// `=>` is `>=` written the other way round, and spaced like it.
	if (cur.kind === 'operator' && COMPARISON_OR_ASSIGN.has(cur.canonicalText ?? curRaw)) {
		return { spaceBefore: true, binary: false, suffix: false };
	}
	if (prev.kind === 'operator' && COMPARISON_OR_ASSIGN.has(prev.canonicalText ?? prev.rawText)) {
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
 * apart from letter case in keywords and identifiers, and the same line
 * continuations in the same places. Formatting that fails this check is thrown
 * away.
 */
export function tokenStreamDifference(before: string, after: string): string | undefined {
	return tokenStreamsDifference(tokenize(before), tokenize(after));
}

function tokenStreamsDifference(a: readonly VbaToken[], b: readonly VbaToken[]): string | undefined {
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
	// The same number in different places is a different module too: a ` _`
	// that loses its space is a stray `_` token, and one that gains a space is a
	// continuation, so a formatted module could trade one for the other and
	// keep the count (found by tests/properties).
	for (let i = 0; i < a.length; i++) {
		if (lineContinuationCount(a[i].leadingTrivia) !== lineContinuationCount(b[i].leadingTrivia)
			|| lineContinuationCount(a[i].trailingTrivia) !== lineContinuationCount(b[i].trailingTrivia)) {
			return `a line continuation moved at token ${i}, line ${a[i].line + 1}`;
		}
	}
	return undefined;
}

/** Counts continuation trivia without allocating per-token arrays or strings. */
function lineContinuationCount(trivia: VbaToken['leadingTrivia']): number {
	let count = 0;
	if (trivia) {
		for (const part of trivia) {
			if (part.kind === 'lineContinuation') { count++; }
		}
	}
	return count;
}
