// Parser state: turns the lexer's flat token stream into logical statements and
// provides a small cursor used by the parser.
//
// Verified against MS-VBAL.pdf, v20250520:
//   - 3.3.1 EOS = *(EOL / ":") -- a logical line ends at a line-terminator or a
//     ':' statement separator. We therefore split the token stream on 'newline'
//     and 'colon' tokens.
//   - 3.2.2 line-continuation already merged into trivia by the lexer, so a
//     continued physical line is a single logical statement here.
//
// Recovery model (Phase 3): statement boundaries are the natural recovery
// points required by the roadmap ("recover at newline boundaries / at colon
// statement separators"), so working at this granularity makes the parser
// inherently error-tolerant.

import { tokenWord } from '../lexer/tokenHelpers';
import { VbaToken } from '../lexer/tokenKinds';

/** A logical statement: the significant tokens between two separators. */
export interface LogicalStatement {
	/** Significant tokens (no 'newline'/'colon' separators; trailing comment kept). */
	tokens: VbaToken[];
	/** Absolute offset of the first token. */
	start: number;
	/** Absolute offset just past the last token. */
	end: number;
	/** Zero-based line of the first token. */
	line: number;
	/** True when a ':' separator, not the end of the line, closed the statement. */
	endedByColon?: boolean;
	/**
	 * True when a single-line If earlier on the same line runs this statement
	 * after a colon, as `b` in `If x Then a: b` (MS-VBAL 5.4.2.9).
	 */
	singleLineIfTail?: boolean;
}

/**
 * Split a token stream into logical statements (MS-VBAL 3.3.1 EOS). 'newline'
 * and 'colon' tokens act as separators and are not included in any statement.
 * Empty statements (blank lines, doubled separators) are dropped.
 */
export function splitLogicalStatements(tokens: readonly VbaToken[]): LogicalStatement[] {
	const statements: LogicalStatement[] = [];
	let current: VbaToken[] = [];
	// A single-line If earlier on this line runs every statement after it to
	// the end of the line: `b` and `c` in `If x Then a: b: c`, and `c` in
	// `If x Then a Else b: c`, all sit in its Then or Else list (MS-VBAL
	// 5.4.2.9). `If x Then:` opens such a list too.
	let inSingleLineIf = false;

	const flush = (endedByColon: boolean) => {
		if (current.length === 0) {
			return;
		}
		const first = current[0];
		const last = current[current.length - 1];
		const statement: LogicalStatement = {
			tokens: current,
			start: first.start,
			end: last.end,
			line: first.line,
			...(endedByColon ? { endedByColon } : {}),
			...(inSingleLineIf ? { singleLineIfTail: true } : {}),
		};
		statements.push(statement);
		inSingleLineIf = inSingleLineIf || (endedByColon && isSingleLineIf(statement));
		current = [];
	};

	for (const token of tokens) {
		if (token.kind === 'newline' || token.kind === 'colon') {
			splitTailBlockOpener();
			flush(token.kind === 'colon');
			if (token.kind === 'newline') {
				inSingleLineIf = false;
			}
			continue;
		}
		current.push(token);
	}
	splitTailBlockOpener();
	flush(false);
	return statements;

	// `If a Then With c: .Add 1: End With` (MS-VBAL 5.4.2.9; issue #128): the
	// block opener in a one-line If's Then or Else tail becomes its own
	// statement, so the block parser opens it and the closers on the same line
	// find it. The If header keeps `endedByColon`, which stops it reading as
	// a block If, and the opener statement is a tail of it.
	function splitTailBlockOpener(): void {
		const split = singleLineIfTailOpenerIndex(current);
		if (split < 0) {
			return;
		}
		const header = current.slice(0, split);
		const tail = current.slice(split);
		current = header;
		flush(true);
		inSingleLineIf = true;
		current = tail;
	}
}

const TAIL_BLOCK_OPENERS: ReadonlySet<string> = new Set(['with', 'for', 'do', 'while', 'select']);

/** Index of the block opener a one-line If's Then/Else tail starts with, or -1. */
function singleLineIfTailOpenerIndex(tokens: readonly VbaToken[]): number {
	const head = tokens[0]?.kind === 'integerLiteral' && /^\d+$/.test(tokens[0].rawText) ? 1 : 0;
	if (tokenWord(tokens[head]) !== 'if') {
		return -1;
	}
	let depth = 0;
	let last = -1;
	for (let i = head + 1; i < tokens.length; i++) {
		const raw = tokens[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (depth === 0) {
			const word = tokenWord(tokens[i]);
			if (word === 'then' || word === 'else') {
				last = i;
			}
		}
	}
	if (last < 0 || last === tokens.length - 1) {
		return -1;
	}
	const opener = tokenWord(tokens[last + 1]);
	if (!opener || !TAIL_BLOCK_OPENERS.has(opener)) {
		return -1;
	}
	if (opener === 'select' && tokenWord(tokens[last + 2]) !== 'case') {
		return -1;
	}
	return last + 1;
}

/** An If statement that is not a block If header, whose Then ends its line. */
function isSingleLineIf(statement: LogicalStatement): boolean {
	const tokens = codeTokens(statement);
	const head = tokens[0]?.kind === 'integerLiteral' && /^\d+$/.test(tokens[0].rawText) ? 1 : 0;
	if (tokenWord(tokens[head]) !== 'if') {
		return false;
	}
	return tokenWord(tokens[tokens.length - 1]) !== 'then' || statement.endedByColon === true;
}

/**
 * A forward cursor over a list of logical statements. The parser consumes
 * statements one at a time; block parsers peek ahead to find their closers.
 */
export class StatementCursor {
	private index = 0;

	constructor(private readonly statements: readonly LogicalStatement[]) {}

	/** True when no statements remain. */
	atEnd(): boolean {
		return this.index >= this.statements.length;
	}

	/** The current statement without consuming it, or undefined at end. */
	peek(): LogicalStatement | undefined {
		return this.statements[this.index];
	}

	/** Consume and return the current statement, or undefined at end. */
	next(): LogicalStatement | undefined {
		return this.statements[this.index++];
	}

	/** Current cursor position (for span bookkeeping). */
	position(): number {
		return this.index;
	}
}

/**
 * Significant tokens of a statement excluding any trailing comment, which is
 * never syntactically meaningful (MS-VBAL 3.3.1 comment-body).
 */
export function codeTokens(statement: LogicalStatement): VbaToken[] {
	return statement.tokens.filter((t) => t.kind !== 'comment');
}

export { tokenWord } from '../lexer/tokenHelpers';
