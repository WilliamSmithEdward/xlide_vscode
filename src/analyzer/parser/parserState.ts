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
}

/**
 * Split a token stream into logical statements (MS-VBAL 3.3.1 EOS). 'newline'
 * and 'colon' tokens act as separators and are not included in any statement.
 * Empty statements (blank lines, doubled separators) are dropped.
 */
export function splitLogicalStatements(tokens: readonly VbaToken[]): LogicalStatement[] {
	const statements: LogicalStatement[] = [];
	let current: VbaToken[] = [];

	const flush = () => {
		if (current.length === 0) {
			return;
		}
		const first = current[0];
		const last = current[current.length - 1];
		statements.push({
			tokens: current,
			start: first.start,
			end: last.end,
			line: first.line,
		});
		current = [];
	};

	for (const token of tokens) {
		if (token.kind === 'newline' || token.kind === 'colon') {
			flush();
			continue;
		}
		current.push(token);
	}
	flush();
	return statements;
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

/**
 * Canonical (case-folded) text of a token used for keyword matching. For
 * keyword tokens the lexer already provides canonicalText; otherwise we lower
 * the raw text because VBA is case-insensitive (MS-VBAL 3.3.5).
 */
export function tokenWord(token: VbaToken | undefined): string {
	if (!token) {
		return '';
	}
	return (token.canonicalText ?? token.rawText).toLowerCase();
}
