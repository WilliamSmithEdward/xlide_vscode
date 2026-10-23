// The expressions inside statements that are not assignments or calls.
//
// The shared expression walk (exprWalk.ts) parses assignments, calls and If
// conditions. A Case item, a Print # or Debug.Print item, a file number, a
// Seek position, a record number or a Lock range is an expression too, and
// the VBE judges it as one: an array or a scalar `Is` there is the same Type
// mismatch it is anywhere (issue #1415, measured in Excel, Word, PowerPoint
// and Access 16.0). This module finds those expressions in a statement's own
// text, so every expression rule sees them through the one walk. The arms of
// a single-line If are statements of their own and read the same way; its
// condition is left to the rules that already read it.
//
// Only operand positions are read. A Put's or an Input's variable may be an
// array, and an Open's path and access words are left alone.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenWord, tokensWithoutLeadingLineNumber } from '../lexer/tokenHelpers';
import type { ExprNode, StatementNode } from '../parser/nodes';
import { parseExpression } from '../parser/parseExpression';

const ROOTS = new WeakMap<StatementNode, { raw: string; roots: ExprNode[] }>();

/** The Case items, each compared with the Select value as by `=` or its `Is` operator. */
const CASE_COMPARANDS = new WeakMap<ExprNode, string>();

/** The operator a Case item is compared by, when `expr` is one: `=`, or `Is >`'s `>`. */
export function caseComparisonOperator(expr: ExprNode): string | undefined {
	return CASE_COMPARANDS.get(expr);
}

/** The operand expressions of a statement, each a root for the expression walk. */
export function statementOperandRoots(node: StatementNode): readonly ExprNode[] {
	const cached = ROOTS.get(node);
	if (cached && cached.raw === node.raw) {
		return cached.roots;
	}
	const toks = tokenize(node.raw)
		.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline')
		.map((tok) => ({ ...tok, start: tok.start + node.span.start, end: tok.end + node.span.start }));
	const roots: ExprNode[] = [];
	const add = (offset: number) => ([from, to, comparison]: OperandRange): void => {
		const expr = parseExpression(toks, offset + from, offset + to).expr;
		if (expr) {
			roots.push(expr);
			if (comparison) {
				CASE_COMPARANDS.set(expr, comparison);
			}
		}
	};
	if (node.singleLineIfBranches?.length) {
		for (const arm of node.singleLineIfBranches) {
			const inArm = toks.filter((tok) => tok.start >= arm.start && tok.end <= arm.end);
			if (inArm.length > 0) {
				operandRanges(inArm).forEach(add(toks.indexOf(inArm[0])));
			}
		}
	} else {
		const body = tokensWithoutLeadingLineNumber(toks);
		operandRanges(body).forEach(add(toks.length - body.length));
	}
	ROOTS.set(node, { raw: node.raw, roots });
	return roots;
}

/** Indices of the top-level tokens in [from, to) that `stop` matches. */
function topLevel(toks: readonly VbaToken[], from: number, to: number, stop: (tok: VbaToken) => boolean): number[] {
	const out: number[] = [];
	let depth = 0;
	for (let i = from; i < to; i++) {
		const raw = toks[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (depth === 0 && stop(toks[i])) {
			out.push(i);
		}
	}
	return out;
}

/** [from, to) of the pieces between top-level separators. */
function pieces(toks: readonly VbaToken[], from: number, to: number, separators: ReadonlySet<string>): Array<[number, number]> {
	const out: Array<[number, number]> = [];
	let start = from;
	for (const at of topLevel(toks, from, to, (tok) => separators.has(tok.rawText))) {
		out.push([start, at]);
		start = at + 1;
	}
	out.push([start, to]);
	return out;
}

const COMMA: ReadonlySet<string> = new Set([',']);
const PRINT_SEPARATORS: ReadonlySet<string> = new Set([',', ';']);
const RELATIONAL: ReadonlySet<string> = new Set(['=', '<>', '<', '>', '<=', '>=']);

/** `#n` or `n` at `from`: the number's range, without its `#`. */
function fileNumber(toks: readonly VbaToken[], from: number, to: number): [number, number] {
	return toks[from]?.rawText === '#' ? [from + 1, to] : [from, to];
}

/** [from, to) of an operand, and for a Case item the operator it is compared by. */
type OperandRange = [number, number, string?];

/** The operand ranges of one statement's tokens, by its leading words. */
function operandRanges(toks: readonly VbaToken[]): OperandRange[] {
	const head = tokenWord(toks[0]);
	const end = toks.length;
	const out: OperandRange[] = [];
	const items = (from: number, separators: ReadonlySet<string>): void => {
		for (const [start, stop] of pieces(toks, from, end, separators)) {
			// `Tab(n)` and `Spc(n)` take their argument; a bare Tab is a column.
			const word = tokenWord(toks[start]);
			if ((word === 'tab' || word === 'spc') && toks[start + 1]?.rawText === '(' && toks[stop - 1]?.rawText === ')') {
				out.push([start + 2, stop - 1]);
			} else if (start < stop) {
				out.push([start, stop]);
			}
		}
	};
	switch (head) {
		case 'case': {
			if (tokenWord(toks[1]) === 'else') {
				break;
			}
			for (const [start, stop] of pieces(toks, 1, end, COMMA)) {
				if (tokenWord(toks[start]) === 'is' && RELATIONAL.has(toks[start + 1]?.rawText ?? '')) {
					out.push([start + 2, stop, toks[start + 1].rawText]);
					continue;
				}
				const to = topLevel(toks, start, stop, (tok) => tokenWord(tok) === 'to')[0];
				if (to === undefined) {
					out.push([start, stop, '=']);
				} else {
					out.push([start, to, '>='], [to + 1, stop, '<=']);
				}
			}
			break;
		}
		case 'print':
		case 'write': {
			if (toks[1]?.rawText !== '#') {
				break;
			}
			const [first, ...rest] = pieces(toks, 1, end, COMMA);
			out.push(fileNumber(toks, first[0], first[1]));
			if (rest.length > 0) {
				items(rest[0][0], head === 'print' ? PRINT_SEPARATORS : COMMA);
			}
			break;
		}
		case 'debug': {
			if (toks[1]?.rawText === '.' && tokenWord(toks[2]) === 'print') {
				items(3, PRINT_SEPARATORS);
			}
			break;
		}
		case 'seek': {
			for (const [start, stop] of pieces(toks, 1, end, COMMA)) {
				out.push(fileNumber(toks, start, stop));
			}
			break;
		}
		case 'put':
		case 'get': {
			// The number and the record; the variable may be an array.
			const [number, record] = pieces(toks, 1, end, COMMA);
			out.push(fileNumber(toks, number[0], number[1]));
			if (record) {
				out.push(record);
			}
			break;
		}
		case 'line':
		case 'input': {
			const at = head === 'line' ? (tokenWord(toks[1]) === 'input' ? 2 : -1) : 1;
			if (at > 0 && toks[at]?.rawText === '#') {
				const [number] = pieces(toks, at, end, COMMA);
				out.push(fileNumber(toks, number[0], number[1]));
			}
			break;
		}
		case 'close': {
			for (const [start, stop] of pieces(toks, 1, end, COMMA)) {
				out.push(fileNumber(toks, start, stop));
			}
			break;
		}
		case 'lock':
		case 'unlock': {
			const [number, range] = pieces(toks, 1, end, COMMA);
			out.push(fileNumber(toks, number[0], number[1]));
			if (range) {
				const to = topLevel(toks, range[0], range[1], (tok) => tokenWord(tok) === 'to')[0];
				if (to === undefined) {
					out.push(range);
				} else {
					out.push([range[0], to], [to + 1, range[1]]);
				}
			}
			break;
		}
		case 'open': {
			const as = topLevel(toks, 1, end, (tok) => tokenWord(tok) === 'as').pop();
			if (as !== undefined) {
				const len = topLevel(toks, as + 1, end, (tok) => tokenWord(tok) === 'len')[0] ?? end;
				out.push(fileNumber(toks, as + 1, len));
			}
			break;
		}
		default:
			break;
	}
	return out.filter(([from, to]) => from < to);
}
