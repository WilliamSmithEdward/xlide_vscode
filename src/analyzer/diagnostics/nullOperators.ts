// Whether an expression gives Null through its operators (issues #324 and
// #556): arithmetic and comparisons propagate Null; And, Or and Imp can be
// decided by the other operand. Concatenation is not inferred as Null.
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenWord as tokenText } from '../lexer/tokenHelpers';

const NULL_PROPAGATING: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '=', '<>', '<', '>', '<=', '>=', 'and', 'or', 'xor', 'eqv', 'imp', '&']);
type TokenRange = { start: number; end: number };
type NullLiteral = number | 'null' | undefined;
interface NullFrame {
	ranges: TokenRange[];
	mode: 'some' | 'every' | 'and' | 'or' | 'imp';
	next: number;
	left?: NullLiteral;
}

/** Whether tokens give Null, querying single-token values through holdsNull. */
export function operatorYieldsNull(toks: readonly VbaToken[], holdsNull: (tok: VbaToken) => boolean): boolean {
	// Common leaf and single-prefix operands need no continuation setup.
	if (toks.length === 1) { return holdsNull(toks[0]); }
	if (toks.length === 2 && (tokenText(toks[0]) === 'not' || tokenText(toks[0]) === '-')) {
		return holdsNull(toks[1]);
	}
	// Match once, only when parentheses are encountered. All windows share the
	// immutable input; nested wrappers neither copy tokens nor grow the stack.
	let parens: Map<number, number> | undefined;
	const matchingParen = (open: number): number | undefined => {
		if (!parens) {
			parens = new Map();
			const pending: number[] = [];
			for (let i = 0; i < toks.length; i++) {
				if (toks[i].rawText === '(') { pending.push(i); }
				else if (toks[i].rawText === ')') {
					const start = pending.pop();
					if (start !== undefined) { parens.set(start, i); }
				}
			}
		}
		return parens.get(open);
	};
	// Keep the existing single outer-pair normalization and literal coercions.
	const isOuterPair = (start: number, end: number): boolean =>
		end - start >= 2 && toks[start].rawText === '(' && matchingParen(start) === end - 1;
	const unwrap = (range: TokenRange): TokenRange =>
		isOuterPair(range.start, range.end)
			? { start: range.start + 1, end: range.end - 1 } : range;
	const literalNumber = (range: TokenRange): number | undefined => {
		const { start, end } = unwrap(range);
		const length = end - start;
		const sign = length === 2 && toks[start].rawText === '-' ? -1 : 1;
		const tok = length === 1 ? toks[start] : length === 2 && (toks[start].rawText === '-' || toks[start].rawText === '+') ? toks[start + 1] : undefined;
		const word = tokenText(tok);
		if (word === 'true' || word === 'false') { return sign * (word === 'true' ? -1 : 0); }
		if (tok?.kind !== 'integerLiteral' && tok?.kind !== 'floatLiteral') { return undefined; }
		const value = Number(tok.rawText.replace(/[%&^!#@]$/, ''));
		return Number.isFinite(value) ? sign * value : undefined;
	};
	const frames: NullFrame[] = [];
	let current: TokenRange = { start: 0, end: toks.length };
	let result: boolean | undefined;
	for (;;) {
		if (result === undefined) {
			let { start, end } = current;
			for (;;) {
				if (isOuterPair(start, end)) { start++; end--; }
				if (end - start === 1) { result = holdsNull(toks[start]); break; }
				if (end <= start) { result = false; break; }
				const head = tokenText(toks[start]);
				if (head === '-' || head === 'not') { start++; continue; }
				if (head === 'abs' && toks[start + 1]?.rawText === '(' && matchingParen(start + 1) === end - 1) {
					start += 2; end--; continue;
				}
				break;
			}
			if (result === undefined) {
				const ranges: TokenRange[] = [];
				const operators: string[] = [];
				let segmentStart = start;
				let depth = 0;
				for (let i = start; i < end; i++) {
					const tok = toks[i];
					if (depth === 0 && tok.rawText === '(') {
						const close = matchingParen(i);
						if (close !== undefined && close < end) { i = close; continue; }
					}
					depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
					const word = tok.kind === 'operator' ? tok.rawText : tokenText(tok);
					if (depth === 0 && i > segmentStart && NULL_PROPAGATING.has(word) && tok.kind !== 'stringLiteral') {
						ranges.push({ start: segmentStart, end: i });
						operators.push(word); segmentStart = i + 1;
					}
				}
				ranges.push({ start: segmentStart, end });
				if (operators.length === 0 || operators.includes('&') || ranges.some(range => range.start === range.end)) {
					result = false;
				} else {
					const logical = operators.find((op): op is 'and' | 'or' | 'imp' => op === 'and' || op === 'or' || op === 'imp');
					frames.push({ ranges, mode: logical ? (operators.length === 1 ? logical : 'every') : 'some', next: 1 });
					current = ranges[0];
					continue;
				}
			}
		}
		// Resume one child at a time to retain callback order and short-circuiting.
		for (;;) {
			const frame = frames[frames.length - 1];
			if (!frame) { return result; }
			if (frame.mode === 'some' || frame.mode === 'every') {
				if ((frame.mode === 'some' ? result : !result) || frame.next === frame.ranges.length) {
					frames.pop(); continue;
				}
				current = frame.ranges[frame.next++]; result = undefined; break;
			}
			const value: NullLiteral = result ? 'null' : literalNumber(frame.ranges[frame.next - 1]);
			if (frame.next === 1) {
				frame.left = value; frame.next = 2; current = frame.ranges[1]; result = undefined; break;
			}
			const left = frame.left;
			const decided = (other: NullLiteral, otherOnLeft: boolean): boolean => other === 'null'
				|| (other !== undefined && (frame.mode === 'and' ? other !== 0 : frame.mode === 'or' ? other === 0 : otherOnLeft ? other !== 0 : other === 0));
			result = (left === 'null' && decided(value, false)) || (value === 'null' && decided(left, true));
			frames.pop();
		}
	}
}
