// Whether an expression gives Null through its operators (issues #324 and
// #556, each measured in Excel 16.0): arithmetic, `+`, unary minus, Not, a
// comparison and Abs give Null when an operand is Null, and so do Xor and
// Eqv; And, Or and Imp give a value when the other side decides it; `&`
// never does.

import type { VbaToken } from '../lexer/tokenKinds';
import { matchParenFrom, tokenText } from './walker';
import { unwrapOuterParens } from './typeInference';

/** The binary operators whose result is Null when an operand is. `&` is here to be refused. */
const NULL_PROPAGATING: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '=', '<>', '<', '>', '<=', '>=', 'and', 'or', 'xor', 'eqv', 'imp', '&']);

/** The number a literal operand is, True as -1 and False as 0: `1`, `-2.5`, `True`. */
function literalNumber(operand: readonly VbaToken[]): number | undefined {
	const toks = unwrapOuterParens([...operand]);
	const sign = toks.length === 2 && toks[0].rawText === '-' ? -1 : 1;
	const tok = toks.length === 1 ? toks[0] : toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+') ? toks[1] : undefined;
	const word = tokenText(tok);
	if (word === 'true' || word === 'false') {
		return sign * (word === 'true' ? -1 : 0);
	}
	if (tok?.kind !== 'integerLiteral' && tok?.kind !== 'floatLiteral') {
		return undefined;
	}
	const value = Number(tok.rawText.replace(/[%&^!#@]$/, ''));
	return Number.isFinite(value) ? sign * value : undefined;
}

/**
 * Whether the tokens give Null. `holdsNull` says whether one token does: the
 * literal Null, or a local known to hold it. A single token is asked whole.
 */
export function operatorYieldsNull(toks: readonly VbaToken[], holdsNull: (tok: VbaToken) => boolean): boolean {
	const part = unwrapOuterParens([...toks]);
	if (part.length === 1) {
		return holdsNull(part[0]);
	}
	const head = tokenText(part[0]);
	if (head === '-' || head === 'not') {
		return operatorYieldsNull(part.slice(1), holdsNull);
	}
	if (head === 'abs' && part[1]?.rawText === '(' && matchParenFrom(part, 1) === part.length - 1) {
		return operatorYieldsNull(part.slice(2, -1), holdsNull);
	}
	const operands: VbaToken[][] = [[]];
	const operators: string[] = [];
	let depth = 0;
	for (const tok of part) {
		depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
		const word = tok.kind === 'operator' ? tok.rawText : tokenText(tok);
		const current = operands[operands.length - 1];
		if (depth === 0 && current.length > 0 && NULL_PROPAGATING.has(word) && tok.kind !== 'stringLiteral') {
			operators.push(word);
			operands.push([]);
		} else {
			current.push(tok);
		}
	}
	if (operators.length === 0 || operators.includes('&') || operands.some((operand) => operand.length === 0)) {
		return false;
	}
	// And, Or and Imp give a value when the other side decides it (issue
	// #556): Null And 0 is 0, but Null And 1 is Null; 40000 Or Null is 40000,
	// but 0 Or Null is Null; Null Imp 12 is 12 and False Imp Null is True, but
	// Null Imp False is Null. Judged with one operator only.
	const logical = operators.find((operator) => operator === 'and' || operator === 'or' || operator === 'imp');
	if (logical) {
		if (operators.length !== 1) {
			return operands.every((operand) => operatorYieldsNull(operand, holdsNull));
		}
		const [left, right] = operands.map((operand) => (operatorYieldsNull(operand, holdsNull) ? 'null' : literalNumber(operand)));
		const decided = (other: number | 'null' | undefined, otherOnLeft: boolean): boolean => other === 'null'
			|| (other !== undefined && (logical === 'and' ? other !== 0 : logical === 'or' ? other === 0 : otherOnLeft ? other !== 0 : other === 0));
		return (left === 'null' && decided(right, false)) || (right === 'null' && decided(left, true));
	}
	return operands.some((operand) => operatorYieldsNull(operand, holdsNull));
}
