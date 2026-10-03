// What an If condition evaluates to when the code makes it plain (issue #273).
//
// A guard whose outcome is known decides which arm runs: with c never Set,
// `If c Is Nothing Then Exit Function` always leaves, so nothing after it
// runs; with n = 0, `If n > 0 Then` never runs its arm. The rules that track
// values and object state ask this module, and skip what cannot run.
//
// Only what is certain is answered: a comparison of two known numbers, two
// strings that differ or match under any Option Compare, `Is Nothing` on an
// object whose state is known, IsNumeric and Len of a known string, and
// Not, And and Or over those. Anything else is undefined.

import type { VbaToken } from '../lexer/tokenKinds';
import { tokenName, tokenWord as tokenText } from '../lexer/tokenHelpers';
import { numericStringVerdict } from './stringConversion';
import { parseVbaIntegerLiteral } from '../constants/integerConstantExpression';
/** A string literal's text, its doubled quotes undone. */
function stringLiteralValue(raw: string): string {
	return raw.replace(/^"/, '').replace(/"$/, '').replace(/""/g, '"');
}

export interface ConditionFacts {
	/** The number or string a name is known to hold, by lowercased name. */
	value(lower: string): number | string | undefined;
	/** Whether an object name is known Nothing (true) or known set (false). */
	isNothing?(lower: string): boolean | undefined;
	/** The whole numbers a name is known to lie between, both included: `Second(Now) + 1000` (issue #565). */
	range?(lower: string): readonly [number, number] | undefined;
	/** Whether a name is known to hold Null (issue #664). */
	isNull?(lower: string): boolean | undefined;
}

/** What a name holds when only its range is known. */
type Range = { range: readonly [number, number] };
type Operand = Value | { nothing: boolean | undefined } | Range;

type Value = number | string | boolean | undefined;

/** The value of `toks` as a condition: true, false, or undefined when it is not certain. */
export function conditionValue(toks: readonly VbaToken[], facts: ConditionFacts): boolean | undefined {
	const parser = new ConditionParser(toks.filter((tok) => tok.kind !== 'comment'), facts);
	const value = parser.orExpr();
	if (!parser.done()) {
		return undefined;
	}
	return truth(value);
}

/** The tokens between `If` (or `ElseIf`) and `Then` in a statement or header. */
export function ifConditionTokens(toks: readonly VbaToken[]): VbaToken[] | undefined {
	const head = tokenText(toks[0]);
	if (head !== 'if' && head !== 'elseif') {
		return undefined;
	}
	let depth = 0;
	for (let i = 1; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (depth === 0 && tokenText(toks[i]) === 'then') {
			return toks.slice(1, i);
		}
	}
	return undefined;
}

function truth(value: Value): boolean | undefined {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	return undefined;
}

class ConditionParser {
	private index = 0;

	constructor(private readonly toks: readonly VbaToken[], private readonly facts: ConditionFacts) {}

	done(): boolean {
		return this.index >= this.toks.length;
	}

	private word(): string {
		return tokenText(this.toks[this.index]);
	}

	orExpr(): Value {
		let value = this.andExpr();
		while (this.word() === 'or') {
			this.index++;
			const right = this.andExpr();
			const a = truth(value);
			const b = truth(right);
			value = a === true || b === true ? true : a === false && b === false ? false : undefined;
		}
		return value;
	}

	private andExpr(): Value {
		let value = this.notExpr();
		while (this.word() === 'and') {
			this.index++;
			const right = this.notExpr();
			const a = truth(value);
			const b = truth(right);
			value = a === false || b === false ? false : a === true && b === true ? true : undefined;
		}
		return value;
	}

	private notExpr(): Value {
		if (this.word() === 'not') {
			this.index++;
			const value = truth(this.notExpr());
			return value === undefined ? undefined : !value;
		}
		return this.comparison();
	}

	private comparison(): Value {
		const left = this.sum();
		const op = this.toks[this.index]?.rawText;
		if (this.word() === 'is' && tokenText(this.toks[this.index + 1]) === 'nothing') {
			this.index += 2;
			return typeof left === 'object' && 'nothing' in left ? left.nothing : undefined;
		}
		if (op !== '=' && op !== '<>' && op !== '<' && op !== '>' && op !== '<=' && op !== '>=') {
			return typeof left === 'object' ? undefined : left;
		}
		this.index++;
		const right = this.sum();
		// A range against a number: decided where every value in it agrees.
		const [lo, hi] = rangeOf(left) ?? [];
		const [rlo, rhi] = rangeOf(right) ?? [];
		if ((typeof left === 'object' && 'range' in left) || (typeof right === 'object' && 'range' in right)) {
			if (lo === undefined || rlo === undefined) {
				return undefined;
			}
			const corners = [compare(lo, rlo, op), compare(lo, rhi!, op), compare(hi!, rlo, op), compare(hi!, rhi!, op)];
			// `=` and `<>` hold at an inner value the corners miss.
			if ((op === '=' || op === '<>') && !(hi! < rlo || rhi! < lo) && !(lo === hi && rlo === rhi)) {
				return undefined;
			}
			return corners.every((corner) => corner) ? true : corners.every((corner) => !corner) ? false : undefined;
		}
		if (typeof left === 'object' || typeof right === 'object' || left === undefined || right === undefined) {
			return undefined;
		}
		if (typeof left === 'number' && typeof right === 'number') {
			return compare(left, right, op);
		}
		if (typeof left === 'string' && typeof right === 'string' && (op === '=' || op === '<>')) {
			// Equal both ways or different both ways, whatever Option Compare says.
			if (left === right) {
				return op === '=';
			}
			if (left.toLowerCase() !== right.toLowerCase()) {
				return op === '<>';
			}
		}
		return undefined;
	}

	/**
	 * Whole numbers joined by +, -, * or Mod: `b Mod 2 = 0` with b known
	 * (issue #565). A result past the Long range, which may overflow, and any
	 * operand that is not a known whole number make it undefined.
	 */
	private sum(): Operand {
		let value = this.product();
		for (let op = this.toks[this.index]?.rawText; op === '+' || op === '-'; op = this.toks[this.index]?.rawText) {
			this.index++;
			value = wholeArithmetic(value, this.product(), op);
		}
		return value;
	}

	private product(): Operand {
		let value = this.operand();
		for (let op = this.toks[this.index]?.rawText; op === '*' || this.word() === 'mod'; op = this.toks[this.index]?.rawText) {
			const mod = this.word() === 'mod';
			this.index++;
			value = wholeArithmetic(value, this.operand(), mod ? 'mod' : '*');
		}
		return value;
	}

	/** A literal, a known name, an object name (for Is Nothing), IsNumeric(...), or a parenthesized condition. */
	private operand(): Operand {
		const tok = this.toks[this.index];
		if (!tok) {
			return undefined;
		}
		if (tok.rawText === '(') {
			this.index++;
			const value = this.orExpr();
			if (this.toks[this.index]?.rawText !== ')') {
				this.index = this.toks.length + 1;
				return undefined;
			}
			this.index++;
			return value;
		}
		if (tok.rawText === '-' && this.toks[this.index + 1]?.kind === 'integerLiteral') {
			this.index += 2;
			const value = parseVbaIntegerLiteral(this.toks[this.index - 1].rawText);
			return value === undefined ? undefined : -value;
		}
		this.index++;
		if (tok.kind === 'integerLiteral') {
			return parseVbaIntegerLiteral(tok.rawText);
		}
		if (tok.kind === 'stringLiteral') {
			return stringLiteralValue(tok.rawText);
		}
		const word = tokenText(tok);
		if (word === 'true' || word === 'false') {
			return word === 'true';
		}
		// IsNull of a local a straight line set to Null, or to a number or a
		// string (issue #664).
		if (word === 'isnull' && this.toks[this.index]?.rawText === '(' && this.toks[this.index + 2]?.rawText === ')') {
			const arg = this.toks[this.index + 1];
			this.index += 3;
			if (tokenText(arg) === 'null') {
				return true;
			}
			const lower = tokenName(arg)?.toLowerCase() ?? '';
			const held = this.facts.isNull?.(lower);
			if (held !== undefined) {
				return held;
			}
			return this.facts.value(lower) !== undefined ? false : undefined;
		}
		if (word === 'isnumeric' && this.toks[this.index]?.rawText === '(' && this.toks[this.index + 2]?.rawText === ')') {
			const arg = this.toks[this.index + 1];
			this.index += 3;
			const value = arg.kind === 'stringLiteral' ? stringLiteralValue(arg.rawText) : this.facts.value(tokenName(arg)?.toLowerCase() ?? '');
			if (typeof value === 'number') {
				return true;
			}
			if (typeof value !== 'string') {
				return undefined;
			}
			const verdict = numericStringVerdict(value);
			return verdict.kind === 'invalid' ? false : verdict.value !== undefined ? true : undefined;
		}
		// Len or LenB of a known String: `If Len(s) > 1 Then` with s empty (issue #577).
		if ((word === 'len' || word === 'lenb') && this.toks[this.index]?.rawText === '(' && this.toks[this.index + 2]?.rawText === ')') {
			const arg = this.toks[this.index + 1];
			this.index += 3;
			const value = arg.kind === 'stringLiteral' ? stringLiteralValue(arg.rawText) : this.facts.value(tokenName(arg)?.toLowerCase() ?? '');
			return typeof value === 'string' ? value.length * (word === 'lenb' ? 2 : 1) : undefined;
		}
		const lower = tokenName(tok)?.toLowerCase();
		if (!lower || this.toks[this.index]?.rawText === '(' || this.toks[this.index]?.rawText === '.') {
			this.index = this.toks.length + 1;
			return undefined;
		}
		if (this.word() === 'is') {
			return { nothing: this.facts.isNothing?.(lower) };
		}
		const known = this.facts.value(lower);
		const range = known === undefined ? this.facts.range?.(lower) : undefined;
		return range ? { range } : known;
	}
}

/** The range an operand lies in: a number is a range of one. */
function rangeOf(operand: Operand): readonly [number, number] | undefined {
	return typeof operand === 'number' && Number.isInteger(operand) ? [operand, operand]
		: typeof operand === 'object' && operand !== null && 'range' in operand ? operand.range
		: undefined;
}

function wholeArithmetic(left: Operand, right: Operand, op: string): Value | Range {
	// A range moves under + and -: `Second(Now) + 1000` lies in 1000 to 1059.
	if ((op === '+' || op === '-') && ((typeof left === 'object' && 'range' in left) || (typeof right === 'object' && 'range' in right))) {
		const a = rangeOf(left);
		const b = rangeOf(right);
		return a && b ? { range: op === '+' ? [a[0] + b[0], a[1] + b[1]] : [a[0] - b[1], a[1] - b[0]] } : undefined;
	}
	if (typeof left !== 'number' || typeof right !== 'number' || !Number.isInteger(left) || !Number.isInteger(right)) {
		return undefined;
	}
	if (op === 'mod' && right === 0) {
		return undefined;
	}
	const value = op === '+' ? left + right : op === '-' ? left - right : op === '*' ? left * right : left % right;
	return Math.abs(value) <= 2147483647 ? value + 0 : undefined;
}

function compare(a: number, b: number, op: string): boolean {
	switch (op) {
		case '=': return a === b;
		case '<>': return a !== b;
		case '<': return a < b;
		case '>': return a > b;
		case '<=': return a <= b;
		default: return a >= b;
	}
}
