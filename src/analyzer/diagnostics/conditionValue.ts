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
import { matchParenFrom, splitTopLevelTokenGroups, tokenName, tokenWord as tokenText } from '../lexer/tokenHelpers';
import type { ModuleCompare } from './knownStringCalls';
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
	/**
	 * The module's Option Compare, which decides strings that differ only in
	 * case and how strings order (issue #686). Without it only what holds
	 * under either is decided.
	 */
	compare?: ModuleCompare;
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
		// `s Like "b*"`, by the module's compare mode (issue #686).
		if (this.word() === 'like') {
			this.index++;
			const pattern = this.sum();
			return typeof left === 'string' && typeof pattern === 'string' ? likeMatch(left, pattern, textCompare(this.facts.compare, left + pattern)) : undefined;
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
		if (typeof left === 'string' && typeof right === 'string') {
			const order = stringOrder(left, right, this.facts.compare);
			if (order !== undefined) {
				return compare(order, 0, op);
			}
			if (op !== '=' && op !== '<>') {
				return undefined;
			}
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
		// LCase, UCase, InStr, StrComp and Replace of known strings (issue #686).
		if (STRING_CALLS.has(word)) {
			const called = this.stringCall(word);
			if (called !== NOT_A_CALL) {
				return called;
			}
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

	/**
	 * A call of a string function at the token before `this.index`, its
	 * arguments each a condition operand. NOT_A_CALL where no `(` follows;
	 * undefined where an argument or the result is not known.
	 */
	private stringCall(word: string): Value | typeof NOT_A_CALL {
		let open = this.index;
		if (this.toks[open]?.rawText === '$') {
			open++;
		}
		if (this.toks[open]?.rawText !== '(' || this.toks[this.index - 2]?.rawText === '.') {
			return NOT_A_CALL;
		}
		const close = matchParenFrom(this.toks, open);
		if (close < 0) {
			return NOT_A_CALL;
		}
		const args = splitTopLevelTokenGroups(this.toks, open + 1, ',', close).map((arg) => this.argumentValue(arg));
		this.index = close + 1;
		const mode = this.facts.compare;
		const compareArg = (value: Value): boolean | undefined => (value === undefined ? undefined : value === 1 ? true : value === 0 ? false : undefined);
		switch (word) {
			case 'lcase':
			case 'ucase': {
				const [s] = args;
				return args.length === 1 && typeof s === 'string' && isAscii(s) ? (word === 'lcase' ? s.toLowerCase() : s.toUpperCase()) : undefined;
			}
			case 'instr': {
				const [start, s1, s2, how] = args.length >= 3 ? args : [1, ...args];
				if (typeof start !== 'number' || !Number.isInteger(start) || start < 1 || typeof s1 !== 'string' || typeof s2 !== 'string' || args.length > 4) {
					return undefined;
				}
				const text = args.length === 4 ? compareArg(how) : textCompare(mode, s1 + s2);
				if (text === undefined || (text && !isAscii(s1 + s2))) {
					return undefined;
				}
				if (s1.length === 0) {
					return 0;
				}
				if (start > s1.length) {
					return s2.length === 0 ? start : 0;
				}
				const found = (text ? s1.toLowerCase() : s1).indexOf(text ? s2.toLowerCase() : s2, start - 1);
				return found + 1;
			}
			case 'strcomp': {
				const [a, b, how] = args;
				if (typeof a !== 'string' || typeof b !== 'string' || args.length > 3) {
					return undefined;
				}
				const text = args.length === 3 ? compareArg(how) : undefined;
				return stringOrder(a, b, args.length === 3 ? (text === undefined ? 'database' : text ? 'text' : 'binary') : mode);
			}
			case 'replace': {
				const [expression, find, replacement, start, count, how] = args;
				if (typeof expression !== 'string' || typeof find !== 'string' || typeof replacement !== 'string'
					|| (start !== undefined && start !== 1) || (count !== undefined && count !== -1) || args.length > 6) {
					return undefined;
				}
				if (find.length === 0) {
					return expression;
				}
				const text = args.length === 6 ? compareArg(how) : textCompare(mode, expression + find);
				if (text === undefined || (text && !isAscii(expression + find))) {
					return undefined;
				}
				return expression.replace(new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), text ? 'gi' : 'g'), () => replacement);
			}
		}
		return undefined;
	}

	/** An argument's value: a condition operand, and vbBinaryCompare or vbTextCompare as 0 or 1. */
	private argumentValue(arg: readonly VbaToken[]): Value {
		const toks = arg.filter((tok) => tok.kind !== 'comment');
		const word = toks.length === 1 ? tokenText(toks[0]) : '';
		if (word === 'vbbinarycompare' || word === 'vbtextcompare') {
			return word === 'vbtextcompare' ? 1 : 0;
		}
		const parser = new ConditionParser(toks, this.facts);
		const value = parser.sum();
		return parser.done() && typeof value !== 'object' ? value : undefined;
	}
}

/** The string functions a condition reads of known strings (issue #686). */
const STRING_CALLS: ReadonlySet<string> = new Set(['lcase', 'ucase', 'instr', 'strcomp', 'replace']);

const NOT_A_CALL = Symbol('not a call');

function isAscii(text: string): boolean {
	return /^[\x00-\x7f]*$/.test(text);
}

/**
 * Whether the module compares these strings as text (true) or binary
 * (false), or undefined where that is not known: no Option Compare given
 * to the rule, Option Compare Database, or text with a character past ASCII,
 * which the locale folds.
 */
function textCompare(mode: ModuleCompare | undefined, sample: string): boolean | undefined {
	if (mode === 'binary') {
		return false;
	}
	return mode === 'text' && isAscii(sample) ? true : undefined;
}

/**
 * How two strings order, -1, 0 or 1, where the module's compare mode makes
 * it certain (issue #686, measured in Excel 16.0): Binary orders ASCII by
 * character code, so "A" < "a"; Text ignores case, and its locale sort is
 * followed only for letters and digits. Equal strings are 0 in any mode.
 */
function stringOrder(a: string, b: string, mode: ModuleCompare | undefined): number | undefined {
	if (a === b) {
		return 0;
	}
	if (mode === 'binary' && isAscii(a + b)) {
		return a < b ? -1 : 1;
	}
	if (mode === 'text' && /^[A-Za-z0-9]*$/.test(a + b)) {
		const x = a.toLowerCase();
		const y = b.toLowerCase();
		return x === y ? 0 : x < y ? -1 : 1;
	}
	return undefined;
}

/**
 * Whether `subject Like pattern` holds: `*`, `?`, `#` and `[list]` with `!`
 * and ranges. Undefined where the compare mode decides a letter and is not
 * known, or the pattern is not one this reads.
 */
function likeMatch(subject: string, pattern: string, text: boolean | undefined): boolean | undefined {
	if (text === undefined && /[A-Za-z]/.test(subject + pattern)) {
		return undefined;
	}
	let source = '';
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '*') {
			source += '[\\s\\S]*';
		} else if (ch === '?') {
			source += '[\\s\\S]';
		} else if (ch === '#') {
			source += '[0-9]';
		} else if (ch === '[') {
			const end = pattern.indexOf(']', i + 1);
			if (end < 0) {
				return undefined;
			}
			let list = pattern.slice(i + 1, end);
			const negate = list.startsWith('!');
			if (negate) {
				list = list.slice(1);
			}
			if (/(.)-(.)/.test(list) && [...list.matchAll(/(.)-(.)/g)].some((range) => range[1] > range[2])) {
				return undefined;
			}
			source += `[${negate ? '^' : ''}${list.replace(/[\\\]^]/g, '\\$&')}]`;
			i = end;
		} else {
			source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		}
	}
	return new RegExp(`^${source}$`, text ? 'i' : '').test(subject);
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
