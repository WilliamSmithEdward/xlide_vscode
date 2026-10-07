// Shared VBA integer-constant expression evaluation.
//
// One evaluator for declared-constant/enum-member integer expressions, used by
// both the project-wide symbol graph (exported constant surfaces) and the
// diagnostics engine (fixed-length strings, runtime argument bounds, division
// by zero). Keeping a single copy guarantees the project-visible constant
// values and the diagnostics rules can never disagree on the same expression.
//
// The grammar is deliberately conservative: +, -, * (binary and unary +/-),
// parentheses, integer literals (decimal, &H hex, &O octal, with an optional
// %/&/^ type suffix), bare constant names, and `Module.Constant` qualified
// names. Anything else evaluates to undefined so callers never guess.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { tokenName } from '../lexer/tokenHelpers';
import { numericStringVerdict, valPrefixValue } from '../diagnostics/stringConversion';

/** The text of a string literal token: its quotes off, a doubled quote one. */
function stringLiteralText(raw: string): string {
	return raw.slice(1, raw.endsWith('"') && raw.length > 1 ? -1 : undefined).replace(/""/g, '"');
}

/** Lookup of integer constant values by lowercased (possibly qualified) name. */
export interface IntegerConstantLookup {
	get(name: string): number | undefined;
}

/** Parses an unsigned decimal integer literal, rejecting unsafe magnitudes. */
export function parseDecimalIntegerLiteral(raw: string): number | undefined {
	if (!/^\d+$/.test(raw)) {
		return undefined;
	}
	const value = Number(raw);
	return Number.isSafeInteger(value) ? value : undefined;
}

/** VBA's rounding to a whole number: banker's rounding at .5. */
export function bankersRound(value: number): number {
	const floor = Math.floor(value);
	const fraction = value - floor;
	if (fraction > 0.5) {
		return floor + 1;
	}
	if (fraction < 0.5) {
		return floor;
	}
	return floor % 2 === 0 ? floor : floor + 1;
}

/** Parses a VBA integer literal (decimal, &H, &O; optional %/&/^ suffix). */
export function parseVbaIntegerLiteral(raw: string): number | undefined {
	const trimmed = raw.trim();
	const suffix = /[%&^]$/.exec(trimmed)?.[0];
	const text = suffix ? trimmed.slice(0, -1) : trimmed;
	const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(text);
	const octal = hex ? undefined : /^&[oO]([0-7]+)$/.exec(text);
	if (hex || octal) {
		const digits = (hex ?? octal)![1];
		const value = Number.parseInt(digits, hex ? 16 : 8);
		if (!Number.isSafeInteger(value)) {
			return undefined;
		}
		// A hex or octal literal is the signed value of its bits (MS-VBAL
		// 3.3.2, measured in Excel 16.0, issue #141): up to four hex digits it
		// is an Integer, so &H8000 is -32768 and &HFFFF is -1; up to eight it
		// is a Long, so &H80000000 is -2147483648. An `&` suffix makes a short
		// literal a Long (&HFFFF& is 65535) but still wraps at 32 bits; a `^`
		// suffix (LongLong) does not wrap here.
		const fits16 = hex ? digits.length <= 4 : value <= 0xFFFF;
		if (suffix !== '&' && suffix !== '^' && fits16 && value > 0x7FFF) {
			return value - 0x10000;
		}
		if (suffix !== '^' && value > 0x7FFFFFFF && value <= 0xFFFFFFFF) {
			return value - 0x100000000;
		}
		return value;
	}
	return parseDecimalIntegerLiteral(text);
}

/** Clamps arithmetic results to safe integers; undefined when out of range. */
export function safeInteger(value: number): number | undefined {
	return Number.isSafeInteger(value) ? value : undefined;
}

/**
 * Raw value expression of an enum member: the explicit initializer when
 * present, otherwise the implicit MS-VBAL rule of previous member + 1 (first
 * member defaults to 0).
 */
export function enumMemberRawExpression(
	explicitRaw: string | undefined,
	previousName: string | undefined,
): string {
	return explicitRaw ?? (previousName ? `${previousName} + 1` : '0');
}

/** Evaluates one raw constant expression against already-known constants. */
export function evaluateIntegerConstantExpression(
	raw: string,
	constants: IntegerConstantLookup,
): number | undefined {
	const evaluation = new IntegerConstantExpressionParser(raw).parse();
	let step = evaluation.next();
	while (!step.done) {
		step = evaluation.next(constants.get(step.value));
	}
	return step.value;
}

/**
 * Recursion-depth ceiling for the descent parser. User-authored Const text is
 * untrusted, so a pathologically deep expression (e.g. thousands of nested
 * parens) must not overflow the JS stack. Past this depth we bail to undefined,
 * honoring the file's "return undefined so callers never guess" contract.
 */
const MAX_RECURSION_DEPTH = 300;

/** Returned where the current token starts no rounding call. */
const NOT_A_CALL = Symbol('not a call');

/** The calls that make a number whole, with the range each result must fit. */
const ROUNDING_CALLS: ReadonlyMap<string, readonly [number, number]> = new Map([
	['cint', [-32768, 32767]],
	['clng', [-2147483648, 2147483647]],
	['cbyte', [0, 255]],
	['int', [-Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
	['fix', [-Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
	['round', [-Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
]);

type IntegerConstantEvaluation = Generator<string, number | undefined, number | undefined>;

/** Yields each lookup without retaining another constant parser on the JS stack. */
class IntegerConstantExpressionParser {
	private readonly tokens: VbaToken[];
	private index = 0;
	private depth = 0;
	/** Inside a call's arguments, where True and False pass as -1 and 0. */
	private inArguments = 0;

	constructor(raw: string) {
		this.tokens = tokenize(raw).filter((token) => token.kind !== 'comment' && token.kind !== 'newline');
	}

	*parse(): IntegerConstantEvaluation {
		if (this.tokens.length === 0) {
			return undefined;
		}
		// Single literals and aliases need no precedence frames. Aliases
		// still yield to the same dependency scheduler.
		if (this.tokens.length === 1) {
			const token = this.tokens[0];
			if (token.kind === 'integerLiteral') {
				return parseVbaIntegerLiteral(token.rawText);
			}
			if (token.kind === 'keyword' && token.rawText.toLowerCase() === 'not') {
				return undefined;
			}
			const name = tokenName(token);
			return name ? yield name.toLowerCase() : undefined;
		}
		if (this.tokens.length === 3 && !(this.tokens[0].kind === 'keyword' && this.tokens[0].rawText.toLowerCase() === 'not')) {
			const qualified = this.qualifiedName();
			if (qualified) {
				return yield qualified.toLowerCase();
			}
		}
		const value = yield* this.expression();
		return value !== undefined && !this.current() ? value : undefined;
	}

	private *expression(): IntegerConstantEvaluation {
		// Depth guard: untrusted Const text can nest arbitrarily deep; bail to
		// undefined rather than overflowing the stack.
		if (++this.depth > MAX_RECURSION_DEPTH) {
			this.depth--;
			return undefined;
		}
		const result = yield* this.logical(0);
		this.depth--;
		return result;
	}

	/**
	 * Xor, Or and And, loosest first, then Not, over whole numbers within the
	 * Long range: `Const K0 = 15 And 255` is 15 (issue #496, measured in Excel
	 * 16.0).
	 */
	private *logical(level: number): IntegerConstantEvaluation {
		if (level === LOGICAL_LEVELS.length) {
			// A flat Not chain needs no recursive frames. Keep the range check
			// even when an even number of operators would cancel each other.
			if (!this.acceptWord('not')) {
				return yield* this.expressionInner();
			}
			let invert = true;
			while (this.acceptWord('not')) {
				invert = !invert;
			}
			const operand = yield* this.expressionInner();
			return operand === undefined || !isLong(operand) ? undefined : invert ? ~operand : operand | 0;
		}
		const word = LOGICAL_LEVELS[level];
		let value = yield* this.logical(level + 1);
		while (value !== undefined && this.acceptWord(word)) {
			const right = yield* this.logical(level + 1);
			if (right === undefined || !isLong(value) || !isLong(right)) {
				return undefined;
			}
			value = word === 'and' ? value & right : word === 'or' ? value | right : value ^ right;
		}
		return value;
	}

	private *expressionInner(): IntegerConstantEvaluation {
		let value = yield* this.modulo();
		while (value !== undefined) {
			if (this.accept('+')) {
				const right = yield* this.modulo();
				value = right === undefined ? undefined : safeInteger(value + right);
				continue;
			}
			if (this.accept('-')) {
				const right = yield* this.modulo();
				value = right === undefined ? undefined : safeInteger(value - right);
				continue;
			}
			break;
		}
		return value;
	}

	/** Mod binds below `\`, and `\` below `*`: `50 Mod 7 + 10` is 11. */
	private *modulo(): IntegerConstantEvaluation {
		let value = yield* this.integerDivision();
		while (value !== undefined && this.acceptWord('mod')) {
			const right = yield* this.integerDivision();
			value = right === undefined || right === 0 ? undefined : safeInteger(value % right);
		}
		return value;
	}

	private *integerDivision(): IntegerConstantEvaluation {
		let value = yield* this.term();
		while (value !== undefined && this.accept('\\')) {
			const right = yield* this.term();
			value = right === undefined || right === 0 ? undefined : safeInteger(Math.trunc(value / right));
		}
		return value;
	}

	private *term(): IntegerConstantEvaluation {
		let value = yield* this.factor();
		while (value !== undefined) {
			if (!this.accept('*')) {
				break;
			}
			const right = yield* this.factor();
			value = right === undefined ? undefined : safeInteger(value * right);
		}
		return value;
	}

	private *factor(): IntegerConstantEvaluation {
		// Depth guard: unary +/- chains and nested parens recurse through factor;
		// bail to undefined once the ceiling is hit (see expression()).
		if (++this.depth > MAX_RECURSION_DEPTH) {
			this.depth--;
			return undefined;
		}
		const result = yield* this.factorInner();
		this.depth--;
		return result;
	}

	private *factorInner(): IntegerConstantEvaluation {
		if (this.accept('+')) {
			return yield* this.factor();
		}
		if (this.accept('-')) {
			const value = yield* this.factor();
			return value === undefined ? undefined : safeInteger(-value);
		}
		if (this.accept('(')) {
			const value = yield* this.expression();
			return value !== undefined && this.accept(')') ? value : undefined;
		}
		const token = this.current();
		if (!token) {
			return undefined;
		}
		if (token.kind === 'integerLiteral') {
			this.index++;
			return parseVbaIntegerLiteral(token.rawText);
		}
		// True is -1 as a number, False 0: `F(False)` (issue #562). Only there:
		// a Byte Const of True holds 255.
		const word = token.rawText.toLowerCase();
		if (this.inArguments > 0 && (word === 'true' || word === 'false') && this.tokens[this.index - 1]?.rawText !== '.') {
			this.index++;
			return word === 'true' ? -1 : 0;
		}
		const qualified = this.qualifiedName();
		if (qualified) {
			return yield qualified.toLowerCase();
		}
		const rounded = yield* this.roundingCall();
		if (rounded !== NOT_A_CALL) {
			return rounded;
		}
		const name = tokenName(token);
		if (name) {
			this.index++;
			// `F()`: a lookup may know a Function's result as `f()` (issue #448).
			if (this.tokens[this.index]?.rawText === '(' && this.tokens[this.index + 1]?.rawText === ')') {
				this.index += 2;
				return yield `${name.toLowerCase()}()`;
			}
			// `F(-1)`: a call with whole-number arguments is `f(-1)` to a lookup (issue #562).
			if (this.tokens[this.index]?.rawText === '(' && this.tokens[this.index - 2]?.rawText !== '.') {
				this.index++;
				this.inArguments++;
				const args: Array<number | undefined> = [(yield* this.expression())];
				while (this.accept(',')) {
					args.push((yield* this.expression()));
				}
				this.inArguments--;
				if (!this.accept(')') || args.some((arg) => arg === undefined)) {
					return undefined;
				}
				return yield `${name.toLowerCase()}(${args.join(',')})`;
			}
			return yield name.toLowerCase();
		}
		return undefined;
	}

	/**
	 * `CInt(3.5)`, `Int(-0.1)`, `Fix(-0.9)`, `Round(0.5)`: a conversion or
	 * rounding of a number, which comes out whole (issue #286, measured in
	 * Excel 16.0). CInt, CLng and Round round half to even, Int down and Fix
	 * toward zero. NOT_A_CALL where the current token starts no such call;
	 * undefined where it does and the argument is not known or the result
	 * does not fit.
	 */
	private *roundingCall(): Generator<string, number | undefined | typeof NOT_A_CALL, number | undefined> {
		const word = tokenName(this.current())?.toLowerCase();
		// `Val("0,5")` is 0 in every locale (issue #703).
		if (word === 'val' && this.tokens[this.index + 1]?.rawText === '(' && this.tokens[this.index + 2]?.kind === 'stringLiteral'
			&& this.tokens[this.index + 3]?.rawText === ')' && this.tokens[this.index - 1]?.rawText !== '.') {
			const value = valPrefixValue(stringLiteralText(this.tokens[this.index + 2].rawText));
			if (value !== undefined) {
				this.index += 4;
				return value;
			}
			return NOT_A_CALL;
		}
		const range = word ? ROUNDING_CALLS.get(word) : undefined;
		if (!range || this.tokens[this.index + 1]?.rawText !== '(' || this.tokens[this.index - 1]?.rawText === '.') {
			return NOT_A_CALL;
		}
		const start = this.index;
		this.index += 2;
		const argument = this.index;
		const negative = this.accept('-');
		const tok = this.current();
		let value: number | undefined;
		if (tok?.kind === 'stringLiteral' && !negative && this.tokens[this.index + 1]?.rawText === ')') {
			// `CInt("(5)")`, `CLng("&H0")`: a string every locale reads alike
			// (issue #703, measured in Excel 16.0).
			this.index++;
			const verdict = numericStringVerdict(stringLiteralText(tok.rawText));
			value = verdict.kind === 'number' ? verdict.value : undefined;
		} else if (tok?.kind === 'floatLiteral' && this.tokens[this.index + 1]?.rawText === ')') {
			this.index++;
			const read = Number(tok.rawText.replace(/[!#@]$/, ''));
			value = Number.isFinite(read) ? (negative ? -read : read) : undefined;
		} else {
			this.index = argument;
			value = yield* this.expression();
		}
		if (value === undefined || !this.accept(')')) {
			this.index = start;
			return NOT_A_CALL;
		}
		const whole = word === 'int' ? Math.floor(value) : word === 'fix' ? Math.trunc(value) : bankersRound(value);
		return whole >= range[0] && whole <= range[1] ? whole + 0 : undefined;
	}

	private qualifiedName(): string | undefined {
		const qualifier = tokenName(this.current());
		const dot = this.tokens[this.index + 1];
		const member = tokenName(this.tokens[this.index + 2]);
		if (!qualifier || dot?.rawText !== '.' || !member) {
			return undefined;
		}
		this.index += 3;
		return `${qualifier}.${member}`;
	}

	private current(): VbaToken | undefined {
		return this.tokens[this.index];
	}

	private accept(raw: string): boolean {
		if (this.current()?.rawText !== raw) {
			return false;
		}
		this.index++;
		return true;
	}

	private acceptWord(word: string): boolean {
		const token = this.current();
		if (token?.kind !== 'keyword' || token.rawText.toLowerCase() !== word) {
			return false;
		}
		this.index++;
		return true;
	}
}

/** The logical operators, loosest first. */
const LOGICAL_LEVELS: readonly string[] = ['xor', 'or', 'and'];

function isLong(value: number): boolean {
	return Number.isInteger(value) && value >= -2147483648 && value <= 2147483647;
}

/**
 * Resolves raw constant expressions (lowercased name -> raw text, undefined
 * for ambiguous duplicates) to integer values, memoized with cycle detection.
 * Names absent from `rawConstants` fall back to the optional `base` map of
 * already-resolved values; the returned map only contains `rawConstants` keys.
 */
export function resolveRawIntegerConstants(
	rawConstants: ReadonlyMap<string, string | undefined>,
	base: ReadonlyMap<string, number | undefined> = new Map(),
): Map<string, number | undefined> {
	const resolved = new Map<string, number | undefined>();
	const resolving = new Set<string>();
	// Suspended parsers keep their cursor and partial arithmetic values. Resume
	// them directly so wide expressions never replay their already-read prefix.
	type Frame = {key: string; evaluation: IntegerConstantEvaluation};
	const pending: Frame[] = [];
	for (const name of rawConstants.keys()) {
		const key = name.toLowerCase();
		if (resolved.has(key)) {
			continue;
		}
		if (!rawConstants.has(key)) {
			base.get(key);
			continue;
		}
		const raw = rawConstants.get(key);
		if (raw === undefined) {
			resolved.set(key, undefined);
			continue;
		}
		pending.push({key, evaluation: new IntegerConstantExpressionParser(raw).parse()});
		resolving.add(key);
		let value: number | undefined;
		while (pending.length > 0) {
			const frame = pending[pending.length - 1];
			const step = frame.evaluation.next(value);
			if (step.done) {
				value = step.value;
				resolved.set(frame.key, value);
				resolving.delete(frame.key);
				pending.pop();
				continue;
			}
			const dependency = step.value.toLowerCase();
			if (resolved.has(dependency)) {
				value = resolved.get(dependency);
				continue;
			}
			if (!rawConstants.has(dependency)) {
				value = base.get(dependency);
				continue;
			}
			if (resolving.has(dependency)) {
				resolved.set(dependency, undefined); value = undefined;
				continue;
			}
			const expression = rawConstants.get(dependency);
			if (expression === undefined) {
				resolved.set(dependency, undefined); value = undefined;
				continue;
			}
			pending.push({key: dependency, evaluation: new IntegerConstantExpressionParser(expression).parse()});
			resolving.add(dependency);
			value = undefined;
		}
	}
	return resolved;
}
