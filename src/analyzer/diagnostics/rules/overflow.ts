// Rule family: overflow the analyzer can prove (issue #116).
//
// VBA types a whole-number literal as Integer when it fits and does Integer
// arithmetic on two Integers, so `60 * 60 * 24` overflows before the Long it
// is assigned to ever sees it. Every case here was measured in Excel 16.0
// (build 20326, 2026-09-26): each compiles and raises error 6 every time it
// runs, or - for a Const - is refused with "Overflow" while compiling.
//
//  - arithmetic-overflow: `secs = 60 * 60 * 24`, `32767 + 1` into a Long,
//    `50000 * 50000`, `2147483647 + 1`, `10 ^ 309`, `Exp(1000)`, Integer Consts
//    multiplied, `CInt(40000)`, `CByte(-1)`, `CLng(2147483647.5)`,
//    `CSng(1E+39)`, `Hex(1E+20)`, `Abs(CInt(-32768))`, `-i` with i = -32768,
//    and an assignment whose folded value the target type cannot hold after
//    rounding: `Byte = 255.5`, `Integer = 32767.5`, `Date = 3000000`.
//  - const-overflow: the same folding on a Const's value, a compile error.
//    Inside any argument or operand too (issue #232): `CStr(CInt(40000))`,
//    `IIf(True, 0, CInt(40000))`, `"x" & CInt(40000)`, `z(CInt(40000))`.
//    LongLong: `CLngLng(1E+19)`, `9223372036854775807^ + 1`; a LongPtr is
//    judged against LongLong's range, which it never exceeds.
//  - for-counter-overflow: `For i = 1 To 32767` with i an Integer, and
//    `For b = 0 To 255` with b a Byte: the increment after the last pass
//    overflows the counter. `To 32766` runs.
//
// The folder follows MS-VBAL 5.6.9.3: two Bytes make a Byte, Byte and
// Integer make Integer, Long makes Long, Single and Double make Double,
// Currency makes Currency, and a Date plus or minus a number is a Date;
// `/` and `^` make Double. A value the folder cannot type stays unknown and
// nothing is reported for it.

import { DATE_EPOCH_MS, DAY_MS, dateLiteralSerial } from '../../constants/dateLiteral';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import type { HostObjectModel } from '../../host/excelObjectModel';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ForBlockNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { bankersRound, isBareOrVbaQualifiedIntrinsicCall, namesIn } from './shared';
import { blockHeaderLeaves, isLoopBlock, selectArms } from '../blockHeaders';
import {
	knownLocalLiteralValues,
	normalizeType,
	stringLiteralValue,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	blockHeaderStatements,
	firstExecutableTokenIndex,
	forEachVariableGroup,
	matchParenFrom,
	statementAndBranchSpans,
	rawExpressionTokens,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

type NumericType = 'byte' | 'integer' | 'long' | 'longlong' | 'single' | 'double' | 'currency' | 'date';

interface Typed {
	value: number;
	type: NumericType;
	/** A LongLong's exact value: a double cannot tell 2^63 - 1 from 2^63. */
	exact?: bigint;
	/**
	 * Made of literals and Consts only, so the VBE folds it while compiling:
	 * negating the Long minimum there wraps to itself (issue #235).
	 */
	constant?: boolean;
}

interface Overflow {
	overflow: true;
	span: Span;
	detail: string;
}

type Folded = Typed | Overflow | undefined;

const RANGES: Readonly<Record<NumericType, { min: number; max: number; label: string }>> = {
	byte: { min: 0, max: 255, label: 'Byte' },
	integer: { min: -32768, max: 32767, label: 'Integer' },
	long: { min: -2147483648, max: 2147483647, label: 'Long' },
	// As doubles these are -2^63 and 2^63; inRange compares a LongLong exactly.
	longlong: { min: -9223372036854775808, max: 9223372036854775807, label: 'LongLong' },
	single: { min: -3.402823e38, max: 3.402823e38, label: 'Single' },
	double: { min: -1.7976931348623157e308, max: 1.7976931348623157e308, label: 'Double' },
	currency: { min: -922337203685477.5807, max: 922337203685477.5807, label: 'Currency' },
	date: { min: -657434, max: 2958465, label: 'Date' },
};

const RANK: Readonly<Record<NumericType, number>> = {
	byte: 0, integer: 1, long: 2, longlong: 3, single: 4, double: 5, currency: 6, date: 7,
};

/**
 * The result type of `a op b` for + - * \ Mod (MS-VBAL 5.6.9.3), as Excel
 * 16.0 computes it (issue #203): two Bytes make a Byte, so 200 + 100
 * overflows; a Date plus or minus a number, or two Dates added, make a Date,
 * so #12/31/9999# + 1 overflows; two Dates subtracted make a Double.
 */
function arithmeticResultType(a: NumericType, b: NumericType, op: string): NumericType {
	if ((a === 'longlong' || b === 'longlong') && (a === 'single' || b === 'single')) {
		return 'double';
	}
	if (a === 'currency' || b === 'currency') {
		return (a === 'double' || b === 'double' || a === 'single' || b === 'single') ? 'double' : 'currency';
	}
	if (a === 'date' || b === 'date') {
		return op === '+' || (op === '-' && !(a === 'date' && b === 'date')) ? 'date' : 'double';
	}
	return RANK[a] >= RANK[b] ? a : b;
}

function isOverflow(folded: Folded): folded is Overflow {
	return folded !== undefined && 'overflow' in folded;
}

/** 2^63: one past the largest LongLong. */
const LONGLONG_LIMIT = 2n ** 63n;

function inRange(value: number, type: NumericType, exact?: bigint): boolean {
	if (type === 'longlong') {
		return exact !== undefined
			? exact >= -LONGLONG_LIMIT && exact < LONGLONG_LIMIT
			: Number.isFinite(value) && value >= -(2 ** 63) && value < 2 ** 63;
	}
	const range = RANGES[type];
	// A Date's range is of days: any time of 12/31/9999 is in it.
	const checked = type === 'date' ? Math.trunc(value) : value;
	return Number.isFinite(value) && checked >= range.min && checked <= range.max;
}


/** A literal's natural type and value: 3 is Integer, 40000 is Long, 3000000000 is Double. */
function literalTyped(tok: VbaToken): Typed | undefined {
	const typed = literalValue(tok);
	return typed ? { ...typed, constant: true } : undefined;
}

function literalValue(tok: VbaToken): Typed | undefined {
	// A Boolean in arithmetic is an Integer: True is -1 and False 0, so
	// `1 / False` divides by zero (issue #235).
	if (tok.kind === 'keyword') {
		const word = tokenText(tok);
		return word === 'true' ? { value: -1, type: 'integer' } : word === 'false' ? { value: 0, type: 'integer' } : undefined;
	}
	if (tok.kind === 'integerLiteral') {
		const raw = tok.rawText;
		const suffix = /[%&^]$/.exec(raw)?.[0];
		if (suffix === '^') {
			const exact = longLongLiteral(raw);
			return exact === undefined ? undefined : { value: Number(exact), type: 'longlong', exact };
		}
		const value = parseVbaIntegerLiteral(raw);
		if (value === undefined) {
			return undefined;
		}
		if (suffix === '%') {
			return inRange(value, 'integer') ? { value, type: 'integer' } : undefined;
		}
		if (suffix === '&') {
			return { value, type: 'long' };
		}
		// A hex or octal literal arrives already signed by its width
		// (parseVbaIntegerLiteral, issue #141): &H8000 is -32768 and an
		// Integer, &H80000000 is -2147483648 and a Long.
		if (inRange(value, 'integer')) {
			return { value, type: 'integer' };
		}
		if (inRange(value, 'long')) {
			return { value, type: 'long' };
		}
		return { value, type: 'double' };
	}
	if (tok.kind === 'floatLiteral') {
		const raw = tok.rawText.replace(/[dD]/g, 'E');
		const suffix = /[!#@]$/.exec(raw)?.[0];
		const value = Number(raw.replace(/[!#@]$/, ''));
		if (!Number.isFinite(value)) {
			return undefined;
		}
		return { value, type: suffix === '!' ? 'single' : suffix === '@' ? 'currency' : 'double' };
	}
	if (tok.kind === 'dateLiteral') {
		const serial = dateLiteralSerial(tok.rawText);
		return serial === undefined ? undefined : { value: serial, type: 'date' };
	}
	return undefined;
}

/**
 * A `^` literal's exact value: `9223372036854775807^`, `&H7FFFFFFFFFFFFFFF^`
 * (a hex or octal literal keeps the sign of its width). A decimal past the
 * range is a syntax error that suffixed-literal-overflow reports.
 */
function longLongLiteral(raw: string): bigint | undefined {
	const body = raw.slice(0, -1);
	let value: bigint;
	if (/^\d+$/.test(body)) {
		value = BigInt(body);
		return value < LONGLONG_LIMIT ? value : undefined;
	}
	const radix = /^&([Hh])([0-9A-Fa-f]+)$|^&[Oo]?([0-7]+)$/.exec(body);
	if (!radix) {
		return undefined;
	}
	value = radix[2] !== undefined ? BigInt(`0x${radix[2]}`) : BigInt(`0o${radix[3]}`);
	return value < 2n ** 64n ? BigInt.asIntN(64, value) : undefined;
}

/**
 * The number a conversion reads from a string (issue #184), for the spellings
 * every locale reads alike: whole digits with an optional exponent, and &H
 * and &O literals, which keep the sign of their width as a literal does
 * ("&H8000" is -32768). A decimal point or a thousands separator is read by
 * the locale and is not judged.
 */
function numberInString(text: string): Typed | 'overflow' | undefined {
	const trimmed = text.trim();
	if (/^[-+]?\d+(?:[eE][-+]?\d+)?$/.test(trimmed)) {
		const value = Number(trimmed);
		return Number.isFinite(value) ? { value, type: 'double' } : 'overflow';
	}
	const radix = /^([-+]?)(&[Hh][0-9A-Fa-f]+|&[Oo]?[0-7]+)$/.exec(trimmed);
	const value = radix ? parseVbaIntegerLiteral(radix[2]) : undefined;
	if (value === undefined) {
		return undefined;
	}
	return { value: radix![1] === '-' ? -value : value, type: 'double' };
}

/**
 * What `Val` reads from a string: a number with an optional fraction and
 * exponent, always with `.` as the decimal point. Only a string that is that
 * number and nothing else is judged, since Val also skips blanks inside one.
 */
function valOfString(text: string): Typed | 'overflow' | undefined {
	const match = /^\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eEdD][-+]?\d+)?)\s*$/.exec(text);
	if (!match) {
		return undefined;
	}
	const value = Number(match[1].replace(/[dD]/, 'E'));
	return Number.isFinite(value) ? { value, type: 'double' } : 'overflow';
}

/** What a name means to the folder: a typed value, or nothing. */
type NameLookup = (lower: string) => Typed | undefined;

/**
 * Folds an arithmetic expression over literals, Consts and known locals with
 * VBA's result typing, stopping at the first operation whose result its type
 * cannot hold.
 */
class TypedFolder {
	private index = 0;

	constructor(
		private readonly toks: readonly VbaToken[],
		private readonly base: number,
		private readonly names: NameLookup,
		/** Told of a division by zero, which a Const cannot hold: "Division by zero" while compiling. */
		private readonly divisionByZero?: (span: Span) => void,
	) {}

	fold(): Folded {
		if (this.toks.length === 0) {
			return undefined;
		}
		// `Not` binds below every arithmetic operator: `Not 255 + 256` is
		// Not 511 (issue #235, measured in Excel 16.0).
		if (this.toks[0].kind === 'keyword' && tokenText(this.toks[0]) === 'not') {
			const operand = new TypedFolder(this.toks.slice(1), this.base, this.names, this.divisionByZero).fold();
			return operand === undefined || isOverflow(operand) ? operand : notOf(operand, this.span(0, this.toks.length - 1));
		}
		const result = this.additive();
		if (isOverflow(result)) {
			return result;
		}
		return this.index === this.toks.length ? result : undefined;
	}

	private span(from: number, to: number): Span {
		return { start: this.base + this.toks[from].start, end: this.base + this.toks[to].end };
	}

	private additive(): Folded {
		const start = this.index;
		let left = this.multiplicative();
		while (left !== undefined && !isOverflow(left)) {
			const op = this.toks[this.index];
			if (!op || op.kind !== 'operator' || (op.rawText !== '+' && op.rawText !== '-')) {
				break;
			}
			this.index++;
			const right = this.multiplicative();
			if (right === undefined || isOverflow(right)) {
				return right;
			}
			left = this.combine(left, right, op.rawText, start, this.index - 1);
		}
		return left;
	}

	// MS-VBAL 5.6.9 arithmetic precedence, highest first: ^, unary minus,
	// * and /, \, Mod, + and -. Folding *, /, \, Mod and ^ at one level read
	// `32000 \ 2 * 4` as 16000 * 4 and reported an overflow on code that
	// runs, and missed `1 Mod 200 * 200`, which does overflow (issue #145).
	private leftAssociative(operators: readonly string[], left: () => Folded, right: () => Folded = left): Folded {
		const start = this.index;
		let value = left();
		while (value !== undefined && !isOverflow(value)) {
			const op = this.toks[this.index];
			if (!op) {
				break;
			}
			const word = op.kind === 'operator' ? op.rawText : tokenText(op);
			if (!operators.includes(word)) {
				break;
			}
			this.index++;
			const operand = right();
			if (operand === undefined || isOverflow(operand)) {
				return operand;
			}
			value = this.combine(value, operand, word, start, this.index - 1);
		}
		return value;
	}

	private multiplicative(): Folded {
		return this.leftAssociative(['mod'], () => this.integerDivision());
	}

	private integerDivision(): Folded {
		return this.leftAssociative(['\\'], () => this.product());
	}

	private product(): Folded {
		return this.leftAssociative(['*', '/'], () => this.unary());
	}

	/** `a ^ b` binds above unary minus (`-2 ^ 2` is -4); the exponent may carry its own sign. */
	private power(): Folded {
		return this.leftAssociative(['^'], () => this.primary(), () => this.unary());
	}

	private unary(): Folded {
		const tok = this.toks[this.index];
		if (tok?.kind === 'operator' && (tok.rawText === '-' || tok.rawText === '+')) {
			const start = this.index;
			this.index++;
			const operand = this.unary();
			if (operand === undefined || isOverflow(operand)) {
				return operand;
			}
			if (tok.rawText === '+') {
				return operand;
			}
			// `-&H80000000` and `-(-2147483647 - 1)` are folded while compiling
			// and give the Long minimum back; `l = -l` overflows (issue #235).
			if (operand.constant && operand.type === 'long' && operand.value === RANGES.long.min) {
				return operand;
			}
			const value = -operand.value;
			const exact = operand.exact !== undefined ? -operand.exact : undefined;
			if (!inRange(value, operand.type, exact)) {
				return { overflow: true, span: this.span(start, this.index - 1), detail: `Negating ${exact !== undefined ? String(operand.exact) : showNumber(operand.value)} gives ${exact !== undefined ? String(exact) : showNumber(-operand.value)}, which does not fit ${RANGES[operand.type].label}` };
			}
			return { value, type: operand.type, ...(exact !== undefined ? { exact } : {}), ...(operand.constant ? { constant: true } : {}) };
		}
		return this.power();
	}

	private primary(): Folded {
		const tok = this.toks[this.index];
		if (!tok) {
			return undefined;
		}
		if (tok.rawText === '(') {
			const close = matchParenFrom(this.toks, this.index);
			if (close < 0) {
				return undefined;
			}
			const inner = new TypedFolder(this.toks.slice(this.index + 1, close), this.base, this.names, this.divisionByZero);
			const value = inner.fold();
			this.index = close + 1;
			return value;
		}
		const literal = literalTyped(tok);
		if (literal) {
			this.index++;
			return literal;
		}
		const name = tokenName(tok);
		if (!name) {
			return undefined;
		}
		// `VBA.CInt(...)` and `CInt(...)`.
		let calleeIndex = this.index;
		if (name.toLowerCase() === 'vba' && this.toks[this.index + 1]?.rawText === '.' && tokenName(this.toks[this.index + 2])) {
			calleeIndex = this.index + 2;
		}
		const callee = tokenName(this.toks[calleeIndex])!.toLowerCase();
		if (this.toks[calleeIndex + 1]?.rawText === '(' && (CONVERSIONS.has(callee) || callee === 'val')) {
			const close = matchParenFrom(this.toks, calleeIndex + 1);
			if (close < 0) {
				return undefined;
			}
			const start = this.index;
			const argument = this.toks.slice(calleeIndex + 2, close);
			this.index = close + 1;
			// A string literal is read as the number it spells (issue #184):
			// `CInt("&H10000")` is CInt(65536), `Val("1e400")` is past a Double.
			if (argument.length === 1 && argument[0].kind === 'stringLiteral') {
				const text = stringLiteralValue(argument[0].rawText);
				const read = callee === 'val' ? valOfString(text) : numberInString(text);
				if (read === 'overflow') {
					return { overflow: true, span: this.span(start, close), detail: `${argument[0].rawText} spells a number past the Double range` };
				}
				if (read === undefined || callee === 'val') {
					return read;
				}
				return this.convert(callee, read, this.span(start, close), argument[0].rawText);
			}
			if (callee === 'val') {
				return undefined;
			}
			const inner = new TypedFolder(argument, this.base, this.names).fold();
			if (inner === undefined || isOverflow(inner)) {
				return inner;
			}
			return this.convert(callee, inner, this.span(start, close));
		}
		if (this.toks[this.index + 1]?.rawText === '.') {
			// `Rows.Count`: a two-part member the lookup may know as a constant.
			const member = tokenName(this.toks[this.index + 2]);
			const after = this.toks[this.index + 3]?.rawText;
			if (member && after !== '.' && after !== '(') {
				const known = this.names(`${name.toLowerCase()}.${member.toLowerCase()}`);
				if (known) {
					this.index += 3;
					return known;
				}
			}
			return undefined;
		}
		if (this.toks[this.index + 1]?.rawText === '(') {
			return undefined; // a call the folder does not know
		}
		const known = this.names(name.toLowerCase());
		if (!known) {
			return undefined;
		}
		this.index++;
		return known;
	}

	private convert(callee: string, inner: Typed, span: Span, shown = showNumber(inner.value)): Folded {
		const target = CONVERSIONS.get(callee)!;
		if (target === 'abs') {
			// Abs of the smallest Long hands it back unchanged: Abs(CLng(-2147483647
			// - 1)) is -2147483648, where Abs(CInt(-32768)) overflows (issue #218,
			// measured in Excel 16.0).
			if (inner.type === 'long' && inner.value === RANGES.long.min) {
				return inner;
			}
			const value = Math.abs(inner.value);
			return inRange(value, inner.type)
				? { value, type: inner.type }
				: { overflow: true, span, detail: `Abs(${showNumber(inner.value)}) does not fit ${RANGES[inner.type].label}` };
		}
		if (target === 'int' || target === 'fix') {
			const value = target === 'int' ? Math.floor(inner.value) : Math.trunc(inner.value);
			return { value, type: inner.type === 'byte' || inner.type === 'integer' || inner.type === 'long' ? inner.type : 'double' };
		}
		if (target === 'exp') {
			const value = Math.exp(inner.value);
			return inRange(value, 'double')
				? { value, type: 'double' }
				: { overflow: true, span, detail: `Exp(${showNumber(inner.value)}) exceeds the Double range` };
		}
		if (target === 'decimal') {
			// A Decimal holds up to 2^96 - 1: CDec("1E28") runs, CDec("1E30") and
			// CDec(1E+30) overflow (issue #218). The result is not typed further.
			return decimalFits(inner.value, shown)
				? undefined
				: { overflow: true, span, detail: `CDec(${shown}) does not fit Decimal` };
		}
		if (target === 'hex' || target === 'oct') {
			// Hex and Oct take a value that fits a Long (or a LongLong on 64-bit
			// for whole numbers; 1E+20 fits neither).
			return inRange(inner.value, 'long') || (Number.isInteger(inner.value) && Math.abs(inner.value) < 9.2e18)
				? undefined
				: { overflow: true, span, detail: `${callee === 'hex' ? 'Hex' : 'Oct'}(${inner.value}) takes a value outside the Long range` };
		}
		if (target === 'longlong' || target === 'longptr') {
			const value = bankersRound(inner.value);
			const exact = inner.exact ?? spelledWhole(shown) ?? (Number.isSafeInteger(value) ? BigInt(value) : undefined);
			return inRange(value, 'longlong', exact)
				? { value, type: 'longlong', ...(exact !== undefined ? { exact } : {}) }
				: { overflow: true, span, detail: `${CONVERSION_NAMES[callee]}(${shown}) does not fit ${target === 'longptr' ? 'a LongPtr, whose range is at most a LongLong\'s' : 'LongLong'}` };
		}
		const type = target as NumericType;
		const value = type === 'single' || type === 'double' || type === 'currency' || type === 'date' ? inner.value : bankersRound(inner.value);
		// A conversion to the type the constant already has is folded away:
		// `-CLng(&H80000000)` wraps, `-CLng(-2147483648#)` overflows (issue #235).
		const constant = inner.constant && inner.type === type ? { constant: true } : {};
		return inRange(value, type)
			? { value, type, ...constant }
			: { overflow: true, span, detail: `${CONVERSION_NAMES[callee]}(${shown}) does not fit ${RANGES[type].label}` };
	}

	private combine(left: Typed, right: Typed, op: string, from: number, to: number): Folded {
		const folded = this.combineValues(left, right, op, from, to);
		// Both halves folded while compiling: so is the result.
		return folded && !isOverflow(folded) && left.constant && right.constant ? { ...folded, constant: true } : folded;
	}

	private combineValues(left: Typed, right: Typed, op: string, from: number, to: number): Folded {
		const span = this.span(from, to);
		// `1 / 0`, `1 \ 0.4` and `1 Mod False` divide by zero; outside a Const
		// that is division-by-zero's to report.
		const divisor = op === '/' ? right.value : op === '\\' || op === 'mod' ? bankersRound(right.value) : undefined;
		if (divisor === 0) {
			this.divisionByZero?.(span);
			return undefined;
		}
		if ((left.type === 'longlong' || right.type === 'longlong') && op !== '/' && op !== '^') {
			return combineLongLong(left, right, op, span);
		}
		let type: NumericType;
		let value: number;
		switch (op) {
			case '/':
				if (right.value === 0) {
					return undefined; // division by zero is another rule's
				}
				type = left.type === 'currency' || right.type === 'currency' ? 'currency' : 'double';
				value = left.value / right.value;
				break;
			case '^':
				type = 'double';
				value = Math.pow(left.value, right.value);
				if (Number.isNaN(value) || (left.value === 0 && right.value < 0)) {
					// 0 ^ -1 raises 5, Invalid procedure call: runtime-value-out-of-range reports it.
					return undefined;
				}
				break;
			case '\\':
			case 'mod': {
				type = arithmeticResultType(left.type, right.type, op);
				const a = bankersRound(left.value);
				const b = bankersRound(right.value);
				if (b === 0) {
					return undefined;
				}
				value = op === 'mod' ? a % b : Math.trunc(a / b);
				break;
			}
			default:
				type = arithmeticResultType(left.type, right.type, op);
				value = op === '+' ? left.value + right.value : op === '-' ? left.value - right.value : left.value * right.value;
				break;
		}
		if (!inRange(value, type)) {
			const result = type === 'date'
				? `falls ${value > 0 ? 'after 12/31/9999' : 'before 1/1/100'}, outside the Date range`
				: `is ${showNumber(value)}, outside the ${RANGES[type].label} range`;
			return {
				overflow: true,
				span,
				detail: `${describe(left)} ${op === 'mod' ? 'Mod' : op} ${describe(right)} ${result}`,
			};
		}
		return { value, type };
	}
}

/** A whole number spelled in digits, as a conversion's string or literal argument shows it. */
function spelledWhole(shown: string): bigint | undefined {
	const digits = /^"?\s*([-+]?)(\d+)\s*"?$/.exec(shown);
	return digits ? (digits[1] === '-' ? -BigInt(digits[2]) : BigInt(digits[2])) : undefined;
}

const WHOLE_TYPES: ReadonlySet<NumericType> = new Set(['byte', 'integer', 'long', 'longlong']);

/** A whole-number operand's exact value, when it has one. */
function exactOf(typed: Typed): bigint | undefined {
	if (typed.exact !== undefined) {
		return typed.exact;
	}
	return WHOLE_TYPES.has(typed.type) && Number.isSafeInteger(typed.value) ? BigInt(typed.value) : undefined;
}

/**
 * `+ - * \ Mod` with a LongLong operand. With another whole number the
 * result is a LongLong, folded exactly; with a Single or Double it is a
 * Double. With a Currency or a Date the result type is not modelled, and
 * nothing is judged.
 */
function combineLongLong(left: Typed, right: Typed, op: string, span: Span): Folded {
	const other = left.type === 'longlong' ? right.type : left.type;
	if (other === 'single' || other === 'double') {
		if (op === '\\' || op === 'mod') {
			return undefined;
		}
		const value = op === '+' ? left.value + right.value : op === '-' ? left.value - right.value : left.value * right.value;
		return inRange(value, 'double') ? { value, type: 'double' } : { overflow: true, span, detail: `${describe(left)} ${op} ${describe(right)} is outside the Double range` };
	}
	const a = exactOf(left);
	const b = exactOf(right);
	if (a === undefined || b === undefined) {
		return undefined;
	}
	let exact: bigint;
	if (op === '\\' || op === 'mod') {
		if (b === 0n) {
			return undefined; // division by zero is another rule's
		}
		exact = op === 'mod' ? a % b : a / b;
	} else {
		exact = op === '+' ? a + b : op === '-' ? a - b : a * b;
	}
	if (!inRange(0, 'longlong', exact)) {
		return { overflow: true, span, detail: `${describe(left)} ${op === 'mod' ? 'Mod' : op} ${describe(right)} is ${exact}, outside the LongLong range` };
	}
	return { value: Number(exact), type: 'longlong', exact };
}

/**
 * `Not x`, the bitwise complement in x's type (issue #235, measured in Excel
 * 16.0): an Integer or a Long stays itself, `Not 32767` is -32768 and
 * `Not 0` is -1; a Single, Double, Currency or Date is rounded to a Long
 * first, so `Not 32768!` is -32769. A Byte stays a Byte.
 */
function notOf(operand: Typed, span: Span): Folded {
	const constant = operand.constant ? { constant: true } : {};
	if (operand.type === 'byte') {
		return { value: 255 - operand.value, type: 'byte', ...constant };
	}
	if (operand.type === 'integer' || operand.type === 'long') {
		return { value: -operand.value - 1, type: operand.type, ...constant };
	}
	if (operand.type === 'longlong') {
		const exact = exactOf(operand);
		return exact === undefined ? undefined : { value: Number(-exact - 1n), type: 'longlong', exact: -exact - 1n, ...constant };
	}
	const whole = bankersRound(operand.value);
	if (!inRange(whole, 'long')) {
		return { overflow: true, span, detail: `Not ${showNumber(operand.value)} rounds to ${showNumber(whole)}, outside the Long range` };
	}
	return { value: -whole - 1, type: 'long', ...constant };
}

function describe(typed: Typed): string {
	if (typed.type === 'date') {
		// Before serial 0 the fraction counts forward from the day's start
		// too: -1.25 is 12/29/1899 6:00:00 AM.
		const day = Math.trunc(typed.value);
		const seconds = Math.round(Math.abs(typed.value - day) * 86400);
		const at = new Date(DATE_EPOCH_MS + day * DAY_MS);
		const date = `${at.getUTCMonth() + 1}/${at.getUTCDate()}/${at.getUTCFullYear()}`;
		if (seconds === 0) {
			return `#${date}# (Date)`;
		}
		const hour = Math.floor(seconds / 3600);
		const clock = `${hour % 12 === 0 ? 12 : hour % 12}:${String(Math.floor(seconds / 60) % 60).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
		return `#${date} ${clock} ${hour < 12 ? 'AM' : 'PM'}# (Date)`;
	}
	return `${typed.exact !== undefined ? String(typed.exact) : showNumber(typed.value)} (${RANGES[typed.type].label})`;
}

/** A type's range as the message shows it; a LongLong's ends print exactly. */
function rangeText(type: NumericType): string {
	return type === 'longlong'
		? '-9223372036854775808 to 9223372036854775807'
		: `${RANGES[type].min} to ${RANGES[type].max}`;
}

/** A folded value converted to a target type, rounded as VBA stores it. */
function storedValue(folded: Typed, target: NumericType): { value: number; exact?: bigint } {
	if (target === 'single' || target === 'double' || target === 'currency' || target === 'date') {
		return { value: folded.value };
	}
	const value = bankersRound(folded.value);
	return target === 'longlong' && folded.exact !== undefined ? { value, exact: folded.exact } : { value };
}

/** A value as VBA would print it: whole numbers plain, huge or fractional ones in E notation. */
function showNumber(value: number): string {
	if (!Number.isFinite(value)) {
		return value > 0 ? 'a value past the Double maximum' : 'a value past the Double minimum';
	}
	if (Number.isInteger(value) && Math.abs(value) < 1e15) {
		return String(value);
	}
	if (Math.abs(value) >= 1e15 || (Math.abs(value) < 1e-4 && value !== 0)) {
		return value.toExponential(4).replace(/\.?0+e/, 'E').replace(/e\+?/, 'E+').replace('E+-', 'E-');
	}
	return String(value);
}

function article(label: string): string {
	return /^[AEIOU]/.test(label) ? 'an' : 'a';
}

const CONVERSIONS: ReadonlyMap<string, NumericType | 'abs' | 'int' | 'fix' | 'exp' | 'hex' | 'oct' | 'decimal' | 'longptr'> = new Map([
	['cbyte', 'byte'], ['cint', 'integer'], ['clng', 'long'], ['csng', 'single'], ['cdbl', 'double'],
	['ccur', 'currency'], ['cdate', 'date'], ['abs', 'abs'], ['int', 'int'], ['fix', 'fix'], ['exp', 'exp'],
	['hex', 'hex'], ['oct', 'oct'], ['cdec', 'decimal'], ['clnglng', 'longlong'], ['clngptr', 'longptr'],
]);

/** 2^96, one past the largest Decimal. */
const DECIMAL_LIMIT = 79228162514264337593543950336n;

/**
 * Whether a Decimal holds the value. A double cannot tell 2^96 - 1 from 2^96,
 * so a whole number spelled in digits is compared exactly.
 */
function decimalFits(value: number, shown: string): boolean {
	const digits = /^"?\s*([-+]?)(\d+)\s*"?$/.exec(shown);
	if (digits) {
		return BigInt(digits[2]) < DECIMAL_LIMIT;
	}
	return Math.abs(value) < 7.9228162514264337e28;
}

const CONVERSION_NAMES: Readonly<Record<string, string>> = {
	cbyte: 'CByte', cint: 'CInt', clng: 'CLng', csng: 'CSng', cdbl: 'CDbl', ccur: 'CCur', cdate: 'CDate',
	clnglng: 'CLngLng', clngptr: 'CLngPtr',
};

function numericTypeOf(asType: string | undefined): NumericType | undefined {
	const normalized = normalizeType(asType);
	if (!normalized) {
		return undefined;
	}
	if (normalized in RANGES) {
		return normalized as NumericType;
	}
	// A LongPtr is a Long or a LongLong by platform: past LongLong's range it
	// overflows on both, and inside it is not judged.
	return normalized === 'longptr' ? 'longlong' : undefined;
}

/**
 * The Consts a procedure can name, folded with their declared or natural
 * type: `Private Const HOURS As Integer = 24` is an Integer 24, and
 * `Const K = 40000` a Long. A Const the folder cannot fold is left out.
 */
/**
 * Folded values of the Consts in `candidates`, layered over `base`: a name
 * declared in both takes the candidate's value (a procedure's own Const wins
 * over the module's), and a candidate's value may refer to a base constant.
 * The module and project layer is folded once per pass and each procedure
 * adds only its own Consts on top; folding the whole project's constants
 * again for every procedure was 15% of a large module's analysis (issue #139).
 */
function constantLookup(
	base: ReadonlyMap<string, Typed>,
	candidates: readonly VbaSymbol[],
): ReadonlyMap<string, Typed> {
	const pending = new Map<string, VbaSymbol>();
	for (const symbol of candidates) {
		if (symbol.kind === 'constant' && symbol.defaultRaw !== undefined) {
			pending.set(symbol.name.toLowerCase(), symbol);
		}
	}
	if (pending.size === 0) {
		return base;
	}
	const folded = new Map<string, Typed>();
	const resolving = new Set<string>();
	const resolve = (lower: string): Typed | undefined => {
		const symbol = pending.get(lower);
		if (!symbol) {
			return base.get(lower);
		}
		if (folded.has(lower)) {
			return folded.get(lower);
		}
		if (resolving.has(lower)) {
			return undefined;
		}
		resolving.add(lower);
		const toks = rawExpressionTokens(symbol.defaultRaw!)
			.filter((tok) => tok.kind !== 'comment');
		const value = new TypedFolder(toks, 0, resolve).fold();
		resolving.delete(lower);
		if (value === undefined || isOverflow(value)) {
			return undefined;
		}
		const declared = numericTypeOf(symbol.asType);
		const kept = declared ? storedValue(value, declared) : undefined;
		const typed: Typed = declared && kept
			? { value: kept.value, type: declared, ...(kept.exact !== undefined ? { exact: kept.exact } : {}), constant: true }
			: value;
		if (declared && kept && !inRange(kept.value, declared, kept.exact)) {
			return undefined;
		}
		folded.set(lower, typed);
		return typed;
	};
	const out = new Map<string, Typed>(base);
	for (const lower of pending.keys()) {
		const typed = resolve(lower);
		// A candidate that did not fold still shadows the base name.
		if (typed) {
			out.set(lower, typed);
		} else {
			out.delete(lower);
		}
	}
	return out;
}

/**
 * Host members whose value is fixed: Excel's `Rows.Count` is 1048576 and
 * `Columns.Count` 16384 on every worksheet since Excel 2007 (a Long each).
 * Absent model means Excel (issue #28).
 */
function hostConstantValues(hostModel: HostObjectModel | undefined): ReadonlyMap<string, Typed> {
	if (hostModel && hostModel.hostName !== undefined && hostModel.hostName !== 'Excel') {
		return new Map();
	}
	return new Map<string, Typed>([
		['rows.count', { value: 1048576, type: 'long' }],
		['columns.count', { value: 16384, type: 'long' }],
	]);
}

export function checkOverflow(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	hostModel: HostObjectModel | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Module-level Consts: a folded overflow is the compile error.
	const moduleConstants = constantLookup(
		new Map(),
		[...(projectVisibleSymbols ?? []), ...(symbols.root.children ?? [])],
	);
	checkConstDeclarations(source, mod.members.filter((m): m is VariableGroupNode => m.kind === 'VariableGroup'), moduleConstants, activity, push);
	const hostValues = hostConstantValues(hostModel);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const constants = constantLookup(moduleConstants, procedureSymbolFor(symbols, member)?.children ?? []);
		const env = typeEnvironmentFor(symbols, member);
		const known = knownLocalLiteralValues(source, member, symbols, activity);
		// Values a straight run of top-level statements has just stored:
		// `i = 32767` followed by `i = i + 1`.
		const justAssigned = new Map<string, Typed>();
		const names: NameLookup = (lower) => {
			const constant = constants.get(lower);
			if (constant) {
				return constant;
			}
			const recent = justAssigned.get(lower);
			if (recent) {
				return recent;
			}
			const local = known.get(lower);
			const type = numericTypeOf(env.get(lower));
			if (local?.kind === 'number' && type) {
				return { value: local.value as number, type };
			}
			return lower.includes('.') && !env.has(lower.slice(0, lower.indexOf('.'))) ? hostValues.get(lower) : undefined;
		};
		const groups: VariableGroupNode[] = [];
		forEachVariableGroup(member.body, (group) => { groups.push(group); }, activity);
		checkConstDeclarations(source, groups, constants, activity, push);
		checkProcedureBody(source, member, env, names, justAssigned, activity, push);
	}
}

function checkConstDeclarations(
	source: string,
	groups: readonly VariableGroupNode[],
	constants: ReadonlyMap<string, Typed>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const group of groups) {
		if (!group.isConst || activity?.isInactive(group.span)) {
			continue;
		}
		for (const decl of group.declarations) {
			if (decl.defaultRaw === undefined) {
				continue;
			}
			const toks = statementTokens(source, decl.span);
			const eq = toks.findIndex((tok) => tok.rawText === '=');
			if (eq < 0) {
				continue;
			}
			const value = toks.slice(eq + 1).filter((tok) => tok.kind !== 'comment');
			let divided: Span | undefined;
			const folded = new TypedFolder(value, decl.span.start, (lower) => constants.get(lower), (span) => {
				divided ??= span;
			}).fold();
			if (divided) {
				push('constEvaluationError', `Const '${decl.name}' divides by zero while it is evaluated. This is a VBE compile error: Division by zero.`, divided);
				continue;
			}
			const refused = stringConstRefusal(value, decl.asType, decl.span.start);
			if (refused) {
				push(refused.overflow ? 'constOverflow' : 'constEvaluationError', `Const '${decl.name}': ${refused.detail}. This is a VBE compile error: ${refused.overflow ? 'Overflow' : 'Type mismatch'}.`, refused.span);
				continue;
			}
			if (isOverflow(folded)) {
				push('constOverflow', `Const '${decl.name}' overflows while it is evaluated: ${folded.detail}. This is a VBE compile error: Overflow.`, folded.span);
				continue;
			}
			const declared = numericTypeOf(decl.asType);
			const kept = folded && declared ? storedValue(folded, declared) : undefined;
			if (folded && declared && kept && !inRange(kept.value, declared, kept.exact)) {
				const label = normalizeType(decl.asType) === 'longptr' ? 'LongPtr' : RANGES[declared].label;
				const shown = folded.exact !== undefined ? String(folded.exact) : String(folded.value);
				push('constOverflow', `Const '${decl.name}' is declared As ${label} but its value ${shown} is outside ${label === 'LongPtr' ? "even a LongLong's range" : 'that range'}. This is a VBE compile error: Overflow.`, { start: decl.span.start + value[0].start, end: decl.span.start + value[value.length - 1].end });
			}
		}
	}
}

/**
 * A Const whose value is a string the declared type cannot take (issue #235,
 * measured in Excel 16.0): `As Long = "abc"`, `As Long = ""`,
 * `As Boolean = "abc"`, `= -"abc"` and `Not ""` are "Type mismatch", and
 * `As Integer = "40000"` is "Overflow". `As Long = "12"`, `As Boolean =
 * "True"` and `-"12"` compile. A string with a digit in it is read by the
 * locale, so only one with none is judged a mismatch.
 */
function stringConstRefusal(
	value: readonly VbaToken[],
	asType: string | undefined,
	base: number,
): { detail: string; span: Span; overflow?: boolean } | undefined {
	const operator = value.length === 2 && (value[0].rawText === '-' || tokenText(value[0]) === 'not') ? value[0] : undefined;
	const literal = value[operator ? 1 : 0];
	if (value.length !== (operator ? 2 : 1) || literal?.kind !== 'stringLiteral') {
		return undefined;
	}
	const span = { start: base + value[0].start, end: base + literal.end };
	const text = stringLiteralValue(literal.rawText);
	const hasDigit = /\d/.test(text);
	if (operator) {
		return hasDigit ? undefined : { detail: `'${operator.rawText}' cannot work on ${literal.rawText}, which is no number`, span };
	}
	const declared = normalizeType(asType);
	if (declared === 'boolean') {
		return hasDigit || /^\s*(true|false)\s*$/i.test(text) ? undefined : { detail: `${literal.rawText} is no Boolean`, span };
	}
	const numeric = numericTypeOf(asType);
	if (!numeric || numeric === 'date') {
		return undefined;
	}
	if (!hasDigit) {
		return { detail: `${literal.rawText} is no number, so it cannot be ${article(RANGES[numeric].label)} ${RANGES[numeric].label}`, span };
	}
	const read = numberInString(text);
	if (read === 'overflow') {
		return { detail: `${literal.rawText} spells a number past the Double range`, span, overflow: true };
	}
	if (read && WHOLE_TYPES.has(numeric) && !inRange(bankersRound(read.value), numeric)) {
		return { detail: `${literal.rawText} is outside the ${RANGES[numeric].label} range`, span, overflow: true };
	}
	return undefined;
}

function checkProcedureBody(
	source: string,
	proc: ProcedureNode,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	justAssigned: Map<string, Typed>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	// Every name a block mentions, its own lines included: `For i = ...` and
	// `If Store(k, n) Then` change what they name as well (issue #237).
	const touchedIn = (node: BodyNode): Set<string> => namesIn(source, node.span);
	const forget = (touched: ReadonlySet<string>): void => {
		for (const lower of touched) {
			justAssigned.delete(lower);
		}
	};
	// `loopTouched`: names an enclosing loop changes, which a block nested in
	// it forgets, since it may run on a later pass.
	const visit = (body: readonly BodyNode[], loopTouched: ReadonlySet<string>): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'ForBlock') {
				checkForCounter(source, node, env, names, push);
			}
			if ('body' in node && Array.isArray(node.body)) {
				// Its header line is evaluated as it is entered: `For i = 1 To
				// CInt(40000)`, `Select Case CInt(40000)` (issue #233). Its body
				// is entered with what is known then: each If arm and each Case
				// from there, a loop's body as its first pass runs it. After it,
				// only what it never names is still known (issue #237).
				const { before, after } = blockHeaderStatements(source, node);
				if (before) {
					checkStatement(source, before.span, env, names, push);
				}
				const touched = touchedIn(node);
				forget(loopTouched);
				// Its own lines run first: `If Store(k, n) Then` changes n.
				for (const header of blockHeaderLeaves(source, node)) {
					forget(namesIn(source, header.span));
				}
				const entry = new Map(justAssigned);
				const restore = (): void => {
					justAssigned.clear();
					for (const [lower, value] of entry) {
						justAssigned.set(lower, value);
					}
				};
				if (node.kind === 'IfBlock') {
					for (const branch of node.branches) {
						restore();
						visit(branch.body, loopTouched);
					}
				} else if (node.kind === 'SelectBlock') {
					for (const arm of selectArms(source, node.body)) {
						restore();
						visit(arm, loopTouched);
					}
				} else {
					visit(node.body as BodyNode[], isLoopBlock(node) ? new Set([...loopTouched, ...touched]) : loopTouched);
				}
				restore();
				forget(touched);
				if (after) {
					checkStatement(source, after.span, env, names, push);
				}
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			const spans = statementAndBranchSpans(node);
			const straightLine = spans.length === 1 && !(node.kind === 'Statement' && node.singleLineIfBranches);
			if (!straightLine) {
				justAssigned.clear();
			}
			for (const span of spans) {
				const stored = checkStatement(source, span, env, names, push);
				if (!straightLine) {
					continue;
				}
				// Any other mention of a tracked name (a ByRef pass, a label a
				// GoTo could reach) ends what is known about it.
				const toks = statementTokens(source, span);
				if (statementLabelDeclaration(source, span) || tokenText(toks[firstExecutableTokenIndex(toks)]) === 'gosub') {
					justAssigned.clear();
					continue;
				}
				for (const tok of toks) {
					const lower = tokenName(tok)?.toLowerCase();
					if (lower && justAssigned.has(lower)) {
						justAssigned.delete(lower);
					}
				}
				if (stored) {
					justAssigned.set(stored.name, stored.value);
				}
			}
		}
	};
	visit(proc.body, new Set());
}

/** The value a bare assignment provably stores, when the rule can tell. */
function checkStatement(
	source: string,
	span: Span,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	push: PushFn,
): { name: string; value: Typed } | undefined {
	const toks = statementTokens(source, span);
	const first = firstExecutableTokenIndex(toks);
	const head = tokenText(toks[first]);
	if (head === 'const' || head === 'dim' || head === 'static' || head === 'redim') {
		return undefined;
	}
	const reported = new Set<string>();
	const report = (folded: Overflow): void => {
		const key = `${folded.span.start}:${folded.span.end}`;
		if (!reported.has(key)) {
			reported.add(key);
			push('arithmeticOverflow', `${folded.detail}. This will raise Run-time error '6': Overflow.`, folded.span);
		}
	};
	let stored: { name: string; value: Typed } | undefined;
	const bare = bareAssignmentTarget(source, span);
	if (bare) {
		const value = bare.valueTokens.filter((tok) => tok.kind !== 'comment');
		const folded = new TypedFolder(value, span.start, names).fold();
		const target = numericTypeOf(env.get(bare.name.toLowerCase()));
		if (isOverflow(folded)) {
			report(folded);
		} else if (folded && target) {
			const kept = storedValue(folded, target);
			if (!inRange(kept.value, target, kept.exact)) {
				const shown = kept.exact !== undefined ? String(kept.exact) : showNumber(kept.value);
				const rounded = kept.value !== folded.value ? ` (${showNumber(folded.value)} rounds to ${showNumber(kept.value)})` : '';
				const label = normalizeType(env.get(bare.name.toLowerCase())) === 'longptr' ? 'LongPtr, which holds no more than a LongLong' : RANGES[target].label;
				push('arithmeticOverflow', `Assignment to '${bare.name}' stores ${shown}${rounded} in ${article(label)} ${label}, whose range is ${rangeText(target)}. This will raise Run-time error '6': Overflow.`, {
					start: span.start + value[0].start,
					end: span.start + value[value.length - 1].end,
				});
				return undefined;
			}
			stored = { name: bare.name.toLowerCase(), value: { value: kept.value, type: target, ...(kept.exact !== undefined ? { exact: kept.exact } : {}) } };
		}
	}
	// Every other part the statement evaluates on its own: a call's
	// arguments, an operand of & or a comparison, an array index, a
	// conversion anywhere (issue #232). `Main = CStr(CInt(40000))` and
	// `IIf(True, 0, CInt(40000))` overflow as `Main = CInt(40000)` does.
	checkParts(toks, span.start, names, report);
	return stored;
}

/**
 * Keywords that stand between separately evaluated parts: the operators the
 * folder does not fold, and the statement words around an expression. A
 * keyword that is an operand or a function, `Date` or `CInt`, is not one:
 * `Date - 32767% - 2%` is Date arithmetic, and its tail is no Integer sum.
 */
const PART_KEYWORDS: ReadonlySet<string> = new Set([
	'and', 'or', 'xor', 'eqv', 'imp', 'not', 'like', 'is', 'if', 'then', 'else', 'elseif',
	'to', 'step', 'print', 'call', 'set', 'let', 'return', 'while', 'until', 'case', 'with',
	'select', 'each', 'in', 'goto', 'gosub', 'on',
]);

/**
 * Tokens that end one separately evaluated part of an expression: a comma,
 * `:=`, every operator the folder does not fold (&, comparisons), and the
 * keywords above, so `Debug.Print 200 * 200` leaves `200 * 200` to fold.
 */
function endsPart(tok: VbaToken): boolean {
	if (tok.rawText === ',' || tok.rawText === ';' || tok.rawText === ':=') {
		return true;
	}
	if (tok.kind === 'operator') {
		return !['+', '-', '*', '/', '\\', '^', '(', ')', '.', '!'].includes(tok.rawText);
	}
	return tok.kind === 'keyword' && PART_KEYWORDS.has(tokenText(tok));
}

/**
 * Folds each part of `toks` that is evaluated on its own; a part that does not
 * fold is searched for parenthesized parts that do, so an argument nested at
 * any depth is reached.
 */
function checkParts(toks: readonly VbaToken[], base: number, names: NameLookup, report: (folded: Overflow) => void): void {
	let from = 0;
	const part = (to: number): void => {
		const piece = toks.slice(from, to).filter((tok) => tok.kind !== 'comment');
		from = to + 1;
		if (piece.length === 0) {
			return;
		}
		const folded = new TypedFolder(piece, base, names).fold();
		if (isOverflow(folded)) {
			report(folded);
			return;
		}
		if (folded !== undefined) {
			return;
		}
		for (let i = 0; i < piece.length; i++) {
			if (piece[i].rawText !== '(') {
				continue;
			}
			const close = matchParenFrom(piece, i);
			if (close < 0) {
				return;
			}
			// A conversion folds with its call: `CInt(40000)` overflows though 40000 does not.
			const callee = tokenText(piece[i - 1]);
			if (CONVERSIONS.has(callee) && isBareOrVbaQualifiedIntrinsicCall(piece, i - 1)) {
				const start = piece[i - 2]?.rawText === '.' ? i - 3 : i - 1;
				const call = new TypedFolder(piece.slice(start, close + 1), base, names).fold();
				if (isOverflow(call)) {
					report(call);
					i = close;
					continue;
				}
			}
			checkParts(piece.slice(i + 1, close), base, names, report);
			i = close;
		}
	};
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (depth === 0 && endsPart(toks[i])) {
			part(i);
		}
	}
	part(toks.length);
}

/**
 * `For i = 1 To 32767` with i an Integer: the counter is incremented past its
 * last value before the exit test, and the increment overflows (measured:
 * `To 32767` raises, `To 32766` runs; `For b = 0 To 255` raises for a Byte).
 */
/**
 * True when a statement in the loop's body can leave the loop before the
 * counter passes its type: `Exit For` (not one belonging to a nested For),
 * `Exit Sub`/`Function`/`Property`, `GoTo`, or `End` (issue #145). Such a
 * loop's overflow is not proved, so it is not reported.
 */
function bodyMayLeaveLoop(source: string, body: readonly BodyNode[]): boolean {
	const LEAVES = new Set(['for', 'sub', 'function', 'property']);
	const visit = (nodes: readonly BodyNode[], insideNestedFor: boolean): boolean => {
		for (const node of nodes) {
			if (isLeafStatement(node)) {
				const toks = statementTokensAfterLeadingLabel(source, node.span);
				for (let i = 0; i < toks.length; i++) {
					const word = tokenText(toks[i]);
					if (word === 'goto' || (word === 'end' && toks.length === 1)) {
						return true;
					}
					if (word === 'exit') {
						const target = tokenText(toks[i + 1]);
						if (LEAVES.has(target) && (target !== 'for' || !insideNestedFor)) {
							return true;
						}
					}
				}
			} else if ('body' in node && Array.isArray(node.body)) {
				if (visit(node.body as BodyNode[], insideNestedFor || node.kind === 'ForBlock')) {
					return true;
				}
			}
		}
		return false;
	};
	return visit(body, false);
}

function checkForCounter(
	source: string,
	node: ForBlockNode,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	push: PushFn,
): void {
	if (node.each || !node.controlVariable) {
		return;
	}
	const type = numericTypeOf(env.get(node.controlVariable.toLowerCase()));
	if (!type || type === 'single' || type === 'double' || type === 'currency' || type === 'date') {
		return;
	}
	const headerEnd = source.indexOf('\n', node.span.start);
	const header = { start: node.span.start, end: headerEnd < 0 ? node.span.end : Math.min(headerEnd, node.span.end) };
	const toks = statementTokensAfterLeadingLabel(source, header);
	const to = toks.findIndex((tok) => tokenText(tok) === 'to');
	if (to < 0) {
		return;
	}
	const step = toks.findIndex((tok) => tokenText(tok) === 'step');
	const limitToks = toks.slice(to + 1, step > 0 ? step : toks.length).filter((tok) => tok.kind !== 'comment');
	const limit = new TypedFolder(limitToks, header.start, names).fold();
	const stepValue = step > 0 ? new TypedFolder(toks.slice(step + 1).filter((tok) => tok.kind !== 'comment'), header.start, names).fold() : { value: 1, type: 'integer' as NumericType };
	if (!limit || isOverflow(limit) || !stepValue || isOverflow(stepValue) || stepValue.value === 0 || !Number.isInteger(stepValue.value)) {
		return;
	}
	// The counter's last value is the start plus whole steps: `For i = 1 To
	// 32766 Step 2` ends at 32765 and runs (measured). A step of 1 or -1 ends
	// at the limit whatever the start.
	const eq = toks.findIndex((tok) => tok.rawText === '=');
	const start = eq > 0 ? new TypedFolder(toks.slice(eq + 1, to).filter((tok) => tok.kind !== 'comment'), header.start, names).fold() : undefined;
	let last: number;
	if (Math.abs(stepValue.value) === 1) {
		last = limit.value;
	} else if (start && !isOverflow(start) && Number.isInteger(start.value)) {
		const passes = Math.floor((limit.value - start.value) / stepValue.value);
		if (passes < 0) {
			return; // the loop body never runs and the counter stays at the start
		}
		last = start.value + passes * stepValue.value;
	} else {
		return;
	}
	const range = RANGES[type];
	let lastShown = String(last);
	let overflows = stepValue.value > 0 ? last + stepValue.value > range.max : last + stepValue.value < range.min;
	if (type === 'longlong') {
		// Exactly: a double cannot tell 2^63 - 1 from 2^63 (issue #232).
		const exactLast = Math.abs(stepValue.value) === 1 ? exactOf(limit) : undefined;
		const exactStep = exactOf(stepValue);
		if (exactLast === undefined || exactStep === undefined) {
			return;
		}
		overflows = !inRange(0, 'longlong', exactLast + exactStep);
		lastShown = String(exactLast);
	}
	if (!overflows || bodyMayLeaveLoop(source, node.body)) {
		return;
	}
	push(
		'forCounterOverflow',
		`Counter '${node.controlVariable}' is ${range.label}; after its last pass at ${lastShown} the loop adds ${stepValue.value}, which does not fit. This will raise Run-time error '6': Overflow.`,
		{ start: header.start + limitToks[0].start, end: header.start + limitToks[limitToks.length - 1].end },
	);
}

