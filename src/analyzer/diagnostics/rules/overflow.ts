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
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ForBlockNode, ModuleNode, ProcedureNode, Span, VariableGroupNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import { jumpTargetLabelDeclaration } from '../../flow/procedureLabels';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { bankersRound, bodyMayLeaveLoop, isBareOrVbaQualifiedIntrinsicCall, namesIn } from './shared';
import { functionResultNamed, knownFunctionResults } from '../functionResults';
import { checkEachCounterPass, loopCountersAt } from '../loopCounters';
import { blockHeaderLeaves, isLoopBlock, selectArms } from '../blockHeaders';
import { fieldChain, moduleTypes, variableRoot, variableSymbolIn, type ModuleTypes } from '../typeFields';
import {
	buildModuleTypeSignatures,
	knownLocalLiteralValuesAt,
	type KnownLocalValue,
	normalizeType,
	stringLiteralValue,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	blockHeaderStatements,
	firstExecutableTokenIndex,
	forEachStatement,
	forEachVariableGroup,
	matchParenFrom,
	statementAndBranchSpans,
	rawExpressionTokens,
	setAssignmentTarget,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';
import { trackedLocalsNamedWhole } from '../dataflow';

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
	/**
	 * A Boolean, which folds as the Integer -1 or 0 but goes into a Byte as
	 * 255 or 0: `b = True` and `CByte(True)` store 255 (issue #326).
	 */
	boolean?: boolean;
	/**
	 * A Date past the Date range that `number + Date` or `number - Date`
	 * made without raising (issue #330): the expression, as the message
	 * names it. Most uses of it raise (issue #405).
	 */
	pastDate?: string;
	/**
	 * Held in a Variant, whose arithmetic widens the type instead of
	 * overflowing: 32767 + 1 is the Long 32768 (issue #480, measured in
	 * Excel 16.0).
	 */
	variant?: boolean;
	/** A Currency value in ten-thousandths, exactly (issue #494). */
	scaled?: bigint;
	/** A whole Decimal's exact value, from CDec (issue #502): the type says Double. */
	decimal?: bigint;
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
	// A Date with a Currency is a Date for + and -, and a Double for *: c + t
	// is VarType 7, t * c VarType 5 (issue #409, measured in Excel 16.0).
	if ((a === 'date' || b === 'date') && (op === '+' || op === '-' || op === '*')) {
		return op === '*' || (a === 'date' && b === 'date' && op === '-') ? 'double' : 'date';
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
		return word === 'true' ? { value: -1, type: 'integer', boolean: true } : word === 'false' ? { value: 0, type: 'integer', boolean: true } : undefined;
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
		const scaled = suffix === '@' ? currencyScaled(raw.replace(/@$/, '')) : undefined;
		return { value, type: suffix === '!' ? 'single' : suffix === '@' ? 'currency' : 'double', ...(scaled !== undefined ? { scaled } : {}) };
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

/**
 * What a name means to the folder: a typed value, or nothing. `declares`,
 * where given, says the procedure or module declares the name, so a host
 * global of that spelling is hidden.
 */
type NameLookup = ((lower: string) => Typed | undefined) & {
	declares?: (lower: string) => boolean;
	/** A local that holds a whole sheet's Cells: `Set r = Cells` (issue #278). */
	wholeSheetCells?: (lower: string) => boolean;
};

/** The operators that read both sides as numbers: arithmetic and comparison. */
const BINARY_ON_NUMBERS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '=', '<>', '<', '>', '<=', '>=',
	// The logical operators convert a String to a number too: `"" And 255`
	// is a Type mismatch (issue #494, measured in Excel 16.0).
	'and', 'or', 'xor', 'eqv', 'imp']);

/** The operators a String beside a number converts under in a Const (issue #494). */
const ARITHMETIC_BESIDE: ReadonlySet<string> = new Set(['+', '-', '*', '/']);

/** Logical precedence; equal precedence splits at the last operator. */
const LOGICAL_PRECEDENCE: ReadonlyMap<string, number> = new Map([
	['imp', 0], ['eqv', 1], ['xor', 2], ['or', 3], ['and', 4],
]);

/** What TypedFolder.logical answers for an expression with no logical operator. */
const NOT_LOGICAL = Symbol('not logical');

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
		const logical = this.logical();
		if (logical !== NOT_LOGICAL) {
			return logical;
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

	/**
	 * `a And b`, Or, Xor, Eqv and Imp, split at the last top-level one of the
	 * lowest precedence. Each operand is converted to a Long first, so one
	 * outside the Long range overflows: `1E10 And 1` is "Overflow" in a Const
	 * (issue #367, measured in Excel 16.0) and error 6 at run time (#323).
	 */
	private logical(): Folded | typeof NOT_LOGICAL {
		let depth = 0;
		let at = -1;
		let precedence = Infinity;
		let word = '';
		// Locate the last top-level operator of the lowest precedence in one
		// pass, including the common arithmetic-only case with no operator.
		for (let i = 0; i < this.toks.length; i++) {
			const tok = this.toks[i];
			depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
			if (depth !== 0 || i === 0 || tok.kind !== 'keyword') { continue; }
			const lower = tokenText(tok);
			const rank = LOGICAL_PRECEDENCE.get(lower);
			if (rank !== undefined && rank <= precedence) {
				at = i;
				precedence = rank;
				word = lower;
			}
		}
		if (at < 0) { return NOT_LOGICAL; }
		const left = this.stringOperand(this.toks.slice(0, at)) ?? new TypedFolder(this.toks.slice(0, at), this.base, this.names, this.divisionByZero).fold();
		if (left === undefined || isOverflow(left)) {
			return left;
		}
		const right = this.stringOperand(this.toks.slice(at + 1)) ?? new TypedFolder(this.toks.slice(at + 1), this.base, this.names, this.divisionByZero).fold();
		if (right === undefined || isOverflow(right)) {
			return right;
		}
		if (left.type === 'longlong' || right.type === 'longlong') {
			return undefined;
		}
		const span = this.span(0, this.toks.length - 1);
		const operands = [left, right].map((operand) => bankersRound(operand.value));
		const outside = [left, right].find((_, k) => !inRange(operands[k], 'long'));
		if (outside) {
			return { overflow: true, span, detail: `${describe(outside)} is outside the Long range that ${this.toks[at].rawText} converts its operands to` };
		}
		const [a, b] = operands;
		const value = word === 'and' ? a & b : word === 'or' ? a | b : word === 'xor' ? a ^ b : word === 'eqv' ? ~(a ^ b) : ~a | b;
		const small = (operand: Typed): boolean => operand.type === 'byte' || operand.type === 'integer';
		const type: NumericType = left.type === 'byte' && right.type === 'byte' ? 'byte' : small(left) && small(right) ? 'integer' : 'long';
		const kept = type === 'byte' ? value & 0xff : value;
		return {
			value: kept,
			type,
			...(left.constant && right.constant ? { constant: true } : {}),
			...(left.boolean && right.boolean ? { boolean: true } : {}),
		};
	}

	/**
	 * A string literal operand of an operator that converts to a number
	 * first, a logical one, `\` or Mod, read as the number it spells:
	 * `"3E9" Or 0` overflows the Long and `"1" \ False` divides by zero in a
	 * Const (issue #458, measured in Excel 16.0). Undefined for anything
	 * else, and for a string that spells no number.
	 */
	private stringOperand(toks: readonly VbaToken[]): Folded {
		const tok = toks[0];
		if (toks.length !== 1 || tok.kind !== 'stringLiteral') {
			return undefined;
		}
		const read = numberInString(stringLiteralValue(tok.rawText));
		const span = { start: this.base + tok.start, end: this.base + tok.end };
		return read === 'overflow'
			? { overflow: true, span, detail: `${tok.rawText} spells a number past the Double range` }
			: read === undefined ? undefined : { ...read, constant: true };
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
			// Negating a Byte gives an Integer: `-b` with b = 200 is -200
			// (issue #362, measured in Excel 16.0).
			const type: NumericType = operand.type === 'byte' ? 'integer' : operand.type;
			if (!inRange(value, type, exact)) {
				return { overflow: true, span: this.span(start, this.index - 1), detail: `Negating ${exact !== undefined ? String(operand.exact) : showNumber(operand.value)} gives ${exact !== undefined ? String(exact) : showNumber(-operand.value)}, which does not fit ${RANGES[type].label}` };
			}
			return { value, type, ...(exact !== undefined ? { exact } : {}), ...(operand.constant ? { constant: true } : {}) };
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
		const date = this.currentDate() ?? this.dateSerial();
		if (date) {
			return date;
		}
		// `"1" \ False`: the string is an operand of `\` or Mod itself, not
		// of a `*`, `/` or `^` that binds tighter.
		const word = (at: number): string => (this.toks[at]?.kind === 'operator' ? this.toks[at].rawText : tokenText(this.toks[at]));
		const before = word(this.index - 1);
		const after = word(this.index + 1);
		if (tok.kind === 'stringLiteral' && (before === '\\' || before === 'mod' || after === '\\' || after === 'mod')
			&& !['*', '/', '^'].includes(after) && !['*', '/', '^', '-', '+'].includes(before)) {
			this.index++;
			return this.stringOperand([tok]);
		}
		// `^` takes both operands as Doubles, a string literal too:
		// `"12" ^ 32767` overflows (issue #331, measured in Excel 16.0).
		if (tok.kind === 'stringLiteral' && (before === '^' || after === '^')) {
			const read = numberInString(stringLiteralValue(tok.rawText));
			if (read === 'overflow') {
				return { overflow: true, span: this.span(this.index, this.index), detail: `${tok.rawText} spells a number past the Double range` };
			}
			if (read !== undefined) {
				this.index++;
				return { value: read.value, type: 'double' };
			}
		}
		// In a Const, a string beside a number in `+`, `-`, `*` or `/` is the
		// number it spells, a Double: `1 - "1E3"` is -999, and
		// `&H7FFFFFFF * "2"` is 4294967294 (issue #556). Beside a Currency it
		// is a Currency: `922337203685477.5807@ + "2"` overflows (issue #494,
		// measured in Excel 16.0). Two strings under `+` join instead.
		const beside = ARITHMETIC_BESIDE.has(before) ? this.toks[this.index - 2] : ARITHMETIC_BESIDE.has(after) ? this.toks[this.index + 2] : undefined;
		const besideType = beside && beside.kind !== 'stringLiteral' ? literalValue(beside)?.type : undefined;
		if (tok.kind === 'stringLiteral' && this.divisionByZero && besideType) {
			const read = numberInString(stringLiteralValue(tok.rawText));
			if (read === 'overflow' || read === undefined) {
				return read === undefined ? undefined : { overflow: true, span: this.span(this.index, this.index), detail: `${tok.rawText} spells a number past the Double range` };
			}
			this.index++;
			const type: NumericType = besideType === 'currency' ? 'currency' : 'double';
			if (!inRange(read.value, type)) {
				return { overflow: true, span: this.span(this.index - 1, this.index - 1), detail: `${tok.rawText} is outside the ${RANGES[type].label} range` };
			}
			const scaled = type === 'currency' ? currencyScaled(String(Math.abs(read.value))) : undefined;
			return { value: read.value, type, constant: true, ...(scaled !== undefined ? { scaled: read.value < 0 ? -scaled : scaled } : {}) };
		}
		// A With member at an operand's start: `.Rows.Count` (issue #411).
		const leading = this.index === 0 || ['(', ',', '='].includes(this.toks[this.index - 1].rawText) || this.toks[this.index - 1].kind === 'operator' || this.toks[this.index - 1].kind === 'keyword';
		if (tok.rawText === '.' && leading && this.names(WITH_SHEET) !== undefined) {
			return this.sheetSize(true);
		}
		const name = tokenName(tok);
		if (!name) {
			return undefined;
		}
		const size = this.sheetSize();
		if (size) {
			return size;
		}
		// A String local known to hold a number, beside -, *, /, \, ^ or Mod,
		// is that number as a Double: `s ^ 32767` with s = "12" overflows
		// (issue #331, measured in Excel 16.0). Beside + two Strings join.
		const numericBeside = (word: string): boolean => ['-', '*', '/', '\\', '^', 'mod'].includes(word);
		const lone = this.toks[this.index + 1]?.rawText !== '(' && this.toks[this.index + 1]?.rawText !== '.' && this.toks[this.index - 1]?.rawText !== '.';
		// A `-` at the start is a sign, not an operator.
		if (lone && (numericBeside(after) || (numericBeside(before) && (before !== '-' || this.index >= 2)))) {
			const spelled = this.names(`"${name.toLowerCase()}`);
			if (spelled) {
				this.index++;
				return { value: spelled.value, type: 'double' };
			}
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
			// `CInt(s)` with s a String local known to hold "40000" (issue
			// #407): the lookup answers a `"` key with the number it spells.
			const held = argument.length === 1 && tokenName(argument[0]) ? this.names(`"${tokenName(argument[0])!.toLowerCase()}`) : undefined;
			if (held && callee !== 'val') {
				return this.convert(callee, held, this.span(start, close), argument[0].rawText);
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
		// `Round(c)` of a Currency comes back a Currency, rounded half to even:
		// Round(922337203685477.5807@) is past the Currency range and raises 6
		// (issue #332, measured in Excel 16.0).
		if (callee === 'round' && this.toks[calleeIndex + 1]?.rawText === '(') {
			const close = matchParenFrom(this.toks, calleeIndex + 1);
			const args = close < 0 ? [] : splitTopLevelTokenGroups(this.toks, calleeIndex + 2, ',', close);
			const inner = args.length === 1 ? new TypedFolder(args[0], this.base, this.names).fold() : undefined;
			if (inner && !isOverflow(inner) && inner.type === 'currency') {
				const rounded = bankersRound(inner.value);
				if (rounded < RANGES.currency.min || rounded > RANGES.currency.max) {
					return { overflow: true, span: this.span(calleeIndex, close), detail: `Round of ${showNumber(inner.value)} gives ${showNumber(rounded)}, past the Currency range` };
				}
			}
		}
		if (this.toks[calleeIndex + 1]?.rawText === '(' && RESULT_FUNCTIONS.has(callee)) {
			const close = matchParenFrom(this.toks, calleeIndex + 1);
			const result = close < 0 ? undefined : this.functionResult(callee, splitTopLevelTokenGroups(this.toks, calleeIndex + 2, ',', close));
			if (result !== undefined) {
				this.index = close + 1;
				return result;
			}
			return undefined;
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
			// `F()`, a Function of the module whose result is known (issue #448).
			const result = this.toks[this.index + 2]?.rawText === ')' && this.toks[this.index - 1]?.rawText !== '.' ? this.names(`${name.toLowerCase()}()`) : undefined;
			if (result) {
				this.index += 3;
			}
			return result; // undefined for a call the folder does not know
		}
		const known = this.names(name.toLowerCase());
		if (!known) {
			return undefined;
		}
		this.index++;
		return known;
	}

	/**
	 * A size Excel fixes, read from a member chain at the current token (issue
	 * #411, measured in Excel 16.0): a worksheet's `Rows.Count` (1048576) or
	 * `Columns.Count` (16384) through ActiveSheet, Application, Worksheets(n)
	 * or a Worksheet local; `Cells(r, c).Row` and `.Column`; and the Row,
	 * Column, Count, Rows.Count and Columns.Count of `Range("A1:B2")`, a
	 * literal address. A Long each. Undefined, and nothing consumed, for any
	 * other chain.
	 */
	private sheetSize(withSubject = false): Folded {
		// `.Rows.Count` inside `With ActiveSheet` reads the With's sheet (issue #411).
		const segments: Array<{ name: string; args?: VbaToken[][] }> = withSubject ? [{ name: WITH_SHEET }] : [];
		let i = withSubject ? this.index + 1 : this.index;
		for (;;) {
			const name = tokenName(this.toks[i]);
			if (!name) {
				return undefined;
			}
			let end = i;
			let args: VbaToken[][] | undefined;
			if (this.toks[i + 1]?.rawText === '(') {
				const close = matchParenFrom(this.toks, i + 1);
				if (close < 0) {
					return undefined;
				}
				args = splitTopLevelTokenGroups(this.toks, i + 2, ',', close);
				end = close;
			}
			segments.push({ name: name.toLowerCase(), ...(args ? { args } : {}) });
			if (this.toks[end + 1]?.rawText !== '.') {
				i = end + 1;
				break;
			}
			i = end + 2;
		}
		if (segments.length < 2 || this.toks[i]?.rawText === '(') {
			return undefined;
		}
		const value = sheetSizeOf(segments, (expr) => {
			const folded = new TypedFolder(expr, this.base, this.names, this.divisionByZero).fold();
			return folded && !isOverflow(folded) ? folded.value : undefined;
		}, this.names);
		if (value === undefined) {
			return undefined;
		}
		const start = this.index;
		this.index = i;
		// Range.Count is a Long, and a sheet has 17,179,869,184 cells:
		// `Cells.Count` raises 6 (issue #278, measured in Excel 16.0).
		if (value > 2147483647) {
			return { overflow: true, span: this.span(start, i - 1), detail: `${this.toks.slice(start, i).map((tok) => tok.rawText).join('')} counts ${value} cells, past the Long range Range.Count returns; CountLarge counts them` };
		}
		return { value, type: 'long' };
	}

	/**
	 * What a VBA function returns for arguments the folder can read (issue
	 * #407, measured in Excel 16.0): Sgn is an Integer, -1, 0 or 1; Choose
	 * and IIf give the argument they pick; Len of `String(n, c)` or of a
	 * literal is a Long; Asc or AscW of a literal, and AscW of `ChrW(n)`, an
	 * Integer. Undefined for anything else.
	 */
	private functionResult(callee: string, args: VbaToken[][]): Folded {
		// Every argument is evaluated, the ones Choose and IIf do not pick
		// too: an overflow in any of them raises (issue #258).
		for (const arg of args) {
			const folded = arg.length > 0 ? new TypedFolder(arg, this.base, this.names, this.divisionByZero).fold() : undefined;
			if (isOverflow(folded)) {
				return folded;
			}
		}
		const fold = (toks: VbaToken[] | undefined): Typed | undefined => {
			const folded = toks && toks.length > 0 ? new TypedFolder(toks, this.base, this.names, this.divisionByZero).fold() : undefined;
			return folded && !isOverflow(folded) ? folded : undefined;
		};
		const call = (toks: VbaToken[] | undefined, name: string): VbaToken[][] | undefined => {
			const head = toks && toks.length >= 3 ? tokenText(toks[0]) : '';
			const open = toks?.[1]?.rawText === '$' ? 2 : 1;
			return toks && head === name && toks[open]?.rawText === '(' && matchParenFrom(toks, open) === toks.length - 1
				? splitTopLevelTokenGroups(toks, open + 1, ',', toks.length - 1) : undefined;
		};
		switch (callee) {
			case 'sgn': {
				const value = args.length === 1 ? fold(args[0]) : undefined;
				return value ? { value: Math.sign(value.value), type: 'integer' } : undefined;
			}
			case 'choose': {
				const index = fold(args[0]);
				const k = index ? Math.trunc(index.value) : undefined;
				return k !== undefined && k >= 1 && k < args.length ? fold(args[k]) : undefined;
			}
			case 'iif': {
				const word = args.length === 3 && args[0].length === 1 ? tokenText(args[0][0]) : '';
				return word === 'true' ? fold(args[1]) : word === 'false' ? fold(args[2]) : undefined;
			}
			case 'len': {
				const only = args.length === 1 ? args[0] : undefined;
				if (only?.length === 1 && only[0].kind === 'stringLiteral') {
					return { value: stringLiteralValue(only[0].rawText).length, type: 'long' };
				}
				const made = call(only, 'string');
				const count = made?.length === 2 ? fold(made[0]) : undefined;
				return count && Number.isInteger(count.value) && count.value >= 0 ? { value: count.value, type: 'long' } : undefined;
			}
			case 'asc':
			case 'ascw': {
				// `Asc("a")` is 97, an Integer (issue #351, measured in Excel 16.0).
				const only = args.length === 1 ? args[0] : undefined;
				if (only?.length === 1 && only[0].kind === 'stringLiteral') {
					const text = stringLiteralValue(only[0].rawText);
					return text.length > 0 && text.charCodeAt(0) < 128 ? { value: text.charCodeAt(0), type: 'integer' } : undefined;
				}
				if (callee === 'asc') {
					return undefined;
				}
				const made = call(only, 'chrw');
				const code = made?.length === 1 ? fold(made[0]) : undefined;
				if (!code || !Number.isInteger(code.value) || code.value < -32768 || code.value > 65535) {
					return undefined;
				}
				return { value: code.value > 32767 ? code.value - 65536 : code.value, type: 'integer' };
			}
		}
		return undefined;
	}

	/**
	 * `Date` and `Now`, a Date whose serial is today's: past an Integer's
	 * 32767 on any clock set after May 1989, so `i = Now` with i As Integer
	 * overflows (issue #327, measured in Excel 16.0). Not when a name of the
	 * procedure's or a call hides them.
	 */
	private currentDate(): Typed | undefined {
		const word = tokenText(this.toks[this.index]);
		const next = this.toks[this.index + 1]?.rawText;
		if ((word !== 'date' && word !== 'now') || next === '(' || next === '.' || next === '$' || this.toks[this.index - 1]?.rawText === '.' || this.names(word) !== undefined) {
			return undefined;
		}
		this.index++;
		const today = Math.floor((Date.now() - DATE_EPOCH_MS) / DAY_MS);
		return { value: word === 'now' ? today + 0.5 : today, type: 'date' };
	}

	/**
	 * `DateSerial(2020, 1, 1)` with whole-number literal arguments: the Date
	 * it names, with month and day rolling over as VBA rolls them (issue
	 * #327, measured in Excel 16.0).
	 */
	private dateSerial(): Typed | undefined {
		if (tokenText(this.toks[this.index]) !== 'dateserial' || this.toks[this.index + 1]?.rawText !== '(' || this.toks[this.index - 1]?.rawText === '.') {
			return undefined;
		}
		const close = matchParenFrom(this.toks, this.index + 1);
		const args = close < 0 ? [] : splitTopLevelTokenGroups(this.toks, this.index + 2, ',', close);
		const parts = args.map((arg) => (arg.length === 1 && arg[0].kind === 'integerLiteral' ? parseVbaIntegerLiteral(arg[0].rawText) : undefined));
		if (parts.length !== 3 || parts.some((part) => part === undefined) || parts[0]! < 100 || parts[0]! > 9999) {
			return undefined;
		}
		const ms = Date.UTC(parts[0]!, parts[1]! - 1, parts[2]!);
		this.index = close + 1;
		return { value: Math.round((ms - DATE_EPOCH_MS) / DAY_MS), type: 'date' };
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
		// Int and Fix of that Date are a Date still past the range, and raise 6;
		// CDate hands it back unchanged (issue #405, measured in Excel 16.0).
		if (inner.pastDate !== undefined && (target === 'int' || target === 'fix')) {
			return { overflow: true, span, detail: `${target === 'int' ? 'Int' : 'Fix'} of ${inner.pastDate} is a Date outside the Date range` };
		}
		if (inner.pastDate !== undefined && target === 'date') {
			return inner;
		}
		if (target === 'int' || target === 'fix') {
			const value = target === 'int' ? Math.floor(inner.value) : Math.trunc(inner.value);
			// A LongLong stays one: `Fix(q) Mod 7` works in LongLong (issue #480).
			return { value, type: inner.type === 'byte' || inner.type === 'integer' || inner.type === 'long' || inner.type === 'longlong' ? inner.type : 'double' };
		}
		if (target === 'exp') {
			const value = Math.exp(inner.value);
			return inRange(value, 'double')
				? { value, type: 'double' }
				: { overflow: true, span, detail: `Exp(${showNumber(inner.value)}) exceeds the Double range` };
		}
		if (target === 'decimal') {
			// A Decimal holds up to 2^96 - 1: CDec("1E28") runs, CDec("1E30") and
			// CDec(1E+30) overflow (issue #218). A whole one is followed exactly,
			// as a Double too wide to tell its neighbours apart (issue #502).
			if (!decimalFits(inner.value, shown)) {
				return { overflow: true, span, detail: `CDec(${shown}) does not fit Decimal` };
			}
			const decimal = spelledWhole(shown) ?? (Number.isSafeInteger(inner.value) ? BigInt(inner.value) : undefined);
			return decimal === undefined ? undefined : { value: Number(decimal), type: 'double', decimal };
		}
		if (target === 'hex' || target === 'oct') {
			// Hex and Oct take a value that fits a Long, or a LongLong on 64-bit
			// once rounded: Hex(3000000000.5) runs there (issue #332, measured in
			// Excel 16.0); 1E+20 fits neither.
			return inRange(inner.value, 'long') || Math.abs(bankersRound(inner.value)) < 9.2e18
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
		const value = type === 'single' || type === 'double' || type === 'currency' || type === 'date' ? inner.value : storedValue(inner, type).value;
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
		// `\` and Mod convert both operands to a Long before they divide, so a
		// Decimal past the Long range overflows first: `m Mod 0` raises 6, not
		// 11 (issue #502, measured in Excel 16.0).
		if ((op === '\\' || op === 'mod') && left.type !== 'longlong' && !inRange(bankersRound(left.value), 'long')) {
			return { overflow: true, span, detail: `${describe(left)} is outside the Long range that ${op === 'mod' ? 'Mod' : op} converts its operands to` };
		}
		// `1 / 0`, `1 \ 0.4` and `1 Mod False` divide by zero; outside a Const
		// that is division-by-zero's to report.
		const divisor = op === '/' ? right.value : op === '\\' || op === 'mod' ? bankersRound(right.value) : undefined;
		if (divisor === 0) {
			this.divisionByZero?.(span);
			return undefined;
		}
		// A whole Decimal with a whole number: `+`, `-` and `*` stay exact and
		// overflow past 2^96 - 1 (issue #502, measured in Excel 16.0).
		if ((left.decimal !== undefined || right.decimal !== undefined) && (op === '+' || op === '-' || op === '*')) {
			const a = left.decimal ?? exactOf(left);
			const b = right.decimal ?? exactOf(right);
			if (a === undefined || b === undefined) {
				return undefined;
			}
			const result = op === '+' ? a + b : op === '-' ? a - b : a * b;
			return result < DECIMAL_LIMIT && result > -DECIMAL_LIMIT
				? { value: Number(result), type: 'double', decimal: result }
				: { overflow: true, span, detail: `${describe(left)} ${op} ${describe(right)} is past the Decimal range` };
		}
		if ((left.type === 'longlong' || right.type === 'longlong') && op !== '/' && op !== '^') {
			return combineLongLong(left, right, op, span);
		}
		// A Currency sum in ten-thousandths, exactly: a double cannot tell
		// 922337203685477.5807 from one past it (issue #494, measured in Excel 16.0).
		if ((op === '+' || op === '-') && arithmeticResultType(left.type, right.type, op) === 'currency') {
			const a = scaledOf(left);
			const b = scaledOf(right);
			if (a !== undefined && b !== undefined) {
				const sum = op === '+' ? a + b : a - b;
				return sum >= -LONGLONG_LIMIT && sum < LONGLONG_LIMIT
					? { value: Number(sum) / 10000, type: 'currency', scaled: sum }
					: { overflow: true, span, detail: `${describe(left)} ${op} ${describe(right)} is past the Currency range` };
			}
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
				// Both operands are converted to a Long first: `1E10 \ 2` and
				// `7 Mod 922337203685477@` overflow (issue #458, measured in
				// Excel 16.0).
				const outside = [left, right].find((_, k) => !inRange([a, b][k], 'long'));
				if (outside) {
					return { overflow: true, span, detail: `${describe(outside)} is outside the Long range that ${op === 'mod' ? 'Mod' : op} converts its operands to` };
				}
				value = op === 'mod' ? a % b : Math.trunc(a / b);
				break;
			}
			default:
				type = arithmeticResultType(left.type, right.type, op);
				value = op === '+' ? left.value + right.value : op === '-' ? left.value - right.value : left.value * right.value;
				break;
		}
		// `number + Date` and `number - Date` hold a serial past the Date range
		// without raising; only a later use fails (issue #330, measured in
		// Excel 16.0). `Date + number` raises 6.
		if (!inRange(value, type) && type === 'date' && left.type !== 'date') {
			return { value, type, pastDate: `${describe(left)} ${op} ${describe(right)}` };
		}
		if ((left.variant || right.variant) && (op === '+' || op === '-' || op === '*')) {
			let widened: NumericType | undefined = type;
			while (widened && !inRange(value, widened)) {
				widened = VARIANT_WIDENING.get(widened);
			}
			return widened ? { value, type: widened, variant: true } : undefined;
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

/** A Currency or whole-number operand in ten-thousandths, exactly, when it has that. */
function scaledOf(typed: Typed): bigint | undefined {
	if (typed.scaled !== undefined) {
		return typed.scaled;
	}
	return WHOLE_TYPES.has(typed.type) && Number.isSafeInteger(typed.value) ? BigInt(typed.value) * 10000n : undefined;
}

/** A Currency literal's digits in ten-thousandths: `922337203685477.5807@`. */
function currencyScaled(text: string): bigint | undefined {
	const match = /^(\d*)(?:\.(\d{0,4}))?$/.exec(text);
	if (!match || (match[1] === '' && !match[2])) {
		return undefined;
	}
	return BigInt(match[1] || '0') * 10000n + BigInt((match[2] ?? '').padEnd(4, '0'));
}

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
	// Not of a Boolean is a Boolean: `b = Not False` stores 255 in a Byte.
	const constant = { ...(operand.constant ? { constant: true } : {}), ...(operand.boolean ? { boolean: true } : {}) };
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
	if (typed.scaled !== undefined) {
		const sign = typed.scaled < 0n ? '-' : '';
		const digits = (typed.scaled < 0n ? -typed.scaled : typed.scaled).toString().padStart(5, '0');
		const fraction = digits.slice(-4).replace(/0+$/, '');
		return `${sign}${digits.slice(0, -4)}${fraction ? `.${fraction}` : ''} (Currency)`;
	}
	if (typed.decimal !== undefined) {
		return `${typed.decimal} (Decimal)`;
	}
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
	if (folded.boolean && target === 'byte') {
		return { value: folded.value === 0 ? 0 : 255 };
	}
	const value = bankersRound(folded.value);
	// A Date goes into a Byte through an Integer, keeping the low byte: -1
	// stores 255 and 1000 stores 232; past the Integer range it overflows.
	// CByte of one does not wrap (issue #624, measured in Excel 16.0).
	if (folded.type === 'date' && target === 'byte' && value >= -32768 && value <= 32767) {
		return { value: ((value % 256) + 256) % 256 };
	}
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

/** The VBA functions whose result the folder works out (issue #407). */
const RESULT_FUNCTIONS: ReadonlySet<string> = new Set(['sgn', 'choose', 'iif', 'len', 'asc', 'ascw']);

/** The type a Variant's arithmetic widens to when a result does not fit (issue #480). */
const VARIANT_WIDENING: ReadonlyMap<NumericType, NumericType> = new Map([
	['byte', 'integer'], ['integer', 'long'], ['long', 'double'], ['single', 'double'],
]);

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
	const enums: VbaSymbol[] = [];
	for (const symbol of candidates) {
		if (symbol.kind === 'constant' && symbol.defaultRaw !== undefined) {
			pending.set(symbol.name.toLowerCase(), symbol);
		} else if (symbol.kind === 'enum') {
			enums.push(symbol);
		}
	}
	if (pending.size === 0 && enums.length === 0) {
		return base;
	}
	const folded = new Map<string, Typed>();
	const resolving = new Set<string>();
	const enumValues = new Map<string, Typed>();
	const resolve = (lower: string): Typed | undefined => {
		const symbol = pending.get(lower);
		if (!symbol) {
			return enumValues.get(lower) ?? base.get(lower);
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
		// `Const T As Boolean = 1` is True.
		if (normalizeType(symbol.asType) === 'boolean') {
			const truth: Typed = { value: value.value === 0 ? 0 : -1, type: 'integer', constant: true, boolean: true };
			folded.set(lower, truth);
			return truth;
		}
		const typed: Typed = declared && kept
			? { value: kept.value, type: declared, ...(kept.exact !== undefined ? { exact: kept.exact } : {}), constant: true }
			: value;
		if (declared && kept && !inRange(kept.value, declared, kept.exact)) {
			return undefined;
		}
		folded.set(lower, typed);
		return typed;
	};
	// An Enum member is a Long: its own value, or one more than the member
	// before it, the first 0 (issue #255). `E.eBig` names it too.
	for (const symbol of enums) {
		let next: number | undefined = 0;
		for (const member of symbol.children ?? []) {
			if (member.kind !== 'enumMember') {
				continue;
			}
			const value: Folded = member.defaultRaw === undefined
				? (next === undefined ? undefined : { value: next, type: 'long' })
				: new TypedFolder(rawExpressionTokens(member.defaultRaw).filter((tok) => tok.kind !== 'comment'), 0, resolve).fold();
			if (value === undefined || isOverflow(value) || !Number.isInteger(value.value) || !inRange(value.value, 'long')) {
				next = undefined;
				continue;
			}
			const typed: Typed = { value: value.value, type: 'long', constant: true };
			enumValues.set(member.name.toLowerCase(), typed);
			enumValues.set(`${symbol.name.toLowerCase()}.${member.name.toLowerCase()}`, typed);
			next = value.value + 1;
		}
	}
	const out = new Map<string, Typed>([...base, ...enumValues]);
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

const SHEET_ROWS = 1048576;
const SHEET_COLUMNS = 16384;

/** The lookup's name for the subject of the With a statement sits in, when that names a sheet. */
const WITH_SHEET = '#with';

/** The segments of a member chain written whole: `ActiveSheet`, `Worksheets(1)`, `ThisWorkbook.Worksheets(2)`. */
function chainSegments(toks: readonly VbaToken[]): ChainSegment[] | undefined {
	const segments: ChainSegment[] = [];
	let i = 0;
	while (i < toks.length) {
		const name = tokenName(toks[i]);
		if (!name) {
			return undefined;
		}
		let end = i;
		let args: VbaToken[][] | undefined;
		if (toks[i + 1]?.rawText === '(') {
			const close = matchParenFrom(toks, i + 1);
			if (close < 0) {
				return undefined;
			}
			args = splitTopLevelTokenGroups(toks, i + 2, ',', close);
			end = close;
		}
		segments.push({ name: name.toLowerCase(), ...(args ? { args } : {}) });
		if (end + 1 === toks.length) {
			return segments;
		}
		if (toks[end + 1].rawText !== '.') {
			return undefined;
		}
		i = end + 2;
	}
	return undefined;
}

interface ChainSegment { name: string; args?: VbaToken[][] }

/**
 * Whether a member chain names a whole worksheet (issue #411): nothing (the
 * active sheet), ActiveSheet, Application, Worksheets(n) from Application,
 * ThisWorkbook, ActiveWorkbook or Workbooks(n), or a local As Worksheet.
 */
function namesSheet(receiver: readonly ChainSegment[], names: NameLookup): boolean {
	const [first, second] = receiver;
	if (receiver.length === 0) {
		return true;
	}
	if (receiver.length === 1) {
		return (!first.args && (first.name === 'activesheet' || first.name === 'application' || (first.name === WITH_SHEET && names(WITH_SHEET) !== undefined)))
			|| (first.name === 'worksheets' && first.args?.length === 1)
			|| (!first.args && names(`${first.name}.rows.count`) !== undefined);
	}
	if (receiver.length === 2 && second.name === 'worksheets' && second.args?.length === 1) {
		return (!first.args && ['thisworkbook', 'activeworkbook', 'application'].includes(first.name)) || (first.name === 'workbooks' && first.args?.length === 1);
	}
	return receiver.length === 2 && !first.args && first.name === 'application' && !second.args && second.name === 'activesheet';
}

/**
 * The procedure's Range, Object or Variant locals that one Set gives a
 * whole sheet's Cells, `Set r = Cells` or `Set r = ActiveSheet.Cells`, and
 * that nothing else assigns or passes whole (issue #278, measured in Excel
 * 16.0: `r.Count` raises 6).
 */
function wholeSheetCellLocals(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	names: NameLookup,
	activity: ConditionalActivityTracker | undefined,
): ReadonlySet<string> {
	const candidates = new Set((procedureSymbolFor(symbols, member)?.children ?? [])
		.filter((child) => child.kind === 'localVariable' && !child.isArray && child.visibility !== 'Static' && ['range', 'object', 'variant'].includes(normalizeType(child.asType) ?? 'variant'))
		.map((child) => child.name.toLowerCase()));
	if (candidates.size === 0) {
		return candidates;
	}
	const sets = new Map<string, VbaToken[][]>();
	const ruled = new Set<string>();
	forEachStatement(member.body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			const toks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
			const target = setAssignmentTarget(source, span)?.name.toLowerCase();
			if (target && candidates.has(target)) {
				sets.set(target, [...(sets.get(target) ?? []), toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1)]);
				continue;
			}
			const bare = bareAssignmentTarget(source, span)?.name.toLowerCase();
			if (bare) {
				ruled.add(bare);
			}
			for (const lower of trackedLocalsNamedWhole(toks, 0, (name) => candidates.has(name), new Set()).keys()) {
				ruled.add(lower);
			}
		}
	}, activity);
	const out = new Set<string>();
	for (const [lower, values] of sets) {
		const chain = values.length === 1 && !ruled.has(lower) ? chainOf(values[0]) : undefined;
		const last = chain?.[chain.length - 1];
		if (chain && last?.name === 'cells' && !last.args && namesSheet(chain.slice(0, -1), names)) {
			out.add(lower);
		}
	}
	return out;
}

/** `a.b(1).c` as member segments, or undefined for anything else. */
function chainOf(toks: readonly VbaToken[]): ChainSegment[] | undefined {
	const out: ChainSegment[] = [];
	for (let i = 0; i < toks.length;) {
		const name = tokenName(toks[i])?.toLowerCase();
		if (!name) {
			return undefined;
		}
		let end = i;
		let args: VbaToken[][] | undefined;
		if (toks[i + 1]?.rawText === '(') {
			const close = matchParenFrom([...toks], i + 1);
			if (close < 0) {
				return undefined;
			}
			args = splitTopLevelTokenGroups([...toks], i + 2, ',', close);
			end = close;
		}
		out.push({ name, ...(args ? { args } : {}) });
		if (end + 1 === toks.length) {
			return out;
		}
		if (toks[end + 1].rawText !== '.') {
			return undefined;
		}
		i = end + 2;
	}
	return undefined;
}

/** The cells `Range("A1:B2")` names, from a literal A1 address: a cell, a block, whole columns or whole rows. */
function literalRange(segment: ChainSegment | undefined): { row: number; column: number; rows: number; columns: number } | undefined {
	const arg = segment?.name === 'range' && segment.args?.length === 1 ? segment.args[0].filter((tok) => tok.kind !== 'comment') : undefined;
	if (!arg || arg.length !== 1 || arg[0].kind !== 'stringLiteral') {
		return undefined;
	}
	const text = stringLiteralValue(arg[0].rawText).replace(/\$/g, '').toUpperCase();
	const column = (letters: string): number => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
	const cells = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(text);
	if (cells) {
		const [r1, c1] = [Number(cells[2]), column(cells[1])];
		const [r2, c2] = cells[3] ? [Number(cells[4]), column(cells[3])] : [r1, c1];
		const valid = [r1, r2].every((r) => r >= 1 && r <= SHEET_ROWS) && [c1, c2].every((c) => c >= 1 && c <= SHEET_COLUMNS);
		return valid ? { row: Math.min(r1, r2), column: Math.min(c1, c2), rows: Math.abs(r2 - r1) + 1, columns: Math.abs(c2 - c1) + 1 } : undefined;
	}
	const wholeColumns = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(text);
	if (wholeColumns) {
		const [c1, c2] = [column(wholeColumns[1]), column(wholeColumns[2])];
		return c1 <= SHEET_COLUMNS && c2 <= SHEET_COLUMNS ? { row: 1, column: Math.min(c1, c2), rows: SHEET_ROWS, columns: Math.abs(c2 - c1) + 1 } : undefined;
	}
	const wholeRows = /^(\d+):(\d+)$/.exec(text);
	if (wholeRows) {
		const [r1, r2] = [Number(wholeRows[1]), Number(wholeRows[2])];
		return [r1, r2].every((r) => r >= 1 && r <= SHEET_ROWS) ? { row: Math.min(r1, r2), column: 1, rows: Math.abs(r2 - r1) + 1, columns: SHEET_COLUMNS } : undefined;
	}
	return undefined;
}

/**
 * The size a member chain reads, where Excel fixes it (issue #411):
 * `ws.Rows.Count`, `Cells(Rows.Count, 1).Row`, `Range("A1:A40000").Rows.Count`.
 */
function sheetSizeOf(chain: readonly ChainSegment[], fold: (expr: VbaToken[]) => number | undefined, names: NameLookup): number | undefined {
	if (names('rows.count') === undefined) {
		return undefined; // not Excel, or a name of the procedure's hides it
	}
	// `cells.Count` with a variable of the code's own named cells.
	if (['cells', 'range', 'rows', 'columns'].includes(chain[0].name) && names.declares?.(chain[0].name)) {
		return undefined;
	}
	// `Set r = Cells`, then `r.Count` reads the sheet's Cells (issue #278).
	const segments = !chain[0].args && names.wholeSheetCells?.(chain[0].name) ? [{ name: 'cells' }, ...chain.slice(1)] : chain;
	const n = segments.length;
	const last = segments[n - 1];
	const before = segments[n - 2];
	if (last.args) {
		return undefined;
	}
	if (last.name === 'count' && (before.name === 'rows' || before.name === 'columns') && !before.args) {
		const receiver = segments.slice(0, n - 2);
		const block = receiver.length >= 1 ? literalRange(receiver[receiver.length - 1]) : undefined;
		if (block && namesSheet(receiver.slice(0, -1), names)) {
			return before.name === 'rows' ? block.rows : block.columns;
		}
		return namesSheet(receiver, names) ? (before.name === 'rows' ? SHEET_ROWS : SHEET_COLUMNS) : undefined;
	}
	const receiver = segments.slice(0, n - 2);
	// `Cells.Count` on the whole sheet, `Cells.Cells.Count` too (issue #278).
	if (last.name === 'count' && before.name === 'cells' && !before.args) {
		const sheet = receiver[receiver.length - 1]?.name === 'cells' && !receiver[receiver.length - 1].args ? receiver.slice(0, -1) : receiver;
		return namesSheet(sheet, names) ? SHEET_ROWS * SHEET_COLUMNS : undefined;
	}
	if (!namesSheet(receiver, names)) {
		return undefined;
	}
	const block = literalRange(before);
	if (block) {
		return last.name === 'count' ? block.rows * block.columns : last.name === 'row' ? block.row : last.name === 'column' ? block.column : undefined;
	}
	if (before.name === 'cells' && before.args?.length === 2 && (last.name === 'row' || last.name === 'column')) {
		return fold(before.args[last.name === 'row' ? 0 : 1]);
	}
	return undefined;
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
	const types = moduleTypes(source, mod, activity);
	const results = knownFunctionResults(source, mod, activity);
	const deftypes = /^[ \t]*Def(Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)[ \t]+[A-Za-z]/im.test(source);
	// Names the project declares, which hide a host global of that spelling.
	const moduleNames = new Set([...(projectVisibleSymbols ?? []), ...(symbols.root.children ?? [])].map((symbol) => symbol.name.toLowerCase()));
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const children = procedureSymbolFor(symbols, member)?.children ?? [];
		// A local or parameter hides a module Const or Enum member of its name.
		const constants = new Map(constantLookup(moduleConstants, children));
		for (const child of children) {
			if (child.kind === 'localVariable' || child.kind === 'parameter') {
				constants.delete(child.name.toLowerCase());
			}
		}
		// A local declared with no type is a Variant, which holds a Double as
		// a Double: `Dim v: v = 3E9` then `v Mod 2` raises 6 (issue #323,
		// measured in Excel 16.0). A DefType statement types it otherwise.
		const untyped = deftypes ? [] : children.filter((child) => child.kind === 'localVariable' && !child.asType && !child.isArray && /\w$/.test(child.name));
		const env: ReadonlyMap<string, string> = untyped.length === 0
			? typeEnvironmentFor(symbols, member)
			: new Map([...typeEnvironmentFor(symbols, member), ...untyped.map((child): [string, string] => [child.name.toLowerCase(), 'Variant'])]);
		// What each local holds as the statement being checked is reached: a
		// value stored only in a branch that does not run is not there
		// (issue #565).
		const knownAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		let at: BodyNode | undefined;
		// Values a straight run of top-level statements has just stored:
		// `i = 32767` followed by `i = i + 1`.
		const justAssigned = new Map<string, Typed>();
		const names: NameLookup = (lower) => {
			const known = knownAt(at);
			// `"s`: the number a String local known to hold one spells, for a
			// conversion to read (issue #407).
			if (lower.startsWith('"')) {
				const held = known.get(lower.slice(1));
				const read = held?.kind === 'string' && !held.contentMutated && normalizeType(env.get(lower.slice(1))) === 'string' ? numberInString(held.value as string) : undefined;
				return read === 'overflow' ? undefined : read;
			}
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
			// A Boolean is an Integer in arithmetic, True -1: `n - b` with n the
			// largest Long and b True overflows (issue #331, measured in Excel 16.0).
			// Kept a Boolean, as True is, so a Byte takes it as 255 (issue #624).
			if (local?.kind === 'number' && normalizeType(env.get(lower)) === 'boolean') {
				return { value: local.value as number, type: 'integer', boolean: true };
			}
			if (!lower.includes('.')) {
				// `b = F()` with F a Function of the module returning 300 (issue #448).
				const result = functionResultNamed(lower.replace(/\(\)$/, ''), results, member, symbols);
				const resultType = result?.kind === 'number' ? numericTypeOf(result.type) : undefined;
				return result?.kind === 'number' && resultType ? { value: result.value, type: resultType } : undefined;
			}
			// `ws.Rows.Count` with ws As Worksheet is the sheet's (issue #411).
			const head = lower.slice(0, lower.indexOf('.'));
			if (env.has(head)) {
				return normalizeType(env.get(head)) === 'worksheet' ? hostValues.get(lower.slice(head.length + 1)) : undefined;
			}
			return hostValues.get(lower);
		};
		names.declares = (lower) => env.has(lower) || moduleNames.has(lower);
		const sheetCells = wholeSheetCellLocals(source, member, symbols, names, activity);
		names.wholeSheetCells = (lower) => sheetCells.has(lower);
		const groups: VariableGroupNode[] = [];
		forEachVariableGroup(member.body, (group) => { groups.push(group); }, activity);
		checkConstDeclarations(source, groups, constants, activity, push);
		// `t.i = t.i + 1`: a numeric member of a Type value as the target (issue #253).
		const memberTarget = types.size === 0 ? undefined : (span: Span): AssignmentTarget | undefined => memberAssignmentTarget(source, span, symbols, member, types);
		checkProcedureBody(source, member, env, names, justAssigned, activity, push, memberTarget, (node) => {
			at = node;
		});
		at = undefined;
		checkByValArguments(source, member, symbols, names, activity, push);
		checkAccumulatingLoops(source, member, env, names, knownAt, activity, push);
	}
}

/**
 * An argument the folder can read, passed to a ByVal number parameter of a
 * procedure of the module that cannot hold it: `TakeI(Rows.Count)` with
 * `ByVal i As Integer` converts 1048576 and raises 6 (issue #411, measured
 * in Excel 16.0). A lone literal is argument-type-mismatch's.
 */
function checkByValArguments(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	names: NameLookup,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const signatures = buildModuleTypeSignatures(symbols);
	if (signatures.size === 0) {
		return;
	}
	const own = new Set([proc.name, ...proc.params.map((param) => param.name), ...(procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => child.name)].map((name) => name.toLowerCase()));
	forEachStatement(proc.body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			const toks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
			for (let i = 0; i < toks.length; i++) {
				const lower = tokenName(toks[i])?.toLowerCase();
				const signature = lower && !own.has(lower) && toks[i - 1]?.rawText !== '.' ? signatures.get(lower) : undefined;
				if (!signature) {
					continue;
				}
				// `F(a, b)` in a value, `Call F(a, b)`, or the statement `F a, b`.
				const parenthesized = toks[i + 1]?.rawText === '(';
				const statementCall = i === 0 && !parenthesized && toks.length > 1 && toks[1].rawText !== '=';
				if (!parenthesized && !statementCall) {
					continue;
				}
				const close = parenthesized ? matchParenFrom(toks, i + 1) : toks.length;
				const args = close < 0 ? [] : splitTopLevelTokenGroups(toks, parenthesized ? i + 2 : i + 1, ',', close);
				args.forEach((arg, k) => {
					const param = signature.params[k];
					const type = param && param.byRef === false && !param.isArray && !param.paramArray ? numericTypeOf(param.type) : undefined;
					const value = arg.filter((tok) => tok.kind !== 'comment');
					if (!type || value.length === 0 || value.some((tok) => tok.rawText === ':=') || literalTyped(value[value.length - 1]) && value.length <= 2) {
						return;
					}
					const folded = new TypedFolder(value, span.start, names).fold();
					if (!folded || isOverflow(folded)) {
						return;
					}
					const kept = storedValue(folded, type);
					if (!inRange(kept.value, type, kept.exact)) {
						push('arithmeticOverflow', `Argument '${param.name}' of '${signature.name}' is ByVal ${RANGES[type].label}, and ${value.map((tok) => tok.rawText).join('')} is ${showNumber(kept.value)}, outside its range ${rangeText(type)}. This will raise Run-time error '6': Overflow.`, {
							start: span.start + value[0].start,
							end: span.start + value[value.length - 1].end,
						});
					}
				});
			}
		}
	}, activity);
}

/** Statement heads after which a loop's pass may not run on. */
const LOOP_LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'resume', 'return', 'on', 'stop']);

/** The most passes a loop is run for to find its overflow. */
const MAX_ACCUMULATED_PASSES = 100000;

/**
 * A whole-number local a loop changes by the same statement every pass,
 * until it no longer fits its type (issue #263, measured in Excel 16.0):
 *
 *  - `For i = 1 To 300: t = t + i` on an Integer t, and `p = p * i` from
 *    p = 1 to 20 on a Long, are run pass by pass from the value the local
 *    holds as the loop starts.
 *  - `Do: i = i + 1: Loop Until i > 32767` on an Integer, and
 *    `While b <= 255: b = b + 1: Wend` on a Byte, cannot end any other way:
 *    no value of the type passes the exit test.
 *
 * The step is the loop's own top-level statement `x = x + k`, `x = x - k` or
 * `x = x * k`, k a whole number or the counter; nothing else in the body
 * names x, and nothing may leave the pass.
 */
function checkAccumulatingLoops(
	source: string,
	proc: ProcedureNode,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	startValues: (node: BodyNode) => ReadonlyMap<string, KnownLocalValue>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span) || !('body' in node) || !Array.isArray(node.body)) {
				continue;
			}
			if (node.kind === 'ForBlock') {
				accumulateFor(source, node, env, names, startValues(node), activity, push);
			} else if (node.kind === 'DoBlock' || node.kind === 'WhileBlock') {
				endlessStep(source, node, env, activity, push, startValues(node));
			}
			visit(node.body as BodyNode[]);
		}
	};
	visit(proc.body);
}

interface LoopStep {
	name: string;
	type: NumericType;
	op: '+' | '-' | '*';
	/** The other operand: a whole number, or the counter. */
	by: number | 'counter';
	span: Span;
	text: string;
}

/** The body's one step of a whole-number local, when the body is plain statements that run every pass. */
function loopStepIn(
	source: string,
	body: readonly BodyNode[],
	env: ReadonlyMap<string, string>,
	counter: string | undefined,
	activity: ConditionalActivityTracker | undefined,
	only?: string,
): LoopStep | undefined {
	const statements: Array<{ toks: readonly VbaToken[]; span: Span }> = [];
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (!isLeafStatement(node) || (node.kind === 'Statement' && node.singleLineIfBranches)) {
			return undefined;
		}
		const toks = statementTokens(source, node.span).filter((tok) => tok.kind !== 'comment');
		if (LOOP_LEAVING_HEADS.has(tokenText(toks[0])) || (tokenText(toks[0]) === 'end' && toks.length === 1) || jumpTargetLabelDeclaration(source, node.span)) {
			return undefined;
		}
		statements.push({ toks, span: node.span });
	}
	let step: LoopStep | undefined;
	for (const { toks, span } of statements) {
		const target = tokenName(toks[0])?.toLowerCase();
		const type = target ? numericTypeOf(env.get(target)) : undefined;
		if (!target || toks[1]?.rawText !== '=' || !type || !WHOLE_TYPES.has(type) || type === 'longlong' || (only && target !== only)) {
			continue;
		}
		const value = toks.slice(2);
		const self = (tok: VbaToken | undefined): boolean => tokenName(tok)?.toLowerCase() === target;
		const operand = (tok: VbaToken | undefined): number | 'counter' | undefined => {
			if (tok?.kind === 'integerLiteral') {
				return parseVbaIntegerLiteral(tok.rawText);
			}
			return counter && tokenName(tok)?.toLowerCase() === counter ? 'counter' : undefined;
		};
		if (value.length !== 3 || !['+', '-', '*'].includes(value[1].rawText)) {
			continue;
		}
		const op = value[1].rawText as LoopStep['op'];
		const by = self(value[0]) ? operand(value[2]) : op !== '-' && self(value[2]) ? operand(value[0]) : undefined;
		if (by === undefined || step) {
			return undefined; // one step only
		}
		step = { name: target, type, op, by, span, text: source.slice(span.start, span.end).trim() };
	}
	if (!step) {
		return undefined;
	}
	// Nothing else names the local.
	const mentions = statements.filter(({ toks }) => toks.some((tok, k) => tokenName(tok)?.toLowerCase() === step!.name && toks[k - 1]?.rawText !== '.'));
	return mentions.length === 1 ? step : undefined;
}

function accumulateFor(
	source: string,
	node: ForBlockNode,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	start: ReadonlyMap<string, KnownLocalValue>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (node.each || !node.controlVariable) {
		return;
	}
	const counter = node.controlVariable.toLowerCase();
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const eq = header.findIndex((tok) => tok.rawText === '=');
	const to = header.findIndex((tok) => tokenText(tok) === 'to');
	const stepAt = header.findIndex((tok) => tokenText(tok) === 'step');
	const fold = (toks: readonly VbaToken[]): number | undefined => {
		const folded = toks.length === 0 ? undefined : new TypedFolder(toks, node.span.start, names).fold();
		return folded && !isOverflow(folded) && Number.isInteger(folded.value) ? folded.value : undefined;
	};
	const first = eq > 0 && to > eq ? fold(header.slice(eq + 1, to)) : undefined;
	const limit = to > 0 ? fold(header.slice(to + 1, stepAt > 0 ? stepAt : header.length)) : undefined;
	const increment = stepAt > 0 ? fold(header.slice(stepAt + 1)) : 1;
	if (first === undefined || limit === undefined || !increment) {
		return;
	}
	const step = loopStepIn(source, node.body as BodyNode[], env, counter, activity);
	const initial = step ? start.get(step.name) : undefined;
	if (!step || step.name === counter || initial?.kind !== 'number' || !Number.isInteger(initial.value)) {
		return;
	}
	let value = initial.value as number;
	let passes = 0;
	for (let c = first; increment > 0 ? c <= limit : c >= limit; c += increment) {
		if (++passes > MAX_ACCUMULATED_PASSES) {
			return;
		}
		const by = step.by === 'counter' ? c : step.by;
		value = step.op === '+' ? value + by : step.op === '-' ? value - by : value * by;
		if (!inRange(value, step.type)) {
			if (passes === 1) {
				return; // the walk into the loop reports its first pass
			}
			push(
				'arithmeticOverflow',
				`On the pass of the For loop where '${node.controlVariable}' is ${c}, '${step.text}' makes '${step.name}' ${value}, which does not fit ${article(RANGES[step.type].label)} ${RANGES[step.type].label}. This will raise Run-time error '6': Overflow.`,
				{ start: step.span.start, end: step.span.start + source.slice(step.span.start, step.span.end).trimEnd().length },
			);
			return;
		}
	}
}

/** `Do ... Loop Until i > 32767` stepping an Integer: no value ends the loop, so the step overflows. */
function endlessStep(
	source: string,
	node: BodyNode,
	env: ReadonlyMap<string, string>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	startValues: ReadonlyMap<string, KnownLocalValue>,
): void {
	// The test: `Do While x`, `Do Until x`, `Loop While x`, `Loop Until x`, `While x`.
	const header = statementTokensAfterLeadingLabel(source, blockHeaderLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const footer = statementTokensAfterLeadingLabel(source, blockFooterLineSpan(source, node.span)).filter((tok) => tok.kind !== 'comment');
	const tests: Array<{ keyword: string; condition: readonly VbaToken[] }> = [];
	for (const line of [header, footer]) {
		const head = tokenText(line[0]);
		const word = head === 'while' ? 'while' : (head === 'do' || head === 'loop') ? tokenText(line[1]) : '';
		if (word === 'while' || word === 'until') {
			tests.push({ keyword: word, condition: line.slice(head === 'while' ? 1 : 2) });
		}
	}
	if (tests.length !== 1) {
		return;
	}
	const { keyword, condition } = tests[0];
	if (condition.length !== 3 && condition.length !== 4) {
		return;
	}
	// `x op c`, the constant signed or not.
	const name = tokenName(condition[0])?.toLowerCase();
	const op = condition[1]?.rawText;
	const negative = condition.length === 4 && condition[2].rawText === '-';
	const literal = condition[negative ? 3 : 2];
	const raw = literal?.kind === 'integerLiteral' ? parseVbaIntegerLiteral(literal.rawText) : undefined;
	if (!name || raw === undefined || !['<', '<=', '>', '>=', '=', '<>'].includes(op)) {
		return;
	}
	const limit = negative ? -raw : raw;
	const start = startValues.get(name);
	// The body is plain statements with no Exit, GoTo or End: loopStepIn
	// finds the step only there.
	const step = loopStepIn(source, (node as { body: BodyNode[] }).body, env, undefined, activity, name);
	if (!step || step.op === '*' || step.by === 'counter' || step.by <= 0) {
		return;
	}
	const range = RANGES[step.type];
	const holds = (w: number): boolean => op === '<' ? w < limit : op === '<=' ? w <= limit : op === '>' ? w > limit : op === '>=' ? w >= limit : op === '=' ? w === limit : w !== limit;
	const ends = (w: number): boolean => keyword === 'while' ? !holds(w) : holds(w);
	// The test changes at the constant, so the range's ends and the values
	// around the constant show whether any value ends the loop.
	const probes = [range.min, range.max, limit - 1, limit, limit + 1].filter((w) => w >= range.min && w <= range.max);
	const condText = condition.map((tok) => tok.rawText).join(' ');
	let reason = keyword === 'while' ? `the loop runs while ${condText}, which ${article(range.label)} ${range.label} always is` : `the loop ends only when ${condText}, which ${article(range.label)} ${range.label} never is`;
	if (probes.some(ends)) {
		// From a known start the loop reaches only start, start + step, ...:
		// `i = 0: Do Until i < 0: i = i + 1000` never ends before 33000
		// overflows an Integer (issue #479, measured in Excel 16.0).
		const held = start?.kind === 'number' && Number.isInteger(start.value) ? start.value as number : start?.kind === 'empty' ? 0 : undefined;
		if (held === undefined || held < range.min || held > range.max) {
			return;
		}
		const delta = step.op === '+' ? step.by : -step.by;
		const last = Math.floor(((delta > 0 ? range.max : range.min) - held) / delta);
		const reached = (k: number): number | undefined => (k >= 0 && k <= last ? held + k * delta : undefined);
		const nearest = [limit - 1, limit, limit + 1].flatMap((w) => [Math.floor((w - held) / delta), Math.ceil((w - held) / delta)]);
		const candidates = [0, last, ...nearest].map(reached).filter((w): w is number => w !== undefined);
		// A test after the step reads the stepped value; one before reads the start too.
		const footTest = tokenText(footer[0]) === 'loop' && ['while', 'until'].includes(tokenText(footer[1]));
		if (candidates.some((w) => ends(w) && !(footTest && w === held))) {
			return;
		}
		reason = keyword === 'while' ? `from ${held}, every value it reaches keeps ${condText} true` : `from ${held}, no value it reaches makes ${condText} true`;
	}
	push(
		'arithmeticOverflow',
		`'${step.name}' is ${article(range.label)} ${range.label}, and ${reason}, so '${step.text}' runs until it does not fit. This will raise Run-time error '6': Overflow.`,
		{ start: step.span.start, end: step.span.start + source.slice(step.span.start, step.span.end).trimEnd().length },
	);
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
	// `"" + 1`, `"abc" < 1`, `1 / ""`: a string with no digit beside a number
	// or a Date in arithmetic or a comparison (issue #367, measured in Excel
	// 16.0). `"2" * 2`, `"a" & 1` and `"a" = "b"` compile.
	if (value.length === 3 && BINARY_ON_NUMBERS.has(tokenText(value[1]) || value[1].rawText)) {
		const [a, b] = [value[0], value[2]];
		const text = a.kind === 'stringLiteral' ? a : b.kind === 'stringLiteral' ? b : undefined;
		const other = text === a ? b : a;
		const numberOrDate = other.kind === 'integerLiteral' || other.kind === 'floatLiteral' || other.kind === 'dateLiteral';
		if (text && numberOrDate && !/\d/.test(stringLiteralValue(text.rawText))) {
			return { detail: `${text.rawText} is no number, so ${value[1].rawText} cannot take it beside ${other.rawText}`, span: { start: base + value[0].start, end: base + value[2].end } };
		}
	}
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
	outerNames: NameLookup,
	justAssigned: Map<string, Typed>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	memberTarget?: (span: Span) => AssignmentTarget | undefined,
	reached: (node: BodyNode | undefined) => void = () => undefined,
): void {
	// Whether each enclosing With names a sheet, innermost last (issue #411).
	const withSheets: boolean[] = [];
	const names: NameLookup = Object.assign(
		(lower: string): Typed | undefined => (lower === WITH_SHEET
			? (withSheets[withSheets.length - 1] ? { value: 0, type: 'long' as const } : undefined)
			: outerNames(lower)),
		outerNames.declares ? { declares: outerNames.declares } : {},
		outerNames.wholeSheetCells ? { wholeSheetCells: outerNames.wholeSheetCells } : {},
	);
	const withNamesSheet = (node: BodyNode): boolean => {
		const header = blockHeaderStatements(source, node).before;
		const toks = header ? statementTokens(source, header.span).filter((tok) => tok.kind !== 'comment') : [];
		const segments = tokenText(toks[0]) === 'with' ? chainSegments(toks.slice(1)) : undefined;
		return segments !== undefined && segments.length > 0 && namesSheet(segments, names);
	};
	// Every name a block mentions, its own lines included: `For i = ...` and
	// `If Store(k, n) Then` change what they name as well (issue #237).
	const touchedIn = (node: BodyNode): Set<string> => namesIn(source, node.span);
	const counters = loopCountersAt(source, proc.body, activity);
	const forget = (touched: ReadonlySet<string>): void => {
		for (const lower of touched) {
			forgetName(justAssigned, lower);
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
					reached(node);
					checkStatement(source, before.span, env, names, push, memberTarget);
				}
				// Each If and ElseIf condition, from the state the block is
				// entered with: `If d And 1 Then` (issue #407).
				if (node.kind === 'IfBlock') {
					for (const header of blockHeaderLeaves(source, node)) {
						checkStatement(source, header.span, env, names, push, memberTarget);
					}
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
				} else if (node.kind === 'WithBlock') {
					withSheets.push(withNamesSheet(node));
					visit(node.body as BodyNode[], loopTouched);
					withSheets.pop();
				} else {
					visit(node.body as BodyNode[], isLoopBlock(node) ? new Set([...loopTouched, ...touched]) : loopTouched);
				}
				restore();
				forget(touched);
				if (after) {
					// A Loop line runs after the body, which may have changed it.
					reached(undefined);
					checkStatement(source, after.span, env, names, push, memberTarget);
				}
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			reached(node);
			const spans = statementAndBranchSpans(node);
			const straightLine = spans.length === 1 && !(node.kind === 'Statement' && node.singleLineIfBranches);
			if (!straightLine) {
				justAssigned.clear();
			}
			for (const span of spans) {
				// A loop counter on its first and last passes as well:
				// `For i = 32760 To 32770` then `CInt(i)` overflows on the last
				// (issue #263).
				let stored: ReturnType<typeof checkStatement>;
				checkEachCounterPass(source, span, counters.get(node), () => undefined, (values, report) => {
					if (values.size === 0) {
						stored = checkStatement(source, span, env, names, report, memberTarget);
						return;
					}
					const passNames: NameLookup = (lower) => {
						const value = values.get(lower);
						const type = value === undefined ? undefined : numericTypeOf(env.get(lower));
						return type ? { value: value!, type } : names(lower);
					};
					checkStatement(source, span, env, passNames, report, memberTarget);
				}, push);
				if (!straightLine) {
					continue;
				}
				// Any other mention of a tracked name (a ByRef pass, a label a
				// GoTo could reach) ends what is known about it.
				const toks = statementTokens(source, span);
				if (jumpTargetLabelDeclaration(source, span) || tokenText(toks[firstExecutableTokenIndex(toks)]) === 'gosub') {
					justAssigned.clear();
					continue;
				}
				for (const tok of toks) {
					const lower = tokenName(tok)?.toLowerCase();
					if (lower) {
						forgetName(justAssigned, lower);
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

/** A name, and every member path under it: `t` forgets `t.i`. */
function forgetName(justAssigned: Map<string, Typed>, lower: string): void {
	justAssigned.delete(lower);
	for (const key of justAssigned.keys()) {
		if (key.startsWith(`${lower}.`)) {
			justAssigned.delete(key);
		}
	}
}

/** What a statement assigns: a name, or a member path or array element with its declared type. */
interface AssignmentTarget {
	name: string;
	valueTokens: VbaToken[];
	asType?: string;
	/** An array element, whose value is not the name's. */
	element?: boolean;
}

/**
 * `t.i = value`, `t.v(2) = value` and `v(2) = value` (issue #253): a numeric
 * member of a Type value, or an element of an array variable or member.
 */
function memberAssignmentTarget(
	source: string,
	span: Span,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
): AssignmentTarget | undefined {
	const toks = statementTokens(source, span);
	const first = firstExecutableTokenIndex(toks);
	const at = tokenText(toks[first]) === 'let' ? first + 1 : first;
	const lower = tokenName(toks[at])?.toLowerCase();
	const variable = lower ? variableSymbolIn(symbols, proc, lower) : undefined;
	if (variable?.isArray && toks[at + 1]?.rawText === '(') {
		const close = matchParenFrom(toks, at + 1);
		return close > 0 && toks[close + 1]?.rawText === '='
			? { name: toks[at].rawText, valueTokens: toks.slice(close + 2), asType: variable.asType?.replace(/\(\s*\)\s*$/, ''), element: true }
			: undefined;
	}
	const root = variableRoot(toks, at, variable, types);
	const step = root ? fieldChain(toks, root, types).at(-1) : undefined;
	const end = step ? step.close ?? step.at : -1;
	if (!step || toks[end + 1]?.rawText !== '=' || (step.field.isArray && step.open === undefined) || (!step.field.isArray && !step.path)) {
		return undefined;
	}
	return { name: step.path ?? step.display, valueTokens: toks.slice(end + 2), asType: step.field.typeName, element: step.field.isArray };
}

/** The value a bare assignment provably stores, when the rule can tell. */
function checkStatement(
	source: string,
	span: Span,
	env: ReadonlyMap<string, string>,
	names: NameLookup,
	push: PushFn,
	memberTarget?: (span: Span) => AssignmentTarget | undefined,
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
	const bare: AssignmentTarget | undefined = bareAssignmentTarget(source, span) ?? memberTarget?.(span);
	if (bare) {
		const value = bare.valueTokens.filter((tok) => tok.kind !== 'comment');
		const folded = new TypedFolder(value, span.start, names).fold();
		const declared = bare.asType ?? env.get(bare.name.toLowerCase());
		const target = numericTypeOf(declared);
		if (isOverflow(folded)) {
			report(folded);
		} else if (folded?.pastDate !== undefined && target === 'date') {
			// A Date local holds it without raising (issue #405).
			if (!bare.element) {
				stored = { name: bare.name.toLowerCase(), value: folded };
			}
		} else if (folded && !target && !bare.element && env.has(bare.name.toLowerCase()) && (normalizeType(declared) ?? 'variant') === 'variant') {
			// A Variant holds the value with its own type: `v = 2147483648#`
			// then `v Mod 7` converts a Double to Long and raises 6 (issue #480).
			stored = { name: bare.name.toLowerCase(), value: { ...folded, variant: true } };
		} else if (folded && target) {
			const kept = storedValue(folded, target);
			if (!inRange(kept.value, target, kept.exact)) {
				const shown = kept.exact !== undefined ? String(kept.exact) : showNumber(kept.value);
				const rounded = kept.value !== folded.value ? ` (${showNumber(folded.value)} rounds to ${showNumber(kept.value)})` : '';
				const label = normalizeType(declared) === 'longptr' ? 'LongPtr, which holds no more than a LongLong' : RANGES[target].label;
				const into = bare.element ? `an element of '${bare.name}'` : `'${bare.name}'`;
				// A Variant's number past the Date range is a Type mismatch, where
				// a typed one's is an Overflow (issue #329, measured in Excel 16.0).
				const variantDate = target === 'date' && folded.variant === true;
				push(variantDate ? 'assignmentTypeMismatch' : 'arithmeticOverflow', `Assignment to ${into} stores ${shown}${rounded} in ${article(label)} ${label}, whose range is ${rangeText(target)}. This will raise Run-time error ${variantDate ? "'13': Type mismatch" : "'6': Overflow"}.`, {
					start: span.start + value[0].start,
					end: span.start + value[value.length - 1].end,
				});
				return undefined;
			}
			if (!bare.element) {
				stored = { name: bare.name.toLowerCase(), value: { value: kept.value, type: target, ...(kept.exact !== undefined ? { exact: kept.exact } : {}) } };
			}
		}
	}
	// A condition is evaluated whole: `If d And 1 Then` with d past the Long
	// range converts it for And and raises 6 (issue #407, measured in Excel
	// 16.0), as `x = d And 1` does.
	const doCondition = (head === 'do' || head === 'loop') && ['while', 'until'].includes(tokenText(toks[first + 1]));
	const conditionFrom = head === 'if' || head === 'elseif' || head === 'while' ? first + 1 : doCondition ? first + 2 : -1;
	if (conditionFrom > 0) {
		const then = toks.findIndex((tok, k) => k >= conditionFrom && tokenText(tok) === 'then');
		const condition = toks.slice(conditionFrom, then < 0 ? toks.length : then).filter((tok) => tok.kind !== 'comment');
		const folded = condition.length > 0 ? new TypedFolder(condition, span.start, names).fold() : undefined;
		if (isOverflow(folded)) {
			report(folded);
		}
	}
	// Every other part the statement evaluates on its own: a call's
	// arguments, an operand of & or a comparison, an array index, a
	// conversion anywhere (issue #232). `Main = CStr(CInt(40000))` and
	// `IIf(True, 0, CInt(40000))` overflow as `Main = CInt(40000)` does.
	const reportPastDate = (folded: Typed, at: Span): void => {
		const key = `${at.start}:${at.end}`;
		if (!reported.has(key)) {
			reported.add(key);
			push('runtimeArgumentValue', `${folded.pastDate} is ${showNumber(folded.value)}, a Date outside the Date range that VBA holds without raising; reading it here as text or as a date raises Run-time error '5': Invalid procedure call or argument.`, at);
		}
	};
	checkParts(toks, span.start, names, report, reportPastDate);
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
function checkParts(
	toks: readonly VbaToken[],
	base: number,
	names: NameLookup,
	report: (folded: Overflow) => void,
	reportPastDate?: (folded: Typed, span: Span) => void,
	callee?: string,
): void {
	let from = 0;
	const part = (to: number): void => {
		const start = from;
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
			if (folded.pastDate !== undefined && readsPastDate(toks, start, to, callee)) {
				reportPastDate?.(folded, { start: base + piece[0].start, end: base + piece[piece.length - 1].end });
			}
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
			checkParts(piece.slice(i + 1, close), base, names, report, reportPastDate, tokenText(piece[i - 1]));
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
 * The built-ins that raise 5 when given a Date past the Date range (issue
 * #405, measured in Excel 16.0). CDate, CVar, CLng, CDbl, CCur, IsDate,
 * IsNumeric, TypeName, VarType and a comparison take it and run.
 */
const PAST_DATE_READERS: ReadonlySet<string> = new Set([
	'cstr', 'str', 'format', 'formatdatetime', 'year', 'month', 'day', 'weekday', 'hour', 'minute', 'second',
	'dateadd', 'datepart', 'datediff', 'datevalue', 'ucase', 'mid', 'left', 'val', 'instr', 'replace',
]);

/**
 * Whether the part `toks[start..to)` is read as text or as a date: an
 * argument of one of the built-ins above, an operand of `&`, or what
 * `Debug.Print` prints.
 */
function readsPastDate(toks: readonly VbaToken[], start: number, to: number, callee: string | undefined): boolean {
	if (callee !== undefined && PAST_DATE_READERS.has(callee)) {
		return true;
	}
	if (toks[start - 1]?.rawText === '&' || toks[to]?.rawText === '&') {
		return true;
	}
	return tokenText(toks[start - 1]) === 'print' && toks[start - 2]?.rawText === '.' && tokenText(toks[start - 3]) === 'debug';
}

/**
 * `For i = 1 To 32767` with i an Integer: the counter is incremented past its
 * last value before the exit test, and the increment overflows (measured:
 * `To 32767` raises, `To 32766` runs; `For b = 0 To 255` raises for a Byte).
 */
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
	// The header with any lines a ` _` continues it onto (issue #289).
	const header = blockHeaderLineSpan(source, node.span);
	const toks = statementTokensAfterLeadingLabel(source, header);
	const to = toks.findIndex((tok) => tokenText(tok) === 'to');
	if (to < 0) {
		return;
	}
	const step = toks.findIndex((tok) => tokenText(tok) === 'step');
	const limitToks = toks.slice(to + 1, step > 0 ? step : toks.length).filter((tok) => tok.kind !== 'comment');
	// The For line converts its start, limit and step to the counter's type
	// as it runs, before the first pass (issue #263, measured in Excel 16.0):
	// `For b = 5 To 3 Step -1` on a Byte raises 6 there, and so does a limit
	// past the type, whatever Exit For the body holds.
	if (type !== 'longlong') {
		const eqAt = toks.findIndex((tok) => tok.rawText === '=');
		const parts: Array<[string, readonly VbaToken[]]> = [
			['start', eqAt > 0 ? toks.slice(eqAt + 1, to) : []],
			['limit', limitToks],
			['step', step > 0 ? toks.slice(step + 1) : []],
		];
		for (const [which, part] of parts) {
			const value = part.filter((tok) => tok.kind !== 'comment');
			const folded = value.length === 0 ? undefined : new TypedFolder(value, header.start, names).fold();
			if (!folded || isOverflow(folded) || inRange(bankersRound(folded.value), type)) {
				continue;
			}
			push(
				'forCounterOverflow',
				`Counter '${node.controlVariable}' is ${RANGES[type].label}, and the For line converts its ${which} ${folded.value} to ${RANGES[type].label} as it starts, which does not fit. This will raise Run-time error '6': Overflow.`,
				{ start: header.start + value[0].start, end: header.start + value[value.length - 1].end },
			);
			return;
		}
	}
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

