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
//  - for-counter-overflow: `For i = 1 To 32767` with i an Integer, and
//    `For b = 0 To 255` with b a Byte: the increment after the last pass
//    overflows the counter. `To 32766` runs.
//
// The folder follows MS-VBAL 5.6.9.3: Byte and Integer operands make Integer
// results, Long makes Long, Single and Double make Double, Currency makes
// Currency; `/` and `^` make Double. A value the folder cannot type stays
// unknown and nothing is reported for it.

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
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';
import {
	knownLocalLiteralValues,
	normalizeType,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	firstExecutableTokenIndex,
	forEachVariableGroup,
	matchParenFrom,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

type NumericType = 'byte' | 'integer' | 'long' | 'single' | 'double' | 'currency' | 'date';

interface Typed {
	value: number;
	type: NumericType;
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
	single: { min: -3.402823e38, max: 3.402823e38, label: 'Single' },
	double: { min: -1.7976931348623157e308, max: 1.7976931348623157e308, label: 'Double' },
	currency: { min: -922337203685477.5807, max: 922337203685477.5807, label: 'Currency' },
	date: { min: -657434, max: 2958465, label: 'Date' },
};

const RANK: Readonly<Record<NumericType, number>> = {
	byte: 0, integer: 1, long: 2, single: 3, double: 4, currency: 5, date: 6,
};

/** The result type of `a op b` for + - * \ Mod (MS-VBAL 5.6.9.3). */
function arithmeticResultType(a: NumericType, b: NumericType): NumericType {
	if (a === 'currency' || b === 'currency') {
		return (a === 'double' || b === 'double' || a === 'single' || b === 'single') ? 'double' : 'currency';
	}
	if (a === 'date' || b === 'date') {
		return 'double';
	}
	const wider = RANK[a] >= RANK[b] ? a : b;
	return wider === 'byte' ? 'integer' : wider;
}

function isOverflow(folded: Folded): folded is Overflow {
	return folded !== undefined && 'overflow' in folded;
}

function inRange(value: number, type: NumericType): boolean {
	const range = RANGES[type];
	return Number.isFinite(value) && value >= range.min && value <= range.max;
}

/** VBA's rounding to a whole number: banker's rounding at .5. */
function bankersRound(value: number): number {
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

/** A literal's natural type and value: 3 is Integer, 40000 is Long, 3000000000 is Double. */
function literalTyped(tok: VbaToken): Typed | undefined {
	if (tok.kind === 'integerLiteral') {
		const raw = tok.rawText;
		const suffix = /[%&^]$/.exec(raw)?.[0];
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
		if (suffix === '^') {
			return undefined; // LongLong is not modelled here
		}
		if (/^&[hHoO]/.test(raw)) {
			// A hex or octal literal of four hex digits (or fewer) is an Integer
			// with 16-bit wraparound: &H8000 is -32768, &HFFFF is -1.
			const digits = raw.replace(/^&[hHoO]/, '').replace(/[%&^]$/, '');
			const isHex = /^&[hH]/.test(raw);
			const fits16 = isHex ? digits.length <= 4 : value <= 0xFFFF;
			if (fits16 && value > 32767) {
				return { value: value - 65536, type: 'integer' };
			}
			if (fits16) {
				return { value, type: 'integer' };
			}
			if (value > 2147483647 && value <= 0xFFFFFFFF) {
				return { value: value - 4294967296, type: 'long' };
			}
			return { value, type: 'long' };
		}
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
	return undefined;
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
	) {}

	fold(): Folded {
		if (this.toks.length === 0) {
			return undefined;
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

	private multiplicative(): Folded {
		const start = this.index;
		let left = this.unary();
		while (left !== undefined && !isOverflow(left)) {
			const op = this.toks[this.index];
			const word = op ? tokenText(op) : '';
			if (!op || !(op.rawText === '*' || op.rawText === '/' || op.rawText === '\\' || word === 'mod' || op.rawText === '^')) {
				break;
			}
			this.index++;
			const right = this.unary();
			if (right === undefined || isOverflow(right)) {
				return right;
			}
			left = this.combine(left, right, op.rawText === '^' ? '^' : word === 'mod' ? 'mod' : op.rawText, start, this.index - 1);
		}
		return left;
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
			const value = -operand.value;
			if (!inRange(value, operand.type)) {
				return { overflow: true, span: this.span(start, this.index - 1), detail: `Negating ${showNumber(operand.value)} gives ${showNumber(-operand.value)}, which does not fit ${RANGES[operand.type].label}` };
			}
			return { value, type: operand.type };
		}
		return this.primary();
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
			const inner = new TypedFolder(this.toks.slice(this.index + 1, close), this.base, this.names);
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
		if (this.toks[calleeIndex + 1]?.rawText === '(' && CONVERSIONS.has(callee)) {
			const close = matchParenFrom(this.toks, calleeIndex + 1);
			if (close < 0) {
				return undefined;
			}
			const inner = new TypedFolder(this.toks.slice(calleeIndex + 2, close), this.base, this.names).fold();
			const start = this.index;
			this.index = close + 1;
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

	private convert(callee: string, inner: Typed, span: Span): Folded {
		const target = CONVERSIONS.get(callee)!;
		if (target === 'abs') {
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
		if (target === 'hex' || target === 'oct') {
			// Hex and Oct take a value that fits a Long (or a LongLong on 64-bit
			// for whole numbers; 1E+20 fits neither).
			return inRange(inner.value, 'long') || (Number.isInteger(inner.value) && Math.abs(inner.value) < 9.2e18)
				? undefined
				: { overflow: true, span, detail: `${callee === 'hex' ? 'Hex' : 'Oct'}(${inner.value}) takes a value outside the Long range` };
		}
		const type = target as NumericType;
		const value = type === 'single' || type === 'double' || type === 'currency' || type === 'date' ? inner.value : bankersRound(inner.value);
		return inRange(value, type)
			? { value, type }
			: { overflow: true, span, detail: `${CONVERSION_NAMES[callee]}(${showNumber(inner.value)}) does not fit ${RANGES[type].label}` };
	}

	private combine(left: Typed, right: Typed, op: string, from: number, to: number): Folded {
		const span = this.span(from, to);
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
				type = arithmeticResultType(left.type, right.type);
				const a = bankersRound(left.value);
				const b = bankersRound(right.value);
				if (b === 0) {
					return undefined;
				}
				value = op === 'mod' ? a % b : Math.trunc(a / b);
				break;
			}
			default:
				type = arithmeticResultType(left.type, right.type);
				value = op === '+' ? left.value + right.value : op === '-' ? left.value - right.value : left.value * right.value;
				break;
		}
		if (!inRange(value, type)) {
			const shown = showNumber(value);
			return {
				overflow: true,
				span,
				detail: `${describe(left)} ${op === 'mod' ? 'Mod' : op} ${describe(right)} is ${shown}, outside the ${RANGES[type].label} range`,
			};
		}
		return { value, type };
	}
}

function describe(typed: Typed): string {
	return `${showNumber(typed.value)} (${RANGES[typed.type].label})`;
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

const CONVERSIONS: ReadonlyMap<string, NumericType | 'abs' | 'int' | 'fix' | 'exp' | 'hex' | 'oct'> = new Map([
	['cbyte', 'byte'], ['cint', 'integer'], ['clng', 'long'], ['csng', 'single'], ['cdbl', 'double'],
	['ccur', 'currency'], ['cdate', 'date'], ['abs', 'abs'], ['int', 'int'], ['fix', 'fix'], ['exp', 'exp'],
	['hex', 'hex'], ['oct', 'oct'],
]);

const CONVERSION_NAMES: Readonly<Record<string, string>> = {
	cbyte: 'CByte', cint: 'CInt', clng: 'CLng', csng: 'CSng', cdbl: 'CDbl', ccur: 'CCur', cdate: 'CDate',
};

function numericTypeOf(asType: string | undefined): NumericType | undefined {
	const normalized = normalizeType(asType);
	if (!normalized) {
		return undefined;
	}
	if (normalized in RANGES) {
		return normalized as NumericType;
	}
	return undefined;
}

/**
 * The Consts a procedure can name, folded with their declared or natural
 * type: `Private Const HOURS As Integer = 24` is an Integer 24, and
 * `Const K = 40000` a Long. A Const the folder cannot fold is left out.
 */
function constantLookup(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): Map<string, Typed> {
	const out = new Map<string, Typed>();
	const candidates: VbaSymbol[] = [
		...(projectVisibleSymbols ?? []),
		...(symbols.root.children ?? []),
		...(proc ? procedureSymbolFor(symbols, proc)?.children ?? [] : []),
	];
	// Later entries shadow earlier ones: the procedure's own Consts win.
	const pending = new Map<string, VbaSymbol>();
	for (const symbol of candidates) {
		if (symbol.kind === 'constant' && symbol.defaultRaw !== undefined) {
			pending.set(symbol.name.toLowerCase(), symbol);
		}
	}
	const resolving = new Set<string>();
	const resolve = (lower: string): Typed | undefined => {
		if (out.has(lower)) {
			return out.get(lower);
		}
		const symbol = pending.get(lower);
		if (!symbol || resolving.has(lower)) {
			return undefined;
		}
		resolving.add(lower);
		const toks = statementTokens(symbol.defaultRaw!, { start: 0, end: symbol.defaultRaw!.length })
			.filter((tok) => tok.kind !== 'comment');
		const folded = new TypedFolder(toks, 0, resolve).fold();
		resolving.delete(lower);
		if (folded === undefined || isOverflow(folded)) {
			return undefined;
		}
		const declared = numericTypeOf(symbol.asType);
		const typed: Typed = declared ? { value: folded.value, type: declared } : folded;
		if (declared && !inRange(bankersRound(folded.value), declared)) {
			return undefined;
		}
		out.set(lower, typed);
		return typed;
	};
	for (const lower of pending.keys()) {
		resolve(lower);
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
	const moduleConstants = constantLookup(symbols, undefined, projectVisibleSymbols);
	checkConstDeclarations(source, mod.members.filter((m): m is VariableGroupNode => m.kind === 'VariableGroup'), moduleConstants, activity, push);
	const hostValues = hostConstantValues(hostModel);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const constants = constantLookup(symbols, member, projectVisibleSymbols);
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
			const folded = new TypedFolder(value, decl.span.start, (lower) => constants.get(lower)).fold();
			if (isOverflow(folded)) {
				push('constOverflow', `Const '${decl.name}' overflows while it is evaluated: ${folded.detail}. This is a VBE compile error: Overflow.`, folded.span);
				continue;
			}
			const declared = numericTypeOf(decl.asType);
			if (folded && declared && !inRange(declared === 'single' || declared === 'double' || declared === 'currency' || declared === 'date' ? folded.value : bankersRound(folded.value), declared)) {
				push('constOverflow', `Const '${decl.name}' is declared As ${RANGES[declared].label} but its value ${folded.value} is outside that range. This is a VBE compile error: Overflow.`, { start: decl.span.start + value[0].start, end: decl.span.start + value[value.length - 1].end });
			}
		}
	}
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
	const visit = (body: readonly BodyNode[], topLevel: boolean): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'ForBlock') {
				checkForCounter(source, node, env, names, push);
			}
			if ('body' in node && Array.isArray(node.body)) {
				// A block may run any number of times: nothing stored before it
				// is known after it, and nothing inside it is straight-line.
				justAssigned.clear();
				visit(node.body as BodyNode[], false);
				justAssigned.clear();
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			const spans = statementAndBranchSpans(node);
			const straightLine = topLevel && spans.length === 1 && !(node.kind === 'Statement' && node.singleLineIfBranches);
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
	visit(proc.body, true);
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
	const bare = bareAssignmentTarget(source, span);
	if (bare) {
		const value = bare.valueTokens.filter((tok) => tok.kind !== 'comment');
		const folded = new TypedFolder(value, span.start, names).fold();
		if (isOverflow(folded)) {
			push('arithmeticOverflow', `${folded.detail}. This will raise Run-time error '6': Overflow.`, folded.span);
			return undefined;
		}
		const target = numericTypeOf(env.get(bare.name.toLowerCase()));
		if (folded && target) {
			const stored = target === 'single' || target === 'double' || target === 'currency' || target === 'date' ? folded.value : bankersRound(folded.value);
			if (!inRange(stored, target)) {
				const rounded = stored !== folded.value ? ` (${showNumber(folded.value)} rounds to ${showNumber(stored)})` : '';
				push('arithmeticOverflow', `Assignment to '${bare.name}' stores ${showNumber(stored)}${rounded} in ${article(RANGES[target].label)} ${RANGES[target].label}, whose range is ${RANGES[target].min} to ${RANGES[target].max}. This will raise Run-time error '6': Overflow.`, {
					start: span.start + value[0].start,
					end: span.start + value[value.length - 1].end,
				});
				return undefined;
			}
			return { name: bare.name.toLowerCase(), value: { value: stored, type: target } };
		}
		return undefined;
	}
	// Conversion calls anywhere else in the statement: `Debug.Print CInt(40000)`.
	for (let i = 0; i + 1 < toks.length; i++) {
		const callee = tokenText(toks[i]);
		if (!CONVERSIONS.has(callee) || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const start = toks[i - 1]?.rawText === '.' ? i - 2 : i;
		const folded = new TypedFolder(toks.slice(start, close + 1), span.start, names).fold();
		if (isOverflow(folded)) {
			push('arithmeticOverflow', `${folded.detail}. This will raise Run-time error '6': Overflow.`, folded.span);
		}
	}
	return undefined;
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
	const overflows = stepValue.value > 0 ? last + stepValue.value > range.max : last + stepValue.value < range.min;
	if (!overflows) {
		return;
	}
	push(
		'forCounterOverflow',
		`Counter '${node.controlVariable}' is ${range.label}; after its last pass at ${last} the loop adds ${stepValue.value}, which does not fit. This will raise Run-time error '6': Overflow.`,
		{ start: header.start + limitToks[0].start, end: header.start + limitToks[limitToks.length - 1].end },
	);
}

