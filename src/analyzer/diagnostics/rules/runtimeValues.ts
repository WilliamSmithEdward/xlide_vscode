// Rule family: deterministic runtime argument/conversion values (audit #0).
//
// Extracted verbatim from analyzeModule.ts: constant arguments that are
// provably outside a runtime function's accepted range and conversions of
// provably invalid literals.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { HostObjectModel } from '../../host/excelObjectModel';
import {
	evaluateIntegerConstantExpression,
	type IntegerConstantLookup,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
} from '../../constants/integerConstantExpression';
import type { VbaToken } from '../../lexer/tokenKinds';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type {
	BodyNode,
	ModuleNode,
	Span,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import {
	type CallableTypeSignature,
	emptyArgSplit,
	splitArgSlots,
} from '../callExtraction';
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import { checkEachCounterPass, loopCountersAt } from '../loopCounters';
import { foldKnownStringCalls, moduleCompare, type KnownStringCallContext } from '../knownStringCalls';
import { bankersRound, isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
import { isInvalidBooleanString, isInvalidDateString, isInvalidNumericString } from '../stringConversion';
import {
	callableTypeSignaturesFor,
	isKnownScalarType,
	knownLocalLiteralValuesAt,
	type KnownLocalValue,
	namedArgumentSlot,
	normalizeType,
	procedureIntegerConstantLookup,
	runtimeCallableSourceShadowed,
	type SourceNameScope,
	sourceNameScopeFor,
	stringLiteralValue,
	typeEnvironmentFor,
	unwrapOuterParens,
} from '../typeInference';
import {
	matchParenFrom,
	rawExpressionTokens,
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

const NO_COUNTER_VALUES: ReadonlyMap<string, number> = new Map();

/** A local's declared type, and its dimensions: 0 for a scalar, the count for a fixed array. */
interface LocalDeclaration {
	asType: string;
	dimensions: number;
}

interface RuntimeArgumentValueSpec {
	canonicalName: string;
	parameterName: string;
	argumentIndex: number;
	minimum?: number;
	maximum?: number;
	/** The value must be strictly above this: Log(0) raises, Log(0.5) runs. */
	exclusiveMinimum?: number;
	/** Single values inside the range that still raise: InStrRev's Start of 0. */
	disallowed?: readonly number[];
	/** A Double parameter, compared as passed: Sqr(-0.4) raises. Others round first. */
	fractional?: boolean;
	/** Whether a whole number is accepted, where no range says it: StrConv's Conversion. */
	accepts?: (value: number) => boolean;
	/** An empty string literal raises: Asc(""), String(3, ""). */
	emptyStringRaises?: boolean;
	/** A string literal must be one of these (case-insensitive): DateAdd's interval. */
	allowedStrings?: readonly string[];
	minimumSlotCount?: number;
	allowNamed?: boolean;
	/** Which `$`-suffixed spelling, if any, the function also has. */
	stringSuffix?: boolean;
	/**
	 * The parameter's type, where no signature gives it: a value outside it
	 * raises error 6, Overflow, before any bound is read (issue #218):
	 * TimeSerial(32768, 0, 0), ChrB(256), String(1E+10, "a"). Without it, a
	 * value past the Long range is left to argument-type-mismatch.
	 */
	overflowType?: 'Byte' | 'Integer' | 'Long';
	/** The bounds raise error 6 rather than 5: Error(65536) (issue #218). */
	boundsOverflow?: boolean;
}

interface RuntimeArgumentValueHit {
	displayName: string;
	parameterName: string;
	value: number | string;
	span: Span;
	/** 6 for Overflow; 5 otherwise. */
	error?: 6;
	/** The whole message, for a check that is not one argument's bound. */
	message?: string;
}

const OVERFLOW_RANGES: Readonly<Record<'Byte' | 'Integer' | 'Long', { min: number; max: number }>> = {
	Byte: { min: 0, max: 255 },
	Integer: { min: -32768, max: 32767 },
	Long: { min: -2147483648, max: 2147483647 },
};

/**
 * Rule: some runtime-library arguments have deterministic value bounds even
 * when the argument type itself is valid. This slice is VBE-oracle-backed for
 * integer bounds on selected string runtime functions, which compile but raise
 * Run-time error 5 when the value is outside the proven range.
 */
export function checkRuntimeArgumentValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	hostModel?: HostObjectModel,
): ProcedureStatementVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	const host = hostModel?.hostName?.toLowerCase();
	const compare = moduleCompare(source);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const constants = procedureIntegerConstantLookup(
			member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel,
		);
		// A local the procedure never assigns holds its default, and one whose
		// every assignment is one literal holds that (issue #118): `Asc(s)` with
		// s never assigned is Asc(""), and `Mid(s, 5, 1) = "x"` after s = "abc"
		// starts past the end.
		// Or the literal the last assignment before the statement stores
		// (issue #180).
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		let known: ReadonlyMap<string, KnownLocalValue> = new Map();
		// A local declared as a scalar, which `Join(n)` refuses (issue #239),
		// or as a fixed array, whose element type and dimensions Join reads.
		const locals = procedureSymbolFor(symbols, member)?.children ?? [];
		const declarationOf = (lower: string): LocalDeclaration | undefined => {
			const local = locals.find((child) => child.name.toLowerCase() === lower);
			if (local?.kind !== 'localVariable') {
				return undefined;
			}
			const type = normalizeType(local.asType);
			if (local.isArray) {
				return local.arrayBounds === undefined
					? undefined // a dynamic array may still be unallocated, which Join takes
					: { asType: local.asType ?? 'Variant', dimensions: splitTopLevelTokenGroups(rawExpressionTokens(local.arrayBounds), 0, ',').length };
			}
			return type !== undefined && type !== 'variant' && isKnownScalarType(type) ? { asType: local.asType!, dimensions: 0 } : undefined;
		};
		const stringsFor = new Map<ReadonlyMap<string, KnownLocalValue>, { strings: Map<string, string>; lengths: Map<string, number> }>();
		const stringsAt = (values: ReadonlyMap<string, KnownLocalValue>): { strings: Map<string, string>; lengths: Map<string, number> } => {
			let out = stringsFor.get(values);
			if (!out) {
				out = { strings: new Map(), lengths: new Map() };
				for (const [lower, value] of values) {
					if (value.kind === 'string') {
						out.lengths.set(lower, (value.value as string).length);
						if (!value.contentMutated) {
							out.strings.set(lower, value.value as string);
						}
					}
				}
				stringsFor.set(values, out);
			}
			return out;
		};
		// A loop counter bound to one pass's value (issue #200).
		let counterValues = NO_COUNTER_VALUES;
		const counters = loopCountersAt(source, member.body, activity);
		const lookup: IntegerConstantLookup = {
			get: (name) => {
				const counter = counterValues.size === 0 ? undefined : counterValues.get(name.toLowerCase());
				if (counter !== undefined) {
					return counter;
				}
				const constant = constants.get(name);
				if (constant !== undefined) {
					return constant;
				}
				const local = known.get(name.toLowerCase());
				return local?.kind === 'number' && Number.isInteger(local.value) ? (local.value as number) : undefined;
			},
		};
		return (stmt) => {
			known = valuesAt(stmt);
			const { strings: knownStrings, lengths: knownStringLengths } = stringsAt(known);
			// A bound of Len(s) reads the length s has as the loop starts.
			const atomValue = (atom: { kind: string; name: string }, counter: { loopNode: BodyNode }): number | undefined =>
				atom.kind === 'len' ? stringsAt(valuesAt(counter.loopNode)).lengths.get(atom.name) : undefined;
			checkEachCounterPass(source, stmt.span, counters.get(stmt), atomValue, (values, report) => {
				counterValues = values;
				const stringCalls: KnownStringCallContext = {
					knownStrings,
					integerValue: (text) => evaluateIntegerConstantExpression(text, lookup),
					shadowed: (name) => runtimeCallableSourceShadowed(name, sourceNames),
					compare,
				};
				for (const hit of runtimeArgumentValueHits(source, stmt.span, moduleSignatures, env, lookup, stringCalls, sourceNames, host, declarationOf)) {
					const raises = hit.error === 6 ? `'6': Overflow` : `'5': Invalid procedure call or argument`;
					report(
						'runtimeArgumentValue',
						hit.message ?? `Argument '${hit.parameterName}' of '${hit.displayName}' is ${hit.value}; this will raise Run-time error ${raises}.`,
						hit.span,
					);
				}
				for (const hit of runtimeStatementValueHits(source, stmt.span, lookup, knownStringLengths, knownStrings, sourceNames)) {
					report('runtimeArgumentValue', hit.message, hit.span);
				}
			}, push);
			counterValues = NO_COUNTER_VALUES;
		};
	};
}

/**
 * Statement and operator forms that raise for a value the code states (issue
 * #118, each measured in Excel 16.0):
 *
 *  - `Err.Raise 0` and `Err.Raise 65536`, `Error 0`: error 5. A number is
 *    valid from 1 to 65535.
 *  - `(-8) ^ (1 / 3)` and `0 ^ -1`: error 5. A negative base takes only a
 *    whole exponent; zero takes only a non-negative one.
 *  - `"b" Like "[z-a]"` and `"b" Like "[a-"`: error 93, Invalid pattern
 *    string. A reversed range or an unterminated character list.
 */
function runtimeStatementValueHits(
	source: string,
	span: Span,
	constants: IntegerConstantLookup,
	knownStringLengths: ReadonlyMap<string, number>,
	knownStrings: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
): Array<{ message: string; span: Span }> {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const out: Array<{ message: string; span: Span }> = [];
	const at = (tok: VbaToken): Span => ({ start: span.start + tok.start, end: span.start + tok.end });
	const first = toks[0];
	// `Mid(s, 5, 1) = "x"` with s holding "abc": the statement form starts
	// past the end of the string, error 5 (issue #118). Only the length
	// matters, which an earlier Mid statement cannot have changed.
	if ((tokenText(first) === 'mid' || tokenText(first) === 'mid$') && toks[1]?.rawText === '(') {
		const close = matchParenFrom(toks, 1);
		if (close > 0 && toks[close + 1]?.rawText === '=') {
			const split = splitArgSlots(toks.slice(2, close), span.start);
			const target = split.slots[0]?.length === 1 ? tokenName(split.slots[0][0])?.toLowerCase() : undefined;
			const length = target !== undefined ? knownStringLengths.get(target) : undefined;
			const startSlot = split.slots[1];
			const start = startSlot ? integerGroupValue(source, span, startSlot, constants) : undefined;
			if (length !== undefined && start !== undefined && start > length) {
				out.push({
					message: `Mid statement start ${start} is past the end of ${split.slots[0][0].rawText}, which is ${length} character(s) long. This will raise Run-time error '5': Invalid procedure call or argument.`,
					span: split.spans[1] ?? at(toks[0]),
				});
			}
		}
	}
	// `Err.Raise n` and `Error n`.
	let numberIndex = -1;
	let form = '';
	if (tokenText(first) === 'err' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'raise') {
		numberIndex = 3;
		form = 'Err.Raise';
	} else if (tokenText(first) === 'error' && toks[1] !== undefined && !runtimeCallableSourceShadowed('Error', sourceNames)) {
		numberIndex = 1;
		form = 'Error';
	}
	if (numberIndex > 0) {
		const group = numberArgumentGroup(toks, numberIndex);
		const value = group ? integerGroupValue(source, span, group, constants) : undefined;
		// Err.Raise takes 1 to 65535 or any negative Long: `vbObjectError +
		// 513` and an HRESULT such as &H80004002 are how a class raises its
		// own errors (issue #142, measured). The Error statement takes only
		// 1 to 65535.
		const invalid = form === 'Err.Raise'
			? value !== undefined && (value === 0 || value > 65535 || value < -2147483648)
			: value !== undefined && (value < 1 || value > 65535);
		if (invalid) {
			const valid = form === 'Err.Raise' ? '1 to 65535, or a negative Long such as vbObjectError + n' : '1 to 65535';
			out.push({
				message: `${form} ${value} is not an error number: valid numbers are ${valid}. This will raise Run-time error '5': Invalid procedure call or argument.`,
				span: { start: span.start + group![0].start, end: span.start + group![group!.length - 1].end },
			});
		}
	}
	for (let i = 1; i < toks.length - 1; i++) {
		const tok = toks[i];
		if (tok.kind === 'operator' && tok.rawText === '^') {
			const base = numericOperandBefore(toks, i);
			const exponent = numericOperandAfter(toks, i);
			if (base !== undefined && exponent !== undefined) {
				if (base < 0 && !Number.isInteger(exponent)) {
					out.push({ message: `A negative number raised to the fractional power ${exponent} has no real value. This will raise Run-time error '5': Invalid procedure call or argument.`, span: at(tok) });
				} else if (base === 0 && exponent < 0) {
					out.push({ message: `Zero raised to the negative power ${exponent} divides by zero. This will raise Run-time error '5': Invalid procedure call or argument.`, span: at(tok) });
				}
			}
			continue;
		}
		if (tokenText(tok) === 'like' && toks[i + 1]?.kind === 'stringLiteral') {
			const pattern = stringLiteralValue(toks[i + 1].rawText);
			const problem = invalidLikePattern(pattern);
			// The matcher meets a bad list only with a character left to
			// compare, so the string matched decides it (issue #193).
			const subject = likeSubject(toks, i, knownStrings);
			if (problem && subject !== undefined && likeReachesBadList(subject, pattern)) {
				out.push({ message: `The Like pattern ${toks[i + 1].rawText} ${problem}, and matching ${JSON.stringify(subject)} reaches it. This will raise Run-time error '93': Invalid pattern string.`, span: at(toks[i + 1]) });
			}
		}
	}
	return out;
}

/** The tokens of the first argument after `index`, up to a top-level comma or the end. */
function numberArgumentGroup(toks: readonly VbaToken[], index: number): VbaToken[] | undefined {
	const group: VbaToken[] = [];
	let depth = 0;
	for (let k = index; k < toks.length; k++) {
		const raw = toks[k].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		} else if (raw === ',' && depth === 0) {
			break;
		}
		if (toks[k].kind !== 'comment') {
			group.push(toks[k]);
		}
	}
	return group.length > 0 ? group : undefined;
}

function integerGroupValue(
	source: string,
	span: Span,
	group: readonly VbaToken[],
	constants: IntegerConstantLookup,
): number | undefined {
	return evaluateIntegerConstantExpression(
		source.slice(span.start + group[0].start, span.start + group[group.length - 1].end),
		constants,
	);
}

/** A numeric literal (optionally signed and parenthesized) right before `index`. */
function numericOperandBefore(toks: readonly VbaToken[], index: number): number | undefined {
	let end = index - 1;
	if (toks[end]?.rawText === ')') {
		let depth = 0;
		let start = end;
		for (; start >= 0; start--) {
			if (toks[start].rawText === ')') { depth++; }
			if (toks[start].rawText === '(') { depth--; if (depth === 0) { break; } }
		}
		if (start < 0) { return undefined; }
		return numericLiteralGroupValue(toks.slice(start + 1, end));
	}
	return numericLiteralGroupValue([toks[end]]);
}

/** A numeric literal (optionally signed and parenthesized) right after `index`. */
function numericOperandAfter(toks: readonly VbaToken[], index: number): number | undefined {
	let start = index + 1;
	if (toks[start]?.rawText === '(') {
		const close = matchParenFrom(toks, start);
		if (close < 0) { return undefined; }
		return numericLiteralGroupValue(toks.slice(start + 1, close));
	}
	const group: VbaToken[] = [];
	if (toks[start]?.rawText === '-' || toks[start]?.rawText === '+') {
		group.push(toks[start]);
		start++;
	}
	if (toks[start]) { group.push(toks[start]); }
	return numericLiteralGroupValue(group);
}

/** The value of `[sign] literal`, or of `a / b` with literal operands. */
function numericLiteralGroupValue(group: readonly VbaToken[]): number | undefined {
	const toks = group.filter((tok) => tok.kind !== 'comment');
	const literal = (tok: VbaToken | undefined): number | undefined => {
		if (!tok) { return undefined; }
		if (tok.kind === 'integerLiteral') { return parseVbaIntegerLiteral(tok.rawText); }
		if (tok.kind === 'floatLiteral') {
			const value = Number(tok.rawText.replace(/[!#@]$/, ''));
			return Number.isFinite(value) ? value : undefined;
		}
		return undefined;
	};
	let sign = 1;
	let rest = toks;
	if (rest[0]?.rawText === '-' || rest[0]?.rawText === '+') {
		sign = rest[0].rawText === '-' ? -1 : 1;
		rest = rest.slice(1);
	}
	if (rest.length === 1) {
		const value = literal(rest[0]);
		return value === undefined ? undefined : sign * value;
	}
	if (rest.length === 3 && rest[1].rawText === '/') {
		const a = literal(rest[0]);
		const b = literal(rest[2]);
		return a === undefined || b === undefined || b === 0 ? undefined : sign * (a / b);
	}
	return undefined;
}

/**
 * The string Like matches at `likeIndex`: a string literal, or a local the
 * procedure makes plain, standing alone on its left.
 */
function likeSubject(toks: readonly VbaToken[], likeIndex: number, knownStrings: ReadonlyMap<string, string>): string | undefined {
	const operand = toks[likeIndex - 1];
	const before = toks[likeIndex - 2];
	const alone = before === undefined || before.rawText === '(' || before.rawText === ',' || before.rawText === '='
		|| ['if', 'elseif', 'while', 'until', 'and', 'or', 'not', 'then'].includes(tokenText(before));
	if (!operand || !alone) {
		return undefined;
	}
	if (operand.kind === 'stringLiteral') {
		return stringLiteralValue(operand.rawText);
	}
	const name = tokenName(operand)?.toLowerCase();
	return name ? knownStrings.get(name) : undefined;
}

/**
 * Whether matching `subject` against `pattern` reaches a malformed character
 * list with a character left to compare, which is when Like raises 93
 * (issue #193, measured in Excel 16.0): "xy" Like "?[" raises, "x" Like "?["
 * is False, and so is "zb" Like "a[z-a]", which fails at the "a". A `*`
 * before the bad list, or a comparison Option Compare Text could decide
 * otherwise, proves nothing.
 */
function likeReachesBadList(subject: string, pattern: string): boolean {
	let p = 0;
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		let matches: (c: string) => boolean | undefined;
		if (ch === '[') {
			const close = pattern.indexOf(']', i + 1);
			const body = close < 0 ? undefined : pattern.slice(i + 1, close);
			if (body === undefined || invalidLikePattern(`[${body}]`)) {
				return p < subject.length;
			}
			const negated = body.startsWith('!');
			const list = negated ? body.slice(1) : body;
			matches = (c) => {
				const exact = charListHas(list, c);
				const folded = charListHas(list.toLowerCase(), c.toLowerCase()) || charListHas(list.toUpperCase(), c.toUpperCase());
				if (exact !== folded) {
					return undefined;
				}
				return negated ? !exact : exact;
			};
			i = close;
		} else if (ch === '*') {
			return false;
		} else if (ch === '?') {
			matches = () => true;
		} else if (ch === '#') {
			matches = (c) => c >= '0' && c <= '9';
		} else {
			matches = (c) => (c === ch ? true : c.toLowerCase() === ch.toLowerCase() ? undefined : false);
		}
		if (p >= subject.length) {
			return false;
		}
		const verdict = matches(subject[p]);
		if (verdict !== true) {
			return false;
		}
		p++;
	}
	return false;
}

/** Whether a character list's body, ranges included, holds `c`. */
function charListHas(list: string, c: string): boolean {
	for (let k = 0; k < list.length; k++) {
		if (list[k + 1] === '-' && k + 2 < list.length) {
			if (c >= list[k] && c <= list[k + 2]) {
				return true;
			}
			k += 2;
		} else if (list[k] === c) {
			return true;
		}
	}
	return false;
}

/** Why a Like pattern raises error 93, or undefined when it is well formed. */
function invalidLikePattern(pattern: string): string | undefined {
	for (let i = 0; i < pattern.length; i++) {
		if (pattern[i] !== '[') {
			continue;
		}
		const close = pattern.indexOf(']', i + 1);
		if (close < 0) {
			return 'opens a character list it never closes';
		}
		let body = pattern.slice(i + 1, close);
		if (body.startsWith('!')) {
			body = body.slice(1);
		}
		for (let k = 1; k + 1 < body.length; k++) {
			if (body[k] === '-' && body[k - 1] > body[k + 1]) {
				return `has the reversed range ${body[k - 1]}-${body[k + 1]}`;
			}
		}
		i = close;
	}
	return undefined;
}

function runtimeArgumentValueHits(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
	constants: IntegerConstantLookup,
	stringCalls: KnownStringCallContext,
	sourceNames: SourceNameScope,
	host: string | undefined,
	declarationOf?: (lower: string) => LocalDeclaration | undefined,
): RuntimeArgumentValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeArgumentValueHit[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const call = runtimeArgumentValueCallAt(toks, i, span, moduleSignatures, env, sourceNames, host);
		if (!call) {
			continue;
		}
		for (const spec of call.specs) {
			const slot = runtimeArgumentValueSlot(call.slots, spec);
			const literal = slot
				? integerArgumentOutsideBounds(source, slot, span.start, spec, constants, stringCalls)
				: undefined;
			if (!literal) {
				continue;
			}
			hits.push({
				displayName: call.displayName,
				parameterName: spec.parameterName,
				value: literal.value,
				span: literal.span,
				...(literal.error === 6 ? { error: 6 as const } : {}),
			});
		}
		const overflow = dateAddPastMaximum(source, span, call, constants)
			?? dateSerialPastMaximum(source, span, call, constants)
			?? argumentRelationHit(source, span, call, constants, declarationOf);
		if (overflow) {
			hits.push(overflow);
		}
	}
	return hits;
}

/** A numeric argument's value: a signed literal, `a / b` of literals, or a constant expression. */
function numericSlotValue(
	source: string,
	span: Span,
	slot: readonly VbaToken[] | undefined,
	constants: IntegerConstantLookup,
): number | undefined {
	const toks = slot ? unwrapOuterParens(slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline')) : [];
	if (toks.length === 0 || namedArgumentSlot(toks)) {
		return undefined;
	}
	return numericLiteralGroupValue(toks) ?? integerGroupValue(source, span, toks, constants);
}

/**
 * Checks that read more than one argument, or an argument's kind rather than
 * its bound (issue #218, each measured in Excel 16.0):
 *
 *  - Partition(Number, Start, Stop, Interval) raises 5 when Start is below 0,
 *    Stop is not above Start, or Interval is below 1, each rounded half to
 *    even: Partition(5, 0, 10, 0.6) runs, 0.4 raises.
 *  - The financial functions raise 5: Pmt with NPer 0; IPmt and PPmt with NPer
 *    or Per not above 0, or Per a whole period past NPer (Per 10.5 of 10 runs,
 *    11 raises); SLN with Life 0; SYD and DDB with Life or Period not above 0,
 *    or Period past Life; DDB with Factor not above 0; NPer where its log has
 *    no value (Rate at or below -1, Rate and Pmt both 0, or a ratio not above
 *    0); Rate with NPer not above 0. PV with Rate -1 divides by zero, error 11.
 *    Whether Rate's iteration converges otherwise is not judged: Rate(10, 100,
 *    1000) raises and Rate(9, 100, 1000) returns -1.73.
 *  - LBound and UBound of Array(...) or Split(...), which have one dimension,
 *    raise 9 for any other Dimension.
 *  - Join and Filter given a string or number where the array goes raise 13.
 *  - Switch with an odd number of arguments raises 5 (issue #219).
 */
function argumentRelationHit(
	source: string,
	span: Span,
	call: { displayName: string; slots: VbaToken[][] },
	constants: IntegerConstantLookup,
	declarationOf: (lower: string) => LocalDeclaration | undefined = () => undefined,
): RuntimeArgumentValueHit | undefined {
	const name = call.displayName.replace(/^VBA\./i, '').replace(/\$$/, '').toLowerCase();
	if (call.slots.some((slot) => namedArgumentSlot(slot))) {
		return undefined;
	}
	const value = (index: number): number | undefined => numericSlotValue(source, span, call.slots[index], constants);
	const slotSpan = (from: number, to = from): Span => {
		const first = call.slots[from].find((t) => t.kind !== 'comment' && t.kind !== 'newline')!;
		const last = [...call.slots[to]].reverse().find((t) => t.kind !== 'comment' && t.kind !== 'newline')!;
		return { start: span.start + first.start, end: span.start + last.end };
	};
	const hit = (message: string, at: Span, error = 5): RuntimeArgumentValueHit => ({
		displayName: call.displayName,
		parameterName: '',
		value: '',
		span: at,
		message: `${message} This will raise Run-time error '${error}': ${error === 11 ? 'Division by zero' : error === 9 ? 'Subscript out of range' : error === 13 ? 'Type mismatch' : 'Invalid procedure call or argument'}.`,
	});
	const present = (count: number): boolean => call.slots.length >= count
		&& call.slots.slice(0, count).every((slot) => slot.some((t) => t.kind !== 'comment' && t.kind !== 'newline'));
	switch (name) {
		case 'partition': {
			if (!present(4)) {
				return undefined;
			}
			const [start, stop, interval] = [value(1), value(2), value(3)].map((v) => (v === undefined ? undefined : bankersRound(v)));
			if (start !== undefined && start < 0) {
				return hit(`Partition's Start is ${start}; it must be 0 or more.`, slotSpan(1));
			}
			if (start !== undefined && stop !== undefined && stop <= start) {
				return hit(`Partition's Stop, ${stop}, is not above its Start, ${start}.`, slotSpan(1, 2));
			}
			if (interval !== undefined && interval < 1) {
				return hit(`Partition's Interval is ${interval}; it must be 1 or more.`, slotSpan(3));
			}
			return undefined;
		}
		case 'pmt':
			return present(3) && value(1) === 0 ? hit('Pmt over 0 periods (NPer 0) has no payment.', slotSpan(1)) : undefined;
		case 'ipmt':
		case 'ppmt': {
			if (!present(4)) {
				return undefined;
			}
			const per = value(1);
			const nper = value(2);
			const label = call.displayName;
			if (nper !== undefined && nper <= 0) {
				return hit(`${label}'s NPer is ${nper}; it must be above 0.`, slotSpan(2));
			}
			if (per !== undefined && per <= 0) {
				return hit(`${label}'s Per is ${per}; it must be above 0.`, slotSpan(1));
			}
			if (per !== undefined && nper !== undefined && per >= nper + 1) {
				return hit(`${label}'s Per, ${per}, is past the last of its ${nper} periods.`, slotSpan(1));
			}
			return undefined;
		}
		case 'sln':
			return present(3) && value(2) === 0 ? hit('SLN over a Life of 0 has no depreciation.', slotSpan(2)) : undefined;
		case 'syd':
		case 'ddb': {
			if (!present(4)) {
				return undefined;
			}
			const life = value(2);
			const period = value(3);
			const label = call.displayName;
			if (life !== undefined && life <= 0) {
				return hit(`${label}'s Life is ${life}; it must be above 0.`, slotSpan(2));
			}
			if (period !== undefined && period <= 0) {
				return hit(`${label}'s Period is ${period}; it must be above 0.`, slotSpan(3));
			}
			if (period !== undefined && life !== undefined && period > life) {
				return hit(`${label}'s Period, ${period}, is past its Life, ${life}.`, slotSpan(3));
			}
			if (name === 'ddb' && present(5)) {
				const factor = value(4);
				if (factor !== undefined && factor <= 0) {
					return hit(`DDB's Factor is ${factor}; it must be above 0.`, slotSpan(4));
				}
			}
			return undefined;
		}
		case 'nper': {
			if (!present(3) || call.slots.length > 5) {
				return undefined;
			}
			const [rate, pmt, pv] = [value(0), value(1), value(2)];
			const fv = call.slots.length >= 4 && present(4) ? value(3) : 0;
			const type = call.slots.length >= 5 && present(5) ? value(4) : 0;
			if (rate === undefined || pmt === undefined || pv === undefined || fv === undefined || type === undefined) {
				return undefined;
			}
			if (rate === 0) {
				return pmt === 0 ? hit('NPer with a Rate and a Pmt of 0 has no number of periods.', slotSpan(0, 1)) : undefined;
			}
			if (rate <= -1) {
				return hit(`NPer's Rate is ${rate}; the logarithm of 1 + Rate has no value at or below -1.`, slotSpan(0));
			}
			const a = pmt * (1 + rate * (type !== 0 ? 1 : 0)) / rate;
			const ratio = (a - fv) / (a + pv);
			return Number.isFinite(ratio) && ratio > 0
				? undefined
				: hit('No number of periods brings these payments to this value: the logarithm NPer takes has no value.', slotSpan(0, Math.min(call.slots.length, 5) - 1));
		}
		case 'rate': {
			const nper = present(3) ? value(0) : undefined;
			return nper !== undefined && Number.isInteger(nper) && nper <= 0
				? hit(`Rate's NPer is ${nper}; it must be above 0.`, slotSpan(0))
				: undefined;
		}
		case 'pv': {
			if (!present(3)) {
				return undefined;
			}
			const [rate, nper] = [value(0), value(1)];
			return rate === -1 && nper !== undefined && nper > 0
				? hit('PV with a Rate of -1 divides by (1 + Rate) ^ NPer, which is 0.', slotSpan(0), 11)
				: undefined;
		}
		case 'lbound':
		case 'ubound': {
			if (call.slots.length !== 2 || !present(2)) {
				return undefined;
			}
			const array = call.slots[0].filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
			const callee = tokenName(array[0])?.toLowerCase();
			const oneDimension = (callee === 'array' || callee === 'split') && array[1]?.rawText === '(' && matchParenFrom(array, 1) === array.length - 1;
			const dimension = value(1);
			return oneDimension && dimension !== undefined && bankersRound(dimension) !== 1
				? hit(`${call.displayName}'s Dimension is ${dimension}, but ${array[0].rawText}(...) has one dimension.`, slotSpan(1), 9)
				: undefined;
		}
		case 'switch':
			// Switch takes condition-value pairs; an odd count compiles and
			// raises 5 whatever the conditions are (issue #219).
			return call.slots.length % 2 === 1
				? hit(`Switch takes its arguments in pairs, a condition and a value, but is given ${call.slots.length}.`, slotSpan(0, call.slots.length - 1))
				: undefined;
		case 'join':
		case 'filter': {
			// Join needs only its array (issue #239): `Join(5)`, `Join(Null)`
			// and Join of a Long or String local raise 13, as Filter of one
			// does (issue #242).
			const join = name === 'join';
			if (!present(join ? 1 : 2)) {
				return undefined;
			}
			const first = call.slots[0].filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
			const scalar = first.length === 1 && (first[0].kind === 'stringLiteral' || first[0].kind === 'integerLiteral' || first[0].kind === 'floatLiteral'
				|| (join && (first[0].kind === 'dateLiteral' || ['null', 'true', 'false'].includes(tokenText(first[0])))));
			if (scalar) {
				return hit(`${call.displayName} takes an array, but ${first[0].rawText} is not one.`, slotSpan(0), 13);
			}
			const declared = first.length === 1 && first[0].kind === 'identifier' ? declarationOf(first[0].rawText.toLowerCase()) : undefined;
			if (!declared) {
				return undefined;
			}
			if (declared.dimensions === 0) {
				return hit(`${call.displayName} takes an array, but '${first[0].rawText}' is declared As ${declared.asType}.`, slotSpan(0), 13);
			}
			// Join reads one dimension of Strings or Variants: an array of Long,
			// or of two dimensions, raises 5 (measured in Excel 16.0).
			// Filter refuses the same arrays with 13 (issue #242).
			const verb = join ? 'joins' : 'filters';
			const error = join ? 5 : 13;
			if (declared.dimensions > 1) {
				return hit(`${call.displayName} ${verb} an array of one dimension, but '${first[0].rawText}' has ${declared.dimensions}.`, slotSpan(0), error);
			}
			const element = normalizeType(declared.asType);
			return element !== 'string' && element !== 'variant'
				? hit(`${call.displayName} ${verb} Strings or Variants, but '${first[0].rawText}' is an array of ${declared.asType}.`, slotSpan(0), error)
				: undefined;
		}
		default:
			return undefined;
	}
}

/**
 * `DateAdd("d", 1, #12/31/9999#)`: adding to a date literal past the last
 * date VBA has (December 31, 9999) raises error 5 (issue #118). Only a date
 * literal with a positive whole-number count and a day, week, month or year
 * interval is decided here.
 */
function dateAddPastMaximum(
	source: string,
	span: Span,
	call: { displayName: string; slots: VbaToken[][] },
	constants: IntegerConstantLookup,
): RuntimeArgumentValueHit | undefined {
	if (call.displayName !== 'DateAdd' || call.slots.length < 3) {
		return undefined;
	}
	const [intervalSlot, numberSlot, dateSlot] = call.slots.map((slot) => unwrapOuterParens(slot.filter((t) => t.kind !== 'comment')));
	if (intervalSlot.length !== 1 || intervalSlot[0].kind !== 'stringLiteral' || dateSlot.length !== 1 || dateSlot[0].kind !== 'dateLiteral') {
		return undefined;
	}
	const interval = stringLiteralValue(intervalSlot[0].rawText).toLowerCase();
	const count = integerGroupValue(source, span, numberSlot, constants);
	if (count === undefined || count <= 0) {
		return undefined;
	}
	const date = parseDateLiteral(dateSlot[0].rawText);
	if (!date) {
		return undefined;
	}
	const maximum = Date.UTC(9999, 11, 31);
	const result = new Date(date);
	switch (interval) {
		case 'd': case 'y': result.setUTCDate(result.getUTCDate() + count); break;
		case 'w': result.setUTCDate(result.getUTCDate() + count); break;
		case 'ww': result.setUTCDate(result.getUTCDate() + 7 * count); break;
		case 'm': result.setUTCMonth(result.getUTCMonth() + count); break;
		case 'q': result.setUTCMonth(result.getUTCMonth() + 3 * count); break;
		case 'yyyy': result.setUTCFullYear(result.getUTCFullYear() + count); break;
		default: return undefined;
	}
	if (result.getTime() <= maximum) {
		return undefined;
	}
	return {
		displayName: 'DateAdd',
		parameterName: 'Date',
		value: `${dateSlot[0].rawText}, which the ${count} ${interval} interval(s) carry past December 31, 9999`,
		span: { start: span.start + dateSlot[0].start, end: span.start + dateSlot[0].end },
	};
}

/**
 * `DateSerial(9999, 13, 1)`: the month and day carry into the year, and a
 * date past December 31, 9999 raises error 5 (issue #189, measured in Excel
 * 16.0). The year alone decides nothing: DateSerial(10000, 0, 1) runs and is
 * December 1, 9999. A year below 100 is read as 19xx or 20xx, so it is not
 * judged.
 */
function dateSerialPastMaximum(
	source: string,
	span: Span,
	call: { displayName: string; slots: VbaToken[][] },
	constants: IntegerConstantLookup,
): RuntimeArgumentValueHit | undefined {
	if (call.displayName.replace(/^VBA\./i, '').toLowerCase() !== 'dateserial' || call.slots.length !== 3) {
		return undefined;
	}
	const [year, month, day] = call.slots.map((slot) => {
		const toks = slot.filter((t) => t.kind !== 'comment');
		return toks.length === 0 ? undefined : integerGroupValue(source, span, toks, constants);
	});
	if (year === undefined || month === undefined || day === undefined || year < 100) {
		return undefined;
	}
	// A part past the Integer range overflows first (issue #218).
	if ([year, month, day].some((part) => part < -32768 || part > 32767)) {
		return undefined;
	}
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, 1);
	date.setUTCDate(day);
	if (date.getUTCFullYear() <= 9999) {
		return undefined;
	}
	const first = call.slots[0].find((t) => t.kind !== 'comment')!;
	const last = [...call.slots[2]].reverse().find((t) => t.kind !== 'comment')!;
	return {
		displayName: call.displayName,
		parameterName: 'Year',
		value: `${year} with month ${month} and day ${day}, a date past December 31, 9999`,
		span: { start: span.start + first.start, end: span.start + last.end },
	};
}

/** A `#m/d/yyyy#` date literal as a UTC date, or undefined for any other spelling. */
function parseDateLiteral(raw: string): Date | undefined {
	const match = /^#\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(?:\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)?\s*#$/i.exec(raw);
	if (!match) {
		return undefined;
	}
	const month = Number(match[1]);
	const day = Number(match[2]);
	const year = Number(match[3]);
	if (month < 1 || month > 12 || day < 1 || day > 31) {
		return undefined;
	}
	return new Date(Date.UTC(year, month - 1, day));
}

function runtimeArgumentValueCallAt(
	toks: readonly VbaToken[],
	index: number,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
	host: string | undefined,
): {
	displayName: string;
	specs: readonly RuntimeArgumentValueSpec[];
	slots: VbaToken[][];
} | undefined {
	const name = tokenName(toks[index]);
	if (!name) {
		return undefined;
	}
	// Only a bare `Left(...)` or a genuine `VBA.Left(...)` is the intrinsic. The
	// shared helper also rejects `obj.vba.Left(...)` via the third-token ('.')
	// check that this ad-hoc logic previously omitted.
	if (!isBareOrVbaQualifiedIntrinsicCall(toks, index)) {
		return undefined;
	}
	// A bare call can be shadowed by a source symbol; a `VBA.`-qualified one
	// cannot, so only the bare form participates in the shadow gate below.
	const qualifier = index >= 2 && toks[index - 1].rawText === '.'
		? tokenName(toks[index - 2])
		: undefined;

	let parenIndex = index + 1;
	let suffix = '';
	if (isRuntimeStringFunctionSuffix(toks[parenIndex])) {
		suffix = toks[parenIndex].rawText;
		parenIndex++;
	}
	if (toks[parenIndex]?.rawText !== '(') {
		return undefined;
	}

	const specs = runtimeArgumentValueSpecs(name, host);
	const canonicalName = specs[0]?.canonicalName ?? RELATION_FUNCTIONS.get(name.toLowerCase());
	if (!canonicalName) {
		return undefined;
	}
	if (suffix && !specs[0]?.stringSuffix) {
		return undefined;
	}
	// `Error (70000)` opening a statement is the Error statement, judged by
	// runtimeStatementValueHits; only the function reads a message.
	if (canonicalName === 'Error' && index === 0) {
		return undefined;
	}
	const lower = canonicalName.toLowerCase();
	if (!qualifier && (
		moduleSignatures.has(lower) ||
		env.has(lower) ||
		runtimeCallableSourceShadowed(name, sourceNames)
	)) {
		return undefined;
	}

	const close = matchParenFrom(toks, parenIndex);
	if (close < 0) {
		return undefined;
	}
	const inner = toks.slice(parenIndex + 1, close);
	const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
	return {
		displayName: `${canonicalName}${suffix}`,
		specs,
		slots: split.slots,
	};
}

/** Functions argumentRelationHit judges, which have no single-argument bound. */
const RELATION_FUNCTIONS: ReadonlyMap<string, string> = new Map(
	['Partition', 'Pmt', 'IPmt', 'PPmt', 'SLN', 'SYD', 'DDB', 'NPer', 'Rate', 'PV', 'LBound', 'UBound', 'Join', 'Filter', 'Switch']
		.map((canonical) => [canonical.toLowerCase(), canonical]),
);

/**
 * The bounds each runtime function's arguments must keep to compile-and-run
 * clean. The first nine were VBE-oracle-backed from the start; the rest were
 * measured one call at a time in Excel 16.0 (build 20326, 2026-09-26, issue
 * #118): every listed value raises error 5 every time, and the nearest value
 * that runs - Mid("abc", 10), Round(1.5, 0), Weekday(Date, 7), Environ(1) -
 * stays quiet.
 *
 * Issue #218 added, each measured in Excel 16.0 (build 20326, 2026-09-30):
 *
 *  - Compare, for all six functions that take one, is 0, 1, or a locale ID
 *    from 3 up that Windows knows (1033 and 66567 run, 16383 and 65536
 *    raise). Only a negative value, and 2 for InStr, StrComp and Filter
 *    outside Access, are refused everywhere; no upper bound holds.
 *  - A string is at most 1073741823 characters: Left, Right and Mid's Length
 *    and String and Space's Number past it raise 5, and so does Mid's Start
 *    past 1073741824.
 *  - CVErr takes 0 to 65535; LeftB, RightB, MidB, AscB and InStrB keep Left's,
 *    Mid's, Asc's and InStr's lower bounds.
 *  - Overflow, error 6: TimeSerial and DateSerial take Integers, ChrB a Byte,
 *    String's Number and InStr's Start a Long, and Error at most 65535.
 */
function runtimeArgumentValueSpecs(name: string, host: string | undefined): readonly RuntimeArgumentValueSpec[] {
	// 2 is vbDatabaseCompare, which only Access accepts in these three.
	const databaseCompare = host === 'access' ? [] : [2];
	switch (name.toLowerCase()) {
		case 'left':
			return [{ canonicalName: 'Left', parameterName: 'Length', argumentIndex: 1, minimum: 0, maximum: MAX_STRING_LENGTH, stringSuffix: true }];
		case 'right':
			return [{ canonicalName: 'Right', parameterName: 'Length', argumentIndex: 1, minimum: 0, maximum: MAX_STRING_LENGTH, stringSuffix: true }];
		case 'string':
			return [
				{ canonicalName: 'String', parameterName: 'Number', argumentIndex: 0, minimum: 0, maximum: MAX_STRING_LENGTH, overflowType: 'Long', stringSuffix: true },
				{ canonicalName: 'String', parameterName: 'Character', argumentIndex: 1, emptyStringRaises: true, stringSuffix: true },
			];
		case 'space':
			return [{ canonicalName: 'Space', parameterName: 'Number', argumentIndex: 0, minimum: 0, maximum: MAX_STRING_LENGTH, stringSuffix: true }];
		case 'mid':
			return [
				{ canonicalName: 'Mid', parameterName: 'Start', argumentIndex: 1, minimum: 1, maximum: MAX_STRING_LENGTH + 1, stringSuffix: true },
				{ canonicalName: 'Mid', parameterName: 'Length', argumentIndex: 2, minimum: 0, maximum: MAX_STRING_LENGTH, stringSuffix: true },
			];
		case 'leftb':
			return [{ canonicalName: 'LeftB', parameterName: 'Length', argumentIndex: 1, minimum: 0, stringSuffix: true }];
		case 'rightb':
			return [{ canonicalName: 'RightB', parameterName: 'Length', argumentIndex: 1, minimum: 0, stringSuffix: true }];
		case 'midb':
			return [
				{ canonicalName: 'MidB', parameterName: 'Start', argumentIndex: 1, minimum: 1, stringSuffix: true },
				{ canonicalName: 'MidB', parameterName: 'Length', argumentIndex: 2, minimum: 0, stringSuffix: true },
			];
		case 'ascb':
			return [{ canonicalName: 'AscB', parameterName: 'String', argumentIndex: 0, emptyStringRaises: true }];
		case 'instrb':
			return [{ canonicalName: 'InStrB', parameterName: 'Start', argumentIndex: 0, minimum: 1, overflowType: 'Long', minimumSlotCount: 3, allowNamed: false }];
		case 'chrb':
			return [{ canonicalName: 'ChrB', parameterName: 'CharCode', argumentIndex: 0, overflowType: 'Byte', stringSuffix: true }];
		case 'cverr':
			return [{ canonicalName: 'CVErr', parameterName: 'ErrorNumber', argumentIndex: 0, minimum: 0, maximum: 65535 }];
		case 'error':
			return [{ canonicalName: 'Error', parameterName: 'ErrorNumber', argumentIndex: 0, maximum: 65535, boundsOverflow: true, stringSuffix: true }];
		case 'timeserial':
			return ['Hour', 'Minute', 'Second'].map((parameterName, argumentIndex) => (
				{ canonicalName: 'TimeSerial', parameterName, argumentIndex, overflowType: 'Integer' as const }
			));
		case 'filter':
			return [{ canonicalName: 'Filter', parameterName: 'Compare', argumentIndex: 3, minimum: 0, disallowed: databaseCompare }];
		case 'replace':
			return [
				{ canonicalName: 'Replace', parameterName: 'Start', argumentIndex: 3, minimum: 1 },
				{ canonicalName: 'Replace', parameterName: 'Count', argumentIndex: 4, minimum: -1 },
				// Replace takes a Compare of 2 even outside Access; -1 raises (issue #189).
				{ canonicalName: 'Replace', parameterName: 'Compare', argumentIndex: 5, minimum: 0 },
			];
		case 'instr':
			return [
				{
					canonicalName: 'InStr',
					parameterName: 'Start',
					argumentIndex: 0,
					minimum: 1,
					overflowType: 'Long',
					minimumSlotCount: 3,
					allowNamed: false,
				},
				{
					canonicalName: 'InStr',
					parameterName: 'Compare',
					argumentIndex: 3,
					minimum: 0,
					disallowed: databaseCompare,
					minimumSlotCount: 4,
					allowNamed: false,
				},
			];
		case 'instrrev':
			return [
				{ canonicalName: 'InStrRev', parameterName: 'Start', argumentIndex: 2, minimum: -1, disallowed: [0] },
				{ canonicalName: 'InStrRev', parameterName: 'Compare', argumentIndex: 3, minimum: 0 },
			];
		case 'chr':
			return [{ canonicalName: 'Chr', parameterName: 'CharCode', argumentIndex: 0, minimum: 0, maximum: 255, stringSuffix: true }];
		case 'chrw':
			return [{ canonicalName: 'ChrW', parameterName: 'CharCode', argumentIndex: 0, minimum: -32768, maximum: 65535 }];
		case 'asc':
			return [{ canonicalName: 'Asc', parameterName: 'String', argumentIndex: 0, emptyStringRaises: true }];
		case 'ascw':
			return [{ canonicalName: 'AscW', parameterName: 'String', argumentIndex: 0, emptyStringRaises: true }];
		case 'sqr':
			return [{ canonicalName: 'Sqr', parameterName: 'Number', argumentIndex: 0, minimum: 0, fractional: true }];
		case 'log':
			return [{ canonicalName: 'Log', parameterName: 'Number', argumentIndex: 0, exclusiveMinimum: 0, fractional: true }];
		case 'monthname':
			return [{ canonicalName: 'MonthName', parameterName: 'Month', argumentIndex: 0, minimum: 1, maximum: 12 }];
		case 'weekdayname':
			return [
				{ canonicalName: 'WeekdayName', parameterName: 'Weekday', argumentIndex: 0, minimum: 1, maximum: 7 },
				{ canonicalName: 'WeekdayName', parameterName: 'FirstDayOfWeek', argumentIndex: 2, minimum: 0, maximum: 7 },
			];
		case 'weekday':
			return [{ canonicalName: 'Weekday', parameterName: 'FirstDayOfWeek', argumentIndex: 1, minimum: 0, maximum: 7 }];
		case 'round':
			return [{ canonicalName: 'Round', parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: 0, maximum: 22 }];
		case 'dateserial':
			// No bound on the Year alone: dateSerialPastMaximum judges the
			// whole date the month and day carry it to (issue #189).
			return ['Year', 'Month', 'Day'].map((parameterName, argumentIndex) => (
				{ canonicalName: 'DateSerial', parameterName, argumentIndex, overflowType: 'Integer' as const }
			));
		case 'strcomp':
			return [{ canonicalName: 'StrComp', parameterName: 'Compare', argumentIndex: 2, minimum: 0, disallowed: databaseCompare }];
		case 'split':
			return [
				{ canonicalName: 'Split', parameterName: 'Limit', argumentIndex: 2, minimum: -1 },
				{ canonicalName: 'Split', parameterName: 'Compare', argumentIndex: 3, minimum: 0 },
			];
		case 'strconv':
			return [{ canonicalName: 'StrConv', parameterName: 'Conversion', argumentIndex: 1, accepts: strConvConversionAnyLocale }];
		case 'formatnumber':
			return [{ canonicalName: 'FormatNumber', parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: -1 }];
		case 'formatcurrency':
			return [{ canonicalName: 'FormatCurrency', parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: -1 }];
		case 'formatpercent':
			return [{ canonicalName: 'FormatPercent', parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: -1 }];
		case 'environ':
			return [{ canonicalName: 'Environ', parameterName: 'Expression', argumentIndex: 0, minimum: 1, stringSuffix: true }];
		case 'dateadd':
			return [{ canonicalName: 'DateAdd', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS }];
		case 'datediff':
			return [{ canonicalName: 'DateDiff', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS }];
		case 'datepart':
			return [{ canonicalName: 'DatePart', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS }];
		default:
			return [];
	}
}

/**
 * Whether some locale accepts a StrConv Conversion (issue #184, measured in
 * Excel 16.0 with the LCIDs of English, Japanese and Chinese). vbWide 4,
 * vbNarrow 8, vbKatakana 16 and vbHiragana 32 run only under an East Asian
 * locale, so they are never judged here. What no locale accepts: vbUnicode 64
 * or vbFromUnicode 128 with any other value, vbWide with vbNarrow, vbKatakana
 * with vbHiragana, and any value past 255 or below 0.
 */
function strConvConversionAnyLocale(value: number): boolean {
	if (value < 0 || value > 255) {
		return false;
	}
	if ((value & 192) !== 0 && value !== 64 && value !== 128) {
		return false;
	}
	return (value & 12) !== 12 && (value & 48) !== 48;
}

/** The longest string VBA builds: Left("abc", 1073741823) runs, 1073741824 raises 5. */
const MAX_STRING_LENGTH = 1073741823;

/** The interval strings DateAdd, DateDiff and DatePart accept. */
const DATE_INTERVALS: readonly string[] = ['yyyy', 'q', 'm', 'y', 'd', 'w', 'ww', 'h', 'n', 's'];

function runtimeArgumentValueSlot(
	slots: readonly VbaToken[][],
	spec: RuntimeArgumentValueSpec,
): VbaToken[] | undefined {
	if (spec.minimumSlotCount !== undefined && slots.length < spec.minimumSlotCount) {
		return undefined;
	}
	let positionalIndex = 0;
	for (const slot of slots) {
		const named = namedArgumentSlot(slot);
		if (named) {
			if (spec.allowNamed === false) {
				continue;
			}
			if (named.name.toLowerCase() === spec.parameterName.toLowerCase()) {
				return named.value;
			}
			continue;
		}
		if (positionalIndex === spec.argumentIndex) {
			return slot;
		}
		positionalIndex++;
	}
	return undefined;
}

function integerArgumentOutsideBounds(
	source: string,
	slot: readonly VbaToken[],
	sliceStart: number,
	spec: RuntimeArgumentValueSpec,
	constants: IntegerConstantLookup,
	stringCalls: KnownStringCallContext,
): { value: number | string; span: Span; error?: 6 } | undefined {
	const { knownStrings } = stringCalls;
	const toks = unwrapOuterParens(
		slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline'),
	);
	if (toks.length === 0) {
		return undefined;
	}
	// String-valued bounds: an empty literal where the function needs a
	// character (Asc(""), String(3, "")), or a literal outside the words the
	// function accepts (DateAdd("x", ...)). A String local the procedure never
	// assigns is "" too (Asc(s)).
	if (spec.emptyStringRaises || spec.allowedStrings) {
		if (toks.length !== 1) {
			return undefined;
		}
		const knownName = tokenName(toks[0])?.toLowerCase();
		const known = knownName !== undefined ? knownStrings.get(knownName) : undefined;
		if (toks[0].kind !== 'stringLiteral' && known === undefined) {
			return undefined;
		}
		const text = toks[0].kind === 'stringLiteral' ? stringLiteralValue(toks[0].rawText) : known!;
		const raises = spec.emptyStringRaises
			? text.length === 0
			: !spec.allowedStrings!.some((allowed) => allowed.toLowerCase() === text.toLowerCase());
		if (!raises) {
			return undefined;
		}
		const value = toks[0].kind === 'stringLiteral'
			? toks[0].rawText
			: `"${text}" (${toks[0].rawText} is never given another value)`;
		return { value, span: { start: sliceStart + toks[0].start, end: sliceStart + toks[0].end } };
	}
	let sign = 1;
	let literal = toks[0];
	let start = literal?.start;
	let literalValue: number | undefined;
	const signedLiteral = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+');
	if (signedLiteral) {
		sign = toks[0].rawText === '-' ? -1 : 1;
		literal = toks[1];
		start = toks[0].start;
	}
	if (literal?.kind === 'integerLiteral' && start !== undefined && (toks.length === 1 || signedLiteral)) {
		const rawValue = parseVbaIntegerLiteral(literal.rawText);
		if (rawValue !== undefined) {
			literalValue = sign * rawValue;
		}
	} else if (literal?.kind === 'floatLiteral' && start !== undefined && (toks.length === 1 || signedLiteral)) {
		// Sqr(-4.5) and Log(0.0) raise like their whole-number neighbours.
		const rawValue = Number(literal.rawText.replace(/[!#@]$/, ''));
		if (Number.isFinite(rawValue)) {
			literalValue = sign * rawValue;
		}
	}
	if (literalValue !== undefined) {
		const verdict = argumentValueVerdict(literalValue, spec);
		if (verdict === 'runs') {
			return undefined;
		}
		return {
			value: shownArgumentValue(literalValue, spec),
			span: { start: sliceStart + start!, end: sliceStart + literal.end },
			...(verdict === 6 ? { error: 6 as const } : {}),
		};
	}

	// `InStr(s, " ") - 1` with s known is `0 - 1` (issue #201).
	const expressionValue = evaluateIntegerConstantExpression(
		foldKnownStringCalls(toks, stringCalls)
			?? source.slice(sliceStart + toks[0].start, sliceStart + toks[toks.length - 1].end),
		constants,
	);
	const verdict = expressionValue === undefined ? 'runs' : argumentValueVerdict(expressionValue, spec);
	if (verdict === 'runs') {
		return undefined;
	}
	return {
		value: shownArgumentValue(expressionValue!, spec),
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
		...(verdict === 6 ? { error: 6 as const } : {}),
	};
}

/**
 * What passing `rawValue` does: runs, raises 5, or overflows (6). A
 * whole-number parameter whose type no spec states is judged only inside the
 * Long range, since past it the conversion overflows first and the typed
 * signature's argument-type-mismatch says so.
 */
function argumentValueVerdict(rawValue: number, spec: RuntimeArgumentValueSpec): 'runs' | 5 | 6 {
	if (!spec.fractional) {
		const value = passedArgumentValue(rawValue, spec);
		const range = OVERFLOW_RANGES[spec.overflowType ?? 'Long'];
		if (value < range.min || value > range.max) {
			return spec.overflowType ? 6 : 'runs';
		}
	}
	if (integerArgumentValueInBounds(rawValue, spec)) {
		return 'runs';
	}
	return spec.boundsOverflow ? 6 : 5;
}

/**
 * The value VBA passes: a whole-number parameter takes the argument rounded
 * half to even, so Space(-0.5) is Space(0) and runs, and Mid(s, 0.5) is
 * Mid(s, 0) and raises (issue #189, measured in Excel 16.0).
 */
function passedArgumentValue(value: number, spec: RuntimeArgumentValueSpec): number {
	return spec.fractional ? value : bankersRound(value);
}

/** The argument as the message states it, with the rounded value VBA uses. */
function shownArgumentValue(value: number, spec: RuntimeArgumentValueSpec): number | string {
	const passed = passedArgumentValue(value, spec);
	return passed === value ? value : `${value}, which VBA rounds to ${passed}`;
}

function integerArgumentValueInBounds(
	rawValue: number,
	spec: RuntimeArgumentValueSpec,
): boolean {
	const value = passedArgumentValue(rawValue, spec);
	if (spec.minimum !== undefined && value < spec.minimum) {
		return false;
	}
	if (spec.exclusiveMinimum !== undefined && value <= spec.exclusiveMinimum) {
		return false;
	}
	if (spec.maximum !== undefined && value > spec.maximum) {
		return false;
	}
	if (spec.disallowed?.includes(value)) {
		return false;
	}
	if (spec.accepts && Number.isInteger(value) && !spec.accepts(value)) {
		return false;
	}
	return true;
}

function isRuntimeStringFunctionSuffix(tok: VbaToken | undefined): boolean {
	return tok?.rawText === '$';
}

function isDeclarationLikeStatement(toks: readonly VbaToken[]): boolean {
	const first = tokenText(toks[0]);
	switch (first) {
		case 'dim':
		case 'static':
		case 'const':
		case 'private':
		case 'public':
		case 'friend':
		case 'declare':
		case 'sub':
		case 'function':
		case 'property':
		case 'type':
		case 'enum':
			return true;
		default:
			return false;
	}
}

interface RuntimeConversionValueHit {
	displayName: string;
	name: string;
	/** What the literal cannot become: 'a number', 'Boolean', 'Date'. */
	target: string;
	span: Span;
}

/**
 * Rule: selected conversion functions compile with Variant-like arguments but
 * can deterministically fail at runtime for literal values that cannot be
 * converted. This first slice is intentionally narrow for CDate string
 * literals that are plainly non-date text.
 */
export function checkRuntimeConversionValues(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
	activity?: ConditionalActivityTracker,
): ProcedureStatementVisitor {
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		// `s = "abc"` then `CLng(s)`: the string a local holds here (issue #238).
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		return (stmt) => {
			const strings = new Map<string, string>();
			for (const [lower, value] of valuesAt(stmt)) {
				if (value.kind === 'string' && !value.contentMutated) {
					strings.set(lower, value.value as string);
				}
			}
			for (const hit of runtimeConversionValueHits(source, stmt.span, sourceNames, strings)) {
				push(
					'runtimeConversionValue',
					`${hit.displayName} cannot convert ${hit.name} to ${hit.target}. This will raise Run-time error '13': Type mismatch.`,
					hit.span,
				);
			}
		};
	};
}

/**
 * The conversion functions and the kind each converts a string literal to
 * (issues #118 and #188, each measured in Excel 16.0). What converts is
 * stringConversion.ts's: `CLng("4x2")`, `CDbl("5%")`, `CBool(" True ")` and
 * `CDate("March")` raise 13, and `CLng("&HFF")` and `CDbl("(5)")` run.
 */
const CONVERSION_TARGETS: Readonly<Record<string, 'numeric' | 'boolean' | 'date'>> = {
	cbyte: 'numeric', cint: 'numeric', clng: 'numeric', clnglng: 'numeric', clngptr: 'numeric',
	csng: 'numeric', cdbl: 'numeric', ccur: 'numeric', cdec: 'numeric', sgn: 'numeric',
	cbool: 'boolean',
	cdate: 'date', cvdate: 'date', datevalue: 'date', timevalue: 'date',
	// The math functions read a number the same way (issue #242): Abs("abc")
	// and Hex("abc") raise 13, and Oct("8") runs.
	abs: 'numeric', sqr: 'numeric', int: 'numeric', fix: 'numeric', round: 'numeric', hex: 'numeric', oct: 'numeric',
	exp: 'numeric', log: 'numeric', sin: 'numeric', cos: 'numeric', tan: 'numeric', atn: 'numeric',
	year: 'date', month: 'date', day: 'date', weekday: 'date', hour: 'date', minute: 'date', second: 'date',
	dateadd: 'date', datepart: 'date', datediff: 'date',
};

/**
 * The arguments each function converts, where it is not the first (issue
 * #218): DateAdd's Date is its third, DatePart's its second, and DateDiff
 * converts its second and third.
 */
const CONVERTED_SLOTS: Readonly<Record<string, readonly number[]>> = {
	dateadd: [2], datepart: [1], datediff: [1, 2],
};

/**
 * The functions that read a number as a Date serial, which runs from
 * -657434 (January 1, 100) to 2958465 (December 31, 9999): Year(2958466) and
 * Day(-657435) raise 13, Year(2958465.9) and Year(-657434.9) run (issue #218,
 * measured in Excel 16.0). CDate raises 6 there, which arithmetic-overflow
 * reports.
 */
const DATE_SERIAL_READERS: ReadonlySet<string> = new Set([
	'year', 'month', 'day', 'weekday', 'hour', 'minute', 'second', 'dateadd', 'datepart', 'datediff',
]);

function runtimeConversionValueHits(
	source: string,
	span: Span,
	sourceNames: SourceNameScope,
	knownStrings: ReadonlyMap<string, string> = new Map(),
): RuntimeConversionValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeConversionValueHit[] = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const name = tokenName(toks[i]);
		const target = name ? CONVERSION_TARGETS[name.toLowerCase()] : undefined;
		if (!name || !target) {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const qualified = toks[i - 1]?.rawText === '.';
		if (!qualified && runtimeCallableSourceShadowed(name, sourceNames)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const split = splitArgSlots(toks.slice(i + 2, close), span.start);
		if (split.slots.some((slot) => namedArgumentSlot(slot))) {
			continue;
		}
		for (const index of CONVERTED_SLOTS[name.toLowerCase()] ?? [0]) {
			const slot = (split.slots[index] ?? []).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
			const at = split.spans[index] ?? (slot.length > 0 ? { start: span.start + slot[0].start, end: span.start + slot[slot.length - 1].end } : undefined);
			const displayName = qualified ? `VBA.${name}` : name;
			const held = slot.length === 1 && slot[0].kind === 'identifier' ? knownStrings.get(slot[0].rawText.toLowerCase()) : undefined;
			if (slot.length === 1 && (slot[0].kind === 'stringLiteral' || held !== undefined) && at) {
				const value = held ?? stringLiteralValue(slot[0].rawText);
				const invalid = target === 'date'
					? isInvalidDateString(value)
					: target === 'boolean'
						? isInvalidBooleanString(value)
						: isInvalidNumericString(value);
				if (invalid) {
					hits.push({
						displayName,
						name: held === undefined ? slot[0].rawText : `${slot[0].rawText}, which holds ${JSON.stringify(held)} here,`,
						target: target === 'date' ? 'Date' : target === 'boolean' ? 'Boolean' : 'a number',
						span: at,
					});
				}
				continue;
			}
			const serial = DATE_SERIAL_READERS.has(name.toLowerCase()) ? signedNumericLiteral(slot) : undefined;
			if (serial !== undefined && at && (serial >= 2958466 || serial <= -657435)) {
				hits.push({
					displayName,
					name: String(serial),
					target: 'a Date, whose serial numbers run from -657434 (January 1, 100) to 2958465 (December 31, 9999)',
					span: at,
				});
			}
		}
	}
	return hits;
}

/** The value of a lone numeric literal, optionally signed. */
function signedNumericLiteral(slot: readonly VbaToken[]): number | undefined {
	const toks = unwrapOuterParens(slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline'));
	const signed = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+');
	if (toks.length !== 1 && !signed) {
		return undefined;
	}
	const literal = toks[toks.length - 1];
	if (literal.kind !== 'integerLiteral' && literal.kind !== 'floatLiteral') {
		return undefined;
	}
	return numericLiteralGroupValue(toks);
}
