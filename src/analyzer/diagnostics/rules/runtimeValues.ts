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
import { isLeafStatement } from '../../parser/nodes';
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
import { dateLiteralSerial } from '../../constants/dateLiteral';
import { knownArrayShapesAt, moduleOptionBase, redimShapesAt } from './arrays';
import { checkEachCounterPass, loopCountersAt } from '../loopCounters';
import { functionIntegerResult, knownFunctionResults } from '../functionResults';
import { straightLineAssignments } from '../straightLineValues';
import { foldKnownStringCalls, knownDate, moduleCompare, parseDateLiteral, type KnownStringCallContext } from '../knownStringCalls';
import { bankersRound, isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
import { isInvalidBooleanString, isInvalidDateString, isInvalidNumericString, isInvalidTimeString } from '../stringConversion';
import { fixedStringLength, moduleTypes } from '../typeFields';
import {
	callableTypeSignaturesFor,
	inferExpressionType,
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
	pickedValues,
} from '../typeInference';
import {
	matchParenFrom,
	rawExpressionTokens,
	statementTokens,
	statementTokensAfterLeadingLabel,
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
	/**
	 * InStr returns before it reads Start or Compare when either string is
	 * empty: `InStr(0, "abc", "")` is 0 (issue #481, measured in Excel 16.0).
	 * A Start past the Long range still overflows.
	 */
	skippedByEmptyString?: boolean;
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
	const types = moduleTypes(source, mod, activity);
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
		let shapesAt: ReturnType<typeof redimShapesAt> | undefined;
		let knownShapesAt: ReturnType<typeof knownArrayShapesAt> | undefined;
		let currentStmt: BodyNode | undefined;
		const declarationOf = (lower: string): LocalDeclaration | undefined => {
			const local = locals.find((child) => child.name.toLowerCase() === lower);
			if (local?.kind !== 'localVariable') {
				return undefined;
			}
			const type = normalizeType(local.asType);
			if (local.isArray && local.arrayBounds === undefined) {
				// A dynamic array may still be unallocated, which Join takes; one
				// a straight line has ReDim'd is judged as a fixed one (issue #342).
				const stmt = currentStmt;
				const shape = stmt && isLeafStatement(stmt)
					? (shapesAt ??= redimShapesAt(source, symbols, member, activity, moduleOptionBase(mod, activity))).get(stmt)?.get(lower)
					: undefined;
				return shape ? { asType: local.asType ?? 'Variant', dimensions: shape.dims.length } : undefined;
			}
			if (local.isArray) {
				return { asType: local.asType ?? 'Variant', dimensions: splitTopLevelTokenGroups(rawExpressionTokens(local.arrayBounds!), 0, ',').length };
			}
			if (type === undefined || type === 'variant') {
				// A Variant holding a block's `.Value` has two dimensions (issue #492).
				const shape = currentStmt && isLeafStatement(currentStmt) ? (knownShapesAt ??= knownArrayShapesAt(source, symbols, member, activity, moduleOptionBase(mod, activity)))(currentStmt).get(lower) : undefined;
				return shape && shape.dims.length > 1 ? { asType: 'Variant', dimensions: shape.dims.length } : undefined;
			}
			return isKnownScalarType(type) ? { asType: local.asType!, dimensions: 0 } : undefined;
		};
		const stringsFor = new Map<ReadonlyMap<string, KnownLocalValue>, { strings: ReadonlyMap<string, string>; lengths: ReadonlyMap<string, number> }>();
		const stringsAt = (values: ReadonlyMap<string, KnownLocalValue>): { strings: ReadonlyMap<string, string>; lengths: ReadonlyMap<string, number> } => {
			let out = stringsFor.get(values);
			if (!out) {
				out = {
					strings: pickedValues(values, (_lower, value) => (value.kind === 'string' && !value.contentMutated ? value.value as string : undefined)),
					lengths: pickedValues(values, (_lower, value) => (value.kind === 'string' ? (value.value as string).length : undefined)),
				};
				stringsFor.set(values, out);
			}
			return out;
		};
		let results: ReturnType<typeof knownFunctionResults> | undefined;
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
				if (local) {
					return local.kind === 'number' && Number.isInteger(local.value) ? (local.value as number) : undefined;
				}
				// A Date local a straight line has just set to a Date literal
				// passes its serial: Chr(d) with d = #1/2/2000# (issue #332).
				if (currentStmt && normalizeType(env.get(name.toLowerCase())) === 'date') {
					const held = (reaching ??= straightLineAssignments(source, member.body, activity)).get(currentStmt)?.get(name.toLowerCase())?.filter((tok) => tok.kind !== 'comment');
					if (held?.length === 1 && held[0].kind === 'dateLiteral') {
						return dateLiteralSerial(held[0].rawText);
					}
				}
				// `Mid(s, F())` with F a Function of the module returning 0 (issue #448).
				return functionIntegerResult(name, results ??= knownFunctionResults(source, mod, activity), member, symbols);
			},
		};
		// The locals a straight line has just set to Null (issue #364).
		let reaching: ReturnType<typeof straightLineAssignments> | undefined;
		return (stmt) => {
			known = valuesAt(stmt);
			currentStmt = stmt;
			const { strings: knownStrings, lengths: knownStringLengths } = stringsAt(known);
			// The argument's type; for a Variant local, the type of what a
			// straight line has just put in it.
			const valueType = (slot: readonly VbaToken[] | undefined): string | undefined => {
				const held = (lower: string): VbaToken[] | undefined =>
					(reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment');
				return staticValueType((slot ?? []).filter((tok) => tok.kind !== 'comment'), env, moduleSignatures, sourceNames, source, held);
			};
			const isNullSlot = (slot: readonly VbaToken[]): boolean => {
				const value = slot.filter((tok) => tok.kind !== 'comment');
				if (value.length !== 1) {
					return false;
				}
				if (tokenText(value[0]) === 'null') {
					return true;
				}
				const lower = tokenName(value[0])?.toLowerCase();
				const held = lower ? (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment') : undefined;
				return held?.length === 1 && tokenText(held[0]) === 'null';
			};
			// `Dim x As Double: x = -1.5: Chr(x)`: a local known to hold a number
			// that is not whole, or Empty, stands in as that literal (issue #332,
			// measured in Excel 16.0). A whole number already comes through the
			// lookup.
			const standIn = (slot: readonly VbaToken[]): readonly VbaToken[] => {
				const value = slot.filter((tok) => tok.kind !== 'comment');
				const lower = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
				// `a = Empty` reads as the keyword, which the literal checks know.
				const reached = lower ? (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment') : undefined;
				if (reached?.length === 1 && tokenText(reached[0]) === 'empty') {
					return [{ ...value[0], kind: 'keyword', rawText: 'Empty' }];
				}
				const held = lower ? known.get(lower) : undefined;
				if (!held || held.contentMutated || held.kind !== 'number' || Number.isInteger(held.value)) {
					return slot;
				}
				return literalTokensFor(held.value as number, value[0]);
			};
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
					// A Date or Variant local a straight line has just given a date literal (issue #559).
					dateOf: (lower) => {
						const held = (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment');
						const declared = normalizeType(env.get(lower));
						return held?.length === 1 && held[0].kind === 'dateLiteral' && (declared === 'date' || declared === 'variant') ? parseDateLiteral(held[0].rawText) : undefined;
					},
				};
				for (const hit of runtimeArgumentValueHits(source, stmt.span, moduleSignatures, env, lookup, stringCalls, sourceNames, host, declarationOf, isNullSlot, valueType, standIn)) {
					const raises = hit.error === 6 ? `'6': Overflow` : `'5': Invalid procedure call or argument`;
					report(
						'runtimeArgumentValue',
						hit.message ?? `Argument '${hit.parameterName}' of '${hit.displayName}' is ${hit.value}; this will raise Run-time error ${raises}.`,
						hit.span,
					);
				}
				// A fixed-length string is always its declared length (issue #248).
				const fixedLengthOf = (slot: readonly VbaToken[]): number | undefined => fixedStringLength(slot, symbols, member, types, lookup);
				for (const hit of runtimeStatementValueHits(source, stmt.span, lookup, knownStringLengths, knownStrings, sourceNames, fixedLengthOf)) {
					report('runtimeArgumentValue', hit.message, hit.span);
				}
				for (const hit of dateDiffPastDateRange(source, stmt.span, stringCalls, lookup)) {
					report('runtimeConversionValue', hit.message, hit.span);
				}
			}, push);
			counterValues = NO_COUNTER_VALUES;
		};
	};
}

/** A count DateAdd reads: it drops the fraction, so DateAdd("d", -0.6, #1/1/100#) runs. */
function truncated(value: number | undefined): number | undefined {
	return value === undefined ? undefined : Math.trunc(value);
}

/** Literal tokens for a number, at the place of the token they stand in for: `-1.5` is `-` and `1.5`. */
function literalTokensFor(value: number, at: VbaToken): VbaToken[] {
	const magnitude = Math.abs(value);
	const literal: VbaToken = {
		...at,
		kind: Number.isInteger(magnitude) ? 'integerLiteral' : 'floatLiteral',
		rawText: String(magnitude).toUpperCase(),
	};
	return value < 0 ? [{ ...at, kind: 'operator', rawText: '-', end: at.start }, literal] : [literal];
}

/**
 * `CDate(DateDiff("s", #2/29/2000#, t))`: DateDiff gives a Variant, and CDate
 * of a Variant number past the date range raises 13, where a Long past it
 * raises 6 (issue #559, measured in Excel 16.0). The dates must be known.
 */
function dateDiffPastDateRange(
	source: string,
	span: Span,
	stringCalls: KnownStringCallContext,
	constants: IntegerConstantLookup,
): Array<{ message: string; span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span).filter((tok) => tok.kind !== 'comment');
	const out: Array<{ message: string; span: Span }> = [];
	for (let i = 0; i + 1 < toks.length; i++) {
		if (tokenText(toks[i]) !== 'cdate' || toks[i + 1].rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i) || stringCalls.shadowed('cdate')) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		const arg = close > i + 2 ? toks.slice(i + 2, close) : [];
		if (!arg.some((tok, k) => tokenText(tok) === 'datediff' && arg[k + 1]?.rawText === '(')) {
			continue;
		}
		const folded = foldKnownStringCalls(arg, stringCalls);
		const value = folded === undefined ? undefined : evaluateIntegerConstantExpression(folded, constants);
		if (value === undefined || (value >= -657434 && value <= 2958465)) {
			continue;
		}
		out.push({
			message: `CDate cannot convert ${value}, the Variant DateDiff gives here, to a date: it is past the range of dates. This will raise Run-time error '13': Type mismatch.`,
			span: { start: span.start + arg[0].start, end: span.start + arg[arg.length - 1].end },
		});
	}
	return out;
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
	fixedLengthOf?: (slot: readonly VbaToken[]) => number | undefined,
): Array<{ message: string; span: Span }> {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const out: Array<{ message: string; span: Span }> = [];
	const at = (tok: VbaToken): Span => ({ start: span.start + tok.start, end: span.start + tok.end });
	const first = toks[0];
	// `On n GoTo L1, L2` and `On n GoSub`: an index outside 0 to 255 raises 5,
	// and text that is no number raises 13 (issue #499, measured in Excel
	// 16.0). 0, or one past the last label, falls through to the next line.
	const jump = toks.findIndex((tok, i) => i > 1 && (tokenText(tok) === 'goto' || tokenText(tok) === 'gosub'));
	if (tokenText(first) === 'on' && tokenText(toks[1]) !== 'error' && tokenText(toks[1]) !== 'local' && jump > 1) {
		const index = toks.slice(1, jump).filter((tok) => tok.kind !== 'comment');
		const value = integerGroupValue(source, span, index, constants);
		const name = index.length === 1 ? tokenName(index[0])?.toLowerCase() : undefined;
		const text = index.length === 1 && index[0].kind === 'stringLiteral' ? stringLiteralValue(index[0].rawText) : name !== undefined ? knownStrings.get(name) : undefined;
		const indexSpan = { start: span.start + index[0].start, end: span.start + index[index.length - 1].end };
		const shown = toks.slice(1, jump).map((tok) => tok.rawText).join(' ');
		if (value !== undefined && (value < 0 || value > 255)) {
			out.push({ message: `On ${tokenText(toks[jump]) === 'goto' ? 'GoTo' : 'GoSub'} takes an index from 0 to 255, and ${shown} is ${value}. This will raise Run-time error '5': Invalid procedure call or argument.`, span: indexSpan });
		} else if (value === undefined && text !== undefined && !/\d/.test(text)) {
			out.push({ message: `On ${tokenText(toks[jump]) === 'goto' ? 'GoTo' : 'GoSub'} takes a number for its index, and ${shown} holds "${text}". This will raise Run-time error '13': Type mismatch.`, span: indexSpan });
		}
	}
	// `Mid(s, 5, 1) = "x"` with s holding "abc": the statement form starts
	// past the end of the string, error 5 (issue #118). Only the length
	// matters, which an earlier Mid statement cannot have changed. A
	// fixed-length string's length is its declaration's, assigned or not
	// (issue #248). `Mid$` lexes as Mid and a `$` of its own.
	const midOpen = toks[1]?.rawText === '$' ? 2 : 1;
	// MidB counts bytes, two to a character: `MidB(s, 9, 1) = "x"` on "abc"
	// starts past its six (issue #327, measured in Excel 16.0).
	const bytes = tokenText(first) === 'midb';
	if ((tokenText(first) === 'mid' || bytes) && toks[midOpen]?.rawText === '(') {
		const close = matchParenFrom(toks, midOpen);
		if (close > 0 && toks[close + 1]?.rawText === '=') {
			const split = splitArgSlots(toks.slice(midOpen + 1, close), span.start);
			const target = split.slots[0]?.length === 1 ? tokenName(split.slots[0][0])?.toLowerCase() : undefined;
			const fixed = split.slots[0]?.length ? fixedLengthOf?.(split.slots[0]) : undefined;
			const characters = fixed ?? (target !== undefined ? knownStringLengths.get(target) : undefined);
			const length = characters !== undefined && bytes ? characters * 2 : characters;
			const startSlot = split.slots[1];
			const start = startSlot ? integerGroupValue(source, span, startSlot, constants) : undefined;
			if (length !== undefined && start !== undefined && start > length) {
				const form = bytes ? 'MidB' : 'Mid';
				const unit = bytes ? 'byte(s)' : 'character(s)';
				out.push({
					message: fixed === undefined
						? `${form} statement start ${start} is past the end of ${split.slots[0][0].rawText}, which is ${length} ${unit} long. This will raise Run-time error '5': Invalid procedure call or argument.`
						: `${form} statement start ${start} is past the end of ${split.slots[0].map((tok) => tok.rawText).join('')}, a fixed-length string of ${length} ${unit}. This will raise Run-time error '5': Invalid procedure call or argument.`,
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
			// A local known to hold a number, and True or False, count too:
			// `z ^ -1` with z never assigned, `0 ^ True` (issue #331, measured
			// in Excel 16.0).
			const named = (operand: VbaToken | undefined, beside: VbaToken | undefined): number | undefined => {
				const word = tokenText(operand);
				if (word === 'true' || word === 'false') {
					return word === 'true' ? -1 : 0;
				}
				const name = operand && beside?.rawText !== '(' && beside?.rawText !== '.' ? tokenName(operand) : undefined;
				return name ? constants.get(name.toLowerCase()) : undefined;
			};
			const before = toks[i - 2]?.rawText === '.' ? undefined : named(toks[i - 1], undefined);
			const base = numericOperandBefore(toks, i) ?? before;
			const exponent = numericOperandAfter(toks, i) ?? named(toks[i + 1], toks[i + 2]);
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
 * with a character left reaches a bad list anywhere after it: "b" Like
 * "*[" raises (issue #336). A comparison Option Compare Text could
 * decide otherwise proves nothing.
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
			// A `*` with a character left reaches a bad list anywhere after
			// it: "b" Like "*[" and "b" Like "*[a][" raise (issue #336).
			return p < subject.length && invalidLikePattern(pattern.slice(i + 1)) !== undefined;
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

/** The built-ins that return Null for a Null argument before checking the rest (issue #364). */
const NULL_RETURNING: ReadonlySet<string> = new Set(['mid', 'left', 'right', 'instr', 'strcomp']);

/**
 * The first-argument types whose Round raises 5 past 22 digits (issue #402,
 * measured in Excel 16.0): Round(3#, 23), Round("3", 23) and
 * Round(#1/2/2000#, 23) raise; an Integer, Long, Byte, Boolean, Currency or
 * Decimal runs at any count, Round(3, 256) and Round(CCur(3.5), 256).
 */
const ROUND_DIGIT_LIMITED: ReadonlySet<string> = new Set(['double', 'single', 'string', 'date']);

/** The type a number literal's suffix gives it; none leaves an integer literal whole and a float a Double. */
const LITERAL_SUFFIX_TYPES: Readonly<Record<string, string>> = { '#': 'double', '!': 'single', '@': 'currency', '%': 'integer', '&': 'long', '^': 'longlong' };

/**
 * The VBA type of a value, lowercase, or undefined where the forms below do
 * not settle it: a literal (by its suffix), a declared local (a Variant by
 * what a straight line just put in it), a call (by its return type), and a
 * `/` of such operands, which gives a Double. A Currency or Decimal operand
 * of `/` is left unsettled.
 */
function staticValueType(
	toks: VbaToken[],
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope,
	source: string,
	held: (lower: string) => VbaToken[] | undefined,
	depth = 0,
): string | undefined {
	const value = unwrapOuterParens(toks);
	const signed = value.length === 2 && (value[0].rawText === '-' || value[0].rawText === '+') ? value[1] : undefined;
	const single = value.length === 1 ? value[0] : signed;
	if (single) {
		if (single.kind === 'stringLiteral') {
			return 'string';
		}
		if (single.kind === 'dateLiteral') {
			return 'date';
		}
		if (single.kind === 'integerLiteral' || single.kind === 'floatLiteral') {
			return LITERAL_SUFFIX_TYPES[single.rawText.slice(-1)] ?? (single.kind === 'integerLiteral' ? 'integer' : 'double');
		}
		const text = tokenText(single);
		if (text === 'true' || text === 'false') {
			return 'boolean';
		}
		const lower = signed ? undefined : tokenName(single)?.toLowerCase();
		const declared = lower ? normalizeType(env.get(lower)) : undefined;
		if (declared === 'variant' && depth === 0) {
			const assigned = held(lower!);
			return assigned?.length ? staticValueType(assigned, env, moduleSignatures, sourceNames, source, held, 1) : undefined;
		}
		return declared;
	}
	if (value.length >= 3 && tokenName(value[0]) && value[1].rawText === '(' && matchParenFrom(value, 1) === value.length - 1) {
		return normalizeType(inferExpressionType(value, 0, env, moduleSignatures, sourceNames, source)?.type);
	}
	// Operands split at each top-level `/`; any other top-level operator leaves it unsettled.
	const operands: VbaToken[][] = [[]];
	let parens = 0;
	for (const tok of value) {
		parens += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
		if (parens === 0 && tok.rawText === '/') {
			operands.push([]);
		} else {
			operands[operands.length - 1].push(tok);
		}
	}
	if (operands.length < 2) {
		return undefined;
	}
	const settled = operands.every((operand) => {
		const type = operand.length ? staticValueType(operand, env, moduleSignatures, sourceNames, source, held, depth) : undefined;
		return type !== undefined && !['currency', 'decimal', 'variant', 'string', 'date'].includes(type);
	});
	return settled ? 'double' : undefined;
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
	isNullSlot: (slot: readonly VbaToken[]) => boolean = () => false,
	valueType: (slot: readonly VbaToken[] | undefined) => string | undefined = () => undefined,
	standIn: (slot: readonly VbaToken[]) => readonly VbaToken[] = (slot) => slot,
): RuntimeArgumentValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeArgumentValueHit[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const found = runtimeArgumentValueCallAt(toks, i, span, moduleSignatures, env, sourceNames, host);
		if (!found) {
			continue;
		}
		// A local known to hold a number is read as that number (issue #332).
		const call = { ...found, slots: found.slots.map((slot) => [...standIn(slot)]) };
		// `Mid(Null, 0)`, `InStr(0, Null, "a")`: a Null argument makes the call
		// return Null before the others are checked (issue #364, measured in
		// Excel 16.0).
		// String checks its Character for Null first: String(-1, Null) is Null
		// (issue #409, measured in Excel 16.0).
		// A Null Start is no position: InStr(v, "abc", "a") with v Null raises
		// 94, where a Null string hands Null back (issue #332, measured).
		if (call.specs[0]?.canonicalName === 'InStr' && call.slots.length >= 3 && isNullSlot(call.slots[0])) {
			const start = call.slots[0].filter((tok) => tok.kind !== 'comment');
			hits.push({
				displayName: call.displayName,
				parameterName: 'Start',
				value: 'Null',
				message: `InStr's Start is Null here, which is no position. This will raise Run-time error '94': Invalid use of Null.`,
				span: { start: span.start + start[0].start, end: span.start + start[start.length - 1].end },
			});
			continue;
		}
		const nullCharacter = call.specs[0]?.canonicalName === 'String' && call.slots[1] !== undefined && isNullSlot(call.slots[1]);
		if (nullCharacter || (NULL_RETURNING.has(call.specs[0]?.canonicalName.toLowerCase() ?? '') && call.slots.some((slot) => isNullSlot(slot)))) {
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
			if (spec.skippedByEmptyString && literal.error !== 6 && call.slots.slice(1, 3).some((text) => knownEmptyString(text, stringCalls.knownStrings))) {
				continue;
			}
			// Round's digit limit follows the value's type, not its fraction:
			// Round(3#, 23) raises 5 and Round(3, 23) runs (issue #402). A count
			// below 0 always raises.
			if (spec.canonicalName === 'Round' && typeof literal.value === 'number' && literal.value > 0 && !ROUND_DIGIT_LIMITED.has(valueType(call.slots[0]) ?? '')) {
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
		const overflow = dateAddPastMaximum(source, span, call, constants, stringCalls)
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
 * `DateAdd("d", 1, #12/31/9999#)` and `DateAdd("yyyy", -1, #1/1/100#)`:
 * adding to a date past the last date VBA has (December 31, 9999) or before
 * the first (January 1, 100) raises error 5 (issues #118 and #262, measured
 * in Excel 16.0 for every interval). The date is a literal or a DateSerial
 * of literals from year 100 on; the count is a whole number. Months,
 * quarters and years keep the day where the month has it, so only the month
 * they reach decides.
 */
function dateAddPastMaximum(
	source: string,
	span: Span,
	call: { displayName: string; slots: VbaToken[][] },
	constants: IntegerConstantLookup,
	stringCalls?: KnownStringCallContext,
): RuntimeArgumentValueHit | undefined {
	if (call.displayName !== 'DateAdd' || call.slots.length < 3) {
		return undefined;
	}
	const [intervalSlot, numberSlot, dateSlot] = call.slots.map((slot) => unwrapOuterParens(slot.filter((t) => t.kind !== 'comment')));
	if (intervalSlot.length !== 1 || intervalSlot[0].kind !== 'stringLiteral' || dateSlot.length === 0) {
		return undefined;
	}
	const interval = stringLiteralValue(intervalSlot[0].rawText).toLowerCase();
	// A count or a date another call gives: `Year(#6/15/5000#)`,
	// `DateValue("12/31/9999")` (issue #510, measured in Excel 16.0).
	const folded = stringCalls && numberSlot.length > 0 ? foldKnownStringCalls(numberSlot, stringCalls) : undefined;
	const count = folded !== undefined
		? evaluateIntegerConstantExpression(folded, constants)
		: integerGroupValue(source, span, numberSlot, constants) ?? truncated(numericLiteralGroupValue(numberSlot));
	if (count === undefined || count === 0) {
		return undefined;
	}
	// The count is held as a Long: 2147483648 of any interval overflows
	// (issue #332, measured in Excel 16.0).
	if (Math.abs(count) > 2147483647) {
		const first = numberSlot[0];
		const last = numberSlot[numberSlot.length - 1];
		return {
			displayName: 'DateAdd',
			parameterName: 'Number',
			value: String(count),
			message: `DateAdd counts its intervals in a Long, and ${source.slice(span.start + first.start, span.start + last.end)} is outside it. This will raise Run-time error '6': Overflow.`,
			span: { start: span.start + first.start, end: span.start + last.end },
			error: 6,
		};
	}
	const date = dateSlot.length === 1 && dateSlot[0].kind === 'dateLiteral'
		? parseDateLiteral(dateSlot[0].rawText)
		: dateSerialOfLiterals(source, span, dateSlot, constants) ?? knownDate(dateSlot, (group) => integerGroupValue(source, span, group, constants), stringCalls?.dateOf);
	if (!date) {
		return undefined;
	}
	const MS_PER_DAY = 86400000;
	const units: Readonly<Record<string, number>> = { d: MS_PER_DAY, y: MS_PER_DAY, w: MS_PER_DAY, ww: 7 * MS_PER_DAY, h: 3600000, n: 60000, s: 1000 };
	const months: Readonly<Record<string, number>> = { m: 1, q: 3, yyyy: 12 };
	let past: 'past December 31, 9999' | 'before January 1, 100' | undefined;
	if (units[interval] !== undefined) {
		const result = date.getTime() + count * units[interval];
		past = result >= Date.UTC(10000, 0, 1) ? 'past December 31, 9999' : result < Date.UTC(100, 0, 1) ? 'before January 1, 100' : undefined;
	} else if (months[interval] !== undefined) {
		// The months are counted in a Long that wraps: DateAdd("yyyy",
		// 2147483647, #1/1/2000#) gives 1/1/1999 (issue #332, measured).
		const month = date.getUTCFullYear() * 12 + date.getUTCMonth() + (((count * months[interval]) % 2 ** 32) + 2 ** 32 + 2 ** 31) % 2 ** 32 - 2 ** 31;
		past = month > 9999 * 12 + 11 ? 'past December 31, 9999' : month < 100 * 12 ? 'before January 1, 100' : undefined;
	}
	if (!past) {
		return undefined;
	}
	const first = dateSlot[0];
	const last = dateSlot[dateSlot.length - 1];
	return {
		displayName: 'DateAdd',
		parameterName: 'Date',
		value: `${source.slice(span.start + first.start, span.start + last.end)}, which the ${count} ${interval} interval(s) carry ${past}`,
		span: { start: span.start + first.start, end: span.start + last.end },
	};
}

/** `DateSerial(y, m, d)` of integer literals, year 100 or later, as a UTC date. */
function dateSerialOfLiterals(source: string, span: Span, toks: readonly VbaToken[], constants: IntegerConstantLookup): Date | undefined {
	let index = 0;
	if (tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.') {
		index = 2;
	}
	if (tokenText(toks[index]) !== 'dateserial' || toks[index + 1]?.rawText !== '(' || matchParenFrom(toks, index + 1) !== toks.length - 1) {
		return undefined;
	}
	const parts = splitTopLevelTokenGroups(toks.slice(index + 2, toks.length - 1), 0, ',').map((group) => integerGroupValue(source, span, group, constants));
	if (parts.length !== 3 || parts.some((part) => part === undefined || part < -32768 || part > 32767)) {
		return undefined;
	}
	const [year, month, day] = parts as number[];
	if (year < 100) {
		return undefined; // read as 19xx or 20xx
	}
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, 1);
	date.setUTCDate(day);
	return date;
}

/**
 * `DateSerial(9999, 13, 1)`: the month and day carry into the year, and a
 * date past December 31, 9999 raises error 5 (issue #189, measured in Excel
 * 16.0). The year alone decides nothing: DateSerial(10000, 0, 1) runs and is
 * December 1, 9999. A year from 0 to 99, after the month carries, is read as
 * 19xx or 20xx, and is judged only where both readings agree.
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
	if (year === undefined || month === undefined || day === undefined) {
		return undefined;
	}
	// A part past the Integer range overflows first (issue #218).
	if ([year, month, day].some((part) => part < -32768 || part > 32767)) {
		return undefined;
	}
	// VBA carries the month into the year first, then reads the year: 0 to
	// 99 as 19xx or 20xx by the system's cut, below 0 counted from 2000. A
	// year past 9999 then raises 5 before the day is added; otherwise the
	// day carries from the first of the month, and the date must fall from
	// January 1, 100 to December 31, 9999. DateSerial(100, 0, 1) is year 99,
	// December, and runs; DateSerial(9999, 13, 0) raises 5 (issues #559 and
	// #604, measured in Excel 16.0 over a 1,716-call grid).
	const carried = year + Math.floor((month - 1) / 12);
	const monthIndex = (((month - 1) % 12) + 12) % 12;
	let before = false;
	if (carried <= 9999) {
		const readings = carried < 0 ? [carried + 2000] : carried < 100 ? [carried + 1900, carried + 2000] : [carried];
		const outside = readings.map((reading) => {
			const date = new Date(0);
			date.setUTCFullYear(reading, monthIndex, 1);
			date.setUTCDate(day);
			const early = date.getUTCFullYear() < 100;
			return { early, outside: early || date.getUTCFullYear() > 9999 };
		});
		// A two-digit year is judged only where both readings agree.
		if (!outside.every((reading) => reading.outside)) {
			return undefined;
		}
		before = outside[0].early;
	}
	const first = call.slots[0].find((t) => t.kind !== 'comment')!;
	const last = [...call.slots[2]].reverse().find((t) => t.kind !== 'comment')!;
	return {
		displayName: call.displayName,
		parameterName: 'Year',
		value: `${year} with month ${month} and day ${day}, a date ${before ? 'before January 1, 100' : 'past December 31, 9999'}`,
		span: { start: span.start + first.start, end: span.start + last.end },
	};
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
	// A project InStr does not take the call from VBA's (issue #280).
	if (!qualifier && lower !== 'instr' && (
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
/**
 * The arguments of a VBA library call that take a whole number with bounds,
 * Mid's Start or Space's Number: each a Long or Integer parameter (issue
 * #298). Empty for a call this table does not know.
 */
export function wholeNumberArguments(name: string, slots: readonly VbaToken[][], host: string | undefined): VbaToken[][] {
	return runtimeArgumentValueSpecs(name, host)
		.filter((spec) => spec.minimum !== undefined)
		.map((spec) => runtimeArgumentValueSlot(slots as VbaToken[][], spec))
		.filter((slot): slot is VbaToken[] => slot !== undefined && slot.length > 0);
}

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
					skippedByEmptyString: true,
				},
				{
					canonicalName: 'InStr',
					parameterName: 'Compare',
					argumentIndex: 3,
					minimum: 0,
					disallowed: databaseCompare,
					minimumSlotCount: 4,
					allowNamed: false,
					skippedByEmptyString: true,
				},
			];
		case 'instrrev':
			return [
				{ canonicalName: 'InStrRev', parameterName: 'Start', argumentIndex: 2, overflowType: 'Long', minimum: -1, disallowed: [0] },
				{ canonicalName: 'InStrRev', parameterName: 'Compare', argumentIndex: 3, overflowType: 'Long', minimum: 0 },
			];
		case 'chr':
			return [{ canonicalName: 'Chr', parameterName: 'CharCode', argumentIndex: 0, overflowType: 'Long', minimum: 0, maximum: 255, stringSuffix: true }];
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
			return [{ canonicalName: 'MonthName', parameterName: 'Month', argumentIndex: 0, overflowType: 'Long', minimum: 1, maximum: 12 }];
		case 'weekdayname':
			return [
				{ canonicalName: 'WeekdayName', parameterName: 'Weekday', argumentIndex: 0, overflowType: 'Long', minimum: 1, maximum: 7 },
				{ canonicalName: 'WeekdayName', parameterName: 'FirstDayOfWeek', argumentIndex: 2, overflowType: 'Long', minimum: 0, maximum: 7 },
			];
		case 'weekday':
			return [{ canonicalName: 'Weekday', parameterName: 'FirstDayOfWeek', argumentIndex: 1, overflowType: 'Long', minimum: 0, maximum: 7 }];
		case 'round':
			return [{ canonicalName: 'Round', parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: 0, maximum: 22 }];
		case 'dateserial':
			// No bound on the Year alone: dateSerialPastMaximum judges the
			// whole date the month and day carry it to (issue #189).
			return ['Year', 'Month', 'Day'].map((parameterName, argumentIndex) => (
				{ canonicalName: 'DateSerial', parameterName, argumentIndex, overflowType: 'Integer' as const }
			));
		case 'strcomp':
			return [{ canonicalName: 'StrComp', parameterName: 'Compare', argumentIndex: 2, overflowType: 'Long', minimum: 0, disallowed: databaseCompare }];
		case 'split':
			return [
				{ canonicalName: 'Split', parameterName: 'Limit', argumentIndex: 2, minimum: -1 },
				{ canonicalName: 'Split', parameterName: 'Compare', argumentIndex: 3, minimum: 0 },
			];
		case 'strconv':
			return [{ canonicalName: 'StrConv', parameterName: 'Conversion', argumentIndex: 1, overflowType: 'Long', accepts: strConvConversionAnyLocale }];
		// A Tristate takes vbUseDefault, vbTrue or vbFalse: -2 to 0 (issue
		// #476, measured in Excel 16.0).
		case 'formatnumber':
		case 'formatcurrency':
		case 'formatpercent': {
			const lower = name.toLowerCase();
			const canonicalName = lower === 'formatnumber' ? 'FormatNumber' : lower === 'formatcurrency' ? 'FormatCurrency' : 'FormatPercent';
			return [
				{ canonicalName, parameterName: 'NumDigitsAfterDecimal', argumentIndex: 1, minimum: -1 },
				...['IncludeLeadingDigit', 'UseParensForNegativeNumbers', 'GroupDigits'].map((parameterName, k) => (
					{ canonicalName, parameterName, argumentIndex: 2 + k, minimum: -2, maximum: 0 }
				)),
			];
		}
		// A rate of -1 divides by zero inside: error 5 (issue #476, measured in
		// Excel 16.0). Pmt and FV take it.
		case 'npv':
			return [{ canonicalName: 'NPV', parameterName: 'Rate', argumentIndex: 0, disallowed: [-1] }];
		case 'irr':
			return [{ canonicalName: 'IRR', parameterName: 'Guess', argumentIndex: 1, disallowed: [-1] }];
		case 'mirr':
			return [
				{ canonicalName: 'MIRR', parameterName: 'FinanceRate', argumentIndex: 1, disallowed: [-1] },
				{ canonicalName: 'MIRR', parameterName: 'ReinvestRate', argumentIndex: 2, disallowed: [-1] },
			];
		case 'environ':
			return [{ canonicalName: 'Environ', parameterName: 'Expression', argumentIndex: 0, minimum: 1, stringSuffix: true }];
		case 'dateadd':
			return [{ canonicalName: 'DateAdd', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS }];
		// A first day of the week runs from 0 to 7 and a first week of the year
		// from 0 to 3; DateDiff reads only the first (issue #262, measured in
		// Excel 16.0: DateDiff with a FirstWeekOfYear of 4 runs).
		case 'datediff':
			return [
				{ canonicalName: 'DateDiff', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS },
				{ canonicalName: 'DateDiff', parameterName: 'FirstDayOfWeek', argumentIndex: 3, minimum: 0, maximum: 7 },
			];
		case 'datepart':
			return [
				{ canonicalName: 'DatePart', parameterName: 'Interval', argumentIndex: 0, allowedStrings: DATE_INTERVALS },
				{ canonicalName: 'DatePart', parameterName: 'FirstDayOfWeek', argumentIndex: 2, minimum: 0, maximum: 7 },
				{ canonicalName: 'DatePart', parameterName: 'FirstWeekOfYear', argumentIndex: 3, minimum: 0, maximum: 3 },
			];
		case 'format':
			return [
				{ canonicalName: 'Format', parameterName: 'FirstDayOfWeek', argumentIndex: 2, minimum: 0, maximum: 7, stringSuffix: true },
				{ canonicalName: 'Format', parameterName: 'FirstWeekOfYear', argumentIndex: 3, minimum: 0, maximum: 3, stringSuffix: true },
			];
		case 'formatdatetime':
			return [{ canonicalName: 'FormatDateTime', parameterName: 'NamedFormat', argumentIndex: 1, minimum: 0, maximum: 4 }];
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

/** Whether an argument is "", vbNullString, or a String local known to hold "". */
function knownEmptyString(slot: readonly VbaToken[], knownStrings: ReadonlyMap<string, string>): boolean {
	const toks = unwrapOuterParens(slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline'));
	if (toks.length !== 1) {
		return false;
	}
	if (toks[0].kind === 'stringLiteral') {
		return stringLiteralValue(toks[0].rawText) === '';
	}
	const lower = tokenName(toks[0])?.toLowerCase();
	return lower === 'vbnullstring' || (lower !== undefined && knownStrings.get(lower) === '');
}

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
		// Asc(Empty) is Asc("") (issue #332, measured in Excel 16.0).
		const known = knownName === 'empty' && spec.canonicalName === 'Asc' ? '' : knownName !== undefined ? knownStrings.get(knownName) : undefined;
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
			: knownName === 'empty'
				? `Empty, which it reads as ""`
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
	// True passes -1, False and Empty 0: Left("abc", True) and Mid("abc",
	// False) raise 5 (issue #434, measured in Excel 16.0).
	const word = toks.length === 1 ? tokenText(toks[0]) : '';
	if (literalValue === undefined && (word === 'true' || word === 'false' || word === 'empty')) {
		literalValue = word === 'true' ? -1 : 0;
	}
	// A Date passes its serial: Chr(#1/2/2000#) is Chr(36527), which raises
	// 5 (issue #332, measured in Excel 16.0).
	if (literalValue === undefined && toks.length === 1 && toks[0].kind === 'dateLiteral') {
		literalValue = dateLiteralSerial(toks[0].rawText);
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
	// A local past the Long range overflows as it converts: Chr(a) with a
	// Double of 1E+300 (issue #336). argument-type-mismatch sees only its type.
	const local = toks.length === 1 && tokenName(toks[0]) !== undefined;
	const pastLong = local && !spec.fractional && expressionValue !== undefined
		&& (expressionValue < OVERFLOW_RANGES.Long.min || expressionValue > OVERFLOW_RANGES.Long.max);
	const verdict = expressionValue === undefined ? 'runs' : pastLong ? 6 : argumentValueVerdict(expressionValue, spec);
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
	// 1E+300 as VBA prints it, not 1e+300.
	const shown = Math.abs(value) >= 1e15 ? value.toExponential().replace('e', 'E') : value;
	return passed === value ? shown : `${shown}, which VBA rounds to ${passed}`;
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
			const strings = pickedValues(valuesAt(stmt), (_lower, value) => (value.kind === 'string' && !value.contentMutated ? value.value as string : undefined));
			const numbers = pickedValues(valuesAt(stmt), (_lower, value) => (value.kind === 'number' && !value.contentMutated ? value.value as number : undefined));
			for (const hit of runtimeConversionValueHits(source, stmt.span, sourceNames, strings, numbers)) {
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
	// Str("abc"), Str("") and Str("True") raise 13, and Str("&H10") runs
	// (issue #332, measured in Excel 16.0).
	str: 'numeric',
	exp: 'numeric', log: 'numeric', sin: 'numeric', cos: 'numeric', tan: 'numeric', atn: 'numeric',
	year: 'date', month: 'date', day: 'date', weekday: 'date', hour: 'date', minute: 'date', second: 'date',
	dateadd: 'date', datepart: 'date', datediff: 'date',
	// FormatNumber("abc") and FormatDateTime("abc") raise 13 (issue #476,
	// measured in Excel 16.0).
	formatnumber: 'numeric', formatcurrency: 'numeric', formatpercent: 'numeric', formatdatetime: 'date',
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
 * Number arguments a String local converts into (issue #332, measured in
 * Excel 16.0): `Chr(s)`, `Space(s)`, `String(s, "a")`, `Left("abc", s)` and
 * the count of `DateAdd` raise 13 when s holds "x" or "".
 */
const NUMBER_SLOTS: Readonly<Record<string, readonly number[]>> = {
	chr: [0], chrw: [0], space: [0], string: [0], left: [1], right: [1], dateadd: [1],
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
	knownNumbers: ReadonlyMap<string, number> = new Map(),
): RuntimeConversionValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeConversionValueHit[] = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const name = tokenName(toks[i]);
		const numberSlots = name ? NUMBER_SLOTS[name.toLowerCase()] : undefined;
		const target = name ? CONVERSION_TARGETS[name.toLowerCase()] ?? (numberSlots ? 'numeric' : undefined) : undefined;
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
		const converted = CONVERSION_TARGETS[name.toLowerCase()] ? CONVERTED_SLOTS[name.toLowerCase()] ?? [0] : [];
		const slots = [...converted.map((index) => ({ index, target, heldOnly: false })), ...(numberSlots ?? []).map((index) => ({ index, target: 'numeric' as const, heldOnly: true }))];
		for (const { index, target, heldOnly } of slots) {
			const slot = (split.slots[index] ?? []).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
			const at = split.spans[index] ?? (slot.length > 0 ? { start: span.start + slot[0].start, end: span.start + slot[slot.length - 1].end } : undefined);
			const displayName = qualified ? `VBA.${name}` : name;
			const held = slot.length === 1 && slot[0].kind === 'identifier' ? knownStrings.get(slot[0].rawText.toLowerCase()) : undefined;
			// A number slot of Chr, Space or Left is judged on a String local
			// only: a literal there is argument-type-mismatch's (issue #332).
			if (slot.length === 1 && ((slot[0].kind === 'stringLiteral' && !heldOnly) || held !== undefined) && at) {
				const value = held ?? stringLiteralValue(slot[0].rawText);
				const lower = name.toLowerCase();
				const invalid = target === 'date'
					? isInvalidDateString(value)
						// A time out of range for every reader of a date, and DateValue
						// or TimeValue of a whole number (issues #262, #444, measured in
						// Excel 16.0): Year("25:00"), DateValue("12"), TimeValue("-1").
						|| isInvalidTimeString(value)
						|| ((lower === 'timevalue' || lower === 'datevalue') && /^[+-]?\d+$/.test(value.replace(/^[ \t]+|[ \t]+$/g, '')))
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
			// A number local known here reads the same way (issue #332).
			const heldNumber = slot.length === 1 && slot[0].kind === 'identifier' ? knownNumbers.get(slot[0].rawText.toLowerCase()) : undefined;
			const serial = DATE_SERIAL_READERS.has(name.toLowerCase()) && !heldOnly ? signedNumericLiteral(slot) ?? heldNumber : undefined;
			if (serial !== undefined && at && (serial >= 2958466 || serial <= -657435)) {
				hits.push({
					displayName,
					name: heldNumber !== undefined && signedNumericLiteral(slot) === undefined ? `${slot[0].rawText}, which holds ${serial} here,` : String(serial),
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
