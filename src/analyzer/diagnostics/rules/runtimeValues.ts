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
import type {
	ModuleNode,
	Span,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import { type PushFn } from '../analysisContext';
import {
	type CallableTypeSignature,
	emptyArgSplit,
	splitArgSlots,
} from '../callExtraction';
import { collectModuleLiteralIntegerConstants } from '../constExpr';
import { bankersRound, isBareOrVbaQualifiedIntrinsicCall } from '../rules/shared';
import { isInvalidBooleanString, isInvalidDateString, isInvalidNumericString } from '../stringConversion';
import {
	callableTypeSignaturesFor,
	knownLocalLiteralValuesAt,
	type KnownLocalValue,
	namedArgumentSlot,
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
	statementTokens,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

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
}

interface RuntimeArgumentValueHit {
	displayName: string;
	parameterName: string;
	value: number | string;
	span: Span;
}

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
		const lookup: IntegerConstantLookup = {
			get: (name) => {
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
			for (const hit of runtimeArgumentValueHits(source, stmt.span, moduleSignatures, env, lookup, knownStrings, sourceNames, host)) {
				push(
					'runtimeArgumentValue',
					`Argument '${hit.parameterName}' of '${hit.displayName}' is ${hit.value}; this will raise Run-time error '5': Invalid procedure call or argument.`,
					hit.span,
				);
			}
			for (const hit of runtimeStatementValueHits(source, stmt.span, lookup, knownStringLengths, knownStrings, sourceNames)) {
				push('runtimeArgumentValue', hit.message, hit.span);
			}
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
	knownStrings: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
	host: string | undefined,
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
				? integerArgumentOutsideBounds(source, slot, span.start, spec, constants, knownStrings)
				: undefined;
			if (!literal) {
				continue;
			}
			hits.push({
				displayName: call.displayName,
				parameterName: spec.parameterName,
				value: literal.value,
				span: literal.span,
			});
		}
		const overflow = dateAddPastMaximum(source, span, call, constants) ?? dateSerialPastMaximum(source, span, call, constants);
		if (overflow) {
			hits.push(overflow);
		}
	}
	return hits;
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
	if (specs.length === 0) {
		return undefined;
	}
	if (suffix && !specs[0].stringSuffix) {
		return undefined;
	}
	const lower = specs[0].canonicalName.toLowerCase();
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
		displayName: `${specs[0].canonicalName}${suffix}`,
		specs,
		slots: split.slots,
	};
}

/**
 * The bounds each runtime function's arguments must keep to compile-and-run
 * clean. The first nine were VBE-oracle-backed from the start; the rest were
 * measured one call at a time in Excel 16.0 (build 20326, 2026-09-26, issue
 * #118): every listed value raises error 5 every time, and the nearest value
 * that runs - Mid("abc", 10), Round(1.5, 0), Weekday(Date, 7), Environ(1) -
 * stays quiet. `vbDatabaseCompare` (2) is valid only where Access is the
 * host, so InStr's Compare bound depends on the host.
 */
function runtimeArgumentValueSpecs(name: string, host: string | undefined): readonly RuntimeArgumentValueSpec[] {
	switch (name.toLowerCase()) {
		case 'left':
			return [{ canonicalName: 'Left', parameterName: 'Length', argumentIndex: 1, minimum: 0, stringSuffix: true }];
		case 'right':
			return [{ canonicalName: 'Right', parameterName: 'Length', argumentIndex: 1, minimum: 0, stringSuffix: true }];
		case 'string':
			return [
				{ canonicalName: 'String', parameterName: 'Number', argumentIndex: 0, minimum: 0, stringSuffix: true },
				{ canonicalName: 'String', parameterName: 'Character', argumentIndex: 1, emptyStringRaises: true, stringSuffix: true },
			];
		case 'space':
			return [{ canonicalName: 'Space', parameterName: 'Number', argumentIndex: 0, minimum: 0, stringSuffix: true }];
		case 'mid':
			return [
				{ canonicalName: 'Mid', parameterName: 'Start', argumentIndex: 1, minimum: 1, stringSuffix: true },
				{ canonicalName: 'Mid', parameterName: 'Length', argumentIndex: 2, minimum: 0, stringSuffix: true },
			];
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
					minimumSlotCount: 3,
					allowNamed: false,
				},
				{
					canonicalName: 'InStr',
					parameterName: 'Compare',
					argumentIndex: 3,
					minimum: 0,
					maximum: host === 'access' ? 2 : 1,
					minimumSlotCount: 4,
					allowNamed: false,
				},
			];
		case 'instrrev':
			return [{ canonicalName: 'InStrRev', parameterName: 'Start', argumentIndex: 2, minimum: -1, disallowed: [0] }];
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
			return [{ canonicalName: 'DateSerial', parameterName: 'Year', argumentIndex: 0 }];
		case 'strcomp':
			return [{ canonicalName: 'StrComp', parameterName: 'Compare', argumentIndex: 2, minimum: 0, maximum: host === 'access' ? 2 : 1 }];
		case 'split':
			return [{ canonicalName: 'Split', parameterName: 'Limit', argumentIndex: 2, minimum: -1 }];
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
	knownStrings: ReadonlyMap<string, string>,
): { value: number | string; span: Span } | undefined {
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
		if (integerArgumentValueInBounds(literalValue, spec)) {
			return undefined;
		}
		return {
			value: shownArgumentValue(literalValue, spec),
			span: { start: sliceStart + start!, end: sliceStart + literal.end },
		};
	}

	const expressionValue = evaluateIntegerConstantExpression(
		source.slice(sliceStart + toks[0].start, sliceStart + toks[toks.length - 1].end),
		constants,
	);
	if (expressionValue === undefined || integerArgumentValueInBounds(expressionValue, spec)) {
		return undefined;
	}
	return {
		value: shownArgumentValue(expressionValue, spec),
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
	};
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
): ProcedureStatementVisitor {
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			for (const hit of runtimeConversionValueHits(source, stmt.span, sourceNames)) {
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
	csng: 'numeric', cdbl: 'numeric', ccur: 'numeric', cdec: 'numeric',
	cbool: 'boolean',
	cdate: 'date', cvdate: 'date', datevalue: 'date', timevalue: 'date',
};

function runtimeConversionValueHits(
	source: string,
	span: Span,
	sourceNames: SourceNameScope,
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
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1 || firstSlot[0].kind !== 'stringLiteral') {
			continue;
		}
		const value = stringLiteralValue(firstSlot[0].rawText);
		const invalid = target === 'date'
			? isInvalidDateString(value)
			: target === 'boolean'
				? isInvalidBooleanString(value)
				: isInvalidNumericString(value);
		if (!invalid) {
			continue;
		}
		hits.push({
			displayName: qualified ? `VBA.${name}` : name,
			name: firstSlot[0].rawText,
			target: target === 'date' ? 'Date' : target === 'boolean' ? 'Boolean' : 'a number',
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
		});
	}
	return hits;
}
