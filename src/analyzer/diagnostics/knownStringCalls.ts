// The whole numbers InStr, InStrRev, Len, Asc and AscW return for strings the
// code states (issue #201).
//
// The "not found" bug reads `Left$(s, InStr(s, " ") - 1)`: with no space
// InStr returns 0, and Left$ gets -1. The value rules evaluate an argument's
// arithmetic, but a call inside it stopped them. Here each such call whose
// strings are literals, or locals holding a literal, becomes its number, so
// the arithmetic around it can be evaluated. Measured in Excel 16.0:
//
//  - InStr([start,] s1, s2 [, compare]) is the first position of s2 in s1 at
//    or after start, 0 when there is none, when s1 is "" or when start is
//    past the end, and start itself when s2 is "", past the end too.
//  - InStrRev(s1, s2 [, start [, compare]]) is the last position of s2 in s1
//    ending at or before start (-1, the default, is the end), 0 when there is
//    none or start is past the end, and start (the length for -1) when s2 is "".
//  - Comparison is binary unless the call passes vbTextCompare (1) or the
//    module says Option Compare Text; Option Compare Database follows the
//    database's locale, so a call is decided there only when both agree.
//  - Asc of a character above 127 depends on the code page and is left alone.

import type { VbaToken } from '../lexer/tokenKinds';
import { matchParenFrom, splitTopLevelTokenGroups } from '../lexer/tokenHelpers';
import { isBareOrVbaQualifiedIntrinsicCall } from './rules/shared';
import { stringLiteralValue } from './typeInference';
import { tokenName, tokenText } from './walker';

export type ModuleCompare = 'binary' | 'text' | 'database';

export interface KnownStringCallContext {
	/** Locals whose value is a known literal, by lower-cased name. */
	knownStrings: ReadonlyMap<string, string>;
	/** The whole number an argument's text evaluates to, if any. */
	integerValue: (text: string) => number | undefined;
	/** Whether the project declares a procedure of this name, hiding VBA's. */
	shadowed: (name: string) => boolean;
	compare: ModuleCompare;
	/** The date a local holds, by lower-cased name, where it is known (issue #559). */
	dateOf?: (lower: string) => Date | undefined;
}

let lastCompare: { source: string; compare: ModuleCompare } | undefined;

/**
 * The module's Option Compare, from its source text. The array rules ask once
 * per procedure, so the last module's answer is kept.
 */
export function moduleCompare(source: string): ModuleCompare {
	if (lastCompare?.source === source) {
		// Adopt an equal-content caller string so subsequent hits compare by identity.
		lastCompare.source = source;
		return lastCompare.compare;
	}
	const match = /^[ \t]*Option[ \t]+Compare[ \t]+(Binary|Text|Database)\b/im.exec(source);
	const compare = (match?.[1].toLowerCase() as ModuleCompare | undefined) ?? 'binary';
	lastCompare = { source, compare };
	return compare;
}

const FOLDED = new Set(['instr', 'instrrev', 'len', 'asc', 'ascw', 'year', 'month', 'day', 'datediff']);

/** A `#m/d/yyyy#` date literal as a UTC date, or undefined for any other spelling. */
export function parseDateLiteral(raw: string): Date | undefined {
	// A year of three digits is that year: #1/1/100# (issue #262).
	const match = /^#\s*(\d{1,2})\/(\d{1,2})\/(\d{3,4})\s*(?:(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?\s*#$/i.exec(raw);
	if (!match) {
		return undefined;
	}
	const month = Number(match[1]);
	const day = Number(match[2]);
	const year = Number(match[3]);
	if (month < 1 || month > 12 || day < 1 || day > 31 || year < 100) {
		return undefined;
	}
	let hour = Number(match[4] ?? 0);
	if (match[7]) {
		hour = hour % 12 + (match[7].toUpperCase() === 'PM' ? 12 : 0);
	}
	const date = new Date(0);
	date.setUTCFullYear(year, month - 1, day);
	date.setUTCHours(hour, Number(match[5] ?? 0), Number(match[6] ?? 0), 0);
	return date;
}

/**
 * The date a date expression names, where it is known (issue #510): a date
 * literal, `DateSerial` of whole numbers from year 100 on, and `DateValue`
 * or `CDate` of a string every locale reads alike, "12/31/9999" with a day
 * past 12 or "9999-12-31". Undefined for anything else, and for a date past
 * the range.
 */
export function knownDate(arg: readonly VbaToken[], integerValue: (toks: readonly VbaToken[]) => number | undefined, dateOf?: (lower: string) => Date | undefined): Date | undefined {
	let toks = arg.filter((tok) => tok.kind !== 'comment');
	while (toks.length > 2 && toks[0].rawText === '(' && matchParenFrom(toks, 0) === toks.length - 1) {
		toks = toks.slice(1, -1);
	}
	if (toks.length === 1 && toks[0].kind === 'dateLiteral') {
		return parseDateLiteral(toks[0].rawText);
	}
	const local = toks.length === 1 && dateOf ? tokenName(toks[0])?.toLowerCase() : undefined;
	if (local) {
		return dateOf!(local);
	}
	const at = tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.' ? 2 : 0;
	const name = tokenText(toks[at]);
	if (toks[at + 1]?.rawText !== '(' || matchParenFrom(toks, at + 1) !== toks.length - 1) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(toks, at + 2, ',', toks.length - 1);
	if (name === 'dateserial' && args.length === 3) {
		const parts = args.map((group) => integerValue(group));
		if (parts.some((part) => part === undefined || part < -32768 || part > 32767) || parts[0]! < 100) {
			return undefined;
		}
		const date = new Date(0);
		date.setUTCFullYear(parts[0]!, parts[1]! - 1, 1);
		date.setUTCDate(parts[2]!);
		return inDateRange(date) ? date : undefined;
	}
	// CDate of a whole number counts days from December 30, 1899:
	// CDate(-10000) is August 13, 1872 (issue #559, measured in Excel 16.0).
	const days = name === 'cdate' && args.length === 1 && args[0][0]?.kind !== 'stringLiteral' ? integerValue(args[0]) : undefined;
	if (days !== undefined) {
		const date = new Date(Date.UTC(1899, 11, 30) + days * 86400000);
		return inDateRange(date) ? date : undefined;
	}
	if ((name === 'datevalue' || name === 'cdate') && args.length === 1 && args[0].length === 1 && args[0][0].kind === 'stringLiteral') {
		const text = stringLiteralValue(args[0][0].rawText).trim();
		const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
		const slashed = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(text);
		let parts: [number, number, number] | undefined;
		if (iso) {
			parts = [Number(iso[1]), Number(iso[2]), Number(iso[3])];
		} else if (slashed) {
			const [a, b] = [Number(slashed[1]), Number(slashed[2])];
			// Only an order every locale agrees on: one of the two is past 12.
			parts = a > 12 && b <= 12 ? [Number(slashed[3]), b, a] : b > 12 && a <= 12 ? [Number(slashed[3]), a, b] : undefined;
		}
		if (!parts || parts[0] < 100 || parts[1] < 1 || parts[1] > 12 || parts[2] < 1) {
			return undefined;
		}
		const date = new Date(0);
		date.setUTCFullYear(parts[0], parts[1] - 1, parts[2]);
		return date.getUTCDate() === parts[2] && inDateRange(date) ? date : undefined;
	}
	return undefined;
}

function inDateRange(date: Date): boolean {
	return date.getTime() >= Date.UTC(100, 0, 1) && date.getUTCFullYear() <= 9999;
}

/** DateDiff of known dates, for the intervals that count whole units. */
function dateDiff(interval: string, from: Date, to: Date): number | undefined {
	const DAY = 86400000;
	const units: Readonly<Record<string, number>> = { h: 3600000, n: 60000, s: 1000 };
	const days = (date: Date): number => Math.floor(date.getTime() / DAY);
	switch (interval) {
		case 'd':
		case 'y':
			return days(to) - days(from);
		case 'w':
			return Math.trunc((days(to) - days(from)) / 7);
		case 'm':
			return (to.getUTCFullYear() - from.getUTCFullYear()) * 12 + to.getUTCMonth() - from.getUTCMonth();
		case 'q':
			return (to.getUTCFullYear() - from.getUTCFullYear()) * 4 + Math.floor(to.getUTCMonth() / 3) - Math.floor(from.getUTCMonth() / 3);
		case 'yyyy':
			return to.getUTCFullYear() - from.getUTCFullYear();
	}
	const unit = units[interval];
	return unit === undefined ? undefined : Math.floor(to.getTime() / unit) - Math.floor(from.getTime() / unit);
}

/**
 * The expression's text with each foldable call replaced by its number, or
 * undefined when it has none.
 */
export function foldKnownStringCalls(toks: readonly VbaToken[], ctx: KnownStringCallContext): string | undefined {
	const out: string[] = [];
	let folded = false;
	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		const lower = tokenText(tok) ?? '';
		const close = FOLDED.has(lower) && toks[i + 1]?.rawText === '(' && isBareOrVbaQualifiedIntrinsicCall(toks, i) && !ctx.shadowed(lower)
			? matchParenFrom(toks, i + 1)
			: -1;
		const value = close > i + 1 ? callValue(lower, splitTopLevelTokenGroups(toks, i + 2, ',', close), ctx) : undefined;
		if (value === undefined) {
			out.push(tok.rawText);
			continue;
		}
		if (toks[i - 1]?.rawText === '.') {
			out.splice(-2, 2); // `VBA.InStr(...)`
		}
		out.push(value < 0 ? `(${value})` : String(value));
		folded = true;
		i = close;
	}
	return folded ? out.join(' ') : undefined;
}

function callValue(name: string, args: VbaToken[][], ctx: KnownStringCallContext): number | undefined {
	switch (name) {
		case 'len':
			return args.length === 1 ? knownString(args[0], ctx)?.length : undefined;
		case 'asc':
		case 'ascw': {
			const text = args.length === 1 ? knownString(args[0], ctx) : undefined;
			const code = text ? text.charCodeAt(0) : undefined;
			return code !== undefined && (name === 'ascw' || code < 128) ? code : undefined;
		}
		case 'instr':
			return inStr(args, ctx);
		case 'instrrev':
			return inStrRev(args, ctx);
		// Parts of a known date, and DateDiff between two (issue #510).
		case 'year':
		case 'month':
		case 'day': {
			const date = args.length === 1 ? knownDate(args[0], (toks) => wholeNumber(toks, ctx), ctx.dateOf) : undefined;
			return !date ? undefined : name === 'year' ? date.getUTCFullYear() : name === 'month' ? date.getUTCMonth() + 1 : date.getUTCDate();
		}
		case 'datediff': {
			const interval = args.length === 3 ? knownString(args[0], ctx)?.toLowerCase() : undefined;
			const from = args.length === 3 ? knownDate(args[1], (toks) => wholeNumber(toks, ctx), ctx.dateOf) : undefined;
			const to = args.length === 3 ? knownDate(args[2], (toks) => wholeNumber(toks, ctx), ctx.dateOf) : undefined;
			return interval === undefined || !from || !to ? undefined : dateDiff(interval, from, to);
		}
		default:
			return undefined;
	}
}

function inStr(args: VbaToken[][], ctx: KnownStringCallContext): number | undefined {
	// With a start, the strings move one place right; a compare needs a start.
	const withStart = args.length >= 3;
	if (args.length < 2 || args.length > 4) {
		return undefined;
	}
	const start = withStart ? wholeNumber(args[0], ctx) : 1;
	const s1 = knownString(args[withStart ? 1 : 0], ctx);
	const s2 = knownString(args[withStart ? 2 : 1], ctx);
	const text = compareMode(args[3], ctx);
	if (start === undefined || start < 1 || s1 === undefined || s2 === undefined || text === undefined) {
		return undefined;
	}
	return decided(s1, s2, text, (a, b) => {
		// An empty s1 gives 0, whatever s2 is: InStr("", "") is 0 (issue
		// #509, measured in Excel 16.0).
		if (a.length === 0) {
			return 0;
		}
		// An empty s2 is found at start, even past the end: InStr(5, "abc", "") is 5.
		if (b.length === 0) {
			return start;
		}
		return a.indexOf(b, start - 1) + 1;
	});
}

function inStrRev(args: VbaToken[][], ctx: KnownStringCallContext): number | undefined {
	if (args.length < 2 || args.length > 4) {
		return undefined;
	}
	const s1 = knownString(args[0], ctx);
	const s2 = knownString(args[1], ctx);
	const start = args[2] ? wholeNumber(args[2], ctx) : -1;
	const text = compareMode(args[3], ctx);
	if (s1 === undefined || s2 === undefined || start === undefined || (start < 1 && start !== -1) || text === undefined) {
		return undefined;
	}
	return decided(s1, s2, text, (a, b) => {
		const end = start === -1 ? a.length : start;
		if (end > a.length) {
			return 0;
		}
		if (b.length === 0) {
			return end;
		}
		return end < b.length ? 0 : a.lastIndexOf(b, end - b.length) + 1;
	});
}

/**
 * The search's result under the comparison that applies, or undefined where
 * it cannot be told: a text comparison over characters beyond ASCII, whose
 * case rules are the locale's, or a Database comparison the two disagree on.
 */
function decided(s1: string, s2: string, text: boolean | 'either', search: (a: string, b: string) => number): number | undefined {
	const binary = search(s1, s2);
	const ascii = /^[\x00-\x7f]*$/.test(s1 + s2);
	const insensitive = ascii ? search(s1.toLowerCase(), s2.toLowerCase()) : undefined;
	if (text === 'either') {
		return insensitive === binary ? binary : undefined;
	}
	return text ? insensitive : binary;
}

/** True for a text comparison, false for binary, 'either' under Option Compare Database. */
function compareMode(arg: VbaToken[] | undefined, ctx: KnownStringCallContext): boolean | 'either' | undefined {
	if (!arg) {
		return ctx.compare === 'text' ? true : ctx.compare === 'database' ? 'either' : false;
	}
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	const word = toks.length === 1 ? tokenName(toks[0])?.toLowerCase() : undefined;
	const value = word === 'vbbinarycompare' ? 0 : word === 'vbtextcompare' ? 1 : wholeNumber(toks, ctx);
	return value === 0 ? false : value === 1 ? true : undefined;
}

function knownString(arg: readonly VbaToken[], ctx: KnownStringCallContext): string | undefined {
	return foldStringExpression(arg, {
		nameValue: (tok) => ctx.knownStrings.get(tokenName(tok)?.toLowerCase() ?? ''),
		integerValue: (toks) => wholeNumber(toks, ctx),
		shadowed: ctx.shadowed,
	});
}

/** What a string expression is folded with. */
export interface StringFoldContext {
	/** A String local's known value, by lowercased name. */
	nameValue: (tok: VbaToken) => string | undefined;
	/** A whole-number argument's value. */
	integerValue: (toks: readonly VbaToken[]) => number | undefined;
	/** Whether the project declares a procedure of this name, hiding VBA's. */
	shadowed?: (name: string) => boolean;
}

/** The VBA string functions the folder runs, each with or without its `$`. */
const STRING_FUNCTIONS: ReadonlySet<string> = new Set(['left', 'right', 'mid', 'lcase', 'ucase', 'strreverse', 'trim', 'ltrim', 'rtrim', 'space', 'string', 'replace']);

/**
 * The text a string expression gives, where every part is known: literals,
 * String locals known to hold one, `&` between them, and Left, Right, Mid,
 * LCase, UCase, StrReverse, Trim, LTrim, RTrim, Space, String and Replace of
 * known arguments (issue #509). Undefined for anything else, and for a call
 * that would raise.
 */
export function foldStringExpression(arg: readonly VbaToken[], ctx: StringFoldContext): string | undefined {
	let toks = arg.filter((tok) => tok.kind !== 'comment');
	while (toks.length > 2 && toks[0].rawText === '(' && matchParenFrom(toks, 0) === toks.length - 1) {
		toks = toks.slice(1, -1);
	}
	if (toks.length === 0) {
		return undefined;
	}
	const parts = splitTopLevelTokenGroups(toks, 0, '&', toks.length);
	if (parts.length > 1) {
		const texts = parts.map((part) => foldStringExpression(part, ctx));
		return texts.every((text) => text !== undefined) ? texts.join('') : undefined;
	}
	if (toks.length === 1) {
		if (toks[0].kind === 'stringLiteral') {
			return stringLiteralValue(toks[0].rawText);
		}
		const name = tokenName(toks[0])?.toLowerCase();
		return name === undefined || toks[0].kind === 'keyword' ? (tokenText(toks[0]) === 'vbnullstring' ? '' : undefined) : ctx.nameValue(toks[0]);
	}
	// `Mid$(s, 1)` lexes as Mid, a `$` and the arguments.
	let at = tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.' ? 2 : 0;
	const name = tokenText(toks[at]);
	at += toks[at + 1]?.rawText === '$' ? 1 : 0;
	if (!STRING_FUNCTIONS.has(name) || toks[at + 1]?.rawText !== '(' || matchParenFrom(toks, at + 1) !== toks.length - 1 || ctx.shadowed?.(name)) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(toks, at + 2, ',', toks.length - 1);
	const text = (k: number): string | undefined => (args[k] ? foldStringExpression(args[k], ctx) : undefined);
	const whole = (k: number): number | undefined => (args[k] && args[k].length > 0 ? ctx.integerValue(args[k]) : undefined);
	switch (name) {
		case 'left':
		case 'right': {
			const s = text(0);
			const n = whole(1);
			if (args.length !== 2 || s === undefined || n === undefined || n < 0) {
				return undefined;
			}
			return name === 'left' ? s.slice(0, n) : s.slice(Math.max(0, s.length - n));
		}
		case 'mid': {
			const s = text(0);
			const start = whole(1);
			const length = args.length === 3 ? whole(2) : s?.length;
			if (args.length < 2 || args.length > 3 || s === undefined || start === undefined || length === undefined || start < 1 || length < 0) {
				return undefined;
			}
			return s.slice(start - 1, start - 1 + length);
		}
		case 'lcase':
		case 'ucase':
		case 'strreverse':
		case 'trim':
		case 'ltrim':
		case 'rtrim': {
			const s = args.length === 1 ? text(0) : undefined;
			if (s === undefined || (!/^[\x00-\x7f]*$/.test(s) && (name === 'lcase' || name === 'ucase'))) {
				return undefined;
			}
			return name === 'lcase' ? s.toLowerCase()
				: name === 'ucase' ? s.toUpperCase()
				: name === 'strreverse' ? [...s].reverse().join('')
				: name === 'trim' ? s.replace(/^ +| +$/g, '')
				: name === 'ltrim' ? s.replace(/^ +/, '')
				: s.replace(/ +$/, '');
		}
		case 'space': {
			const n = args.length === 1 ? whole(0) : undefined;
			return n === undefined || n < 0 || n > 65535 ? undefined : ' '.repeat(n);
		}
		case 'string': {
			const n = whole(0);
			const c = text(1);
			return args.length !== 2 || n === undefined || n < 0 || n > 65535 || !c ? undefined : c[0].repeat(n);
		}
		case 'replace': {
			const [s, find, by] = [text(0), text(1), text(2)];
			if (args.length !== 3 || s === undefined || find === undefined || by === undefined) {
				return undefined;
			}
			// A letter in find matches by Option Compare, which this does not read.
			return /[A-Za-z]/.test(find) ? undefined : find.length === 0 ? s : s.split(find).join(by);
		}
	}
	return undefined;
}

/** A whole-number argument, itself folded first: `InStr(InStr(s, "X") + 1, s, "X")`. */
function wholeNumber(arg: readonly VbaToken[], ctx: KnownStringCallContext): number | undefined {
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	if (toks.length === 0) {
		return undefined;
	}
	const text = foldKnownStringCalls(toks, ctx) ?? toks.map((tok) => tok.rawText).join(' ');
	const value = ctx.integerValue(text);
	return value !== undefined && Number.isInteger(value) ? value : undefined;
}
