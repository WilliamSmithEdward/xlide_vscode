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
}

/** The module's Option Compare, from its source text. */
export function moduleCompare(source: string): ModuleCompare {
	const match = /^[ \t]*Option[ \t]+Compare[ \t]+(Binary|Text|Database)\b/im.exec(source);
	return (match?.[1].toLowerCase() as ModuleCompare | undefined) ?? 'binary';
}

const FOLDED = new Set(['instr', 'instrrev', 'len', 'asc', 'ascw']);

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
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	if (toks.length !== 1) {
		return undefined;
	}
	if (toks[0].kind === 'stringLiteral') {
		return stringLiteralValue(toks[0].rawText);
	}
	const name = tokenName(toks[0])?.toLowerCase();
	return name === undefined ? undefined : ctx.knownStrings.get(name);
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
