// Literal tokens the VBE refuses while compiling (MS-VBAL 3.3.2 number
// tokens, 3.3.3 date tokens). Each form below was measured in Excel 16.0
// (build 20326): the refused ones are "Syntax error" (issues #125 and #133),
// and the accepted ones at the edges compile.
//
// Type-suffixed integers (suffix_integer_pct_* oracle cases, and #133):
//   32768%, -32768%, &H10000%, 1E3%, 1.5%          refused   (% is Integer)
//   32767%, &H8000%, &HFFFF%, &O100000%             accepted  (hex wraps to 16 bits)
//   2147483648&, -2147483648&, &H100000000&         refused   (& is Long)
//   2147483647&, &HFFFFFFFF&                        accepted  (hex wraps to 32 bits)
//   9223372036854775808^                            refused   (^ is LongLong)
//   9223372036854775807^                            accepted
// Type-suffixed floats:
//   3.5E+38!                                        refused   (! is Single)
//   3.402823E+38!                                   accepted
//   922337203685477.5808@, 922337203685478@         refused   (@ is Currency)
//   922337203685477.5807@                           accepted
//   1.8E+308#, 1E+309 (unsuffixed)                  refused   (Double)
//   1.79769313486231E+308#, 1E+308, 99999999999999999999   accepted
// Radix prefix with no digits: &H                   refused
// Date literals:
//   #1/1/10000#, #2/30/2000#, #1/0/2000#            refused   (year, day)
//   #25:00#, #1/1/2000 24:00:00#, ##                refused   (time, empty)
//   #12/31/9999#, #1/1/100#, #13/1/2000#, #13:00 PM#  accepted (13/1 reads as 13 January)
//
// The `&` suffix is ambiguous with concatenation: `s = 3000000000&"x"` is
// accepted as `3000000000 & "x"` (oracle suffix_long_amp_glued_concat_accepted),
// so a &-suffixed literal is judged only when nothing that could start an
// operand follows it.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { PushFn } from '../analysisContext';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';

const INTEGER_MAX = 32767;
const LONG_MAX = 2147483647;
const LONGLONG_MAX = 9223372036854775807n;
const SINGLE_MAX = 3.402823e38;
const CURRENCY_MAX = 922337203685477.5807;
const OPERAND_STARTS: ReadonlySet<string> = new Set(['identifier', 'keyword', 'integerLiteral', 'floatLiteral', 'stringLiteral', 'dateLiteral', 'bracketedIdentifier']);

export function checkSuffixedLiteralOverflow(
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const tokens = tokenizeCached(source);
	for (let index = 0; index < tokens.length; index++) {
		const tok = tokens[index];
		const span = { start: tok.start, end: tok.end };
		if (tok.kind === 'floatLiteral') {
			if (activity?.isInactive(span)) {
				continue;
			}
			checkFloat(tok, tokens[index + 1], push);
			continue;
		}
		if (tok.kind === 'dateLiteral') {
			if (activity?.isInactive(span)) {
				continue;
			}
			const problem = dateLiteralProblem(tok.rawText);
			if (problem) {
				push('dateLiteralInvalid', `The date literal ${tok.rawText} ${problem}. VBE rejects this at compile time as a Syntax error.`, span);
			}
			continue;
		}
		if (tok.kind !== 'integerLiteral' || activity?.isInactive(span)) {
			continue;
		}
		checkInteger(tok, tokens[index - 1], tokens[index + 1], push);
	}
}

function checkInteger(tok: VbaToken, previous: VbaToken | undefined, next: VbaToken | undefined, push: PushFn): void {
	const raw = tok.rawText;
	const span = { start: tok.start, end: tok.end };
	const reject = (message: string): void => {
		push('suffixedLiteralOverflow', `${message} VBE rejects this at compile time as a Syntax error.`, span);
	};
	const radix = /^&([hHoO]?)([0-9A-Fa-f]*)([%&^]?)$/.exec(raw);
	if (radix) {
		const [, letter, digits, suffix] = radix;
		if (digits.length === 0) {
			reject(`'${raw}' names a radix with no digits after it.`);
			return;
		}
		const base = letter.toLowerCase() === 'h' ? 16 : 8;
		let value: bigint;
		try {
			value = base === 16 ? BigInt(`0x${digits}`) : BigInt(`0o${digits}`);
		} catch {
			return;
		}
		// A hex or octal literal wraps: 16 bits with %, 32 with &, 64 with ^,
		// and without a suffix 16 then 32 bits as the digits need.
		if (suffix === '%' && value > 0xFFFFn) {
			reject(`The literal '${raw}' does not fit the Integer its '%' suffix asks for: at most four hex digits (&HFFFF).`);
		} else if (suffix === '&' && value > 0xFFFFFFFFn) {
			reject(`The literal '${raw}' does not fit the Long its '&' suffix asks for: at most eight hex digits (&HFFFFFFFF).`);
		} else if (suffix === '^' && value > 0xFFFFFFFFFFFFFFFFn) {
			reject(`The literal '${raw}' does not fit the LongLong its '^' suffix asks for.`);
		}
		// An unsuffixed hex or octal literal wider than 32 bits has not been
		// measured against the VBE and is left alone.
		return;
	}
	const decimal = /^(\d+)([%&^]?)$/.exec(raw);
	if (!decimal) {
		return;
	}
	const [, digits, suffix] = decimal;
	// The sign belongs to the token's value: -32768% is refused as well as
	// 32768% (oracle suffix_integer_pct_neg_compile), so only the magnitude counts.
	const value = BigInt(digits);
	if (suffix === '%' && value > BigInt(INTEGER_MAX)) {
		reject(`The literal '${raw}' is outside the Integer range -32768 to 32767 of its '%' type suffix.`);
	} else if (suffix === '&' && value > BigInt(LONG_MAX)) {
		// `3000000000&"x"` reads as concatenation; only a `&` nothing follows,
		// or an operator follows, is the Long suffix.
		if (next === undefined || next.kind === 'newline' || next.kind === 'colon' || next.kind === 'comment' || next.kind === 'operator' || next.kind === 'punctuation') {
			if (!(next && OPERAND_STARTS.has(next.kind))) {
				reject(`The literal '${raw}' is outside the Long range -2147483648 to 2147483647 of its '&' type suffix.`);
			}
		}
	} else if (suffix === '^' && value > LONGLONG_MAX) {
		reject(`The literal '${raw}' is outside the LongLong range of its '^' type suffix (at most 9223372036854775807).`);
	}
	void previous;
}

function checkFloat(tok: VbaToken, next: VbaToken | undefined, push: PushFn): void {
	const raw = tok.rawText;
	const span = { start: tok.start, end: tok.end };
	const suffix = /[!#@]$/.exec(raw)?.[0] ?? '';
	const value = Number(raw.replace(/[dD]/g, 'E').replace(/[!#@]$/, ''));
	if (!Number.isFinite(value)) {
		push('floatLiteralOverflow', `The literal '${raw}' is outside the Double range (about 1.8E+308). VBE rejects this at compile time as a Syntax error.`, span);
		return;
	}
	if (suffix === '!' && Math.abs(value) > SINGLE_MAX) {
		push('floatLiteralOverflow', `The literal '${raw}' is outside the Single range (about 3.402823E+38) of its '!' type suffix. VBE rejects this at compile time as a Syntax error.`, span);
		return;
	}
	if (suffix === '@' && currencyOverflows(raw)) {
		push('floatLiteralOverflow', `The literal '${raw}' is outside the Currency range (at most 922337203685477.5807) of its '@' type suffix. VBE rejects this at compile time as a Syntax error.`, span);
		return;
	}
	// `1.5%` and `1E3%`: the Integer suffix on a number that is not an integer token.
	if (next && next.kind === 'unknown' && next.rawText === '%' && next.start === tok.end) {
		push(
			'suffixedLiteralOverflow',
			`The literal '${raw}%' puts the '%' Integer type suffix on a fractional or exponent literal, which has no Integer form. VBE rejects this at compile time as a Syntax error.`,
			{ start: tok.start, end: next.end },
		);
	}
}

/**
 * Whether a Currency literal exceeds 922337203685477.5807 in magnitude,
 * compared in scaled integers: at 9.2E+14 a Double's spacing is 0.125, so
 * .5807 and .5808 would read as the same number as floats.
 */
function currencyOverflows(raw: string): boolean {
	const plain = /^(\d+)(?:\.(\d*))?@$/.exec(raw);
	if (!plain) {
		const value = Number(raw.replace(/[dD]/g, 'E').replace(/@$/, ''));
		return Number.isFinite(value) && Math.abs(value) > CURRENCY_MAX;
	}
	const fraction = (plain[2] ?? '').padEnd(4, '0');
	if (fraction.length > 4) {
		return false; // more than four decimals: rounding the VBE applies is not modelled
	}
	return BigInt(plain[1]) * 10000n + BigInt(fraction) > 9223372036854775807n;
}

/**
 * What is wrong with a `#...#` date literal, or undefined when it is one
 * the VBE accepts or one this check does not judge (named months and other
 * regional forms are left alone).
 */
function dateLiteralProblem(raw: string): string | undefined {
	const body = raw.slice(1, -1).trim();
	if (body.length === 0) {
		return 'is empty';
	}
	const match = /^(?:(\d{1,2})\/(\d{1,2})\/(\d{1,5}))?\s*(?:(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?)?$/.exec(body);
	if (!match || (match[1] === undefined && match[4] === undefined)) {
		return undefined;
	}
	const [, first, second, yearText, hourText, minuteText, secondText, meridiem] = match;
	if (first !== undefined) {
		let month = Number(first);
		let day = Number(second);
		const year = Number(yearText);
		if (year > 9999) {
			return 'names a year past 9999';
		}
		if (month > 12) {
			// The VBE reads #13/1/2000# as 13 January when the first number
			// cannot be a month and the second can.
			if (day <= 12) {
				[month, day] = [day, month];
			} else {
				return `names month ${month}, which no calendar has`;
			}
		}
		if (month < 1) {
			return 'names month 0';
		}
		const fullYear = yearText.length <= 2 ? (year < 30 ? 2000 + year : 1900 + year) : year;
		const daysInMonth = new Date(Date.UTC(fullYear, month, 0)).getUTCDate();
		if (day < 1 || day > daysInMonth) {
			return `names day ${day} in a month of ${daysInMonth} days`;
		}
	}
	if (hourText !== undefined) {
		const hour = Number(hourText);
		const minute = Number(minuteText);
		const second = secondText === undefined ? 0 : Number(secondText);
		if (hour > 23 || (meridiem && hour > 12 && hour > 23)) {
			return `names hour ${hour}; hours run 0 to 23`;
		}
		if (minute > 59) {
			return `names minute ${minute}; minutes run 0 to 59`;
		}
		if (second > 59) {
			return `names second ${second}; seconds run 0 to 59`;
		}
	}
	return undefined;
}
