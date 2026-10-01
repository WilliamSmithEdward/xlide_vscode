// What VBA makes of a string it converts to a number, a Boolean or a Date
// (issue #188). Measured in Excel 16.0 (build 20326, en-US, 2026-09-29),
// with `x = "..."` into a typed variable and through CLng, CDbl, CBool and
// CDate, which agree.
//
// A string converts through the locale: its decimal point, thousands
// separator and currency symbol. "2.5", "1,000", "1 000" and "$5" read
// differently, or not at all, under another locale, so a string is judged
// invalid only when no locale can read it:
//
//  - Numbers. Hex and octal strings convert like the literal ("&HFF" is 255,
//    "&HFFFF" is -1, "&17" is 15). Otherwise a sign before or after, a
//    parenthesized negative, a currency symbol at either end, digits with
//    separators and an exponent run. Empty, no digits at all ("abc", ".",
//    "$"), a sign before &H ("-&H10"), and any other character among the
//    digits ("4x2", "5%", "1/2/2020", "12:30") raise 13.
//  - Boolean. "True" and "False" in any case, exactly, and any number, which
//    is True unless zero. " True " with blanks raises 13, as do words.
//  - Date. Any number converts, "$5" included. No digits at all ("May",
//    "Monday", "True", ".") raises 13, and so does a character no date or
//    number uses ("5%", "-&H10").

import { parseVbaIntegerLiteral } from '../constants/integerConstantExpression';

/** How a string converts to a number: never, or with the value when every locale reads it alike. */
export type NumericStringVerdict = { kind: 'invalid' } | { kind: 'number'; value?: number };

const BLANK_EDGES = /^[ \t]+|[ \t]+$/g;

/** The Date serials of 1/1/100 and 12/31/9999, the range a Date holds. */
const DATE_SERIAL_MIN = -657434;
const DATE_SERIAL_MAX = 2958465;

export function numericStringVerdict(text: string): NumericStringVerdict {
	const trimmed = text.replace(BLANK_EDGES, '');
	if (trimmed.length === 0) {
		return { kind: 'invalid' };
	}
	if (/^&[Hh][0-9A-Fa-f]+$|^&[Oo]?[0-7]+$/.test(trimmed)) {
		// "&17" is octal like "&O17".
		const value = parseVbaIntegerLiteral(/^&[0-7]/.test(trimmed) ? `&O${trimmed.slice(1)}` : trimmed);
		return value === undefined ? { kind: 'number' } : { kind: 'number', value };
	}
	if (!/[0-9]/.test(trimmed)) {
		return { kind: 'invalid' };
	}
	let body = trimmed;
	let negative = false;
	let exact = true;
	for (let pass = 0; pass < 6; pass++) {
		const before = body;
		if (body.startsWith('(') && body.endsWith(')')) {
			body = body.slice(1, -1).replace(BLANK_EDGES, '');
			negative = !negative;
		} else if (/^[+-]/.test(body)) {
			negative = body[0] === '-' ? !negative : negative;
			body = body.slice(1).replace(BLANK_EDGES, '');
		} else if (/[+-]$/.test(body)) {
			negative = body[body.length - 1] === '-' ? !negative : negative;
			body = body.slice(0, -1).replace(BLANK_EDGES, '');
		} else {
			// A currency symbol, which some locale spells in letters ("kr").
			body = body.replace(/^[\p{Sc}\p{L}]+\.?[ \t]*/u, '').replace(/[ \t]*[\p{Sc}\p{L}]+\.?$/u, '');
			if (body !== before) {
				exact = false;
			}
		}
		if (body === before) {
			break;
		}
	}
	if (!/^(?:\d|[.,]\d)[\d.,' \t ]*(?:[eEdD][+-]?\d+)?$/.test(body)) {
		return { kind: 'invalid' };
	}
	if (!exact || !/^\d+(?:[eEdD][+-]?\d+)?$/.test(body)) {
		return { kind: 'number' };
	}
	const value = Number(body.replace(/[dD]/, 'e'));
	return Number.isFinite(value) ? { kind: 'number', value: negative ? -value : value } : { kind: 'number' };
}

/** Whether no locale converts the string to a number. */
export function isInvalidNumericString(text: string): boolean {
	return numericStringVerdict(text).kind === 'invalid';
}

/** Whether no locale converts the string to a Boolean. */
export function isInvalidBooleanString(text: string): boolean {
	if (/^(?:true|false)$/i.test(text)) {
		return false;
	}
	return isInvalidNumericString(text);
}

/** Whether the Gregorian calendar has this day: 1900 is no leap year, 2000 is. */
function isCalendarDay(year: number, month: number, day: number): boolean {
	if (month < 1 || month > 12 || day < 1) {
		return false;
	}
	const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
	const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
	return day <= days;
}

/**
 * Whether a string written as hours, minutes and seconds names no time:
 * "25:00", "10:60", "10:00:60", "1:2:3:4" (issue #262, measured in Excel
 * 16.0 with CDate, DateValue and TimeValue). "13:00 PM" runs.
 */
export function isInvalidTimeString(text: string): boolean {
	const trimmed = text.replace(BLANK_EDGES, '');
	if (/^\d+(?::\d+){3,}$/.test(trimmed)) {
		return true;
	}
	const time = /^(\d+):(\d+)(?::(\d+))?$/.exec(trimmed);
	return time !== null && (Number(time[1]) > 23 || Number(time[2]) > 59 || Number(time[3] ?? 0) > 59);
}

/** Whether no locale converts the string to a Date. */
export function isInvalidDateString(text: string): boolean {
	const trimmed = text.replace(BLANK_EDGES, '');
	// A number every locale reads alike is that day's serial: "&HFF" is
	// 9/11/1900, and one past 12/31/9999 raises (issue #336, measured in
	// Excel 16.0).
	const number = numericStringVerdict(trimmed);
	if (number.kind !== 'invalid' && number.value !== undefined) {
		return number.value < DATE_SERIAL_MIN || number.value >= DATE_SERIAL_MAX + 1;
	}
	if (!/[0-9]/.test(trimmed)) {
		return true;
	}
	if (!isInvalidNumericString(trimmed)) {
		return false;
	}
	// A year first, then a month and a day in either order, that names no
	// day: "2020-02-30" (issue #239, measured in Excel 16.0).
	const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(trimmed);
	if (iso) {
		const [year, a, b] = iso.slice(1).map(Number);
		return year >= 100 && !isCalendarDay(year, a, b) && !isCalendarDay(year, b, a);
	}
	// Dates use letters (month names, AM and PM), digits and these separators.
	return /[^\p{L}\p{N} \t.,/:'-]/u.test(trimmed);
}
