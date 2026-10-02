// What a worksheet function refuses when its arguments are literals (issue
// #442). Through WorksheetFunction a worksheet error value is raised as
// Run-time error 1004; through Application the same call returns the error
// value and runs, so only WorksheetFunction is judged. Each case measured in
// Excel 16.0 (build 20430, 2026-10-02):
//
//  - Ln(0), Ln(-1) and Log10(0): no logarithm of a number at or below 0.
//  - Power(-1, 0.5): a negative base to a fractional power; Power(0, -1).
//  - Sum("abc"), Max("abc"), Average("x", 1): a string that is no number.
//    Sum("5") is 5.
//  - Dec2Bin(1000) and Dec2Bin(-513): outside -512 to 511.
//  - Large(Array(1, 2), 3) and Small(Array(1, 2), 0): k outside 1 to Count.
//  - Index(Array(1, 2), 3): past the last element.
//  - Match("zzz", Array("a", "b"), 0): an exact match finds nothing. Text
//    compares without case: Match("B", ...) is 2.

import { parseVbaIntegerLiteral } from '../../constants/integerConstantExpression';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import { stringLiteralValue } from '../typeInference';
import { matchParenFrom, tokenText } from '../walker';

type Literal = number | string;

/** A number or string literal, a number optionally signed. */
function literal(arg: readonly VbaToken[] | undefined): Literal | undefined {
	if (!arg) {
		return undefined;
	}
	const toks = arg.filter((tok) => tok.kind !== 'comment');
	if (toks.length === 1 && toks[0].kind === 'stringLiteral') {
		return stringLiteralValue(toks[0].rawText);
	}
	const sign = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+') ? (toks[0].rawText === '-' ? -1 : 1) : undefined;
	const number = sign === undefined ? toks[0] : toks[1];
	if (toks.length !== (sign === undefined ? 1 : 2) || !number) {
		return undefined;
	}
	const value = number.kind === 'integerLiteral' ? parseVbaIntegerLiteral(number.rawText)
		: number.kind === 'floatLiteral' ? Number(number.rawText.replace(/[!#@]$/, '')) : undefined;
	return value === undefined || !Number.isFinite(value) ? undefined : (sign ?? 1) * value;
}

/** The literals of `Array(...)`, each a number or string, or undefined. */
function arrayLiteral(arg: readonly VbaToken[] | undefined): Literal[] | undefined {
	const toks = arg?.filter((tok) => tok.kind !== 'comment');
	if (!toks || tokenText(toks[0]) !== 'array' || toks[1]?.rawText !== '(' || matchParenFrom(toks, 1) !== toks.length - 1) {
		return undefined;
	}
	const items = toks.length === 3 ? [] : splitTopLevelTokenGroups(toks, 2, ',', toks.length - 1).map(literal);
	return items.every((item) => item !== undefined) ? items as Literal[] : undefined;
}

function isNumericText(text: string): boolean {
	return /^\s*[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?\s*$/i.test(text);
}

/**
 * Why the worksheet function refuses these literal arguments, or undefined.
 * `name` is the function's name, lowercased.
 */
export function worksheetFunctionRefusal(name: string, args: readonly (readonly VbaToken[])[]): string | undefined {
	const values = args.map(literal);
	const [a, b] = values;
	switch (name) {
		case 'ln':
		case 'log10':
			return typeof a === 'number' && a <= 0 ? `${name === 'ln' ? 'Ln' : 'Log10'} has no value at ${a}: a logarithm takes a number above 0` : undefined;
		case 'power':
			if (typeof a === 'number' && typeof b === 'number') {
				if (a < 0 && !Number.isInteger(b)) {
					return `Power(${a}, ${b}) raises a negative number to a fractional power`;
				}
				if (a === 0 && b < 0) {
					return `Power(0, ${b}) divides by zero`;
				}
			}
			return undefined;
		case 'sum':
		case 'max':
		case 'min':
		case 'average':
		case 'product': {
			const text = values.find((value): value is string => typeof value === 'string' && !isNumericText(value));
			return text !== undefined ? `${JSON.stringify(text)} is no number for the worksheet function to take` : undefined;
		}
		case 'dec2bin':
			return typeof a === 'number' && Number.isInteger(a) && (a > 511 || a < -512) ? `Dec2Bin takes -512 to 511, and ${a} is outside that` : undefined;
		case 'large':
		case 'small': {
			const items = arrayLiteral(args[0]);
			return items && typeof b === 'number' && Number.isInteger(b) && (b < 1 || b > items.length)
				? `the array holds ${items.length} value${items.length === 1 ? '' : 's'}, so k = ${b} names none`
				: undefined;
		}
		case 'index': {
			const items = arrayLiteral(args[0]);
			return items && args.length === 2 && typeof b === 'number' && Number.isInteger(b) && b > items.length
				? `the array holds ${items.length} value${items.length === 1 ? '' : 's'}, so index ${b} is past the end`
				: undefined;
		}
		case 'match': {
			const items = arrayLiteral(args[1]);
			const exact = values[2] === 0;
			if (!items || !exact || a === undefined || items.some((item) => typeof item !== typeof a)) {
				return undefined;
			}
			const same = (item: Literal): boolean => (typeof a === 'string' && typeof item === 'string' ? a.toLowerCase() === item.toLowerCase() : a === item);
			return items.some(same) ? undefined : `an exact Match finds ${JSON.stringify(a)} nowhere in the array`;
		}
	}
	return undefined;
}
