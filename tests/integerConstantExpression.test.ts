import { describe, expect, it } from 'vitest';
import {
	enumMemberRawExpression,
	evaluateIntegerConstantExpression,
	parseDecimalIntegerLiteral,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
	safeInteger,
} from '../src/analyzer/constants/integerConstantExpression';

const noConstants = new Map<string, number | undefined>();

describe('parseVbaIntegerLiteral', () => {
	it('parses decimal, hex, and octal literals with optional type suffixes', () => {
		expect(parseVbaIntegerLiteral('42')).toBe(42);
		expect(parseVbaIntegerLiteral('42%')).toBe(42);
		expect(parseVbaIntegerLiteral('42&')).toBe(42);
		expect(parseVbaIntegerLiteral('42^')).toBe(42);
		expect(parseVbaIntegerLiteral('&HFF')).toBe(255);
		expect(parseVbaIntegerLiteral('&hff&')).toBe(255);
		expect(parseVbaIntegerLiteral('&O17')).toBe(15);
		expect(parseVbaIntegerLiteral('&o17%')).toBe(15);
	});

	it('rejects non-integer text and unsafe magnitudes', () => {
		expect(parseVbaIntegerLiteral('1.5')).toBeUndefined();
		expect(parseVbaIntegerLiteral('-1')).toBeUndefined();
		expect(parseVbaIntegerLiteral('abc')).toBeUndefined();
		expect(parseVbaIntegerLiteral('')).toBeUndefined();
		expect(parseVbaIntegerLiteral('9007199254740993')).toBeUndefined();
	});
});

describe('parseDecimalIntegerLiteral / safeInteger', () => {
	it('accepts plain unsigned decimals only', () => {
		expect(parseDecimalIntegerLiteral('7')).toBe(7);
		expect(parseDecimalIntegerLiteral('7%')).toBeUndefined();
		expect(parseDecimalIntegerLiteral('&H7')).toBeUndefined();
	});

	it('clamps to safe integers', () => {
		expect(safeInteger(3)).toBe(3);
		expect(safeInteger(Number.MAX_SAFE_INTEGER + 1)).toBeUndefined();
	});
});

describe('enumMemberRawExpression', () => {
	it('uses the explicit raw value when present', () => {
		expect(enumMemberRawExpression('5', 'Prev')).toBe('5');
	});

	it('falls back to previous member + 1, then 0', () => {
		expect(enumMemberRawExpression(undefined, 'Prev')).toBe('Prev + 1');
		expect(enumMemberRawExpression(undefined, undefined)).toBe('0');
	});
});

describe('evaluateIntegerConstantExpression', () => {
	it('evaluates +, -, * with precedence, parens, and unary signs', () => {
		expect(evaluateIntegerConstantExpression('1 + 2 * 3', noConstants)).toBe(7);
		expect(evaluateIntegerConstantExpression('(1 + 2) * 3', noConstants)).toBe(9);
		expect(evaluateIntegerConstantExpression('-4 + +6', noConstants)).toBe(2);
		expect(evaluateIntegerConstantExpression('--5', noConstants)).toBe(5);
		expect(evaluateIntegerConstantExpression('10 - 2 - 3', noConstants)).toBe(5);
	});

	it('resolves bare and Module.Constant qualified names case-insensitively', () => {
		const constants = new Map<string, number | undefined>([
			['width', 10],
			['globals.height', 4],
		]);
		expect(evaluateIntegerConstantExpression('WIDTH + 1', constants)).toBe(11);
		expect(evaluateIntegerConstantExpression('Globals.Height * 2', constants)).toBe(8);
		expect(evaluateIntegerConstantExpression('[Width] - 1', constants)).toBe(9);
	});

	it('returns undefined for unknown names, partial parses, and non-integers', () => {
		expect(evaluateIntegerConstantExpression('missing', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('1 +', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('1 2', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('(1', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('1.5', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('"text"', noConstants)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('', noConstants)).toBeUndefined();
	});

	it('propagates overflow to undefined instead of clamping', () => {
		expect(
			evaluateIntegerConstantExpression('9007199254740991 + 1', noConstants),
		).toBeUndefined();
	});
});

describe('resolveRawIntegerConstants', () => {
	it('resolves chained constant references', () => {
		const resolved = resolveRawIntegerConstants(
			new Map<string, string | undefined>([
				['base', '2'],
				['derived', 'base * 3'],
				['further', 'derived + 1'],
			]),
		);
		expect(resolved.get('base')).toBe(2);
		expect(resolved.get('derived')).toBe(6);
		expect(resolved.get('further')).toBe(7);
	});

	it('treats ambiguous duplicates (undefined raw) and cycles as unknown', () => {
		const resolved = resolveRawIntegerConstants(
			new Map<string, string | undefined>([
				['dup', undefined],
				['a', 'b + 1'],
				['b', 'a + 1'],
				['uses', 'dup + 1'],
			]),
		);
		expect(resolved.get('dup')).toBeUndefined();
		expect(resolved.get('a')).toBeUndefined();
		expect(resolved.get('b')).toBeUndefined();
		expect(resolved.get('uses')).toBeUndefined();
	});

	it('falls back to the base map for names it does not own (diagnostics call site)', () => {
		const base = new Map<string, number | undefined>([['project_max', 100]]);
		const resolved = resolveRawIntegerConstants(
			new Map<string, string | undefined>([['local_max', 'project_max - 1']]),
			base,
		);
		expect(resolved.get('local_max')).toBe(99);
		// Only the raw-constant keys are materialized in the result.
		expect(resolved.has('project_max')).toBe(false);
	});

	it('treats unknown names as unresolved without a base map (project call site)', () => {
		const resolved = resolveRawIntegerConstants(
			new Map<string, string | undefined>([['uses_hidden', 'hidden + 1']]),
		);
		expect(resolved.get('uses_hidden')).toBeUndefined();
	});
});
