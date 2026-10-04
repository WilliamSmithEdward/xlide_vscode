import { describe, expect, it } from 'vitest';
import { evaluateIntegerConstantExpression } from '../src/analyzer/constants/integerConstantExpression';

const empty = new Map<string, number | undefined>();

describe('integer constant Not chains', () => {
	it.each([20000, 20001])('evaluates %i Not operators without overflowing the JavaScript stack', (count) => {
		expect(evaluateIntegerConstantExpression('Not '.repeat(count) + '7', empty)).toBe(count % 2 ? -8 : 7);
	});

	it('keeps arithmetic inside Not and logical operators outside it', () => {
		expect(evaluateIntegerConstantExpression('Not Not 1 + 2 * 3 And 3 Or 8 Xor 1', empty)).toBe(10);
		expect(evaluateIntegerConstantExpression('Not (Not 7 And 3)', empty)).toBe(-1);
	});

	it('reads a constant operand once even for a long chain', () => {
		const queries: string[] = [];
		expect(evaluateIntegerConstantExpression('Not '.repeat(20001) + 'Module.Value', {
			get(name) { queries.push(name); return -2147483648; },
		})).toBe(2147483647);
		expect(queries).toEqual(['module.value']);
	});

	it('rejects an out-of-Long operand even when an even number of Nots would cancel', () => {
		for (const operand of ['2147483648', '-2147483649', 'missing', '']) {
			expect(evaluateIntegerConstantExpression('Not '.repeat(20000) + operand, empty)).toBeUndefined();
		}
		expect(evaluateIntegerConstantExpression('2147483648', empty)).toBe(2147483648);
	});

	it('normalizes negative zero after an even number of Nots', () => {
		expect(evaluateIntegerConstantExpression('Not Not -0', empty)).toBe(0);
	});

	it('preserves the nesting guard for parentheses and unary signs', () => {
		expect(evaluateIntegerConstantExpression('('.repeat(1000) + 'Not 1' + ')'.repeat(1000), empty)).toBeUndefined();
		expect(evaluateIntegerConstantExpression('Not ' + '-'.repeat(1000) + '1', empty)).toBeUndefined();
	});
});
