import { describe, expect, it } from 'vitest';
import { createConditionalActivityTracker, evaluateConditionalExpression, indexConditionalCompilation } from '../src/analyzer/conditional/conditionalCompilation';
import { MAX_EXPRESSION_DEPTH } from '../src/analyzer/parser/expressionLimits';
import { parseModule } from '../src/analyzer/parser/parseModule';

describe('conditional expression stack safety', () => {
	it.each([
		['Not '.repeat(20000) + '1', 1],
		['Not '.repeat(20001) + '1', -2],
		['- '.repeat(20000) + '1', 1],
		['- '.repeat(20001) + '1', -1],
		['+ - '.repeat(10000) + '2 ^ 2', 4],
	] as const)('evaluates a flat unary chain without recursion (%#)', (expression, value) => {
		expect(evaluateConditionalExpression(expression)).toBe(value);
	});

	it.each([
		['Not Not "1.5"', 2], ['Not Not Empty', 0], ['Not Not True', true],
		['Not Not Null', { kind: 'null' }], ['Not Not Nothing', undefined],
		['Not Not Missing', undefined], ['Not 1 = 2', true], ['Not Not 1 = 2', false],
		['- + - "2"', 2], ['- + True', 1], ['- + Empty', -0],
		['- + Null', { kind: 'null' }], ['- + Nothing', undefined],
		['- + Missing', undefined], ['-2 ^ 2', -4], ['2 ^ -1', 0.5],
		['2 ^ --1', undefined], ['2 ^ 3 ^ 2', 64], ['Not', undefined],
		['- +', undefined], ['Not Not 1 2', undefined],
	] as const)('preserves coercion and precedence: %s', (expression, value) => {
		expect(evaluateConditionalExpression(expression)).toEqual(value);
	});

	it('evaluates parentheses at the shared limit and resets depth for siblings', () => {
		const expression = '('.repeat(MAX_EXPRESSION_DEPTH) + '1' + ')'.repeat(MAX_EXPRESSION_DEPTH);
		expect(evaluateConditionalExpression(expression + ' + ' + expression)).toBe(2);
	});

	it.each([MAX_EXPRESSION_DEPTH + 1, 1000, 20000])('leaves %i nested parentheses unknown', (depth) => {
		const expression = '('.repeat(depth) + '1' + ')'.repeat(depth);
		expect(evaluateConditionalExpression(expression)).toBeUndefined();
		expect(evaluateConditionalExpression(expression + ' Or True')).toBeUndefined();
	});

	it.each([
		['Not '.repeat(20001) + '1', -2, 'active'],
		['- '.repeat(20000) + '0', 0, 'inactive'],
		['('.repeat(1000) + '1' + ')'.repeat(1000), undefined, 'unknown'],
	] as const)('recovers through constant indexing and branch tracking (%#)', (expression, value, activity) => {
		const constants = indexConditionalCompilation(parseModule('#Const FLAG = ' + expression + '\n'));
		expect(constants.constants[0].value).toEqual(value);
		const source = '#If ' + expression + ' Then\nPublic Value As Long\n#End If\n';
		const start = source.indexOf('Public');
		expect(createConditionalActivityTracker(parseModule(source))?.activityForSpan({ start, end: start + 6 })).toBe(activity);
	});

	it('uses fresh compiler constants on each evaluation', () => {
		expect(evaluateConditionalExpression('Not Not FLAG', { compilerConstants: { FLAG: 1 } })).toBe(1);
		expect(evaluateConditionalExpression('Not Not FLAG', { compilerConstants: { FLAG: true } })).toBe(true);
	});
});
