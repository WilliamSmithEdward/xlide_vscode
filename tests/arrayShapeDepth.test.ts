import { describe, expect, it } from 'vitest';
import { arrayValueShape } from '../src/analyzer/diagnostics/rules/arrays';
import { rawExpressionTokens } from '../src/analyzer/diagnostics/walker';
import { MAX_EXPRESSION_DEPTH } from '../src/analyzer/parser/expressionLimits';
function shape(expression: string) { return arrayValueShape(rawExpressionTokens(expression), 'a', 1, undefined, 'binary'); }
function nestedArray(depth: number) { return 'Array('.repeat(depth) + '1' + ')'.repeat(depth); }
function levels(value: ReturnType<typeof shape>) { let count = 0; while (value) { count++; value = value.elements?.[0]; } return count; }
describe('array-shape expression depth', () => {
 it('preserves nested bounds and values within the shared limit', () => {
  const result = shape(nestedArray(50));
  expect(levels(result)).toBe(50);
  let leaf = result;
  for (let i = 1; i < 50; i++) leaf = leaf?.elements?.[0];
  expect(leaf?.values).toEqual([1]);
  expect(leaf?.dims).toEqual([{ lower: 1, upper: 1, explicitLower: true }]);
 });
 it('keeps outer bounds and stops collecting elements at the shared limit', () => {
  const result = shape(nestedArray(MAX_EXPRESSION_DEPTH + 1));
  expect(levels(result)).toBe(MAX_EXPRESSION_DEPTH);
  expect(result?.dims[0].upper).toBe(1);
 });
 it('does not throw for thousands of nested arrays', () => {
  expect(levels(shape(nestedArray(5000)))).toBe(MAX_EXPRESSION_DEPTH);
 });
 it('applies the same budget through nested Filter inputs', () => {
  const nested = (depth: number) => 'Filter('.repeat(depth) + 'Array("a", "b")' + ', "a")'.repeat(depth);
  expect(shape(nested(20))?.values).toEqual(['a']);
  expect(shape(nested(2000))).toBeUndefined();
 });
 it('shares the budget across alternating Array and Filter calls', () => {
  let expression = 'Array("a")';
  for (let i = 0; i < MAX_EXPRESSION_DEPTH; i++) expression = i % 2 ? `Filter(${expression}, "a")` : `Array(${expression})`;
  expect(() => shape(expression)).not.toThrow();
 });
});
