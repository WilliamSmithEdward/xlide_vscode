import { describe, expect, it } from 'vitest';
import { conditionValue, numberValue } from '../src/analyzer/diagnostics/conditionValue';
import { rawExpressionTokens } from '../src/analyzer/diagnostics/walker';
import { MAX_EXPRESSION_DEPTH } from '../src/analyzer/parser/expressionLimits';
const facts = { value: () => undefined };
const numeric = (text: string) => numberValue(rawExpressionTokens(text), facts);
const condition = (text: string) => conditionValue(rawExpressionTokens(text), facts);
describe('condition value depth recovery', () => {
 it('preserves numeric calls at the boundary and stops beyond it', () => {
  expect(numeric('Abs('.repeat(MAX_EXPRESSION_DEPTH - 1) + '1' + ')'.repeat(MAX_EXPRESSION_DEPTH - 1))).toBe(1);
  expect(numeric('Abs('.repeat(MAX_EXPRESSION_DEPTH) + '1' + ')'.repeat(MAX_EXPRESSION_DEPTH))).toBeUndefined();
 });
 it.each(['Abs', 'Len', 'IIf'])('recovers from thousands of nested %s calls', (fn) => {
  const text = fn === 'IIf' ? 'IIf('.repeat(5000) + '1' + ', 1, 0)'.repeat(5000) : (fn + '(').repeat(5000) + '1' + ')'.repeat(5000);
  expect(numeric(text)).toBeUndefined();
 });
 it('bounds parentheses and prefix operators and shares their call budget', () => {
  expect(condition('('.repeat(5000) + 'True' + ')'.repeat(5000))).toBeUndefined();
  expect(condition('Not '.repeat(5000) + 'True')).toBeUndefined();
  expect(numeric('-'.repeat(5000) + '1')).toBeUndefined();
  expect(numeric('Abs('.repeat(128) + '('.repeat(128) + '1' + ')'.repeat(256))).toBeUndefined();
 });
 it('does not spend a sibling budget on independent operands', () => {
  expect(numeric(Array.from({ length: 500 }, () => 'Abs(1)').join(' + '))).toBe(500);
  expect(condition('Not Not (Abs(-2) = 2)')).toBe(true);
 });
});
