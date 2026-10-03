import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { MAX_EXPRESSION_DEPTH } from '../src/analyzer/parser/expressionLimits';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

function wrapped(depth: number, mixed = false): string {
 let text = '"bad"';
 for (let i = 0; i < depth; i++) text = mixed && i % 2 ? `StrConv(${text}, 1)` : `CStr(${text})`;
 return text;
}
function hits(expression: string): string[] {
 const source = `Sub P()\nDim n As Long\nn = ${expression}\nn = "bad"\nEnd Sub`;
 const mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const result: string[] = [];
 checkAssignmentTypes(source, mod, symbols, undefined, {}, undefined, (code) => { result.push(code); });
 return result;
}
describe('assignment text folding depth recovery', () => {
 it.each([false, true])('folds up to the shared limit and recovers beyond it (mixed=%s)', (mixed) => {
  expect(hits(wrapped(MAX_EXPRESSION_DEPTH - 1, mixed))).toEqual(['assignmentTypeMismatch', 'assignmentTypeMismatch']);
  expect(hits(wrapped(MAX_EXPRESSION_DEPTH, mixed))).toEqual(['assignmentTypeMismatch']);
 });
 it('counts nested calls rather than concatenation siblings', () => {
  expect(hits(Array.from({ length: MAX_EXPRESSION_DEPTH + 1 }, () => 'CStr("bad")').join(' & '))).toEqual(['assignmentTypeMismatch', 'assignmentTypeMismatch']);
 });
 it.each(['CStr', 'StrConv'])('continues full analysis after 5,000 nested %s calls', (fn) => {
  const expression = fn === 'CStr' ? 'CStr('.repeat(5000) + '"bad"' + ')'.repeat(5000) : 'StrConv('.repeat(5000) + '"bad"' + ', 1)'.repeat(5000);
  const source = `Sub P()\nDim n As Long\nn = ${expression}\nn = "bad"\nEnd Sub`;
  const errors: unknown[] = [];
  const diagnostics = analyzeModule(source, { onInternalError: (error, where) => { errors.push({ error: String(error), where }); } });
  expect(errors).toEqual([]);
  expect(diagnostics.filter(d => d.code === 'assignment-type-mismatch')).toHaveLength(1);
 }, 60000);
 it('keeps ordinary nested conversion, substring and case folds', () => {
  expect(hits('Left$(UCase(StrConv(Trim(CStr("bad")), 2)), 2)')).toHaveLength(2);
  expect(hits('CStr("123")')).toHaveLength(1);
 });
});
