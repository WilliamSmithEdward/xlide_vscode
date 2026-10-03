import { describe, expect, it } from 'vitest';
import { checkMissingReturnAssignments } from '../src/analyzer/diagnostics/rules/assignments';
import { forEachStatement } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import type { Span } from '../src/analyzer/parser/nodes';
type Hit = Parameters<Parameters<typeof checkMissingReturnAssignments>[7]>;
function run(source: string, vba7?: boolean, isInterface = false, watch = false) {
 const mod = parseModule(source), environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
 const symbols = buildModuleSymbols('M', 'class', source, { parsedModule: mod, conditionalCompilation: environment });
 const tracker = environment ? createConditionalActivityTracker(mod, environment) : undefined;
 let reads = 0, activityChecks = 0;
 const activity = tracker ? { ...tracker, isInactive(span: Span) { activityChecks++; return tracker.isInactive(span); } } : undefined;
 const watched: { node: Parameters<Parameters<typeof forEachStatement>[1]>[0]; span: Span }[] = [];
 if (watch) for (const proc of mod.members) if (proc.kind === 'Procedure') forEachStatement(proc.body, node => { const span = node.span; watched.push({ node, span }); Object.defineProperty(node, 'span', { configurable: true, enumerable: true, get() { reads++; return span; } }); });
 const hits: Hit[] = [];
 try { checkMissingReturnAssignments(source, mod, symbols, undefined, activity, 'M', isInterface ? new Set(['m']) : undefined, (...hit) => { hits.push(hit); }); }
 finally { for (const { node, span } of watched) Object.defineProperty(node, 'span', { configurable: true, enumerable: true, writable: true, value: span }); }
 return { hits, reads, activityChecks };
}
describe('missing-return raise processing', () => {
 it.each(['early-raise', 'early-error', 'late-raise', 'no-raise', 'early-return'])('bounds span reads while preserving outcomes (%s)', (mode) => {
  const source = 'Function P() As Long\n' + (mode === 'early-raise' ? 'Err.Raise 5\n' : mode === 'early-error' ? 'Error 5\n' : mode === 'early-return' ? 'P = 1\n' : '') + 'Beep\n'.repeat(1000) + (mode === 'late-raise' ? 'Err.Raise 5\n' : '') + 'End Function';
  const result = run(source, undefined, false, true);
  expect(result.hits).toHaveLength(mode === 'no-raise' ? 1 : 0);
  if (mode === 'early-raise' || mode === 'early-error') expect(result.reads).toBeLessThan(1500);
  else if (mode === 'early-return') expect(result.reads).toBeLessThan(5);
  else expect(result.reads).toBeGreaterThanOrEqual(3000);
 });
 it('ignores inactive raises and refreshes activity for reused parsed source', () => {
  const source = 'Function P() As Long\n#If VBA7 Then\nErr.Raise 5\n#Else\nBeep\n#End If\nBeep\nEnd Function';
  expect(run(source, true).hits).toEqual([]);
  expect(run(source, false).hits).toHaveLength(1);
  expect(run(source, true).hits).toEqual([]);
 });
 it('continues checking conditional activity throughout the body after a raise', () => {
  const result = run('Function P() As Long\n#If VBA7 Then\nErr.Raise 5\n#Else\nBeep\n#End If\n' + 'Beep\n'.repeat(1000) + 'End Function', true);
  expect(result.hits).toEqual([]);
  expect(result.activityChecks).toBeGreaterThanOrEqual(3000);
 });
 it.each(['Function P() As Long\n', 'Property Get P() As Long\n'])('preserves empty interface stubs and executable body distinctions (%s)', (header) => {
  const end = header.startsWith('Function') ? 'End Function' : 'End Property';
  const stub = header + "' comment\nDim n As Long\n" + end;
  expect(run(stub, undefined, true).hits).toEqual([]);
  expect(run(stub).hits).toHaveLength(1);
  expect(run(header + 'Beep\n' + end, undefined, true).hits).toHaveLength(1);
  expect(run(header + 'Error 5\nBeep\n' + end).hits).toEqual([]);
 });
 it('keeps the raise result specific to its procedure', () => {
  const source = 'Function P() As Long\nErr.Raise 5\nBeep\nEnd Function\nFunction Q() As Long\nBeep\nEnd Function';
  const hits = run(source).hits;
  expect(hits).toHaveLength(1); expect(hits[0][1]).toContain("'Q'");
 });
 it('retains existing active nested and single-line raise recognition', () => {
  expect(run('Function P() As Long\nIf False Then\nErr.Raise 5\nEnd If\nBeep\nEnd Function').hits).toEqual([]);
  expect(run('Function P() As Long\nIf True Then Beep Else Err.Raise 5\nBeep\nEnd Function').hits).toEqual([]);
 });
});
