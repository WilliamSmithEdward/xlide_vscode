import { describe, expect, it } from 'vitest';
import { analyzeProjectModule, projectOptions } from './diagnostics/helpers';
import { resolveAssignmentValueCompletion } from '../src/analyzer/completion/assignmentValueCompletion';
import { resolveMemberCompletions, resolveReceiverTypeAt } from '../src/analyzer/completion/memberAccess';
import { parseModule } from '../src/analyzer/parser/parseModule';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
const types = 'Public Type Record\nFlag As Boolean\nCount As Long\nEnd Type\nPublic Property Get Snapshot() As Record\nEnd Property';
const library = { moduleName: 'Types', moduleKind: 'standard' as const, source: types };
function source(body: string, prelude = '') { return `Option Explicit\n${prelude}\nSub T()\n${body}\nEnd Sub`; }
function findings(body: string, prelude = '', modules = [library]) { return analyzeProjectModule(source(body, prelude), modules, 'Caller'); }
function choice(text: string, modules = [library]) {
 const options = projectOptions([{ moduleName: 'Caller', source: text }, ...modules], 'Caller');
 return resolveAssignmentValueCompletion(text, text.indexOf(' = ') + 3, { projectClassMembers: options.projectClassMembers, projectSymbols: options.projectVisibleSymbols });
}
describe('exported getter receiver assignments', () => {
 for (const target of ['Snapshot.Flag', 'Types.Snapshot.Flag', 'Snapshot().Flag', 'With Snapshot\n.Flag']) {
  it(`checks the value in ${target}`, () => {
   const end = target.startsWith('With') ? '\nEnd With' : '';
   const bad = `${target} = "nonsense"${end}`;
   expect(findings(bad).filter(d => d.severity === 'error')).toEqual([expect.objectContaining({ code: 'assignment-type-mismatch' })]);
   expect(findings(`${target} = True${end}`).filter(d => d.severity === 'error')).toEqual([]);
  });
 }
 it('offers Boolean constants for the returned field', () => {
  const text = source('Snapshot.Flag = ');
  expect(choice(text)?.enumName).toBe('Boolean');
 });
 it('threads exported function result types too', () => {
  const functions = [{ ...library, source: types.replace('Property Get Snapshot()', 'Function Snapshot()').replace('End Property', 'End Function') }];
  expect(findings('Snapshot().Flag = "nonsense"', '', functions)).toContainEqual(expect.objectContaining({ code: 'assignment-type-mismatch' }));
 });
 for (const decl of ['Dim Snapshot As Object', 'Dim Snapshot', 'Const Snapshot = 1']) {
  it(`preserves local shadowing: ${decl}`, () => {
   expect(findings(`${decl}\nSnapshot.Flag = "nonsense"`).some(d => d.code === 'assignment-type-mismatch')).toBe(false);
  });
 }
 it('prefers an own procedure to the export', () => {
  expect(findings('Snapshot.Flag = "nonsense"', 'Function Snapshot() As Object\nEnd Function').some(d => d.code === 'assignment-type-mismatch')).toBe(false);
 });
 it('leaves ambiguous exports unresolved', () => {
  expect(findings('Snapshot.Flag = "nonsense"', '', [library, { ...library, moduleName: 'Other' }]).some(d => d.code === 'assignment-type-mismatch')).toBe(false);
 });
 it('does not export a private getter', () => {
  expect(findings('Snapshot.Flag = "nonsense"', '', [{ ...library, source: types.replace('Public Property Get', 'Private Property Get') }]).some(d => d.code === 'assignment-type-mismatch')).toBe(false);
 });
 it('does not export a class getter as a bare name', () => {
  expect(findings('Snapshot.Flag = "nonsense"', '', [{ ...library, moduleKind: 'class' as never }]).some(d => d.code === 'assignment-type-mismatch')).toBe(false);
 });
 it('lets an exported getter shadow a host global', () => {
  const rows = [{ ...library, source: types.replaceAll('Snapshot', 'Rows') }];
  expect(findings('Rows.Flag = "nonsense"', '', rows)).toContainEqual(expect.objectContaining({ code: 'assignment-type-mismatch' }));
  expect(findings('Rows.Flag = True', '', rows).filter(d => d.severity === 'error')).toEqual([]);
 });
});
it('indexes export names once for repeated receiver requests', () => {
 let reads = 0;
 const text = 'Sub T()\nSnapshot.Flag = True\nEnd Sub';
 const options = projectOptions([{ moduleName: 'Caller', source: text }, library], 'Caller');
 const symbol = options.projectVisibleSymbols!.find(s => s.name === 'Snapshot')!;
 const projectSymbols = Array.from({ length: 1000 }, (_, i) => ({ ...symbol, get name() { reads++; return i === 0 ? 'Snapshot' : `Getter${i}`; } }));
 const ctx = { projectSymbols, projectClassMembers: options.projectClassMembers, parsedModule: parseModule(text) };
 const offset = text.indexOf('.Flag') + 1;
 expect(resolveReceiverTypeAt(text, offset, ctx)).toBe('project:record');
 expect(reads).toBeGreaterThanOrEqual(1000);
 reads = 0;
 for (let i = 0; i < 20; i++) expect(resolveReceiverTypeAt(text, offset, ctx)).toBe('project:record');
 expect(reads).toBe(0);
});
it('keeps a setter-only export from falling through to host Rows', () => {
 const text = 'Sub T()\nRows.\nEnd Sub';
 const options = projectOptions([{ moduleName: 'Caller', source: text }, { ...library, source: 'Public Property Let Rows(ByVal value As Boolean)\nEnd Property' }], 'Caller');
 expect(resolveMemberCompletions(text, text.indexOf('Rows.') + 5, { projectSymbols: options.projectVisibleSymbols, projectClassMembers: options.projectClassMembers })).toEqual([]);
});
