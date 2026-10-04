import { describe, expect, it, vi } from 'vitest';
const work = vi.hoisted(() => ({ characters: 0 }));
vi.mock('../src/vbaSourceScan', async importOriginal => {
 const actual = await importOriginal<typeof import('../src/vbaSourceScan')>();
 return {...actual, stripVba: (text: string) => { work.characters += text.length; return actual.stripVba(text); }};
});
import { callSitesOf } from '../src/analyzer/refactor/callSites';
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for (const eol of ['\n', '\r\n', '\r']) {
 describe('call-site physical bounds ' + JSON.stringify(eol), () => {
  it('does not include neighboring live statements in a bare call edit', () => {
   const source = ['Option Explicit', 'Sub Caller()', 'Go 1', 'Debug.Print 2', 'End Sub', ''].join(eol);
   const sites = callSitesOf(source, 'Go');
   expect(sites).toHaveLength(1);
   expect(applyVbaTextEdits(source, sites.map(site => ({span:site.argumentInsert,newText:site.argumentText('3')}))))
    .toBe(source.replace('Go 1', 'Go 1, 3'));
  });
  it('skips a declaration without skipping later calls', () => {
   const source = ['Sub Go()', 'End Sub', 'Sub Caller()', 'Go', 'End Sub', ''].join(eol);
   const sites = callSitesOf(source, 'Go');
   expect(sites).toHaveLength(1);
   expect(sites[0].offset).toBe(source.lastIndexOf('Go'));
  });
  it('does not match a bracket across a physical line break', () => {
   const source = ['Option Explicit', 'Sub Caller()', 'Go(', 'Debug.Print 1)', 'End Sub', ''].join(eol);
   expect(callSitesOf(source, 'Go')).toEqual([]);
  });
  it('rejects a physical break even inside an unterminated string', () => {
   const source = ['Option Explicit', 'Sub Caller()', 'Go("broken', 'next")', 'End Sub', ''].join(eol);
   expect(callSitesOf(source, 'Go')).toEqual([]);
  });
  it('preserves trailing comments and qualified calls', () => {
   const source = ['Option Explicit', 'Sub Caller()', "Module1.Go 1 'note", 'Debug.Print 2', 'End Sub', ''].join(eol);
   const sites = callSitesOf(source, 'Go', {qualifier:'Module1'});
   expect(sites).toHaveLength(1);
   expect(applyVbaTextEdits(source, sites.map(site => ({span:site.argumentInsert,newText:site.argumentText('3')}))))
    .toBe(source.replace('Go 1', 'Go 1, 3'));
  });
  it('introduces a parameter at all calls without deleting following code', () => {
   const source = ['Option Explicit', 'Public Sub Report()', 'Dim limit As Long', 'limit = 3', 'Debug.Print limit', 'End Sub', 'Sub Caller()', 'Report', 'Call Report()', 'Debug.Print 2', 'End Sub', ''].join(eol);
   const result = introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1'});
   if (!result.ok) { throw new Error(result.reason); }
   expect(applyVbaTextEdits(source, result.edits)).toBe(['Option Explicit', 'Public Sub Report(ByVal limit As Long)', 'Debug.Print limit', 'End Sub', 'Sub Caller()', 'Report 3', 'Call Report(3)', 'Debug.Print 2', 'End Sub', ''].join(eol));
  });
  it('bounds stripped text by line length rather than module size', () => {
   const count = 1000;
   const source = ['Option Explicit', 'Sub Caller()', ...Array(count).fill('Go 1'), 'End Sub', ''].join(eol);
   work.characters = 0;
   const sites = callSitesOf(source, 'Go');
   expect(sites).toHaveLength(count);
   expect(work.characters).toBeLessThanOrEqual(count * 10);
  });
 });
}
