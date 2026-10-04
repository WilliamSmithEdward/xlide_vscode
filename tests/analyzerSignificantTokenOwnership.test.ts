import { describe, expect, it, vi } from 'vitest';
vi.mock('../src/analyzer/lexer/tokenHelpers', async (importOriginal) => {
 const actual = await importOriginal<typeof import('../src/analyzer/lexer/tokenHelpers')>();
 return { ...actual, statementTokensCached(source: string, span: { start: number; end: number }) {
  const tokens = actual.statementTokensCached(source, span);
  for (const token of tokens) Object.freeze(token);
  Object.freeze(tokens);
  return tokens;
 } };
});
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
const cases = [
 ['labels and continuations', 'Sub P()\nDim n As Long\nL: n = _\n 40000 \' tail\n10 n = n + 1\nEnd Sub'],
 ['branches and file numbers', 'Sub P(ByVal flag As Boolean)\nIf flag Then Print #1,"a" Else Print #2,"b" \' comment\nEnd Sub'],
 ['array erase through a callee', 'Sub Free(p() As Long)\n10 Erase p \' tail\nEnd Sub\nSub P()\nDim a() As Long\nReDim a(2)\nFree a\nDebug.Print UBound(a)\nEnd Sub'],
 ['conditional declarations', '#If VBA7 Then\nDeclare PtrSafe Function f Lib "" () As Long\n#Else\nDeclare Function f Lib "" () As Long\n#End If\nSub P()\nDebug.Print f()\nEnd Sub'],
 ['known function results', 'Function F() As String\nF="abc"\nEnd Function\nSub P()\nDim n As Long\nn=F()\nEnd Sub'],
 ['held classes and collections', 'Sub P()\nDim c As Collection\nDim o As Object\nSet c=New Collection\nSet o=New Collection\nc.Add 1\nSet o=c(1)\nEnd Sub'],
 ['loop and branch facts', 'Sub P()\nDim a(2) As Long\nDim i As Long\nFor i=0 To 3\na(i)=i\nNext\nDo While i<5\ni=i+1\nLoop\nSelect Case i\nCase 5\ni=6\nCase Else\ni=7\nEnd Select\nEnd Sub'],
 ['late-bound Else state', 'Sub P(ByVal flag As Boolean)\nDim d As Object\nIf flag Then Set d=CreateObject("Scripting.Dictionary") Else Set d=CreateObject("Scripting.Dictionary")\nd.Add "k",1\nDebug.Print d.Item("k")\nEnd Sub'],
];
describe('analyzer significant token ownership', () => {
 it.each(cases)('reads frozen cached arrays and tokens: %s', (_, source) => {
  for (const host of ['excel', 'word', 'powerpoint', 'access']) {
   const errors: unknown[] = [];
   expect(() => analyzeModule(source, { host, onInternalError: (error) => { errors.push(error); } })).not.toThrow();
   expect(errors).toEqual([]);
  }
 });
 it('retains Else stripping in file-number diagnostics', () => {
  const errors: unknown[] = [];
  const diagnostics = analyzeModule('Sub P(ByVal flag As Boolean)\nIf flag Then Print #1,"a" Else Print #2,"b"\nEnd Sub', { projectOpenedFileNumbers: { any: false, numbers: new Set() }, onInternalError: (error) => { errors.push(error); } });
  expect(diagnostics.filter(d => d.code === 'file-number-zero')).toHaveLength(2);
  expect(errors).toEqual([]);
 });
});
