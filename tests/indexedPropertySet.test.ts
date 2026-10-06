import { describe, expect, it } from 'vitest';
import { analyzeProjectModule } from './diagnostics/helpers';
const setter = 'Public Property Set Item(ByVal index As Long, ByVal value As Worksheet)\nEnd Property';
function codes(statement: string, member = setter) {
 const source = `Option Explicit\nSub T(ByVal holder As Widget, ByVal ws As Worksheet)\n${statement}\nEnd Sub`;
 return analyzeProjectModule(source, [{ moduleName: 'Widget', moduleKind: 'class', source: member }], 'Caller').filter(d => d.severity === 'error').map(d => d.code);
}
describe('indexed Property Set values', () => {
 for (const wrap of [
  (rhs: string) => `Set holder.Item(1) = ${rhs}`,
  (rhs: string) => `With holder\nSet .Item(1) = ${rhs}\nEnd With`,
  (rhs: string) => `If True Then Set holder.Item(1) = ${rhs}`,
 ]) {
  it('accepts Worksheet and Nothing', () => { for (const rhs of ['ws', 'Nothing']) expect(codes(wrap(rhs)), wrap(rhs)).toEqual([]); });
  it('reports incompatible objects once', () => { expect(codes(wrap('New Collection'))).toEqual(['assignment-object-type-mismatch']); });
  it('reports a scalar as Object required once', () => { expect(codes(wrap('5'))).toEqual(['set-requires-object']); });
 }
 it('retains the Set-only Let diagnostic', () => { expect(codes('holder.Item(1) = 5')).toEqual(['invalid-property-use']); });
});
