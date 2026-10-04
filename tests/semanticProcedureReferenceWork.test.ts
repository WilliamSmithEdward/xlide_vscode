import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import * as helpers from '../src/analyzer/lexer/tokenHelpers';
import { collectTypeNameReferences, resolveTypeSemanticTokens } from '../src/analyzer/semantic/typeSemanticTokens';
afterEach(() => vi.restoreAllMocks());

describe('type references in unchanged procedure nodes', () => {
  it.each(['\n', '\r\n', '\r'])('does not rewalk retained procedures after a body edit (%j)', eol => {
    const source = Array.from({length:100}, (_,i) => ['Sub Ref'+i+'()', 'Dim item As Collection', ...Array.from({length:10},()=> 'Set item = New Collection'), 'End Sub'].join(eol)).join(eol) + eol + ['Sub Editing()', 'Dim n As Long', 'n = 1', 'End Sub'].join(eol);
    const original = parseModule(source);
    const expected = collectTypeNameReferences(source);
    const scan = vi.spyOn(helpers, 'statementTokensCached');
    for (let i=2;i<12;i++) {
      const changed = source.replace('n = 1', 'n = '+i);
      const parsed = parseModule(changed);
      expect(parsed.members[0]).toBe(original.members[0]);
      expect(collectTypeNameReferences(changed)).toEqual(expected);
    }
    expect(scan.mock.calls.length).toBeLessThanOrEqual(30);
  });
  it('refreshes an edited procedure, rebased suffix spans and project type context', () => {
    const source = "' padding\n".repeat(500) + 'Sub First()\nDim first As Alpha\nSet first = New Alpha\nEnd Sub\nSub Last()\nDim last As Beta\nEnd Sub\n';
    collectTypeNameReferences(source);
    for (const type of ['Gamma', 'Alpha', 'Gamma']) {
      const changed = source.replaceAll('Alpha', type);
      const refs = collectTypeNameReferences(changed);
      expect(refs.map(r=>[r.name,changed.slice(r.span.start,r.span.end)])).toEqual([[type,type],[type,type],['Beta','Beta']]);
      for (const kind of ['class','enum','class'] as const) {
        expect(resolveTypeSemanticTokens(changed,{projectTypes:[{name:type,kind},{name:'Beta',kind:'userType'}]}).map(t=>t.tokenType)).toEqual([kind,kind,'struct']);
      }
    }
  });
});
