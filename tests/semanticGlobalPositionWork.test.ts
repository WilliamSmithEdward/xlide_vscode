import { afterEach, describe, expect, it, vi } from 'vitest';
import * as host from '../src/analyzer/host/hostModel';
import { collectHostGlobalTokens } from '../src/analyzer/semantic/typeSemanticTokens';

afterEach(() => vi.restoreAllMocks());
describe('host global semantic position gates', () => {
  it.each(['\n', '\r\n', '\r'])('rejects member and type positions before host lookup (%j)', eol => {
    const source = ['Sub Run()', 'Dim receiver As Object', 'Dim typed As Application', 'Set receiver = New Application', ...Array.from({ length: 1000 }, () => 'receiver.Application: receiver!xlUp: receiver.UnknownMember'), 'Debug.Print Application', 'End Sub'].join(eol);
    const lookups = [vi.spyOn(host, 'resolveHostGlobal'), vi.spyOn(host, 'resolveHostGlobalMember'), vi.spyOn(host, 'resolveHostConstant')];
    const start = source.lastIndexOf('Application');
    expect(collectHostGlobalTokens(source)).toEqual([{ name: 'Application', tokenType: 'variable', span: { start, end: start + 11 }, modifiers: ['defaultLibrary'] }]);
    for (const lookup of lookups) {
      expect(lookup.mock.calls.filter(([name]) => ['Application', 'xlUp', 'UnknownMember', 'Object'].includes(name)).length).toBe(lookup === lookups[0] ? 1 : 0);
    }
  });
  it('retains bare globals, methods and constants after comment/newline filtering', () => {
    const source = 'Sub Run()\nDebug.Print Application\nUnion Range("A1"), Range("B1")\nx = xlUp\nreceiver.\n\' comment\nApplication\nEnd Sub';
    const result = collectHostGlobalTokens(source);
    expect(result.map(t => [t.name, t.tokenType])).toEqual([['Application', 'variable'], ['Union', 'function'], ['Range', 'variable'], ['Range', 'variable'], ['xlUp', 'enumMember']]);
    for (const token of result) expect(source.slice(token.span.start, token.span.end)).toBe(token.name);
  });
});
