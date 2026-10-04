import { expect, it } from 'vitest';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';
import { collectSymbolReferences } from '../src/vbaReferenceResolution';

function collect(caller: string, type?: string) {
    const modules = [
        { moduleName: 'Library', source: 'Public Sub Greet()\nEnd Sub\n' },
        { moduleName: 'Caller', source: caller, type },
    ];
    const project = buildVbaProjectIndex(modules);
    const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
    const at = modules[0].source.indexOf('Greet');
    return () => collectSymbolReferences(byModule, project, modules, modules[0].source,
        'Library', modules[0], 'Greet', at + 5, at, true);
}

it.each(['\n', '\r\n', '\r'])('bounds With-window token reads (%j)', eol => {
    const count = 1000;
    const caller = ['Sub UseIt()', ' With Library', ...Array(count).fill('    .Greet'), ' End With', 'End Sub', ''].join(eol);
    const run = collect(caller);
    const tokens = tokenizeCached(caller);
    let reads = 0;
    for (const token of tokens) {
        const end = token.end;
        Object.defineProperty(token, 'end', { get() { reads++; return end; } });
        Object.freeze(token);
    }
    Object.freeze(tokens);
    const result = run();
    expect(result).toEqual({ references: [
        { moduleName: 'Library', line: 0, column: 11, length: 5, kind: 'write' },
        ...Array.from({ length: count }, (_, i) => ({ moduleName: 'Caller', line: i + 2, column: 5, length: 5, kind: 'read' })),
    ], hasSymbol: true, ambiguous: [] });
    expect(reads).toBeLessThan(50 * tokens.length);
});

it('keeps procedure and module contexts independent', () => {
    const caller = 'Sub First()\n With Library\n .Greet\n End With\nEnd Sub\nSub Second()\n With Library\n .Greet\n End With\nEnd Sub\n';
    const result = collect(caller)();
    expect(result.references.map(ref => [ref.moduleName, ref.line])).toEqual([['Library', 0], ['Caller', 2], ['Caller', 7]]);
});

it('discards With facts when a later collection changes the receiver', () => {
    const source = 'Sub UseIt()\n With Library\n .Greet\n End With\nEnd Sub\n';
    const first = collect(source)();
    expect(first.references.map(ref => ref.moduleName)).toEqual(['Library', 'Caller']);
    const changed = collect(source.replace('With Library', 'With MissingReceiver'))();
    expect(changed.references.map(ref => ref.moduleName)).toEqual(['Library']);
    expect(collect(source)()).toEqual(first);
});
