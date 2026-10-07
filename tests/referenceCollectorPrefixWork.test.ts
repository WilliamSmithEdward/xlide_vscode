import { expect, it } from 'vitest';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { buildVbaProjectIndex } from '../src/vbaProjectAnalysis';
import { collectSymbolReferences } from '../src/vbaReferenceResolution';

it.each([true, false])('bounds copied token work for qualified references (declaration=%s)', includeDeclaration => {
    const count = 1000;
    const modules = [
        { moduleName: 'Library', source: 'Public Sub Greet()\nEnd Sub\n' },
        { moduleName: 'Caller', source: 'Sub UseIt()\n' + Array(count).fill('    Library.Greet').join('\n') + '\nEnd Sub\n' },
    ];
    const project = buildVbaProjectIndex(modules);
    const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
    const shared = new Set(modules.flatMap(mod => tokenizeCached(mod.source)));
    for (const mod of modules) {
        const tokens = tokenizeCached(mod.source);
        for (const token of tokens) { Object.freeze(token); }
        Object.freeze(tokens);
    }
    const slice = Array.prototype.slice;
    let copied = 0;
    // Count copies of actual cached tokens, including filtered arrays that retain
    // those objects. Restore before assertions so observation cannot affect them.
    Array.prototype.slice = function(start?: number, end?: number) {
        const result = slice.call(this, start, end);
        if (this.length > 0 && shared.has(this[0])) { copied += result.length; }
        return result;
    };
    let result;
    const at = modules[0].source.indexOf('Greet');
    try {
        result = collectSymbolReferences(byModule, project, modules, modules[0].source,
            'Library', modules[0], 'Greet', at + 5, at, includeDeclaration);
    } finally { Array.prototype.slice = slice; }
    const references = [
        ...(includeDeclaration ? [{ moduleName: 'Library', line: 0, column: 11, length: 5, kind: 'write' }] : []),
        ...Array.from({ length: count }, (_, i) => ({ moduleName: 'Caller', line: i + 1, column: 12, length: 5, kind: 'read' })),
    ];
    expect(result).toEqual({ references, hasSymbol: true, ambiguous: [] });
    expect(copied).toBeLessThan(20 * shared.size);
});
