import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { variableSymbolIn } from '../src/analyzer/diagnostics/typeFields';

describe('variable symbol scopes', () => {
    it('resolves module variables and local/parameter shadows while a constant hides the module value', () => {
        const source = 'Dim n As Long\nSub P()\nDim n As String\nEnd Sub\nSub Q(ByVal n As Byte)\nEnd Sub\nSub R()\nConst n = 1\nEnd Sub\nSub S()\nEnd Sub';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const procs = mod.members.filter(member => member.kind === 'Procedure');
        const expected = ['String', 'Byte', undefined, 'Long'];
        for (const [index, proc] of procs.entries()) {
            expect(variableSymbolIn(symbols, proc, 'n')?.asType).toBe(expected[index]);
            expect(variableSymbolIn(symbols, proc, 'absent')).toBeUndefined();
        }
        const changed = { ...symbols, root: { ...symbols.root, children: symbols.root.children?.map(symbol =>
            symbol.kind === 'moduleVariable' ? { ...symbol, asType: 'Double' } : symbol) } };
        expect(variableSymbolIn(changed, procs[3], 'n')?.asType).toBe('Double');
        expect(variableSymbolIn(symbols, procs[3], 'n')?.asType).toBe('Long');
    });

    it('indexes the module once rather than rescanning declarations for every procedure', () => {
        const source = 'Dim n As Long\n'
            + Array.from({ length: 200 }, (_, i) => `Const K${i} = ${i}\n`).join('')
            + Array.from({ length: 200 }, (_, i) => `Sub P${i}()\nEnd Sub\n`).join('');
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        let kindReads = 0;
        const counted = { ...symbols, root: { ...symbols.root, children: symbols.root.children?.map(symbol => {
            const copy = { ...symbol };
            Object.defineProperty(copy, 'kind', { get: () => { kindReads++; return symbol.kind; } });
            return copy;
        }) } };
        for (const proc of mod.members.filter(member => member.kind === 'Procedure')) {
            expect(variableSymbolIn(counted, proc, 'n')?.asType).toBe('Long');
        }
        expect(kindReads).toBeLessThanOrEqual(counted.root.children!.length * 2);
    });
});
