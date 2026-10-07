import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { defaultedStraightLine, stringConstantsInScope } from '../src/analyzer/diagnostics/typeInference';

describe('constant values in procedure starts', () => {
    it('keeps module values separate from local constants, parameters and local variables', () => {
        const source = 'Const K = 99\nConst Flag = True\nConst Text = "hi"\n'
            + 'Sub P()\nDebug.Print K\nEnd Sub\n'
            + 'Sub Q()\nConst K = 7\nDebug.Print K\nEnd Sub\n'
            + 'Sub R(ByVal K As Long)\nDebug.Print K\nEnd Sub\n'
            + 'Sub S()\nDim K As Long\nDebug.Print K\nEnd Sub\n'
            + 'Sub T()\nConst K = Unknown + 1\nDebug.Print K\nEnd Sub';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const procs = mod.members.filter(member => member.kind === 'Procedure');
        for (const [index, expected] of ['99', '7', undefined, '0', undefined].entries()) {
            const proc = procs[index];
            const facts = defaultedStraightLine(source, proc, symbols, undefined).get(proc.body.at(-1)!);
            expect(facts?.get('k')?.map(token => token.rawText).join('')).toBe(expected);
            expect(facts?.get('flag')?.map(token => token.rawText).join('')).toBe('-1');
            expect(facts?.get('text')?.map(token => token.rawText).join('')).toBe('"hi"');
            expect(facts?.has('k')).toBe(expected !== undefined);
            expect(facts?.size).toBe(expected === undefined ? 2 : 3);
            expect(new Map(facts).size).toBe(facts?.size);
        }
    });

    it('keeps string constant scopes mutable without leaking changes into other procedures or snapshots', () => {
        const source = 'Const K = "module"\nSub P()\nEnd Sub\nSub Q()\nConst K = "local"\nEnd Sub\nSub R(ByVal K As String)\nEnd Sub\nSub S()\nConst K = Unknown\nEnd Sub';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const procs = mod.members.filter(member => member.kind === 'Procedure');
        const changed = { ...symbols, root: { ...symbols.root, children: symbols.root.children?.map(symbol =>
            symbol.kind === 'constant' ? { ...symbol, defaultRaw: '"changed"' } : symbol) } };
        const p = stringConstantsInScope(symbols, procs[0]);
        expect(p.get('k')).toBe('module');
        p.set('k', 'caller mutation');
        expect(stringConstantsInScope(symbols, procs[0]).get('k')).toBe('module');
        expect(stringConstantsInScope(symbols, procs[1]).get('k')).toBe('local');
        expect(stringConstantsInScope(symbols, procs[2]).has('k')).toBe(false);
        expect(stringConstantsInScope(symbols, procs[3]).has('k')).toBe(false);
        expect(stringConstantsInScope(changed, procs[0]).get('k')).toBe('changed');
        expect(stringConstantsInScope(symbols, procs[0]).get('k')).toBe('module');
    });
});
