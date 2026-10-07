import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { statementMayChangeModuleVariable } from '../src/analyzer/diagnostics/typeInference';

function fixture(source: string) {
    const mod = parseModule(source);
    const procedures = mod.members.filter((m): m is ProcedureNode => m.kind === 'Procedure');
    return { source, procedures, symbols: buildModuleSymbols('M', 'standard', source, { parsedModule: mod }) };
}

describe('code-running scope lookup', () => {
    it('does not walk the module symbol list for each statement', () => {
        const source = 'Option Explicit\nPrivate tracked As Long\n'
            + Array.from({ length: 500 }, (_, i) => `Private Const K${i} As Long = ${i}\n`).join('')
            + 'Sub P()\nDim x As Long\n'
            + Array.from({ length: 500 }, () => 'x = K499 + 1\n').join('') + 'End Sub\n';
        const f = fixture(source);
        let moduleReads = 0;
        const root = { ...f.symbols.root };
        Object.defineProperty(root, 'children', { get: () => { moduleReads++; return f.symbols.root.children; } });
        const symbols = { ...f.symbols, root };
        const proc = f.procedures[0];
        for (const node of proc.body.filter(node => node.kind === 'Assignment')) {
            expect(statementMayChangeModuleVariable(source, symbols, proc, node.span, 'tracked')).toBe(false);
        }
        expect(moduleReads).toBeLessThanOrEqual(3);
    });

    it('keeps ByRef and ByVal parameter scopes separate', () => {
        const f = fixture('Private tracked As Long\nSub A(ByRef value As Long)\nDebug.Print value\nEnd Sub\n'
            + 'Sub B(ByVal value As Long)\nDebug.Print value\nEnd Sub\n');
        for (const [index, expected] of [[0, true], [1, false], [0, true]] as const) {
            const proc = f.procedures[index];
            expect(statementMayChangeModuleVariable(f.source, f.symbols, proc, proc.body[0].span, 'tracked')).toBe(expected);
        }
    });

    it('does not carry safe module names into another symbol snapshot', () => {
        const f = fixture('Private tracked As Long\nSub P()\nUnknownName = 1\nEnd Sub\n');
        const proc = f.procedures[0];
        const tracked = f.symbols.root.children!.find(symbol => symbol.name === 'tracked')!;
        const changed = { ...f.symbols, root: { ...f.symbols.root, children: [
            ...f.symbols.root.children!, { ...tracked, name: 'UnknownName' },
        ] } };
        for (const [symbols, expected] of [[f.symbols, true], [changed, false], [f.symbols, true]] as const) {
            expect(statementMayChangeModuleVariable(f.source, symbols, proc, proc.body[0].span, 'tracked')).toBe(expected);
        }
    });

    it('still treats external member calls and New classes as code-running', () => {
        const f = fixture('Private tracked As Long\nSub P()\nDim x As Long, obj As Object\n'
            + 'x = Len("abc")\nobj.Run\nSet obj = New Worker\nEnd Sub\n');
        const proc = f.procedures[0];
        expect(proc.body.slice(1).map(node => statementMayChangeModuleVariable(f.source, f.symbols, proc, node.span, 'tracked')))
            .toEqual([false, true, true]);
    });
});
