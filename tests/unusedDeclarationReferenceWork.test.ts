import { describe, expect, it, vi } from 'vitest';
import { checkUnusedDeclarations } from '../src/analyzer/diagnostics/rules/deadCode';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import * as references from '../src/analyzer/references/referenceKinds';

describe('unused declaration reference work', () => {
    it('classifies variable candidates only, preserving constant use and write-only findings', () => {
        const source = 'Option Explicit\nPrivate Const K = 1\nSub Run()\nDim x As Long\nx = K\n'
            + Array.from({ length: 1000 }, () => 'Debug.Print K\n').join('') + 'End Sub\n';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source);
        const spy = vi.spyOn(references, 'classifyReferenceKinds');
        const findings: string[] = [];
        try {
            checkUnusedDeclarations(source, mod, symbols, undefined, (rule) => findings.push(rule));
            expect(findings).toEqual(['variableNeverRead']);
            expect(spy).toHaveBeenCalledTimes(1);
            expect(spy.mock.calls[0][1]).toEqual([source.indexOf('x = K')]);
        } finally { spy.mockRestore(); }
    });
});
