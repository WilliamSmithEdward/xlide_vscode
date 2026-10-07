import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { checkModuleMemberForms } from '../src/analyzer/diagnostics/rules/moduleMembers';
import { walkProcedureStatements } from '../src/analyzer/diagnostics/walker';

describe('module member candidate lookup', () => {
    it('indexes visible declarations once instead of scanning them for every ordinary assignment', () => {
        const source = 'Option Explicit\n'
            + Array.from({ length: 500 }, (_, i) => `Public Const K${i} As Long = ${i}\n`).join('')
            + 'Sub P()\nDim x As Long\n'
            + Array.from({ length: 500 }, () => 'x = 1\n').join('') + 'End Sub\n';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        let visibleNameReads = 0;
        const visible = symbols.root.children!.filter(symbol => symbol.kind === 'constant').map(symbol => {
            const external = { ...symbol, moduleName: 'Other' };
            Object.defineProperty(external, 'name', { get: () => { visibleNameReads++; return symbol.name; } });
            return external;
        });
        const findings: string[] = [];
        const visitor = checkModuleMemberForms(source, symbols, {}, visible, true, code => findings.push(code));
        walkProcedureStatements(mod, undefined, [visitor]);
        expect(findings).toEqual([]);
        expect(visibleNameReads).toBeLessThanOrEqual(visible.length * 3);
    });
});
