import { describe, expect, it, vi } from 'vitest';
import { checkOverflow } from '../../src/analyzer/diagnostics/rules/overflow';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';

describe('overflow folding depth recovery', () => {
    it.each([
        ['parentheses', '('.repeat(3000)+'1'+')'.repeat(3000)],
        ['conversions', 'CInt('.repeat(1500)+'1'+')'.repeat(1500)],
        ['intrinsics', 'Sgn('.repeat(1500)+'1'+')'.repeat(1500)],
        ['Not', 'Not '.repeat(3000)+'1'],
    ])('keeps later overflow findings after excessive %s nesting', (_name, expression) => {
        const source='Private Const Deep = '+expression+'\nPrivate Const Bad As Integer = 40000\nSub Main()\nDim i As Integer\ni = 40000\nEnd Sub';
        const module=parseModule(source);
        const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
        const push=vi.fn();
        expect(()=>checkOverflow(source,module,symbols,undefined,undefined,undefined,push)).not.toThrow();
        expect(push.mock.calls.filter(call=>call[0]==='constOverflow')).toHaveLength(1);
        expect(push.mock.calls.filter(call=>call[0]==='arithmeticOverflow')).toHaveLength(1);
    });

    it('keeps checking after deep unknown calls in a procedure statement', () => {
        const expression='F('.repeat(1500)+'1'+')'.repeat(1500);
        const source='Sub Main()\nDim i As Integer\nFoo '+expression+'\ni = 40000\nEnd Sub';
        const module=parseModule(source);
        const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
        const push=vi.fn();
        expect(()=>checkOverflow(source,module,symbols,undefined,undefined,undefined,push)).not.toThrow();
        expect(push.mock.calls.filter(call=>call[0]==='arithmeticOverflow')).toHaveLength(1);
    });

    it('still folds ordinary nested expressions and reports their original spans', () => {
        const expression='CInt('.repeat(20)+'32767 + 1'+')'.repeat(20);
        const source='Private Const Deep = '+expression;
        const module=parseModule(source);
        const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
        const push=vi.fn();
        checkOverflow(source,module,symbols,undefined,undefined,undefined,push);
        expect(push).toHaveBeenCalledTimes(1);
        expect(push.mock.calls[0][0]).toBe('constOverflow');
        const span=push.mock.calls[0][2];
        expect(source.slice(span.start,span.end)).toBe('32767 + 1');
    });
});
