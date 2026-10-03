import { describe, expect, it, vi } from 'vitest';
import { checkOverflow } from '../../src/analyzer/diagnostics/rules/overflow';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';

function setup(count: number) {
    const source='Private Const C = '+Array.from({length:count},()=> '1').join(' And ')+'\nPrivate Const Bad As Integer = 40000';
    const module=parseModule(source);
    const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
    return {source,module,symbols};
}

describe('numeric folding long logical chains', () => {
    it('copies token operands in proportion to expression length', () => {
        const count=400;
        const {source,module,symbols}=setup(count);
        let copied=0;
        const original=Array.prototype.slice;
        Array.prototype.slice=function(this: unknown[], start?: number, end?: number) {
            const result=original.call(this,start,end);
            const first=this[0];
            if(typeof first==='object' && first!==null && 'rawText' in first) { copied+=result.length; }
            return result;
        };
        const push=vi.fn();
        try { checkOverflow(source,module,symbols,undefined,undefined,undefined,push); }
        finally { Array.prototype.slice=original; }
        expect(push.mock.calls.filter(call=>call[0]==='constOverflow')).toHaveLength(1);
        expect(copied).toBeLessThan(count*60);
    });

    it('does not lose later findings to a logical-chain stack failure', () => {
        const {source,module,symbols}=setup(5000);
        const push=vi.fn();
        expect(()=>checkOverflow(source,module,symbols,undefined,undefined,undefined,push)).not.toThrow();
        expect(push.mock.calls.filter(call=>call[0]==='constOverflow')).toHaveLength(1);
    });

    it.each([
        ['(3E9 And 1) Or CInt(40000)', '3E9 And 1'],
        ['1 Or (3E9 And 1)', '3E9 And 1'],
        ['Missing Or CInt(40000)', undefined],
        ['9223372036854775807^ Or (3E9 And 1)', '3E9 And 1'],
    ])('preserves the first fold failure for %s', (expression, spanText) => {
        const source='Private Const C = '+expression;
        const module=parseModule(source);
        const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
        const push=vi.fn();
        checkOverflow(source,module,symbols,undefined,undefined,undefined,push);
        const found=push.mock.calls.filter(call=>call[0]==='constOverflow');
        expect(found).toHaveLength(Number(spanText!==undefined));
        if(spanText!==undefined)expect(source.slice(found[0][2].start,found[0][2].end)).toBe(spanText);
    });
});
