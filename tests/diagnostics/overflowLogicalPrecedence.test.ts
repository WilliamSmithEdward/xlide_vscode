import { describe, expect, it, vi } from 'vitest';
import { checkOverflow } from '../../src/analyzer/diagnostics/rules/overflow';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../../src/analyzer/symbols/buildModuleSymbols';

describe('numeric folding logical precedence', () => {
    it.each([
        ['1 Or 256 And 0', false],
        ['256 Or 1 Xor 256', false],
        ['0 Imp 0 And 1', true],
        ['(0 Imp 0) And 1', false],
        ['0 Imp 0 Imp 0', false],
        ['0 Imp (0 Imp 0)', true],
        ['Not 255 + 256', true],
        ['(Not 255) + 256', false],
    ])('preserves the Byte Const verdict for %s', (expression, overflow) => {
        const source='Private Const C As Byte = '+expression;
        const module=parseModule(source);
        const symbols=buildModuleSymbols('Module','standard',source,{parsedModule:module});
        const push=vi.fn();
        checkOverflow(source,module,symbols,undefined,undefined,undefined,push);
        expect(push.mock.calls.filter(call=>call[0]==='constOverflow')).toHaveLength(Number(overflow));
    });
});
