import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { checkOverflow } from '../src/analyzer/diagnostics/rules/overflow';

const findings = (source: string) => {
    const mod = parseModule(source);
    const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
    const out: { code: string; text: string }[] = [];
    checkOverflow(source, mod, symbols, undefined, undefined, undefined, (code, _message, span) => out.push({ code, text: source.slice(span.start, span.end) }));
    return out;
};

describe('overflow constant scopes', () => {
    it('retains module constants while parameters, locals and local constants shadow them', () => {
        const source = 'Const K As Long = 2147483647\n'
            + 'Sub P()\nDebug.Print K + 1\nEnd Sub\n'
            + 'Sub Q(K As Double)\nDebug.Print K + 1\nEnd Sub\n'
            + 'Sub R()\nDim K As Double\nK = 1\nDebug.Print K + 1\nEnd Sub\n'
            + 'Sub S()\nConst K As Long = 1\nDebug.Print K + 1\nEnd Sub\n'
            + 'Sub T()\nConst K = UnknownValue\nDebug.Print K + 1\nEnd Sub';
        expect(findings(source)).toEqual([{ code: 'arithmeticOverflow', text: 'K + 1' }]);
    });

    it('folds local constants using module constants without hiding them in later procedures', () => {
        const source = 'Const K As Long = 2147483647\n'
            + 'Sub P()\nConst NextK = K + 1\nEnd Sub\n'
            + 'Sub Q()\nDim K As Long\nK = 1\nEnd Sub\n'
            + 'Sub R()\nDebug.Print K + 1\nEnd Sub';
        expect(findings(source)).toEqual([
            { code: 'constOverflow', text: 'K + 1' },
            { code: 'arithmeticOverflow', text: 'K + 1' },
        ]);
    });
});
