import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { checkFixedArraySubscriptBounds } from '../src/analyzer/diagnostics/rules/arrays';

const findings = (source: string, flag?: boolean) => {
    const mod = parseModule(source);
    const activity = flag === undefined ? undefined : createConditionalActivityTracker(mod, { compilerConstants: { FLAG: flag } });
    const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
    const spans: string[] = [];
    checkFixedArraySubscriptBounds(source, mod, symbols, activity, (_code, _message, span) => spans.push(source.slice(span.start, span.end)));
    return spans;
};

describe('array-return function candidates', () => {
    it('keeps direct and explicit return-array subscripts while respecting local and parameter shadows', () => {
        const source = 'Function Values() As Variant\nValues = Array(1, 2)\nEnd Function\n'
            + 'Function WithArg(i As Long) As Variant\nWithArg = Array(1, 2)\nEnd Function\n'
            + 'Sub P()\nx = Values(5)\nx = WithArg(5)\nx = WithArg(1)(5)\nEnd Sub\n'
            + 'Sub Q()\nDim Values(9) As Long\nx = Values(5)\nEnd Sub\n'
            + 'Sub R(Values() As Long)\nx = Values(5)\nEnd Sub';
        expect(findings(source)).toEqual(['5', '5']);
    });

    it('uses only the active parameterless-function declarations on each pass', () => {
        const source = '#If FLAG Then\nFunction Values() As Variant\nValues = Array(1, 2)\nEnd Function\n'
            + '#Else\nFunction Values(i As Long) As Variant\nValues = Array(1, 2)\nEnd Function\n'
            + '#End If\nSub P()\nx = Values(5)\nEnd Sub';
        expect(findings(source, true)).toEqual(['5']);
        expect(findings(source, false)).toEqual([]);
        expect(findings(source, true)).toEqual(['5']);
    });
});
