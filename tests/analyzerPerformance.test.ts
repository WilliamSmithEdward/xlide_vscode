import { describe, expect, it, vi } from 'vitest';
import { firstTokenAtOrAfter } from '../src/analyzer/lexer/tokenHelpers';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { isNumericType, knownLocalLiteralValuesAt } from '../src/analyzer/diagnostics/typeInference';
import { classMemberValues } from '../src/analyzer/symbols/classMemberFacts';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import type { ProcedureNode } from '../src/analyzer/parser/nodes';

describe('analyzer performance invariants', () => {
    it('finds token boundaries with logarithmic reads, including gaps and past the end', () => {
        const tokens = tokenize('Option Explicit\nSub Test()\n' + 'x = 1\n'.repeat(4000) + 'End Sub');
        let reads = 0;
        const indexed = new Proxy(tokens, { get(target, key, receiver) {
            if (typeof key === 'string' && /^\d+$/.test(key)) reads++;
            return Reflect.get(target, key, receiver);
        } });
        for (const offset of [-1, 0, 4, 15, tokens[tokens.length - 1].start, tokens[tokens.length - 1].end + 1]) {
            reads = 0;
            const expected = tokens.findIndex(token => token.start >= offset);
            expect(firstTokenAtOrAfter(indexed, offset)).toBe(expected < 0 ? tokens.length : expected);
            expect(reads).toBeLessThanOrEqual(Math.ceil(Math.log2(tokens.length + 1)));
        }
        expect(firstTokenAtOrAfter([], 10)).toBe(0);
    });

    it.each(['byte', 'integer', 'long', 'longlong', 'longptr', 'single', 'double', 'currency', 'decimal'])('keeps %s numeric', type => {
        expect(isNumericType(type)).toBe(true);
    });
    it.each(['string', 'boolean', 'date', 'object', 'variant', 'Long', 'long()'])('does not classify %s as a normalized numeric type', type => {
        expect(isNumericType(type)).toBe(false);
    });

    it('shares statement values across rules and invalidates source, symbols, and conditional activity', () => {
        const source = 'Option Explicit\n#If True Then\nSub Run()\nDim x As Long\nx = 2\nDebug.Print x\nx = 4\nDebug.Print x\nEnd Sub\n#End If';
        const mod = parseModule(source);
        const proc = mod.members.find(m => m.kind === 'Procedure') as ProcedureNode;
        const symbols = buildModuleSymbols('M', 'standard', source);
        const activity = createConditionalActivityTracker(mod);
        const at = knownLocalLiteralValuesAt(source, proc, symbols, activity);
        expect(knownLocalLiteralValuesAt(source, proc, symbols, activity)).toBe(at);
        expect(at(proc.body[2]).get('x')?.value).toBe(2);
        expect(at(proc.body[4]).get('x')?.value).toBe(4);
        expect(at(proc.body[2]).get('x')?.value).toBe(2);
        expect(knownLocalLiteralValuesAt(source, proc, symbols, createConditionalActivityTracker(mod))).not.toBe(at);
        expect(knownLocalLiteralValuesAt(source, proc, buildModuleSymbols('M', 'standard', source), activity)).not.toBe(at);
        expect(knownLocalLiteralValuesAt(source + '\n', proc, symbols, activity)).not.toBe(at);
    });

    it('keeps malformed header diagnostics at the end of a large module', () => {
        const prefix = 'Option Explicit\n' + Array.from({ length: 300 }, (_, i) => `Sub S${i}()\nEnd Sub\n`).join('');
        const source = prefix + 'Sub Broken(a b)\nEnd Sub\n';
        const hits = analyzeModule(source).filter(d => d.code === 'malformed-statement');
        expect(hits).toHaveLength(1);
        expect(source.slice(hits[0].span.start, hits[0].span.end)).toBe('b');
    });

    it('preserves class facts for field mentions, shadowed names, getters and parameterized functions', () => {
        const source = `Option Explicit
Public Untouched As Object
Public Written As Object
Public ReadElsewhere As Variant
Public Property Get Result() As Variant
Result = 3
End Property
Public Function EmptyResult() As Variant
End Function
Public Function NothingResult() As Object
Set NothingResult = Nothing
End Function
Public Function WithParam(ByVal x As Long) As Variant
WithParam = x
End Function
Public Sub Assign()
Set Written = New Collection
Debug.Print ReadElsewhere
End Sub`;
        const symbols = buildModuleSymbols('C', 'class', source);
        expect([...classMemberValues(source, symbols.root.children ?? [])]).toEqual([
            ['untouched', 'nothing'], ['result', 'scalar'], ['emptyresult', 'empty'], ['nothingresult', 'nothing'], ['withparam', 'scalar'],
        ]);
    });

    it('does not scan all class tokens for each getter or field', () => {
        const source = 'Option Explicit\n' + Array.from({ length: 400 }, (_, i) =>
            `Public Field${i} As Object\nPublic Function Get${i}() As Variant\nGet${i} = ${i}\nEnd Function\n`).join('');
        const symbols = buildModuleSymbols('C', 'class', source);
        let offsetReads = 0;
        const tokens = tokenize(source).map(token => new Proxy(token, {
            get(target, key, receiver) {
                if (key === 'start') offsetReads++;
                return Reflect.get(target, key, receiver);
            },
        }));
        const spy = vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(tokens);
        try {
            expect(classMemberValues(source, symbols.root.children ?? []).size).toBe(800);
            // The old per-getter filter reads millions of offsets here.
            expect(offsetReads).toBeLessThan(tokens.length * 5);
        } finally {
            spy.mockRestore();
        }
    });
});
