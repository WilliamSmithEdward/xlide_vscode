import { describe, expect, it } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { knownLocalLiteralValuesAt } from '../src/analyzer/diagnostics/typeInference';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';

describe('statement value fact lifetime', () => {
    it('follows separate numeric and variant writes among many module constants', () => {
        const constants = Array.from({ length: 300 }, (_, i) => `Const K${i} = ${i}`).join('\n');
        const source = `${constants}\nDim a As Long, b As Variant, untouched As Long\nSub P()\nb = "hello"\na = 7\nDebug.Print a, b, untouched\nEnd Sub`;
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const reader = knownLocalLiteralValuesAt(source, proc, symbols, undefined);
        const values = reader(proc.body[2]);
        expect(values.get('a')).toMatchObject({ kind: 'number', value: 7 });
        expect(values.get('b')).toMatchObject({ kind: 'string', value: 'hello' });
        expect(values.get('untouched')).toMatchObject({ kind: 'number', value: 0 });
        expect(reader(proc.body[2])).toBe(values);
    });

    it('retains module values at each statement without sharing later writes or shadowing locals', () => {
        const source = 'Dim n As Long\nSub P()\nn = 2\nDebug.Print n\nn = 3\nDebug.Print n\nEnd Sub\nSub Q()\nDim n As Long\nDebug.Print n\nEnd Sub';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const procs = mod.members.filter(member => member.kind === 'Procedure');
        const p = procs[0];
        const reader = knownLocalLiteralValuesAt(source, p, symbols, undefined);
        const first = reader(p.body[1]);
        const later = reader(p.body[3]);
        expect(first.get('n')).toMatchObject({ value: 2 });
        expect(later.get('n')).toMatchObject({ value: 3 });
        expect(reader(p.body[1])).toBe(first);
        expect(reader(p.body[3])).toBe(later);
        const q = knownLocalLiteralValuesAt(source, procs[1], symbols, undefined);
        expect(q(procs[1].body[1]).get('n')).toMatchObject({ value: 0 });
    });

    it('recomputes the same parsed statement after conditional activity changes', () => {
        const source = 'Dim n As Long\nSub P()\n#If FLAG Then\nn = 2\n#Else\nn = 3\n#End If\nDebug.Print n\nEnd Sub';
        const mod = parseModule(source);
        const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const stmt = proc.body.at(-1)!;
        for (const flag of [true, false, true]) {
            const activity = createConditionalActivityTracker(mod, { compilerConstants: { FLAG: flag } });
            const reader = knownLocalLiteralValuesAt(source, proc, symbols, activity);
            const facts = reader(stmt);
            expect(facts.get('n')).toMatchObject({ value: flag ? 2 : 3 });
            expect(reader(stmt)).toBe(facts);
        }
    });

    it('keeps followable module declarations within their symbol snapshot', () => {
        const source = 'Dim n As Long\nSub P()\nn = 2\nDebug.Print n\nEnd Sub';
        const mod = parseModule(source);
        const proc = mod.members.find(member => member.kind === 'Procedure')!;
        const base = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
        const changed = { ...base, root: { ...base.root, children: base.root.children?.map(symbol =>
            symbol.kind === 'moduleVariable' ? { ...symbol, asType: 'Object' } : symbol) } };
        for (const [symbols, expected] of [[base, 2], [changed, undefined], [base, 2]] as const) {
            const reader = knownLocalLiteralValuesAt(source, proc, symbols, undefined);
            expect(reader(proc.body[1]).get('n')?.value).toBe(expected);
        }
    });
});
