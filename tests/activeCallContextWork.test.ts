import { afterEach, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import * as cursorContext from '../src/analyzer/completion/cursorContext';
import { findActiveCallSite, callableCompletionShouldInsertParens } from '../src/analyzer/call/callContext';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

const lineContext = cursorContext.completionLineCursorContext;
afterEach(() => vi.restoreAllMocks());

function compareLegacyPrefix(source: string, offset: number): void {
    // Supply the full-prefix stream used before this change to the same call
    // classifiers, then compare every field/decision with the bounded stream.
    const probe = vi.spyOn(cursorContext, 'completionLineCursorContext');
    probe.mockImplementation((text, caret) => cursorContext.completionCursorContext(text, caret));
    const legacySite = findActiveCallSite(source, offset);
    const legacyParens = callableCompletionShouldInsertParens(source, offset);
    probe.mockRestore();
    expect(findActiveCallSite(source, offset)).toEqual(legacySite);
    expect(callableCompletionShouldInsertParens(source, offset)).toBe(legacyParens);
}

describe('active-call logical-line work', () => {
    it.each([
        'Call SaveFile(1, Nested(2, 3), value)',
        'Debug.Print Left$("a,b", Len("text"))',
        '10 Workbooks.Open "a.xlsx", , False',
        'label: .Offset 1, Array(2, 3)',
        'Call Target(1, _\r\n    Nested(2, _\r\n        3))',
        'Call Target(1, _\n    2): Other 3, 4',
        'Call Target(1, _\r    2)',
        "Call Target(1, 2) ' comment (fake, target)",
        'Rem comment (fake, target)',
        'value = "text : (Fake, 2)"',
        '#If VBA7 Then\nCall Target(1, 2)\n#End If',
        'Call ไทย(1, _\n    名字(2, 3))',
        'Call Target(1, [select], Left$("a", 2))',
        'Call Target(1, ): value = Lef',
        'Call Target(1,\nOther(2, 3)',
        'End Sub\n\n',
    ])('preserves full-prefix decisions at every caret in %j', statement => {
        const prefix = 'Sub Demo()\nPrevious(1,\nlabel: ignored = 3\n';
        const source = prefix + statement + '\nEnd Sub\n';
        for (let offset = prefix.length; offset <= source.length; offset++) {
            compareLegacyPrefix(source, offset);
        }
        compareLegacyPrefix(source, -1);
        compareLegacyPrefix(source, source.length + 10);
    });

    it('bounds token visits by the current statement after 24,000 unrelated calls', () => {
        const source = 'Sub Demo()\n' + 'Ignored(1, 2)\r\n'.repeat(24000) + 'Call Target(1, Nested(2, ';
        tokenizeCached(source); // Full lexing is separately cached; warm it first.
        const context = lineContext(source, source.length);
        let visits = 0;
        const tokens = new Proxy(context.significantTokens, {
            get(target, key, receiver) {
                if (typeof key === 'string' && /^\d+$/.test(key)) { visits++; }
                return Reflect.get(target, key, receiver);
            },
        });
        vi.spyOn(cursorContext, 'completionLineCursorContext').mockReturnValue({ ...context, significantTokens: tokens });
        const fullPrefix = vi.spyOn(cursorContext, 'completionCursorContext');
        expect(findActiveCallSite(source, source.length)).toMatchObject({ calleeName: 'Nested', activeParameter: 1 });
        expect(visits).toBeLessThan(80);
        expect(context.significantTokens.length).toBeLessThan(20);
        expect(fullPrefix).not.toHaveBeenCalled();
        visits = 0;
        expect(callableCompletionShouldInsertParens(source, source.length)).toBe(true);
        expect(visits).toBeLessThan(80);
        expect(fullPrefix).not.toHaveBeenCalled();
    });

    it.skipIf(!process.env.XLIDE_INCREMENTAL_PARSE_CORPUS)('preserves private-corpus call contexts without logging source', () => {
        const source = readFileSync(process.env.XLIDE_INCREMENTAL_PARSE_CORPUS!, 'utf8');
        const lines = [...source.matchAll(/\r\n|\r|\n/g)];
        const middle = lines[Math.min(1500, lines.length - 1)].index!;
        for (const center of [middle, source.length - 80]) {
            for (let step = -16; step <= 16; step++) {
                compareLegacyPrefix(source, Math.max(0, Math.min(source.length, center + step)));
            }
        }
    });
});
