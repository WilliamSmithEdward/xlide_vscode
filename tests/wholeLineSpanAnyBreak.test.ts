import { describe, expect, it } from 'vitest';
import { lineStartAtAnyBreak, wholeLineSpanAnyBreak } from '../src/vbaSourceScan';
describe('wholeLineSpanAnyBreak', () => {
    it.each(['\n', '\r\n', '\r'])('includes one complete physical line and its break (%j)', eol => {
        const first = 'before' + eol, target = '    Dim one As Long', source = first + target + eol + 'after';
        const span = wholeLineSpanAnyBreak(source, { start: first.length + 4, end: first.length + target.length });
        expect(source.slice(span.start, span.end)).toBe(target + eol);
        expect(source.slice(0, span.start) + source.slice(span.end)).toBe(first + 'after');
    });
    it('preserves nearest-break results for every offset, including fractional and non-finite positions', () => {
        for (const source of ['', 'single', '\nfirst\nlast', 'before\rnext', 'Café\r\nΔΕΛΤΑ\nend\rtail']) {
            const offsets = [...Array.from({ length: source.length * 2 + 7 }, (_, i) => i / 2 - 1), NaN, Infinity, -Infinity];
            for (const offset of offsets) {
                const expected = offset <= 0 ? 0 : Math.max(source.lastIndexOf('\n', offset - 1), source.lastIndexOf('\r', offset - 1)) + 1;
                expect(lineStartAtAnyBreak(source, offset)).toBe(expected);
            }
        }
    });

    it('ends at EOF without adding a nonexistent break', () => {
        expect(wholeLineSpanAnyBreak('before\r  final', { start: 9, end: 14 })).toEqual({ start: 7, end: 14 });
        expect(wholeLineSpanAnyBreak('', { start: 0, end: 0 })).toEqual({ start: 0, end: 0 });
    });
});
