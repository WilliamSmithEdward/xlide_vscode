import { describe, expect, it, vi } from 'vitest';
import { nameAt } from '../src/analyzer/refactor/shared';
function original(source: string, offset: number): string | undefined {
    const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(0, offset));
    const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(offset));
    const name = `${before?.[0] ?? ''}${after?.[0] ?? ''}`;
    return /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(name) ? name : undefined;
}
describe('line-bounded refactor identifier lookup', () => {
    it.each(['\r\n', '\n', '\r'])('matches the prior lookup at every caret boundary with %j breaks', (eol) => {
        const source = ['Public Sub Work()', '  foo = bar12 + _leading', '  x = 123foo + 12', '  obj.Прибор = café + ก้', '  x = 𐐀name + _á', '  x = [two words]', '  text = "name" \' comment', '  x = abc\u2028def\u2029ghi', 'End Sub', ''].join(eol);
        for (let offset = -source.length - 2; offset <= source.length + 2; offset++) {
            expect(nameAt(source, offset), 'offset ' + offset).toBe(original(source, offset));
            expect(nameAt(source, offset + 0.5)).toBe(original(source, offset + 0.5));
        }
        for (const offset of [NaN, Infinity, -Infinity]) expect(nameAt(source, offset)).toBe(original(source, offset));
    });

    it('slices only the caret line rather than a module-sized prefix or suffix', () => {
        const prefix = 'Dim unrelated As Long\r\n'.repeat(20000);
        const line = '    result = obj.Прибор';
        const source = prefix + line + '\r\n' + prefix;
        const offset = prefix.length + line.length - 2;
        const slice = vi.spyOn(String.prototype, 'slice');
        try {
            expect(nameAt(source, offset)).toBe('Прибор');
            const ranges = slice.mock.calls;
            expect(ranges).toEqual([[prefix.length, offset], [offset, prefix.length + line.length]]);
        } finally { slice.mockRestore(); }
    });
});
