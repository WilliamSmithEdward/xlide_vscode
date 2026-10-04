import { describe, expect, it } from 'vitest';
import { nameAt } from '../src/analyzer/refactor/shared';
import { IDENT_RE } from '../src/analyzer/lexer/tokenHelpers';
const previousNameAt = (source: string, offset: number) => {
    const before = /[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.exec(source.slice(0, offset));
    const after = /^[\p{L}\p{M}\p{N}_]*/u.exec(source.slice(offset));
    const name = `${before?.[0] ?? ''}${after?.[0] ?? ''}`;
    return IDENT_RE.test(name) ? name : undefined;
};
describe('caret-local refactor name lookup', () => {
    it.each(['value', '123value', 'value123', '_value', 'é', 'ค่า', '𐐀name', '123', '"value"', 'obj.value', 'x + y'])('preserves word boundaries for %s', word => {
        const source = 'Sub Demo()\n' + word + '\nEnd Sub';
        for (let offset = -1; offset <= source.length + 1; offset++) {
            expect(nameAt(source, offset), String(offset)).toBe(previousNameAt(source, offset));
        }
    });
    it('finds the word at the end of a large class prefix', () => {
        const source = 'Dim something As Long\n'.repeat(26000) + 'LatencyValue';
        expect(nameAt(source, source.length)).toBe('LatencyValue');
        const start = performance.now();
        for (let i = 0; i < 500; i++) { nameAt(source, source.length - i % 10); }
        process.stdout.write('Large-class caret name lookup, 500 calls (ms): ' + (performance.now() - start) + '\n');
    });
});
