import { afterEach, describe, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';

afterEach(() => vi.restoreAllMocks());

describe('identifier completion prefix work', () => {
    it('does not copy preceding procedures for identifier or member-position checks', () => {
        const source = Array.from({ length: 1200 }, (_, i) =>
            'Sub Padding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('') +
            'Sub Probe()\nDim CanonicalValue As Long\ncanonicalvalue\nThisWorkbook.Sheets(1).Na\nEnd Sub';
        const identifierOffset = source.lastIndexOf('canonicalvalue') + 14;
        const memberOffset = source.lastIndexOf('.Na') + 3;
        expect(resolveIdentifierCompletions(source, identifierOffset).map(item => item.name)).toContain('CanonicalValue');
        const tokens = lexer.tokenizeCached(source);
        let reads = 0;
        vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(new Proxy(tokens, {
            get(target, property, receiver) {
                if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                return Reflect.get(target, property, receiver);
            },
        }));
        for (const offset of [identifierOffset - 1, identifierOffset - 2, memberOffset]) {
            reads = 0;
            const result = resolveIdentifierCompletions(source, offset);
            expect(reads).toBeLessThan(100);
            if (offset === memberOffset) { expect(result).toEqual([]); }
            else { expect(result.map(item => item.name)).toContain('CanonicalValue'); }
        }
    });

    it('preserves the preceding-token policy at a blank position without crossing a second newline', () => {
        const resolve = (text: string) => {
            const source = 'Sub Probe()\n' + text;
            return resolveIdentifierCompletions(source, source.length).map(item => item.name);
        };
        expect(resolve('Dim item As\n')).toEqual([]);
        expect(resolve('Dim item As\n\n')).toContain('ThisWorkbook');
        expect(resolve('item =\n')).toContain('True');
        expect(resolve('item =\n\n')).not.toContain('True');
        expect(resolve('Call\n')).toContain('MsgBox');
        expect(resolve('Call\n')).not.toContain('DoEvents');
        expect(resolve('Call\n\n')).toContain('DoEvents');
    });

    it('preserves explicit Call, continued statements and newline expression boundaries', () => {
        for (const prefix of ['Call ', '10 Call _\r\n ', 'x = 1: Call ']) {
            const source = 'Sub Probe()\n' + prefix + 'MsgB\nEnd Sub';
            const result = resolveIdentifierCompletions(source, source.indexOf('MsgB') + 4);
            expect(result.map(item => item.name)).toContain('MsgBox');
        }
        for (const prefix of ['If True Then\n', 'x = 1\n', 'Call MsgBox(1)\n']) {
            const source = 'Sub Probe()\n' + prefix + 'Th\nEnd Sub';
            expect(resolveIdentifierCompletions(source, source.indexOf('\nTh') + 3).map(item => item.name)).toContain('ThisWorkbook');
        }
    });
});
