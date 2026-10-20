import { afterEach, describe, expect, it, vi } from 'vitest';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';

afterEach(() => vi.restoreAllMocks());

describe('identifier completion prefix work', () => {
    it.each([['identifier', 1], ['identifier', 2], ['member', 0]] as const)(
        'does not copy preceding procedures for %s checks (%i characters back)', (position, backtrack) => {
            const source = Array.from({ length: 1200 }, (_, i) =>
                'Sub Padding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('') +
                'Sub Probe()\nDim CanonicalValue As Long\ncanonicalvalue\nThisWorkbook.Sheets(1).Na\nEnd Sub';
            const identifierOffset = source.lastIndexOf('canonicalvalue') + 14;
            const memberOffset = source.lastIndexOf('.Na') + 3;
            expect(resolveIdentifierCompletions(source, identifierOffset).map(item => item.name)).toContain('CanonicalValue');
            let reads = 0;
            let lexedCharacters = 0;
            const countReads = (tokens: ReturnType<typeof lexer.tokenize>) => new Proxy(tokens, {
                get(target, property, receiver) {
                    if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                    return Reflect.get(target, property, receiver);
                },
            });
            // Warm symbols above, but use fresh cursor offsets below so the line
            // context cache cannot hide grammar work. Count both the current line
            // lexer and cached token reads if a full-prefix path is reintroduced.
            const tokenize = lexer.tokenize;
            const tokenizeCached = lexer.tokenizeCached;
            vi.spyOn(lexer, 'tokenize').mockImplementation(text => {
                lexedCharacters += text.length;
                return countReads(tokenize(text));
            });
            vi.spyOn(lexer, 'tokenizeCached').mockImplementation(text => countReads(tokenizeCached(text)));
            const offset = position === 'member' ? memberOffset : identifierOffset - backtrack;
            const result = resolveIdentifierCompletions(source, offset);
            expect(reads).toBeGreaterThan(0);
            expect(reads).toBeLessThan(100);
            expect(lexedCharacters).toBeGreaterThan(0);
            expect(lexedCharacters).toBeLessThan(100);
            if (position === 'member') { expect(result).toEqual([]); }
            else { expect(result.map(item => item.name)).toContain('CanonicalValue'); }
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
