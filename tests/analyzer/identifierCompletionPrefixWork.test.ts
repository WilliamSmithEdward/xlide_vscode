import { describe, expect, it, vi } from 'vitest';

// Issue #1443: the guard counts the lexing identifier completion really does.
// The lexer module itself is wrapped, so every importer's calls are seen:
// each tokenize() call's characters, and each tokenizeCached() call with the
// reads of the stream it returns. A spy on one export saw none of the
// line-window lexing, and its zero passed every upper bound.
const work = { lexCalls: 0, lexedChars: 0, cachedCalls: 0, cachedChars: 0, cachedReads: 0 };
const resetWork = () => { for (const key of Object.keys(work) as (keyof typeof work)[]) { work[key] = 0; } };
vi.mock('../../src/analyzer/lexer/tokenize', async (load) => {
    const real = await load<typeof import('../../src/analyzer/lexer/tokenize')>();
    return {
        ...real,
        tokenize: (...args: Parameters<typeof real.tokenize>) => {
            work.lexCalls += 1;
            work.lexedChars += args[0].length;
            return real.tokenize(...args);
        },
        tokenizeCached: (...args: Parameters<typeof real.tokenizeCached>) => {
            work.cachedCalls += 1;
            work.cachedChars += args[0].length;
            return new Proxy(real.tokenizeCached(...args), {
                get(target, property, receiver) {
                    if (typeof property === 'string' && /^\d+$/.test(property)) { work.cachedReads += 1; }
                    return Reflect.get(target, property, receiver);
                },
            });
        },
    };
});

const { resolveIdentifierCompletions } = await import('../../src/analyzer/completion/identifierCompletion');

const padded = Array.from({ length: 1200 }, (_, i) =>
    'Sub Padding' + i + '()\nDebug.Print ' + i + '\nEnd Sub\n').join('') +
    'Sub Probe()\nDim CanonicalValue As Long\ncanonicalvalue\nThisWorkbook.Sheets(1).Na\nEnd Sub';

describe('identifier completion prefix work', () => {
    it('lexes only the current line on each keystroke once the module is warm', () => {
        const source = padded;
        const identifierOffset = source.lastIndexOf('canonicalvalue') + 14;
        const memberOffset = source.lastIndexOf('.Na') + 3;
        // Warm the module's symbol analysis on purpose: the guard is about the
        // per-keystroke grammar work, not the first request's.
        expect(resolveIdentifierCompletions(source, identifierOffset).map(item => item.name)).toContain('CanonicalValue');
        // Fresh offsets, so no cursor-context cache answers for them.
        for (const offset of [identifierOffset - 1, identifierOffset - 2, memberOffset]) {
            resetWork();
            const result = resolveIdentifierCompletions(source, offset);
            expect(work.lexedChars, `offset ${offset}`).toBeGreaterThan(0);
            expect(work.lexedChars, `offset ${offset}`).toBeLessThanOrEqual(120);
            // Measured at 41 to 55 characters, and the shared stream untouched.
            expect(work.cachedCalls, `offset ${offset}`).toBe(0);
            if (offset === memberOffset) { expect(result).toEqual([]); }
            else { expect(result.map(item => item.name)).toContain('CanonicalValue'); }
        }
    });

    it('keeps cold start separately measurable: the first request lexes the module', () => {
        const source = padded.replace('Sub Probe()', 'Sub ColdProbe()');
        resetWork();
        const offset = source.lastIndexOf('canonicalvalue') + 14;
        expect(resolveIdentifierCompletions(source, offset).map(item => item.name)).toContain('CanonicalValue');
        expect(work.lexedChars + work.cachedChars).toBeGreaterThanOrEqual(source.length);
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
