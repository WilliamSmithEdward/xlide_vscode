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

    it('starts a fresh statement on the row below any line, finished or not (#1527)', () => {
        const resolve = (text: string) => {
            const source = 'Sub Probe()\n' + text;
            return resolveIdentifierCompletions(source, source.length).map(item => item.name);
        };
        // The VBE ends a statement at the line break: each of these rows is a
        // syntax error on its own line, and the empty row below starts anew.
        for (const line of ['Dim item As', 'Set c = New', 'Dim', 'Const', 'ReDim', 'Static', 'Private', 'x.', 'Call', 'item =']) {
            for (const breakText of ['\n', '\r\n']) {
                const names = resolve(line + breakText);
                expect(names, line).toContain('ThisWorkbook');
                expect(names, line).toContain('DoEvents');
                expect(names, line).not.toContain('True');
            }
        }
        // A continuation keeps the statement going.
        expect(resolve('item = _\n')).toContain('True');
        expect(resolve('item = _\r\n')).toContain('True');
    });

    it('lexes only the caret row below a long tail of blank rows (#1527)', () => {
        const source = padded + '\n'.repeat(4000);
        resolveIdentifierCompletions(source, source.length);
        resetWork();
        expect(resolveIdentifierCompletions(source, source.length - 1).map(item => item.name)).toContain('ThisWorkbook');
        expect(work.lexedChars).toBeLessThanOrEqual(4);
        expect(work.cachedCalls).toBe(0);
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
