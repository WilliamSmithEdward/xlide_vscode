import { afterEach, describe, expect, it, vi } from 'vitest';

// Issue #1443: the guards count the lexing completion really does. The lexer
// module itself is wrapped, so every importer's calls are seen: each
// tokenize() call's characters, and each tokenizeCached() call with the
// reads of the stream it returns. A spy on one export saw none of the
// bounded-window lexing, and its zero passed every upper bound.
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

const { resolveTypeCompletions } = await import('../../src/analyzer/completion/typeCompletion');
const { resolveCanonicalCaseEdits } = await import('../../src/analyzer/completion/canonicalCasing');

afterEach(() => { vi.restoreAllMocks(); });

describe('type-position lookback work', () => {
    it('lexes only the last lines of a large ordinary expression, never the whole stream', () => {
        const source = 'Debug.Print 1\n'.repeat(1200) + 'ordinaryValue';
        resetWork();
        expect(resolveTypeCompletions(source, source.length)).toEqual([]);
        // Positive: the bounded window was lexed. Bounded: a few lines of it.
        expect(work.lexedChars).toBeGreaterThan(0);
        expect(work.lexedChars).toBeLessThanOrEqual(80);
        expect(work.cachedCalls).toBe(0);
        expect(work.cachedReads).toBe(0);
    });

    it('falls back to the shared token stream only behind a long tail of comment lines', async () => {
        const source = 'value = 1\n' + "' note\n".repeat(40);
        resetWork();
        expect(resolveTypeCompletions(source, source.length)).toEqual([]);
        expect(work.lexCalls).toBe(32);
        expect(work.cachedCalls).toBe(1);
        expect(work.cachedChars).toBe(source.length);
        // The fallback walks the shared stream once: a few reads per token.
        const { tokenize: realTokenize } = await vi.importActual<typeof import('../src/analyzer/lexer/tokenize')>('../../src/analyzer/lexer/tokenize');
        expect(work.cachedReads).toBeGreaterThan(0);
        expect(work.cachedReads).toBeLessThanOrEqual(3 * realTokenize(source).length);
    });
    it('retains As New mode for the longest qualified suffix', () => {
        const source = 'Dim item As New Shapes.Pe';
        const result = resolveTypeCompletions(source, source.length, {
            projectTypes: [
                { name: 'Person', kind: 'class', moduleName: 'Shapes' },
                { name: 'Period', kind: 'enum', moduleName: 'Shapes' },
            ],
        });
        expect(result.map(item => item.name)).toEqual(['Person']);
    });

    it('preserves type position through continuation and newline tokens', () => {
        for (const source of ['Dim item As _\r\n New _\r\n Shapes.Pe', 'Dim item As\nNew\nShapes.Pe']) {
            expect(resolveTypeCompletions(source, source.length, {
                projectTypes: [{ name: 'Person', kind: 'class', moduleName: 'Shapes' }],
            }).map(item => item.name)).toEqual(['Person']);
        }
    });

    it('keeps expression New, library qualifiers and non-type positions distinct', () => {
        const source = 'Set item = New Excel.Wor';
        expect(resolveTypeCompletions(source, source.length).map(item => item.name)).toContain('Workbook');
        for (const ordinary of ['value = library.Name', 'Dim item As Long: value', 'Dim item As New Shapes.Pe + value']) {
            expect(resolveTypeCompletions(ordinary, ordinary.length)).toEqual([]);
        }
        const library = 'Dim item As Excel.';
        expect(resolveTypeCompletions(library, library.length).map(item => item.name)).toContain('Worksheet');
    });
});

it.skipIf(process.env.XLIDE_TYPE_LOOKBACK_BENCH !== '1')('measures warm type checks and casing at the end of a large module', () => {
    const source = Array.from({ length: 1200 }, (_, index) =>
        'Sub Padding' + index + '()\nDebug.Print ' + index + '\nEnd Sub\n').join('') +
        'Sub LookbackProbe()\nDim CanonicalValue As Long\ncanonicalvalue = canonicalvalue + canonicalvalue\nEnd Sub\n';
    const start = source.lastIndexOf('\ncanonicalvalue') + 1;
    const end = source.indexOf('\nEnd Sub', start);
    expect(resolveTypeCompletions(source, end)).toEqual([]);
    expect(resolveCanonicalCaseEdits(source, { start, end }, { identifier: { includeGlobals: false, includeRuntime: false } })).toHaveLength(3);
    const medianMs: Record<string, number> = {};
    for (const [name, run] of [
        ['typeCheck', () => resolveTypeCompletions(source, end)],
        ['lineCasing', () => resolveCanonicalCaseEdits(source, { start, end }, { identifier: { includeGlobals: false, includeRuntime: false } })],
    ] as const) {
        run();
        const times = Array.from({ length: 21 }, () => {
            const before = performance.now();
            run();
            return performance.now() - before;
        }).sort((a, b) => a - b);
        medianMs[name] = times[10];
    }
    process.stdout.write('Type lookback benchmark: ' + JSON.stringify({ bytes: source.length, medianMs }) + '\n');
});
