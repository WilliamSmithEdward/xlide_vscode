import { afterEach, describe, expect, it, vi } from 'vitest';
import * as cursor from '../src/analyzer/completion/cursorContext';
import { resolveTypeCompletions } from '../src/analyzer/completion/typeCompletion';
import { resolveCanonicalCaseEdits } from '../src/analyzer/completion/canonicalCasing';

afterEach(() => { vi.restoreAllMocks(); });

describe('type-position lookback work', () => {
    it('reads only the trailing grammar tokens in a large ordinary expression', () => {
        const source = 'Debug.Print 1\n'.repeat(1200) + 'ordinaryValue';
        const actual = cursor.completionCursorContext(source, source.length);
        let reads = 0;
        const tokens = new Proxy(actual.significantTokens, {
            get(target, property, receiver) {
                if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
                return Reflect.get(target, property, receiver);
            },
        });
        vi.spyOn(cursor, 'completionCursorContext').mockReturnValue({ ...actual, significantTokens: tokens });
        expect(resolveTypeCompletions(source, source.length)).toEqual([]);
        expect(reads).toBeLessThanOrEqual(7);
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
