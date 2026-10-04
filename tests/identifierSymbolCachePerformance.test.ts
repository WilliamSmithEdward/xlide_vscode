import { beforeEach, describe, expect, it, vi } from 'vitest';
import { resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';
import { resolveCanonicalCaseEdits } from '../src/analyzer/completion/canonicalCasing';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

vi.mock('../src/analyzer/symbols/buildModuleSymbols', async () => {
    const actual = await vi.importActual<typeof import('../src/analyzer/symbols/buildModuleSymbols')>('../src/analyzer/symbols/buildModuleSymbols');
    return { ...actual, buildModuleSymbols: vi.fn(actual.buildModuleSymbols) };
});
beforeEach(() => { vi.mocked(buildModuleSymbols).mockClear(); });

const context = { includeGlobals: false, includeRuntime: false };
const fixture = (name: string) => 'Sub ' + name + '()\nDim CachedValue As Long\ncached\nEnd Sub';
const complete = (source: string, ctx = context) => resolveIdentifierCompletions(source, source.indexOf('\ncached') + 7, ctx);

describe('identifier symbol snapshot reuse', () => {
    it('shares projection between repeated completion and casing without sharing returned records', () => {
        const source = fixture('CacheRepeated');
        const first = complete(source);
        expect(first.map(item => item.name)).toEqual(['CachedValue']);
        first[0].name = 'Poisoned';
        first[0].detail = 'Poisoned';
        expect(complete(source)[0]).toMatchObject({ name: 'CachedValue', detail: 'local variable As Long' });
        resolveCanonicalCaseEdits(source, { start: 0, end: source.length }, { identifier: context });
        expect(buildModuleSymbols).toHaveBeenCalledTimes(1);
    });

    it('reprojects edited declarations and never leaks locals from another procedure', () => {
        const source = fixture('CacheEdited') + '\nSub CacheSecond()\nDim CachedOther As String\ncached\nEnd Sub';
        expect(complete(source).map(item => item.name)).toEqual(['CachedValue']);
        expect(resolveIdentifierCompletions(source, source.lastIndexOf('cached') + 6, context).map(item => item.name)).toEqual(['CachedOther']);
        const changed = source.replace('CachedValue As Long', 'CachedRenamed As String');
        expect(complete(changed)).toEqual([expect.objectContaining({ name: 'CachedRenamed', detail: 'local variable As String' })]);
        expect(buildModuleSymbols).toHaveBeenCalledTimes(2);
    });

    it('keys projection by module name and kind while refreshing external metadata', () => {
        const source = fixture('CacheIdentity');
        const resolve = (moduleName: string, moduleKind: 'standard' | 'class', codeName: string) =>
            resolveIdentifierCompletions(source, source.indexOf('\ncached') + 7,
                { ...context, moduleName, moduleKind, codeNames: [codeName] });
        expect(resolve('One', 'standard', 'CachedSheetOne').map(item => item.name)).toContain('CachedSheetOne');
        expect(resolve('One', 'standard', 'CachedSheetTwo').map(item => item.name)).not.toContain('CachedSheetOne');
        resolve('Two', 'standard', 'CachedSheetTwo');
        resolve('Two', 'class', 'CachedSheetTwo');
        expect(buildModuleSymbols).toHaveBeenCalledTimes(3);
        expect(vi.mocked(buildModuleSymbols).mock.calls.map(call => call.slice(0, 2))).toEqual([
            ['One', 'standard'], ['Two', 'standard'], ['Two', 'class'],
        ]);
    });

    it('evicts the least recently used snapshot after eight entries', () => {
        const sources = Array.from({ length: 9 }, (_, index) => fixture('CacheEviction' + index));
        sources.slice(0, 8).forEach(source => complete(source));
        complete(sources[0]);
        complete(sources[8]);
        complete(sources[0]);
        expect(buildModuleSymbols).toHaveBeenCalledTimes(9);
        complete(sources[1]);
        expect(buildModuleSymbols).toHaveBeenCalledTimes(10);
    });

    it('keeps declaration-name completion lazy', () => {
        const source = fixture('CacheLazy');
        expect(resolveIdentifierCompletions(source, source.indexOf('CachedValue') + 11, context)).toEqual([]);
        expect(buildModuleSymbols).not.toHaveBeenCalled();
    });
});

it.skipIf(process.env.XLIDE_IDENTIFIER_CACHE_BENCH !== '1')('measures repeated identifier completion and casing', () => {
    const source = fixture('CacheBenchmark').replace('\ncached\n', '\ncachedvalue\n') + Array.from({ length: 1200 }, (_, index) =>
        "\n' Documentation for procedure " + index + "\nPublic Function Extra" + index + "() As Long\nExtra" + index + " = " + index + "\nEnd Function\n").join('');
    const offset = source.indexOf('\ncachedvalue') + 12;
    expect(resolveIdentifierCompletions(source, offset, context).map(item => item.name)).toEqual(['CachedValue']);
    expect(resolveCanonicalCaseEdits(source, { start: source.indexOf('\ncachedvalue') + 1, end: offset }, { identifier: context })).toHaveLength(1);
    const samples: Record<string, number> = {};
    for (const [name, run] of [
        ['completion', () => resolveIdentifierCompletions(source, offset, context)],
        ['casing', () => resolveCanonicalCaseEdits(source, { start: source.indexOf('\ncached') + 1, end: offset }, { identifier: context })],
    ] as const) {
        run();
        const times = Array.from({ length: 21 }, () => {
            const start = performance.now();
            run();
            return performance.now() - start;
        }).sort((a, b) => a - b);
        samples[name] = times[10];
    }
    process.stdout.write('Identifier symbol cache benchmark: ' + JSON.stringify({ bytes: source.length, medianMs: samples }) + '\n');
});
