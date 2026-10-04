import { beforeEach, describe, expect, it, vi } from 'vitest';
import { resolveHover } from '../src/analyzer/hover/resolveHover';
import { resolveIdentifierCompletions } from '../src/analyzer/completion/identifierCompletion';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { DocRegistry } from '../src/analyzer/docs/docRegistry';

vi.mock('../src/analyzer/symbols/buildModuleSymbols', async () => {
    const actual = await vi.importActual<typeof import('../src/analyzer/symbols/buildModuleSymbols')>('../src/analyzer/symbols/buildModuleSymbols');
    return { ...actual, buildModuleSymbols: vi.fn(actual.buildModuleSymbols) };
});
beforeEach(() => { vi.mocked(buildModuleSymbols).mockClear(); });

const fixture = (name: string) => 'Sub ' + name + '()\nDim HoverValue As Long\nhovervalue = 1\nEnd Sub';
const at = (source: string) => source.indexOf('\nhovervalue') + 3;

describe('hover symbol snapshot reuse', () => {
    it('shares one projection with completion and returns fresh hover records', () => {
        const source = fixture('HoverShared');
        const offset = at(source);
        expect(resolveIdentifierCompletions(source, offset + 8, { includeGlobals: false, includeRuntime: false })
            .map(item => item.name)).toEqual(['HoverValue']);
        const first = resolveHover(source, offset)!;
        expect(first.signature).toBe('HoverValue As Long');
        first.signature = 'Poisoned';
        first.details.push('Poisoned');
        expect(resolveHover(source, offset)).toMatchObject({ signature: 'HoverValue As Long' });
        expect(resolveHover(source, offset)?.details).not.toContain('Poisoned');
        expect(buildModuleSymbols).toHaveBeenCalledTimes(1);
    });

    it('refreshes edited declarations and selects scope at every mouse position', () => {
        const source = fixture('HoverFirst') + '\nSub HoverSecond()\nDim HoverValue As String\nhovervalue = ""\nEnd Sub';
        expect(resolveHover(source, at(source))?.signature).toBe('HoverValue As Long');
        expect(resolveHover(source, source.lastIndexOf('hovervalue') + 2)?.signature).toBe('HoverValue As String');
        const changed = source.replace('HoverValue As Long', 'HoverValue As Double');
        expect(resolveHover(changed, at(changed))?.signature).toBe('HoverValue As Double');
        expect(buildModuleSymbols).toHaveBeenCalledTimes(2);
    });

    it('keeps module identity and live external documentation independent', () => {
        const source = fixture('HoverIdentity');
        const docs = new DocRegistry();
        docs.add([{ name: 'One.HoverValue', doc: { summary: 'First metadata', params: [], source: 'external' } }]);
        expect(resolveHover(source, at(source), { moduleName: 'One', docRegistry: docs })?.documentation).toContain('First metadata');
        docs.clear();
        docs.add([{ name: 'One.HoverValue', doc: { summary: 'Updated metadata', params: [], source: 'external' } }]);
        expect(resolveHover(source, at(source), { moduleName: 'One', docRegistry: docs })?.documentation).toContain('Updated metadata');
        resolveHover(source, at(source), { moduleName: 'Two' });
        resolveHover(source, at(source), { moduleName: 'Two', moduleKind: 'class' });
        expect(buildModuleSymbols).toHaveBeenCalledTimes(3);
    });

    it('rechecks host-global shadowing after an edit', () => {
        const source = 'Sub HoverShadow()\nDebug.Print ThisWorkbook\nEnd Sub';
        expect(resolveHover(source, source.indexOf('ThisWorkbook') + 2)?.signature).toBe('ThisWorkbook As Workbook');
        resolveHover(source, source.indexOf('ThisWorkbook') + 2);
        const changed = source.replace('\nDebug', '\nDim ThisWorkbook As String\nDebug');
        expect(resolveHover(changed, changed.lastIndexOf('ThisWorkbook') + 2)?.signature).toBe('ThisWorkbook As String');
        expect(buildModuleSymbols).toHaveBeenCalledTimes(2);
    });
});

it.skipIf(process.env.XLIDE_HOVER_SYMBOL_BENCH !== '1')('measures warm bare-symbol hover in a large module', () => {
    const source = Array.from({ length: 1200 }, (_, index) =>
        'Sub HoverPadding' + index + '()\nDebug.Print ' + index + '\nEnd Sub\n').join('') +
        fixture('HoverBenchmark') + '\nSub HostProbe()\nDebug.Print ThisWorkbook\nEnd Sub';
    const medianMs: Record<string, number> = {};
    for (const [name, offset, signature] of [
        ['local', at(source), 'HoverValue As Long'],
        ['hostGlobal', source.lastIndexOf('ThisWorkbook') + 2, 'ThisWorkbook As Workbook'],
    ] as const) {
        expect(resolveHover(source, offset)?.signature).toBe(signature);
        const times = Array.from({ length: 21 }, () => {
            const start = performance.now();
            resolveHover(source, offset);
            return performance.now() - start;
        }).sort((a, b) => a - b);
        medianMs[name] = times[10];
    }
    process.stdout.write('Hover symbol snapshot benchmark: ' + JSON.stringify({ bytes: source.length, medianMs }) + '\n');
});
