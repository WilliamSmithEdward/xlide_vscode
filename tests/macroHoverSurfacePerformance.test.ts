import { afterEach, describe, expect, it, vi } from 'vitest';
import { writeFileSync } from 'node:fs';
import * as lexer from '../src/analyzer/lexer/tokenize';
import { macroNameStringAt } from '../src/analyzer/completion/macroNames';
import { hoverMayResolveAt, resolveHover } from '../src/analyzer/hover/resolveHover';
import { buildLiveVbaProjectIndex } from '../src/vbaProjectAnalysis';

vi.mock('vscode', async () => ({
    ...(await import('./helpers/vscodeMock')).vscodeMock(),
    Hover: class { constructor(public contents: unknown, public range: unknown) {} },
    MarkdownString: class {
        value = '';
        appendCodeblock(text: string) { this.value += text; return this; }
        appendMarkdown(text: string) { this.value += text; return this; }
    },
}));
import * as vscode from 'vscode';
import { VbaHoverSignatureProvider } from '../src/vbaHoverSignatureProvider';

afterEach(() => vi.restoreAllMocks());
const prefix = Array.from({ length: 1200 }, (_, index) =>
    'Sub Probe' + index + '()\n    Debug.Print "sample"\nEnd Sub\n').join('\n');

function documentFor(source: string) {
    return {
        version: 1, uri: { scheme: 'xlide-vba' }, getText: () => source,
        offsetAt: (position: vscode.Position) => position.character,
        positionAt: (offset: number) => new vscode.Position(0, offset),
    } as unknown as vscode.TextDocument;
}
function context() {
    return {
        cachedEditorProjectContext: vi.fn(() => undefined),
        cheapEditorProjectContext: vi.fn(() => ({})),
        localEditorProjectContext: vi.fn(() => ({})),
        warmEditorProjectContext: vi.fn(),
        buildEditorProjectContextWithin: vi.fn(async () => undefined),
    };
}
function countedTokens(source: string) {
    const tokens = lexer.tokenizeCached(source);
    let reads = 0;
    const proxy = new Proxy(tokens, {
        get(target, property, receiver) {
            if (typeof property === 'string' && /^\d+$/.test(property)) { reads++; }
            return Reflect.get(target, property, receiver);
        },
    });
    vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(proxy);
    return { reads: () => reads, length: tokens.length };
}

describe('bounded macro string lookup', () => {
    it('checks fewer than 30 tokens for dot completion at the end of a large module', () => {
        const source = prefix + '\nSub Active()\n    ThisWorkbook.Sheets(1).\nEnd Sub\n';
        const counter = countedTokens(source);
        expect(counter.length).toBeGreaterThan(10000);
        expect(macroNameStringAt(source, source.lastIndexOf('.') + 1)).toBeUndefined();
        expect(counter.reads()).toBeLessThan(30);
    });

    it.each(['shp.OnAction = "Demo.Run"', 'Call Wire(handlerProc:="Demo.Run")'])(
        'uses neighboring tokens without a prefix scan at %s', line => {
            const source = prefix + '\nSub Active()\n    ' + line + '\nEnd Sub\n';
            const counter = countedTokens(source);
            expect(macroNameStringAt(source, source.lastIndexOf('"Demo.') + 6)?.text).toBe('Demo.Run');
            expect(counter.reads()).toBeLessThan(30);
        });

    it('preserves string boundaries, trivia, and incomplete literals', () => {
        const source = 'shp.OnAction = "Demo.Run"  ';
        const start = source.indexOf('"'), end = source.lastIndexOf('"') + 1;
        expect(macroNameStringAt(source, start)).toBeUndefined();
        expect(macroNameStringAt(source, start + 1)?.text).toBe('Demo.Run');
        expect(macroNameStringAt(source, end)?.text).toBe('Demo.Run');
        expect(macroNameStringAt(source, end + 1)).toBeUndefined();
        expect(macroNameStringAt('shp.OnAction = "Demo.R', 22)?.text).toBe('Demo.R');
        expect(macroNameStringAt('', 0)).toBeUndefined();
    });
});

describe('hover preflight', () => {
    it.each(["' Counter", '42', '#1/1/2026#', '+', '    ', "' note _\nCounter"])(
        'avoids all project context work for %j', async fragment => {
            const source = prefix + '\n' + fragment + '\n';
            const offset = prefix.length + 1 + Math.floor(fragment.length / 2);
            const ctx = context();
            expect(await new VbaHoverSignatureProvider(ctx as never).provideHover(
                documentFor(source), new vscode.Position(0, offset))).toBeUndefined();
            for (const spy of Object.values(ctx)) { expect(spy).not.toHaveBeenCalled(); }
        });

    it('keeps identifier boundaries and metadata-dependent strings eligible', () => {
        const source = 'Counter+shp.OnAction="Demo.Run"';
        for (const offset of [0, 3, 7, 8, source.indexOf('"') + 1, source.lastIndexOf('"') + 1]) {
            expect(hoverMayResolveAt(source, offset)).toBe(true);
        }
        expect(hoverMayResolveAt(source, source.length + 1)).toBe(false);
        expect(hoverMayResolveAt('', 0)).toBe(false);
    });

    it('still loads project context for an unknown identifier', async () => {
        const source = 'UnknownName';
        const ctx = context();
        expect(await new VbaHoverSignatureProvider(ctx as never).provideHover(
            documentFor(source), new vscode.Position(0, 3))).toBeUndefined();
        expect(ctx.buildEditorProjectContextWithin).toHaveBeenCalledTimes(1);
    });

    it('still resolves a macro-name hover from project metadata', async () => {
        const source = 'shp.OnAction = "Demo.Run"';
        const projectProcedures = [{
            name: 'Run', moduleName: 'Demo', kind: 'sub', visibility: 'public',
            params: [], span: { start: 0, end: 0 },
        }];
        const ctx = context();
        ctx.buildEditorProjectContextWithin.mockResolvedValue({ projectProcedures } as never);
        const hover = await new VbaHoverSignatureProvider(ctx as never).provideHover(
            documentFor(source), new vscode.Position(0, source.indexOf('Demo') + 2));
        expect(hover).toBeDefined();
        expect(ctx.buildEditorProjectContextWithin).toHaveBeenCalledTimes(1);
        expect(resolveHover(source, source.indexOf('Demo') + 2, { projectProcedures } as never)?.signature).toContain('Run');
    });
});

it.skipIf(!process.env.XLIDE_MACRO_HOVER_BENCHMARK_OUTPUT)('measures macro lookup and empty hover work', async () => {
    const source = prefix + "\nSub Active()\n    ThisWorkbook.Sheets(1).\n    ' comment\nEnd Sub\n";
    const dot = source.lastIndexOf('Sheets(1).') + 'Sheets(1).'.length;
    const comment = source.lastIndexOf("' comment") + 3;
    const doc = documentFor(source);
    const ctx = context();
    ctx.localEditorProjectContext.mockImplementation(() => {
        buildLiveVbaProjectIndex([{ moduleName: 'Active', moduleKind: 'standard', source }]);
        return {};
    });
    const provider = new VbaHoverSignatureProvider(ctx as never);
    const medians: Record<string, number> = {};
    const measure = async (name: string, invoke: () => unknown) => {
        await invoke();
        const times: number[] = [];
        for (let sample = 0; sample < 21; sample++) {
            const start = performance.now(); await invoke(); times.push(performance.now() - start);
        }
        medians[name] = times.sort((a, b) => a - b)[10];
    };
    await measure('nonStringMacroLookupMs', () => macroNameStringAt(source, dot));
    await measure('emptyHoverMs', () => provider.provideHover(doc, new vscode.Position(0, comment)));
    writeFileSync(process.env.XLIDE_MACRO_HOVER_BENCHMARK_OUTPUT!, JSON.stringify({
        procedures: 1200, bytes: source.length, samples: 21, medians,
        hoverContextBuilds: ctx.buildEditorProjectContextWithin.mock.calls.length,
    }, null, 2));
});
