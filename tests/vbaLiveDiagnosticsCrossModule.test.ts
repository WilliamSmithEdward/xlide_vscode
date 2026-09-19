// Live diagnostics across modules, through the real registration and the real
// project index over a fake engine. A module's findings depend on the others:
// `HelperMod.Greet` is only defined while HelperMod exists. So when another
// module is created or changes - an agent writing it, a save from its editor -
// the module that calls it is analyzed again, with no edit of its own and no
// change of editor. It used to keep its "not defined" until something touched
// its own editor.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

const state = vi.hoisted(() => ({
    published: new Map<string, string[]>(),
    publishes: new Map<string, number>(),
}));

vi.mock('vscode', async () => {
    const helper = await import('./helpers/vscodeMock');
    class Diagnostic {
        source?: string;
        code?: unknown;
        tags?: unknown[];
        constructor(public range: unknown, public message: string, public severity: number) {}
    }
    const noop = () => ({ dispose: () => undefined });
    return helper.vscodeMock({
        Diagnostic,
        DiagnosticSeverity: { Error: 0, Warning: 1, Information: 2, Hint: 3 },
        DiagnosticTag: { Unnecessary: 1, Deprecated: 2 },
        languages: {
            createDiagnosticCollection: () => ({
                set: (uri: unknown, diagnostics: Array<{ code?: unknown }>) => {
                    state.published.set(String(uri), diagnostics.map((d) => String(d.code)));
                    state.publishes.set(String(uri), (state.publishes.get(String(uri)) ?? 0) + 1);
                },
                delete: (uri: unknown) => { state.published.delete(String(uri)); },
                get: () => undefined,
                clear: () => undefined,
                dispose: () => undefined,
            }),
        },
        window: {
            activeTextEditor: undefined,
            onDidChangeActiveTextEditor: noop,
            onDidChangeTextEditorSelection: noop,
        },
    });
});

import * as vscode from 'vscode';
import { registerVbaDiagnostics } from '../src/vbaLiveDiagnostics';
import { VbaSymbolIndex } from '../src/vbaSymbolIndex';
import { VbaProjectIndexService } from '../src/vbaProjectIndexService';
import { fakeProjectEngine, type FakeBridgeModule } from './helpers/fakeProjectEngine';

const BOOK = process.platform === 'win32' ? 'C:/Book.xlsm' : '/work/Book.xlsm';
const BOOK_URI_PATH = BOOK.startsWith('/') ? BOOK : `/${BOOK}`;

const CALLER = 'Option Explicit\n\nPublic Sub Run()\n    HelperMod.Greet "x"\n    Greet2 "y"\nEnd Sub\n';
const HELPER = 'Option Explicit\n\nPublic Sub Greet(ByVal who As String)\n    Debug.Print who\nEnd Sub\n';

/** A module of the project, open in an editor. */
function moduleDocument(moduleName: string, source: string): vscodeTypes.TextDocument {
    const lines = source.split('\n');
    const lineStarts: number[] = [];
    let offset = 0;
    for (const line of lines) {
        lineStarts.push(offset);
        offset += line.length + 1;
    }
    const positionAt = (at: number) => {
        let line = 0;
        while (line + 1 < lineStarts.length && lineStarts[line + 1] <= at) {
            line++;
        }
        return new vscode.Position(line, at - lineStarts[line]);
    };
    const value = `xlide-vba:${BOOK_URI_PATH}/${moduleName}.bas`;
    return {
        uri: { scheme: 'xlide-vba', path: `${BOOK_URI_PATH}/${moduleName}.bas`, toString: () => value },
        languageId: 'vba',
        version: 1,
        isClosed: false,
        lineCount: lines.length,
        getText: () => source,
        lineAt: (line: number) => ({ text: lines[line] ?? '' }),
        positionAt,
        offsetAt: (position: { line: number; character: number }) => lineStarts[position.line] + position.character,
    } as unknown as vscodeTypes.TextDocument;
}

const documents = (): vscodeTypes.TextDocument[] =>
    vscode.workspace.textDocuments as vscodeTypes.TextDocument[];

const published = (document: vscodeTypes.TextDocument): string[] | undefined =>
    state.published.get(document.uri.toString());
const publishes = (document: vscodeTypes.TextDocument): number =>
    state.publishes.get(document.uri.toString()) ?? 0;

/**
 * Waits until the check holds. A pass reads the project's settings file from
 * disk, which fake timers do not wait for, so this advances the debounce
 * timers and lets real time pass in turn.
 */
function until(check: () => void): Promise<void> {
    return vi.waitFor(check, { timeout: 2000, interval: 25 });
}

/** The engine, the index over it, and diagnostics published for every open document. */
async function project(modules: FakeBridgeModule[], open: vscodeTypes.TextDocument[]): Promise<VbaSymbolIndex> {
    const index = new VbaSymbolIndex(fakeProjectEngine(modules));
    documents().push(...open);
    registerVbaDiagnostics(
        { subscriptions: [] } as unknown as vscodeTypes.ExtensionContext,
        new VbaProjectIndexService(index),
    );
    await until(() => {
        for (const document of open) {
            expect(publishes(document)).toBeGreaterThan(0);
        }
    });
    return index;
}

beforeEach(() => {
    vi.useFakeTimers();
    state.published.clear();
    state.publishes.clear();
    documents().length = 0;
});

afterEach(() => {
    vi.useRealTimers();
});

describe('diagnostics across modules', () => {
    it('analyze a module again when the module it calls is created', async () => {
        const modules: FakeBridgeModule[] = [{ name: 'CallerMod', type: 'standard', source: CALLER }];
        const caller = moduleDocument('CallerMod', CALLER);
        const index = await project(modules, [caller]);
        expect(published(caller)).toContain('undeclared-variable');

        // What a module write does: the engine has the module now, and the
        // project is invalidated. Nothing touches CallerMod.
        modules.push({ name: 'HelperMod', type: 'standard', source: HELPER });
        index.invalidate(BOOK);

        await until(() => expect(published(caller)).not.toContain('undeclared-variable'));
    });

    it('analyze every other open module again when a module is saved with what one of them calls', async () => {
        const helperSource = 'Option Explicit\n\nPublic Sub Other()\nEnd Sub\n';
        const modules: FakeBridgeModule[] = [
            { name: 'CallerMod', type: 'standard', source: CALLER },
            { name: 'HelperMod', type: 'standard', source: HELPER },
            { name: 'Helper2', type: 'standard', source: helperSource },
        ];
        const caller = moduleDocument('CallerMod', CALLER);
        const helper = moduleDocument('HelperMod', HELPER);
        const index = await project(modules, [caller, helper]);
        expect(published(caller)).toContain('unknown-call');
        const helperBefore = publishes(helper);

        // A save from Helper2's editor, which is how the index hears of it.
        index.updateModuleSource(BOOK, 'Helper2', `${helperSource}\nPublic Sub Greet2(ByVal who As String)\nEnd Sub\n`);

        await until(() => expect(published(caller)).toEqual([]));
        await until(() => expect(publishes(helper)).toBe(helperBefore + 1));
    });

    it('do not analyze the saved module itself again, since its own text did not change', async () => {
        const modules: FakeBridgeModule[] = [
            { name: 'CallerMod', type: 'standard', source: CALLER },
            { name: 'HelperMod', type: 'standard', source: HELPER },
        ];
        const caller = moduleDocument('CallerMod', CALLER);
        const helper = moduleDocument('HelperMod', HELPER);
        const index = await project(modules, [caller, helper]);
        const callerBefore = publishes(caller);
        const helperBefore = publishes(helper);

        index.updateModuleSource(BOOK, 'CallerMod', CALLER);

        // HelperMod's run shows the change was taken in and analyzed.
        await until(() => expect(publishes(helper)).toBe(helperBefore + 1));
        expect(publishes(caller)).toBe(callerBefore);
    });
});
