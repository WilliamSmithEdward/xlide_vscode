// xlide.analysis.ignoreFilesOutsideTree, driven through the real live
// diagnostics registration: a loose file on disk is not analyzed by default,
// is analyzed when the setting is off, and loses its findings the moment the
// setting is turned back on.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type * as vscodeTypes from 'vscode';

const state = vi.hoisted(() => ({
    ignoreOutsideTree: undefined as boolean | undefined,
    published: new Map<string, string[]>(),
    deleted: [] as string[],
    configListeners: [] as Array<(e: { affectsConfiguration(section: string): boolean }) => void>,
}));

vi.mock('vscode', async () => {
    const helper = await import('./helpers/vscodeMock');
    class Range {
        constructor(..._args: unknown[]) {}
    }
    class Diagnostic {
        source?: string;
        code?: unknown;
        tags?: unknown[];
        constructor(public range: unknown, public message: string, public severity: number) {}
    }
    const noop = () => ({ dispose: () => undefined });
    return helper.vscodeMock({
        Range,
        Diagnostic,
        DiagnosticSeverity: { Error: 0, Warning: 1, Information: 2, Hint: 3 },
        DiagnosticTag: { Unnecessary: 1, Deprecated: 2 },
        languages: {
            createDiagnosticCollection: () => ({
                set: (uri: unknown, diagnostics: Array<{ code?: unknown }>) => {
                    state.published.set(String(uri), diagnostics.map((d) => String(d.code)));
                },
                delete: (uri: unknown) => {
                    state.deleted.push(String(uri));
                    state.published.delete(String(uri));
                },
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
        workspace: {
            getConfiguration: () => ({
                get: (key: string, fallback?: unknown) =>
                    key === 'analysis.ignoreFilesOutsideTree' && state.ignoreOutsideTree !== undefined
                        ? state.ignoreOutsideTree
                        : fallback,
                inspect: () => ({}),
            }),
            onDidChangeConfiguration: (listener: (typeof state.configListeners)[number]) => {
                state.configListeners.push(listener);
                return { dispose: () => undefined };
            },
        },
    });
});

import * as vscode from 'vscode';
import { registerVbaDiagnostics } from '../src/vbaLiveDiagnostics';

/** A module saved on disk that no project claims, the way an export leaves one. */
function looseDocument(source: string, fileName = 'Module1.bas'): vscodeTypes.TextDocument {
    const fsPath = `C:\\work\\export\\${fileName}`;
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
    return {
        uri: {
            scheme: 'file',
            fsPath,
            path: `/${fsPath.replace(/\\/g, '/')}`,
            toString: () => `file:///${fsPath.replace(/\\/g, '/')}`,
        },
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

// `Me` in a standard module does not compile, so an analyzed copy reports it.
const WITH_A_FINDING = 'Option Explicit\nPublic Sub T()\n    Debug.Print Me.Name\nEnd Sub\n';

const documents = (): vscodeTypes.TextDocument[] =>
    vscode.workspace.textDocuments as vscodeTypes.TextDocument[];

async function openAndAnalyze(document: vscodeTypes.TextDocument): Promise<void> {
    documents().push(document);
    registerVbaDiagnostics({ subscriptions: [] } as unknown as vscodeTypes.ExtensionContext, {} as never);
    await vi.advanceTimersByTimeAsync(10_000);
}

function settingChanged(): void {
    for (const listener of state.configListeners) {
        listener({ affectsConfiguration: (section) => 'xlide.analysis.ignoreFilesOutsideTree'.startsWith(section) });
    }
}

beforeEach(() => {
    vi.useFakeTimers();
    state.ignoreOutsideTree = undefined;
    state.published.clear();
    state.deleted = [];
    state.configListeners = [];
    documents().length = 0;
});

afterEach(() => {
    vi.useRealTimers();
});

describe('files outside the XLIDE tree', () => {
    it('are not analyzed by default', async () => {
        const document = looseDocument(WITH_A_FINDING);

        await openAndAnalyze(document);

        expect(state.published.get(document.uri.toString()) ?? []).toEqual([]);
    });

    it('are analyzed as standalone modules when the setting is off', async () => {
        state.ignoreOutsideTree = false;
        const document = looseDocument(WITH_A_FINDING);

        await openAndAnalyze(document);

        expect(state.published.get(document.uri.toString())).toContain('me-outside-object-module');
    });

    it('lose their findings as soon as the setting is turned back on', async () => {
        state.ignoreOutsideTree = false;
        const document = looseDocument(WITH_A_FINDING);
        await openAndAnalyze(document);
        expect(state.published.get(document.uri.toString())).toContain('me-outside-object-module');

        state.ignoreOutsideTree = true;
        settingChanged();
        await vi.advanceTimersByTimeAsync(10_000);

        expect(state.published.has(document.uri.toString())).toBe(false);
        expect(state.deleted).toContain(document.uri.toString());
    });
});
