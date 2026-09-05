import * as vscode from 'vscode';
import {
    encapsulateField,
    extractMethod,
    extractVariable,
    implementInterface,
    inlineVariable,
    introduceParameter,
    moveToModule,
    type VbaRefactorResult,
} from '../analyzer';
import { registerXlideCommand } from '../xlideCommandRegistration';
import { moduleLocationOfDocument } from '../vbaDocumentLocation';
import { isVbaDocument } from '../xlideFileSystem';
import { statusMessage, type CommandDeps } from './shared';

/**
 * The seven refactorings beyond rename (issue #69), as commands.
 *
 * The engine is pure and lives in `src/analyzer/refactor`; this is the layer
 * that finds the caret, fetches whatever project text a refactoring needs, and
 * turns its answer into a workspace edit. A refusal is surfaced as its own
 * sentence, because the refusals are the design: each one names a rule that
 * would otherwise be broken silently.
 */

export const XLIDE_REFACTOR_COMMANDS = [
    'xlide.refactor.extractMethod',
    'xlide.refactor.extractVariable',
    'xlide.refactor.inlineVariable',
    'xlide.refactor.encapsulateField',
    'xlide.refactor.implementInterface',
    'xlide.refactor.moveToModule',
    'xlide.refactor.introduceParameter',
] as const;

export function registerRefactorCommands(deps: CommandDeps): vscode.Disposable[] {
    return [
        registerXlideCommand('xlide.refactor.extractMethod', () => runLocal(
            (source, editor) => extractMethod({ source, span: selectionSpan(editor) }),
        )),
        registerXlideCommand('xlide.refactor.extractVariable', () => runLocal(
            (source, editor) => extractVariable({ source, span: selectionSpan(editor) }),
        )),
        registerXlideCommand('xlide.refactor.inlineVariable', () => runLocal(
            (source, editor) => inlineVariable({ source, offset: caretOffset(editor) }),
        )),
        registerXlideCommand('xlide.refactor.encapsulateField', () => runLocal(
            (source, editor) => encapsulateField({ source, offset: caretOffset(editor) }),
        )),
        registerXlideCommand('xlide.refactor.implementInterface', () => runProjectWide(
            deps,
            async (source, editor, project) => implementInterface({
                source,
                moduleSources: project.sources,
                ...(await pickInterface(source) ?? {}),
            }),
        )),
        registerXlideCommand('xlide.refactor.introduceParameter', () => runProjectWide(
            deps,
            async (source, editor, project) => introduceParameter({
                source,
                offset: caretOffset(editor),
                moduleName: project.moduleName,
                otherModuleSources: project.sources,
            }),
        )),
        registerXlideCommand('xlide.refactor.moveToModule', () => runProjectWide(
            deps,
            async (source, editor, project) => {
                const target = await pickModule(Object.keys(project.sources), project.moduleName);
                if (!target) {
                    return undefined;
                }
                return moveToModule({
                    source,
                    offset: caretOffset(editor),
                    moduleName: project.moduleName,
                    targetModuleName: target,
                    otherModuleSources: project.sources,
                });
            },
        )),
    ];
}

/** A refactoring that needs nothing but the open module. */
async function runLocal(
    compute: (source: string, editor: vscode.TextEditor) => VbaRefactorResult,
): Promise<void> {
    const editor = vbaEditor();
    if (!editor) { return; }
    await applyResult(editor, compute(editor.document.getText(), editor));
}

interface ProjectText {
    moduleName: string;
    /** Every OTHER module in the project, keyed by name. */
    sources: Record<string, string>;
}

/** A refactoring that reads or writes other modules of the same project. */
async function runProjectWide(
    deps: CommandDeps,
    compute: (
        source: string,
        editor: vscode.TextEditor,
        project: ProjectText,
    ) => Promise<VbaRefactorResult | undefined>,
): Promise<void> {
    const editor = vbaEditor();
    if (!editor) { return; }
    const location = moduleLocationOfDocument(editor.document);
    if (!location) {
        vscode.window.showWarningMessage('XLIDE: this document is not a module of a project.');
        return;
    }

    const sources: Record<string, string> = {};
    try {
        const modules = await deps.bridge.call<Array<{ name: string }>>(
            'listModules', { path: location.projectPath },
        );
        for (const module of modules) {
            if (module.name.toLowerCase() === location.moduleName.toLowerCase()) { continue; }
            const read = await deps.bridge.call<{ source: string }>(
                'readModule', { path: location.projectPath, module: module.name },
            );
            sources[module.name] = read.source;
        }
    } catch (err) {
        vscode.window.showErrorMessage(
            `XLIDE: could not read the project's modules: ${err instanceof Error ? err.message : String(err)}`,
        );
        return;
    }

    const result = await compute(
        editor.document.getText(),
        editor,
        { moduleName: location.moduleName, sources },
    );
    if (!result) { return; }
    await applyResult(editor, result, deps, location.projectPath);
}

/**
 * Applies a result, or reports the refusal. Edits to other modules go through
 * the project engine rather than a workspace edit, since a project module has
 * no file of its own until it is opened.
 */
async function applyResult(
    editor: vscode.TextEditor,
    result: VbaRefactorResult,
    deps?: CommandDeps,
    projectPath?: string,
): Promise<void> {
    if (!result.ok) {
        vscode.window.showInformationMessage(`XLIDE: ${result.reason}`);
        return;
    }

    const edit = new vscode.WorkspaceEdit();
    for (const textEdit of result.edits) {
        edit.replace(
            editor.document.uri,
            new vscode.Range(
                editor.document.positionAt(textEdit.span.start),
                editor.document.positionAt(textEdit.span.end),
            ),
            textEdit.newText,
        );
    }
    if (!await vscode.workspace.applyEdit(edit)) {
        vscode.window.showErrorMessage('XLIDE: the refactoring could not be applied.');
        return;
    }

    if (result.otherModules?.length && deps && projectPath) {
        // All or nothing: a signature changed in one module and not repointed
        // in another leaves the project not compiling, so a failure here is
        // reported rather than swallowed.
        for (const module of result.otherModules) {
            try {
                const read = await deps.bridge.call<{ source: string }>(
                    'readModule', { path: projectPath, module: module.moduleName },
                );
                const updated = applyToText(read.source, module.edits);
                await deps.bridge.call('writeModule', {
                    path: projectPath, module: module.moduleName, source: updated,
                });
            } catch (err) {
                vscode.window.showErrorMessage(
                    `XLIDE: '${module.moduleName}' was not updated: `
                    + `${err instanceof Error ? err.message : String(err)}. `
                    + 'The other edits were applied; check the project before running it.',
                );
                return;
            }
        }
    }

    if (result.renameSpan) {
        const at = editor.document.positionAt(result.renameSpan.start);
        editor.selection = new vscode.Selection(
            at,
            editor.document.positionAt(result.renameSpan.end),
        );
        editor.revealRange(new vscode.Range(at, at));
    }
    statusMessage(`XLIDE: ${result.title}`);
}

function applyToText(source: string, edits: readonly { span: { start: number; end: number }; newText: string }[]): string {
    const ordered = [...edits].sort((a, b) => b.span.start - a.span.start);
    let out = source;
    for (const edit of ordered) {
        out = out.slice(0, edit.span.start) + edit.newText + out.slice(edit.span.end);
    }
    return out;
}

function vbaEditor(): vscode.TextEditor | undefined {
    const editor = vscode.window.activeTextEditor;
    if (!editor || !isVbaDocument(editor.document)) {
        vscode.window.showWarningMessage('XLIDE: open a VBA module first.');
        return undefined;
    }
    return editor;
}

function caretOffset(editor: vscode.TextEditor): number {
    return editor.document.offsetAt(editor.selection.active);
}

function selectionSpan(editor: vscode.TextEditor): { start: number; end: number } {
    return {
        start: editor.document.offsetAt(editor.selection.start),
        end: editor.document.offsetAt(editor.selection.end),
    };
}

/** Which interface, when the class implements more than one. */
async function pickInterface(source: string): Promise<{ interfaceName: string } | undefined> {
    const names = [...source.matchAll(/^[ \t]*Implements[ \t]+([\p{L}_][\p{L}\p{M}\p{N}_.]*)/gimu)]
        .map((match) => match[1]);
    if (names.length <= 1) {
        return undefined;
    }
    const picked = await vscode.window.showQuickPick(names, {
        title: 'Implement which interface?',
    });
    return picked ? { interfaceName: picked } : undefined;
}

async function pickModule(names: string[], exclude: string): Promise<string | undefined> {
    const choices = names.filter((name) => name.toLowerCase() !== exclude.toLowerCase()).sort();
    if (choices.length === 0) {
        vscode.window.showInformationMessage('XLIDE: the project has no other module to move this to.');
        return undefined;
    }
    return vscode.window.showQuickPick(choices, { title: 'Move the procedure to which module?' });
}
