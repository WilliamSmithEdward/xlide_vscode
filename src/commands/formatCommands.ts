// Format All Modules: the document formatter over every module of a file.
//
// Each module goes through the same provider Shift+Alt+F uses, on its own
// document, so a module with unsaved edits is formatted where those edits
// are rather than overwritten on disk. Modules that were clean are saved
// afterwards; ones that were dirty keep the formatting and stay open for the
// user to save. The command counts first and asks before it writes.

import * as path from 'node:path';
import * as vscode from 'vscode';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import type { ModuleEntry } from '../vba/projectService';
import { moduleDocumentUri } from '../vbaDocumentLocation';
import { errorMessage } from '../util/errors';
import { resolveProjectPath, statusMessage, type CommandDeps } from './shared';

export interface FormatAllModulesOptions {
	/** Skip the confirmation; the integration suite has no one to click it. */
	withoutPrompt?: boolean;
}

interface PlannedModule {
	name: string;
	document: vscode.TextDocument;
	edits: vscode.TextEdit[];
	/** The document version the edits were computed against. */
	version: number;
	/** The document had unsaved edits before formatting: format in place, do not save. */
	wasDirty: boolean;
}

export function registerFormatCommands(deps: CommandDeps): vscode.Disposable[] {
	const { bridge, out } = deps;
	return [
		registerXlideCommand(
			'xlide.formatAllModules',
			async (node?: XlideNode, options: FormatAllModulesOptions = {}) => {
				const projectPath = resolveProjectPath(node);
				if (!projectPath) {
					void vscode.window.showInformationMessage(
						'XLIDE: Select a file in the XLIDE tree, or open one of its modules, to format its modules.',
					);
					return;
				}
				const name = path.basename(projectPath);
				const modules = (await bridge.call<ModuleEntry[]>('readModules', { path: projectPath }))
					.filter((module) => module.source !== undefined);
				const planned = await vscode.window.withProgress(
					{ location: vscode.ProgressLocation.Notification, title: `XLIDE: Checking ${modules.length} modules in ${name}`, cancellable: true },
					(progress, token) => planFormatting(projectPath, modules, progress, token),
				);
				if (!planned) {
					return;
				}
				if (planned.length === 0) {
					statusMessage(`XLIDE: All ${modules.length} modules in ${name} are already formatted`);
					return;
				}
				if (!options.withoutPrompt) {
					const choice = await vscode.window.showInformationMessage(
						`Format ${planned.length} of ${modules.length} modules in ${name}?`,
						{
							modal: true,
							detail: 'Each module is re-indented, its keywords are cased, and its spacing is normalized. '
								+ 'A module with unsaved edits is formatted in its editor and left for you to save.',
						},
						'Format',
					);
					if (choice !== 'Format') {
						return;
					}
				}
				const { formatted, leftUnsaved, failed } = await vscode.window.withProgress(
					{ location: vscode.ProgressLocation.Notification, title: `XLIDE: Formatting ${planned.length} modules in ${name}` },
					(progress) => applyFormatting(planned, progress, out),
				);
				const parts = [`XLIDE: Formatted ${formatted} of ${modules.length} modules in ${name}`];
				if (leftUnsaved > 0) {
					parts.push(`${leftUnsaved} left unsaved`);
				}
				if (failed > 0) {
					parts.push(`${failed} failed, see the XLIDE output`);
				}
				statusMessage(parts.join(', '));
			},
			{ errorPrefix: 'Format All Modules failed', logTag: 'format', log: (line) => out.appendLine(line) },
		),
	];
}

/** The modules the formatter would change, with their edits; undefined when cancelled. */
async function planFormatting(
	projectPath: string,
	modules: readonly ModuleEntry[],
	progress: vscode.Progress<{ message?: string; increment?: number }>,
	token: vscode.CancellationToken,
): Promise<PlannedModule[] | undefined> {
	const planned: PlannedModule[] = [];
	for (const module of modules) {
		if (token.isCancellationRequested) {
			return undefined;
		}
		progress.report({ message: module.name, increment: 100 / Math.max(1, modules.length) });
		const uri = moduleDocumentUri(projectPath, { moduleName: module.name, filePath: module.filePath });
		const document = await vscode.workspace.openTextDocument(uri);
		const edits = await formattingEdits(document);
		if (edits.length > 0) {
			planned.push({ name: module.name, document, edits, version: document.version, wasDirty: document.isDirty });
		}
	}
	return planned;
}

async function applyFormatting(
	planned: readonly PlannedModule[],
	progress: vscode.Progress<{ message?: string; increment?: number }>,
	out: vscode.OutputChannel,
): Promise<{ formatted: number; leftUnsaved: number; failed: number }> {
	let formatted = 0;
	let leftUnsaved = 0;
	let failed = 0;
	for (const module of planned) {
		progress.report({ message: module.name, increment: 100 / planned.length });
		try {
			const { document } = module;
			// The text may have moved on since the count (a keystroke, the
			// casing controller); the edits must fit the document as it is now.
			const edits = document.version === module.version
				? module.edits
				: await formattingEdits(document);
			if (edits.length === 0) {
				continue;
			}
			const edit = new vscode.WorkspaceEdit();
			edit.set(document.uri, edits);
			if (!(await vscode.workspace.applyEdit(edit))) {
				throw new Error('the editor refused the edit');
			}
			formatted++;
			if (module.wasDirty) {
				leftUnsaved++;
			} else if (!(await document.save())) {
				throw new Error('the document could not be saved');
			}
		} catch (err) {
			failed++;
			out.appendLine(`[format] ${module.name}: ${errorMessage(err)}`);
		}
	}
	return { formatted, leftUnsaved, failed };
}

/** The formatter's edits for a document, under the editor's own indentation settings for it. */
async function formattingEdits(document: vscode.TextDocument): Promise<vscode.TextEdit[]> {
	const editor = vscode.workspace.getConfiguration('editor', document);
	const options: vscode.FormattingOptions = {
		tabSize: editor.get<number>('tabSize', 4),
		insertSpaces: editor.get<boolean>('insertSpaces', true),
	};
	const edits = await vscode.commands.executeCommand<vscode.TextEdit[]>(
		'vscode.executeFormatDocumentProvider', document.uri, options,
	);
	return edits ?? [];
}
