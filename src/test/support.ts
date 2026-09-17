// Shared helpers for the integration suites, which run inside a real VS Code
// against the throwaway workspace .vscode-test.mjs builds: a copy of
// FormFixture.xlsm and a loose Loose.bas.

import * as assert from 'node:assert/strict';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { encodeModuleUri } from '../xlideFileSystem';
import { writeModule as engineWriteModule } from '../vba/projectService';

export const EXTENSION_ID = 'WilliamSmithE.xlide';

/** The workspace folder the runner opened. */
export function workspaceRoot(): string {
	const folder = vscode.workspace.workspaceFolders?.[0];
	assert.ok(folder, 'the runner should open the fixture workspace');
	return folder.uri.fsPath;
}

/** The disposable copy of the fixture workbook. */
export function workbookPath(): string {
	return path.join(workspaceRoot(), 'FormFixture.xlsm');
}

/** The virtual document of a module in the fixture workbook. */
export function moduleUri(moduleName: string): vscode.Uri {
	return encodeModuleUri(workbookPath(), moduleName);
}

/** Activates XLIDE, which the runner's `--disable-extensions` leaves alone. */
export async function activate(): Promise<void> {
	const extension = vscode.extensions.getExtension(EXTENSION_ID);
	assert.ok(extension, `${EXTENSION_ID} should be available in the development host`);
	if (!extension.isActive) {
		await extension.activate();
	}
}

/**
 * Writes a module into the workbook with the engine the extension itself
 * uses, creating the module when the workbook lacks it. The extension reads
 * the workbook by modification time, so it sees the new content at once.
 */
export async function writeModule(moduleName: string, source: string): Promise<vscode.Uri> {
	await activate();
	engineWriteModule(workbookPath(), moduleName, source, 'standard');
	return moduleUri(moduleName);
}

/** Opens a document and shows it, so providers that read the active editor see it. */
export async function open(uri: vscode.Uri): Promise<vscode.TextDocument> {
	const document = await vscode.workspace.openTextDocument(uri);
	await vscode.window.showTextDocument(document, { preview: false });
	return document;
}

/**
 * Opens a document without an editor. XLIDE's typing automation (the
 * canonical-casing controller among it) works on editors, so a test that
 * reads text and formats it against that text must not race one.
 */
export function openHidden(uri: vscode.Uri): Thenable<vscode.TextDocument> {
	return vscode.workspace.openTextDocument(uri);
}

/** Polls until the check holds, failing with a message that says what never came. */
export async function until<T>(
	produce: () => T | undefined | Promise<T | undefined>,
	message: string,
	timeoutMs = 20000,
): Promise<T> {
	const deadline = Date.now() + timeoutMs;
	for (;;) {
		const value = await produce();
		if (value !== undefined) {
			return value;
		}
		assert.ok(Date.now() < deadline, message);
		await new Promise((resolve) => setTimeout(resolve, 200));
	}
}

/**
 * The document's text after the edits, the way the editor would apply them.
 * VS Code hands back minimal edits - several per line, in any order - so
 * they are applied from the end of the text backwards by offset.
 */
export function applyEdits(text: string, edits: readonly vscode.TextEdit[]): string {
	const lineStarts = [0];
	for (let i = 0; i < text.length; i++) {
		if (text[i] === '\r') {
			if (text[i + 1] === '\n') {
				i++;
			}
			lineStarts.push(i + 1);
		} else if (text[i] === '\n') {
			lineStarts.push(i + 1);
		}
	}
	const offsetAt = (position: vscode.Position): number => lineStarts[position.line] + position.character;
	const sorted = [...edits].sort((a, b) => offsetAt(b.range.start) - offsetAt(a.range.start));
	let out = text;
	for (const edit of sorted) {
		out = out.slice(0, offsetAt(edit.range.start)) + edit.newText + out.slice(offsetAt(edit.range.end));
	}
	return out;
}

export function closeAllEditors(): Thenable<unknown> {
	return vscode.commands.executeCommand('workbench.action.closeAllEditors');
}
