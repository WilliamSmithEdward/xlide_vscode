// Deleting a form from the tree, in a real VS Code. The tree offered no Delete
// for a form; it does now, and the form's designer goes with it, as do the
// editors on its markup and its designer.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as vscode from 'vscode';
import { encodeFormMarkupUri } from '../xlideFileSystem';
import { addFormModule, listModules } from '../vba/projectService';
import { openMacroContainer } from '../vba/macroContainer';
import { activate, closeAllEditors, until, workbookPath } from './support';

suite('Deleting a form', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('removes the form and its designer, and closes its markup and designer', async () => {
		addFormModule(workbookPath(), 'DoomedForm');
		const markup = encodeFormMarkupUri(workbookPath(), 'DoomedForm');
		const document = await vscode.workspace.openTextDocument(markup);
		await vscode.languages.setTextDocumentLanguage(document, 'xml');
		await vscode.window.showTextDocument(markup, { viewColumn: vscode.ViewColumn.One, preview: false });
		await vscode.commands.executeCommand('vscode.openWith', markup, 'xlideFormDesigner', vscode.ViewColumn.Two);
		const tabsOnMarkup = () => vscode.window.tabGroups.all
			.flatMap((group) => group.tabs)
			.filter((tab) => (tab.input instanceof vscode.TabInputText || tab.input instanceof vscode.TabInputCustom)
				&& tab.input.uri.toString() === markup.toString());
		assert.equal(tabsOnMarkup().length, 2, 'the markup and the designer should be open');

		const window = vscode.window as { showWarningMessage: typeof vscode.window.showWarningMessage };
		const showWarningMessage = window.showWarningMessage;
		window.showWarningMessage = (async () => 'Delete') as typeof vscode.window.showWarningMessage;
		try {
			await vscode.commands.executeCommand('xlide.deleteModule', {
				kind: 'module',
				label: 'DoomedForm',
				filePath: workbookPath(),
				moduleName: 'DoomedForm',
				moduleType: 'userform',
			});
		} finally {
			window.showWarningMessage = showWarningMessage;
		}

		assert.ok(!listModules(workbookPath()).some((module) => module.name === 'DoomedForm'));
		const cfb = openMacroContainer(fs.readFileSync(workbookPath())).vbaCfb();
		assert.equal(cfb.hasStoragePath(['DoomedForm']), false, 'the designer should go with the form');
		await until(
			() => (tabsOnMarkup().length === 0 ? true : undefined),
			'the markup and the designer should close',
			5000,
		);
	});
});
