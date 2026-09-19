// Renaming a module that is open, in a real VS Code. The editor stayed on the
// old name, showing a module that no longer existed, and its text kept that
// module alive for analysis: a module still calling the old name showed no
// error.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { encodeFormMarkupUri } from '../xlideFileSystem';
import { listModules } from '../vba/projectService';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath } from './support';

async function agentTool(name: string, input: Record<string, unknown>): Promise<void> {
	await vscode.lm.invokeTool(name, {
		input: { filePath: workbookPath(), ...input },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

function tabsOn(moduleName: string): vscode.Tab[] {
	const uri = moduleUri(moduleName).toString();
	return vscode.window.tabGroups.all
		.flatMap((group) => group.tabs)
		.filter((tab) => tab.input instanceof vscode.TabInputText && tab.input.uri.toString() === uri);
}

suite('Renaming an open module', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('moves its editor to the new name, and a call to the old name is reported', async () => {
		await agentTool('xlide_writeModule', { moduleName: 'RenameMe', source: 'Option Explicit\r\n\r\nPublic Sub Proc()\r\nEnd Sub\r\n' });
		await agentTool('xlide_writeModule', { moduleName: 'RenameCaller', source: 'Option Explicit\r\n\r\nSub Run()\r\n    RenameMe.Proc\r\nEnd Sub\r\n' });
		const caller = await vscode.workspace.openTextDocument(moduleUri('RenameCaller'));
		await open(moduleUri('RenameMe'));
		const onTheCall = () => vscode.languages.getDiagnostics(caller.uri).filter((d) => d.range.start.line === 3);

		await agentTool('xlide_renameModule', { moduleName: 'RenameMe', newName: 'RenamedModule' });

		await until(
			() => (tabsOn('RenameMe').length === 0 && tabsOn('RenamedModule').length === 1 ? true : undefined),
			'the editor should follow the module to its new name',
			5000,
		);
		await until(
			() => (onTheCall().length > 0 ? true : undefined),
			'a call to the old name should be reported once the module is renamed',
			8000,
		);
	});

	test("moves a form's markup and designer to the new name, where the form still reads", async () => {
		// The engine left the designer under the old name, so the form could
		// not be read at all after a rename, and only the code editor followed.
		const form = listModules(workbookPath()).find((module) => module.type === 'userform');
		assert.ok(form, 'the fixture should have a UserForm');
		const oldMarkup = encodeFormMarkupUri(workbookPath(), form.name);
		const newMarkup = encodeFormMarkupUri(workbookPath(), 'RenamedPicker');
		const markup = await vscode.workspace.openTextDocument(oldMarkup);
		await vscode.languages.setTextDocumentLanguage(markup, 'xml');
		await vscode.window.showTextDocument(oldMarkup, { viewColumn: vscode.ViewColumn.One, preview: false });
		await vscode.commands.executeCommand('vscode.openWith', oldMarkup, 'xlideFormDesigner', vscode.ViewColumn.Two);
		const tabsAt = (uri: vscode.Uri) => vscode.window.tabGroups.all
			.flatMap((group) => group.tabs)
			.filter((tab) => (tab.input instanceof vscode.TabInputText || tab.input instanceof vscode.TabInputCustom)
				&& tab.input.uri.toString() === uri.toString());

		try {
			await agentTool('xlide_renameModule', { moduleName: form.name, newName: 'RenamedPicker' });

			await until(
				() => (tabsAt(oldMarkup).length === 0 && tabsAt(newMarkup).length === 2 ? true : undefined),
				'the markup and the designer should follow the form to its new name',
				8000,
			);
			const renamed = vscode.workspace.textDocuments.find((document) => document.uri.toString() === newMarkup.toString());
			assert.equal(renamed?.languageId, 'xml');
			assert.match(renamed?.getText() ?? '', /RenamedPicker/);
		} finally {
			await closeAllEditors();
			await agentTool('xlide_renameModule', { moduleName: 'RenamedPicker', newName: form.name });
		}
	});

	test('renames a form from the tree, and the code that uses it', async () => {
		// The tree offered no Rename for a form, and a rename elsewhere left
		// `Dim f As FrmPicker` and `FrmPicker.Show` naming a form that was gone.
		const form = listModules(workbookPath()).find((module) => module.type === 'userform');
		assert.ok(form, 'the fixture should have a UserForm');
		await agentTool('xlide_writeModule', {
			moduleName: 'FormCaller',
			source: [
				'Option Explicit',
				'',
				'Public Sub ShowPicker()',
				`    Dim f As ${form.name}`,
				`    Set f = New ${form.name}`,
				`    ${form.name}.Show`,
				`    Unload ${form.name}`,
				'End Sub',
				'',
			].join('\r\n'),
		});
		const window = vscode.window as { showInputBox: typeof vscode.window.showInputBox };
		const showInputBox = window.showInputBox;
		window.showInputBox = async () => 'ChooserForm';
		try {
			await vscode.commands.executeCommand('xlide.renameModule', {
				kind: 'module',
				label: form.name,
				filePath: workbookPath(),
				moduleName: form.name,
				moduleType: 'userform',
			});
		} finally {
			window.showInputBox = showInputBox;
		}

		try {
			assert.ok(listModules(workbookPath()).some((module) => module.name === 'ChooserForm' && module.type === 'userform'));
			const caller = await vscode.workspace.openTextDocument(moduleUri('FormCaller'));
			assert.deepEqual(caller.getText().split('\r\n').slice(3, 7), [
				'    Dim f As ChooserForm',
				'    Set f = New ChooserForm',
				'    ChooserForm.Show',
				'    Unload ChooserForm',
			]);
			assert.equal(await caller.save(), true);
		} finally {
			await closeAllEditors();
			await agentTool('xlide_renameModule', { moduleName: 'ChooserForm', newName: form.name });
		}
	});
});
