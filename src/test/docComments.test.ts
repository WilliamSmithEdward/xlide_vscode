// The doc comment rules through live diagnostics in a real VS Code: the
// findings arrive as warnings, one of them inside the comment itself, and the
// quick fixes edit the comment where the findings say.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, applyEdits, closeAllEditors, open, until, writeModule } from './support';

suite('Doc comment diagnostics', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('reports a stale <param>, the parameter it lost and a missing <returns>, with their fixes', async () => {
		const uri = await writeModule('DocProbe', [
			'Option Explicit',
			'',
			"''' <summary>Calculates the total.</summary>",
			"''' <param name=\"Rate\">The tax rate.</param>",
			'Public Function Total(ByVal TaxRate As Double) As Double',
			'    Total = TaxRate',
			'End Function',
			'',
		].join('\r\n'));
		const document = await open(uri);

		const diagnostics = await until(() => {
			const found = vscode.languages.getDiagnostics(document.uri)
				.filter((d) => typeof d.code === 'string' && d.code.startsWith('doc-'));
			return found.length === 3 ? found : undefined;
		}, 'the three doc comment findings should be published');

		for (const diagnostic of diagnostics) {
			assert.equal(diagnostic.severity, vscode.DiagnosticSeverity.Warning, String(diagnostic.code));
		}
		const byCode = new Map(diagnostics.map((d) => [d.code as string, d]));
		assert.equal(document.getText(byCode.get('doc-param-unknown')!.range), 'Rate');
		assert.equal(document.getText(byCode.get('doc-param-missing')!.range), 'TaxRate');
		assert.equal(document.getText(byCode.get('doc-returns-missing')!.range), 'Total');

		const fix = async (code: string, title: string): Promise<string[]> => {
			const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>(
				'vscode.executeCodeActionProvider', document.uri, byCode.get(code)!.range,
			) ?? [];
			const action = actions.find((candidate) => candidate.title === title);
			assert.ok(action?.edit, `quick fixes offered for ${code}: ${actions.map((candidate) => candidate.title).join(', ')}`);
			return applyEdits(document.getText(), action.edit.get(document.uri)).split('\r\n');
		};
		assert.equal(
			(await fix('doc-param-unknown', "Rename the <param> to 'TaxRate'"))[3],
			"''' <param name=\"TaxRate\">The tax rate.</param>",
		);
		assert.deepEqual((await fix('doc-returns-missing', 'Add a <returns>')).slice(2, 6), [
			"''' <summary>Calculates the total.</summary>",
			"''' <param name=\"Rate\">The tax rate.</param>",
			"''' <returns></returns>",
			'Public Function Total(ByVal TaxRate As Double) As Double',
		]);
	});

	test('renaming a parameter renames its <param> too', async () => {
		const uri = await writeModule('DocRenameProbe', [
			'Option Explicit',
			'',
			"''' <summary>Adds an item.</summary>",
			"''' <param name=\"item\">The item.</param>",
			'Public Sub AddItem(ByVal item As String)',
			'    Debug.Print item',
			'End Sub',
			'',
		].join('\r\n'));
		const document = await open(uri);
		const at = document.positionAt(document.getText().indexOf('item As String'));
		const edit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>(
			'vscode.executeDocumentRenameProvider', document.uri, at, 'product',
		);
		assert.ok(edit, 'the rename should produce an edit');
		assert.deepEqual(applyEdits(document.getText(), edit.get(document.uri)).split('\r\n').slice(3, 6), [
			"''' <param name=\"product\">The item.</param>",
			'Public Sub AddItem(ByVal product As String)',
			'    Debug.Print product',
		]);
	});
});
