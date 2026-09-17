// The dead-code rules through live diagnostics in a real VS Code: the
// findings arrive as Information with the Unnecessary tag, and the quick fix
// the rule attaches is offered where the finding is.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Dead code diagnostics', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('reports an unused variable, an uncalled Private Sub and unreachable code as faded information', async () => {
		const uri = await writeModule('DeadProbe', [
			'Option Explicit',
			'',
			'Private Sub Helper()',
			'End Sub',
			'',
			'Public Sub Main()',
			'    Dim unused As Long',
			'    Exit Sub',
			'    Debug.Print 1',
			'End Sub',
			'',
		].join('\r\n'));
		const document = await open(uri);

		const diagnostics = await until(() => {
			const found = vscode.languages.getDiagnostics(document.uri)
				.filter((d) => typeof d.code === 'string' && ['unused-variable', 'unused-procedure', 'unreachable-code'].includes(d.code));
			return found.length === 3 ? found : undefined;
		}, 'the three dead-code findings should be published');

		for (const diagnostic of diagnostics) {
			assert.equal(diagnostic.severity, vscode.DiagnosticSeverity.Information, String(diagnostic.code));
			assert.deepEqual(diagnostic.tags, [vscode.DiagnosticTag.Unnecessary], String(diagnostic.code));
		}
		const byCode = new Map(diagnostics.map((d) => [d.code as string, d]));
		assert.equal(document.getText(byCode.get('unused-variable')!.range), 'unused');
		assert.equal(document.getText(byCode.get('unused-procedure')!.range), 'Helper');
		assert.equal(document.getText(byCode.get('unreachable-code')!.range), 'Debug.Print 1');

		const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>(
			'vscode.executeCodeActionProvider', document.uri, byCode.get('unused-variable')!.range,
		) ?? [];
		assert.ok(
			actions.some((action) => action.title === "Remove unused declaration of 'unused'"),
			`quick fixes offered: ${actions.map((action) => action.title).join(', ')}`,
		);
	});
});
