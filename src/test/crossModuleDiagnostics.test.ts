// Diagnostics across modules in a real VS Code. A module's findings depend on
// the others - `HelperMod.Greet` is only defined while HelperMod exists - so
// when an agent writes one module and then creates or edits the one it calls,
// the first module's "not defined" must clear by itself. It used to stay until
// something unrelated touched that module's editor. The writes go through the
// agent's own tool, which is the path an agent takes.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath } from './support';

/** Writes a module the way an agent does: through XLIDE's write tool. */
async function agentWrite(moduleName: string, source: string): Promise<void> {
	await vscode.lm.invokeTool('xlide_writeModule', {
		input: { filePath: workbookPath(), moduleName, source },
		toolInvocationToken: undefined,
	}, new vscode.CancellationTokenSource().token);
}

/** The document's current diagnostics that mention `text`. */
function mentioning(uri: vscode.Uri, text: string): vscode.Diagnostic[] {
	return vscode.languages.getDiagnostics(uri).filter((diagnostic) => diagnostic.message.includes(text));
}

suite('Diagnostics across modules', () => {
	suiteSetup(async () => {
		await activate();
	});

	teardown(async () => {
		await closeAllEditors();
	});

	test('clear in the calling module once an agent creates the module it calls', async () => {
		await agentWrite('CallerMod', [
			'Option Explicit',
			'',
			'Public Sub Run()',
			'    HelperMod.Greet "x"',
			'End Sub',
			'',
		].join('\r\n'));
		const caller = await open(moduleUri('CallerMod'));
		await until(
			() => (mentioning(caller.uri, 'HelperMod').length > 0 ? true : undefined),
			'a call into a module that does not exist yet should be reported first',
		);

		const created = Date.now();
		await agentWrite('HelperMod', [
			'Option Explicit',
			'',
			'Public Sub Greet(ByVal who As String)',
			'    Debug.Print who',
			'End Sub',
			'',
		].join('\r\n'));
		await until(
			() => (mentioning(caller.uri, 'HelperMod').length === 0 ? true : undefined),
			'CallerMod should stop reporting HelperMod once the agent has created it',
			8000,
		);
		console.log(`    (cleared ${Date.now() - created} ms after HelperMod was created)`);
	});

	test('clear in the calling module once an agent adds the procedure to another module', async () => {
		await agentWrite('HelperTwo', 'Option Explicit\r\n\r\nPublic Sub Other()\r\nEnd Sub\r\n');
		await agentWrite('CallerTwo', [
			'Option Explicit',
			'',
			'Public Sub Run()',
			'    GreetTwo "x"',
			'End Sub',
			'',
		].join('\r\n'));
		const caller = await open(moduleUri('CallerTwo'));
		await until(
			() => (mentioning(caller.uri, 'GreetTwo').length > 0 ? true : undefined),
			'a call to a procedure no module defines should be reported first',
		);

		const edited = Date.now();
		await agentWrite('HelperTwo', [
			'Option Explicit',
			'',
			'Public Sub Other()',
			'End Sub',
			'',
			'Public Sub GreetTwo(ByVal who As String)',
			'    Debug.Print who',
			'End Sub',
			'',
		].join('\r\n'));
		await until(
			() => (mentioning(caller.uri, 'GreetTwo').length === 0 ? true : undefined),
			'CallerTwo should stop reporting GreetTwo once the agent has added it to HelperTwo',
			8000,
		);
		console.log(`    (cleared ${Date.now() - edited} ms after HelperTwo was edited)`);
		assert.equal(mentioning(caller.uri, 'GreetTwo').length, 0);
	});
});
