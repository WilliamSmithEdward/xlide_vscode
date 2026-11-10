// Issue #1622: completion and hover offer a referenced project's exposed
// class members, in a real VS Code. ProbeConsumerFixture.xlsm references
// ProbeLibraryFixture.xlsm (as ProbeLibrary) beside it; the test gives the
// library an exposed class, Widget, written with XLIDE's own engine.

import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { encodeModuleUri } from '../xlideFileSystem';
import { writeModule as engineWriteModule, listModules, listReferences } from '../vba/projectService';
import { referencedProjectFile } from '../vba/referencedProjectPath';
import { activate, closeAllEditors, until, workspaceRoot } from './support';

const FIXTURES = path.join(__dirname, '..', '..', 'tests', 'fixtures', 'binaries');

const WIDGET = [
	"'@Exposed",
	'Option Explicit',
	'Public Value As Long',
	'Public Child As Widget',
	'Private mSecret As Long',
	'Public Event Changed()',
	'Public Function Twice(ByVal n As Long) As Long',
	'    Twice = Value * 2 * n',
	'End Function',
	'Friend Sub Inner()',
	'End Sub',
	'Public Sub Poke(ByVal n As Long)',
	'End Sub',
	'',
].join('\r\n');

const labelOf = (item: vscode.CompletionItem): string => (typeof item.label === 'string' ? item.label : item.label.label);

function percentile(samples: number[], p: number): number {
	const sorted = [...samples].sort((a, b) => a - b);
	return sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))];
}

suite('A referenced class in completion and hover (issue #1622)', () => {
	let dir: string;
	let consumer: string;
	let resolved: string | undefined;

	suiteSetup(async () => {
		await activate();
		dir = path.join(workspaceRoot(), 'referenced-1622');
		fs.mkdirSync(dir, { recursive: true });
		consumer = path.join(dir, 'ProbeConsumerFixture.xlsm');
		fs.copyFileSync(path.join(FIXTURES, 'ProbeConsumerFixture.xlsm'), consumer);
		const library = path.join(dir, 'ProbeLibraryFixture.xlsm');
		fs.copyFileSync(path.join(FIXTURES, 'ProbeLibraryFixture.xlsm'), library);
		// The consumer also records the provider's absolute path where Excel saved
		// it, in the checkout the fixtures were made in, and the VBE and the engine
		// read that first. Where it exists, the copy beside the consumer is not
		// the one read, so the test cannot give it a class.
		const reference = listReferences(consumer).find((candidate) => candidate.name === 'ProbeLibrary');
		resolved = reference && referencedProjectFile(reference.libid, reference.relativeLibid, consumer);
		engineWriteModule(library, 'Widget', WIDGET, 'class');
		engineWriteModule(consumer, 'Module1', [
			'Option Explicit',
			'Public Sub Main()',
			'    Dim w As ProbeLibrary.Widget',
			'    Debug.Print w.Value',
			'    w.',
			'End Sub',
			'',
		].join('\r\n'));
	});

	suiteTeardown(async () => {
		await closeAllEditors();
		// On Windows the removal failed with EPERM in 2 of 4 full integration
		// runs on 2026-11-08, as it does while another process holds a file in
		// the folder open. rmSync retries that from Node 24.21, which VS Code's
		// extension host has.
		fs.rmSync(dir, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 });
	});

	test('offers Widget\'s Public members after w., and hovers one', async function () {
		if (!resolved || !listModules(resolved).some((module) => module.name === 'Widget')) {
			console.log(`[1622] skipped: the consumer reads ProbeLibrary from ${resolved}, which has no Widget`);
			this.skip();
		}
		const uri = encodeModuleUri(consumer, 'Module1');
		const document = await vscode.workspace.openTextDocument(uri);
		await vscode.window.showTextDocument(document, { preview: false });
		const caret = new vscode.Position(4, document.lineAt(4).text.length);
		const complete = async (): Promise<string[]> => {
			const list = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', uri, caret, '.');
			return (list?.items ?? []).map(labelOf);
		};
		const coldStart = performance.now();
		// Twice is nowhere in the module's text, so the editor's word list cannot supply it.
		const offered = await until(async () => {
			const labels = await complete();
			return labels.includes('Twice') ? labels : undefined;
		}, 'w. should offer Widget\'s members once the referenced project is read');
		const firstFull = performance.now() - coldStart;
		for (const member of ['Value', 'Child', 'Twice', 'Poke']) { assert.ok(offered.includes(member), `${member} in ${offered.join(',')}`); }
		for (const absent of ['mSecret', 'Inner', 'Changed']) { assert.ok(!offered.includes(absent), absent); }

		const warm: number[] = [];
		for (let i = 0; i < 30; i++) {
			const start = performance.now();
			const labels = await complete();
			warm.push(performance.now() - start);
			assert.ok(labels.includes('Value'));
		}
		console.log(`[1622] completion after w.: first full list ${firstFull.toFixed(1)} ms (includes the project read);`
			+ ` warm p50 ${percentile(warm, 50).toFixed(1)} ms, p95 ${percentile(warm, 95).toFixed(1)} ms, max ${Math.max(...warm).toFixed(1)} ms`);

		const hoverAt = new vscode.Position(3, document.lineAt(3).text.indexOf('Value') + 2);
		const hovers: number[] = [];
		let text = '';
		for (let i = 0; i < 20; i++) {
			const start = performance.now();
			const result = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider', uri, hoverAt);
			hovers.push(performance.now() - start);
			text = (result ?? []).flatMap((hover) => hover.contents.map((part) => (typeof part === 'string' ? part : part.value))).join('\n');
		}
		assert.match(text, /Value/, `hover should describe Value: ${text}`);
		console.log(`[1622] hover on w.Value: p50 ${percentile(hovers, 50).toFixed(1)} ms, p95 ${percentile(hovers, 95).toFixed(1)} ms, max ${Math.max(...hovers).toFixed(1)} ms`);
	});
});
