// Shapes in a real VS Code: the tree's shape commands are there to be
// called, and a shape added with its look through the agent tool reads back
// through the listing tool, so the tool schemas, the engine's parameter
// checks and the Office write coordination are exercised together.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, workbookPath } from './support';

/** A tool's answer, as the agent reads it. */
async function invoke(name: string, input: object): Promise<string> {
	const result = await vscode.lm.invokeTool(name, { input, toolInvocationToken: undefined }, new vscode.CancellationTokenSource().token);
	return result.content.map((part) => (part instanceof vscode.LanguageModelTextPart ? part.value : '')).join('');
}

interface ListedShape {
	name: string;
	rotation?: number;
	hidden?: boolean;
	fill?: { type: string; color?: string; transparency?: number };
	line?: { type: string };
	font?: { bold?: boolean; size?: number };
	text?: string;
	zOrder?: number;
}

suite('Shapes', () => {
	suiteSetup(activate);

	test('registers the tree\'s shape commands', async () => {
		const commands = await vscode.commands.getCommands(true);
		for (const command of [
			'xlide.editShape', 'xlide.addShape', 'xlide.deleteShape',
			'xlide.linkShapeMacro', 'xlide.unlinkShapeMacro', 'xlide.goToShapeMacro',
		]) {
			assert.ok(commands.includes(command), `${command} should be registered`);
		}
	});

	test('adds a shape with its look through the agent tool, and lists it back', async () => {
		const filePath = workbookPath();
		await invoke('xlide_editShape', {
			filePath, surface: 'Sheet1', action: 'add', type: 'oval', name: 'IntegrationDot', range: 'B2:C4',
			text: 'Hi', rotation: 30, fill: { type: 'solid', color: '#FF0000', transparency: 25 },
			line: { type: 'none' }, font: { bold: true, size: 14 },
		});
		try {
			const surfaces = JSON.parse(await invoke('xlide_listShapes', { filePath, surface: 'Sheet1' })) as Array<{ codeName?: string; shapes: ListedShape[] }>;
			assert.equal(surfaces[0].codeName, 'Sheet1');
			const dot = surfaces[0].shapes.find((shape) => shape.name === 'IntegrationDot');
			assert.ok(dot, 'the added shape should be listed');
			assert.deepEqual(
				{ text: dot.text, rotation: dot.rotation, fill: dot.fill, line: dot.line, bold: dot.font?.bold, size: dot.font?.size, zOrder: dot.zOrder },
				{ text: 'Hi', rotation: 30, fill: { type: 'solid', color: '#FF0000', transparency: 25 }, line: { type: 'none' }, bold: true, size: 14, zOrder: 1 },
			);

			await invoke('xlide_editShape', { filePath, surface: 'Sheet1', action: 'update', name: 'IntegrationDot', hidden: true, font: { color: '' } });
			const again = JSON.parse(await invoke('xlide_listShapes', { filePath, surface: 'Sheet1' })) as Array<{ shapes: ListedShape[] }>;
			assert.equal(again[0].shapes.find((shape) => shape.name === 'IntegrationDot')?.hidden, true);
		} finally {
			await invoke('xlide_editShape', { filePath, surface: 'Sheet1', action: 'delete', name: 'IntegrationDot' });
		}
	});

	test('refuses a malformed look before it reaches the file', async () => {
		await assert.rejects(
			invoke('xlide_editShape', { filePath: workbookPath(), surface: 'Sheet1', action: 'add', type: 'oval', range: 'B2:C4', fill: { type: 'gradient' } }),
			/fill's type must be none or solid/,
		);
	});
});
