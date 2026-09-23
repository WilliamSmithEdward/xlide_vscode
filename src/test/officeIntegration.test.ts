// The Office integration surface, as a real VS Code sees it: the Open
// commands and the settings that decide what happens when a file is open in
// its application. Nothing here starts an Office application - the unit suite
// pins the scripts, and they were verified live against all four applications.

import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { activate, EXTENSION_ID } from './support';

suite('Office integration', () => {
	suiteSetup(activate);

	test('registers one Open pair for every application, and keeps the ids it had for Excel', async () => {
		const commands = new Set(await vscode.commands.getCommands(true));
		for (const id of [
			'xlide.openInOfficeApp',
			'xlide.openInOfficeAppReadOnly',
			// Bound to a key by someone, somewhere: they must keep working.
			'xlide.openWorkbook',
			'xlide.openWorkbookReadOnly',
		]) {
			assert.ok(commands.has(id), `${id} should be registered`);
		}
	});

	test('declares only the host-neutral Open commands', () => {
		const declared = (vscode.extensions.getExtension(EXTENSION_ID)?.packageJSON.contributes.commands as Array<{
			command: string;
			title: string;
		}>);
		const titles = new Map(declared.map((entry) => [entry.command, entry.title]));
		assert.equal(titles.get('xlide.openInOfficeApp'), 'Open in Office Application');
		assert.equal(titles.get('xlide.openInOfficeAppReadOnly'), 'Open in Office Application (Read Only)');
		assert.equal(titles.has('xlide.openWorkbook'), false);
		assert.deepEqual(
			declared.filter((entry) => /excel|workbook/i.test(entry.title)).map((entry) => entry.title),
			[],
			'no command title should name one application',
		);
	});

	test('contributes the settings under their host-neutral names, with the safe defaults', () => {
		const config = vscode.workspace.getConfiguration('xlide');
		assert.equal(config.inspect('officeIntegration.coordinationMode')?.defaultValue, 'block');
		assert.equal(config.inspect('officeIntegration.trackOpenedFiles')?.defaultValue, true);
		assert.equal(config.inspect('officeIntegration.reopenAfterClose')?.defaultValue, true);
		assert.equal(config.inspect('officeIntegration.reopenMode')?.defaultValue, 'lastState');
		// On since the refresh leaves any copy with unsaved work alone, keeps
		// the reader's place, stays behind the editor and runs no
		// Workbook_Open: a read-only Excel window follows each save.
		assert.equal(config.inspect('officeIntegration.reopenReadOnlyAfterSave')?.defaultValue, true);
		assert.equal(config.inspect('officeIntegration.attachToRunning')?.defaultValue, true);
	});

	test('still contributes the names the settings had for Excel, marked deprecated', () => {
		const properties = vscode.extensions.getExtension(EXTENSION_ID)?.packageJSON.contributes.configuration
			.properties as Record<string, { deprecationMessage?: string }>;
		for (const [legacy, renamed] of [
			['xlide.attachToRunningExcel', 'xlide.officeIntegration.attachToRunning'],
			['xlide.excelIntegration.coordinationMode', 'xlide.officeIntegration.coordinationMode'],
			['xlide.excelIntegration.trackOpenedWorkbooks', 'xlide.officeIntegration.trackOpenedFiles'],
		]) {
			assert.match(properties[legacy]?.deprecationMessage ?? '', new RegExp(`^Renamed to ${renamed.replace(/\./g, '\\.')}\\.`));
			assert.equal(properties[renamed]?.deprecationMessage, undefined);
		}
	});
});
