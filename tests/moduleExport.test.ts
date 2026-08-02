import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { WorkbookEngine } from '../src/workbookEngine';
import { WorkbookEngineError, JSONRPC_METHOD_NOT_FOUND } from '../src/workbookEngineErrors';
import {
	exportWorkbookModule,
	exportWorkbookModules,
} from '../src/moduleExport';
import {
	readWorkbookSettings,
	settingsPathForWorkbook,
	writeWorkbookSettings,
} from '../src/workbookSettings';

interface FakeModule {
	name: string;
	type: string;
	source: string;
}

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string; exportFolder: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-export-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	const exportFolder = path.join(root, 'repo');
	fs.writeFileSync(workbook, '', 'utf8');
	return { root, workbook, exportFolder };
}

function fakeBridge(modules: readonly FakeModule[]): WorkbookEngine {
	return {
		async call<T>(method: string, args: Record<string, unknown>): Promise<T> {
			if (method === 'listModules') {
				return modules.map((mod) => ({ name: mod.name, type: mod.type })) as T;
			}
			if (method === 'readModule') {
				const moduleName = String(args.module ?? '').toLowerCase();
				const mod = modules.find((candidate) => candidate.name.toLowerCase() === moduleName);
				if (!mod) {
					throw new Error(`Unknown module ${String(args.module)}`);
				}
				return { source: mod.source } as T;
			}
			throw new WorkbookEngineError(`Method not found: ${method}`, JSONRPC_METHOD_NOT_FOUND);
		},
	} as WorkbookEngine;
}

describe('moduleExport', () => {
	it('exports one module and writes only workbook sync settings to the sidecar', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
		});
		const bridge = fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Attribute VB_Name = "Module1"\nSub T()\nEnd Sub\n' },
			{ name: 'Other', type: 'standard', source: 'Sub Other()\nEnd Sub\n' },
		]);

		const result = await exportWorkbookModule(bridge, {
			filePath: workbook,
			moduleName: 'module1',
		});

		expect(result).toMatchObject({
			moduleName: 'Module1',
			moduleType: 'standard',
			relativeName: 'Module1.bas',
			written: true,
			writtenFiles: ['Module1.bas'],
		});
		expect(fs.readFileSync(path.join(exportFolder, 'Module1.bas'), 'utf8')).toBe(
			'Attribute VB_Name = "Module1"\nSub T()\nEnd Sub\n',
		);
		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
		});
		expect(settingsPathForWorkbook(workbook)).toBe(path.join(path.dirname(workbook), 'Book.xlsm.xlide_settings.json'));
	});

	it('keeps all-module true-up behavior on the shared module-file writer', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		fs.mkdirSync(exportFolder, { recursive: true });
		fs.writeFileSync(path.join(exportFolder, 'Stale.bas'), 'old', 'utf8');
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
		});
		const bridge = fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
			{ name: 'Person', type: 'class', source: 'VERSION 1.0 CLASS\n' },
		]);

		const result = await exportWorkbookModules(bridge, { filePath: workbook });

		expect(result).toMatchObject({
			writtenCount: 2,
			removedCount: 1,
			writtenFiles: ['Module1.bas', 'Person.cls'],
			removedFiles: ['Stale.bas'],
		});
		expect(fs.existsSync(path.join(exportFolder, 'Stale.bas'))).toBe(false);
		expect(fs.existsSync(path.join(exportFolder, 'Module1.bas'))).toBe(true);
		expect(fs.existsSync(path.join(exportFolder, 'Person.cls'))).toBe(true);
		expect(settingsPathForWorkbook(workbook)).toBe(path.join(path.dirname(workbook), 'Book.xlsm.xlide_settings.json'));
	});

	it('only deletes root bas/cls module files during true-up', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		fs.mkdirSync(exportFolder, { recursive: true });
		fs.writeFileSync(path.join(exportFolder, 'Stale.bas'), 'old', 'utf8');
		fs.writeFileSync(path.join(exportFolder, 'StaleClass.cls'), 'old', 'utf8');
		fs.writeFileSync(path.join(exportFolder, 'Notes.txt'), 'keep', 'utf8');
		fs.writeFileSync(path.join(exportFolder, 'UserForm1.frm'), 'keep', 'utf8');
		fs.mkdirSync(path.join(exportFolder, 'nested'));
		fs.writeFileSync(path.join(exportFolder, 'nested', 'StaleClass.cls'), 'keep', 'utf8');
		await writeWorkbookSettings(workbook, { exportFolder, exportMode: 'trueUp' });

		const result = await exportWorkbookModules(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), { filePath: workbook });

		expect(result.removedFiles).toEqual(['Stale.bas', 'StaleClass.cls']);
		expect(fs.existsSync(path.join(exportFolder, 'Stale.bas'))).toBe(false);
		expect(fs.existsSync(path.join(exportFolder, 'StaleClass.cls'))).toBe(false);
		expect(fs.existsSync(path.join(exportFolder, 'Notes.txt'))).toBe(true);
		expect(fs.existsSync(path.join(exportFolder, 'UserForm1.frm'))).toBe(true);
		expect(fs.existsSync(path.join(exportFolder, 'nested', 'StaleClass.cls'))).toBe(true);
	});

	it('preserves workbook analysis overrides when exporting modules', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'exportAll',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await exportWorkbookModules(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), { filePath: workbook, exportMode: 'trueUp' });

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});

});
