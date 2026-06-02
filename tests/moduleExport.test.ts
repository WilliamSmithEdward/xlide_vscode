import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { PythonBridge } from '../src/pythonBridge';
import {
	configPathForWorkbook,
	exportWorkbookModule,
	exportWorkbookModules,
	legacyConfigPathForWorkbook,
	readWorkbookRepoConfig,
	writeWorkbookRepoConfig,
} from '../src/moduleExport';

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

function fakeBridge(modules: readonly FakeModule[]): PythonBridge {
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
			throw new Error(`Unexpected bridge call ${method}`);
		},
	} as PythonBridge;
}

describe('moduleExport', () => {
	it('exports one module and writes only workbook sync settings to the sidecar', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookRepoConfig(workbook, {
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
		expect(await readWorkbookRepoConfig(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
		});
		expect(JSON.parse(fs.readFileSync(configPathForWorkbook(workbook), 'utf8'))).not.toHaveProperty('managedFiles');
	});

	it('normalizes legacy replaceExistingOnly mode to exportAll and creates new current-module files', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookRepoConfig(workbook, {
			exportFolder,
			exportMode: 'replaceExistingOnly',
		});
		const bridge = fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]);

		const result = await exportWorkbookModule(bridge, {
			filePath: workbook,
			moduleName: 'Module1',
		});

		expect(result).toMatchObject({
			exportMode: 'exportAll',
			relativeName: 'Module1.bas',
			written: true,
			writtenFiles: ['Module1.bas'],
		});
		expect(fs.existsSync(path.join(exportFolder, 'Module1.bas'))).toBe(true);
		expect(await readWorkbookRepoConfig(workbook)).toMatchObject({
			exportFolder,
			exportMode: 'exportAll',
		});
		expect(JSON.parse(fs.readFileSync(configPathForWorkbook(workbook), 'utf8'))).not.toHaveProperty('managedFiles');
	});

	it('keeps all-module true-up behavior on the shared module-file writer', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		fs.mkdirSync(exportFolder, { recursive: true });
		fs.writeFileSync(path.join(exportFolder, 'Stale.bas'), 'old', 'utf8');
		await writeWorkbookRepoConfig(workbook, {
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
		expect(JSON.parse(fs.readFileSync(configPathForWorkbook(workbook), 'utf8'))).not.toHaveProperty('managedFiles');
		expect(configPathForWorkbook(workbook)).toBe(path.join(path.dirname(workbook), 'Book.xlsm.repo.json'));
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
		await writeWorkbookRepoConfig(workbook, { exportFolder, exportMode: 'trueUp' });

		const result = await exportWorkbookModules(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), { filePath: workbook });

		expect(result.removedFiles).toEqual(['Stale.bas', 'StaleClass.cls']);
		expect(fs.existsSync(path.join(exportFolder, 'Stale.bas'))).toBe(false);
		expect(fs.existsSync(path.join(exportFolder, 'StaleClass.cls'))).toBe(false);
		expect(fs.existsSync(path.join(exportFolder, 'Notes.txt'))).toBe(true);
		expect(fs.existsSync(path.join(exportFolder, 'UserForm1.frm'))).toBe(true);
		expect(fs.existsSync(path.join(exportFolder, 'nested', 'StaleClass.cls'))).toBe(true);
		expect(JSON.parse(fs.readFileSync(configPathForWorkbook(workbook), 'utf8'))).not.toHaveProperty('managedFiles');
	});

	it('reads legacy sidecars and writes the new sidecar name on update', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		fs.writeFileSync(
			legacyConfigPathForWorkbook(workbook),
			`${JSON.stringify({
				exportFolder,
				exportMode: 'replaceExistingOnly',
				managedFiles: ['Legacy.bas'],
			}, null, 2)}\n`,
			'utf8',
		);

		expect(await readWorkbookRepoConfig(workbook)).toMatchObject({
			exportFolder,
			exportMode: 'exportAll',
		});
		expect(await readWorkbookRepoConfig(workbook)).not.toHaveProperty('managedFiles');

		await writeWorkbookRepoConfig(workbook, {
			exportFolder,
			exportMode: 'trueUp',
		});

		expect(fs.existsSync(configPathForWorkbook(workbook))).toBe(true);
		expect(await readWorkbookRepoConfig(workbook)).toMatchObject({
			exportFolder,
			exportMode: 'trueUp',
		});
		expect(JSON.parse(fs.readFileSync(configPathForWorkbook(workbook), 'utf8'))).not.toHaveProperty('managedFiles');
	});
});
