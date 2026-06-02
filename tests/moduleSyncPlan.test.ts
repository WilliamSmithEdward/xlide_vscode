import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { PythonBridge } from '../src/pythonBridge';
import {
	buildExportModuleSyncPlan,
	buildImportModuleSyncPlan,
	buildSideBySideDiff,
} from '../src/moduleSyncPlan';
import { writeWorkbookRepoConfig } from '../src/moduleExport';

interface FakeModule {
	name: string;
	type: string;
	source: string;
	documentType?: string;
}

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string; repo: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-sync-plan-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	const repo = path.join(root, 'repo');
	fs.writeFileSync(workbook, '', 'utf8');
	fs.mkdirSync(repo, { recursive: true });
	return { root, workbook, repo };
}

function fakeBridge(modules: readonly FakeModule[]): PythonBridge {
	return {
		async call<T>(method: string, args: Record<string, unknown>): Promise<T> {
			if (method === 'listModules') {
				return modules.map((mod) => ({
					name: mod.name,
					type: mod.type,
					documentType: mod.documentType,
				})) as T;
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

describe('module sync plan', () => {
	it('plans export row statuses from one workbook-vs-repo comparison path', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Existing.bas'), 'Sub Same()\nEnd Sub\n', 'utf8');
		fs.writeFileSync(path.join(repo, 'Changed.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');

		const plan = await buildExportModuleSyncPlan(fakeBridge([
			{ name: 'Existing', type: 'standard', source: 'Sub Same()\nEnd Sub\n' },
			{ name: 'Changed', type: 'standard', source: 'Sub Newer()\nEnd Sub\n' },
			{ name: 'NewModule', type: 'standard', source: 'Sub NewModule()\nEnd Sub\n' },
		]), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'replaceExistingOnly',
		});

		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('Existing')).toMatchObject({
			status: 'unchanged',
			checked: false,
			selectable: true,
		});
		expect(byName.get('Changed')).toMatchObject({
			status: 'will-write',
			checked: true,
			selectable: true,
		});
		expect(byName.get('NewModule')).toMatchObject({
			status: 'skipping-export',
			checked: false,
			selectable: false,
		});
		expect(byName.get('NewModule')?.warning).toContain('replaceExistingOnly');
	});

	it('surfaces true-up stale managed files as removable preview rows', async () => {
		const { workbook, repo } = tempWorkbook();
		fs.writeFileSync(path.join(repo, 'Stale.bas'), 'Sub Old()\nEnd Sub\n', 'utf8');
		await writeWorkbookRepoConfig(workbook, {
			exportFolder: repo,
			exportMode: 'trueUp',
			managedFiles: ['Stale.bas'],
		});

		const plan = await buildExportModuleSyncPlan(fakeBridge([
			{ name: 'Module1', type: 'standard', source: 'Sub T()\nEnd Sub\n' },
		]), {
			workbookPath: workbook,
			exportFolder: repo,
			exportMode: 'trueUp',
		});

		const stale = plan.items.find((item) => item.relativeName === 'Stale.bas');
		expect(stale).toMatchObject({
			moduleName: 'Stale',
			status: 'will-remove',
			checked: true,
			selectable: true,
			existsInWorkbook: false,
			existsInRepo: true,
		});
		expect(stale?.warning).toContain('no longer exists');
	});

	it('warns and skips missing document modules while allowing existing document updates', async () => {
		const { workbook, repo } = tempWorkbook();
		const sheetBase = 'Attribute VB_Base = "{00020820-0000-0000-C000-000000000046}"\n';
		fs.writeFileSync(path.join(repo, 'Sheet1.cls'), `${sheetBase}Private Sub Worksheet_Change()\nEnd Sub\n`, 'utf8');
		fs.writeFileSync(path.join(repo, 'Sheet2.cls'), `${sheetBase}Private Sub Worksheet_Activate()\nEnd Sub\n`, 'utf8');

		const plan = await buildImportModuleSyncPlan(fakeBridge([
			{
				name: 'Sheet1',
				type: 'document',
				documentType: 'worksheet',
				source: `${sheetBase}Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n`,
			},
		]), {
			workbookPath: workbook,
			importFolder: repo,
		});

		const byName = new Map(plan.items.map((item) => [item.moduleName, item]));
		expect(byName.get('Sheet1')).toMatchObject({
			status: 'will-update',
			checked: true,
			selectable: true,
			unsupportedDirectCreation: false,
		});
		expect(byName.get('Sheet1')?.warning).toContain('code can be updated');
		expect(byName.get('Sheet2')).toMatchObject({
			status: 'skipping-import',
			checked: false,
			selectable: true,
			unsupportedDirectCreation: true,
		});
		expect(byName.get('Sheet2')?.warning).toContain('cannot be created directly');
		expect(plan.warnings).toEqual([
			'Sheet2: skipping import unless the module already exists in the workbook.',
		]);
	});

	it('builds a side-by-side diff with changed line metadata', () => {
		const diff = buildSideBySideDiff('A\nB\nC', 'A\nX\nC');
		expect(diff.map((line) => line.kind)).toContain('changed');
		expect(diff.find((line) => line.kind === 'changed')).toMatchObject({
			left: 'B',
			right: 'X',
		});
	});
});
