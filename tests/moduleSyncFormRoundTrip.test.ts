// Issue #36, end to end against the real engine: export a workbook with a
// form, plan the import straight back, and the form reads unchanged like its
// neighbours. FormFixture.xlsm was authored by live Excel (see
// vbaMacroContainers.test.ts), so the designer bytes are Office ground truth.
import { afterAll, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as svc from '../src/vba/workbookService';
import type { WorkbookEngine } from '../src/workbookEngine';
import { exportWorkbookModules } from '../src/moduleExport';
import { buildImportModuleSyncPlan } from '../src/moduleSyncPlan';

/** The engine's dispatch for the calls export and the sync plan make. */
function realBridge(): WorkbookEngine {
	return {
		async call<T>(method: string, p: Record<string, unknown>): Promise<T> {
			switch (method) {
				case 'listModules': return svc.listModules(String(p.path)) as T;
				case 'readModules': return svc.readModules(String(p.path), Boolean(p.full)) as T;
				case 'readModule': return svc.readModule(String(p.path), String(p.module), Boolean(p.full)) as T;
				case 'readFormExport': return svc.readFormExport(String(p.path), String(p.module)) as T;
				default: throw new Error(`test bridge: unexpected call ${method}`);
			}
		},
	} as WorkbookEngine;
}

const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-sync-form-'));
afterAll(() => {
	fs.rmSync(tempRoot, { recursive: true, force: true });
});

async function exportedFixture(): Promise<{ workbook: string; repo: string; frmPath: string }> {
	const dir = fs.mkdtempSync(path.join(tempRoot, 'case-'));
	const workbook = path.join(dir, 'FormFixture.xlsm');
	fs.copyFileSync(path.join(__dirname, 'fixtures', 'binaries', 'FormFixture.xlsm'), workbook);
	const repo = path.join(dir, 'repo');
	fs.mkdirSync(repo);
	await exportWorkbookModules(realBridge(), { filePath: workbook, exportFolder: repo });
	return { workbook, repo, frmPath: path.join(repo, 'FrmPicker.frm') };
}

describe('a form round-trips clean through export and import plan (issue #36)', () => {
	it('every row of a fresh round trip reads unchanged, the form included', async () => {
		const { workbook, repo } = await exportedFixture();
		const plan = await buildImportModuleSyncPlan(realBridge(), {
			workbookPath: workbook,
			importFolder: repo,
		});
		expect(plan.items.map((item) => `${item.relativeName}:${item.status}`).sort()).toEqual([
			'FrmPicker.frm:unchanged',
			'Sheet1.cls:unchanged',
			'ThisWorkbook.cls:unchanged',
			'XlideFormProbe.bas:unchanged',
		]);
	});

	it('a code edit in the repo .frm reads will-update', async () => {
		const { workbook, repo, frmPath } = await exportedFixture();
		fs.appendFileSync(frmPath, 'Public Sub Added()\r\nEnd Sub\r\n', 'utf8');
		const plan = await buildImportModuleSyncPlan(realBridge(), {
			workbookPath: workbook,
			importFolder: repo,
		});
		expect(plan.items.find((item) => item.relativeName === 'FrmPicker.frm')?.status)
			.toBe('will-update');
	});

	it('a designer-only edit reads unchanged: the code half is the compare (issue #36 scope)', async () => {
		// The accepted day-one trade: forms compare on the half the module
		// text can say. A designer-only repo change does not surface as a
		// pending row; the designer still travels whenever the row is applied.
		const { workbook, repo, frmPath } = await exportedFixture();
		const frm = fs.readFileSync(frmPath, 'utf8');
		const edited = frm.replace(/^(\s*Caption\s*=\s*).*$/m, '$1"Renamed"');
		expect(edited).not.toBe(frm);
		fs.writeFileSync(frmPath, edited, 'utf8');
		const plan = await buildImportModuleSyncPlan(realBridge(), {
			workbookPath: workbook,
			importFolder: repo,
		});
		expect(plan.items.find((item) => item.relativeName === 'FrmPicker.frm')?.status)
			.toBe('unchanged');
	});
});
