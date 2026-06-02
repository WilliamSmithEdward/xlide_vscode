import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	effectiveWorkbookModuleSyncSettings,
	setWorkbookModuleSyncExportMode,
	updateWorkbookModuleSyncSettings,
} from '../src/workbookModuleSyncSettings';
import {
	readWorkbookSettings,
	settingsPathForWorkbook,
	writeWorkbookSettings,
} from '../src/workbookSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-module-sync-settings-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	fs.writeFileSync(workbook, '', 'utf8');
	return { root, workbook };
}

describe('workbook module sync settings', () => {
	it('resolves missing workbook sync settings to built-in defaults', async () => {
		const { workbook } = tempWorkbook();

		await expect(effectiveWorkbookModuleSyncSettings(workbook)).resolves.toEqual({
			folderPath: undefined,
			folderPathSource: 'missing',
			exportMode: 'exportAll',
			exportModeSource: 'default',
			importMode: 'updateOnly',
			importModeSource: 'default',
			settingsPath: settingsPathForWorkbook(workbook),
		});
	});

	it('reports workbook sync overrides with workbook provenance', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'trueUpStandardClass',
		});

		await expect(effectiveWorkbookModuleSyncSettings(workbook)).resolves.toEqual({
			folderPath: 'C:/repo',
			folderPathSource: 'workbook',
			exportMode: 'trueUp',
			exportModeSource: 'workbook',
			importMode: 'trueUpStandardClass',
			importModeSource: 'workbook',
			settingsPath: settingsPathForWorkbook(workbook),
		});
	});

	it('updates only provided sync settings and preserves unrelated workbook settings', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(updateWorkbookModuleSyncSettings(workbook, { exportMode: 'trueUp' })).resolves.toMatchObject({
			folderPath: 'C:/repo',
			exportMode: 'trueUp',
			exportModeSource: 'workbook',
			importMode: 'updateOnly',
			importModeSource: 'default',
		});
		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});

	it('sets export mode through the module sync settings owner', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			exportMode: 'exportAll',
			importMode: 'trueUpStandardClass',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(setWorkbookModuleSyncExportMode(workbook, 'trueUp')).resolves.toMatchObject({
			folderPath: 'C:/repo',
			exportMode: 'trueUp',
			exportModeSource: 'workbook',
			importMode: 'trueUpStandardClass',
			importModeSource: 'workbook',
		});
		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'trueUpStandardClass',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});
});
