import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { readWorkbookSettings, writeWorkbookSettings } from '../src/workbookSettings';
import {
	effectiveWorkbookTestSettings,
	effectiveWorkbookTestSettingsFromConfig,
	updateWorkbookTestSettings,
} from '../src/workbookTestSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-workbook-test-settings-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	fs.writeFileSync(workbook, '', 'utf8');
	return { root, workbook };
}

describe('workbook test settings', () => {
	it('uses default artifact settings until workbook overrides exist', async () => {
		const { workbook } = tempWorkbook();

		await expect(effectiveWorkbookTestSettings(workbook)).resolves.toMatchObject({
			artifactFolder: 'tests',
			artifactFolderSource: 'default',
			artifactRetention: 20,
			artifactRetentionSource: 'default',
		});
	});

	it('resolves artifact settings from already-loaded workbook config', () => {
		expect(effectiveWorkbookTestSettingsFromConfig('Book.xlsm', {
			tests: {
				artifactFolder: 'ci-artifacts',
				artifactRetention: 5,
			},
		})).toMatchObject({
			artifactFolder: 'ci-artifacts',
			artifactFolderSource: 'workbook',
			artifactRetention: 5,
			artifactRetentionSource: 'workbook',
		});
	});

	it('updates test artifact settings without disturbing sync or analysis settings', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			analysis: {
				untrackedRules: ['argument-count'],
			},
		});

		await expect(updateWorkbookTestSettings(workbook, {
			artifactFolder: 'ci-artifacts',
			artifactRetention: 8,
		})).resolves.toMatchObject({
			artifactFolder: 'ci-artifacts',
			artifactFolderSource: 'workbook',
			artifactRetention: 8,
			artifactRetentionSource: 'workbook',
		});

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			analysis: {
				untrackedRules: ['argument-count'],
			},
			tests: {
				artifactFolder: 'ci-artifacts',
				artifactRetention: 8,
			},
		});
	});
});
