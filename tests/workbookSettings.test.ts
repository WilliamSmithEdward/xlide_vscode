import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	isWorkbookSettingsError,
	readWorkbookSettings,
	resolveWorkbookSetting,
	settingsPathForWorkbook,
	updateWorkbookSettings,
	writeWorkbookSettings,
} from '../src/workbookSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-workbook-settings-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	fs.writeFileSync(workbook, '', 'utf8');
	return { root, workbook };
}

describe('workbookSettings', () => {
	it('resolves workbook overrides over global defaults with explicit source', () => {
		const fallback = { value: ['error'], source: 'default' as const };
		expect(resolveWorkbookSetting(['warning'], fallback)).toEqual({
			value: ['warning'],
			source: 'workbook',
		});
		expect(resolveWorkbookSetting(undefined, fallback)).toEqual({
			value: ['error'],
			source: 'default',
		});
	});

	it('treats a missing workbook settings sidecar as no workbook settings', async () => {
		const { workbook } = tempWorkbook();

		await expect(readWorkbookSettings(workbook)).resolves.toEqual({});
	});

	it('writes normalized workbook settings to the workbook sidecar', async () => {
		const { workbook } = tempWorkbook();

		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			analysis: {
				visibleSeverities: ['warning', 'warning'],
				untrackedRules: [' Option-Explicit-Missing ', 'option-explicit-missing'],
			},
		});

		expect(JSON.parse(fs.readFileSync(settingsPathForWorkbook(workbook), 'utf8'))).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			analysis: {
				visibleSeverities: ['warning', 'warning'],
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});

	it('updates workbook settings through one read-normalize-write owner', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder: 'C:/repo',
			exportMode: 'exportAll',
			analysis: {
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(updateWorkbookSettings(workbook, (existing) => ({
			...existing,
			importMode: 'trueUpStandardClass',
			analysis: {
				...existing.analysis,
				untrackedRules: [' Argument-Count ', 'option-explicit-missing'],
			},
		}))).resolves.toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'exportAll',
			importMode: 'trueUpStandardClass',
			analysis: {
				untrackedRules: ['argument-count', 'option-explicit-missing'],
			},
		});
	});

	it('can skip workbook settings writes when an update is unnecessary', async () => {
		const { workbook } = tempWorkbook();
		await writeWorkbookSettings(workbook, { exportFolder: 'C:/repo' });
		const before = fs.readFileSync(settingsPathForWorkbook(workbook), 'utf8');

		await expect(updateWorkbookSettings(workbook, () => undefined)).resolves.toEqual({
			exportFolder: 'C:/repo',
		});

		expect(fs.readFileSync(settingsPathForWorkbook(workbook), 'utf8')).toBe(before);
	});

	it('rejects invalid workbook settings JSON with the sidecar path', async () => {
		const { workbook } = tempWorkbook();
		fs.writeFileSync(settingsPathForWorkbook(workbook), '{ nope', 'utf8');

		await expect(readWorkbookSettings(workbook)).rejects.toMatchObject({
			settingsPath: settingsPathForWorkbook(workbook),
			name: 'WorkbookSettingsError',
		});
		await expect(readWorkbookSettings(workbook)).rejects.toThrow('Expected valid JSON');
	});

	it('rejects non-object workbook settings', async () => {
		const { workbook } = tempWorkbook();
		fs.writeFileSync(settingsPathForWorkbook(workbook), '[]', 'utf8');

		await expect(readWorkbookSettings(workbook)).rejects.toThrow('Expected the root value to be a JSON object');
	});

	it('rejects unknown workbook settings keys instead of ignoring them', async () => {
		const { workbook } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForWorkbook(workbook),
			`${JSON.stringify({ exportFolder: 'C:/repo', typoMode: true }, null, 2)}\n`,
			'utf8',
		);

		await expect(readWorkbookSettings(workbook)).rejects.toThrow('Unknown setting "typoMode"');
	});

	it('rejects invalid workbook sync modes from disk', async () => {
		const { workbook } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForWorkbook(workbook),
			`${JSON.stringify({ exportFolder: 'C:/repo', exportMode: 'anythingElse' }, null, 2)}\n`,
			'utf8',
		);

		await expect(readWorkbookSettings(workbook)).rejects.toThrow('Expected "exportMode" to be "exportAll" or "trueUp"');
	});

	it('rejects invalid workbook analysis settings from disk', async () => {
		const { workbook } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForWorkbook(workbook),
			`${JSON.stringify({
				analysis: {
					visibleSeverities: ['error', 'hint'],
				},
			}, null, 2)}\n`,
			'utf8',
		);

		try {
			await readWorkbookSettings(workbook);
			throw new Error('Expected readWorkbookSettings to reject');
		} catch (err) {
			expect(isWorkbookSettingsError(err)).toBe(true);
			expect(err).toMatchObject({
				settingsPath: settingsPathForWorkbook(workbook),
			});
			expect(err instanceof Error ? err.message : String(err))
				.toContain('Expected "analysis.visibleSeverities" entries');
		}
	});
});
