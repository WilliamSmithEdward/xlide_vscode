import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const mockConfig = vi.hoisted(() => ({
	visibleSeverities: ['error', 'warning', 'information'] as string[],
	untrackedRules: [] as string[],
	machineKeys: new Set<string>(),
}));

vi.mock('vscode', () => ({
	workspace: {
		getConfiguration: () => ({
			get: (key: string, fallback: unknown) => {
				if (key === 'analysis.visibleSeverities') {
					return mockConfig.visibleSeverities;
				}
				if (key === 'analysis.untrackedRules') {
					return mockConfig.untrackedRules;
				}
				return fallback;
			},
			inspect: (key: string) => mockConfig.machineKeys.has(key)
				? { globalValue: key === 'analysis.visibleSeverities'
					? mockConfig.visibleSeverities
					: mockConfig.untrackedRules }
				: {},
		}),
	},
}));

import { readWorkbookSettings, writeWorkbookSettings } from '../src/workbookSettings';
import {
	effectiveWorkbookAnalysisSettings,
	resetWorkbookAnalysisRuleTracking,
	resetWorkbookAnalysisSettings,
	resetWorkbookAnalysisVisibleSeverities,
	setWorkbookAnalysisRuleTracked,
	setWorkbookAnalysisVisibleSeverities,
} from '../src/workbookAnalysisSettings';

const tempRoots: string[] = [];

beforeEach(() => {
	mockConfig.visibleSeverities = ['error', 'warning', 'information'];
	mockConfig.untrackedRules = [];
	mockConfig.machineKeys.clear();
});

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; workbook: string; exportFolder: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-analysis-settings-'));
	tempRoots.push(root);
	const workbook = path.join(root, 'Book.xlsm');
	const exportFolder = path.join(root, 'repo');
	fs.writeFileSync(workbook, '', 'utf8');
	return { root, workbook, exportFolder };
}

describe('workbook analysis settings', () => {
	it('uses global defaults until a workbook override exists', async () => {
		const { workbook } = tempWorkbook();
		mockConfig.visibleSeverities = ['error', 'information'];
		mockConfig.untrackedRules = ['option-explicit-missing'];
		mockConfig.machineKeys.add('analysis.visibleSeverities');
		mockConfig.machineKeys.add('analysis.untrackedRules');

		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['error', 'information'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['option-explicit-missing'],
			untrackedRulesSource: 'machine',
		});
	});

	it('stores workbook severity overrides without disturbing sync settings', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
			importMode: 'updateOnly',
		});

		await setWorkbookAnalysisVisibleSeverities(workbook, ['warning']);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			analysis: {
				visibleSeverities: ['warning'],
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['warning'],
			visibleSeveritiesSource: 'workbook',
			untrackedRulesSource: 'default',
		});
	});

	it('starts rule tracking overrides from the effective global default', async () => {
		const { workbook } = tempWorkbook();
		mockConfig.untrackedRules = ['argument-count'];
		mockConfig.machineKeys.add('analysis.untrackedRules');

		const update = await setWorkbookAnalysisRuleTracked(workbook, 'Option-Explicit-Missing', false);

		expect(update).toMatchObject({
			tracked: false,
			changed: true,
			untrackedRules: ['argument-count', 'option-explicit-missing'],
		});
		expect((await readWorkbookSettings(workbook)).analysis?.untrackedRules)
			.toEqual(['argument-count', 'option-explicit-missing']);
	});

	it('resets one workbook analysis override while preserving the other', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		mockConfig.visibleSeverities = ['error'];
		mockConfig.machineKeys.add('analysis.visibleSeverities');
		await writeWorkbookSettings(workbook, {
			exportFolder,
			analysis: {
				visibleSeverities: ['warning'],
				untrackedRules: ['argument-count'],
			},
		});

		await resetWorkbookAnalysisVisibleSeverities(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			analysis: {
				untrackedRules: ['argument-count'],
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['error'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['argument-count'],
			untrackedRulesSource: 'workbook',
		});
	});

	it('resets rule tracking and removes empty analysis settings', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		mockConfig.untrackedRules = ['option-explicit-missing'];
		mockConfig.machineKeys.add('analysis.untrackedRules');
		await writeWorkbookSettings(workbook, {
			exportFolder,
			importMode: 'trueUpStandardClass',
			analysis: {
				untrackedRules: ['argument-count'],
			},
		});

		await resetWorkbookAnalysisRuleTracking(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			importMode: 'trueUpStandardClass',
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			untrackedRules: ['option-explicit-missing'],
			untrackedRulesSource: 'machine',
		});
	});

	it('resets all workbook analysis overrides without disturbing workbook-only settings', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
			analysis: {
				visibleSeverities: ['warning'],
				untrackedRules: ['argument-count'],
			},
		});

		await resetWorkbookAnalysisSettings(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeveritiesSource: 'default',
			untrackedRulesSource: 'default',
		});
	});
});
