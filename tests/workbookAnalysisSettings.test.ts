import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const mockConfig = vi.hoisted(() => ({
	visibleSeverities: ['error', 'warning', 'information'] as string[],
	untrackedRules: [] as string[],
	ruleSeverityOverrides: {} as Record<string, string>,
	machineKeys: new Set<string>(),
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
	workspace: {
		getConfiguration: () => ({
			get: (key: string, fallback: unknown) => {
				if (key === 'analysis.visibleSeverities') {
					return mockConfig.visibleSeverities;
				}
				if (key === 'analysis.untrackedRules') {
					return mockConfig.untrackedRules;
				}
				if (key === 'analysis.ruleSeverityOverrides') {
					return mockConfig.ruleSeverityOverrides;
				}
				return fallback;
			},
			inspect: (key: string) => mockConfig.machineKeys.has(key)
				? { globalValue: key === 'analysis.visibleSeverities'
					? mockConfig.visibleSeverities
					: key === 'analysis.untrackedRules'
						? mockConfig.untrackedRules
						: mockConfig.ruleSeverityOverrides }
				: {},
		}),
	},
}));

import { readWorkbookSettings, writeWorkbookSettings } from '../src/workbookSettings';
import {
	effectiveWorkbookAnalysisSettings,
	effectiveWorkbookAnalysisSettingsFromConfig,
	resetWorkbookAnalysisRuleTracking,
	setWorkbookAnalysisRuleTracked,
} from '../src/workbookAnalysisSettings';

const tempRoots: string[] = [];

beforeEach(() => {
	mockConfig.visibleSeverities = ['error', 'warning', 'information'];
	mockConfig.untrackedRules = [];
	mockConfig.ruleSeverityOverrides = {};
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
		mockConfig.ruleSeverityOverrides = { 'unknown-call': 'warning' };
		mockConfig.machineKeys.add('analysis.visibleSeverities');
		mockConfig.machineKeys.add('analysis.untrackedRules');
		mockConfig.machineKeys.add('analysis.ruleSeverityOverrides');

		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['error', 'information'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['option-explicit-missing'],
			untrackedRulesSource: 'machine',
			workbookUntrackedRules: [],
			ruleSeverityOverrides: { 'unknown-call': 'warning' },
			ruleSeverityOverridesSource: 'machine',
		});
	});

	it('resolves effective settings from an already-loaded workbook config', () => {
		mockConfig.visibleSeverities = ['error'];
		mockConfig.machineKeys.add('analysis.visibleSeverities');

		expect(effectiveWorkbookAnalysisSettingsFromConfig('Book.xlsm', {
			analysis: {
				untrackedRules: ['argument-count'],
			},
		})).toMatchObject({
			visibleSeverities: ['error'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['argument-count'],
			untrackedRulesSource: 'workbook',
			workbookUntrackedRules: ['argument-count'],
			ruleSeverityOverridesSource: 'default',
		});
	});

	it('stores workbook rule tracking separately from the effective global default', async () => {
		const { workbook } = tempWorkbook();
		mockConfig.untrackedRules = ['argument-count'];
		mockConfig.machineKeys.add('analysis.untrackedRules');

		const update = await setWorkbookAnalysisRuleTracked(workbook, 'Option-Explicit-Missing', false);

		expect(update).toMatchObject({
			tracked: false,
			changed: true,
			untrackedRules: ['option-explicit-missing'],
		});
		expect((await readWorkbookSettings(workbook)).analysis?.untrackedRules)
			.toEqual(['option-explicit-missing']);
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			untrackedRules: ['argument-count', 'option-explicit-missing'],
			untrackedRulesSource: 'workbook',
			workbookUntrackedRules: ['option-explicit-missing'],
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
			workbookUntrackedRules: [],
		});
	});

});
