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
	resetWorkbookAnalysisRuleSeverities,
	resetWorkbookAnalysisSettings,
	resetWorkbookAnalysisVisibleSeverities,
	setWorkbookAnalysisRuleSeverityOverride,
	setWorkbookAnalysisRuleTracked,
	setWorkbookAnalysisVisibleSeverities,
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

	it('stores workbook severity overrides without disturbing sync settings', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 3,
			},
		});

		await setWorkbookAnalysisVisibleSeverities(workbook, ['warning']);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 3,
			},
			analysis: {
				visibleSeverities: ['warning'],
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['warning'],
			visibleSeveritiesSource: 'workbook',
			untrackedRulesSource: 'default',
			workbookUntrackedRules: [],
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

	it('resets one workbook analysis override while preserving the other', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		mockConfig.visibleSeverities = ['error'];
		mockConfig.machineKeys.add('analysis.visibleSeverities');
		await writeWorkbookSettings(workbook, {
			exportFolder,
			analysis: {
				visibleSeverities: ['warning'],
				untrackedRules: ['argument-count'],
				ruleSeverityOverrides: { 'unknown-call': 'warning' },
			},
		});

		await resetWorkbookAnalysisVisibleSeverities(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			analysis: {
				untrackedRules: ['argument-count'],
				ruleSeverityOverrides: { 'unknown-call': 'warning' },
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeverities: ['error'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['argument-count'],
			untrackedRulesSource: 'workbook',
			workbookUntrackedRules: ['argument-count'],
			ruleSeverityOverrides: { 'unknown-call': 'warning' },
			ruleSeverityOverridesSource: 'workbook',
		});
	});

	it('stores workbook rule severity overrides from the effective global default', async () => {
		const { workbook } = tempWorkbook();
		mockConfig.ruleSeverityOverrides = { 'member-not-found': 'warning' };
		mockConfig.machineKeys.add('analysis.ruleSeverityOverrides');

		await setWorkbookAnalysisRuleSeverityOverride(workbook, 'Unknown-Call', 'warning');

		expect((await readWorkbookSettings(workbook)).analysis?.ruleSeverityOverrides).toEqual({
			'member-not-found': 'warning',
			'unknown-call': 'warning',
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			ruleSeverityOverrides: {
				'member-not-found': 'warning',
				'unknown-call': 'warning',
			},
			ruleSeverityOverridesSource: 'workbook',
		});
	});

	it('drops disallowed workbook rule severity overrides through the guarded normalizer', async () => {
		const { workbook } = tempWorkbook();

		await setWorkbookAnalysisRuleSeverityOverride(workbook, 'option-explicit-missing', 'error');

		expect((await readWorkbookSettings(workbook)).analysis?.ruleSeverityOverrides).toBeUndefined();
	});

	it('resets workbook rule severity overrides independently', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		mockConfig.ruleSeverityOverrides = { 'member-not-found': 'warning' };
		mockConfig.machineKeys.add('analysis.ruleSeverityOverrides');
		await writeWorkbookSettings(workbook, {
			exportFolder,
			analysis: {
				untrackedRules: ['argument-count'],
				ruleSeverityOverrides: { 'unknown-call': 'warning' },
			},
		});

		await resetWorkbookAnalysisRuleSeverities(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			analysis: {
				untrackedRules: ['argument-count'],
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			ruleSeverityOverrides: { 'member-not-found': 'warning' },
			ruleSeverityOverridesSource: 'machine',
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

	it('resets all workbook analysis overrides without disturbing workbook-only settings', async () => {
		const { workbook, exportFolder } = tempWorkbook();
		await writeWorkbookSettings(workbook, {
			exportFolder,
			exportMode: 'trueUp',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 4,
			},
			analysis: {
				visibleSeverities: ['warning'],
				untrackedRules: ['argument-count'],
				ruleSeverityOverrides: { 'unknown-call': 'warning' },
			},
		});

		await resetWorkbookAnalysisSettings(workbook);

		expect(await readWorkbookSettings(workbook)).toEqual({
			exportFolder,
			exportMode: 'trueUp',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 4,
			},
		});
		await expect(effectiveWorkbookAnalysisSettings(workbook)).resolves.toMatchObject({
			visibleSeveritiesSource: 'default',
			untrackedRulesSource: 'default',
			workbookUntrackedRules: [],
		});
	});
});
