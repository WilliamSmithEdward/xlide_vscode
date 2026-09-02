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

import { readProjectSettings, writeProjectSettings } from '../src/projectSettings';
import {
	effectiveProjectAnalysisSettings,
	effectiveProjectAnalysisSettingsFromConfig,
	resetProjectAnalysisRuleTracking,
	setProjectAnalysisRuleTracked,
} from '../src/projectAnalysisSettings';

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

function tempWorkbook(): { root: string; project: string; exportFolder: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-analysis-settings-'));
	tempRoots.push(root);
	const project = path.join(root, 'Book.xlsm');
	const exportFolder = path.join(root, 'repo');
	fs.writeFileSync(project, '', 'utf8');
	return { root, project, exportFolder };
}

describe('project analysis settings', () => {
	it('uses global defaults until a project override exists', async () => {
		const { project } = tempWorkbook();
		mockConfig.visibleSeverities = ['error', 'information'];
		mockConfig.untrackedRules = ['option-explicit-missing'];
		mockConfig.ruleSeverityOverrides = { 'unknown-call': 'warning' };
		mockConfig.machineKeys.add('analysis.visibleSeverities');
		mockConfig.machineKeys.add('analysis.untrackedRules');
		mockConfig.machineKeys.add('analysis.ruleSeverityOverrides');

		await expect(effectiveProjectAnalysisSettings(project)).resolves.toMatchObject({
			visibleSeverities: ['error', 'information'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['option-explicit-missing'],
			untrackedRulesSource: 'machine',
			projectUntrackedRules: [],
			ruleSeverityOverrides: { 'unknown-call': 'warning' },
			ruleSeverityOverridesSource: 'machine',
		});
	});

	it('resolves effective settings from an already-loaded project config', () => {
		mockConfig.visibleSeverities = ['error'];
		mockConfig.machineKeys.add('analysis.visibleSeverities');

		expect(effectiveProjectAnalysisSettingsFromConfig('Book.xlsm', {
			analysis: {
				untrackedRules: ['argument-count'],
			},
		})).toMatchObject({
			visibleSeverities: ['error'],
			visibleSeveritiesSource: 'machine',
			untrackedRules: ['argument-count'],
			untrackedRulesSource: 'project',
			projectUntrackedRules: ['argument-count'],
			ruleSeverityOverridesSource: 'default',
		});
	});

	it('stores project rule tracking separately from the effective global default', async () => {
		const { project } = tempWorkbook();
		mockConfig.untrackedRules = ['argument-count'];
		mockConfig.machineKeys.add('analysis.untrackedRules');

		const update = await setProjectAnalysisRuleTracked(project, 'Option-Explicit-Missing', false);

		expect(update).toMatchObject({
			tracked: false,
			changed: true,
			untrackedRules: ['option-explicit-missing'],
		});
		expect((await readProjectSettings(project)).analysis?.untrackedRules)
			.toEqual(['option-explicit-missing']);
		await expect(effectiveProjectAnalysisSettings(project)).resolves.toMatchObject({
			untrackedRules: ['argument-count', 'option-explicit-missing'],
			untrackedRulesSource: 'project',
			projectUntrackedRules: ['option-explicit-missing'],
		});
	});

	it('resets rule tracking and removes empty analysis settings', async () => {
		const { project, exportFolder } = tempWorkbook();
		mockConfig.untrackedRules = ['option-explicit-missing'];
		mockConfig.machineKeys.add('analysis.untrackedRules');
		await writeProjectSettings(project, {
			exportFolder,
			importMode: 'trueUpStandardClass',
			analysis: {
				untrackedRules: ['argument-count'],
			},
		});

		await resetProjectAnalysisRuleTracking(project);

		expect(await readProjectSettings(project)).toEqual({
			exportFolder,
			importMode: 'trueUpStandardClass',
		});
		await expect(effectiveProjectAnalysisSettings(project)).resolves.toMatchObject({
			untrackedRules: ['option-explicit-missing'],
			untrackedRulesSource: 'machine',
			projectUntrackedRules: [],
		});
	});

});
