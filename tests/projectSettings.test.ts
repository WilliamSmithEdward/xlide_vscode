import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	isProjectSettingsError,
	readProjectSettings,
	resolveProjectSetting,
	settingsPathForProject,
	updateProjectSettings,
	writeProjectSettings,
} from '../src/projectSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; project: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-workbook-settings-'));
	tempRoots.push(root);
	const project = path.join(root, 'Book.xlsm');
	fs.writeFileSync(project, '', 'utf8');
	return { root, project };
}

describe('projectSettings', () => {
	it('resolves project overrides over global defaults with explicit source', () => {
		const fallback = { value: ['error'], source: 'default' as const };
		expect(resolveProjectSetting(['warning'], fallback)).toEqual({
			value: ['warning'],
			source: 'project',
		});
		expect(resolveProjectSetting(undefined, fallback)).toEqual({
			value: ['error'],
			source: 'default',
		});
	});

	it('treats a missing project settings sidecar as no project settings', async () => {
		const { project } = tempWorkbook();

		await expect(readProjectSettings(project)).resolves.toEqual({});
	});

	it('writes normalized project settings to the project sidecar', async () => {
		const { project } = tempWorkbook();

		await writeProjectSettings(project, {
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 7,
			},
			analysis: {
				visibleSeverities: ['warning', 'warning'],
				untrackedRules: [' Option-Explicit-Missing ', 'option-explicit-missing'],
				ruleSeverityOverrides: {
					' Unknown-Call ': 'warning',
					'option-explicit-missing': 'error',
				},
			},
		});

		expect(JSON.parse(fs.readFileSync(settingsPathForProject(project), 'utf8'))).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'updateOnly',
			tests: {
				artifactFolder: 'custom-tests',
				artifactRetention: 7,
			},
			analysis: {
				visibleSeverities: ['warning', 'warning'],
				untrackedRules: ['option-explicit-missing'],
				ruleSeverityOverrides: {
					'unknown-call': 'warning',
				},
			},
		});
	});

	it('updates project settings through one read-normalize-write owner', async () => {
		const { project } = tempWorkbook();
		await writeProjectSettings(project, {
			exportFolder: 'C:/repo',
			exportMode: 'exportAll',
			analysis: {
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(updateProjectSettings(project, (existing) => ({
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

	it('can skip project settings writes when an update is unnecessary', async () => {
		const { project } = tempWorkbook();
		await writeProjectSettings(project, { exportFolder: 'C:/repo' });
		const before = fs.readFileSync(settingsPathForProject(project), 'utf8');

		await expect(updateProjectSettings(project, () => undefined)).resolves.toEqual({
			exportFolder: 'C:/repo',
		});

		expect(fs.readFileSync(settingsPathForProject(project), 'utf8')).toBe(before);
	});

	it('rejects invalid project settings JSON with the sidecar path', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(settingsPathForProject(project), '{ nope', 'utf8');

		await expect(readProjectSettings(project)).rejects.toMatchObject({
			settingsPath: settingsPathForProject(project),
			name: 'ProjectSettingsError',
		});
		await expect(readProjectSettings(project)).rejects.toThrow('Expected valid JSON');
	});

	it('recovers a project settings sidecar with trailing duplicate JSON and normalizes it on update', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({ exportFolder: 'C:/old' })}\n${JSON.stringify({ exportFolder: 'C:/repo' })}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).resolves.toEqual({
			exportFolder: 'C:/repo',
		});

		await updateProjectSettings(project, (existing) => ({
			...existing,
			analysis: {
				untrackedRules: ['option-explicit-missing'],
			},
		}));

		expect(fs.readFileSync(settingsPathForProject(project), 'utf8'))
			.toBe(`${JSON.stringify({
				exportFolder: 'C:/repo',
				analysis: {
					untrackedRules: ['option-explicit-missing'],
				},
			}, null, 2)}\n`);
	});

	it('serializes concurrent project settings updates for the same sidecar', async () => {
		const { project } = tempWorkbook();

		await Promise.all([
			updateProjectSettings(project, (existing) => ({
				...existing,
				exportFolder: 'C:/repo',
			})),
			updateProjectSettings(project, (existing) => ({
				...existing,
				analysis: {
					...existing.analysis,
					untrackedRules: ['option-explicit-missing'],
				},
			})),
		]);

		expect(await readProjectSettings(project)).toEqual({
			exportFolder: 'C:/repo',
			analysis: {
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});

	it('rejects non-object project settings', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(settingsPathForProject(project), '[]', 'utf8');

		await expect(readProjectSettings(project)).rejects.toThrow('Expected the root value to be a JSON object');
	});

	it('rejects unknown project settings keys instead of ignoring them', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({ exportFolder: 'C:/repo', typoMode: true }, null, 2)}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).rejects.toThrow('Unknown setting "typoMode"');
	});

	it('lenient read keeps the valid subset of a stale sidecar instead of throwing', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({
				exportFolder: 'C:/repo',
				legacyOption: true,
				analysis: { ruleSeverityOverrides: { 'since-removed-rule': 'off' } },
			}, null, 2)}\n`,
			'utf8',
		);

		// Strict read (the settings editor) still surfaces a stale key as a typo.
		await expect(readProjectSettings(project)).rejects.toThrow('Unknown setting "legacyOption"');

		// Lenient read (the diagnostics/apply path) never throws on version skew:
		// it drops the unknown key and stale rule code and keeps the valid subset,
		// so a stale sidecar cannot blast an error across every module.
		const lenient = await readProjectSettings(project, { lenient: true });
		expect(lenient.exportFolder).toBe('C:/repo');
		expect(lenient.analysis?.ruleSeverityOverrides ?? {}).toEqual({});
	});

	it('rejects invalid project sync modes from disk', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({ exportFolder: 'C:/repo', exportMode: 'anythingElse' }, null, 2)}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).rejects.toThrow('Expected "exportMode" to be "exportAll" or "trueUp"');
	});

	it('rejects invalid project analysis settings from disk', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({
				analysis: {
					visibleSeverities: ['error', 'hint'],
					ruleSeverityOverrides: {
						'option-explicit-missing': 'error',
					},
				},
			}, null, 2)}\n`,
			'utf8',
		);

		try {
			await readProjectSettings(project);
			throw new Error('Expected readProjectSettings to reject');
		} catch (err) {
			expect(isProjectSettingsError(err)).toBe(true);
			expect(err).toMatchObject({
				settingsPath: settingsPathForProject(project),
			});
			expect(err instanceof Error ? err.message : String(err))
				.toContain('Expected "analysis.visibleSeverities" entries');
		}
	});

	it('rejects invalid project test settings from disk', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({
				tests: {
					artifactFolder: 123,
					artifactRetention: 0,
				},
			}, null, 2)}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).rejects.toThrow(
			'Expected "tests.artifactFolder" to be a string.',
		);

		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({
				tests: {
					artifactFolder: 'tests',
					artifactRetention: 0,
				},
			}, null, 2)}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).rejects.toThrow(
			'Expected "tests.artifactRetention" to be a positive integer.',
		);
	});

	it('rejects invalid project rule severity overrides from disk', async () => {
		const { project } = tempWorkbook();
		fs.writeFileSync(
			settingsPathForProject(project),
			`${JSON.stringify({
				analysis: {
					ruleSeverityOverrides: {
						'option-explicit-missing': 'error',
					},
				},
			}, null, 2)}\n`,
			'utf8',
		);

		await expect(readProjectSettings(project)).rejects.toThrow(
			'Expected "analysis.ruleSeverityOverrides.option-explicit-missing" to be one of: off.',
		);
	});
});
