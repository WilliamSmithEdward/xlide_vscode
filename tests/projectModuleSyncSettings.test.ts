import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	effectiveProjectModuleSyncSettings,
	setProjectModuleSyncExportMode,
	updateProjectModuleSyncSettings,
} from '../src/projectModuleSyncSettings';
import {
	readProjectSettings,
	settingsPathForProject,
	writeProjectSettings,
} from '../src/projectSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; project: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-module-sync-settings-'));
	tempRoots.push(root);
	const project = path.join(root, 'Book.xlsm');
	fs.writeFileSync(project, '', 'utf8');
	return { root, project };
}

describe('project module sync settings', () => {
	it('resolves missing project sync settings to built-in defaults', async () => {
		const { project } = tempWorkbook();

		await expect(effectiveProjectModuleSyncSettings(project)).resolves.toEqual({
			folderPath: undefined,
			folderPathSource: 'missing',
			exportMode: 'exportAll',
			exportModeSource: 'default',
			importMode: 'updateOnly',
			importModeSource: 'default',
			settingsPath: settingsPathForProject(project),
		});
	});

	it('reports project sync overrides with project provenance', async () => {
		const { project } = tempWorkbook();
		await writeProjectSettings(project, {
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			importMode: 'trueUpStandardClass',
		});

		await expect(effectiveProjectModuleSyncSettings(project)).resolves.toEqual({
			folderPath: 'C:/repo',
			folderPathSource: 'project',
			exportMode: 'trueUp',
			exportModeSource: 'project',
			importMode: 'trueUpStandardClass',
			importModeSource: 'project',
			settingsPath: settingsPathForProject(project),
		});
	});

	it('updates only provided sync settings and preserves unrelated project settings', async () => {
		const { project } = tempWorkbook();
		await writeProjectSettings(project, {
			exportFolder: 'C:/repo',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(updateProjectModuleSyncSettings(project, { exportMode: 'trueUp' })).resolves.toMatchObject({
			folderPath: 'C:/repo',
			exportMode: 'trueUp',
			exportModeSource: 'project',
			importMode: 'updateOnly',
			importModeSource: 'default',
		});
		expect(await readProjectSettings(project)).toEqual({
			exportFolder: 'C:/repo',
			exportMode: 'trueUp',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});
	});

	it('sets export mode through the module sync settings owner', async () => {
		const { project } = tempWorkbook();
		await writeProjectSettings(project, {
			exportFolder: 'C:/repo',
			exportMode: 'exportAll',
			importMode: 'trueUpStandardClass',
			analysis: {
				visibleSeverities: ['error', 'information'],
				untrackedRules: ['option-explicit-missing'],
			},
		});

		await expect(setProjectModuleSyncExportMode(project, 'trueUp')).resolves.toMatchObject({
			folderPath: 'C:/repo',
			exportMode: 'trueUp',
			exportModeSource: 'project',
			importMode: 'trueUpStandardClass',
			importModeSource: 'project',
		});
		expect(await readProjectSettings(project)).toEqual({
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
