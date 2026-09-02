import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
	effectiveWorkbookTestSettings,
	effectiveWorkbookTestSettingsFromConfig,
} from '../src/workbookTestSettings';

const tempRoots: string[] = [];

afterEach(() => {
	for (const root of tempRoots.splice(0)) {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

function tempWorkbook(): { root: string; project: string } {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-workbook-test-settings-'));
	tempRoots.push(root);
	const project = path.join(root, 'Book.xlsm');
	fs.writeFileSync(project, '', 'utf8');
	return { root, project };
}

describe('project test settings', () => {
	it('uses default artifact settings until project overrides exist', async () => {
		const { project } = tempWorkbook();

		await expect(effectiveWorkbookTestSettings(project)).resolves.toMatchObject({
			artifactFolder: 'tests',
			artifactFolderSource: 'default',
			artifactRetention: 20,
			artifactRetentionSource: 'default',
		});
	});

	it('resolves artifact settings from already-loaded project config', () => {
		expect(effectiveWorkbookTestSettingsFromConfig('Book.xlsm', {
			tests: {
				artifactFolder: 'ci-artifacts',
				artifactRetention: 5,
			},
		})).toMatchObject({
			artifactFolder: 'ci-artifacts',
			artifactFolderSource: 'project',
			artifactRetention: 5,
			artifactRetentionSource: 'project',
		});
	});

});
