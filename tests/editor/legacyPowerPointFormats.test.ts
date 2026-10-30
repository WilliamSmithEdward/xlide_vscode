import { afterEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock());

import { ProjectEngine } from '../../src/projectEngine';
import { hostTokenForFileName } from '../../src/analyzer/host/hostRegistry';
import { containerHostForPath, MACRO_CONTAINER_EXTENSIONS } from '../../src/macroContainerUi';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';
import { listModules, readModule, validateProject, writeModule } from '../../src/vba/projectService';
import { setExtensionAssetRoot } from '../../src/extensionAssets';

// Issue #1533. PowerPoint 97-2003 saves a template (.pot) and a show (.pps)
// in the same compound file as a presentation (.ppt). PowerPointFixture.pot
// and .pps were saved by PowerPoint 16.0 (build 20527) from
// PowerPointFixture.ppt with SaveCopyAs ppSaveAsTemplate (5) and
// ppSaveAsShow (7), with macros force-disabled.

const FIXTURES = path.join(__dirname, '..', 'fixtures', 'binaries');
// The blank templates createProject and addVbaProject seed from.
const REPO_ROOT = path.join(__dirname, '..', '..');
setExtensionAssetRoot(REPO_ROOT);
const engine = new ProjectEngine({ extensionUri: { scheme: 'file', path: REPO_ROOT, fsPath: REPO_ROOT } } as never);
const temps: string[] = [];

function tempDir(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-legacy-ppt-'));
	temps.push(dir);
	return dir;
}

afterEach(() => {
	for (const dir of temps.splice(0)) { fs.rmSync(dir, { recursive: true, force: true }); }
});

describe('legacy PowerPoint templates and shows', () => {
	it.each(['pot', 'pps'])('treats .%s as a PowerPoint file everywhere the extension decides', (extension) => {
		expect(hostTokenForFileName(`Design.${extension}`)).toBe('powerpoint');
		expect(containerHostForPath(`C:\\work\\Design.${extension}`)).toBe('powerpoint');
		expect(MACRO_CONTAINER_EXTENSIONS).toContain(extension);
	});

	it.each(['ppt', 'pot', 'pps'])('checks code in a .%s against the PowerPoint model', (extension) => {
		const source = 'Option Explicit\r\nPublic Sub Probe()\r\n    ActivePresentation.Slides.Nope\r\nEnd Sub\r\n';
		const messages = analyzeVbaModuleSource({
			source,
			moduleName: 'Module1',
			moduleKind: 'standard',
			host: hostTokenForFileName(`Design.${extension}`),
		}).diagnostics.map((diagnostic) => diagnostic.message);
		expect(messages.join('\n')).toContain('Method or data member not found');
	});

	it.each(['PowerPointFixture.pot', 'PowerPointFixture.pps'])('reads and writes %s like the .ppt', (file) => {
		const target = path.join(tempDir(), file);
		fs.copyFileSync(path.join(FIXTURES, file), target);
		expect(listModules(target).map((module) => module.name)).toEqual(expect.arrayContaining(['Module1', 'CDeck']));
		expect(readModule(target, 'CDeck', true).source).toContain('ActivePresentation.Name');
		writeModule(target, 'MAdded', 'Public Sub AddedProc()\r\nEnd Sub\r\n');
		expect(readModule(target, 'MAdded', true).source).toContain('Public Sub AddedProc()');
		expect(readModule(target, 'Module1', true).source).toContain('ActivePresentation.Slides.');
		expect(validateProject(target).issues).toEqual([]);
	});

	it('creates a .pptm, so the refusals below are not an unset template root', async () => {
		const target = path.join(tempDir(), 'New.pptm');
		await engine.call('createProject', { path: target });
		expect(fs.readFileSync(target).equals(fs.readFileSync(path.join(REPO_ROOT, 'assets', 'templates', 'blank.pptm')))).toBe(true);
	});

	it.each(['pot', 'pps', 'ppsx'])('refuses to create a .%s and writes nothing', async (extension) => {
		const target = path.join(tempDir(), `New.${extension}`);
		await expect(engine.call('createProject', { path: target })).rejects.toThrow();
		expect(fs.existsSync(target)).toBe(false);
	});

	it.each([['pot', 'potm'], ['pps', 'ppsm']])('names the format to save a .%s as before adding a project', async (extension, modern) => {
		const target = path.join(tempDir(), `Design.${extension}`);
		fs.copyFileSync(path.join(FIXTURES, `PowerPointFixture.${extension}`), target);
		const before = fs.readFileSync(target);
		await expect(engine.call('addVbaProject', { path: target })).rejects.toThrow(`Save it as .${modern} first`);
		expect(fs.readFileSync(target).equals(before)).toBe(true);
	});
});
