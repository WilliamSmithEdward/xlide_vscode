import { afterAll, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
vi.mock('vscode', async () => (await import('../helpers/vscodeMock')).vscodeMock());
import * as vscode from 'vscode';
import { ProjectEngine } from '../../src/projectEngine';

// A legacy file keeps its project in records XLIDE does not author, so a new
// one, or a first project in one that has none, is refused with the format
// to save it as. An extension missing from those lists fell through to the
// default and seeded an Excel workbook under the PowerPoint file's name.

/** An engine that finds the shipped blank templates, as the extension does. */
function engine(): ProjectEngine {
	return new ProjectEngine({ extensionUri: vscode.Uri.file(path.join(__dirname, '..', '..')) } as never);
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-legacy-create-'));
afterAll(() => { fs.rmSync(dir, { recursive: true, force: true }); });

describe.each([
	['ppt', 'pptm'],
	['pot', 'potm'],
	['pps', 'ppsm'],
])('a legacy .%s', (extension, modern) => {
	it('is not created', async () => {
		const target = path.join(dir, `New.${extension}`);
		const bridge = engine();
		await expect(bridge.call('createProject', { path: target }))
			.rejects.toThrow(`Creating new legacy-format files (.${extension}) is not supported`);
		expect(fs.existsSync(target)).toBe(false);
		bridge.dispose();
	});

	it('is told which format takes a first project', async () => {
		const target = path.join(dir, `Empty.${extension}`);
		fs.copyFileSync(path.join(__dirname, '..', 'fixtures', 'binaries', 'NoVbaFixture.ppt'), target);
		const before = fs.readFileSync(target);
		const bridge = engine();
		await expect(bridge.call('addVbaProject', { path: target }))
			.rejects.toThrow(`Save it as .${modern} first`);
		expect(fs.readFileSync(target).equals(before)).toBe(true);
		bridge.dispose();
	});
});

it('a .ppsx show, which holds no macros, is not created as a workbook', async () => {
	const target = path.join(dir, 'New.ppsx');
	const bridge = engine();
	await expect(bridge.call('createProject', { path: target }))
		.rejects.toThrow('.ppsx is not a macro-enabled format');
	expect(fs.existsSync(target)).toBe(false);
	bridge.dispose();
});
