import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import {
	applyFormDesignerOp,
	readFormExport,
	readFormMarkup,
	readFormPreview,
	readModules,
	resetWorkbookCacheForTests,
} from '../src/vba/workbookService';

// Canvas gestures: the same mutations the markup diff performs, addressed by
// control name (VBA requires names unique form-wide) and applied one gesture
// per write. Additions run through the authoring path live Excel verified.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

const tempDirs: string[] = [];
afterEach(() => {
	resetWorkbookCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function workbook(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-designer-'));
	tempDirs.push(dir);
	const wb = path.join(dir, 'Forms.xlsm');
	fs.copyFileSync(FIXTURE, wb);
	return wb;
}

describe('geometry gestures', () => {
	it('moves a top-level control', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'geometry', name: 'OkButton', left: 250, top: 250 });
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup)
			.toContain('<CommandButton Name="OkButton" Left="250" Top="250"');
	});

	it('moves a control nested in a Frame, found by name alone', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'geometry', name: 'PickAir', left: 10, top: 36 });
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup)
			.toContain('<OptionButton Name="PickAir" Left="10" Top="36"');
	});

	it('resizes a control and a container', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'geometry', name: 'NameBox', width: 150, height: 22 });
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'geometry', name: 'Options', width: 100, height: 72 });
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toMatch(/NameBox[^\r\n]*Width="150" Height="22"/);
		expect(markup).toMatch(/<Frame Name="Options"[^\r\n]*Width="100" Height="72"/);
	});

	it('reports a no-op honestly and writes nothing', () => {
		const wb = workbook();
		const before = fs.readFileSync(wb);
		const result = applyFormDesignerOp(wb, 'EntryForm', { kind: 'geometry', name: 'OkButton', left: 262 });
		expect(result.signatureDropped).toBe(false);
		expect(fs.readFileSync(wb).equals(before)).toBe(true);
	});
});

describe('toolbox additions', () => {
	it('adds at a point with a generated unique name', () => {
		const wb = workbook();
		const result = applyFormDesignerOp(wb, 'EntryForm', {
			kind: 'add', container: '', controlKind: 'Label', left: 12, top: 264,
		});
		// NameLabel and ViewNote exist; the generator still starts fresh at 1.
		expect(result.newName).toBe('Label1');
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup)
			.toContain('<Label Name="Label1" Left="12" Top="264"');
	});

	it('adds INTO a Frame and onto a Page by container name', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', {
			kind: 'add', container: 'Options', controlKind: 'OptionButton', left: 8, top: 50,
		});
		applyFormDesignerOp(wb, 'EntryForm', {
			kind: 'add', container: 'Page2', controlKind: 'TextBox', left: 10, top: 10,
		});
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		const frame = markup.slice(markup.indexOf('<Frame Name="Options"'), markup.indexOf('</Frame>'));
		expect(frame).toContain('OptionButton1');
		const page2 = markup.slice(markup.indexOf('<Page Name="Page2"'));
		expect(page2).toContain('TextBox1');
		// The analyzer sees additions the way it sees any designer control.
		const mods = readModules(wb, 'EntryForm' ? false : false);
		const form = mods.find((m) => m.name === 'EntryForm');
		expect(form?.implicitMembers?.map((m) => m.name)).toContain('OptionButton1');
		expect(form?.implicitMembers?.map((m) => m.name)).toContain('TextBox1');
	});
});

describe('removal', () => {
	it('removes by name, wherever it nests', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'remove', name: 'PickAir' });
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup).not.toContain('PickAir');
	});

	it('refuses removing a Page from the canvas, by name', () => {
		const wb = workbook();
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'remove', name: 'Page2' }))
			.toThrow(/remove pages through the form markup/);
	});
});

describe('form resize', () => {
	it('resizes the client area and keeps the VBFrame in step', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'formSize', width: 400, height: 320 });
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toContain('Width="400" Height="320"');
		// The VBFrame's twips client box repeats the size and must follow.
		const { frm } = readFormExport(wb, 'EntryForm');
		expect(frm).toMatch(/ClientWidth\s*=\s*8000/);
		expect(frm).toMatch(/ClientHeight\s*=\s*6400/);
	});
});

describe('the interactive canvas contract', () => {
	it('tags every control and surface for the gesture script', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain('data-name="OkButton"');
		expect(html).toContain('data-surface=""');
		expect(html).toContain('data-surface="Options"');
		expect(html).toContain('data-surface="Page1"');
		expect(html).toContain('id="toolbar"');
		expect(html).toContain('data-kind="CommandButton"');
		expect(html).toContain('acquireVsCodeApi');
	});

	it('restores a selection across re-renders', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm', 'OkButton');
		expect(html).toMatch(/class="ctl button selected" data-name="OkButton"/);
	});
});
