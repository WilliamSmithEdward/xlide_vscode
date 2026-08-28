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

describe('reparent', () => {
	const between = (markup: string, name: string, openTag: RegExp, closeTag: string): boolean => {
		const open = markup.search(openTag);
		const close = markup.indexOf(closeTag, open);
		const at = markup.indexOf(`Name="${name}"`);
		return open >= 0 && at > open && at < close;
	};

	it('moves a top-level control into a Frame', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'OkButton', container: 'Options', left: 12, top: 48 });
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(between(markup, 'OkButton', /<Frame Name="Options"/, '</Frame>')).toBe(true);
		expect(markup).toContain('<CommandButton Name="OkButton" Left="12" Top="48"');
	});

	it('moves a nested control out to the form', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'PickAir', container: '', left: 200, top: 6 });
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(between(markup, 'PickAir', /<Frame Name="Options"/, '</Frame>')).toBe(false);
		expect(markup).toContain('<OptionButton Name="PickAir" Left="200" Top="6"');
	});

	it('moves a whole Frame onto a Page, children intact', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'Options', container: 'Page1', left: 6, top: 6 });
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(between(markup, 'Options', /<Page Name="Page1"/, '</Page>')).toBe(true);
		expect(between(markup, 'PickAir', /<Frame Name="Options"/, '</Frame>')).toBe(true);
	});

	it('treats a drop on its own surface as a plain move', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'OkButton', container: '', left: 111, top: 222 });
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup)
			.toContain('<CommandButton Name="OkButton" Left="111" Top="222"');
	});

	it('refuses to move a Page, and a Frame into itself', () => {
		const wb = workbook();
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'Page2', container: '', left: 0, top: 0 }))
			.toThrow(/pages stay inside their MultiPage/);
		expect(() => applyFormDesignerOp(wb, 'EntryForm', { kind: 'reparent', name: 'Options', container: 'Options', left: 0, top: 0 }))
			.toThrow(/cannot move into itself/);
	});
});

describe('property writes', () => {
	const set = (wb: string, name: string, prop: string, value: string) =>
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'setProp', name, prop, value });

	it('sets a caption, a color, a font style, and a tab index', () => {
		const wb = workbook();
		set(wb, 'OkButton', 'Caption', 'Go');
		set(wb, 'NameBox', 'BackColor', '#FF0000');
		set(wb, 'NameBox', 'Font.Bold', 'True');
		set(wb, 'OkButton', 'TabIndex', '9');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toContain('Caption="Go"');
		expect(markup).toMatch(/Name="NameBox"[^>]*BackColor="#ff0000"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*Font\.Bold="True"/);
		expect(markup).toMatch(/Name="OkButton"[^>]*TabIndex="9"/);
	});

	it('writes the boolean flags, printed only off their kind defaults', () => {
		const wb = workbook();
		set(wb, 'OkButton', 'Enabled', 'False');
		set(wb, 'OkButton', 'Default', 'True');
		set(wb, 'OkButton', 'TabStop', 'False');
		set(wb, 'NameBox', 'MultiLine', 'True');
		set(wb, 'NameBox', 'Locked', 'True');
		set(wb, 'NameLabel', 'WordWrap', 'False');
		set(wb, 'Taxable', 'Visible', 'False');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toMatch(/Name="OkButton"[^>]*Enabled="False"/);
		expect(markup).toMatch(/Name="OkButton"[^>]*Default="True"/);
		expect(markup).toMatch(/Name="OkButton"[^>]*TabStop="False"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*MultiLine="True"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*Locked="True"/);
		expect(markup).toMatch(/Name="NameLabel"[^>]*WordWrap="False"/);
		expect(markup).toMatch(/Name="Taxable"[^>]*Visible="False"/);
		// Setting a flag BACK to its default silences it again.
		set(wb, 'OkButton', 'Enabled', 'True');
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup).not.toContain('Enabled=');
	});

	it('writes alignment, list behavior, combo style, and the extra font effects', () => {
		const wb = workbook();
		set(wb, 'NameLabel', 'TextAlign', 'Center');
		set(wb, 'NameBox', 'TextAlign', 'Right');
		set(wb, 'NameBox', 'Font.Underline', 'True');
		set(wb, 'NameBox', 'Font.Strikethrough', 'True');
		set(wb, 'HistoryList', 'MultiSelect', '1');
		set(wb, 'HistoryList', 'ColumnCount', '2');
		set(wb, 'RegionPick', 'Style', '2');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toMatch(/Name="NameLabel"[^>]*TextAlign="Center"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*TextAlign="Right"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*Font\.Underline="True"/);
		expect(markup).toMatch(/Name="NameBox"[^>]*Font\.Strikethrough="True"/);
		expect(markup).toMatch(/Name="HistoryList"[^>]*MultiSelect="1"/);
		expect(markup).toMatch(/Name="HistoryList"[^>]*ColumnCount="2"/);
		expect(markup).toMatch(/Name="RegionPick"[^>]*Style="2"/);
		// Style back to the dropdown combo goes quiet again.
		set(wb, 'RegionPick', 'Style', '0');
		resetWorkbookCacheForTests();
		expect(readFormMarkup(wb, 'EntryForm').markup).not.toContain('Style=');
	});

	it('refuses alignment and style where they do not belong', () => {
		const wb = workbook();
		expect(() => set(wb, 'NameBox', 'TextAlign', 'Justified')).toThrow(/not Left, Center, or Right/);
		expect(() => set(wb, 'NameBox', 'Style', '2')).toThrow(/has no Style/);
		expect(() => set(wb, 'RegionPick', 'Style', '1')).toThrow(/not 0 or 2/);
	});

	it('refuses a flag the kind does not carry', () => {
		const wb = workbook();
		expect(() => set(wb, 'OkButton', 'MultiLine', 'True')).toThrow(/has no MultiLine/);
		expect(() => set(wb, 'NameLabel', 'TabStop', 'False')).toThrow(/has no TabStop/);
		expect(() => set(wb, 'NameBox', 'Default', 'True')).toThrow(/has no Default/);
		expect(() => set(wb, 'OkButton', 'Enabled', 'maybe')).toThrow(/not True or False/);
	});

	it('renames a control and reports the new name', () => {
		const wb = workbook();
		const result = set(wb, 'OkButton', 'Name', 'GoButton');
		expect(result.newName).toBe('GoButton');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toContain('Name="GoButton"');
		expect(markup).not.toContain('Name="OkButton"');
	});

	it('refuses a taken name, an illegal name, and an unknown property', () => {
		const wb = workbook();
		expect(() => set(wb, 'OkButton', 'Name', 'NameBox')).toThrow(/already exists/);
		expect(() => set(wb, 'OkButton', 'Name', '1Bad')).toThrow(/not a legal control name/);
		expect(() => set(wb, 'OkButton', 'Bogus', 'x')).toThrow(/has no Bogus/);
	});

	it('writes captions where each container keeps its own', () => {
		const wb = workbook();
		set(wb, 'Options', 'Caption', 'Choices');
		set(wb, 'Page1', 'Caption', 'First Things');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toMatch(/<Frame Name="Options"[^>]*Caption="Choices"/);
		expect(markup).toMatch(/<Page Name="Page1" Caption="First Things"/);
	});

	it('writes the form itself: caption via the VBFrame, size with its twips echo', () => {
		const wb = workbook();
		set(wb, '', 'Caption', 'Entry Station');
		set(wb, '', 'Width', '400');
		set(wb, '', 'BackColor', '#00FF00');
		resetWorkbookCacheForTests();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		expect(markup).toMatch(/<Form Name="EntryForm" Caption="Entry Station"[^>]*Width="400"/);
		expect(markup).toMatch(/<Form[^>]*BackColor="#00ff00"/);
		const { frm } = readFormExport(wb, 'EntryForm');
		expect(frm).toMatch(/ClientWidth\s*=\s*8000/);
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

	it('starts with grid snap on and neighbor snap off, and keeps them exclusive', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain('id="snapGrid" checked');
		expect(html).toContain('id="snapNeighbors">');
		// Checking either snap clears the other in the gesture script.
		expect(html).toContain('neighborsBox.checked = false');
		expect(html).toContain('gridBox.checked = false');
	});

	it('paints the 6pt lattice on every surface while grid snap is on', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain('body.grid-on [data-surface]::before');
		expect(html).toContain('syncGridDots');
	});

	it('carries the Properties pane and its rows', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain('id="props"');
		expect(html).toContain("type: 'setProp'");
		expect(html).toContain('"prop":"Caption"');
	});

	it('paints stored pictures instead of placeholders', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		// The fixture carries two real BMPs: the Badge image and a picture ON
		// the OK button. Both should arrive as data URIs; the Badge drops its
		// placeholder label, the button keeps its caption over the picture.
		// Single-quoted url() - the style attribute is double-quoted, and a
		// double-quoted URI truncated it silently once (measured in the page:
		// class survived, backgrounds died).
		expect(html.split("url('data:image/bmp;base64,").length - 1).toBe(2);
		expect(html).toMatch(/data-name="Badge"[^>]*background-size:contain/);
		expect(html).not.toMatch(/data-name="Badge"[^>]*><span>/);
	});

	it('draws the disabled gray, alignment, and text decoration it can set', () => {
		const wb = workbook();
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'setProp', name: 'OkButton', prop: 'Enabled', value: 'False' });
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'setProp', name: 'NameLabel', prop: 'TextAlign', value: 'Center' });
		applyFormDesignerOp(wb, 'EntryForm', { kind: 'setProp', name: 'NameBox', prop: 'Font.Underline', value: 'True' });
		resetWorkbookCacheForTests();
		const { html } = readFormPreview(wb, 'EntryForm');
		expect(html).toMatch(/data-name="OkButton"[^>]*color:#6d6d6d/);
		expect(html).toMatch(/data-name="NameLabel"[^>]*text-align:center/);
		expect(html).toMatch(/data-name="NameBox"[^>]*text-decoration:underline/);
	});

	it('jumps to the default event handler on double-click', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain("type: 'openHandler'");
		expect(html).toContain('DEFAULT_EVENTS');
		expect(html).toContain("MultiPage: 'Change'");
	});

	it('offers cross-surface drops in the gesture script', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm');
		expect(html).toContain("type: 'reparent'");
		expect(html).toContain('.drop-target');
	});

	it('activates the form itself when the selection is the empty name', () => {
		const { html } = readFormPreview(workbook(), 'EntryForm', '');
		expect(html).toContain('class="dialog form-selected"');
		expect(html).toContain("post({ type: 'formResize', width, height })");
	});
});
