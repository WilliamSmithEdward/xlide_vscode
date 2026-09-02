import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { parseFormPackage } from '../src/vba/oforms/formPackage';
import { renderFormPreviewHtml } from '../src/vba/oforms/preview';
import { readFormPreview } from '../src/vba/projectService';

// The preview is a RENDERER of the same projection the markup edits: real
// bounds, captions, colors and fonts from the binary, with honest hatched
// placeholders where fidelity runs out (pictures, foreign controls). These
// tests pin the rendering against the Excel-authored 19-control fixture.

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');

function fixtureHtml(): string {
	const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(FIXTURE)).readVbaProject());
	const pkg = parseFormPackage(cfb, ['EntryForm']);
	return renderFormPreviewHtml(pkg, { formName: 'EntryForm', caption: 'Quarter Entry' });
}

describe('the canvas stylesheet places every control where the scene says', () => {
	// A control's box comes from .ctl { position: absolute } plus the inline
	// left/top the adapter computed. A kind rule that names position again
	// wins on order and drops the control into normal flow, which the canvas
	// reports at runtime as "misplaced ... pos=relative" - measured on a
	// scroll bar, whose own rule did exactly that.
	const KIND_CLASSES = ['label', 'edit', 'combo', 'list', 'opt', 'button', 'image', 'spin', 'scroll', 'tabstrip',
		'frame', 'multipage', 'foreign', 'line', 'shape', 'timer', 'data', 'ole', 'picture'];

	it('never lets a control kind override the absolute position .ctl sets', () => {
		const style = /<style>([\s\S]*?)<\/style>/.exec(fixtureHtml())?.[1];
		expect(style).toBeDefined();
		for (const cls of KIND_CLASSES) {
			const rule = new RegExp(`(^|[^-\\w.])\\.${cls}\\s*\\{([^}]*)\\}`, 'g');
			for (let m = rule.exec(style!); m; m = rule.exec(style!)) {
				const position = /(^|;)\s*position\s*:\s*([a-z]+)/.exec(m[2]);
				expect(position?.[2] ?? 'absolute', `.${cls} { position: ${position?.[2]} }`).toBe('absolute');
			}
		}
	});
});

describe('the rendered form', () => {
	it('draws the dialog at the form size with its caption', () => {
		const html = fixtureHtml();
		expect(html).toContain('Quarter Entry');
		expect(html).toContain('width: 348pt');
		expect(html).toContain('height: 291pt');
	});

	it('places every control at its real bounds', () => {
		const html = fixtureHtml();
		// NameLabel: Left=12 Top=14 Width=66 Height=16, in points.
		expect(html).toMatch(/class="ctl label" data-name="NameLabel" style="left:12pt;top:14pt;width:66pt;height:16pt;[^"]*" title="NameLabel"/);
		expect(html).toContain('>Customer</div>');
	});

	it('renders each kind as itself', () => {
		const html = fixtureHtml();
		expect(html).toContain('title="NameBox"');       // TextBox
		expect(html).toContain('class="ctl edit combo"'); // ComboBox
		expect(html).toMatch(/class="box"/);              // CheckBox glyph
		expect(html).toMatch(/class="radio"/);            // OptionButton glyph
		expect(html).toContain('class="ctl button"');     // CommandButton
		expect(html).toContain('>Start</div>');
	});

	it('nests Frame contents inside the frame surface', () => {
		const html = fixtureHtml();
		const frame = html.indexOf('title="Options"');
		const ground = html.indexOf('title="PickGround"');
		expect(frame).toBeGreaterThan(-1);
		expect(ground).toBeGreaterThan(frame);
	});

	it('renders MultiPage pages with their captions and contents', () => {
		const html = fixtureHtml();
		expect(html).toContain('>Page1</span>');
		expect(html).toContain('>Page2</span>');
		expect(html).toContain('title="Agree"');
		// Only the first page shows; the second is hidden until its tab.
		expect(html).toMatch(/<div class="page" id="[^"]*" data-surface="Page2" hidden>/);
	});

	it('shows tab captions on the standalone TabStrip', () => {
		const html = fixtureHtml();
		expect(html).toContain('>Tab1</span>');
		expect(html).toContain('>Tab2</span>');
	});

	it('paints the picture-bearing Image with its real BMP', () => {
		const html = fixtureHtml();
		expect(html).toMatch(/class="ctl image pictured"[^>]*title="Badge"/);
		expect(html).toMatch(/data-name="Badge"[^>]*data:image\/bmp;base64,/);
	});

	it('carries the fixture font onto controls', () => {
		const html = fixtureHtml();
		expect(html).toContain("font-family:'Tahoma'");
		expect(html).toContain('font-size:8.25pt');
	});
});

describe('the service entry point', () => {
	it('renders straight from a project path', () => {
		const { html } = readFormPreview(FIXTURE, 'EntryForm');
		expect(html).toContain('Quarter Entry');
		expect(html).toContain('title="OkButton"');
	});
});

describe('the toolbox drag', () => {
	it('carries its hooks in the canvas script', () => {
		const { html } = readFormPreview(FIXTURE, 'EntryForm');
		expect(html).toContain('toolDrag');
		expect(html).toContain("className = 'ghost'");
		expect(html).toContain('.ghost {');
	});
});
