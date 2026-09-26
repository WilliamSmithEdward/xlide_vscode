// What the MSForms designer writes alongside a property, and the defaults it
// gives a new control (issue #138, carried over from pyOpenVBA #31). Every
// expectation was measured in Word 16.0 (build 20326) on 2026-09-26 by
// setting the property through the VBE designer, saving the document and
// reading the bytes back with this engine.

import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { readFormMarkup, applyFormMarkup, resetProjectCacheForTests } from '../src/vba/projectService';
import { Cfb } from '../src/vba/cfb';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { parseFormPackage, walkPackages, type FormPackage } from '../src/vba/oforms/formPackage';
import { siteName, siteId } from '../src/vba/oforms/formStream';
import type { ParsedRecord } from '../src/vba/oforms/records';

const FIXTURE = path.join('tests', 'fixtures', 'binaries', 'FormFixtureVbide.xlsm');
const CRLF = '\r\n';

const tempDirs: string[] = [];
afterEach(() => {
	resetProjectCacheForTests();
	for (const dir of tempDirs.splice(0)) {
		fs.rmSync(dir, { recursive: true, force: true });
	}
});

function project(): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-form-couplings-'));
	tempDirs.push(dir);
	const wb = path.join(dir, 'Forms.xlsm');
	fs.copyFileSync(FIXTURE, wb);
	return wb;
}

function formPackage(wb: string): FormPackage {
	const cfb = Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(wb)).readVbaProject());
	return parseFormPackage(cfb, ['EntryForm']);
}

interface Found { record?: ParsedRecord; site: FormPackage['form']['sites'][number]; owner: FormPackage }

function find(pkg: FormPackage, name: string): Found {
	let found: Found | undefined;
	walkPackages(pkg, (p) => {
		p.form.sites.forEach((site, i) => {
			if (siteName(site) === name) {
				const entry = p.entries[i];
				found = { record: entry?.kind === 'record' ? entry.record : undefined, site, owner: p };
			}
		});
	});
	if (!found) { throw new Error(`no control ${name}`); }
	return found;
}

function addToForm(wb: string, elements: string): string {
	const markup = readFormMarkup(wb, 'EntryForm').markup;
	const edited = markup.replace(/<\/Form>\s*$/, `${elements}${CRLF}</Form>${CRLF}`);
	applyFormMarkup(wb, 'EntryForm', edited);
	resetProjectCacheForTests();
	return readFormMarkup(wb, 'EntryForm').markup;
}

describe('new-control defaults the designer gives (issue #138)', () => {
	it('gives a new control the font of the surface it lands on, in twips', () => {
		const wb = project();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		// The fixture form carries no font of its own; give it one, then add.
		const withFont = markup
			.replace(/<Form /, '<Form Font.Name="Aptos" Font.Size="10.5" ')
			.replace(/<\/Form>\s*$/, `    <CommandButton Name="Inherit1" Left="8" Top="8" Width="72" Height="24" />${CRLF}</Form>${CRLF}`);
		expect(withFont).toContain('Font.Name="Aptos" Font.Size="10.5"');
		applyFormMarkup(wb, 'EntryForm', withFont);
		resetProjectCacheForTests();
		const { record } = find(formPackage(wb), 'Inherit1');
		expect(record?.textProps?.strings.get('FontName')?.text).toBe('Aptos');
		expect(record?.textProps?.values.get('FontHeight')).toBe(210);
	});

	it('writes a Label site without TabStop and a Frame with the etched border', () => {
		const wb = project();
		addToForm(wb, `    <Label Name="Lbl1" Left="8" Top="8" Width="72" Height="12" Caption="x" />${CRLF}    <Frame Name="Frm1" Left="8" Top="30" Width="120" Height="90" />`);
		const pkg = formPackage(wb);
		expect(find(pkg, 'Lbl1').site.values.get('BitFlags')).toBe(0x32);
		const frame = find(pkg, 'Frm1');
		const inner = frame.owner.containers.get(siteId(frame.site))!;
		expect(inner.form.record.values.get('SpecialEffect')).toBe(3);
	});
});

describe('what the designer writes alongside a property (issue #138)', () => {
	it('keeps TakeFocusOnClick False as PropMask bit 9 and prints it back', () => {
		const wb = project();
		const markup = addToForm(wb, `    <CommandButton Name="Btn1" Left="8" Top="8" Width="72" Height="24" TakeFocusOnClick="False" />`);
		expect(markup).toContain('Name="Btn1"');
		expect(markup).toContain('TakeFocusOnClick="False"');
		const { record } = find(formPackage(wb), 'Btn1');
		expect((record!.maskLo & (1 << 9)) !== 0).toBe(true);
	});

	it('keeps TripleState in the MultiSelect field and prints it back', () => {
		const wb = project();
		const markup = addToForm(wb, `    <CheckBox Name="Chk1" Left="8" Top="8" Width="90" Height="15" TripleState="True" />`);
		expect(markup).toContain('TripleState="True"');
		expect(find(formPackage(wb), 'Chk1').record?.values.get('MultiSelect')).toBe(1);
	});

	it('disables the font with the control, except on a ListBox', () => {
		const wb = project();
		addToForm(wb, `    <TextBox Name="Txt1" Left="8" Top="8" Width="96" Height="18" Enabled="False" />${CRLF}    <ListBox Name="Lst1" Left="8" Top="30" Width="96" Height="60" Enabled="False" />`);
		const pkg = formPackage(wb);
		expect(find(pkg, 'Txt1').record?.textProps?.values.get('FontEffects')).toBe(0x40002000);
		expect(find(pkg, 'Lst1').record?.textProps?.values.get('FontEffects')).toBeUndefined();
	});

	it('sets auto colour with a font style and weight 700 with bold', () => {
		const wb = project();
		addToForm(wb, `    <ToggleButton Name="Tgl1" Left="8" Top="8" Width="72" Height="24" Font.Bold="True" />`);
		const tp = find(formPackage(wb), 'Tgl1').record?.textProps;
		expect(tp?.values.get('FontEffects')).toBe(0x40000001);
		expect(tp?.values.get('FontWeight')).toBe(700);
	});

	it('clears SpecialEffect for a BorderStyle and BorderStyle for a SpecialEffect', () => {
		const wb = project();
		addToForm(wb, `    <TextBox Name="Txt2" Left="8" Top="8" Width="96" Height="18" BorderStyle="1" />${CRLF}    <Image Name="Img1" Left="8" Top="30" Width="72" Height="54" SpecialEffect="2" />`);
		const pkg = formPackage(wb);
		expect(find(pkg, 'Txt2').record?.values.get('SpecialEffect')).toBe(0);
		expect(find(pkg, 'Img1').record?.values.get('BorderStyle')).toBe(0);
	});

	it('disables both arrows of a ScrollBar and lets Min pull Position up', () => {
		const wb = project();
		addToForm(wb, `    <ScrollBar Name="Scr1" Left="8" Top="8" Width="13" Height="90" Enabled="False" />${CRLF}    <SpinButton Name="Spn1" Left="30" Top="8" Width="13" Height="36" Min="10" Position="5" />`);
		const pkg = formPackage(wb);
		const scroll = find(pkg, 'Scr1').record!;
		expect(scroll.values.get('PrevEnabled')).toBe(0);
		expect(scroll.values.get('NextEnabled')).toBe(0);
		expect(find(pkg, 'Spn1').record?.values.get('Position')).toBe(10);
	});

	it('stores the form scroll bars as bars plus KeepScrollBarsVisible and prints both', () => {
		const wb = project();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		applyFormMarkup(wb, 'EntryForm', markup.replace(/<Form /, '<Form ScrollBars="3" KeepScrollBarsVisible="1" '));
		resetProjectCacheForTests();
		expect(formPackage(wb).form.record.values.get('ScrollBars')).toBe(3 | (1 << 2));
		const printed = readFormMarkup(wb, 'EntryForm').markup;
		expect(printed).toContain('ScrollBars="3"');
		expect(printed).toContain('KeepScrollBarsVisible="1"');
	});

	it('places a new page 1.5pt in and below the tab strip, sized to the MultiPage', () => {
		const wb = project();
		const markup = readFormMarkup(wb, 'EntryForm').markup;
		applyFormMarkup(wb, 'EntryForm', markup.replace(
			'<Page Name="Page2" Caption="Page2" />',
			`<Page Name="Page2" Caption="Page2" />${CRLF}        <Page Name="Details" Caption="Details" />`,
		));
		resetProjectCacheForTests();
		const pkg = formPackage(wb);
		const details = find(pkg, 'Details');
		// The MultiPage's first page is the witness the engine copies: same
		// origin, same size.
		const reference = find(pkg, 'Page1');
		expect(details.site.position).toEqual(reference.site.position);
		const detailsSize = details.owner.containers.get(siteId(details.site))!.form.record.sizes.get('DisplayedSize');
		const referenceSize = reference.owner.containers.get(siteId(reference.site))!.form.record.sizes.get('DisplayedSize');
		expect(detailsSize).toEqual(referenceSize);
	});
});
