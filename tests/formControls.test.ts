import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { ZipArchive } from '../src/vba/zip';

// Excel 16 authored ControlsFixture.xlsm with one of every form control,
// each wired to the cell or list it drives, plus one plain AutoShape so the
// drawing layer is not all controls.
//
// The round trip is pinned here because it is easy to lose by accident.
// pyOpenVBA's issue #25 is exactly that: its writer rebuilds a control's
// properties from a kind lookup, so an option button, spinner, scroll bar or
// group box read from a file and written back comes out a Button, and a
// spinner loses its maximum. XLIDE splices the existing formControlPr
// instead - it sets the attributes an edit names and leaves every other byte
// alone - and these tests are what say so.
const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'ControlsFixture.xlsm');

const open = (): XlsxWorkbook => XlsxWorkbook.fromBuffer(fs.readFileSync(FIXTURE));
const reopen = (book: XlsxWorkbook): XlsxWorkbook => XlsxWorkbook.fromBuffer(book.toBytes());
const sheet = (book: XlsxWorkbook) => book.shapes()[0];
const kindOf = (book: XlsxWorkbook): string[] =>
	sheet(book).shapes.map((s) => `${s.name}=${s.kind}`);

/** Every ctrlProp part's formControlPr tag, keyed by part name. */
function controlProperties(bytes: Buffer): Map<string, string> {
	const zip = ZipArchive.read(bytes);
	const out = new Map<string, string>();
	for (const name of zip.names()) {
		if (!name.includes('ctrlProps/')) { continue; }
		out.set(name, /<formControlPr\b[^>]*>/.exec(zip.read(name).toString('utf8'))?.[0] ?? '(none)');
	}
	return out;
}

/**
 * Re-link every control that drives a cell, which is the edit that rewrites
 * its ctrlProp part. Alt text and a rename do NOT: they change the sheet's
 * <controls> entry and leave the part alone, so an assertion built on one of
 * those would hold however the part was written.
 */
const LINKABLE = ['Tick', 'Untick', 'First', 'Second', 'Pick', 'Many', 'Step', 'Slide'];

function relinkEveryControl(book: XlsxWorkbook): void {
	const on = sheet(book).sheet;
	for (const [index, name] of LINKABLE.entries()) {
		book.editShape(on, { action: 'update', name, linkedCell: `$F$${index + 1}` });
	}
}

describe('reading every form control', () => {
	it('tells the nine kinds apart, as Excel does', () => {
		expect(kindOf(open())).toEqual([
			'Plain=shape',
			'Go=button',
			'Tick=checkBox',
			'Untick=checkBox',
			'First=optionButton',
			'Second=optionButton',
			'Pick=dropDown',
			'Many=listBox',
			'Step=spinner',
			'Slide=scrollBar',
			'Set=groupBox',
			'Caption=label',
		]);
	});

	it('reports what each control drives', () => {
		const shapes = sheet(open()).shapes;
		const named = (name: string) => shapes.find((s) => s.name === name);

		expect(named('Go')?.macro).toBe('Clicked');
		expect(named('Tick')?.linkedCell).toBe('$D$1');
		expect(named('Pick')).toMatchObject({ linkedCell: '$D$3', inputRange: '$H$1:$H$3' });
		expect(named('Many')).toMatchObject({ linkedCell: '$D$4', inputRange: '$H$1:$H$3' });
		expect(named('Step')?.linkedCell).toBe('$D$5');
		expect(named('Slide')?.linkedCell).toBe('$D$6');
	});
});

describe('editing a form control', () => {
	it('keeps every kind, where rebuilding the properties would lose four', () => {
		const book = open();
		relinkEveryControl(book);

		expect(kindOf(reopen(book))).toEqual(kindOf(open()));
	});

	it('changes only the attribute the edit names, on every rewritten part', () => {
		const before = controlProperties(fs.readFileSync(FIXTURE));
		const book = open();
		relinkEveryControl(book);
		const after = controlProperties(book.toBytes());

		// Every linkable control's part is genuinely rewritten, so nothing
		// below can hold just because the writer never ran.
		const rewritten = [...after].filter(([name, tag]) => before.get(name) !== tag);
		expect(rewritten).toHaveLength(LINKABLE.length);

		for (const [name, tag] of rewritten) {
			const was = before.get(name)!;
			const strip = (one: string) => one.replace(/\s*\bfmlaLink="[^"]*"/, '');
			expect(strip(tag), name).toBe(strip(was));
			expect(tag, name).toMatch(/\bfmlaLink="\$F\$\d+"/);
		}
	});

	it("keeps a spinner's and a scroll bar's maximum through a re-link", () => {
		// The concrete property a regenerating writer drops: a control that
		// cannot exceed zero sits at zero, in a file that opens cleanly.
		const book = open();
		relinkEveryControl(book);
		const tags = [...controlProperties(book.toBytes()).values()];

		expect(tags.some((tag) => /objectType="Spin"[^>]*\bmax="50"/.test(tag))).toBe(true);
		expect(tags.some((tag) => /objectType="Scroll"[^>]*\bmax="100"/.test(tag))).toBe(true);
	});

	it('still writes the properties an edit names', () => {
		const book = open();
		const on = sheet(book).sheet;
		book.editShape(on, { action: 'update', name: 'Step', linkedCell: '$F$9' });
		book.editShape(on, { action: 'update', name: 'Pick', inputRange: '$H$1:$H$2' });

		const after = sheet(reopen(book)).shapes;
		expect(after.find((s) => s.name === 'Step')).toMatchObject({ kind: 'spinner', linkedCell: '$F$9' });
		expect(after.find((s) => s.name === 'Pick')).toMatchObject({ kind: 'dropDown', inputRange: '$H$1:$H$2' });
	});

	it('renames and deletes a control of any kind', () => {
		const book = open();
		const on = sheet(book).sheet;
		book.editShape(on, { action: 'update', name: 'Slide', newName: 'Slider' });
		book.editShape(on, { action: 'delete', name: 'Set' });

		const after = sheet(reopen(book)).shapes;
		expect(after.find((s) => s.name === 'Slider')?.kind).toBe('scrollBar');
		expect(after.some((s) => s.name === 'Set')).toBe(false);
	});
});
