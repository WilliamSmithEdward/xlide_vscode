import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { Cfb } from '../src/vba/cfb';
import { XlsxWorkbook } from '../src/vba/xlsx';
import {
	parseFormStream,
	serializeFormStream,
	parseObjectStream,
	serializeObjectStream,
	siteName,
} from '../src/vba/oforms/formStream';
import { parseFormPackage, writeFormPackage } from '../src/vba/oforms/formPackage';
import { himetricToPoints, pointsToHimetric } from '../src/vba/oforms/bytes';

// The [MS-OFORMS] engine's contract is BYTE IDENTITY: parse a form Office
// wrote, serialize it back, get the same bytes. Padding inside a DataBlock is
// undefined ([MS-OFORMS] 2.1.1.2.4) - Excel writes whatever was in its buffer
// - so the engine captures and replays it rather than recomputing. These four
// fixtures were authored by live Excel, Word, and PowerPoint; between them
// they cover every standard control kind, fonts, two multi-hundred-KB
// pictures, a Frame, and a MultiPage with two Pages in nested storages.

const FIXTURES = 'tests/fixtures/binaries';

const FORMS: ReadonlyArray<readonly [string, string]> = [
	[path.join(FIXTURES, 'FormFixture.xlsm'), 'FrmPicker'],
	[path.join(FIXTURES, 'FormFixtureVbide.xlsm'), 'EntryForm'],
	[path.join(FIXTURES, 'WordFormFixture.docm'), 'FrmNotice'],
	[path.join(FIXTURES, 'PowerPointFormFixture.pptm'), 'FrmDeck'],
];

function openVba(book: string): Cfb {
	return Cfb.fromBytes(XlsxWorkbook.fromBuffer(fs.readFileSync(book)).readVbaProject());
}

function allStreams(cfb: Cfb, root: string): Map<string, Buffer> {
	const out = new Map<string, Buffer>();
	const walk = (p: string[]): void => {
		for (const s of cfb.listStreamsAtPath(p)) {
			out.set([...p, s].join('/'), cfb.getStreamAtPath(p, s));
		}
		for (const c of cfb.listStoragesAtPath(p)) { walk([...p, c]); }
	};
	walk([root]);
	return out;
}

describe('form stream byte identity', () => {
	for (const [book, form] of FORMS) {
		it(`round-trips ${form} exactly`, () => {
			const cfb = openVba(book);
			const f = cfb.getStreamAtPath([form], 'f');
			const o = cfb.getStreamAtPath([form], 'o');
			const model = parseFormStream(f);
			expect(serializeFormStream(model).equals(f)).toBe(true);
			const entries = parseObjectStream(o, model.sites);
			expect(serializeObjectStream(entries).equals(o)).toBe(true);
		});
	}

	it('round-trips every nested container storage of the 19-control form', () => {
		const cfb = openVba(path.join(FIXTURES, 'FormFixtureVbide.xlsm'));
		// Frame, MultiPage, and both Pages - including the MultiPage's own f
		// whose sites are its Pages and whose o holds its TabStrip record.
		for (const p of [['EntryForm', 'i06'], ['EntryForm', 'i10'], ['EntryForm', 'i10', 'i12'], ['EntryForm', 'i10', 'i13']]) {
			const f = cfb.getStreamAtPath(p, 'f');
			const o = cfb.getStreamAtPath(p, 'o');
			const model = parseFormStream(f);
			expect(serializeFormStream(model).equals(f), `${p.join('/')} f`).toBe(true);
			const entries = parseObjectStream(o, model.sites);
			expect(serializeObjectStream(entries).equals(o), `${p.join('/')} o`).toBe(true);
		}
	});
});

describe('whole-package recursive round trip', () => {
	for (const [book, form] of FORMS) {
		it(`writes ${form} back byte-identically across every stream`, () => {
			const cfb = openVba(book);
			const before = allStreams(cfb, form);
			writeFormPackage(cfb, [form], parseFormPackage(cfb, [form]));
			const after = allStreams(cfb, form);
			expect(after.size).toBe(before.size);
			for (const [key, bytes] of before) {
				expect(after.get(key)?.equals(bytes), key).toBe(true);
			}
		});
	}
});

describe('the parse agrees with the proven control reader', () => {
	it('sees the same controls formDesigner.ts reports for FrmPicker', () => {
		const cfb = openVba(path.join(FIXTURES, 'FormFixture.xlsm'));
		const model = parseFormStream(cfb.getStreamAtPath(['FrmPicker'], 'f'));
		expect(model.sites.map(siteName)).toEqual(['RegionPick', 'Taxable', 'NameBox', 'OkButton']);
	});
});

describe('HIMETRIC conversion', () => {
	it('round-trips whole points exactly', () => {
		for (const pt of [0, 12, 66, 228, 291]) {
			expect(himetricToPoints(pointsToHimetric(pt))).toBeCloseTo(pt, 1);
		}
	});
});
