// How often a worksheet part is inflated to list its shapes, which listing a
// workbook's modules does for every sheet. A sheet part holds the sheet's
// rows, so inflating it costs in proportion to the data on the sheet, and it
// used to be inflated three times: for its code name, for its drawing
// references and for its controls. The counts here are of whole-part reads,
// which do not depend on the machine.

import { afterEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { ZipArchive } from '../src/vba/zip';
import { openMacroContainer } from '../src/vba/macroContainer';
import { readModulesFromBuffer } from '../src/vba/projectService';

const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const TEMPLATES = path.join(__dirname, '..', 'assets', 'templates');
const SHEET1 = 'xl/worksheets/sheet1.xml';
const SHEET2 = 'xl/worksheets/sheet2.xml';

/**
 * The fixture with one sheet grown past the head that is read for a code
 * name. The rows do not repeat, so the part stays large compressed.
 */
function largeSheetWorkbook(fixture: string, part = SHEET1, change: (xml: string) => string = (xml) => xml): Buffer {
	const zip = ZipArchive.read(fs.readFileSync(path.join(FIXTURES, fixture)));
	const rows = Array.from({ length: 60000 }, (_, i) =>
		`<row r="${i + 2}"><c r="A${i + 2}"><v>${(i * 7919) % 100003}</v></c><c r="B${i + 2}" t="str"><v>item ${i.toString(36)}</v></c></row>`);
	const xml = change(zip.read(part).toString('utf8')
		.replace(/<sheetData\b[^>]*\/>|<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${rows.join('')}</sheetData>`));
	expect(xml.length).toBeGreaterThan(2_000_000);
	zip.write(part, Buffer.from(xml, 'utf8'));
	return zip.toBytes();
}

/** How many times `run` inflates the whole of a part. */
function wholeReads(part: string, run: () => void): number {
	const read = vi.spyOn(ZipArchive.prototype, 'read');
	run();
	return read.mock.calls.filter(([name]) => name === part).length;
}

/** A sheet's code name as it was read before: the first sheetPr anywhere in the whole part. */
function codeNameOfWholePart(data: Buffer, part: string): string | undefined {
	const sheetPr = /<sheetPr\b[^>]*>/.exec(ZipArchive.read(data).read(part).toString('utf8'))?.[0];
	return sheetPr ? /\bcodeName="([^"]*)"/.exec(sheetPr)?.[1] || undefined : undefined;
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('how often a worksheet part is read', () => {
	it('reads a sheet once to list its shapes', () => {
		for (const fixture of ['SheetsFixture.xlsm', 'ShapesFixture.xlsm']) {
			const xlsx = openMacroContainer(largeSheetWorkbook(fixture)).xlsx!;
			expect(wholeReads(SHEET1, () => { xlsx.shapes(); }), fixture).toBe(1);
			vi.restoreAllMocks();
		}
	});

	it('reads a sheet once to list the workbook\'s modules, and still lists its module as a Worksheet', () => {
		for (const fixture of ['SheetsFixture.xlsm', 'ShapesFixture.xlsm']) {
			const data = largeSheetWorkbook(fixture);
			let entries: ReturnType<typeof readModulesFromBuffer> = [];
			expect(wholeReads(SHEET1, () => { entries = readModulesFromBuffer(data); }), fixture).toBe(1);
			const sheet = entries.find((entry) => entry.name === 'Sheet1');
			expect(sheet?.designerClass, fixture).toBe('Excel.Worksheet');
			expect(sheet?.implicitMembers, fixture).toEqual([]);
			vi.restoreAllMocks();
		}
	});

	// Sheet2 of ShapesFixture has no sheetPr. Its head cannot show that none
	// follows, so the part is read whole for the code name as well.
	it('reads a sheet with no sheetPr once more, to look for one', () => {
		const data = largeSheetWorkbook('ShapesFixture.xlsm', SHEET2);
		expect(codeNameOfWholePart(data, SHEET2)).toBeUndefined();
		const xlsx = openMacroContainer(data).xlsx!;
		let listed: ReturnType<typeof xlsx.shapes> = [];
		expect(wholeReads(SHEET2, () => { listed = xlsx.shapes(); })).toBe(2);
		expect(listed[1].codeName).toBeUndefined();
	});
});

describe('a sheet\'s code name, read from the head of its part', () => {
	const workbooks = [FIXTURES, TEMPLATES].flatMap((dir) => fs.readdirSync(dir)
		.filter((name) => /\.(xlsm|xlam|xltm)$/i.test(name))
		.map((name) => path.join(dir, name)));

	it('covers the fixtures', () => {
		expect(workbooks.length).toBeGreaterThan(10);
	});

	it.each(workbooks.map((file) => [path.basename(file), file]))('is the one the whole part gives: %s', (_name, file) => {
		const data = fs.readFileSync(file);
		const xlsx = openMacroContainer(data).xlsx;
		if (!xlsx?.hasSheetSurface()) {
			return;
		}
		const zip = ZipArchive.read(data);
		const parts = zip.names().filter((name) => /^xl\/worksheets\/[^/]+\.xml$/.test(name));
		const whole = parts.map((part) => codeNameOfWholePart(data, part)).filter((name) => name !== undefined).sort();
		const listed = xlsx.shapes().map((sheet) => sheet.codeName).filter((name) => name !== undefined).sort();
		expect(listed).toEqual(whole);
	});

	// The same answer on a sheet too large for its head to hold it all,
	// whatever the sheetPr looks like and wherever it is.
	it.each<[string, (xml: string) => string, string | undefined]>([
		['as Excel writes it', (xml) => xml, 'Sheet1'],
		['with child elements', (xml) => xml.replace('<sheetPr codeName="Sheet1"/>', '<sheetPr codeName="Sheet1"><pageSetUpPr fitToPage="1"/></sheetPr>'), 'Sheet1'],
		['with an empty code name', (xml) => xml.replace('codeName="Sheet1"', 'codeName=""'), undefined],
		['with no code name', (xml) => xml.replace('<sheetPr codeName="Sheet1"/>', '<sheetPr/>'), undefined],
		['with no sheetPr', (xml) => xml.replace('<sheetPr codeName="Sheet1"/>', ''), undefined],
		['with an entity in the value', (xml) => xml.replace('codeName="Sheet1"', 'codeName="A&amp;B"'), 'A&amp;B'],
		['with a prefixed attribute', (xml) => xml.replace('codeName="Sheet1"', 'x:codeName="Sheet1"'), 'Sheet1'],
		['with the attribute twice', (xml) => xml.replace('codeName="Sheet1"', 'codeName="Sheet1" codeName="Second"'), 'Sheet1'],
		// Past the rows, where the schema does not allow it and the head does not reach.
		['after the rows', (xml) => xml.replace('<sheetPr codeName="Sheet1"/>', '').replace('</sheetData>', '</sheetData><sheetPr codeName="Sheet1"/>'), 'Sheet1'],
	])('on a large sheet, a sheetPr %s', (name, change, expected) => {
		const data = largeSheetWorkbook('SheetsFixture.xlsm', SHEET1, (xml) => {
			const changed = change(xml);
			expect(changed !== xml).toBe(name !== 'as Excel writes it');
			return changed;
		});
		expect(codeNameOfWholePart(data, SHEET1)).toBe(expected);
		expect(openMacroContainer(data).xlsx!.shapes()[0].codeName).toBe(expected);
	});
});
