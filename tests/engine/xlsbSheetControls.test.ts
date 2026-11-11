// The ActiveX controls of an .xlsb's worksheets, read from their binary
// parts (issue #1555). ExcelControlsDisabledSaveFixture.xlsb is an Excel 16.0
// file with eleven controls on Sheet1, one on Data and a chart sheet;
// SheetsFixture.xlsb has five sheets and no control.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { xlsbSheetControls } from '../../src/vba/workbookSheets';
import { ZipArchive } from '../../src/vba/zip';

const FIXTURES = path.join(__dirname, '..', 'fixtures', 'binaries');
const SHEET1 = ['CtlCommandButton', 'CtlTextBox', 'CtlCheckBox', 'CtlOptionButton', 'CtlToggleButton', 'CtlComboBox', 'CtlListBox', 'CtlLabel', 'CtlImage', 'CtlScrollBar', 'CtlSpinButton'];

function zipOf(file: string): ZipArchive {
	return ZipArchive.read(fs.readFileSync(path.join(FIXTURES, file)));
}

/** The worksheet part that holds a code name. */
function partOf(zip: ZipArchive, codeName: string): string {
	const parts = zip.names().filter((name) => /^xl\/worksheets\/[^/]+\.bin$/.test(name));
	const part = parts.find((name) => zip.read(name).includes(Buffer.from(codeName, 'utf16le')));
	expect(part, codeName).toBeDefined();
	return part!;
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('the ActiveX controls of an .xlsb worksheet', () => {
	it('lists each sheet\'s controls in the order Excel wrote them, by code name', () => {
		const controls = xlsbSheetControls(zipOf('ExcelControlsDisabledSaveFixture.xlsb'));
		expect([...controls.keys()].sort()).toEqual(['data', 'sheet1']);
		expect(controls.get('sheet1')).toEqual(SHEET1);
		expect(controls.get('data')).toEqual(['DataBtn']);
	});

	it('gives a worksheet with no control an empty list, reading no more of it than its code name needs', () => {
		const zip = zipOf('SheetsFixture.xlsb');
		const read = vi.spyOn(ZipArchive.prototype, 'read');
		const prefix = vi.spyOn(ZipArchive.prototype, 'readPrefix');
		const controls = xlsbSheetControls(zip);
		expect(controls.size).toBeGreaterThan(0);
		expect([...controls.values()].every((names) => names.length === 0)).toBe(true);
		// The head read for the code name reads a part this small whole, and
		// nothing else may.
		const sheetParts = (calls: unknown[][]) => calls.map(([name]) => String(name)).filter((name) => /^xl\/worksheets\/[^/]+\.bin$/.test(name)).sort();
		expect(sheetParts(prefix.mock.calls).length).toBeGreaterThan(0);
		expect(sheetParts(read.mock.calls)).toEqual(sheetParts(prefix.mock.calls));
	});

	it('finds the control relationship whatever the element\'s prefix, the attribute\'s quotes or a character reference in it', () => {
		const zip = zipOf('ExcelControlsDisabledSaveFixture.xlsb');
		const rels = `xl/worksheets/_rels/${partOf(zip, 'Sheet1').split('/').pop()}.rels`;
		const xml = zip.read(rels).toString('utf8');
		const prefixed = xml.replace(/<(\/?)Relationship(s?)\b/g, '<$1r:Relationship$2').replace('<r:Relationships ', '<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships" ')
			.replace(/Type="([^"]*)"/g, "Type='$1'");
		expect(prefixed).not.toBe(xml);
		zip.write(rels, Buffer.from(prefixed, 'utf8'));
		expect(xlsbSheetControls(zip).get('sheet1')).toEqual(SHEET1);
		// A character reference in the type names the same relationship.
		const referenced = prefixed.replace(/relationships\/control'/g, "relationships/contr&#111;l'");
		expect(referenced).not.toBe(prefixed);
		zip.write(rels, Buffer.from(referenced, 'utf8'));
		expect(xlsbSheetControls(zip).get('sheet1')).toEqual(SHEET1);
	});

	it('leaves out a sheet whose control block is never closed', () => {
		const zip = zipOf('ExcelControlsDisabledSaveFixture.xlsb');
		const part = partOf(zip, 'Sheet1');
		const whole = Buffer.from(zip.read(part));
		// BrtEndActiveXControls (0x0285, written 85 05) with no body, made another record.
		const end = whole.indexOf(Buffer.from([0x85, 0x05, 0x00]));
		expect(end).toBeGreaterThan(0);
		whole[end] = 0x86;
		zip.write(part, whole);
		const controls = xlsbSheetControls(zip);
		expect(controls.has('sheet1')).toBe(false);
		expect(controls.get('data')).toEqual(['DataBtn']);
	});

	it('leaves out a sheet whose part ends early, and still lists the others', () => {
		const zip = zipOf('ExcelControlsDisabledSaveFixture.xlsb');
		const part = partOf(zip, 'Sheet1');
		const whole = zip.read(part);
		// Cut inside the control block, which the part's last records follow.
		const block = whole.indexOf(Buffer.from('CtlTextBox', 'utf16le'));
		expect(block).toBeGreaterThan(0);
		zip.write(part, whole.subarray(0, block));
		const controls = xlsbSheetControls(zip);
		expect(controls.has('sheet1')).toBe(false);
		expect(controls.get('data')).toEqual(['DataBtn']);
	});
});
