// A workbook's own word on its sheet modules (issue #225): a worksheet's module
// is listed as an Excel.Worksheet whose members include the sheet's ActiveX
// controls, so its surface can prove a name absent.

import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { ZipArchive } from '../src/vba/zip';
import { readModulesFromBuffer } from '../src/vba/projectService';

const FIXTURE = path.join(__dirname, 'fixtures', 'binaries', 'ShapesFixture.xlsm');

/** ShapesFixture with an ActiveX CommandButton1 added to Sheet1, in the layout Excel writes. */
function withActiveX(): Buffer {
	const zip = ZipArchive.read(fs.readFileSync(FIXTURE));
	const text = (name: string): string => zip.read(name).toString('utf8');
	const put = (name: string, value: string): void => zip.write(name, Buffer.from(value, 'utf8'));
	put('xl/worksheets/sheet1.xml', text('xl/worksheets/sheet1.xml').replace('<controls>',
		'<controls><control shapeId="2049" r:id="rIdAx1" name="CommandButton1"/>'));
	put('xl/worksheets/_rels/sheet1.xml.rels', text('xl/worksheets/_rels/sheet1.xml.rels').replace('</Relationships>',
		'<Relationship Id="rIdAx1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml"/>'
		+ '</Relationships>'));
	return zip.toBytes();
}

describe("a worksheet's module (issue #225)", () => {
	it('is an Excel.Worksheet with its ActiveX controls as members', () => {
		const sheet = readModulesFromBuffer(withActiveX()).find((entry) => entry.name === 'Sheet1');
		expect(sheet?.designerClass).toBe('Excel.Worksheet');
		expect(sheet?.implicitMembers).toEqual([{ name: 'CommandButton1', type: 'Object' }]);
	});

	it('has no members from its form controls, and ThisWorkbook gains nothing', () => {
		const entries = readModulesFromBuffer(fs.readFileSync(FIXTURE));
		expect(entries.find((entry) => entry.name === 'Sheet1')?.implicitMembers).toEqual([]);
		const book = entries.find((entry) => entry.name === 'ThisWorkbook');
		expect(book?.designerClass).toBeUndefined();
		expect(book?.implicitMembers).toBeUndefined();
	});
});
