import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { XlsxWorkbook } from '../src/vba/xlsx';
import { ZipArchive } from '../src/vba/zip';
import { editShape, listShapes, writeModule } from '../src/vba/projectService';
import type { ShapeInfo } from '../src/vba/xlsxShapes';

// Excel 16 saved both fixtures. ShapesFixture has, on Sheet1, a rectangle
// running DoIt, a rounded rectangle running Macros.Other, an oval with alt
// text, a text box, a line, a group of two, and three form controls: a button
// running DoIt, a check box linked to $H$10 and a drop-down listing
// $J$1:$J$3 into $H$11; Sheet2 has one rectangle. ShapesPictureChartFixture
// has a picture and a chart that run DoIt. Every kind of edit here was also
// checked by having Excel open the file the engine wrote and report its shapes.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const BLANK = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');

const open = (name: string): XlsxWorkbook => XlsxWorkbook.fromBuffer(fs.readFileSync(path.join(FIXTURES, name)));
const reopen = (book: XlsxWorkbook): XlsxWorkbook => XlsxWorkbook.fromBuffer(book.toBytes());
const part = (book: XlsxWorkbook, name: string): string => ZipArchive.read(book.toBytes()).read(name).toString('utf8');
const hasPart = (book: XlsxWorkbook, name: string): boolean => ZipArchive.read(book.toBytes()).has(name);
const shapesOn = (book: XlsxWorkbook, sheet: string): ShapeInfo[] => book.shapes(sheet)[0].shapes;
const shape = (book: XlsxWorkbook, sheet: string, name: string): ShapeInfo | undefined =>
	shapesOn(book, sheet).find((s) => s.name === name);

describe('listing shapes', () => {
	it('reports each shape as Excel does: kind, cells, macro, text and links', () => {
		const shapes = shapesOn(open('ShapesFixture.xlsm'), 'Sheet1');

		expect(shapes).toEqual([
			{ name: 'RunButton', kind: 'shape', geometry: 'rect', range: 'B3:D5', macro: 'DoIt', text: 'Run it' },
			{ name: 'Rectangle: Rounded Corners 2', kind: 'shape', geometry: 'roundRect', range: 'E3:G5', macro: 'Macros.Other' },
			{ name: 'Oval 3', kind: 'shape', geometry: 'ellipse', range: 'G3:H5', altText: 'an oval' },
			{ name: 'TextBox 4', kind: 'textBox', range: 'B7:E9', text: 'A note' },
			{ name: 'Straight Connector 5', kind: 'line', range: 'B11:F11' },
			{
				name: 'Pair', kind: 'group', range: 'G9:I10', shapes: [
					{ name: 'Rectangle 6', kind: 'shape', geometry: 'rect' },
					{ name: 'Oval 7', kind: 'shape', geometry: 'ellipse' },
				],
			},
			{ name: 'Button 1', kind: 'button', range: 'B14:D16', macro: 'DoIt', text: 'Go' },
			{ name: 'Check Box 2', kind: 'checkBox', range: 'D14:F16', linkedCell: '$H$10', text: 'Check Box 2' },
			{ name: 'Drop Down 3', kind: 'dropDown', range: 'G14:I16', linkedCell: '$H$11', inputRange: '$J$1:$J$3' },
		]);
	});

	it('lists every sheet when none is named, and pictures and charts with their macros', () => {
		expect(open('ShapesFixture.xlsm').shapes().map((s) => [s.sheet, s.shapes.length])).toEqual([['Sheet1', 9], ['Sheet2', 1]]);
		const shapes = shapesOn(open('ShapesPictureChartFixture.xlsm'), 'Sheet1');
		expect(shapes.filter((s) => s.kind === 'picture' || s.kind === 'chart')).toEqual([
			{ name: 'Picture 4', kind: 'picture', range: 'A6:B9', macro: 'DoIt' },
			{ name: 'Chart 5', kind: 'chart', range: 'C6:F13', macro: 'DoIt' },
		]);
	});

	it('lists nothing on a sheet without shapes, and refuses a sheet that is not there', () => {
		expect(XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK)).shapes()).toEqual([{ sheet: 'Sheet1', shapes: [] }]);
		expect(() => open('ShapesFixture.xlsm').shapes('Nope')).toThrow(/Worksheet not found: Nope/);
	});
});

describe('changing a shape', () => {
	it('moves, relabels, relinks and renames an AutoShape', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'update', name: 'runbutton', range: 'B20:C21', text: 'Two\nlines', macro: 'Macros.Other', altText: 'runs Other', newName: 'Runner' });

		expect(shape(reopen(book), 'Sheet1', 'Runner')).toEqual({
			name: 'Runner', kind: 'shape', geometry: 'rect', range: 'B20:C21', macro: 'Macros.Other', text: 'Two\nlines', altText: 'runs Other',
		});
		// The first run's format carries to the new text.
		expect(part(book, 'xl/drawings/drawing1.xml')).toMatch(/<a:r><a:rPr[^>]*\/?>(?:<\/a:rPr>)?<a:t>Two<\/a:t><\/a:r><\/a:p><a:p>(?:<a:pPr[^>]*\/>)?<a:r><a:rPr/);
	});

	it('removes a macro and alt text given as empty', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'update', name: 'RunButton', macro: '' });
		book.editShape('Sheet1', { action: 'update', name: 'Oval 3', altText: '' });

		expect(shape(book, 'Sheet1', 'RunButton')?.macro).toBeUndefined();
		expect(shape(book, 'Sheet1', 'Oval 3')?.altText).toBeUndefined();
		expect(part(book, 'xl/drawings/drawing1.xml')).toContain('<xdr:sp macro="" textlink="">');
	});

	it('changes a form control in all four places Excel keeps it', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'update', name: 'Button 1', text: 'Run', macro: 'Other', range: 'B18:C19', newName: 'Runner' });

		expect(shape(reopen(book), 'Sheet1', 'Runner')).toEqual({ name: 'Runner', kind: 'button', range: 'B18:C19', macro: 'Other', text: 'Run' });
		const vml = part(book, 'xl/drawings/vmlDrawing1.vml');
		expect(vml).toContain('<x:FmlaMacro>Other</x:FmlaMacro>');
		expect(vml).toMatch(/<x:Anchor>\s*1, 0, 17, 0, 3, 0, 19, 0<\/x:Anchor>/);
		const sheet = part(book, 'xl/worksheets/sheet1.xml');
		expect(sheet).toMatch(/<control shapeId="1025" r:id="rId3" name="Runner"><controlPr [^>]*macro="Other"/);
		// The hidden DrawingML twin Excel 2010 and later draw the button from.
		const drawing = part(book, 'xl/drawings/drawing1.xml');
		expect(drawing).toMatch(/<xdr:cNvPr id="1025" name="Runner" hidden="1"/);
		expect(drawing).toMatch(/<a:t>Run<\/a:t>/);
	});

	it('relinks a check box and a drop-down, on this sheet or another, and unlinks them', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'update', name: 'Check Box 2', linkedCell: 'sheet2!b4' });
		book.editShape('Sheet1', { action: 'update', name: 'Drop Down 3', linkedCell: 'Sheet1!K1', inputRange: 'L1:L5' });

		expect(shape(book, 'Sheet1', 'Check Box 2')?.linkedCell).toBe('Sheet2!$B$4');
		expect(shape(book, 'Sheet1', 'Drop Down 3')).toMatchObject({ linkedCell: '$K$1', inputRange: '$L$1:$L$5' });
		expect(part(book, 'xl/ctrlProps/ctrlProp3.xml')).toMatch(/fmlaLink="\$K\$1" fmlaRange="\$L\$1:\$L\$5"|fmlaRange="\$L\$1:\$L\$5"[^>]*fmlaLink="\$K\$1"|fmlaLink="\$K\$1"[^>]*fmlaRange="\$L\$1:\$L\$5"/);
		expect(part(book, 'xl/drawings/vmlDrawing1.vml')).toContain('<x:FmlaRange>$L$1:$L$5</x:FmlaRange>');

		book.editShape('Sheet1', { action: 'update', name: 'Check Box 2', linkedCell: '' });
		expect(shape(book, 'Sheet1', 'Check Box 2')?.linkedCell).toBeUndefined();
		expect(part(book, 'xl/ctrlProps/ctrlProp2.xml')).not.toContain('fmlaLink');
		expect(part(book, 'xl/drawings/vmlDrawing1.vml')).not.toContain('$H$10');
	});

	it('quotes another sheet\'s name in a link when a formula would', () => {
		const book = open('ShapesFixture.xlsm');
		const zip = ZipArchive.read(book.toBytes());
		zip.write('xl/workbook.xml', Buffer.from(zip.read('xl/workbook.xml').toString('utf8').replace('name="Sheet2"', 'name="My Data"'), 'utf8'));
		const renamed = XlsxWorkbook.fromBuffer(zip.toBytes());
		renamed.editShape('Sheet1', { action: 'update', name: 'Check Box 2', linkedCell: "'my data'!$A$1" });

		expect(shape(renamed, 'Sheet1', 'Check Box 2')?.linkedCell).toBe("'My Data'!$A$1");
	});

	it('refuses what the shape cannot take, and changes nothing', () => {
		const book = open('ShapesFixture.xlsm');
		const before = book.toBytes();
		const refusals: Array<[Parameters<XlsxWorkbook['editShape']>[1], RegExp]> = [
			[{ action: 'update', name: 'Missing' }, /No shape named 'Missing' on sheet 'Sheet1'/],
			[{ action: 'update' }, /give its name; xlide_listShapes lists them/],
			[{ action: 'update', name: 'Straight Connector 5', text: 'x' }, /is a line, which holds no text/],
			[{ action: 'update', name: 'Pair', macro: 'DoIt' }, /is a group, which Excel runs no macro for/],
			[{ action: 'update', name: 'Rectangle 6', range: 'A1:B2' }, /is in a group; move the group instead/],
			[{ action: 'delete', name: 'Oval 7' }, /is in a group; delete the group/],
			[{ action: 'update', name: 'RunButton', linkedCell: '$A$1' }, /is not a form control, so it has no cell link/],
			[{ action: 'update', name: 'Button 1', linkedCell: '$A$1' }, /is a button, which has no cell link/],
			[{ action: 'update', name: 'Check Box 2', inputRange: '$A$1:$A$3' }, /only a drop-down or list box has an input range/],
			[{ action: 'update', name: 'Drop Down 3', text: 'x' }, /is a dropDown, which has no caption/],
			[{ action: 'update', name: 'Check Box 2', linkedCell: 'A1:A2' }, /a cell link is one cell/],
			[{ action: 'update', name: 'Check Box 2', linkedCell: 'Nowhere!A1' }, /The workbook has no sheet named 'Nowhere'/],
			[{ action: 'update', name: 'Check Box 2', linkedCell: 'XFE1' }, /runs past XFD1048576/],
			[{ action: 'update', name: 'Check Box 2', linkedCell: 'not a cell' }, /is not a cell link; give cells such as \$H\$10/],
			[{ action: 'update', name: 'RunButton', range: 'B2:' }, /is not a range of cells/],
			[{ action: 'update', name: 'RunButton', range: 'A1:A1048577' }, /runs past XFD1048576/],
			[{ action: 'update', name: 'RunButton', newName: 'oval 3' }, /already has a shape named 'oval 3'/],
			[{ action: 'add', range: 'A1:B2' }, /needs its type/],
			[{ action: 'add', type: 'oval' }, /needs the cells it covers/],
			[{ action: 'add', type: 'heart' as never, range: 'A1' }, /'heart' is not a shape XLIDE adds/],
			[{ action: 'add', type: 'toString' as never, range: 'A1' }, /'toString' is not a shape XLIDE adds/],
			[{ action: 'add', type: 'button', range: 'A1', name: 'go', newName: 'x' }, /takes its name from name/],
			[{ action: 'add', type: 'oval', range: 'A1', name: 'button 1' }, /already has a shape named 'button 1'/],
		];
		for (const [edit, message] of refusals) {
			expect(() => book.editShape('Sheet1', edit), JSON.stringify(edit)).toThrow(message);
		}
		expect(book.toBytes().equals(before)).toBe(true);
	});
});

describe('deleting a shape', () => {
	it('removes a form control from all four places, and the parts it alone used', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'delete', name: 'Check Box 2' });

		expect(shapesOn(reopen(book), 'Sheet1').map((s) => s.name)).not.toContain('Check Box 2');
		expect(hasPart(book, 'xl/ctrlProps/ctrlProp2.xml')).toBe(false);
		expect(part(book, '[Content_Types].xml')).not.toContain('ctrlProp2.xml');
		expect(part(book, 'xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('ctrlProp2.xml');
		expect(part(book, 'xl/worksheets/sheet1.xml')).not.toContain('shapeId="1026"');
		expect(part(book, 'xl/drawings/vmlDrawing1.vml')).not.toContain('_x0000_s1026');
		expect(part(book, 'xl/drawings/drawing1.xml')).not.toContain('id="1026"');
	});

	it('removes the VML part and <controls> with the last form control', () => {
		const book = open('ShapesFixture.xlsm');
		for (const name of ['Button 1', 'Check Box 2', 'Drop Down 3']) {
			book.editShape('Sheet1', { action: 'delete', name });
		}

		const sheet = part(book, 'xl/worksheets/sheet1.xml');
		expect(sheet).not.toMatch(/<legacyDrawing\b|<controls>|<mc:AlternateContent/);
		expect(sheet).toContain('<drawing r:id="rId1"/>');
		expect(hasPart(book, 'xl/drawings/vmlDrawing1.vml')).toBe(false);
		expect(part(book, 'xl/worksheets/_rels/sheet1.xml.rels')).not.toMatch(/vmlDrawing|ctrlProp/);
		expect(shapesOn(reopen(book), 'Sheet1')).toHaveLength(6);
	});

	it('removes the drawing part with the last shape on a sheet', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet2', { action: 'delete', name: 'Rectangle 1' });

		expect(part(book, 'xl/worksheets/sheet2.xml')).not.toContain('<drawing');
		expect(hasPart(book, 'xl/drawings/drawing2.xml')).toBe(false);
		expect(hasPart(book, 'xl/worksheets/_rels/sheet2.xml.rels')).toBe(false);
		expect(part(book, '[Content_Types].xml')).not.toContain('drawing2.xml');
		expect(reopen(book).shapes('Sheet2')).toEqual([{ sheet: 'Sheet2', shapes: [] }]);
	});

	it('removes a picture\'s image and a chart\'s part with them', () => {
		const book = open('ShapesPictureChartFixture.xlsm');
		book.editShape('Sheet1', { action: 'delete', name: 'Picture 4' });
		expect(hasPart(book, 'xl/media/image1.png')).toBe(false);
		expect(hasPart(book, 'xl/charts/chart1.xml')).toBe(true);

		book.editShape('Sheet1', { action: 'delete', name: 'chart 5' });
		expect(hasPart(book, 'xl/charts/chart1.xml')).toBe(false);
		expect(part(book, '[Content_Types].xml')).not.toContain('chart1.xml');
		expect(hasPart(book, 'xl/drawings/_rels/drawing1.xml.rels')).toBe(false);
		expect(shapesOn(reopen(book), 'Sheet1').map((s) => s.name)).toEqual(['Rectangle 1', 'Rectangle 2', 'Button 1']);
	});

	it('removes a whole group', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'delete', name: 'Pair' });

		expect(shapesOn(book, 'Sheet1').map((s) => s.name)).not.toContain('Pair');
		expect(part(book, 'xl/drawings/drawing1.xml')).not.toContain('Rectangle 6');
	});
});

describe('adding a shape', () => {
	it('builds the drawing, VML and control parts a blank sheet needs', () => {
		const book = XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK));
		book.editShape('Sheet1', { action: 'add', type: 'roundedRectangle', range: 'B2:D4', text: 'Run', macro: 'DoIt' });
		book.editShape('Sheet1', { action: 'add', type: 'textBox', range: 'F2:H3', name: 'Note', text: 'Read me', altText: 'a note' });
		book.editShape('Sheet1', { action: 'add', type: 'button', range: 'B6:C7', text: 'Go', macro: 'DoIt' });

		expect(shapesOn(reopen(book), 'Sheet1')).toEqual([
			{ name: 'Rectangle: Rounded Corners 1', kind: 'shape', geometry: 'roundRect', range: 'B2:D4', macro: 'DoIt', text: 'Run' },
			{ name: 'Note', kind: 'textBox', range: 'F2:H3', text: 'Read me', altText: 'a note' },
			{ name: 'Button 1', kind: 'button', range: 'B6:C7', macro: 'DoIt', text: 'Go' },
		]);
		// The worksheet's children stay in the order the schema fixes.
		const sheet = part(book, 'xl/worksheets/sheet1.xml');
		const order = ['<pageMargins', '<drawing ', '<legacyDrawing ', '<mc:AlternateContent'].map((tag) => sheet.indexOf(tag));
		expect(order.every((at, i) => at >= 0 && (i === 0 || at > order[i - 1]))).toBe(true);
		expect(sheet).toMatch(/<worksheet [^>]*xmlns:x14="http:\/\/schemas.microsoft.com\/office\/spreadsheetml\/2009\/9\/main"/);
		const types = part(book, '[Content_Types].xml');
		expect(types).toContain('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
		expect(types).toContain('<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>');
		expect(types).toMatch(/<Default Extension="vml" ContentType="application\/vnd.openxmlformats-officedocument.vmlDrawing"\/>/);
		expect(part(book, 'xl/drawings/vmlDrawing1.vml')).toContain('<o:idmap v:ext="edit" data="1"/>');
	});

	it('numbers a new control after the sheet\'s others, and a new VML part after the workbook\'s', () => {
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'add', type: 'button', range: 'K2:L3' });
		book.editShape('Sheet2', { action: 'add', type: 'button', range: 'D2:E3', name: 'Other button' });

		expect(shape(book, 'Sheet1', 'Button 4')).toEqual({ name: 'Button 4', kind: 'button', range: 'K2:L3', text: 'Button 4' });
		expect(part(book, 'xl/drawings/vmlDrawing1.vml')).toContain('id="_x0000_s1028"');
		// Sheet2 had no VML part: its own takes the next id block.
		expect(part(book, 'xl/drawings/vmlDrawing2.vml')).toContain('<o:idmap v:ext="edit" data="2"/>');
		expect(part(book, 'xl/drawings/vmlDrawing2.vml')).toContain('id="_x0000_s2049"');
		expect(shape(reopen(book), 'Sheet2', 'Other button')).toMatchObject({ kind: 'button', range: 'D2:E3' });
	});

	it('gives a new AutoShape an id below the controls\' and a name Excel would', () => {
		// Excel numbers a shape's name one below its id; the group took id 9.
		const book = open('ShapesFixture.xlsm');
		book.editShape('Sheet1', { action: 'add', type: 'oval', range: 'K5' });

		expect(shape(book, 'Sheet1', 'Oval 9')).toEqual({ name: 'Oval 9', kind: 'shape', geometry: 'ellipse', range: 'K5' });
		expect(part(book, 'xl/drawings/drawing1.xml')).toContain('<xdr:cNvPr id="10" name="Oval 9"/>');
	});

	it('keeps text with dollar signs and markup characters as typed', () => {
		const book = open('ShapesFixture.xlsm');
		const text = 'Pay $& now <b> & $$ $1 $\' $`';
		book.editShape('Sheet1', { action: 'add', type: 'button', range: 'K2:L3', name: 'N$&', text, altText: text });
		book.editShape('Sheet1', { action: 'add', type: 'textBox', range: 'K5:L6', name: 'T$$', text, altText: text });
		book.editShape('Sheet1', { action: 'update', name: 'RunButton', text, altText: text, newName: 'R$`' });
		book.editShape('Sheet1', { action: 'update', name: 'Button 1', text, altText: text });

		const shapes = shapesOn(reopen(book), 'Sheet1');
		for (const name of ['N$&', 'T$$', 'R$`', 'Button 1']) {
			expect(shapes.find((s) => s.name === name), name).toMatchObject({ text });
		}
		expect(shapes.find((s) => s.name === 'T$$')?.altText).toBe(text);
		expect(shapes.find((s) => s.name === 'R$`')?.altText).toBe(text);
	});
});

describe('ActiveX controls', () => {
	// No ActiveX control could be made here (the Trust Center setting that
	// disables them stays as it is), so this sheet follows the layout Excel is
	// documented to write: the entry repeated in an mc:Fallback, and the VML
	// shape named for the control with its number in o:spid.
	function withActiveX(): XlsxWorkbook {
		const zip = ZipArchive.read(open('ShapesFixture.xlsm').toBytes());
		const text = (name: string): string => zip.read(name).toString('utf8');
		const put = (name: string, value: string): void => zip.write(name, Buffer.from(value, 'utf8'));
		const mc = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
		put('xl/worksheets/sheet2.xml', text('xl/worksheets/sheet2.xml').replace('<drawing r:id="rId1"/>',
			'<drawing r:id="rId1"/><legacyDrawing r:id="rId2"/>'
			+ `<mc:AlternateContent ${mc}><mc:Choice Requires="x14"><controls><mc:AlternateContent ${mc}><mc:Choice Requires="x14">`
			+ '<control shapeId="2049" r:id="rId3" name="CommandButton1"><controlPr defaultSize="0" autoLine="0" r:id="rId4">'
			+ '<anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
			+ '<to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control>'
			+ '</mc:Choice><mc:Fallback><control shapeId="2049" r:id="rId3" name="CommandButton1"/></mc:Fallback></mc:AlternateContent>'
			+ '</controls></mc:Choice></mc:AlternateContent>'));
		put('xl/worksheets/_rels/sheet2.xml.rels', text('xl/worksheets/_rels/sheet2.xml.rels').replace('</Relationships>',
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing2.vml"/>'
			+ '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml"/>'
			+ '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/>'
			+ '</Relationships>'));
		put('xl/drawings/vmlDrawing2.vml', '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
			+ 'xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>'
			+ '<v:shape id="CommandButton1" o:spid="_x0000_s2049" type="#_x0000_t75" style=\'position:absolute\'>'
			+ '<v:imagedata o:relid="rId1" o:title=""/><x:ClientData ObjectType="Pict"><x:SizeWithCells/>'
			+ '<x:Anchor>1, 0, 7, 0, 3, 0, 9, 0</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape></xml>');
		return XlsxWorkbook.fromBuffer(zip.toBytes());
	}

	it('lists one, once, and neither changes nor removes it', () => {
		const book = withActiveX();

		expect(shapesOn(book, 'Sheet2')).toEqual([
			{ name: 'Rectangle 1', kind: 'shape', geometry: 'rect', range: 'A2:B5', macro: 'Other' },
			{ name: 'CommandButton1', kind: 'activeX', range: 'B8:C9' },
		]);
		expect(() => book.editShape('Sheet2', { action: 'update', name: 'CommandButton1', macro: 'DoIt' }))
			.toThrow(/is an ActiveX control, which XLIDE does not edit; its code is event procedures in the sheet's module, such as CommandButton1_Click/);
		expect(() => book.editShape('Sheet2', { action: 'delete', name: 'CommandButton1' })).toThrow(/does not remove; delete it in Excel/);
	});
});

describe('linking shapes to macros through the project', () => {
	let dir: string;
	let book: string;

	beforeEach(() => {
		dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-shapes-'));
		book = path.join(dir, 'Shapes Book.xlsm');
		fs.copyFileSync(path.join(FIXTURES, 'ShapesFixture.xlsm'), book);
	});

	afterEach(() => {
		fs.rmSync(dir, { recursive: true, force: true });
	});

	const macroOf = (name: string): string | undefined =>
		listShapes(book, 'Sheet1').surfaces[0].shapes.find((s) => s.name === name)?.macro;

	it('links a Public Sub, spelled as the project spells it', () => {
		editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro: 'other' });
		expect(macroOf('RunButton')).toBe('Other');
		editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro: 'MACROS.doit' });
		expect(macroOf('RunButton')).toBe('Macros.DoIt');
		editShape(book, 'Sheet1', { action: 'update', name: 'Button 1', macro: '[0]!Other' });
		expect(macroOf('Button 1')).toBe('Other');
		editShape(book, 'Sheet1', { action: 'update', name: 'Button 1', macro: "'shapes book.xlsm'!DoIt" });
		expect(macroOf('Button 1')).toBe('DoIt');
	});

	it('links a Sub in a sheet\'s module when it is named with the module', () => {
		writeModule(book, 'Sheet1', 'Public Sub OnSheet()\r\nEnd Sub\r\n', 'standard');
		editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro: 'Sheet1.OnSheet' });
		expect(macroOf('RunButton')).toBe('Sheet1.OnSheet');
		expect(() => editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro: 'OnSheet' }))
			.toThrow(/no Public Sub named OnSheet in a standard module/);
	});

	it('refuses what a click could not run, and leaves the file as it was', () => {
		writeModule(book, 'Helpers', [
			'Private Sub Hidden()', 'End Sub',
			'Public Function Answer()', 'End Function',
			'Public Sub NeedsArg(ByVal x As Long)', 'End Sub',
			'Public Sub OptionalArg(Optional ByVal x As Long)', 'End Sub',
			'Public Sub DoIt()', 'End Sub',
		].join('\r\n') + '\r\n', 'standard');
		writeModule(book, 'Thing', 'Public Sub Poke()\r\nEnd Sub\r\n', 'class');
		const before = fs.readFileSync(book);
		const refusals: Array<[string, RegExp]> = [
			['Hidden', /Helpers.Hidden is Private; Excel's Assign Macro offers only Public Subs/],
			['Answer', /Helpers.Answer is a Function; a shape runs a Sub/],
			['NeedsArg', /Helpers.NeedsArg takes parameters; a click passes none/],
			['DoIt', /DoIt is a Public Sub in Macros and Helpers; say which, as Module.DoIt/],
			['Missing', /no Public Sub named Missing in a standard module/],
			['Nowhere.DoIt', /The project has no module named Nowhere/],
			['Thing.Poke', /Thing is a class module; a shape runs a Sub in a standard module/],
			["'Other Book.xlsm'!DoIt", /runs a macro in Other Book.xlsm; XLIDE links shapes only to Subs in this workbook/],
			['Macros.DoIt(1)', /is not a macro Excel can run from a shape/],
		];
		for (const [macro, message] of refusals) {
			expect(() => editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro }), macro).toThrow(message);
		}
		expect(fs.readFileSync(book).equals(before)).toBe(true);

		editShape(book, 'Sheet1', { action: 'update', name: 'RunButton', macro: 'OptionalArg' });
		expect(macroOf('RunButton')).toBe('OptionalArg');
	});

	it('adds a button that runs a Sub, saved to the file, and names what it made', () => {
		expect(editShape(book, 'Sheet1', { action: 'add', type: 'button', range: 'K2:L3', text: 'Start', macro: 'macros.other' }))
			.toEqual({ ok: true, name: 'Button 4' });
		expect(listShapes(book).surfaces[0].shapes.find((s) => s.name === 'Button 4')).toEqual({
			name: 'Button 4', kind: 'button', range: 'K2:L3', macro: 'Macros.Other', text: 'Start',
		});
		expect(editShape(book, 'Sheet1', { action: 'update', name: 'button 4', newName: 'Start' })).toEqual({ ok: true, name: 'Start' });
		expect(editShape(book, 'Sheet1', { action: 'delete', name: 'START' })).toEqual({ ok: true, name: 'Start' });
	});
});
