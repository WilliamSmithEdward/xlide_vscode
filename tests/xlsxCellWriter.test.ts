import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { XlsxError, XlsxWorkbook } from '../src/vba/xlsx';
import { ZipArchive } from '../src/vba/zip';

// The fixtures were saved by Excel 16: CellWriterFixture has a formatted cell,
// a formatted empty cell before a value, a shared formula and a calc chain;
// CellStylesFixture a formatted column and row and a CSE array; and
// DynamicArrayFixture spilling formulas with their cell metadata. Each case
// here was also checked by having Excel open the file the engine wrote.
const FIXTURES = path.join(__dirname, 'fixtures', 'binaries');
const BLANK = path.join(__dirname, '..', 'assets', 'templates', 'blank.xlsm');
const SHEET = 'xl/worksheets/sheet1.xml';

const open = (name: string): XlsxWorkbook => XlsxWorkbook.fromBuffer(fs.readFileSync(path.join(FIXTURES, name)));
const zipOf = (book: XlsxWorkbook): ZipArchive => ZipArchive.read(book.toBytes());
const part = (book: XlsxWorkbook, name: string): string => zipOf(book).read(name).toString('utf8');
const cell = (book: XlsxWorkbook, ref: string): string | undefined =>
	new RegExp(`<c r="${ref}"[^>]*?(?:/>|>.*?</c>)`, 's').exec(part(book, SHEET))?.[0];

describe('writeCells', () => {
	it('keeps the cells beside a written one exactly', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'C12', [[5]]);
		const row = /<row r="12"[^>]*>.*?<\/row>/s.exec(part(book, SHEET))![0];
		expect(row).toContain('<c r="A12" s="2"/><c r="B12"><v>7</v></c><c r="C12"><v>5</v></c>');
		expect(row.match(/r="B12"/g)).toHaveLength(1);
	});

	it('keeps a cell\'s format when writing over it or clearing it', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'A1', [[99]]);
		expect(cell(book, 'A1')).toBe('<c r="A1" s="1"><v>99</v></c>');
		book.writeCells('Sheet1', 'A1', [[null]]);
		expect(cell(book, 'A1')).toBe('<c r="A1" s="1"/>');
	});

	it('formats a new cell from its row, else its column, as Excel does', () => {
		const book = open('CellStylesFixture.xlsm');
		book.writeCells('Sheet1', 'F5', [[12.5]]);
		book.writeCells('Sheet1', 'F20', [[0.25, 0.5]]);
		expect(cell(book, 'F5')).toBe('<c r="F5" s="1"><v>12.5</v></c>');
		expect(cell(book, 'F20')).toBe('<c r="F20" s="2"><v>0.25</v></c>');
		expect(cell(book, 'G20')).toBe('<c r="G20" s="2"><v>0.5</v></c>');
	});

	it('refuses cells past the edge of a worksheet', () => {
		const book = open('CellWriterFixture.xlsm');
		const before = part(book, SHEET);
		for (const ref of ['A0', 'XFE1', 'A1048577']) {
			expect(() => book.writeCells('Sheet1', ref, [[1]])).toThrow(/Invalid cell reference/);
		}
		expect(() => book.writeCells('Sheet1', 'A1048576', [[1], [2]])).toThrow(/run past XFD1048576/);
		expect(part(book, SHEET)).toBe(before);
	});

	it('drops the calc chain only when a formula cell is overwritten', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'F20', [[1]]);
		expect(zipOf(book).has('xl/calcChain.xml')).toBe(true);
		book.writeCells('Sheet1', 'D1', [[5]]);
		expect(zipOf(book).has('xl/calcChain.xml')).toBe(false);
		expect(part(book, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
		expect(part(book, '[Content_Types].xml')).not.toContain('calcChain');
	});

	it('asks Excel to recalculate every formula when it opens the file', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'A2', [[100]]);
		expect(part(book, 'xl/workbook.xml')).toMatch(/<calcPr fullCalcOnLoad="1" calcId="\d+"\/>/);
	});

	it('gives a shared group its formula cell by cell when its first cell is overwritten', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'C3', [[1]]);
		expect(cell(book, 'C4')).toBe('<c r="C4"><f>A4*B4</f><v>160</v></c>');
		expect(cell(book, 'C10')).toBe('<c r="C10"><f>A10*B10</f><v>1000</v></c>');
	});

	it('leaves a shared group alone when another of its cells is written', () => {
		const book = open('CellWriterFixture.xlsm');
		book.writeCells('Sheet1', 'C5', [[1]]);
		expect(cell(book, 'C4')).toBe('<c r="C4"><f t="shared" si="0"/><v>160</v></c>');
		expect(cell(book, 'C5')).toBe('<c r="C5"><v>1</v></c>');
	});

	it('refuses part of an array formula and replaces all of one', () => {
		const book = open('CellStylesFixture.xlsm');
		const before = part(book, SHEET);
		expect(() => book.writeCells('Sheet1', 'H3', [[1]])).toThrow(/H2:H4 holds an array formula/);
		expect(part(book, SHEET)).toBe(before);
		book.writeCells('Sheet1', 'H2', [[1], [2], [3]]);
		expect(cell(book, 'H2')).toBe('<c r="H2"><v>1</v></c>');
		expect(cell(book, 'H3')).toBe('<c r="H3"><v>2</v></c>');
	});

	it('blocks a spill when a value lands in its range', () => {
		const book = open('DynamicArrayFixture.xlsm');
		book.writeCells('Sheet1', 'F3', [[99]]);
		expect(cell(book, 'F2')).toBe('<c r="F2" cm="1"><f t="array" ref="F2">_xlfn._xlws.SORT(A2:A4)</f><v>1</v></c>');
		expect(cell(book, 'F3')).toBe('<c r="F3"><v>99</v></c>');
		expect(cell(book, 'F4')).toBeUndefined();
	});

	it('clears a spill when its anchor is overwritten', () => {
		const book = open('DynamicArrayFixture.xlsm');
		book.writeCells('Sheet1', 'F2', [[5]]);
		expect(cell(book, 'F2')).toBe('<c r="F2"><v>5</v></c>');
		expect(cell(book, 'F3')).toBeUndefined();
		expect(cell(book, 'F4')).toBeUndefined();
	});

	it('stores a formula as Excel 365 does, adding the metadata a workbook lacks', () => {
		const book = XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK));
		book.writeCells('Sheet1', 'A1', [['=SEQUENCE(3)', '=XLOOKUP(2,A1:A3,A1:A3)']]);
		expect(cell(book, 'A1')).toBe('<c r="A1" cm="1"><f t="array" ref="A1">_xlfn.SEQUENCE(3)</f></c>');
		expect(cell(book, 'B1')).toBe('<c r="B1" cm="1"><f t="array" ref="B1">_xlfn.XLOOKUP(2,A1:A3,A1:A3)</f></c>');
		const metadata = part(book, 'xl/metadata.xml');
		expect(metadata).toContain('<metadataTypes count="1"><metadataType name="XLDAPR"');
		expect(metadata).toContain('</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>');
		expect(part(book, 'xl/_rels/workbook.xml.rels')).toContain('relationships/sheetMetadata" Target="metadata.xml"');
		expect(part(book, '[Content_Types].xml')).toContain('<Override PartName="/xl/metadata.xml"');
		// A second write reuses the entry rather than adding another.
		book.writeCells('Sheet1', 'C1', [['=SUM(A1#)']]);
		expect(cell(book, 'C1')).toBe('<c r="C1" cm="1"><f t="array" ref="C1">SUM(_xlfn.ANCHORARRAY(A1))</f></c>');
		expect(part(book, 'xl/metadata.xml')).toBe(metadata);
	});

	it('reuses the dynamic array metadata Excel wrote', () => {
		const book = open('DynamicArrayFixture.xlsm');
		const metadata = part(book, 'xl/metadata.xml');
		book.writeCells('Sheet1', 'L2', [['=LET(x,2,x*3)']]);
		expect(cell(book, 'L2')).toBe('<c r="L2" cm="1"><f t="array" ref="L2">_xlfn.LET(_xlpm.x,2,_xlpm.x*3)</f></c>');
		expect(part(book, 'xl/metadata.xml')).toBe(metadata);
	});

	it('writes nothing when a formula is one Excel would refuse', () => {
		const book = XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK));
		const before = part(book, SHEET);
		expect(() => book.writeCells('Sheet1', 'A1', [[1, '=SUM(A1;A2)']]))
			.toThrow(/Nothing was written: the formula for B1 is not one Excel accepts - arguments are separated by commas/);
		expect(() => book.writeCells('Sheet1', 'A1', [['=SPLIT("a,b",",")']])).toThrow(XlsxError);
		expect(part(book, SHEET)).toBe(before);
		expect(zipOf(book).has('xl/metadata.xml')).toBe(false);
	});

	it('widens the dimension to the written cells', () => {
		const book = XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK));
		book.writeCells('Sheet1', 'C3', [[1, 2]]);
		expect(part(book, SHEET)).toContain('<dimension ref="A1:D3"/>');
	});
});

describe('readCells', () => {
	it('reads formulas as Excel shows them', () => {
		const book = open('DynamicArrayFixture.xlsm');
		expect(book.readCells('Sheet1', 'F2:K2', false)).toEqual([[
			'=SORT(A2:A4)', '=SEQUENCE(3)', '=A2:A4*2', '=XLOOKUP(2,A2:A4,A2:A4)', '=A2:A4*2', '=LET(x,2,x*3)',
		]]);
	});

	it('reads a formula back as it was written', () => {
		const book = XlsxWorkbook.fromBuffer(fs.readFileSync(BLANK));
		const typed = ['=FILTER(A1:A3,A1:A3>1)', '=SUM(B1#)', '=@A1:A3', '=SUM(A1:.A3)', '=GROUPBY(A1:A3,A1:A3,SUM)'];
		book.writeCells('Sheet1', 'D1', [typed]);
		expect(book.readCells('Sheet1', 'D1:H1', false)).toEqual([typed]);
	});

	it('translates shared formulas without touching function names', () => {
		const book = open('FormulaTableFixture.xlsm');
		expect(book.readCells('Sheet1', 'J6:M6', false)).toEqual([['=LOG10(A6)+A6', '=DEC2BIN(A6)&A6', "='Q1'!A6+A6", '=SUM(6:6)+A6']]);
		expect(book.readCells('Sheet1', 'N8:P8', false)).toEqual([['=SUM(B:B)+B$1', '=SUM(C:C)+C$1', '=SUM(D:D)+D$1']]);
	});

	it('reads dates in a workbook that counts them from 1904', () => {
		const book = open('Date1904Fixture.xlsm');
		expect(book.readCells('Sheet1', 'A2', true)).toEqual([['2024-03-15']]);
	});

	it('cuts an oversized range to the cells the sheet has', () => {
		const book = open('CellWriterFixture.xlsm');
		const grid = book.readCells('Sheet1', 'A1:XFD1048576', true);
		expect(grid).toHaveLength(12);
		expect(grid[0]).toHaveLength(4);
	});
});
