// OOXML workbook access: XLIDE's sheet/cell surface.
//
// Reads sheet names, used ranges, cell values and formulas, defined names, and
// the embedded vbaProject.bin; writes cell values back. Writes splice into the
// original sheet XML - only the touched rows are re-serialized, so styles,
// conditional formatting, charts, pivot caches and every other part survive
// byte-for-byte.

import {
	FormulaError,
	MAX_COLUMN,
	MAX_ROW,
	columnToIndex,
	formulaForDisplay,
	formulaForFile,
	indexToColumn,
	shiftFormula,
	type FormulaContext,
} from './xlsxFormula';
import { editSheetShape, listSheetShapes, type ShapeEdit, type ShapeInfo } from './xlsxShapes';
import { ZipArchive } from './zip';

export class XlsxError extends Error {}

export interface SheetSummary {
	name: string;
	dimensions: string;
}

export interface NamedRange {
	name: string;
	ref: string;
}

export type CellValue = string | number | boolean | null;

// ---------------------------------------------------------------- XML helpers

interface Tag {
	name: string;
	attrs: Record<string, string>;
	selfClosing: boolean;
	start: number;
	end: number;
}

/** Scan the next element tag at or after `from`, honouring quoted attributes. */
function nextTag(xml: string, from: number): Tag | undefined {
	let i = xml.indexOf('<', from);
	while (i >= 0) {
		const c = xml[i + 1];
		if (c === '?' || c === '!' ) {
			i = xml.indexOf('<', i + 1);
			continue;
		}
		break;
	}
	if (i < 0) { return undefined; }
	let j = i + 1;
	let quote: string | undefined;
	while (j < xml.length) {
		const ch = xml[j];
		if (quote) {
			if (ch === quote) { quote = undefined; }
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === '>') {
			break;
		}
		j++;
	}
	const inner = xml.slice(i + 1, j);
	const selfClosing = inner.endsWith('/');
	const body = selfClosing ? inner.slice(0, -1) : inner;
	const nameMatch = /^\/?([\w:.-]+)/.exec(body);
	const name = nameMatch ? nameMatch[1] : '';
	const attrs: Record<string, string> = {};
	const attrRe = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
	let m: RegExpExecArray | null;
	while ((m = attrRe.exec(body)) !== null) {
		attrs[m[1]] = decodeXmlText(m[3] ?? m[4] ?? '');
	}
	return { name: body.startsWith('/') ? `/${name}` : name, attrs, selfClosing, start: i, end: j + 1 };
}

function decodeXmlText(text: string): string {
	// XML line-ending normalization ([XML] 2.11) happens before entity
	// expansion, so a literal CRLF in the file becomes LF while an explicit
	// &#13; character reference survives as CR.
	const normalized = text.includes('\r') ? text.replace(/\r\n?/g, '\n') : text;
	if (!normalized.includes('&')) { return normalized; }
	return normalized
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&#x([0-9a-fA-F]+);/g, (_s, h) => String.fromCodePoint(parseInt(h, 16)))
		.replace(/&#(\d+);/g, (_s, d) => String.fromCodePoint(Number(d)))
		.replace(/&amp;/g, '&');
}

function encodeXmlText(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		// Control characters are illegal in XML 1.0; Excel drops them too.
		.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
}

function encodeXmlAttr(text: string): string {
	return encodeXmlText(text).replace(/"/g, '&quot;');
}

// ------------------------------------------------------------ A1 conversions

export function parseCellRef(ref: string): { row: number; col: number } {
	const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref.trim());
	const row = m ? Number(m[2]) : 0;
	const col = m ? columnToIndex(m[1]) : 0;
	if (row < 1 || row > MAX_ROW || col > MAX_COLUMN) {
		throw new XlsxError(`Invalid cell reference '${ref}': expected A1 notation within A1:XFD1048576, such as 'B3'.`);
	}
	return { row, col };
}

function parseRangeRef(ref: string): { r1: number; c1: number; r2: number; c2: number } {
	const cleaned = ref.trim().replace(/^.*!/, '');
	const [a, b] = cleaned.split(':');
	const start = parseCellRef(a);
	const end = b ? parseCellRef(b) : start;
	return {
		r1: Math.min(start.row, end.row), c1: Math.min(start.col, end.col),
		r2: Math.max(start.row, end.row), c2: Math.max(start.col, end.col),
	};
}

// ----------------------------------------------------------------- workbook

/** The three places Office hosts keep the VBA project inside an OOXML zip. */
const VBA_PROJECT_PARTS = ['xl/vbaProject.bin', 'word/vbaProject.bin', 'ppt/vbaProject.bin'];

const WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels';
const CONTENT_TYPES = '[Content_Types].xml';

/** A read larger than this is cut to the cells a sheet has. */
const MAX_READ_CELLS = 1_000_000;

/** Days from Excel's 1900 date origin to its 1904 one. */
const DATE_1904_OFFSET = 1462;

/** A workbook relationship target, relative to xl/ unless absolute. */
function workbookPartPath(target: string): string {
	return target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
}

function escapeRegExp(text: string): string {
	return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ------------------------------------------------------ dynamic array metadata
//
// A formula that can spill is marked by its cell's `cm` attribute, a 1-based
// index into the cellMetadata of xl/metadata.xml. The entry there points at an
// XLDAPR metadata type and a futureMetadata block with fDynamic="1". The XML
// below is what Excel 16 writes.

const SHEET_METADATA_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const SHEET_METADATA_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const DYNAMIC_ARRAY_NAMESPACE = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const EMPTY_METADATA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
	+ '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></metadata>';
const DYNAMIC_ARRAY_TYPE = '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
	+ 'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
	+ 'coerce="1" cellMeta="1"/>';
const DYNAMIC_ARRAY_BLOCK = '<bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
	+ '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>';

function blocks(xml: string): string[] {
	return [...xml.matchAll(/<bk\b[^>]*>([\s\S]*?)<\/bk>/g)].map((m) => m[1]);
}

function metadataTypeIndex(xml: string): number {
	return [...xml.matchAll(/<metadataType\b[^>]*?\bname="([^"]*)"/g)].map((m) => m[1]).indexOf('XLDAPR') + 1;
}

function dynamicArrayFuture(xml: string): RegExpExecArray | null {
	return /<futureMetadata\b[^>]*\bname="XLDAPR"[^>]*>([\s\S]*?)<\/futureMetadata>/.exec(xml);
}

/** The 1-based cellMetadata indexes that mark a dynamic array formula. */
function dynamicArrayIndexes(xml: string): Set<number> {
	const type = metadataTypeIndex(xml);
	const dynamic = blocks(dynamicArrayFuture(xml)?.[1] ?? '').map((bk) => /\bfDynamic="(?:1|true)"/.test(bk));
	const cells = /<cellMetadata\b[^>]*>([\s\S]*?)<\/cellMetadata>/.exec(xml)?.[1] ?? '';
	const out = new Set<number>();
	blocks(cells).forEach((bk, i) => {
		const t = Number(/\bt="(\d+)"/.exec(bk)?.[1]);
		const v = Number(/\bv="(\d+)"/.exec(bk)?.[1]);
		if (type > 0 && t === type && dynamic[v]) {
			out.add(i + 1);
		}
	});
	return out;
}

function withCount(xml: string, element: string, count: number): string {
	return xml.replace(new RegExp(`<${element}\\b[^>]*>`), (tag) => (/\bcount="\d*"/.test(tag)
		? tag.replace(/\bcount="\d*"/, `count="${count}"`)
		: tag.replace(new RegExp(`^<${element}\\b`), `<${element} count="${count}"`)));
}

/**
 * Insert `text` after the last of the root's children named in `elements`;
 * the children's order is fixed by the schema.
 */
function insertAfter(xml: string, elements: string[], text: string): string {
	const at = Math.max(...elements.map((name) => {
		const close = xml.lastIndexOf(`</${name}>`);
		return close < 0 ? -1 : close + name.length + 3;
	}));
	return at < 0 ? xml : `${xml.slice(0, at)}${text}${xml.slice(at)}`;
}

/**
 * Metadata XML with an entry marking dynamic array formulas, and the
 * cellMetadata index of that entry. Existing entries are reused.
 */
function withDynamicArrayMetadata(original: string): { xml: string; index: number } {
	const existing = [...dynamicArrayIndexes(original)];
	if (existing.length > 0) {
		return { xml: original, index: Math.min(...existing) };
	}
	let xml = original.includes(DYNAMIC_ARRAY_NAMESPACE)
		? original
		: original.replace(/<metadata\b/, `<metadata xmlns:xda="${DYNAMIC_ARRAY_NAMESPACE}"`);

	let type = metadataTypeIndex(xml);
	if (type === 0) {
		if (!/<metadataTypes\b/.test(xml)) {
			xml = xml.replace(/(<metadata\b[^>]*>)/, '$1<metadataTypes count="0"></metadataTypes>');
		}
		type = [...xml.matchAll(/<metadataType\b/g)].length + 1;
		xml = withCount(xml.replace(/<\/metadataTypes>/, `${DYNAMIC_ARRAY_TYPE}</metadataTypes>`), 'metadataTypes', type);
	}

	let future = dynamicArrayFuture(xml);
	if (!future) {
		xml = insertAfter(xml, ['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata'],
			'<futureMetadata name="XLDAPR" count="0"></futureMetadata>');
		future = dynamicArrayFuture(xml)!;
	}
	let value = blocks(future[1]).findIndex((bk) => /\bfDynamic="(?:1|true)"/.test(bk));
	if (value < 0) {
		value = blocks(future[1]).length;
		const replaced = future[0].replace(/<\/futureMetadata>$/, `${DYNAMIC_ARRAY_BLOCK}</futureMetadata>`)
			.replace(/\bcount="\d*"/, `count="${value + 1}"`);
		xml = xml.replace(future[0], replaced);
	}

	const entry = `<bk><rc t="${type}" v="${value}"/></bk>`;
	if (!/<cellMetadata\b/.test(xml)) {
		xml = insertAfter(xml, ['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata'],
			'<cellMetadata count="0"></cellMetadata>');
	}
	const cells = /<cellMetadata\b[^>]*>([\s\S]*?)<\/cellMetadata>/.exec(xml)!;
	const index = blocks(cells[1]).length + 1;
	xml = withCount(xml.replace(/<\/cellMetadata>/, `${entry}</cellMetadata>`), 'cellMetadata', index);
	return { xml, index };
}

interface SheetRef {
	name: string;
	path: string;
}

export class XlsxWorkbook {
	private sharedStrings: string[] | undefined;
	private dateStyles: Set<number> | undefined;
	private date1904: boolean | undefined;

	private constructor(private readonly zip: ZipArchive) {}

	static fromBuffer(data: Buffer): XlsxWorkbook {
		return new XlsxWorkbook(ZipArchive.read(data));
	}

	toBytes(): Buffer {
		return this.zip.toBytes();
	}

	/**
	 * The package's archive, for the readers that work on parts this class
	 * knows nothing about - a slide's drawing tree, a Word story. Edits to it
	 * are edits to this package, and `toBytes` picks them up.
	 */
	zipArchive(): ZipArchive {
		return this.zip;
	}

	/**
	 * The Office host this OOXML package belongs to, decided by which root
	 * document part it carries - content, not the file extension. Undefined
	 * for a zip that is none of the three.
	 */
	packageHost(): 'excel' | 'word' | 'powerpoint' | undefined {
		if (this.zip.has('xl/workbook.xml') || this.zip.has('xl/workbook.bin')) { return 'excel'; }
		if (this.zip.has('word/document.xml')) { return 'word'; }
		if (this.zip.has('ppt/presentation.xml')) { return 'powerpoint'; }
		return undefined;
	}

	/**
	 * Whether the worksheet/cell surface is readable: .xlsb keeps its
	 * workbook part as binary `xl/workbook.bin`, which this XML reader does
	 * not parse - VBA editing is unaffected, but sheet APIs must refuse
	 * honestly rather than fail on the missing XML part.
	 */
	hasSheetSurface(): boolean {
		return this.zip.has('xl/workbook.xml');
	}

	/** Where this package keeps its VBA project, when it has one. */
	private vbaProjectPath(): string | undefined {
		return VBA_PROJECT_PARTS.find((part) => this.zip.has(part))
			?? this.zip.names().find((name) => /(^|\/)vbaProject\.bin$/.test(name));
	}

	/**
	 * Whether the package carries a VBA project at all. A macro-enabled
	 * format does not imply one: a workbook saved as .xlsm before any macro
	 * is written gets no vbaProject.bin part, which is "no code yet" rather
	 * than a package that could not be read.
	 */
	hasVbaProject(): boolean {
		return this.vbaProjectPath() !== undefined;
	}

	readVbaProject(): Buffer {
		const path = this.vbaProjectPath();
		if (!path) {
			throw new XlsxError('Package contains no VBA project (no vbaProject.bin part).');
		}
		return this.zip.read(path);
	}

	writeVbaProject(data: Buffer): void {
		// Written back to wherever this package keeps it; which containers
		// accept writes at all is macroContainer's decision, not this layer's.
		this.zip.write(this.vbaProjectPath() ?? 'xl/vbaProject.bin', data);
	}

	/** Worksheets in workbook order (chartsheets and dialog sheets excluded). */
	sheets(): SheetRef[] {
		const workbookXml = this.zip.read('xl/workbook.xml').toString('utf8');
		const rels = this.readRelationships(WORKBOOK_RELS);
		const out: SheetRef[] = [];
		let pos = 0;
		for (;;) {
			const tag = nextTag(workbookXml, pos);
			if (!tag) { break; }
			pos = tag.end;
			if (tag.name !== 'sheet') { continue; }
			const rid = tag.attrs['r:id'] ?? tag.attrs['id'];
			const target = rid ? rels.get(rid) : undefined;
			if (!target) { continue; }
			const path = workbookPartPath(target);
			if (!path.includes('/worksheets/')) { continue; }
			out.push({ name: tag.attrs['name'] ?? '', path });
		}
		return out;
	}

	sheetSummaries(): SheetSummary[] {
		return this.sheets().map((sheet) => ({
			name: sheet.name,
			dimensions: this.sheetDimensions(sheet.path),
		}));
	}

	definedNames(): NamedRange[] {
		const workbookXml = this.zip.read('xl/workbook.xml').toString('utf8');
		const out: NamedRange[] = [];
		let pos = 0;
		for (;;) {
			const tag = nextTag(workbookXml, pos);
			if (!tag) { break; }
			pos = tag.end;
			if (tag.name !== 'definedName' || tag.selfClosing) { continue; }
			const close = workbookXml.indexOf('</definedName>', tag.end);
			const ref = close < 0 ? '' : decodeXmlText(workbookXml.slice(tag.end, close));
			out.push({ name: tag.attrs['name'] ?? '', ref });
		}
		return out;
	}

	readCells(sheetName: string, range: string, dataOnly: boolean): CellValue[][] {
		const sheet = this.requireSheet(sheetName);
		const xml = this.zip.read(sheet.path).toString('utf8');
		// A sheet with no rows at all yields no rows for any range, rather than a
		// rectangle of blanks - a range selection can only span rows that exist.
		if (!/<row\b/.test(xml)) {
			return [];
		}
		const bounds = parseRangeRef(range);
		const { r1, c1 } = bounds;
		let { r2, c2 } = bounds;
		// An oversized range such as A1:XFD1048576 is cut to the cells the
		// sheet has, rather than building billions of empty ones.
		if ((r2 - r1 + 1) * (c2 - c1 + 1) > MAX_READ_CELLS) {
			let lastRow = 0;
			let lastCol = 0;
			for (const cell of iterateCells(xml)) {
				lastRow = Math.max(lastRow, cell.row);
				lastCol = Math.max(lastCol, cell.col);
			}
			r2 = Math.min(r2, lastRow);
			c2 = Math.min(c2, lastCol);
			if (r2 < r1 || c2 < c1) {
				return [];
			}
			if ((r2 - r1 + 1) * (c2 - c1 + 1) > MAX_READ_CELLS) {
				throw new XlsxError(`${range} holds more than ${MAX_READ_CELLS} cells of data; read it in smaller ranges.`);
			}
		}
		const grid: CellValue[][] = [];
		for (let r = r1; r <= r2; r++) {
			grid.push(new Array<CellValue>(c2 - c1 + 1).fill(null));
		}
		// Shared formulas store their text once on a master cell; every other
		// participant carries only the group id and must be translated.
		const shared = dataOnly ? undefined : collectSharedFormulas(xml);
		for (const cell of iterateCells(xml)) {
			if (cell.row < r1 || cell.row > r2 || cell.col < c1 || cell.col > c2) { continue; }
			const resolved = shared && cell.formula === undefined && cell.sharedIndex !== undefined
				? { ...cell, formula: translateSharedFormula(shared, cell) }
				: cell;
			grid[cell.row - r1][cell.col - c1] = this.cellValue(resolved, dataOnly);
		}
		return grid;
	}

	/**
	 * Write values from `startCell`, a row of `data` per sheet row. A string
	 * starting with = is a formula as typed into Excel; anything Excel would
	 * refuse is refused here, since Excel will not open a file holding it.
	 */
	writeCells(sheetName: string, startCell: string, data: CellValue[][]): void {
		const sheet = this.requireSheet(sheetName);
		const start = parseCellRef(startCell);
		const width = data.reduce((most, row) => Math.max(most, row.length), 0);
		if (start.row + data.length - 1 > MAX_ROW || start.col + width - 1 > MAX_COLUMN) {
			throw new XlsxError(
				`${data.length} row(s) of ${width} value(s) from ${startCell} run past XFD1048576, the last cell of a worksheet.`,
			);
		}
		const writes = new Map<number, Map<number, CellValue>>();
		data.forEach((row, r) => {
			writes.set(start.row + r, new Map(row.map((value, c) => [start.col + c, value])));
		});

		let context: FormulaContext | undefined;
		const result = rewriteSheetCells(this.zip.read(sheet.path).toString('utf8'), writes, {
			dynamicArrays: this.dynamicArrayCellMetadata(),
			formula: (text, ref) => {
				try {
					return formulaForFile(text, context ??= this.formulaContext());
				} catch (e) {
					if (e instanceof FormulaError) {
						throw new XlsxError(`Nothing was written: the formula for ${ref} is not one Excel accepts - ${e.message}`);
					}
					throw e;
				}
			},
			cellMetadata: () => this.ensureDynamicArrayCellMetadata(),
		});
		this.zip.write(sheet.path, Buffer.from(result.xml, 'utf8'));
		if (result.formulasChanged) {
			this.dropCalcChain();
		}
		this.requestFullCalculation();
	}

	/** The shapes on each worksheet, or on the one named. */
	shapes(sheetName?: string): Array<{ sheet: string; codeName?: string; shapes: ShapeInfo[] }> {
		const sheets = sheetName === undefined ? this.sheets() : [this.requireSheet(sheetName)];
		return sheets.map((sheet) => {
			const codeName = this.sheetCodeName(sheet.path);
			return { sheet: sheet.name, ...(codeName ? { codeName } : {}), shapes: listSheetShapes(this.zip, sheet) };
		});
	}

	/**
	 * A worksheet's code name, the name of its module in the VBA project, from
	 * its sheetPr. Excel writes one once the sheet has its module, which it
	 * makes when the VBA editor is opened after the sheet is added; a sheet
	 * saved before then has none, as Sheet2 of ShapesFixture shows.
	 */
	private sheetCodeName(path: string): string | undefined {
		const xml = this.zip.read(path).toString('utf8');
		const sheetPr = /<sheetPr\b[^>]*>/.exec(xml)?.[0];
		return sheetPr ? /\bcodeName="([^"]*)"/.exec(sheetPr)?.[1] || undefined : undefined;
	}

	/** Add, change or remove one shape on a worksheet; gives the shape's name after the edit. */
	editShape(sheetName: string, edit: ShapeEdit): string {
		return editSheetShape(this.zip, this.requireSheet(sheetName), edit, this.sheets().map((sheet) => sheet.name));
	}

	// ------------------------------------------------------------- internals

	private requireSheet(name: string): SheetRef {
		const needle = name.toLowerCase();
		const sheet = this.sheets().find((s) => s.name.toLowerCase() === needle);
		if (!sheet) {
			throw new XlsxError(`Worksheet not found: ${name}`);
		}
		return sheet;
	}

	private readRelationships(path: string): Map<string, string> {
		const out = new Map<string, string>();
		if (!this.zip.has(path)) { return out; }
		const xml = this.zip.read(path).toString('utf8');
		let pos = 0;
		for (;;) {
			const tag = nextTag(xml, pos);
			if (!tag) { break; }
			pos = tag.end;
			if (tag.name === 'Relationship' && tag.attrs['Id'] && tag.attrs['Target']) {
				out.set(tag.attrs['Id'], tag.attrs['Target']);
			}
		}
		return out;
	}

	/** The part the workbook relates to by a relationship type ending in `typeSuffix`. */
	private workbookPart(typeSuffix: string): string | undefined {
		if (!this.zip.has(WORKBOOK_RELS)) { return undefined; }
		const xml = this.zip.read(WORKBOOK_RELS).toString('utf8');
		for (let tag = nextTag(xml, 0); tag; tag = nextTag(xml, tag.end)) {
			if (tag.name === 'Relationship' && tag.attrs['Target'] && (tag.attrs['Type'] ?? '').endsWith(typeSuffix)) {
				return workbookPartPath(tag.attrs['Target']);
			}
		}
		return undefined;
	}

	private editPart(path: string, edit: (xml: string) => string): void {
		if (!this.zip.has(path)) { return; }
		const xml = this.zip.read(path).toString('utf8');
		const edited = edit(xml);
		if (edited !== xml) {
			this.zip.write(path, Buffer.from(edited, 'utf8'));
		}
	}

	/**
	 * Remove the calc chain. It names every formula cell, and one naming a
	 * cell that no longer holds a formula makes Excel refuse the whole file;
	 * Excel builds a new chain when it loads the workbook.
	 */
	private dropCalcChain(): void {
		const path = this.workbookPart('/calcChain') ?? 'xl/calcChain.xml';
		this.zip.delete(path);
		this.editPart(WORKBOOK_RELS, (xml) => xml.replace(/<Relationship\b[^>]*\/calcChain"[^>]*\/>/g, ''));
		this.editPart(CONTENT_TYPES, (xml) => xml.replace(
			new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(path)}"[^>]*/>`, 'g'),
			'',
		));
	}

	/**
	 * Ask Excel to recalculate every formula when it next opens the file. A
	 * file keeps each formula's last result, so after a write the formulas that
	 * depend on a changed cell would show stale values; Excel clears the flag
	 * when it saves.
	 */
	private requestFullCalculation(): void {
		this.editPart('xl/workbook.xml', (xml) => {
			const calcPr = /<calcPr\b[^>]*>/.exec(xml);
			if (calcPr) {
				const tag = calcPr[0];
				return xml.replace(tag, /\bfullCalcOnLoad\s*=/.test(tag)
					? tag.replace(/\bfullCalcOnLoad\s*=\s*"[^"]*"/, 'fullCalcOnLoad="1"')
					: tag.replace(/^<calcPr\b/, '<calcPr fullCalcOnLoad="1"'));
			}
			// calcPr goes after definedNames and before these, in schema order.
			const next = /<(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/
				.exec(xml);
			return next ? `${xml.slice(0, next.index)}<calcPr fullCalcOnLoad="1"/>${xml.slice(next.index)}` : xml;
		});
	}

	/** The `cm` values that mark a cell's formula as a dynamic array. */
	private dynamicArrayCellMetadata(): Set<number> {
		const path = this.workbookPart('/sheetMetadata');
		return path && this.zip.has(path) ? dynamicArrayIndexes(this.zip.read(path).toString('utf8')) : new Set();
	}

	/**
	 * The `cm` value for a new dynamic array formula, adding the workbook's
	 * metadata part, or the entries in it, when they are missing.
	 */
	private ensureDynamicArrayCellMetadata(): number {
		let path = this.workbookPart('/sheetMetadata');
		if (!path) {
			path = 'xl/metadata.xml';
			this.editPart(WORKBOOK_RELS, (xml) => {
				const ids = [...xml.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]));
				const id = `rId${Math.max(0, ...ids) + 1}`;
				return xml.replace(/<\/Relationships>/, `<Relationship Id="${id}" Type="${SHEET_METADATA_TYPE}" Target="metadata.xml"/></Relationships>`);
			});
			this.editPart(CONTENT_TYPES, (xml) => xml.includes('PartName="/xl/metadata.xml"')
				? xml
				: xml.replace(/<\/Types>/, `<Override PartName="/xl/metadata.xml" ContentType="${SHEET_METADATA_CONTENT_TYPE}"/></Types>`));
		}
		const current = this.zip.has(path) ? this.zip.read(path).toString('utf8') : EMPTY_METADATA;
		const { xml, index } = withDynamicArrayMetadata(current);
		if (xml !== current || !this.zip.has(path)) {
			this.zip.write(path, Buffer.from(xml, 'utf8'));
		}
		return index;
	}

	/** What a formula written into this workbook may refer to. */
	private formulaContext(): FormulaContext {
		const workbookXml = this.zip.read('xl/workbook.xml').toString('utf8');
		const sheets = new Set<string>();
		for (let tag = nextTag(workbookXml, 0); tag; tag = nextTag(workbookXml, tag.end)) {
			if (tag.name === 'sheet' && tag.attrs['name']) {
				sheets.add(tag.attrs['name'].toUpperCase());
			}
		}
		const names = new Set(this.definedNames().map((n) => n.name.toUpperCase()));
		const tables = new Map<string, Set<string>>();
		for (const part of this.zip.names().filter((name) => /^xl\/tables\/[^/]+\.xml$/.test(name))) {
			const xml = this.zip.read(part).toString('utf8');
			const table = /<table\b[^>]*>/.exec(xml)?.[0] ?? '';
			const columns = new Set([...xml.matchAll(/<tableColumn\b[^>]*?\bname="([^"]*)"/g)]
				.map((m) => decodeXmlText(m[1]).toUpperCase()));
			for (const attr of [/\bname="([^"]*)"/, /\bdisplayName="([^"]*)"/]) {
				const value = attr.exec(table)?.[1];
				if (value) {
					names.add(decodeXmlText(value).toUpperCase());
					tables.set(decodeXmlText(value).toUpperCase(), columns);
				}
			}
		}
		return { sheets, names, tables };
	}

	private sheetDimensions(path: string): string {
		const xml = this.zip.read(path).toString('utf8');
		let pos = 0;
		for (;;) {
			const tag = nextTag(xml, pos);
			if (!tag) { break; }
			pos = tag.end;
			if (tag.name === 'dimension' && tag.attrs['ref']) {
				// A single-cell used range is reported as "A1:A1", matching how
				// spreadsheet tooling normalizes a degenerate dimension.
				const ref = tag.attrs['ref'];
				return ref.includes(':') ? ref : `${ref}:${ref}`;
			}
			if (tag.name === 'sheetData') { break; }
		}
		// No dimension hint: derive it from the cells present.
		let minRow = Infinity, minCol = Infinity, maxRow = 0, maxCol = 0;
		for (const cell of iterateCells(xml)) {
			minRow = Math.min(minRow, cell.row); maxRow = Math.max(maxRow, cell.row);
			minCol = Math.min(minCol, cell.col); maxCol = Math.max(maxCol, cell.col);
		}
		if (maxRow === 0) { return ''; }
		return `${indexToColumn(minCol)}${minRow}:${indexToColumn(maxCol)}${maxRow}`;
	}

	private strings(): string[] {
		if (this.sharedStrings) { return this.sharedStrings; }
		const out: string[] = [];
		if (this.zip.has('xl/sharedStrings.xml')) {
			const xml = this.zip.read('xl/sharedStrings.xml').toString('utf8');
			// Each <si> may hold several <t> runs; concatenate them.
			const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
			let m: RegExpExecArray | null;
			while ((m = siRe.exec(xml)) !== null) {
				const body = m[1] ?? '';
				let text = '';
				const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g;
				let t: RegExpExecArray | null;
				while ((t = tRe.exec(body)) !== null) {
					text += decodeXmlText(t[1] ?? '');
				}
				out.push(text);
			}
		}
		this.sharedStrings = out;
		return out;
	}

	/** Style indices whose number format renders as a date/time. */
	private dateStyleIndices(): Set<number> {
		if (this.dateStyles) { return this.dateStyles; }
		const out = new Set<number>();
		if (this.zip.has('xl/styles.xml')) {
			const xml = this.zip.read('xl/styles.xml').toString('utf8');
			const customDate = new Map<number, boolean>();
			let pos = 0;
			for (;;) {
				const tag = nextTag(xml, pos);
				if (!tag) { break; }
				pos = tag.end;
				if (tag.name === 'numFmt' && tag.attrs['numFmtId'] && tag.attrs['formatCode']) {
					customDate.set(Number(tag.attrs['numFmtId']), isDateFormatCode(tag.attrs['formatCode']));
				}
			}
			const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
			if (cellXfs) {
				let p = 0;
				let index = 0;
				for (;;) {
					const tag = nextTag(cellXfs[1], p);
					if (!tag) { break; }
					p = tag.end;
					if (tag.name !== 'xf') { continue; }
					const id = Number(tag.attrs['numFmtId'] ?? '0');
					if (isBuiltinDateFormat(id) || customDate.get(id) === true) {
						out.add(index);
					}
					index++;
				}
			}
		}
		this.dateStyles = out;
		return out;
	}

	/** Whether serial dates count from 1904, as workbooks from older Macs do. */
	private usesDate1904(): boolean {
		this.date1904 ??= /<workbookPr\b[^>]*\bdate1904\s*=\s*"(?:1|true)"/
			.test(this.zip.read('xl/workbook.xml').toString('utf8'));
		return this.date1904;
	}

	private cellValue(cell: RawCell, dataOnly: boolean): CellValue {
		if (!dataOnly && cell.formula !== undefined) {
			return `=${formulaForDisplay(cell.formula)}`;
		}
		const type = cell.type ?? 'n';
		if (cell.inlineText !== undefined) { return cell.inlineText; }
		if (cell.value === undefined) { return null; }
		switch (type) {
			case 's': {
				const idx = Number(cell.value);
				return this.strings()[idx] ?? null;
			}
			case 'str': return decodeXmlText(cell.value);
			case 'b': return cell.value === '1';
			case 'e': return decodeXmlText(cell.value);
			default: {
				const num = Number(cell.value);
				if (!Number.isFinite(num)) { return null; }
				if (cell.style !== undefined && this.dateStyleIndices().has(cell.style)) {
					return excelSerialToIso(this.usesDate1904() ? num + DATE_1904_OFFSET : num);
				}
				return num;
			}
		}
	}
}

// ---------------------------------------------------------------- cell scan

interface RawCell {
	ref: string;
	row: number;
	col: number;
	/** The cell element exactly as the sheet has it. */
	xml: string;
	type?: string;
	style?: number;
	value?: string;
	/** Whether the cell has an <f> element, even one naming only a shared group. */
	hasFormula: boolean;
	formula?: string;
	/** The <f> element's kind (shared, array, dataTable) and the range it covers. */
	formulaType?: string;
	formulaRef?: string;
	/** The cell metadata index (`cm`) that marks a dynamic array formula. */
	cellMetadata?: number;
	inlineText?: string;
	/** Group id of a shared formula (`<f t="shared" si="N"/>`). */
	sharedIndex?: number;
}

interface SharedFormula {
	formula: string;
	row: number;
	col: number;
}

/** Master cell of each shared-formula group, keyed by group id. */
function collectSharedFormulas(xml: string): Map<number, SharedFormula> {
	const out = new Map<number, SharedFormula>();
	for (const cell of iterateCells(xml)) {
		if (cell.sharedIndex !== undefined && cell.formula !== undefined && !out.has(cell.sharedIndex)) {
			out.set(cell.sharedIndex, { formula: cell.formula, row: cell.row, col: cell.col });
		}
	}
	return out;
}

/** A shared formula's text for one of the group's cells, from the master's. */
function translateSharedFormula(shared: Map<number, SharedFormula>, cell: RawCell): string | undefined {
	const master = cell.sharedIndex === undefined ? undefined : shared.get(cell.sharedIndex);
	return master && shiftFormula(master.formula, cell.row - master.row, cell.col - master.col);
}

function* iterateCells(xml: string): Generator<RawCell> {
	const cellRe = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
	let m: RegExpExecArray | null;
	while ((m = cellRe.exec(xml)) !== null) {
		const attrText = m[1];
		const body = m[3] ?? '';
		const ref = /\br\s*=\s*"([^"]*)"/.exec(attrText)?.[1];
		if (!ref) { continue; }
		let pos: { row: number; col: number };
		try { pos = parseCellRef(ref); } catch { continue; }
		const type = /\bt\s*=\s*"([^"]*)"/.exec(attrText)?.[1];
		const styleRaw = /\bs\s*=\s*"([^"]*)"/.exec(attrText)?.[1];
		const value = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1];
		const fTag = /<f\b([^>]*?)(\/>|>([\s\S]*?)<\/f>)/.exec(body);
		const formulaText = fTag?.[3];
		const formula = formulaText === undefined || formulaText === '' ? undefined : formulaText;
		const sharedRaw = fTag ? /\bsi\s*=\s*"(\d+)"/.exec(fTag[1])?.[1] : undefined;
		const cellMetadataRaw = /\bcm\s*=\s*"(\d+)"/.exec(attrText)?.[1];
		const inline = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)?.[1];
		let inlineText: string | undefined;
		if (inline !== undefined) {
			inlineText = '';
			const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
			let t: RegExpExecArray | null;
			while ((t = tRe.exec(inline)) !== null) {
				inlineText += decodeXmlText(t[1]);
			}
		}
		yield {
			ref,
			row: pos.row,
			col: pos.col,
			xml: m[0],
			type,
			style: styleRaw === undefined ? undefined : Number(styleRaw),
			value,
			hasFormula: fTag !== null,
			formula: formula === undefined ? undefined : decodeXmlText(formula),
			formulaType: fTag ? /\bt\s*=\s*"([^"]*)"/.exec(fTag[1])?.[1] : undefined,
			formulaRef: fTag ? /\bref\s*=\s*"([^"]*)"/.exec(fTag[1])?.[1] : undefined,
			cellMetadata: cellMetadataRaw === undefined ? undefined : Number(cellMetadataRaw),
			inlineText,
			sharedIndex: sharedRaw === undefined ? undefined : Number(sharedRaw),
		};
	}
}

// ------------------------------------------------------------- write splice
//
// A write re-serializes only the rows it touches, and around the written cells
// keeps the sheet as Excel would: a cell keeps its format, an array formula is
// replaced whole or not at all, a spill is cleared or blocked, and when the
// first cell of a shared formula is overwritten every other cell of the group
// is given the formula as its own.

interface SheetWriteOptions {
	/** The `cm` values that mark a dynamic array formula. */
	dynamicArrays: ReadonlySet<number>;
	/** A typed formula (without the =) as the file stores it; throws for one Excel would refuse. */
	formula(text: string, ref: string): string;
	/** The `cm` value for a new dynamic array formula. */
	cellMetadata(): number;
}

interface RowBlock {
	/** The row element's attributes, spans dropped: Excel recomputes them and a stale value hides cells. */
	attrs: string;
	text: string;
	body: string;
	cells?: Map<number, RawCell>;
}

interface ColumnStyle {
	min: number;
	max: number;
	style: number;
}

function readRows(body: string): Map<number, RowBlock> {
	const rows = new Map<number, RowBlock>();
	const rowRe = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
	let row = 0;
	let m: RegExpExecArray | null;
	while ((m = rowRe.exec(body)) !== null) {
		// A row without r follows the one before it.
		const r = /\br\s*=\s*"(\d+)"/.exec(m[1])?.[1];
		row = r === undefined ? row + 1 : Number(r);
		rows.set(row, { attrs: m[1].replace(/\s+spans\s*=\s*"[^"]*"/, ''), text: m[0], body: m[3] ?? '' });
	}
	return rows;
}

function rowCells(block: RowBlock): Map<number, RawCell> {
	block.cells ??= new Map([...iterateCells(block.body)].map((cell) => [cell.col, cell]));
	return block.cells;
}

/** The format a new cell takes, as Excel gives it one: its row's when the row is formatted, else its column's. */
function newCellStyle(block: RowBlock | undefined, col: number, columns: ColumnStyle[]): number | undefined {
	if (block && /\bcustomFormat\s*=\s*"(?:1|true)"/.test(block.attrs)) {
		const style = /\bs\s*=\s*"(\d+)"/.exec(block.attrs)?.[1];
		if (style !== undefined) { return Number(style); }
	}
	return columns.find((c) => col >= c.min && col <= c.max)?.style;
}

function serializeCell(
	ref: string,
	value: CellValue,
	style: number | undefined,
	formula?: { text: string; cellMetadata: number },
): string | null {
	const s = style ? ` s="${style}"` : '';
	if (formula) {
		// A formula is stored as Excel 365 stores one typed into a cell: a
		// dynamic array formula, which spills when its result is an array.
		return `<c r="${ref}"${s} cm="${formula.cellMetadata}"><f t="array" ref="${ref}">${encodeXmlText(formula.text)}</f></c>`;
	}
	if (value === null || value === undefined || (typeof value === 'number' && !Number.isFinite(value))) {
		return s ? `<c r="${ref}"${s}/>` : null;
	}
	if (typeof value === 'number') {
		return `<c r="${ref}"${s}><v>${value}</v></c>`;
	}
	if (typeof value === 'boolean') {
		return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
	}
	// Inline strings avoid mutating the shared-string table (and its refcounts).
	return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${encodeXmlText(String(value))}</t></is></c>`;
}

/** A cell emptied but for its format, or gone when it has none. */
function clearedCell(cell: RawCell): string | null {
	return cell.style ? `<c r="${cell.ref}" s="${cell.style}"/>` : null;
}

/**
 * A spill anchor with something now in its way: its range shrinks to itself,
 * and Excel's recalculation on open shows #SPILL! until the range is clear.
 * The cached result stays; Excel will not open a #SPILL! written without the
 * rich-value metadata it keeps for one.
 */
function blockedSpill(cell: RawCell): string {
	return cell.xml.replace(/(<f\b[^>]*?)\bref\s*=\s*"[^"]*"/, `$1ref="${cell.ref}"`);
}

/** A shared group's cell with the group's formula as its own. */
function standaloneFormula(cell: RawCell, formula: string): string {
	return cell.xml.replace(/<f\b([^>]*?)(?:\/>|>[\s\S]*?<\/f>)/, (_whole, attrs: string) =>
		`<f${attrs.replace(/\s+(?:t|ref|si)\s*=\s*"[^"]*"/g, '')}>${encodeXmlText(formula)}</f>`);
}

function rewriteSheetCells(
	xml: string,
	writes: Map<number, Map<number, CellValue>>,
	options: SheetWriteOptions,
): { xml: string; formulasChanged: boolean } {
	const sheetDataOpen = /<sheetData\b[^>]*?(\/>|>)/.exec(xml);
	if (!sheetDataOpen) {
		throw new XlsxError('Worksheet XML has no <sheetData> element.');
	}
	// Self-closing <sheetData/>: expand it so rows can be inserted.
	let working = xml;
	if (sheetDataOpen[1] === '/>') {
		const at = sheetDataOpen.index + sheetDataOpen[0].length;
		working = `${xml.slice(0, sheetDataOpen.index)}<sheetData></sheetData>${xml.slice(at)}`;
	}
	const openMatch = /<sheetData\b[^>]*>/.exec(working)!;
	const dataStart = openMatch.index + openMatch[0].length;
	const dataEnd = working.indexOf('</sheetData>', dataStart);
	const body = working.slice(dataStart, dataEnd);
	const rows = readRows(body);

	const isWritten = (row: number, col: number): boolean => writes.get(row)?.has(col) === true;
	const edits = new Map<number, Map<number, string | null>>();
	const edit = (row: number, col: number, cellXml: string | null): void => {
		const byCol = edits.get(row) ?? new Map<number, string | null>();
		byCol.set(col, cellXml);
		edits.set(row, byCol);
	};
	// Whether a formula cell was overwritten or re-stored, which leaves the
	// calc chain naming formulas that are gone.
	let formulasChanged = false;

	// Formula groups reach past the written cells; only a sheet that has one
	// pays for reading every cell to find them.
	const grouped = /\bt\s*=\s*"(?:array|dataTable|shared)"/.test(body)
		? [...rows.values()].flatMap((block) => [...rowCells(block).values()]).filter((cell) => cell.formulaType !== undefined)
		: [];
	for (const cell of grouped) {
		if ((cell.formulaType !== 'array' && cell.formulaType !== 'dataTable') || !cell.formulaRef) { continue; }
		let range: { r1: number; c1: number; r2: number; c2: number };
		try { range = parseRangeRef(cell.formulaRef); } catch { continue; }
		const inside = (row: number, col: number): boolean =>
			row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
		let hits = 0;
		for (const [row, cols] of writes) {
			for (const col of cols.keys()) {
				if (inside(row, col)) { hits++; }
			}
		}
		const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
		if (hits === 0 || hits === area) { continue; }
		const dynamic = cell.formulaType === 'array' && cell.cellMetadata !== undefined
			&& options.dynamicArrays.has(cell.cellMetadata);
		if (!dynamic) {
			const what = cell.formulaType === 'dataTable' ? 'a data table' : 'an array formula';
			throw new XlsxError(
				`${cell.formulaRef} holds ${what}, and Excel does not change part of one. Write all of ${cell.formulaRef}, or leave it.`,
			);
		}
		// A spill: a value in its anchor replaces the formula and the spill
		// goes; a value anywhere else in its range blocks it, as #SPILL!.
		for (const [row, block] of rows) {
			if (row < range.r1 || row > range.r2) { continue; }
			for (const [col, member] of rowCells(block)) {
				if (member !== cell && inside(row, col) && !isWritten(row, col)) { edit(row, col, clearedCell(member)); }
			}
		}
		if (!isWritten(cell.row, cell.col)) { edit(cell.row, cell.col, blockedSpill(cell)); }
		formulasChanged = true;
	}

	// A shared formula is stored on the group's first cell. When a write
	// replaces that cell, each other cell of the group takes its own copy.
	const masters = new Map<number, RawCell>();
	for (const cell of grouped) {
		if (cell.formulaType === 'shared' && cell.formula !== undefined && cell.sharedIndex !== undefined
			&& isWritten(cell.row, cell.col) && !masters.has(cell.sharedIndex)) {
			masters.set(cell.sharedIndex, cell);
		}
	}
	for (const cell of grouped) {
		const master = cell.sharedIndex === undefined ? undefined : masters.get(cell.sharedIndex);
		if (master?.formula !== undefined && cell !== master && !isWritten(cell.row, cell.col)) {
			edit(cell.row, cell.col, standaloneFormula(cell, shiftFormula(master.formula, cell.row - master.row, cell.col - master.col)));
			formulasChanged = true;
		}
	}

	// Every written formula is checked before anything is serialized, so a
	// bad one leaves the workbook as it was.
	const formulas = new Map<string, string>();
	for (const [row, cols] of writes) {
		for (const [col, value] of cols) {
			if (typeof value === 'string' && value.startsWith('=')) {
				const ref = `${indexToColumn(col)}${row}`;
				formulas.set(ref, options.formula(value.slice(1), ref));
			}
		}
	}
	const cellMetadata = formulas.size > 0 ? options.cellMetadata() : 0;
	const columns: ColumnStyle[] = [...working.matchAll(/<col\b[^>]*>/g)].flatMap((m) => {
		const attr = (name: string): string | undefined => new RegExp(`\\b${name}\\s*=\\s*"(\\d+)"`).exec(m[0])?.[1];
		const style = attr('style');
		return style === undefined ? [] : [{ min: Number(attr('min')), max: Number(attr('max')), style: Number(style) }];
	});
	for (const [row, cols] of writes) {
		const block = rows.get(row);
		for (const [col, value] of cols) {
			const ref = `${indexToColumn(col)}${row}`;
			const existing = block ? rowCells(block).get(col) : undefined;
			formulasChanged ||= existing?.hasFormula === true;
			const text = formulas.get(ref);
			const style = existing ? existing.style : newCellStyle(block, col, columns);
			edit(row, col, serializeCell(ref, value, style, text === undefined ? undefined : { text, cellMetadata }));
		}
	}

	const rebuilt = new Map<number, string>();
	for (const [row, byCol] of edits) {
		const block = rows.get(row);
		const cells = block
			? [...rowCells(block).values()].filter((cell) => !byCol.has(cell.col)).map((cell) => ({ col: cell.col, xml: cell.xml }))
			: [];
		for (const [col, cellXml] of byCol) {
			if (cellXml !== null) { cells.push({ col, xml: cellXml }); }
		}
		if (!block && cells.length === 0) { continue; }
		cells.sort((a, b) => a.col - b.col);
		const attrs = block ? block.attrs : ` r="${row}"`;
		rebuilt.set(row, cells.length > 0 ? `<row${attrs}>${cells.map((c) => c.xml).join('')}</row>` : `<row${attrs}/>`);
	}
	const order = [...new Set([...rows.keys(), ...rebuilt.keys()])].sort((a, b) => a - b);
	const sheetData = order.map((row) => rebuilt.get(row) ?? rows.get(row)!.text).join('');
	return {
		xml: expandDimension(`${working.slice(0, dataStart)}${sheetData}${working.slice(dataEnd)}`, writes),
		formulasChanged,
	};
}

/** Widen the sheet's dimension hint to cover the written cells. */
function expandDimension(xml: string, writes: Map<number, Map<number, CellValue>>): string {
	const m = /<dimension\b[^>]*\bref\s*=\s*"([^"]*)"[^>]*\/>/.exec(xml);
	if (!m) { return xml; }
	let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
	for (const [row, cols] of writes) {
		for (const col of cols.keys()) {
			r1 = Math.min(r1, row); r2 = Math.max(r2, row);
			c1 = Math.min(c1, col); c2 = Math.max(c2, col);
		}
	}
	if (r2 === 0) { return xml; }
	try {
		const cur = parseRangeRef(m[1]);
		r1 = Math.min(r1, cur.r1); c1 = Math.min(c1, cur.c1);
		r2 = Math.max(r2, cur.r2); c2 = Math.max(c2, cur.c2);
	} catch { /* malformed dimension: fall back to the written extent */ }
	const ref = `${indexToColumn(c1)}${r1}:${indexToColumn(c2)}${r2}`;
	return xml.replace(m[0], `<dimension ref="${encodeXmlAttr(ref)}"/>`);
}

// ------------------------------------------------------------- date formats

function isBuiltinDateFormat(id: number): boolean {
	return (id >= 14 && id <= 22) || (id >= 45 && id <= 47);
}

function isDateFormatCode(code: string): boolean {
	// Strip literals/colour and currency sections before sniffing date tokens.
	const stripped = code
		.replace(/\[[^\]]*\]/g, '')
		.replace(/"[^"]*"/g, '')
		.replace(/\\./g, '');
	return /[dmyhs]/i.test(stripped) && !/^[^dmyhs]*[#0?,.%E+-]+[^dmyhs]*$/i.test(stripped);
}

/**
 * Excel serial date to an ISO-8601 string, matching what the previous backend
 * emitted (datetime objects were JSON-encoded via isoformat()).
 */
function excelSerialToIso(serial: number): string {
	// Excel's epoch is 1899-12-30 and it treats 1900 as a leap year.
	const ms = Math.round((serial - 25569) * 86400 * 1000);
	const date = new Date(ms);
	if (Number.isNaN(date.getTime())) { return String(serial); }
	const iso = date.toISOString();
	const hasTime = Math.abs(serial - Math.floor(serial)) > 1e-9;
	return hasTime ? iso.replace('Z', '').replace(/\.000$/, '') : iso.slice(0, 10);
}
