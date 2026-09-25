// The sheets of a workbook - name, code name, kind and visibility - from
// whichever form the file keeps them in: the OOXML parts of an .xlsm, the
// binary parts of an .xlsb, or the BIFF8 records of an .xls. The tree lists
// a workbook's sheets from this, in tab order, and matches each sheet's code
// name to the module the VBA project has for it. A sheet has a code name
// only once Excel has given it a module, which it does when the VBA editor
// is opened after the sheet was added; a sheet added through automation and
// saved has none.
//
// The layouts were pinned against SheetsFixture.xlsm, .xlsb and .xls in
// tests/fixtures/binaries, which Excel 16.0 saved from one workbook of five
// sheets, and against what Excel itself reported for each of them.

import type { Cfb } from './cfb';
import type { ZipArchive } from './zip';

export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface WorkbookSheet {
	name: string;
	/** The name of the sheet's module in the VBA project; absent until Excel has given the sheet one. */
	codeName?: string;
	kind: SheetKind;
	/** hidden: the user can unhide it; veryHidden: only code can. */
	state?: 'hidden' | 'veryHidden';
}

export class WorkbookSheetsError extends Error {}

/**
 * How much of a sheet part to read for its properties. `sheetPr` is the
 * first child of the root element, so it sits within the first few hundred
 * bytes after the namespace declarations; the compressed head inflates to
 * several times this.
 */
const PART_HEAD_BYTES = 16 * 1024;

// ------------------------------------------------------------------- OOXML

/** The sheets of an OOXML workbook (`xl/workbook.xml`), in tab order. */
export function sheetsOfOoxml(zip: ZipArchive): WorkbookSheet[] {
	const workbookXml = zip.read('xl/workbook.xml').toString('utf8');
	const rels = relationships(zip, 'xl/_rels/workbook.xml.rels');
	const out: WorkbookSheet[] = [];
	const sheetsElement = /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/.exec(workbookXml)?.[1] ?? '';
	for (const tag of sheetsElement.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
		const attrs = attributes(tag[1]);
		const target = rels.get(attrs['r:id'] ?? attrs['id'] ?? '');
		const kind = target ? sheetKindOfPath(target) : undefined;
		if (!target || !kind) {
			continue;
		}
		const sheet: WorkbookSheet = { name: decodeXml(attrs['name'] ?? ''), kind };
		const codeName = ooxmlCodeName(zip, partPath(target));
		if (codeName) { sheet.codeName = codeName; }
		const state = sheetState(attrs['state']);
		if (state) { sheet.state = state; }
		out.push(sheet);
	}
	return out;
}

/**
 * The `codeName` of a sheet part's `sheetPr`, read from the head of the
 * part. When the head shows an element that comes after `sheetPr` in the
 * schema, or the whole part, the sheet has none; a head that shows neither
 * (a root tag longer than the head, say) is read whole.
 */
function ooxmlCodeName(zip: ZipArchive, path: string): string | undefined {
	if (!zip.has(path)) {
		return undefined;
	}
	const head = zip.readPrefix(path, PART_HEAD_BYTES).toString('utf8');
	const found = codeNameIn(head);
	if (found.decided) {
		return found.codeName;
	}
	return codeNameIn(zip.read(path).toString('utf8')).codeName;
}

const AFTER_SHEET_PR_RE = /<(?:dimension|sheetViews|sheetFormatPr|cols|sheetData|sheetProtection|pageMargins|drawing)\b|<\/(?:worksheet|chartsheet|dialogsheet|macroSheet)>/;

function codeNameIn(xml: string): { decided: boolean; codeName?: string } {
	const sheetPr = /<sheetPr\b([^>]*?)\/?>/.exec(xml);
	if (sheetPr) {
		const codeName = decodeXml(attributes(sheetPr[1])['codeName'] ?? '');
		return { decided: true, ...(codeName ? { codeName } : {}) };
	}
	return { decided: AFTER_SHEET_PR_RE.test(xml) };
}

// -------------------------------------------------------------------- xlsb

/** Record types of MS-XLSB that the sheet list needs. */
const BRT_BUNDLE_SH = 0x009c;
const BRT_WS_PROP = 0x0093;
const BRT_CS_PROP = 0x028b;
const BRT_BEGIN_SHEET_DATA = 0x0091;
const BRT_END_SHEET = 0x0082;

/** Where the sheet's name sits in its properties record: after the flags and the tab color. */
const WS_PROP_NAME_OFFSET = 19;
const CS_PROP_NAME_OFFSET = 10;

/** The sheets of a binary workbook (`xl/workbook.bin`), in tab order. */
export function sheetsOfXlsb(zip: ZipArchive): WorkbookSheet[] {
	const workbook = zip.read('xl/workbook.bin');
	const rels = relationships(zip, 'xl/_rels/workbook.bin.rels');
	const out: WorkbookSheet[] = [];
	for (const record of xlsbRecords(workbook)) {
		if (record.type !== BRT_BUNDLE_SH) {
			continue;
		}
		const body = record.body;
		const hsState = body.readUInt32LE(0);
		const relId = xlsbWideString(body, 8);
		const name = xlsbWideString(body, relId.next);
		const target = relId.text === undefined ? undefined : rels.get(relId.text);
		const kind = target ? sheetKindOfPath(target) : undefined;
		if (!target || !kind) {
			continue;
		}
		const sheet: WorkbookSheet = { name: name.text ?? '', kind };
		const codeName = xlsbCodeName(zip, partPath(target));
		if (codeName) { sheet.codeName = codeName; }
		const state = hsState === 1 ? 'hidden' : hsState === 2 ? 'veryHidden' : undefined;
		if (state) { sheet.state = state; }
		out.push(sheet);
	}
	return out;
}

/**
 * The code name in a sheet part's properties record, which comes right
 * after the part's begin record. The head of the part is enough unless it
 * is cut inside the record.
 */
function xlsbCodeName(zip: ZipArchive, path: string): string | undefined {
	if (!zip.has(path)) {
		return undefined;
	}
	const head = zip.readPrefix(path, PART_HEAD_BYTES);
	const found = xlsbCodeNameIn(head);
	return found.decided ? found.codeName : xlsbCodeNameIn(zip.read(path)).codeName;
}

function xlsbCodeNameIn(data: Buffer): { decided: boolean; codeName?: string } {
	for (const record of xlsbRecords(data)) {
		if (record.type === BRT_WS_PROP || record.type === BRT_CS_PROP) {
			const offset = record.type === BRT_WS_PROP ? WS_PROP_NAME_OFFSET : CS_PROP_NAME_OFFSET;
			const name = xlsbTrailingString(record.body, offset);
			return { decided: true, ...(name ? { codeName: name } : {}) };
		}
		if (record.type === BRT_BEGIN_SHEET_DATA || record.type === BRT_END_SHEET) {
			// Past where the properties would be: the sheet has none.
			return { decided: true };
		}
		if (record.truncated) {
			return { decided: false };
		}
	}
	// Neither record seen, and the data ended: a head cut between records,
	// or a chart sheet whose properties record is optional. Only the whole
	// part can say, and reading it whole answers the same way with the end
	// of the part as the end of the search.
	return { decided: false };
}

/**
 * The string a properties record ends with: at its known offset when the
 * record's length agrees, else wherever a length field accounts for exactly
 * the rest of the record.
 */
function xlsbTrailingString(body: Buffer, offset: number): string | undefined {
	const at = (pos: number): string | undefined => {
		if (pos + 4 > body.length) {
			return undefined;
		}
		const cch = body.readUInt32LE(pos);
		if (cch === 0xffffffff) {
			return pos + 4 === body.length ? '' : undefined;
		}
		return pos + 4 + cch * 2 === body.length ? body.subarray(pos + 4).toString('utf16le') : undefined;
	};
	const known = at(offset);
	if (known !== undefined) {
		return known || undefined;
	}
	for (let pos = 0; pos + 4 <= body.length; pos += 1) {
		const candidate = at(pos);
		if (candidate !== undefined) {
			return candidate || undefined;
		}
	}
	return undefined;
}

interface XlsbRecord {
	type: number;
	body: Buffer;
	/** The data ended inside this record: the last of a head read. */
	truncated: boolean;
}

/** The records of a binary part: a 1- or 2-byte type, a 1- to 4-byte size, the body. */
function* xlsbRecords(data: Buffer): Generator<XlsbRecord> {
	let pos = 0;
	while (pos < data.length) {
		const type = xlsbVarint(data, pos, 2);
		const size = xlsbVarint(data, type.next, 4);
		if (type.next >= data.length || size.next > data.length) {
			return;
		}
		const end = size.next + size.value;
		if (end > data.length) {
			yield { type: type.value, body: data.subarray(size.next), truncated: true };
			return;
		}
		yield { type: type.value, body: data.subarray(size.next, end), truncated: false };
		pos = end;
	}
}

function xlsbVarint(data: Buffer, pos: number, maxBytes: number): { value: number; next: number } {
	let value = 0;
	for (let i = 0; i < maxBytes && pos + i < data.length; i += 1) {
		const byte = data[pos + i];
		value |= (byte & 0x7f) << (7 * i);
		if ((byte & 0x80) === 0) {
			return { value, next: pos + i + 1 };
		}
	}
	return { value, next: pos + maxBytes };
}

/** An XLWideString: a 4-byte count of UTF-16 units, 0xFFFFFFFF for none, then the units. */
function xlsbWideString(body: Buffer, pos: number): { text?: string; next: number } {
	const cch = body.readUInt32LE(pos);
	if (cch === 0xffffffff) {
		return { next: pos + 4 };
	}
	return { text: body.subarray(pos + 4, pos + 4 + cch * 2).toString('utf16le'), next: pos + 4 + cch * 2 };
}

// ------------------------------------------------------------------- BIFF8

const BIFF_BOF = 0x0809;
const BIFF_EOF = 0x000a;
const BIFF_BOUNDSHEET = 0x0085;
const BIFF_CODENAME = 0x01ba;

/**
 * The sheets of a legacy workbook, from the BOUNDSHEET records of the
 * Workbook stream's globals and the CODENAME record each sheet's own
 * substream carries once Excel has given it a module.
 */
export function sheetsOfBiff(cfb: Cfb): WorkbookSheet[] {
	if (!cfb.hasStream('Workbook')) {
		throw new WorkbookSheetsError(
			cfb.hasStream('Book')
				? 'This is an Excel 5.0/95 workbook, whose sheet list XLIDE does not read.'
				: 'The compound file has no Workbook stream.',
		);
	}
	const data = cfb.getStream('Workbook');
	const bounds: Array<{ sheet: WorkbookSheet; offset: number }> = [];
	// The globals substream: from the first BOF to its EOF.
	for (const record of biffRecords(data, 0)) {
		if (record.type === BIFF_EOF) {
			break;
		}
		if (record.type !== BIFF_BOUNDSHEET || record.body.length < 8) {
			continue;
		}
		const body = record.body;
		const kind = biffSheetKind(body[5]);
		if (!kind) {
			continue;
		}
		const sheet: WorkbookSheet = { name: biffShortString(body, 6), kind };
		const hsState = body[4] & 0x03;
		const state = hsState === 1 ? 'hidden' : hsState === 2 ? 'veryHidden' : undefined;
		if (state) { sheet.state = state; }
		bounds.push({ sheet, offset: body.readUInt32LE(0) });
	}
	for (const { sheet, offset } of bounds) {
		const codeName = biffCodeName(data, offset);
		if (codeName) { sheet.codeName = codeName; }
	}
	return bounds.map((bound) => bound.sheet);
}

/**
 * The CODENAME record of the substream that starts with the BOF at `offset`.
 * A sheet's substream can hold a chart's, BOF to EOF, so only the EOF that
 * matches the sheet's own BOF ends the search.
 */
function biffCodeName(data: Buffer, offset: number): string | undefined {
	if (offset + 4 > data.length || data.readUInt16LE(offset) !== BIFF_BOF) {
		return undefined;
	}
	let depth = 0;
	for (const record of biffRecords(data, offset)) {
		if (record.type === BIFF_BOF) {
			depth += 1;
		} else if (record.type === BIFF_EOF) {
			depth -= 1;
			if (depth === 0) {
				return undefined;
			}
		} else if (depth === 1 && record.type === BIFF_CODENAME && record.body.length >= 3) {
			return biffString(record.body, 0) || undefined;
		}
	}
	return undefined;
}

function* biffRecords(data: Buffer, from: number): Generator<{ type: number; body: Buffer }> {
	let pos = from;
	while (pos + 4 <= data.length) {
		const type = data.readUInt16LE(pos);
		const size = data.readUInt16LE(pos + 2);
		const end = Math.min(pos + 4 + size, data.length);
		yield { type, body: data.subarray(pos + 4, end) };
		pos = end;
	}
}

/** BOUNDSHEET's dt: 0 a worksheet (or dialog sheet), 1 an Excel 4.0 macro sheet, 2 a chart. */
function biffSheetKind(dt: number): SheetKind | undefined {
	switch (dt) {
		case 0: return 'worksheet';
		case 1: return 'macrosheet';
		case 2: return 'chartsheet';
		default: return undefined;
	}
}

/** ShortXLUnicodeString: a 1-byte count, a flags byte, then the characters. */
function biffShortString(body: Buffer, pos: number): string {
	return biffChars(body, pos + 2, body[pos], body[pos + 1]);
}

/** XLUnicodeString: a 2-byte count, a flags byte, then the characters. */
function biffString(body: Buffer, pos: number): string {
	return biffChars(body, pos + 3, body.readUInt16LE(pos), body[pos + 2]);
}

/** The characters are UTF-16 when the flags say so, else one byte each. */
function biffChars(body: Buffer, pos: number, count: number, flags: number): string {
	return flags & 0x01
		? body.subarray(pos, pos + count * 2).toString('utf16le')
		: body.subarray(pos, pos + count).toString('latin1');
}

// ------------------------------------------------------------------ shared

function sheetKindOfPath(target: string): SheetKind | undefined {
	if (/(^|\/)worksheets\//.test(target)) { return 'worksheet'; }
	if (/(^|\/)chartsheets\//.test(target)) { return 'chartsheet'; }
	if (/(^|\/)dialogsheets\//.test(target)) { return 'dialogsheet'; }
	if (/(^|\/)macrosheets\//.test(target)) { return 'macrosheet'; }
	return undefined;
}

/** A relationship target as a part path: relative to `xl/`, or absolute from the package root. */
function partPath(target: string): string {
	return target.startsWith('/') ? target.slice(1) : `xl/${target}`;
}

function sheetState(value: string | undefined): WorkbookSheet['state'] | undefined {
	return value === 'hidden' || value === 'veryHidden' ? value : undefined;
}

/** The Id -> Target map of a relationships part; empty when the part is missing. */
function relationships(zip: ZipArchive, path: string): Map<string, string> {
	const out = new Map<string, string>();
	if (!zip.has(path)) {
		return out;
	}
	const xml = zip.read(path).toString('utf8');
	for (const tag of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
		const attrs = attributes(tag[1]);
		if (attrs['Id'] && attrs['Target']) {
			out.set(attrs['Id'], decodeXml(attrs['Target']));
		}
	}
	return out;
}

function attributes(text: string): Record<string, string> {
	const out: Record<string, string> = {};
	for (const attr of text.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) {
		out[attr[1]] = attr[2];
	}
	return out;
}

function decodeXml(text: string): string {
	return text.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (_whole, entity: string) => {
		switch (entity) {
			case 'amp': return '&';
			case 'lt': return '<';
			case 'gt': return '>';
			case 'quot': return '"';
			case 'apos': return "'";
			default:
				return entity.startsWith('#x')
					? String.fromCodePoint(parseInt(entity.slice(2), 16))
					: String.fromCodePoint(parseInt(entity.slice(1), 10));
		}
	});
}
