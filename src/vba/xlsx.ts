// OOXML workbook access: XLIDE's sheet/cell surface.
//
// Reads sheet names, used ranges, cell values and formulas, defined names, and
// the embedded vbaProject.bin; writes cell values back. Writes splice into the
// original sheet XML - only the touched rows are re-serialized, so styles,
// conditional formatting, charts, pivot caches and every other part survive
// byte-for-byte.

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

export function columnToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) {
		n = n * 26 + (ch.charCodeAt(0) - 64);
	}
	return n;
}

export function indexToColumn(index: number): string {
	let n = index;
	let out = '';
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out;
}

export function parseCellRef(ref: string): { row: number; col: number } {
	const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref.trim());
	if (!m) {
		throw new XlsxError(`Invalid cell reference '${ref}': expected A1 notation such as 'B3'.`);
	}
	return { row: Number(m[2]), col: columnToIndex(m[1]) };
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

interface SheetRef {
	name: string;
	path: string;
}

export class XlsxWorkbook {
	private sharedStrings: string[] | undefined;
	private dateStyles: Set<number> | undefined;

	private constructor(private readonly zip: ZipArchive) {}

	static fromBuffer(data: Buffer): XlsxWorkbook {
		return new XlsxWorkbook(ZipArchive.read(data));
	}

	toBytes(): Buffer {
		return this.zip.toBytes();
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
		const rels = this.readRelationships('xl/_rels/workbook.xml.rels');
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
			const path = target.startsWith('/')
				? target.slice(1)
				: `xl/${target.replace(/^\.\//, '')}`;
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
		const { r1, c1, r2, c2 } = parseRangeRef(range);
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

	writeCells(sheetName: string, startCell: string, data: CellValue[][]): void {
		const sheet = this.requireSheet(sheetName);
		const original = this.zip.read(sheet.path).toString('utf8');
		const start = parseCellRef(startCell);

		// Collect the target values by row.
		const updates = new Map<number, Map<number, CellValue>>();
		let maxRow = start.row;
		let maxCol = start.col;
		data.forEach((row, rOffset) => {
			const rowNum = start.row + rOffset;
			const byCol = updates.get(rowNum) ?? new Map<number, CellValue>();
			row.forEach((value, cOffset) => {
				const col = start.col + cOffset;
				byCol.set(col, value);
				maxCol = Math.max(maxCol, col);
			});
			updates.set(rowNum, byCol);
			maxRow = Math.max(maxRow, rowNum);
		});

		const updated = spliceRows(original, updates);
		this.zip.write(sheet.path, Buffer.from(expandDimension(updated, maxRow, maxCol), 'utf8'));
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

	private cellValue(cell: RawCell, dataOnly: boolean): CellValue {
		if (!dataOnly && cell.formula !== undefined) {
			return `=${cell.formula}`;
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
					return excelSerialToIso(num);
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
	type?: string;
	style?: number;
	value?: string;
	formula?: string;
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

/**
 * Re-target a shared formula from its master cell to `cell`, shifting relative
 * references by the row/column delta and leaving $-anchored parts fixed.
 */
function translateSharedFormula(shared: Map<number, SharedFormula>, cell: RawCell): string | undefined {
	const master = cell.sharedIndex === undefined ? undefined : shared.get(cell.sharedIndex);
	if (!master) { return undefined; }
	const rowDelta = cell.row - master.row;
	const colDelta = cell.col - master.col;
	if (rowDelta === 0 && colDelta === 0) { return master.formula; }

	const formula = master.formula;
	let out = '';
	let i = 0;
	while (i < formula.length) {
		const ch = formula[i];
		if (ch === '"') {
			// Copy string literals verbatim ("" escapes an inner quote).
			const start = i++;
			while (i < formula.length) {
				if (formula[i] === '"') {
					if (formula[i + 1] === '"') { i += 2; continue; }
					i++;
					break;
				}
				i++;
			}
			out += formula.slice(start, i);
			continue;
		}
		const ref = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/.exec(formula.slice(i));
		// A reference must not continue an identifier (e.g. the "G10" in LOG10).
		const prev = out.length > 0 ? out[out.length - 1] : '';
		if (ref && !/[A-Za-z0-9_]/.test(prev)) {
			const [, colAbs, colLetters, rowAbs, rowDigits] = ref;
			const col = colAbs ? columnToIndex(colLetters) : columnToIndex(colLetters) + colDelta;
			const row = rowAbs ? Number(rowDigits) : Number(rowDigits) + rowDelta;
			if (col >= 1 && col <= 16384 && row >= 1 && row <= 1048576) {
				out += `${colAbs}${indexToColumn(col)}${rowAbs}${row}`;
				i += ref[0].length;
				continue;
			}
			out += '#REF!';
			i += ref[0].length;
			continue;
		}
		out += ch;
		i++;
	}
	return out;
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
			type,
			style: styleRaw === undefined ? undefined : Number(styleRaw),
			value,
			formula: formula === undefined ? undefined : decodeXmlText(formula),
			inlineText,
			sharedIndex: sharedRaw === undefined ? undefined : Number(sharedRaw),
		};
	}
}

// ------------------------------------------------------------- write splice

function serializeCell(ref: string, value: CellValue): string {
	if (value === null || value === undefined) {
		return `<c r="${ref}"/>`;
	}
	if (typeof value === 'number') {
		return Number.isFinite(value) ? `<c r="${ref}"><v>${value}</v></c>` : `<c r="${ref}"/>`;
	}
	if (typeof value === 'boolean') {
		return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
	}
	const text = String(value);
	if (text.startsWith('=')) {
		return `<c r="${ref}"><f>${encodeXmlText(text.slice(1))}</f></c>`;
	}
	// Inline strings avoid mutating the shared-string table (and its refcounts).
	return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${encodeXmlText(text)}</t></is></c>`;
}

/** Rewrite only the rows named in `updates`, preserving all other XML. */
function spliceRows(xml: string, updates: Map<number, Map<number, CellValue>>): string {
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

	interface RowBlock { row: number; text: string }
	const blocks: RowBlock[] = [];
	const rowRe = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
	let m: RegExpExecArray | null;
	while ((m = rowRe.exec(body)) !== null) {
		const rowNum = Number(/\br\s*=\s*"(\d+)"/.exec(m[1])?.[1] ?? '0');
		blocks.push({ row: rowNum, text: m[0] });
	}

	const byRow = new Map(blocks.map((b) => [b.row, b]));
	for (const [rowNum, cells] of updates) {
		const existing = byRow.get(rowNum);
		byRow.set(rowNum, { row: rowNum, text: rewriteRow(existing?.text, rowNum, cells) });
	}
	const ordered = [...byRow.values()].sort((a, b) => a.row - b.row);
	return `${working.slice(0, dataStart)}${ordered.map((b) => b.text).join('')}${working.slice(dataEnd)}`;
}

function rewriteRow(existing: string | undefined, rowNum: number, cells: Map<number, CellValue>): string {
	let attrs = ` r="${rowNum}"`;
	const kept: Array<{ col: number; text: string }> = [];
	if (existing) {
		const m = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/.exec(existing);
		if (m) {
			// Drop stale spans; Excel recomputes them and a wrong value hides cells.
			attrs = m[1].replace(/\s+spans\s*=\s*"[^"]*"/, '');
			for (const cell of iterateCells(m[3] ?? '')) {
				if (!cells.has(cell.col)) {
					const cellRe = new RegExp(`<c\\b[^>]*\\br\\s*=\\s*"${cell.ref}"[^>]*(?:/>|>[\\s\\S]*?</c>)`);
					const raw = cellRe.exec(m[3] ?? '')?.[0];
					if (raw) { kept.push({ col: cell.col, text: raw }); }
				}
			}
		}
	}
	for (const [col, value] of cells) {
		kept.push({ col, text: serializeCell(`${indexToColumn(col)}${rowNum}`, value) });
	}
	kept.sort((a, b) => a.col - b.col);
	return `<row${attrs}>${kept.map((k) => k.text).join('')}</row>`;
}

function expandDimension(xml: string, maxRow: number, maxCol: number): string {
	const m = /<dimension\b[^>]*\bref\s*=\s*"([^"]*)"[^>]*\/>/.exec(xml);
	if (!m) { return xml; }
	let r1 = 1, c1 = 1, r2 = maxRow, c2 = maxCol;
	try {
		const cur = parseRangeRef(m[1]);
		r1 = cur.r1;
		c1 = cur.c1;
		r2 = Math.max(cur.r2, maxRow);
		c2 = Math.max(cur.c2, maxCol);
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
