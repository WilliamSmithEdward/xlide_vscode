// Native workbook service: every VBA/cell operation XLIDE needs, implemented
// directly against the OOXML package and the MS-OVBA VBA project. This replaces
// the external backend entirely.
//
// Writes are atomic: the updated package is written to a sibling temp file and
// renamed over the original, so a failure part-way through never leaves a
// truncated workbook.

import * as fs from 'fs';
import * as path from 'path';
import { Cfb } from './cfb';
import { decodeCodePage, encodeCodePage } from './codePages';
import {
	composeFormFrx,
	composeFrmDesignerBlock,
	mergeVbFrameFromFrm,
	parseFormDesignerStreams,
	parseFormFrx,
} from './formDesigner';
import { openMacroContainer, type MacroContainer } from './macroContainer';
import {
	detectSignature,
	synthesizeClassHeader,
	synthesizeStandardHeader,
	VbaProject,
	type VbaModule,
} from './vbaProject';
import { XlsxWorkbook, type CellValue, type NamedRange, type SheetSummary } from './xlsx';

export type ModuleType = 'standard' | 'class' | 'document' | 'userform';
export type DocumentType = 'workbook' | 'worksheet' | 'chart' | 'document';

export interface ModuleEntry {
	name: string;
	type: ModuleType;
	documentType?: DocumentType;
	source?: string;
	/**
	 * A form's designer-declared controls, read natively from the designer
	 * storage inside vbaProject.bin. Present only for userform modules whose
	 * designer parsed cleanly; absent means "not known", never "none".
	 */
	implicitMembers?: { name: string; type: string }[];
}

export interface ProcedureEntry {
	name: string;
	kind: string;
	line: number;
}

export interface ProtectionInfo {
	isPasswordProtected: boolean;
	isSigned: boolean;
}

export interface WriteResult {
	ok: true;
	signatureDropped: boolean;
}

const WORKBOOK_CLSID = '{00020819-0000-0000-C000-000000000046}';
const WORKSHEET_CLSID = '{00020820-0000-0000-C000-000000000046}';
const CHART_CLSID = '{00020821-0000-0000-C000-000000000046}';
const WORD_DOCUMENT_CLSID = '{00020906-0000-0000-C000-000000000046}';
const GUID_RE = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g;
const DOCUMENT_NAME_RE = /^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha)\d*$/i;
const ATTR_LINE_RE = /^Attribute\s+VB_/i;
// \p{L}, not \w: VBA identifiers may use any locale letter (a Russian
// project legitimately declares `Sub Proverka()` in Cyrillic), and the
// ASCII-only \w made such procedures vanish from the explorer tree. \p{M}
// after the first character because Thai and Devanagari build a letter from
// a base plus a combining mark, and stopping at the mark made those
// procedures vanish the same way.
const PROC_RE =
	/^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\p{L}_][\p{L}\p{M}\p{N}_]*)\s*[(\r\n]/gimu;

/**
 * Split module source into (hidden header, visible body) exactly as the editor
 * surface expects: the VERSION/BEGIN/END class preamble plus the contiguous run
 * of `Attribute VB_*` lines are hidden; the body has leading blank lines removed.
 */
export function splitVbaSource(source: string): { header: string; body: string } {
	const lines = source.split(/(?<=\n)/); // keep line endings
	let i = 0;
	if (lines.length > 0 && /^VERSION\s+\d/i.test(lines[0])) {
		i++;
		// A class preamble opens with a bare BEGIN; a form's designer block
		// opens with `Begin {GUID} Name` and can NEST further Begin blocks for
		// its controls. Both are one balanced block, walked by depth - without
		// this, importing a .frm spliced the designer text into the code module.
		const opener = i < lines.length ? lines[i].replace(/[\r\n]+$/, '').trim() : '';
		if (/^BEGIN\b/i.test(opener)) {
			let depth = 1;
			i++;
			while (i < lines.length && depth > 0) {
				const line = lines[i].replace(/[\r\n]+$/, '').trim();
				if (/^Begin\b/i.test(line)) {
					depth++;
				} else if (/^End$/i.test(line)) {
					depth--;
				}
				i++;
			}
		}
	}
	while (i < lines.length && ATTR_LINE_RE.test(lines[i])) {
		i++;
	}
	return {
		header: lines.slice(0, i).join(''),
		body: lines.slice(i).join('').replace(/^[\r\n]+/, ''),
	};
}

export function joinVbaSource(header: string, body: string): string {
	if (!header) { return body; }
	return `${header.replace(/[\r\n]+$/, '')}\r\n${body}`;
}

function attributeValue(source: string, attribute: string): string {
	const re = new RegExp(`^\\s*Attribute\\s+${attribute}\\s*=\\s*"([^"]*)"`, 'im');
	return re.exec(source)?.[1] ?? '';
}

export function classifyModuleType(name: string, source: string): ModuleType {
	const vbBase = attributeValue(source, 'VB_Base');
	if (vbBase) {
		// UserForms carry TWO GUIDs in VB_Base; classes and documents carry one.
		if ((vbBase.match(GUID_RE) ?? []).length >= 2) { return 'userform'; }
		const upper = vbBase.toUpperCase();
		if (upper.includes(WORKBOOK_CLSID) || upper.includes(WORKSHEET_CLSID) || upper.includes(CHART_CLSID)
			|| upper.includes(WORD_DOCUMENT_CLSID)) {
			return 'document';
		}
	}
	// Host code-behind that names no CLSID: Word's ThisDocument declares
	// VB_Base = "1Normal.ThisDocument" (measured against a live-authored
	// .docm). Office marks every document module PredeclaredId + Exposed and
	// nothing else it authors gets both, so the pair is the host-generic
	// document signature (forms are Exposed = False and caught above).
	if (/^True$/i.test(attributeValue(source, 'VB_PredeclaredId'))
		&& /^True$/i.test(attributeValue(source, 'VB_Exposed'))) {
		return 'document';
	}
	if (name === 'ThisWorkbook' || name === 'ThisDocument' || DOCUMENT_NAME_RE.test(name)) { return 'document'; }
	return 'standard';
}

export function classifyDocumentType(name: string, source: string): DocumentType | undefined {
	const vbBase = attributeValue(source, 'VB_Base').toUpperCase();
	if (vbBase.includes(WORKBOOK_CLSID)) { return 'workbook'; }
	if (vbBase.includes(WORKSHEET_CLSID)) { return 'worksheet'; }
	if (vbBase.includes(CHART_CLSID)) { return 'chart'; }
	if (vbBase.includes(WORD_DOCUMENT_CLSID)) { return 'document'; }
	if (name === 'ThisWorkbook') { return 'workbook'; }
	if (name === 'ThisDocument') { return 'document'; }
	if (/^Chart\d*$/i.test(name)) { return 'chart'; }
	if (DOCUMENT_NAME_RE.test(name)) { return 'worksheet'; }
	return undefined;
}

function moduleEntry(module: VbaModule): ModuleEntry {
	let type: ModuleType;
	if (module.kind === 'standard') {
		type = 'standard';
	} else {
		// MODULETYPE 'other' covers class, document and designer modules.
		// Both classifiers read only `Attribute VB_*` lines, which live at the
		// very top of the module, so the header prefix is all they need.
		type = classifyModuleType(module.name, module.sourceHeader);
		if (type === 'standard') { type = 'class'; }
	}
	const entry: ModuleEntry = { name: module.name, type };
	if (type === 'document') {
		const documentType = classifyDocumentType(module.name, module.sourceHeader);
		if (documentType) { entry.documentType = documentType; }
	}
	return entry;
}

interface OpenWorkbook {
	container: MacroContainer;
	cfb: Cfb;
	project: VbaProject;
}

// -------------------------------------------------------------- parse cache
//
// Re-reading and re-parsing the file IS the cost of a read now that the engine
// is in-process, and callers arrive in bursts that hit the same workbook: an
// explorer expansion is listModules + a protection probe + one listSubs per
// module, and every one of those re-opened the file. Reads share one parse per
// workbook, validated against (mtimeMs, size) on every call so out-of-band
// writers (Excel, git, another window) are always seen.
//
// Writes never touch the cache: they parse fresh, because a mutating save that
// fails halfway must not leave a poisoned parse behind for readers - and every
// mutation lands through atomicWrite, which drops the entry. The mtime check
// is the backstop for writers that bypass this process entirely.

interface WorkbookCacheEntry {
	mtimeMs: number;
	size: number;
	container: MacroContainer;
	/** Built on first VBA access; sheet/cell reads never pay for the project. */
	cfb?: Cfb;
	project?: VbaProject;
}

/**
 * Small on purpose: an entry retains the package plus decompressed module
 * sources (a few MB for a large workbook), and a session's hot set is the
 * handful of workbooks whose trees or editors are open.
 */
const WORKBOOK_CACHE_MAX = 4;
const workbookCache = new Map<string, WorkbookCacheEntry>();
let cacheHits = 0;
let cacheMisses = 0;

function cachedPackage(filePath: string): WorkbookCacheEntry {
	// Stat BEFORE reading: if a writer swaps the file between the stat and the
	// read, this entry holds the new bytes under the old mtime, so the next
	// call mismatches and rebuilds - a stale parse can never survive.
	const stat = fs.statSync(filePath);
	const hit = workbookCache.get(filePath);
	if (hit && hit.mtimeMs === stat.mtimeMs && hit.size === stat.size) {
		cacheHits++;
		// Refresh recency: Map iteration order is insertion order.
		workbookCache.delete(filePath);
		workbookCache.set(filePath, hit);
		return hit;
	}
	cacheMisses++;
	const entry: WorkbookCacheEntry = {
		mtimeMs: stat.mtimeMs,
		size: stat.size,
		container: openMacroContainer(fs.readFileSync(filePath)),
	};
	workbookCache.delete(filePath);
	workbookCache.set(filePath, entry);
	while (workbookCache.size > WORKBOOK_CACHE_MAX) {
		const oldest = workbookCache.keys().next().value;
		if (oldest === undefined) { break; }
		workbookCache.delete(oldest);
	}
	return entry;
}

/** Shared read-only parse. Callers must not mutate the returned project. */
function openWorkbook(filePath: string): OpenWorkbook {
	const entry = cachedPackage(filePath);
	entry.cfb ??= entry.container.vbaCfb();
	entry.project ??= VbaProject.parse(entry.cfb);
	return { container: entry.container, cfb: entry.cfb, project: entry.project };
}

/** Fresh parse for mutating operations; never aliases the shared cache. */
function openWorkbookForWrite(filePath: string): OpenWorkbook {
	const container = openMacroContainer(fs.readFileSync(filePath));
	if (!container.writable) {
		throw new Error(`${path.basename(filePath)} is ${container.description}.`);
	}
	const cfb = container.vbaCfb();
	return { container, cfb, project: VbaProject.parse(cfb) };
}

/** The OOXML sheet/cell surface, or an honest refusal for containers without one. */
function sheetSurface(filePath: string): XlsxWorkbook {
	const entry = cachedPackage(filePath);
	const { container } = entry;
	if (container.kind !== 'excel' || !container.xlsx) {
		throw new Error(
			`${path.basename(filePath)} is ${container.description}; it has no worksheet surface.`,
		);
	}
	if (!container.xlsx.hasSheetSurface()) {
		// .xlsb: the workbook part is binary (xl/workbook.bin), which the
		// sheet reader does not parse. VBA editing is unaffected.
		throw new Error(
			`${path.basename(filePath)} is a binary Excel workbook (.xlsb); its worksheet data is `
			+ 'stored in a binary format XLIDE does not read. VBA editing is unaffected - save the '
			+ 'workbook as .xlsm to use the sheet and cell tools.',
		);
	}
	return container.xlsx;
}

/** Test hooks; product code never reads these. */
export function workbookCacheStatsForTests(): { hits: number; misses: number; size: number } {
	return { hits: cacheHits, misses: cacheMisses, size: workbookCache.size };
}

export function resetWorkbookCacheForTests(): void {
	workbookCache.clear();
	cacheHits = 0;
	cacheMisses = 0;
}

/** Write the mutated VBA project back into the container, atomically. */
function saveWorkbook(filePath: string, wb: OpenWorkbook): void {
	wb.project.save(wb.cfb);
	atomicWrite(filePath, wb.container.toFileBytes(wb.cfb));
}

export function atomicWrite(filePath: string, data: Buffer): void {
	const dir = path.dirname(path.resolve(filePath));
	const tmp = path.join(dir, `.xlide-${process.pid}-${Date.now()}.tmp`);
	try {
		fs.writeFileSync(tmp, data);
		try {
			// Preserve the original file mode; a fresh temp file would otherwise
			// narrow permissions on POSIX.
			const stat = fs.statSync(filePath);
			fs.chmodSync(tmp, stat.mode);
		} catch { /* new file: keep the default mode */ }
		fs.renameSync(tmp, filePath);
	} catch (err) {
		try { fs.unlinkSync(tmp); } catch { /* nothing to clean up */ }
		throw err;
	}
	// Every in-process mutation funnels through here, so this is the one place
	// cache invalidation cannot be forgotten. (A caller-supplied path with
	// different casing would miss this delete; the per-call mtime check still
	// catches that, so a stale entry can cost a re-parse but never stale data.)
	workbookCache.delete(filePath);
}

// ------------------------------------------------------------------ read API

export function listModules(filePath: string): ModuleEntry[] {
	const { cfb, project } = openWorkbook(filePath);
	return project.modules.map((module) => moduleEntryWithDesigner(cfb, project, module));
}

export function readModules(filePath: string, full = false): ModuleEntry[] {
	const { cfb, project } = openWorkbook(filePath);
	const out: ModuleEntry[] = [];
	for (const module of project.modules) {
		try {
			const entry = moduleEntryWithDesigner(cfb, project, module);
			entry.source = full ? module.source : splitVbaSource(module.source).body;
			out.push(entry);
		} catch {
			// Keep workbook-wide reads best-effort at the module boundary.
			continue;
		}
	}
	return out;
}

/**
 * A module entry, with a userform's controls read from its designer storage.
 * The storage is named after the module and its `f` stream is the MS-OFORMS
 * FormControl; the module's own text never mentions the controls at all.
 */
function moduleEntryWithDesigner(cfb: Cfb, project: VbaProject, module: VbaModule): ModuleEntry {
	const entry = moduleEntry(module);
	if (entry.type !== 'userform') {
		return entry;
	}
	try {
		const f = cfb.getStreamInStorage(module.name, 'f');
		const o = cfb.hasStreamInStorage(module.name, 'o')
			? cfb.getStreamInStorage(module.name, 'o')
			: undefined;
		const controls = parseFormDesignerStreams(f, o, (bytes, compressed) =>
			compressed ? decodeCodePage(bytes, project.codePage) : bytes.toString('utf16le'));
		if (controls) {
			entry.implicitMembers = controls;
		}
	} catch {
		// No designer storage, or a shape this reader does not understand:
		// the entry simply carries no members, same as before.
	}
	return entry;
}

export function readModule(filePath: string, moduleName: string, full = false): { source: string } {
	const { project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	return { source: full ? module.source : splitVbaSource(module.source).body };
}

/**
 * A form's export pair, composed natively from the workbook: the `.frm` text
 * (designer block from the VBFrame stream, `OleObjectBlob` naming the sidecar,
 * then the module's own attributes and code) and the `.frx` sidecar packaging
 * the designer storage's binary streams.
 */
export function readFormExport(filePath: string, moduleName: string): { frm: string; frx: Buffer } {
	const { cfb, project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	const designer = readDesignerStorage(cfb, module.name);
	if (!designer) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const safeName = module.name.replace(/[<>:"/\|?*\x00-\x1f]/g, '_');
	const frm = composeFrmDesignerBlock(designer.vbFrame, `${safeName}.frx`) + module.source;
	return { frm, frx: composeFormFrx(designer) };
}

/**
 * Writes a form's designer back into the workbook from a `.frx` sidecar (the
 * binary control tree) and, when provided, the `.frm`'s designer block (the
 * form's own textual properties). The module's code is untouched: that is
 * writeModule's job, and the two writes compose.
 */
export function writeFormDesigner(
	filePath: string,
	moduleName: string,
	frx: Buffer,
	frmDesignerBlock?: string,
): WriteResult {
	const streams = parseFormFrx(frx);
	if (!streams) {
		throw new Error('Not a .frx sidecar this importer understands.');
	}
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	const existing = readDesignerStorage(wb.cfb, module.name);
	if (!existing) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	wb.cfb.writeStreamInStorage(module.name, 'f', streams.f);
	wb.cfb.writeStreamInStorage(module.name, 'o', streams.o);
	if (frmDesignerBlock) {
		const merged = mergeVbFrameFromFrm(frmDesignerBlock, existing.vbFrame);
		wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(merged, wb.project.codePage));
	}
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

const VBFRAME_STREAM = '\x03VBFrame';

/** The designer storage's streams for a module, or undefined when it has none. */
function readDesignerStorage(
	cfb: Cfb,
	moduleName: string,
): { f: Buffer; o: Buffer; compObj?: Buffer; vbFrame: string } | undefined {
	try {
		const f = cfb.getStreamInStorage(moduleName, 'f');
		const o = cfb.getStreamInStorage(moduleName, 'o');
		const compObjName = '\x01CompObj';
		const compObj = cfb.hasStreamInStorage(moduleName, compObjName)
			? cfb.getStreamInStorage(moduleName, compObjName)
			: undefined;
		const vbFrameBytes = cfb.getStreamInStorage(moduleName, VBFRAME_STREAM);
		return { f, o, compObj, vbFrame: vbFrameBytes.toString('latin1') };
	} catch {
		return undefined;
	}
}

export function listSubs(filePath: string, moduleName: string): ProcedureEntry[] {
	const { source } = readModule(filePath, moduleName, true);
	const { body } = splitVbaSource(source);
	const out: ProcedureEntry[] = [];
	PROC_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	// Line numbers come from one forward pass. Counting them per match with
	// `body.slice(0, m.index).split('\n')` re-walked the module from the start
	// every time - quadratic, and on a 26,000-line class it made expanding the
	// module in the explorer cost about 370 ms of which 1.5 ms was the search.
	// Matches arrive in increasing order, so each character is counted once.
	let scanned = 0;
	let line = 1;
	while ((m = PROC_RE.exec(body)) !== null) {
		for (; scanned < m.index; scanned += 1) {
			if (body.charCodeAt(scanned) === 10) {
				line += 1;
			}
		}
		out.push({
			name: m[2],
			kind: m[1].replace(/\s+/g, ' ').trim(),
			line,
		});
	}
	return out;
}

export function getProtectionInfo(filePath: string): ProtectionInfo {
	const { cfb, project } = openWorkbook(filePath);
	return { isPasswordProtected: project.hasPassword, isSigned: detectSignature(cfb).present };
}

export function getModulesAndProtectionInfo(filePath: string): ProtectionInfo & { modules: ModuleEntry[] } {
	const { cfb, project } = openWorkbook(filePath);
	return {
		modules: project.modules.map(moduleEntry),
		isPasswordProtected: project.hasPassword,
		isSigned: detectSignature(cfb).present,
	};
}

export function listSheets(filePath: string): { sheets: SheetSummary[] } {
	return { sheets: sheetSurface(filePath).sheetSummaries() };
}

export function getWorkbookInfo(filePath: string): {
	sheets: SheetSummary[];
	namedRanges: NamedRange[];
	modules: ModuleEntry[];
	isPasswordProtected: boolean;
	isSigned: boolean;
} {
	const { container, cfb, project } = openWorkbook(filePath);
	// Only the OOXML Excel container has a READABLE sheet surface (.xlsb
	// keeps a binary workbook part); for every other shape the modules and
	// protection facts still answer.
	const xlsx = container.kind === 'excel' && container.xlsx?.hasSheetSurface()
		? container.xlsx
		: undefined;
	return {
		sheets: xlsx ? xlsx.sheetSummaries() : [],
		namedRanges: xlsx ? xlsx.definedNames() : [],
		modules: project.modules.map(moduleEntry),
		isPasswordProtected: project.hasPassword,
		isSigned: detectSignature(cfb).present,
	};
}

export function readCells(filePath: string, sheet: string, range: string): { data: CellValue[][] } {
	return { data: sheetSurface(filePath).readCells(sheet, range, true) };
}

export function readFormulas(filePath: string, sheet: string, range: string): { data: CellValue[][] } {
	return { data: sheetSurface(filePath).readCells(sheet, range, false) };
}

/**
 * Structural checks over the VBA project - the native equivalent of the
 * previous backend's validate(): every dir-declared module must resolve to a
 * readable stream, names must be unique, and the PROJECT stream must agree.
 */
export function validateWorkbook(filePath: string): { issues: string[] } {
	const issues: string[] = [];
	let wb: OpenWorkbook;
	try {
		wb = openWorkbook(filePath);
	} catch (err) {
		return { issues: [`VBA project could not be parsed: ${err instanceof Error ? err.message : String(err)}`] };
	}
	const seen = new Set<string>();
	for (const module of wb.project.modules) {
		const key = module.name.toLowerCase();
		if (seen.has(key)) {
			issues.push(`Duplicate module name: ${module.name}`);
		}
		seen.add(key);
		if (!module.name) {
			issues.push('Module with an empty name in the dir stream.');
		}
		const streamName = module.streamName || module.name;
		// Root-level fallback mirrors VbaProject.parse: legacy containers and
		// the Access synthetic CFB keep module streams outside a VBA storage.
		if (!wb.cfb.hasStreamInStorage('VBA', streamName) && !wb.cfb.hasStream(streamName)) {
			issues.push(`Module ${module.name} references missing stream '${streamName}'.`);
		}
		if (module.sourceHeader === '' && module.prefixBytes.length === 0) {
			issues.push(`Module ${module.name} has no readable source stream.`);
		}
	}
	return { issues };
}

// ----------------------------------------------------------------- write API

/** A name as the project's ANSI code page stores it ('?' for what it cannot express). */
function foldedModuleName(name: string, codePage: number): string {
	return decodeCodePage(encodeCodePage(name, codePage), codePage);
}

/**
 * A module name beyond the project's ANSI code page is legal: the unicode
 * dir records and the CFB stream name carry the real name, and the ANSI
 * records plus the PROJECT stream hold its '?'-folded projection - the same
 * shape Office itself produces. Verified against live Excel (2026-08-18):
 * the VBE lists the unicode name, Application.Run executes the module, and
 * an Excel re-save reads back intact. What CANNOT coexist are two modules
 * whose projections collide: the PROJECT stream would declare the same
 * folded name twice, which Excel treats as corruption.
 */
function assertFoldedNameDistinct(
	modules: readonly { name: string }[],
	codePage: number,
	name: string,
	previousName?: string,
): void {
	const folded = foldedModuleName(name, codePage).toLowerCase();
	const prevLower = previousName?.toLowerCase();
	const nameLower = name.toLowerCase();
	for (const other of modules) {
		const otherLower = other.name.toLowerCase();
		if (otherLower === nameLower || otherLower === prevLower) {
			continue; // exact duplicates are the project layer's own error
		}
		if (foldedModuleName(other.name, codePage).toLowerCase() === folded) {
			throw new Error(
				`Module name "${name}" cannot coexist with "${other.name}": this project's code page ` +
				`(${codePage}) stores both as "${foldedModuleName(name, codePage)}", and Excel treats the ` +
				'duplicate PROJECT declarations as corruption. Choose a name with a distinct stored form.',
			);
		}
	}
}

export function writeModule(
	filePath: string,
	moduleName: string,
	source: string,
	kind: 'standard' | 'class' = 'standard',
): WriteResult {
	// Callers may pass a bare body or a full export; strip any incoming header
	// so the workbook's own header is always the one that persists.
	const { body } = splitVbaSource(source);
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const existing = wb.project.getModule(moduleName);
	if (existing) {
		const { header } = splitVbaSource(existing.source);
		wb.project.setModuleSource(existing.name, joinVbaSource(header, body));
	} else {
		assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, moduleName);
		const header = kind === 'class'
			? synthesizeClassHeader(moduleName)
			: synthesizeStandardHeader(moduleName);
		wb.project.addModule(moduleName, joinVbaSource(header, body), kind === 'class' ? 'other' : 'standard');
	}
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

export function renameModule(filePath: string, moduleName: string, newName: string): WriteResult {
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, newName, moduleName);
	wb.project.renameModule(moduleName, newName);
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

export function deleteModule(filePath: string, moduleName: string): WriteResult {
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	wb.project.deleteModule(moduleName);
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

export function writeCells(
	filePath: string,
	sheet: string,
	startCell: string,
	data: CellValue[][],
): { ok: true } {
	// Mutates the package, so never the cached instance readers share.
	const container = openMacroContainer(fs.readFileSync(filePath));
	if (container.kind !== 'excel' || !container.xlsx) {
		throw new Error(
			`${path.basename(filePath)} is ${container.description}; cell writes need an OOXML Excel workbook.`,
		);
	}
	if (!container.xlsx.hasSheetSurface()) {
		throw new Error(
			`${path.basename(filePath)} is a binary Excel workbook (.xlsb); its worksheet data is `
			+ 'stored in a binary format XLIDE does not write. VBA editing is unaffected - save the '
			+ 'workbook as .xlsm to use the sheet and cell tools.',
		);
	}
	container.xlsx.writeCells(sheet, startCell, data);
	atomicWrite(filePath, container.xlsx.toBytes());
	return { ok: true };
}

/**
 * Create a new macro-enabled file by copying the bundled blank template for
 * its extension byte for byte. Overwrites `filePath` if it exists - callers
 * gate that: the New File command's save dialog confirms replacement
 * natively, and the agent tool refuses existing paths outright.
 */
export function createWorkbook(filePath: string, templatePath: string): { ok: true; path: string } {
	const template = fs.readFileSync(templatePath);
	atomicWrite(filePath, template);
	return { ok: true, path: filePath };
}
