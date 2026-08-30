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
import { parseFormPackage, writeFormPackage, walkPackages as walkOformsPackages, controlKindOfSite as oformsControlKind } from './oforms/formPackage';
import { siteName as oformsSiteName } from './oforms/formStream';
import { printFormMarkup as printOformsMarkup, parseFormMarkup as parseOformsMarkup, applyFormMarkup as applyOformsMarkup } from './oforms/markup';
import { formatPointsShortest } from './oforms/bytes';
import { composeNewForm } from './oforms/newForm';
import { renderFormPreviewHtml } from './oforms/preview';
import {
	addControlAt as designerAddControlAt,
	listFormProperties as designerListFormProperties,
	reconcileMarkupIdentities as designerReconcileMarkupIdentities,
	removeControl as designerRemoveControl,
	reparentControl as designerReparentControl,
	setControlGeometry as designerSetControlGeometry,
	setControlProperty as designerSetControlProperty,
	setFormSize as designerSetFormSize,
} from './oforms/designerOps';
import type { OformsTextCodec } from './oforms/records';
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
	/**
	 * True when the module carries `Attribute VB_PredeclaredId = True`, giving
	 * it a default instance so its own name is usable as a value. Absent means
	 * the attribute header was not read, never "no".
	 */
	predeclaredId?: boolean;
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

/**
 * The value of an `Attribute VB_*` header line.
 *
 * Both spellings count. The VBE quotes a string value (`VB_Name = "Ticket"`)
 * and leaves a boolean bare (`VB_PredeclaredId = True`), and reading only the
 * quoted form made every boolean attribute answer the empty string - which
 * silently disabled the document-module fallback below, whose whole job is to
 * recognise a host module by `PredeclaredId` and `Exposed` both being True.
 */
function attributeValue(source: string, attribute: string): string {
	const re = new RegExp(`^\\s*Attribute\\s+${attribute}\\s*=\\s*(?:"([^"]*)"|([^\\r\\n]*))`, 'im');
	const match = re.exec(source);
	return (match?.[1] ?? match?.[2] ?? '').trim();
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
	// Only vouch when the header is actually there. A module the container
	// stored without one has an UNKNOWN default instance, and answering
	// `false` would turn every predeclared class red (issue #47).
	if (/^\s*Attribute\s+VB_PredeclaredId\s*=/im.test(module.sourceHeader)) {
		entry.predeclaredId = /^True$/i.test(attributeValue(module.sourceHeader, 'VB_PredeclaredId'));
	}
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
	// VBA exposes EVERY control as a member of the form, however deeply it
	// nests - `Me.PickAir` works when PickAir sits inside a Frame - so the
	// member surface must walk the whole package tree. The flat top-level
	// read this replaced under-reported: nested controls were missing, and a
	// code-behind touching one was called undeclared.
	try {
		const pkg = parseFormPackage(cfb, [module.name], oformsCodec(project.codePage));
		const controls: { name: string; type: string }[] = [];
		walkOformsPackages(pkg, (surface) => {
			for (const surfaceEntry of surface.entries) {
				const name = oformsSiteName(surfaceEntry.site);
				if (!name) { continue; }
				controls.push({
					name,
					type: oformsControlKind(surfaceEntry.site, surfaceEntry.kind === 'record' ? surfaceEntry.record : undefined) === 'ActiveX'
						? 'ActiveX.Control'
						: `MSForms.${oformsControlKind(surfaceEntry.site, surfaceEntry.kind === 'record' ? surfaceEntry.record : undefined)}`,
				});
			}
			// Container controls are members too: the Frame, the MultiPage,
			// and each Page answer to their names on the form.
			for (const site of surface.form.sites) {
				const kind = oformsControlKind(site);
				if (kind !== 'Frame' && kind !== 'MultiPage' && kind !== 'Page') { continue; }
				const name = oformsSiteName(site);
				if (name) { controls.push({ name, type: `MSForms.${kind}` }); }
			}
		});
		entry.implicitMembers = controls;
	} catch {
		// The new engine could not read the storage; fall back to the proven
		// flat reader so behavior never regresses below what it was.
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
			// No designer storage, or a shape neither reader understands: the
			// entry simply carries no members, same as before.
		}
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

/** MBCS/UTF-16 codec bound to one project's code page. */
function oformsCodec(codePage: number): OformsTextCodec {
	return {
		decode: (bytes, compressed) =>
			compressed ? decodeCodePage(bytes, codePage) : bytes.toString('utf16le'),
		encode: (text, compressed) =>
			compressed ? encodeCodePage(text, codePage) : Buffer.from(text, 'utf16le'),
	};
}

/**
 * A form's design, projected to XLIDE form markup. Read natively from the
 * designer storage; no host is involved.
 */
export function readFormMarkup(filePath: string, moduleName: string): { markup: string } {
	const { cfb, project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!cfb.hasStoragePath([module.name])) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const pkg = parseFormPackage(cfb, [module.name], oformsCodec(project.codePage));
	const frame = decodeCodePage(cfb.getStreamInStorage(module.name, VBFRAME_STREAM), project.codePage);
	const captionFallback = /^\s*Caption\s*=\s*"([^"]*)"/m.exec(frame)?.[1];
	return { markup: printOformsMarkup(pkg, module.name, { captionFallback, vbFrame: vbFramePropsOf(frame) }) };
}

/** The form rendered as a self-contained HTML preview document. */
/** The form properties the VBFrame text carries, for the Properties pane. */
function vbFramePropsOf(frame: string): { showModal?: string; startUpPosition?: string; whatsThisButton?: string } {
	const numberOf = (key: string): string | undefined =>
		new RegExp(`^\\s*${key}\\s*=\\s*(-?\\d+)`, 'm').exec(frame)?.[1];
	const modal = numberOf('ShowModal');
	const whats = numberOf('WhatsThisButton');
	return {
		startUpPosition: numberOf('StartUpPosition') ?? '1',
		// Absent means the DEFAULT: modal True, the help button False.
		showModal: modal === undefined ? 'True' : (modal === '0' ? 'False' : 'True'),
		whatsThisButton: whats === undefined ? 'False' : (whats === '0' ? 'False' : 'True'),
	};
}

/** Sets one numeric VBFrame line, replacing it or inserting it before End. */
function setVbFrameLine(frame: string, key: string, value: string): string {
	const line = new RegExp(`^(\\s*${key}\\s*=\\s*).*$`, 'm');
	if (line.test(frame)) {
		return frame.replace(line, `$1${value}`);
	}
	const eol = frame.includes('\r\n') ? '\r\n' : '\n';
	return frame.replace(/^End\b/m, `   ${key}       =   ${value}${eol}End`);
}

export function readFormPreview(
	filePath: string,
	moduleName: string,
	selected?: string,
	markup?: string,
	identityPath?: string,
): { html: string } {
	const { cfb, project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!cfb.hasStoragePath([module.name])) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const pkg = parseFormPackage(cfb, [module.name], oformsCodec(project.codePage));
	const frame = decodeCodePage(cfb.getStreamInStorage(module.name, VBFRAME_STREAM), project.codePage);
	const caption = /^\s*Caption\s*=\s*"([^"]*)"/m.exec(frame)?.[1];
	const properties = designerListFormProperties(pkg, module.name, caption, vbFramePropsOf(frame));
	return {
		html: renderFormPreviewHtml(pkg, {
			formName: module.name,
			caption,
			selected,
			properties,
			// The designer may render from a scratch copy; its identity - what
			// F5 launches, what the state names - is the real workbook.
			identity: { workbook: identityPath ?? filePath, module: module.name },
			// The pane shows the document's own spelling; when no document
			// exists yet, the engine's canonical print stands in.
			markup: markup ?? readFormMarkup(filePath, moduleName).markup,
		}),
	};
}

/**
 * Applies an edited markup document back to the form's designer storage.
 * The document parses whole first - a parse error applies nothing - and the
 * apply is a name-keyed diff. Within the dialect's vocabulary the document
 * is TOTAL: an attribute quiet at its default means the default, so an edit
 * that returned a property to its default survives the save. Anything the
 * dialect cannot spell (pictures, foreign payloads) is never touched.
 */
export function applyFormMarkup(
	filePath: string,
	moduleName: string,
	markup: string,
): WriteResult & { applied: string[] } {
	const root = parseOformsMarkup(markup);
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!wb.cfb.hasStoragePath([module.name])) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const codec = oformsCodec(wb.project.codePage);
	const pkg = parseFormPackage(wb.cfb, [module.name], codec);
	// Renames and reparents pair IN PLACE before the name-keyed diff, so a
	// renamed control keeps what the dialect cannot spell - its picture, its
	// icon, an ActiveX payload - instead of dying as remove-plus-add.
	const reconciled = designerReconcileMarkupIdentities(pkg, root);
	const outcome = applyOformsMarkup(pkg, root);
	outcome.applied.unshift(...reconciled);

	// The form's own caption and the VBFrame trio are persisted in the
	// VBFrame text, so the document's <Form> attrs diff against that rather
	// than the f record. Absent trio attrs mean the DEFAULTS - the document
	// carries the whole form state, which is what makes text undo honest.
	const frame = decodeCodePage(wb.cfb.getStreamInStorage(module.name, VBFRAME_STREAM), wb.project.codePage);
	const currentCaption = /^\s*Caption\s*=\s*"([^"]*)"/m.exec(frame)?.[1];
	const documentCaption = root.attrs.get('Caption');
	let vbFrameUpdated = frame;
	let vbFrameChanged = false;
	if (documentCaption !== undefined && currentCaption !== undefined && documentCaption !== currentCaption) {
		vbFrameUpdated = vbFrameUpdated.replace(
			/^(\s*Caption\s*=\s*)"[^"]*"/m,
			`$1"${documentCaption.replace(/"/g, '')}"`,
		);
		vbFrameChanged = true;
		outcome.applied.push('Caption of the form');
	}
	const frameProps = vbFramePropsOf(frame);
	const sup = root.attrs.get('StartUpPosition') ?? '1';
	if (!/^[0-3]$/.test(sup)) {
		throw new Error(`StartUpPosition="${sup}" is not 0-3`);
	}
	if (sup !== (frameProps.startUpPosition ?? '1')) {
		vbFrameUpdated = setVbFrameLine(vbFrameUpdated, 'StartUpPosition', sup);
		vbFrameChanged = true;
		outcome.applied.push('StartUpPosition of the form');
	}
	for (const [attr, fallback] of [['ShowModal', 'True'], ['WhatsThisButton', 'False']] as const) {
		const text = root.attrs.get(attr) ?? fallback;
		if (!/^(true|false)$/i.test(text)) {
			throw new Error(`${attr}="${text}" is not True or False`);
		}
		const want = /^true$/i.test(text) ? 'True' : 'False';
		const current = attr === 'ShowModal' ? frameProps.showModal : frameProps.whatsThisButton;
		if (want !== (current ?? fallback)) {
			vbFrameUpdated = setVbFrameLine(vbFrameUpdated, attr, want === 'True' ? "-1  'True" : "0   'False");
			vbFrameChanged = true;
			outcome.applied.push(`${attr} of the form`);
		}
	}
	// A real form resize keeps the VBFrame's twips echo in step, as the
	// designer's own resize gesture always has.
	if (outcome.applied.includes('Form size')) {
		const width = Number(root.attrs.get('Width'));
		const height = Number(root.attrs.get('Height'));
		if (Number.isFinite(width)) {
			vbFrameUpdated = vbFrameUpdated.replace(/^(\s*ClientWidth\s*=\s*)\d+/m, `$1${Math.round(width * 20)}`);
		}
		if (Number.isFinite(height)) {
			vbFrameUpdated = vbFrameUpdated.replace(/^(\s*ClientHeight\s*=\s*)\d+/m, `$1${Math.round(height * 20)}`);
		}
		vbFrameChanged = vbFrameChanged || vbFrameUpdated !== frame;
	}

	if (outcome.applied.length === 0) {
		return { ok: true, signatureDropped: false, applied: [] };
	}
	writeFormPackage(wb.cfb, [module.name], pkg, codec);
	if (vbFrameChanged) {
		wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(vbFrameUpdated, wb.project.codePage));
	}
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped, applied: outcome.applied };
}

/**
 * One canvas gesture applied to a form: move or resize a control, add one at
 * a point, remove one, or resize the form itself. Each gesture is one
 * parse-mutate-write of the designer storage, through the same primitives
 * the markup apply uses.
 */
export function applyFormDesignerOp(
	filePath: string,
	moduleName: string,
	op:
		| { kind: 'geometry'; name: string; left?: number; top?: number; width?: number; height?: number }
		| { kind: 'add'; container: string; controlKind: string; left: number; top: number }
		| { kind: 'remove'; name: string }
		| { kind: 'reparent'; name: string; container: string; left: number; top: number }
		| { kind: 'setProp'; name: string; prop: string; value: string }
		| { kind: 'formSize'; width: number; height: number },
): WriteResult & { newName?: string } {
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!wb.cfb.hasStoragePath([module.name])) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const codec = oformsCodec(wb.project.codePage);
	const pkg = parseFormPackage(wb.cfb, [module.name], codec);
	let newName: string | undefined;
	if (op.kind === 'geometry') {
		const applied = designerSetControlGeometry(pkg, op.name, op);
		if (applied.length === 0) {
			return { ok: true, signatureDropped: false };
		}
	} else if (op.kind === 'add') {
		newName = designerAddControlAt(pkg, op.container, op.controlKind, op.left, op.top);
	} else if (op.kind === 'remove') {
		designerRemoveControl(pkg, op.name);
	} else if (op.kind === 'reparent') {
		designerReparentControl(pkg, op.name, op.container, op.left, op.top);
	} else if (op.kind === 'setProp') {
		if (op.name === '' && op.prop === 'Caption') {
			// The form's caption is persisted in the VBFrame text.
			const frame = decodeCodePage(wb.cfb.getStreamInStorage(module.name, VBFRAME_STREAM), wb.project.codePage);
			const updated = frame.replace(
				/^(\s*Caption\s*=\s*)"[^"]*"/m,
				`$1"${op.value.replace(/"/g, '')}"`,
			);
			if (updated === frame) {
				return { ok: true, signatureDropped: false };
			}
			wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(updated, wb.project.codePage));
		} else if (op.name === '' && ['ShowModal', 'WhatsThisButton', 'StartUpPosition'].includes(op.prop)) {
			// These live in the VBFrame text, like the caption. VB spells the
			// booleans 0 and -1.
			let text: string;
			if (op.prop === 'StartUpPosition') {
				if (!/^[0-3]$/.test(op.value)) {
					throw new Error(`StartUpPosition="${op.value}" is not 0, 1, 2, or 3`);
				}
				text = op.value;
			} else {
				if (!/^(true|false)$/i.test(op.value)) {
					throw new Error(`${op.prop}="${op.value}" is not True or False`);
				}
				text = /^true$/i.test(op.value) ? "-1  'True" : "0   'False";
			}
			const frame = decodeCodePage(wb.cfb.getStreamInStorage(module.name, VBFRAME_STREAM), wb.project.codePage);
			const updated = setVbFrameLine(frame, op.prop, text);
			if (updated === frame) {
				return { ok: true, signatureDropped: false };
			}
			wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(updated, wb.project.codePage));
		} else if (op.name === '' && (op.prop === 'Width' || op.prop === 'Height')) {
			const n = Number(op.value);
			if (!Number.isFinite(n)) {
				throw new Error(`${op.prop}="${op.value}" is not a number`);
			}
			const size = pkg.form.record.sizes.get('DisplayedSize') ?? { width: 0, height: 0 };
			const widthPt = op.prop === 'Width' ? n : Number(formatPointsShortest(size.width));
			const heightPt = op.prop === 'Height' ? n : Number(formatPointsShortest(size.height));
			designerSetFormSize(pkg, widthPt, heightPt);
			const frame = decodeCodePage(wb.cfb.getStreamInStorage(module.name, VBFRAME_STREAM), wb.project.codePage);
			const updated = frame
				.replace(/^(\s*ClientWidth\s*=\s*)\d+/m, `$1${Math.round(widthPt * 20)}`)
				.replace(/^(\s*ClientHeight\s*=\s*)\d+/m, `$1${Math.round(heightPt * 20)}`);
			wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(updated, wb.project.codePage));
		} else {
			const result = designerSetControlProperty(pkg, op.name, op.prop, op.value);
			if (result.applied.length === 0) {
				return { ok: true, signatureDropped: false };
			}
			newName = result.renamed;
		}
	} else {
		designerSetFormSize(pkg, op.width, op.height);
		// The VBFrame's client box repeats the size in twips and must follow.
		const frame = decodeCodePage(wb.cfb.getStreamInStorage(module.name, VBFRAME_STREAM), wb.project.codePage);
		const updated = frame
			.replace(/^(\s*ClientWidth\s*=\s*)\d+/m, `$1${Math.round(op.width * 20)}`)
			.replace(/^(\s*ClientHeight\s*=\s*)\d+/m, `$1${Math.round(op.height * 20)}`);
		wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(updated, wb.project.codePage));
	}
	writeFormPackage(wb.cfb, [module.name], pkg, codec);
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped, newName };
}

/**
 * A byte-true snapshot of one form's whole designer storage - every stream,
 * however deeply the containers nest - keyed by its path inside the module's
 * storage. The designer's undo walks these: restoring one puts back exactly
 * the bytes a gesture changed, and only those (the code face lives outside
 * the designer storage and is never touched).
 */
export function readFormDesignerSnapshot(
	filePath: string,
	moduleName: string,
): { streams: Record<string, string> } {
	const { cfb, project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!cfb.hasStoragePath([module.name])) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const streams: Record<string, string> = {};
	const walk = (rel: string[]): void => {
		const at = [module.name, ...rel];
		for (const name of cfb.listStreamsAtPath(at)) {
			streams[[...rel, name].join('/')] = cfb.getStreamAtPath(at, name).toString('base64');
		}
		for (const child of cfb.listStoragesAtPath(at)) {
			walk([...rel, child]);
		}
	};
	walk([]);
	return { streams };
}

/** Puts a designer snapshot back, byte for byte - the undo/redo restore. */
export function restoreFormDesignerSnapshot(
	filePath: string,
	moduleName: string,
	streams: Record<string, string>,
): WriteResult {
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (wb.cfb.hasStoragePath([module.name])) {
		wb.cfb.removeStorageAtPath([module.name]);
	}
	wb.cfb.addStorageAtPath([], module.name);
	const ensured = new Set<string>(['']);
	for (const key of Object.keys(streams).sort()) {
		const parts = key.split('/');
		const streamName = parts.pop()!;
		let parent: string[] = [];
		for (const part of parts) {
			const pathKey = [...parent, part].join('/');
			if (!ensured.has(pathKey)) {
				wb.cfb.addStorageAtPath([module.name, ...parent], part);
				ensured.add(pathKey);
			}
			parent = [...parent, part];
		}
		wb.cfb.setStreamAtPath([module.name, ...parts], streamName, Buffer.from(streams[key], 'base64'));
	}
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

/**
 * Creates a UserForm module natively: the module stream with its exported
 * header, the BaseClass registration, and a designer storage holding a
 * minimal FormControl, an empty object stream, the VBFrame, and the Forms
 * 2.0 CompObj - the same shape live Excel writes for a new form.
 */
export function addFormModule(
	filePath: string,
	moduleName: string,
	body = '',
): WriteResult {
	const wb = openWorkbookForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	if (wb.project.getModule(moduleName)) {
		throw new Error(`Module already exists: ${moduleName}`);
	}
	assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, moduleName);
	if (wb.cfb.hasStoragePath([moduleName])) {
		throw new Error(`A designer storage named ${moduleName} already exists.`);
	}
	const streams = composeNewForm({ name: moduleName });
	wb.project.addModule(
		moduleName,
		joinVbaSource(streams.header, body),
		'other',
		{ projectKeyword: 'BaseClass' },
	);
	wb.cfb.addStorageAtPath([], moduleName);
	wb.cfb.setStreamAtPath([moduleName], 'f', streams.f);
	wb.cfb.setStreamAtPath([moduleName], 'o', streams.o);
	wb.cfb.setStreamAtPath([moduleName], VBFRAME_STREAM, encodeCodePage(streams.vbFrame, wb.project.codePage));
	wb.cfb.setStreamAtPath([moduleName], '\x01CompObj', streams.compObj);
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

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
