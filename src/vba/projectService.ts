// Native project service: every VBA/cell operation XLIDE needs, implemented
// directly against the OOXML package and the MS-OVBA VBA project. This replaces
// the external backend entirely.
//
// Writes are atomic: the updated package is written to a sibling temp file and
// renamed over the original, so a failure part-way through never leaves a
// truncated project.

import * as fs from 'fs';
import * as path from 'path';
import { Cfb } from './cfb';
import { decodeCodePage, encodeCodePage } from './codePages';
import { parseFormPackage, writeFormPackage, walkPackages as walkOformsPackages, controlKindOfSite as oformsControlKind } from './oforms/formPackage';
import { siteName as oformsSiteName } from './oforms/formStream';
import { printFormMarkup as printOformsMarkup, parseFormMarkup as parseOformsMarkup, applyFormMarkup as applyOformsMarkup } from './oforms/markup';
import { formatPointsShortest } from './oforms/bytes';
import { composeNewForm } from './oforms/newForm';
import { renderFormPreviewHtml, renderFormSceneHtml } from './oforms/preview';
import {
	addControlAt as designerAddControlAt,
	listFormProperties as designerListFormProperties,
	reconcileMarkupIdentities as designerReconcileMarkupIdentities,
	removeControl as designerRemoveControl,
	reparentControl as designerReparentControl,
	setControlGeometry as designerSetControlGeometry,
	setGeometryBatch as designerSetGeometryBatch,
	setTabOrder as designerSetTabOrder,
	setZOrder as designerSetZOrder,
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
import { atomicWrite } from './atomicWrite';
import { attributeValue, joinVbaSource, listProcedures, splitVbaSource, type ProcedureEntry } from './moduleSource';
import {
	isVb6ProjectPath,
	listVb6Modules,
	listVb6Procedures,
	readVb6Module,
	readVb6Modules,
	readVb6FormHeader,
	validateVb6Project,
	writeVb6Module,
	type Vb6ModuleEntry,
	VB6_CODE_PAGE,
} from './vb6/vb6Project';
import { frmFrxRefs, parseFrmHeader, printFrmHeader, type FrmHeader } from './vb6/frmHeader';
import { readFrxRecords, type FrxValue } from './vb6/frx';
import { listFrmProperties, sceneOfFrmHeader, type FrxLookup } from './vb6/frmScene';
import { applyFrmDesignerOp, type FrmDesignerOp, type FrmDesignerOpResult } from './vb6/frmDesignerOps';

/**
 * A module's kind. The last three exist only in a VB6 project, whose manifest
 * names UserControls, PropertyPages and Designers beside its forms: they are
 * listed so the project reads whole, and their designers stay opaque.
 */
export type ModuleType = 'standard' | 'class' | 'document' | 'userform'
	| 'usercontrol' | 'propertypage' | 'designer';
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
	implicitMembers?: { name: string; type: string; array?: boolean }[];
	/**
	 * A VB6 designer's own class (`VB.Form`, `VB.MDIForm`, `VB.UserControl`),
	 * which decides what `Me` is and what its event handlers are called.
	 * Office forms carry none: a UserForm is always an MSForms.UserForm.
	 */
	designerClass?: string;
	/**
	 * True when the module carries `Attribute VB_PredeclaredId = True`, giving
	 * it a default instance so its own name is usable as a value. Absent means
	 * the attribute header was not read, never "no".
	 */
	predeclaredId?: boolean;
	/**
	 * The module's own file, for the containers whose modules ARE files (a
	 * VB6 project). Absent for a project module, which lives in a stream.
	 */
	filePath?: string;
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
// The module text helpers live in moduleSource.ts, shared with the VB6
// project reader; they are re-exported here so every caller keeps its import.
export { joinVbaSource, splitVbaSource } from './moduleSource';
export type { ProcedureEntry } from './moduleSource';
export { atomicWrite } from './atomicWrite';

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

interface OpenContainer {
	container: MacroContainer;
	cfb: Cfb;
	project: VbaProject;
}

// -------------------------------------------------------------- parse cache
//
// Re-reading and re-parsing the file IS the cost of a read now that the engine
// is in-process, and callers arrive in bursts that hit the same project: an
// explorer expansion is listModules + a protection probe + one listSubs per
// module, and every one of those re-opened the file. Reads share one parse per
// project, validated against (mtimeMs, size) on every call so out-of-band
// writers (Excel, git, another window) are always seen.
//
// Writes never touch the cache: they parse fresh, because a mutating save that
// fails halfway must not leave a poisoned parse behind for readers - and every
// mutation lands through atomicWrite, which drops the entry. The mtime check
// is the backstop for writers that bypass this process entirely.

interface ProjectCacheEntry {
	mtimeMs: number;
	size: number;
	container: MacroContainer;
	/** Built on first VBA access; sheet/cell reads never pay for the project. */
	cfb?: Cfb;
	project?: VbaProject;
}

/**
 * Small on purpose: an entry retains the package plus decompressed module
 * sources (a few MB for a large project), and a session's hot set is the
 * handful of projects whose trees or editors are open.
 */
const WORKBOOK_CACHE_MAX = 4;
const projectCache = new Map<string, ProjectCacheEntry>();
let cacheHits = 0;
let cacheMisses = 0;

function cachedPackage(filePath: string): ProjectCacheEntry {
	// Stat BEFORE reading: if a writer swaps the file between the stat and the
	// read, this entry holds the new bytes under the old mtime, so the next
	// call mismatches and rebuilds - a stale parse can never survive.
	const stat = fs.statSync(filePath);
	const hit = projectCache.get(filePath);
	if (hit && hit.mtimeMs === stat.mtimeMs && hit.size === stat.size) {
		cacheHits++;
		// Refresh recency: Map iteration order is insertion order.
		projectCache.delete(filePath);
		projectCache.set(filePath, hit);
		return hit;
	}
	cacheMisses++;
	const entry: ProjectCacheEntry = {
		mtimeMs: stat.mtimeMs,
		size: stat.size,
		container: openMacroContainer(fs.readFileSync(filePath)),
	};
	projectCache.delete(filePath);
	projectCache.set(filePath, entry);
	while (projectCache.size > WORKBOOK_CACHE_MAX) {
		const oldest = projectCache.keys().next().value;
		if (oldest === undefined) { break; }
		projectCache.delete(oldest);
	}
	return entry;
}

/** Shared read-only parse. Callers must not mutate the returned project. */
function openContainer(filePath: string): OpenContainer {
	if (isVb6ProjectPath(filePath)) {
		throw vb6ProjectRefusal(filePath);
	}
	const entry = cachedPackage(filePath);
	entry.cfb ??= entry.container.vbaCfb();
	entry.project ??= VbaProject.parse(entry.cfb);
	return { container: entry.container, cfb: entry.cfb, project: entry.project };
}

/** Fresh parse for mutating operations; never aliases the shared cache. */
function openContainerForWrite(filePath: string): OpenContainer {
	if (isVb6ProjectPath(filePath)) {
		throw vb6ProjectRefusal(filePath);
	}
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

/** Test hook: a VB6 form's parsed designer header through the service's own dispatch. */
export const readVb6FormHeaderForTests = readVb6FormHeader;

/**
 * Sidecar records a designer has placed that the file has not taken yet:
 * they follow the bytes on disk in order, from `base`, and reach the file
 * when the document saves (`appendVb6Sidecar`).
 */
export interface Vb6PendingSidecar {
	/** The sidecar's file name, as the header spells it. */
	file: string;
	/** The sidecar's byte length on disk when the first record was placed. */
	base: number;
	/** The records, base64, in offset order from `base`. */
	records: string[];
}

/**
 * A VB6 form rendered for the designer, from the header text the document
 * holds right now, the sidecar on disk, and the records a designer holds
 * pending until the document saves. The document IS the module's file, so
 * nothing is copied and nothing is applied: the header is parsed from
 * `text`, and the `.frx` it names is read beside `modulePath` for the
 * strings and pictures the header keeps there. The pane's markup is the
 * header itself, as written.
 */
export function readVb6FormPreview(
	modulePath: string,
	text: string,
	selected?: string,
	vbpPath?: string,
	pending?: Vb6PendingSidecar,
): { html: string; headerEnd: number } {
	const header = parseFrmHeader(text);
	if (!header) {
		throw new Error(`${path.basename(modulePath)} has no designer header (VERSION 5.00 / Begin VB.Form ... End).`);
	}
	const formName = header.form.name;
	const frx = vb6FrxLookup(header, path.dirname(modulePath), pending);
	const scene = sceneOfFrmHeader(header, { formName, frx });
	return {
		html: renderFormSceneHtml(scene, {
			formName,
			selected,
			properties: listFrmProperties(header, { formName, frx }),
			identity: { project: vbpPath ?? modulePath, module: formName },
			markup: printFrmHeader(header),
		}),
		headerEnd: header.endOffset,
	};
}

/** The sidecar a header names on its references, or undefined when it keeps nothing there. */
function vb6SidecarNamedBy(header: FrmHeader): string | undefined {
	return frmFrxRefs(header)[0]?.property.frx?.file;
}

/** The sidecar a designer's header names, else the one VB6 pairs with the module's extension (.frx, .ctx, .pgx, .dsx). */
export function vb6SidecarFileFor(modulePath: string, header: FrmHeader | undefined): string {
	const named = header ? vb6SidecarNamedBy(header) : undefined;
	if (named) {
		return named;
	}
	const ext = path.extname(modulePath);
	const sidecarExt = { '.ctl': '.ctx', '.pag': '.pgx', '.dsr': '.dsx' }[ext.toLowerCase()] ?? '.frx';
	return `${path.basename(modulePath, ext)}${sidecarExt}`;
}

/** The sidecar's bytes, or none when the file does not exist yet; any other failure is the caller's to see, never an empty file. */
function readVb6Sidecar(frxPath: string): Buffer {
	try {
		return fs.readFileSync(frxPath);
	} catch (err) {
		if ((err as NodeJS.ErrnoException).code === 'ENOENT') {
			return Buffer.alloc(0);
		}
		throw new Error(`Cannot read the form's sidecar ${path.basename(frxPath)}: ${err instanceof Error ? err.message : String(err)}`);
	}
}

/** Sidecars as last read, by path, good while the file's stamp and size hold: a render per gesture must not reread a picture set. */
const sidecarCache = new Map<string, { mtimeMs: number; size: number; blob: Buffer }>();

/** The sidecar's bytes for a render, cached by the file's stamp; undefined when there is no file to read. */
function readVb6SidecarForRender(frxPath: string): Buffer | undefined {
	let stat: fs.Stats;
	try {
		stat = fs.statSync(frxPath);
	} catch {
		return undefined;
	}
	const cached = sidecarCache.get(frxPath);
	if (cached && cached.mtimeMs === stat.mtimeMs && cached.size === stat.size) {
		return cached.blob;
	}
	try {
		const blob = fs.readFileSync(frxPath);
		sidecarCache.set(frxPath, { mtimeMs: stat.mtimeMs, size: stat.size, blob });
		return blob;
	} catch {
		return undefined;
	}
}

/** A sidecar record a gesture placed: the host holds it with the document until the save that writes it. */
export interface Vb6SidecarRecord {
	file: string;
	/** The sidecar's byte length on disk when the record was placed. */
	base: number;
	/** The record's offset: `base` plus the pending bytes before it. */
	offset: number;
	/** The record's bytes, base64: a length, then the string in the form's code page. */
	record: string;
}

/**
 * A designer gesture on a VB6 form: the header rewritten in the document's
 * text, the code kept (frmDesignerOps.ts). A string the header cannot hold
 * inline - a multi-line Text - becomes a sidecar record in the layout the
 * reader measures (a length, then the bytes) at the offset it will have
 * after the records already pending (`pendingBytes` of them), and the
 * header points there. Nothing is written here: the host keeps the record
 * with the document and `appendVb6Sidecar` writes it when the document
 * saves, so the file and its sidecar move together.
 */
export function applyVb6FormDesignerOp(
	modulePath: string,
	text: string,
	op: FrmDesignerOp,
	pendingBytes = 0,
): FrmDesignerOpResult & { sidecar?: Vb6SidecarRecord } {
	let sidecar: Vb6SidecarRecord | undefined;
	const result = applyFrmDesignerOp(text, op, {
		storeString: (value, header) => {
			const file = vb6SidecarFileFor(modulePath, header);
			const base = readVb6Sidecar(path.join(path.dirname(modulePath), file)).length;
			const bytes = encodeCodePage(value, VB6_CODE_PAGE);
			const long = bytes.length > 255;
			const length = long ? Buffer.alloc(4) : Buffer.from([bytes.length]);
			if (long) { length.writeUInt32LE(bytes.length, 0); }
			const offset = base + pendingBytes;
			sidecar = { file, base, offset, record: Buffer.concat([length, bytes]).toString('base64') };
			return { file, offset, long };
		},
	});
	return sidecar ? { ...result, sidecar } : result;
}

/**
 * Writes a document's pending sidecar records, at save: they go after the
 * bytes on disk in the order they were placed, which is the order their
 * offsets assume. Refused when the sidecar is not the length the records
 * were placed against - another program wrote it - because appending would
 * put every pending reference at the wrong offset.
 */
export function appendVb6Sidecar(modulePath: string, file: string, base: number, records: readonly string[]): { length: number } {
	const frxPath = path.join(path.dirname(modulePath), file);
	const blob = readVb6Sidecar(frxPath);
	if (blob.length !== base) {
		throw new Error(
			`${file} is ${blob.length} bytes but the form's pending values were placed against ${base}; `
			+ 'the sidecar changed on disk meanwhile, so those values were not written. Undo the edits that set them and set them again.',
		);
	}
	const out = Buffer.concat([blob, ...records.map((record) => Buffer.from(record, 'base64'))]);
	atomicWrite(frxPath, out);
	sidecarCache.delete(frxPath);
	return { length: out.length };
}

/**
 * The sidecar values a header points at, keyed by offset: the bytes on disk
 * with the pending records after them when they still fit (the file has not
 * moved on since they were placed). A form has one sidecar, named on every
 * reference; a missing or unreadable sidecar leaves the header to draw
 * alone (the strings it keeps there are then blank).
 */
function vb6FrxLookup(header: FrmHeader, dir: string, pending?: Vb6PendingSidecar): FrxLookup | undefined {
	const file = vb6SidecarNamedBy(header);
	if (!file) {
		return undefined;
	}
	let blob = readVb6SidecarForRender(path.join(dir, file)) ?? Buffer.alloc(0);
	if (pending && pending.file.toLowerCase() === file.toLowerCase() && blob.length === pending.base) {
		blob = Buffer.concat([blob, ...pending.records.map((record) => Buffer.from(record, 'base64'))]);
	}
	if (blob.length === 0) {
		return undefined;
	}
	const byOffset = new Map<number, FrxValue>();
	for (const record of readFrxRecords(header, blob, (b) => decodeCodePage(b, VB6_CODE_PAGE))) {
		byOffset.set(record.offset, record.value);
	}
	return (ref) => (ref.file.toLowerCase() === file.toLowerCase() ? byOffset.get(ref.offset) : undefined);
}

/** Test hooks; product code never reads these. */
export function projectCacheStatsForTests(): { hits: number; misses: number; size: number } {
	return { hits: cacheHits, misses: cacheMisses, size: projectCache.size };
}

export function resetProjectCacheForTests(): void {
	projectCache.clear();
	cacheHits = 0;
	cacheMisses = 0;
}

/** Write the mutated VBA project back into the container, atomically. */
function saveContainer(filePath: string, wb: OpenContainer): void {
	wb.project.save(wb.cfb);
	atomicContainerWrite(filePath, wb.container.toFileBytes(wb.cfb));
}

/**
 * Every in-process project mutation lands through here, so this is the one
 * place cache invalidation cannot be forgotten. (A caller-supplied path with
 * different casing would miss this delete; the per-call mtime check still
 * catches that, so a stale entry can cost a re-parse but never stale data.)
 */
function atomicContainerWrite(filePath: string, data: Buffer): void {
	atomicWrite(filePath, data);
	projectCache.delete(filePath);
}

// ------------------------------------------------------------------ read API

/** A VB6 module entry in the project vocabulary. */
function vb6ModuleEntry(entry: Vb6ModuleEntry): ModuleEntry {
	const out: ModuleEntry = { name: entry.name, type: entry.type, filePath: entry.filePath };
	if (entry.source !== undefined) { out.source = entry.source; }
	if (entry.implicitMembers) { out.implicitMembers = entry.implicitMembers; }
	if (entry.designerClass !== undefined) { out.designerClass = entry.designerClass; }
	if (entry.predeclaredId !== undefined) { out.predeclaredId = entry.predeclaredId; }
	return out;
}

function vb6ProjectRefusal(filePath: string): Error {
	return new Error(
		`${path.basename(filePath)} is a Visual Basic 6 project: its modules are files beside it, `
		+ 'and this project operation has no meaning for them.',
	);
}

export function listModules(filePath: string): ModuleEntry[] {
	if (isVb6ProjectPath(filePath)) {
		return listVb6Modules(filePath).map(vb6ModuleEntry);
	}
	const { cfb, project } = openContainer(filePath);
	return project.modules.map((module) => moduleEntryWithDesigner(cfb, project, module));
}

export function readModules(filePath: string, full = false): ModuleEntry[] {
	if (isVb6ProjectPath(filePath)) {
		return readVb6Modules(filePath, full).map(vb6ModuleEntry);
	}
	const { cfb, project } = openContainer(filePath);
	const out: ModuleEntry[] = [];
	for (const module of project.modules) {
		try {
			const entry = moduleEntryWithDesigner(cfb, project, module);
			entry.source = full ? module.source : splitVbaSource(module.source).body;
			out.push(entry);
		} catch {
			// Keep project-wide reads best-effort at the module boundary.
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
		// One entry per control: MSForms names are unique across the whole
		// form, so a name already taken is the SAME control reached twice -
		// a container arrives once with its record and again in the site
		// sweep below, and listing it twice would be a duplicate member.
		const named = new Set<string>();
		const take = (name: string | undefined, type: string): void => {
			if (!name || named.has(name.toLowerCase())) { return; }
			named.add(name.toLowerCase());
			controls.push({ name, type });
		};
		walkOformsPackages(pkg, (surface) => {
			for (const surfaceEntry of surface.entries) {
				const kind = oformsControlKind(
					surfaceEntry.site,
					surfaceEntry.kind === 'record' ? surfaceEntry.record : undefined,
				);
				take(oformsSiteName(surfaceEntry.site),
					kind === 'ActiveX' ? 'ActiveX.Control' : `MSForms.${kind}`);
			}
			// Container controls are members too: the Frame, the MultiPage,
			// and each Page answer to their names on the form. Most arrive
			// with the entries above; this catches one whose surface carries
			// no record of its own.
			for (const site of surface.form.sites) {
				const kind = oformsControlKind(site);
				if (kind !== 'Frame' && kind !== 'MultiPage' && kind !== 'Page') { continue; }
				take(oformsSiteName(site), `MSForms.${kind}`);
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
	if (isVb6ProjectPath(filePath)) {
		return { source: readVb6Module(filePath, moduleName, full).source ?? '' };
	}
	const { project } = openContainer(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	return { source: full ? module.source : splitVbaSource(module.source).body };
}

/**
 * A form's export pair, composed natively from the project: the `.frm` text
 * (designer block from the VBFrame stream, `OleObjectBlob` naming the sidecar,
 * then the module's own attributes and code) and the `.frx` sidecar packaging
 * the designer storage's binary streams.
 */
export function readFormExport(filePath: string, moduleName: string): { frm: string; frx: Buffer } {
	const { cfb, project } = openContainer(filePath);
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
	return { frm, frx: composeFormFrx(cfb, module.name, designer.vbFrame) };
}

/**
 * Writes a form's designer back into the project from a `.frx` sidecar (the
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
	const wb = openContainerForWrite(filePath);
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
	// The container storages travel too: a sidecar's Frame and MultiPage
	// children replace the module's own, class CLSIDs and all - the flat
	// pair alone silently dropped every container's contents (hunt eight).
	if (streams.tree) {
		const tree = streams.tree;
		for (const child of wb.cfb.listChildrenAtPath([module.name])) {
			if (child.kind === 'storage') { wb.cfb.removeStorageAtPath([module.name, child.name]); }
		}
		const plant = (srcPath: string[], dstPath: string[]): void => {
			for (const child of tree.listChildrenAtPath(srcPath)) {
				if (child.kind === 'stream') {
					if (srcPath.length === 0) { continue; } // f/o written above; root CompObj stays the project's
					wb.cfb.setStreamAtPath(dstPath, child.name, tree.getStreamAtPath(srcPath, child.name));
				} else {
					wb.cfb.addStorageAtPath(dstPath, child.name, tree.storageClsidAtPath([...srcPath, child.name]));
					plant([...srcPath, child.name], [...dstPath, child.name]);
				}
			}
		};
		plant([], [module.name]);
	}
	if (frmDesignerBlock) {
		const merged = mergeVbFrameFromFrm(frmDesignerBlock, existing.vbFrame);
		wb.cfb.writeStreamInStorage(module.name, VBFRAME_STREAM, encodeCodePage(merged, wb.project.codePage));
	}
	saveContainer(filePath, wb);
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
	const { cfb, project } = openContainer(filePath);
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
	const { cfb, project } = openContainer(filePath);
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
			identity: { project: identityPath ?? filePath, module: module.name },
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
/**
 * The markup lines one named control occupies: a single self-closing line, or
 * an open line through its matching close at the SAME indent, which is all the
 * printer emits. Copy and delete both work on these spans, so a container's
 * children travel with it without either having to understand nesting.
 */
function formElementBlock(
	lines: readonly string[],
	name: string,
): { start: number; end: number; tag: string } {
	const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
	const opener = new RegExp(`^(\\s*)<(\\w+) Name="${escaped}"`);
	let at = -1;
	let indent = '';
	let tag = '';
	for (let i = 0; i < lines.length; i++) {
		const found = opener.exec(lines[i]);
		if (found) { at = i; indent = found[1]; tag = found[2]; break; }
	}
	if (at < 0) { throw new Error(`no control named ${name}`); }
	if (/\/>\s*$/.test(lines[at])) { return { start: at, end: at, tag }; }
	const closing = `${indent}</${tag}>`;
	for (let i = at + 1; i < lines.length; i++) {
		if (lines[i] === closing) { return { start: at, end: i, tag }; }
	}
	throw new Error(`${name}: its element is not closed`);
}

/**
 * DELETING A SELECTION in one step: every named element's span is cut, so a
 * multi-control delete is a single document edit and a single undo - and a
 * container drags its children out with it. A name nested inside another name
 * being deleted is dropped rather than refused: the parent already takes it.
 *
 * All-or-nothing on unknown names, so a stale selection cannot half-delete.
 */
export function removeFormControls(
	filePath: string,
	moduleName: string,
	names: readonly string[],
): WriteResult & { removed: string[] } {
	const markup = readFormMarkup(filePath, moduleName).markup;
	const lines = markup.split('\r\n');
	const found = names.map((name) => ({ name, block: formElementBlock(lines, name) }));
	for (const one of found) {
		if (one.block.tag === 'Page') {
			throw new Error(`${one.name} is a Page; remove pages through the form markup`);
		}
	}
	// Keep only the outermost spans, and only one entry per span.
	const seen = new Set<number>();
	const outermost = found.filter((one) => {
		if (seen.has(one.block.start)) { return false; }
		seen.add(one.block.start);
		return !found.some((other) => other !== one
			&& other.block.start < one.block.start && one.block.end <= other.block.end);
	});
	if (outermost.length === 0) {
		return { ok: true, signatureDropped: false, removed: [] };
	}
	// Last span first, so a cut never shifts an earlier span's index.
	for (const one of [...outermost].sort((a, b) => b.block.start - a.block.start)) {
		lines.splice(one.block.start, one.block.end - one.block.start + 1);
	}
	const result = applyFormMarkup(filePath, moduleName, lines.join('\r\n'));
	return { ok: true, signatureDropped: result.signatureDropped, removed: outermost.map((one) => one.name) };
}

/**
 * COPY AND PASTE, as a document transform: each named control's element is
 * cloned in place, given fresh names, nudged so the copy is visible, and the
 * whole document re-applied. Working on the MARKUP rather than on the sites
 * is deliberate - the clone then travels the same authoring path every other
 * addition takes, containers included, whose children come along and are
 * renamed with them.
 *
 * Returns the new top-level names, so the canvas can select what it pasted.
 */
export function duplicateFormControls(
	filePath: string,
	moduleName: string,
	names: readonly string[],
	offsetPt = 6,
): WriteResult & { newNames: string[] } {
	const markup = readFormMarkup(filePath, moduleName).markup;
	const lines = markup.split('\r\n');
	const taken = new Set<string>();
	for (const line of lines) {
		const found = /\sName="([^"]*)"/.exec(line);
		if (found) { taken.add(found[1].toLowerCase()); }
	}
	const freshName = (kind: string): string => {
		for (let i = 1; ; i++) {
			const candidate = `${kind}${i}`;
			if (!taken.has(candidate.toLowerCase())) {
				taken.add(candidate.toLowerCase());
				return candidate;
			}
		}
	};
	const newNames: string[] = [];
	// Last block first, so an earlier clone never shifts a later block's index.
	const blocks = names.map((name) => ({ name, block: formElementBlock(lines, name) }))
		.sort((a, b) => b.block.start - a.block.start);
	for (const { name, block } of blocks) {
		if (block.tag === 'Page') {
			throw new Error(`${name} is a Page; add pages through the form markup`);
		}
		const clone = lines.slice(block.start, block.end + 1).map((line) => {
			const kind = /^\s*<(\w+)/.exec(line)?.[1];
			return line.replace(/(\sName=")([^"]*)(")/, (_all, lead: string, was: string, tail: string) => {
				const fresh = freshName(kind ?? was.replace(/\d+$/, ''));
				if (was.toLowerCase() === name.toLowerCase()) { newNames.push(fresh); }
				return `${lead}${fresh}${tail}`;
			});
		});
		// Only the copy's own position moves; a child's is relative to it.
		clone[0] = clone[0]
			.replace(/\sLeft="(-?[\d.]+)"/, (_a, v: string) => ` Left="${Number(v) + offsetPt}"`)
			.replace(/\sTop="(-?[\d.]+)"/, (_a, v: string) => ` Top="${Number(v) + offsetPt}"`);
		lines.splice(block.end + 1, 0, ...clone);
	}
	if (newNames.length === 0) {
		return { ok: true, signatureDropped: false, newNames: [] };
	}
	const result = applyFormMarkup(filePath, moduleName, lines.join('\r\n'));
	return { ok: true, signatureDropped: result.signatureDropped, newNames: newNames.reverse() };
}

export function applyFormMarkup(
	filePath: string,
	moduleName: string,
	markup: string,
): WriteResult & { applied: string[] } {
	const root = parseOformsMarkup(markup);
	const wb = openContainerForWrite(filePath);
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
	saveContainer(filePath, wb);
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
		| { kind: 'formSize'; width: number; height: number }
		| { kind: 'geometryBatch'; items: readonly { name: string; left?: number; top?: number; width?: number; height?: number }[] }
		| { kind: 'zOrder'; name: string; toFront: boolean }
		| { kind: 'tabOrder'; container: string; names: readonly string[] },
): WriteResult & { newName?: string } {
	const wb = openContainerForWrite(filePath);
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
	} else if (op.kind === 'geometryBatch') {
		if (designerSetGeometryBatch(pkg, op.items).length === 0) {
			return { ok: true, signatureDropped: false };
		}
	} else if (op.kind === 'zOrder') {
		if (designerSetZOrder(pkg, op.name, op.toFront).length === 0) {
			return { ok: true, signatureDropped: false };
		}
	} else if (op.kind === 'tabOrder') {
		if (designerSetTabOrder(pkg, op.container, op.names).length === 0) {
			return { ok: true, signatureDropped: false };
		}
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
	saveContainer(filePath, wb);
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
	const { cfb, project } = openContainer(filePath);
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
	const wb = openContainerForWrite(filePath);
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
	saveContainer(filePath, wb);
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
	const wb = openContainerForWrite(filePath);
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
	saveContainer(filePath, wb);
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
	if (isVb6ProjectPath(filePath)) {
		return listVb6Procedures(filePath, moduleName);
	}
	const { source } = readModule(filePath, moduleName, true);
	return listProcedures(splitVbaSource(source).body);
}

export function getProtectionInfo(filePath: string): ProtectionInfo {
	if (isVb6ProjectPath(filePath)) {
		// Files on disk: nothing to lock and nothing to sign.
		return { isPasswordProtected: false, isSigned: false };
	}
	const { cfb, project } = openContainer(filePath);
	return { isPasswordProtected: project.hasPassword, isSigned: detectSignature(cfb).present };
}

export function getModulesAndProtectionInfo(filePath: string): ProtectionInfo & { modules: ModuleEntry[] } {
	if (isVb6ProjectPath(filePath)) {
		return { modules: listModules(filePath), isPasswordProtected: false, isSigned: false };
	}
	const { cfb, project } = openContainer(filePath);
	return {
		modules: project.modules.map(moduleEntry),
		isPasswordProtected: project.hasPassword,
		isSigned: detectSignature(cfb).present,
	};
}

export function listSheets(filePath: string): { sheets: SheetSummary[] } {
	return { sheets: sheetSurface(filePath).sheetSummaries() };
}

export function getProjectInfo(filePath: string): {
	sheets: SheetSummary[];
	namedRanges: NamedRange[];
	modules: ModuleEntry[];
	isPasswordProtected: boolean;
	isSigned: boolean;
} {
	if (isVb6ProjectPath(filePath)) {
		return { sheets: [], namedRanges: [], modules: listModules(filePath), isPasswordProtected: false, isSigned: false };
	}
	const { container, cfb, project } = openContainer(filePath);
	// Only the OOXML Excel container has a READABLE sheet surface (.xlsb
	// keeps a binary project part); for every other shape the modules and
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
export function validateProject(filePath: string): { issues: string[] } {
	if (isVb6ProjectPath(filePath)) {
		return validateVb6Project(filePath);
	}
	const issues: string[] = [];
	let wb: OpenContainer;
	try {
		wb = openContainer(filePath);
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
	if (isVb6ProjectPath(filePath)) {
		writeVb6Module(filePath, moduleName, source);
		return { ok: true, signatureDropped: false };
	}
	// Callers may pass a bare body or a full export; strip any incoming header
	// so the project's own header is always the one that persists.
	const { body } = splitVbaSource(source);
	const wb = openContainerForWrite(filePath);
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
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped };
}

export function renameModule(filePath: string, moduleName: string, newName: string): WriteResult {
	if (isVb6ProjectPath(filePath)) {
		throw new Error(`Renaming a module of a VB6 project is not supported yet; rename ${moduleName} in the .vbp and its file.`);
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, newName, moduleName);
	wb.project.renameModule(moduleName, newName);
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped };
}

export function deleteModule(filePath: string, moduleName: string): WriteResult {
	if (isVb6ProjectPath(filePath)) {
		throw new Error(`Deleting a module of a VB6 project is not supported yet; remove ${moduleName} from the .vbp and delete its file.`);
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	wb.project.deleteModule(moduleName);
	saveContainer(filePath, wb);
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
	atomicContainerWrite(filePath, container.xlsx.toBytes());
	return { ok: true };
}

/**
 * Create a new macro-enabled file by copying the bundled blank template for
 * its extension byte for byte. Overwrites `filePath` if it exists - callers
 * gate that: the New File command's save dialog confirms replacement
 * natively, and the agent tool refuses existing paths outright.
 */
export function createProject(filePath: string, templatePath: string): { ok: true; path: string } {
	const template = fs.readFileSync(templatePath);
	atomicContainerWrite(filePath, template);
	return { ok: true, path: filePath };
}
