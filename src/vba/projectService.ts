// Native project service: every VBA/cell operation XLIDE needs, implemented
// directly against the OOXML package and the MS-OVBA VBA project. This replaces
// the external backend entirely.
//
// Writes are atomic: the updated package is written to a sibling temp file and
// renamed over the original, so a failure part-way through never leaves a
// truncated project.

import { hostPlatform } from './hostPlatform';
import * as path from 'path';
import { evictOldest } from '../util/boundedMap';
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
import { NoVbaProjectError, openMacroContainer, type AccessContainerDesign, type MacroContainer } from './macroContainer';
import { ZipArchive } from './zip';
import {
	AccessVbaWriter,
	accessDesignModuleName,
	type AccessDesignEntry,
} from './access/accessVbaWriter';
import {
	accessDesignProperties,
	applyAccessDesignMarkup,
	printAccessDesignMarkup,
	setAccessDesignPropertyText,
} from './access/accessDesignMarkup';
import type { MarkupElement } from './oforms/markup';
import { sceneOfAccessDesign } from './access/accessDesignScene';
import {
	addDesignControl,
	removeDesignControl,
	reparentDesignControl,
	setDesignProperty,
	setDesignTabOrder,
	setDesignZOrder,
} from './access/accessDesignEdit';
import { ACCESS_DESIGN_CLASSES, accessDesignObjectName, type AccessDesign } from './access/accessDesign';
import {
	detectSignature,
	synthesizeClassHeader,
	synthesizeStandardHeader,
	VbaProject,
	type VbaModule,
} from './vbaProject';
import { XlsxWorkbook, type CellValue, type NamedRange, type SheetSummary } from './xlsx';
import type { ShapeEdit, ShapeInfo } from './shapes';
import { editSlideShape, listSlideShapes, presentationSlides, requireSlide } from './pptShapes';
import {
	WORD_SHAPE_MACRO_REFUSAL,
	defaultStory,
	documentStories,
	editStoryShape,
	listStoryShapes,
	requireStory,
} from './docShapes';
import { atomicWrite } from './atomicWrite';
import { addVbaProjectToPackage } from './addVbaProject';
import {
	HOST_LIBRARIES,
	buildMsFormsReference,
	buildRegisteredReference,
	hasMsFormsReference,
	MSFORMS_REFERENCE_NAME,
	readProjectReferences,
	removeReferenceRecords,
	type VbaProjectReference,
} from './vbaProjectReferences';
import { attributeValue, joinVbaSource, listProcedures, splitVbaSource, type ProcedureEntry } from './moduleSource';
import { readFolderAnnotation } from './folderAnnotation';
import { validateVbaModuleName } from '../vbaSourceScan';
import { readAttributeAnnotations } from '../analyzer/annotations/attributeAnnotations';
import { applyAttributeAnnotations } from '../analyzer/annotations/attributeRewriter';
import { parseModule } from '../analyzer/parser/parseModule';
import { hostTokenForFileName } from '../analyzer/host/hostRegistry';
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
	// An Access form or report: the design, with the module behind it when it
	// has one. Access names that module `Form_<design>` / `Report_<design>`,
	// which is what the VBE lists, so the design and its code are one entry.
	| 'accessform' | 'accessreport'
	| 'usercontrol' | 'propertypage' | 'designer';
export type DocumentType = 'workbook' | 'worksheet' | 'chart' | 'document';

export interface ModuleEntry {
	name: string;
	type: ModuleType;
	documentType?: DocumentType;
	source?: string;
	/**
	 * A form's designer-declared controls, read natively from the designer
	 * storage inside vbaProject.bin, or an Access design's named sections and
	 * controls. Present only where the designer parsed cleanly; absent means
	 * "not known", never "none". `eventClass` is the class a member's events
	 * come from where that is not its `type` (an Access report's sections).
	 */
	implicitMembers?: { name: string; type: string; array?: boolean; eventClass?: string }[];
	/**
	 * The class a designer makes the module, where that is not an
	 * MSForms.UserForm: a VB6 designer's own (`VB.Form`, `VB.MDIForm`,
	 * `VB.UserControl`) or an Access design's (`Access.Form`,
	 * `Access.Report`). It decides what `Me` is. A UserForm carries none.
	 */
	designerClass?: string;
	/**
	 * The PROJECT's conditional compilation arguments (the VBE project
	 * property), repeated on every entry of one read so a caller that only sees
	 * the module list still gets them. Absent when the project declares none.
	 */
	projectConditionalConstants?: string;
	/**
	 * The type libraries the PROJECT references, repeated on every entry of
	 * one read the way the conditional constants are. A project that
	 * references another application's library can name its types, so the
	 * analyzer needs the list to know that `Excel.Application` means anything
	 * in a Word document. Absent when the read did not reach them.
	 */
	projectReferences?: VbaProjectReference[];
	/**
	 * True when the module carries `Attribute VB_PredeclaredId = True`, giving
	 * it a default instance so its own name is usable as a value. Absent means
	 * the attribute header was not read, never "no".
	 */
	predeclaredId?: boolean;
	/**
	 * The normalized `@Folder` annotation from the module's declarations
	 * section, which places it in the explorer's folder layout. Absent when
	 * the module names no folder, which puts it at the project's root.
	 */
	folder?: string;
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
	/** Hidden attributes an annotation in the code set, in words a notice can carry. */
	attributeChanges?: string[];
}

const WORKBOOK_CLSID = '{00020819-0000-0000-C000-000000000046}';
const WORKSHEET_CLSID = '{00020820-0000-0000-C000-000000000046}';
const CHART_CLSID = '{00020821-0000-0000-C000-000000000046}';
const WORD_DOCUMENT_CLSID = '{00020906-0000-0000-C000-000000000046}';
/** The base the VBE writes on every class module it creates. */
const CLASS_CLSID = '{FCFB3D2A-A0FA-1068-A738-08002B3371B5}';
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
		// The VBE's own class: whatever its attributes, it is no document.
		if (upper.includes(CLASS_CLSID)) { return 'standard'; }
	}
	// Host code-behind that names no CLSID: Word's ThisDocument declares
	// VB_Base = "1Normal.ThisDocument" (measured against a live-authored
	// .docm). Office marks every document module PredeclaredId + Exposed, so
	// with a base named the pair is the host-generic document signature (forms
	// are Exposed = False and caught above). Without a base it is not: a class
	// can have both - a factory in an add-in, or `'@PredeclaredId` and
	// `'@Exposed` written through XLIDE - and was taken for a document, listed
	// as one and refused a rename or a delete.
	if (vbBase
		&& /^True$/i.test(attributeValue(source, 'VB_PredeclaredId'))
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
	const folder = folderOfModule(module);
	if (folder) { entry.folder = folder; }
	return entry;
}

/**
 * The module's `@Folder` annotation, read from the header prefix and NOTHING
 * more. `sourceHeader` is already inflated for the attribute checks above, so
 * this costs nothing; reaching for `module.source` would break the property
 * docs/xlide_performance_budgets.md calls out by name - "touching
 * `module.source` in a path that only needs names or types silently doubles
 * every read in the extension". It does: falling back to the whole source
 * doubled a cold `listModules` over the test corpus (16 ms to 33 ms) and 7x'd
 * one workbook, inflating 1.4 MB that a listing has no other use for, and
 * changed not one answer.
 *
 * The bound is the prefix: an annotation below the first ~4 KB of a module is
 * not found, and the module sits at the project root. `@Folder` is a
 * top-of-module convention - Rubberduck writes it as the first line - so that
 * is a module whose declarations run for a hundred lines BEFORE naming its
 * folder. The failure is the same place an unannotated module lands, which
 * makes it visible and fixable by moving the annotation up.
 */
function folderOfModule(module: VbaModule): string | undefined {
	return readFolderAnnotation(module.sourceHeader, { truncated: true }).folder;
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
const PROJECT_CACHE_MAX = 4;
const projectCache = new Map<string, ProjectCacheEntry>();
let cacheHits = 0;
let cacheMisses = 0;

function cachedPackage(filePath: string): ProjectCacheEntry {
	// Stat BEFORE reading: if a writer swaps the file between the stat and the
	// read, this entry holds the new bytes under the old mtime, so the next
	// call mismatches and rebuilds - a stale parse can never survive.
	const stat = hostPlatform().stat(filePath);
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
		container: openMacroContainer(hostPlatform().readFile(filePath)),
	};
	projectCache.delete(filePath);
	projectCache.set(filePath, entry);
	evictOldest(projectCache, PROJECT_CACHE_MAX);
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

/**
 * The same read, for callers that list what a project holds: a container with
 * no VBA project in it yet answers `undefined` rather than throwing.
 *
 * Having no code is an ordinary state of a macro-enabled file, not a failure
 * to read one - Excel writes no vbaProject.bin part at all into a workbook
 * saved as .xlsm before the first macro exists. Listing surfaces say "nothing
 * yet"; surfaces asked for a NAMED module or a write still refuse, through
 * {@link noVbaProjectRefusal}, so the reason reaches the user as itself.
 */
function openContainerIfAnyVba(filePath: string): OpenContainer | undefined {
	try {
		return openContainer(filePath);
	} catch (err) {
		if (err instanceof NoVbaProjectError) {
			return undefined;
		}
		throw err;
	}
}

/**
 * The container alone, for callers that ask about the FILE rather than the
 * code in it - which host it is, what designs it carries. A file with no VBA
 * project in it still answers every one of those.
 */
function containerOf(filePath: string): MacroContainer {
	if (isVb6ProjectPath(filePath)) {
		throw vb6ProjectRefusal(filePath);
	}
	return cachedPackage(filePath).container;
}

/**
 * Whether the file holds a VBA project at all, which is a different question
 * from whether that project holds any modules: a blank Access database has an
 * empty project, and a freshly saved .xlsm has none. Only the first can take
 * a new module, so callers that offer to write one ask this.
 */
export function hasVbaProject(filePath: string): boolean {
	if (isVb6ProjectPath(filePath)) {
		return true;
	}
	return openContainerIfAnyVba(filePath) !== undefined;
}

/**
 * The refusal a named read or a write gets from a file with no VBA in it.
 * XLIDE cannot start a project in a container that has none - what makes one
 * is a vbaProject.bin whose document modules match the file's own sheets or
 * stories - so the message says what to do instead rather than implying the
 * operation might work on a retry.
 */
function noVbaProjectRefusal(filePath: string, err: NoVbaProjectError, hostApp: string): Error {
	return new Error(
		`${path.basename(filePath)} has no VBA in it yet: it is ${err.containerDescription} that has `
		+ 'never held any code, so there is nothing to read and nowhere to write. XLIDE cannot put the '
		+ `first macro into a file that has none - write one in ${hostApp} and save, and the project is `
		+ "XLIDE's from there.",
	);
}

/** The application a container belongs to, for the refusal above. */
function containerAppName(container: MacroContainer): string {
	return { excel: 'Excel', word: 'Word', powerpoint: 'PowerPoint', access: 'Access' }[container.kind];
}

/**
 * A read of something NAMED inside the project. Turns the container seam's
 * "no VBA project" state into a refusal that names the file and says what to
 * do, so nothing downstream sees the raw internal message.
 */
function requireVbaProject(filePath: string): OpenContainer {
	try {
		return openContainer(filePath);
	} catch (err) {
		if (err instanceof NoVbaProjectError) {
			// The container itself parsed - it is only the project inside it
			// that is absent - so this is a cache hit, not a second read.
			throw noVbaProjectRefusal(filePath, err, containerAppName(containerOf(filePath)));
		}
		throw err;
	}
}

/** Fresh parse for mutating operations; never aliases the shared cache. */
function openContainerForWrite(filePath: string): OpenContainer {
	if (isVb6ProjectPath(filePath)) {
		throw vb6ProjectRefusal(filePath);
	}
	const container = openMacroContainer(hostPlatform().readFile(filePath));
	if (!container.writable) {
		throw new Error(`${path.basename(filePath)} is ${container.description}.`);
	}
	let cfb: Cfb;
	try {
		cfb = container.vbaCfb();
	} catch (err) {
		if (err instanceof NoVbaProjectError) {
			throw noVbaProjectRefusal(filePath, err, containerAppName(container));
		}
		throw err;
	}
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

/**
 * The Access form or report a module name stands for, or undefined when the
 * container has none. An Access design is named `Form_<design>` in the module
 * list, which is what the VBE calls the module behind it.
 */
function accessDesignFor(
	filePath: string,
	moduleName: string,
): { entry: AccessDesignEntry } | undefined {
	if (isVb6ProjectPath(filePath)) {
		return undefined;
	}
	const design = containerOf(filePath).designs?.().find(
		(entry) => entry.moduleName.toLowerCase() === moduleName.toLowerCase(),
	);
	if (!design) {
		return undefined;
	}
	const found = new AccessVbaWriter(hostPlatform().readFile(filePath)).designs().find(
		(entry) => entry.name === design.name && entry.kind === design.kind,
	);
	return found ? { entry: found } : undefined;
}

/**
 * The design name behind what the tree calls a module. A design is listed as
 * the module Access binds its code to, `Form_Orders`, so a new name typed with
 * or without that prefix means the same design: the prefix is Access's, not
 * the user's to drop.
 */
function accessDesignRename(
	kind: 'form' | 'report',
	newName: string,
): { design: string; module: string } {
	const prefix = accessDesignModuleName(kind, '');
	const design = newName.toLowerCase().startsWith(prefix.toLowerCase())
		? newName.slice(prefix.length)
		: newName;
	return { design, module: accessDesignModuleName(kind, design) };
}

/** Apply markup to an Access design and write the database back. */
function applyAccessMarkup(
	filePath: string,
	designName: string,
	root: MarkupElement,
): WriteResult & { applied: string[] } {
	const writer = new AccessVbaWriter(hostPlatform().readFile(filePath));
	let applied: string[] = [];
	writer.editDesign(designName, (design) => {
		const outcome = applyAccessDesignMarkup(
			design,
			root,
			writer.designPrototypesFor(designName),
			() => hostPlatform().randomBytes(16),
		);
		applied = outcome.applied;
		return outcome.design;
	});
	// Through the one write that cannot forget to invalidate the cache.
	atomicContainerWrite(filePath, writer.toBuffer());
	return { ok: true, signatureDropped: false, applied };
}

/**
 * A designer gesture on an Access design.
 *
 * Every gesture the canvas makes lands here and the designer then re-reads the
 * markup, so a gesture is the design edit it stands for and takes the same
 * path a markup edit takes. The canvas measures in points and Access stores
 * twips, which is the twentyfold everywhere below.
 */
function applyAccessDesignerOp(
	filePath: string,
	designName: string,
	op: FormDesignerOp,
): WriteResult & { newName?: string } {
	const writer = new AccessVbaWriter(hostPlatform().readFile(filePath));
	let newName: string | undefined;
	writer.editDesign(designName, (design) => {
		switch (op.kind) {
			case 'geometry':
				return setAccessGeometry(design, op.name, op);
			case 'geometryBatch':
				return op.items.reduce(
					(next: AccessDesign, item) => setAccessGeometry(next, item.name, item), design,
				);
			case 'setProp':
				// The pane sends what a reader would type, so it lands the same
				// way the markup's own attribute does.
				return setAccessDesignPropertyText(design, op.name || undefined, op.prop, op.value);
			case 'zOrder':
				return setDesignZOrder(design, op.name, op.toFront);
			case 'tabOrder':
				return setDesignTabOrder(design, op.container, op.names);
			case 'reparent':
				return reparentDesignControl(
					design, op.name, op.container, twips(op.left), twips(op.top),
				);
			case 'formSize':
				// Access sizes a form by its sections rather than by one outer
				// box: the width is the design's, the height the detail band's.
				return setDesignProperty(
					setDesignProperty(design, undefined, 'Width', twips(op.width)),
					'Detail', 'Height', twips(op.height),
				);
			case 'remove':
				return removeDesignControl(design, op.name);
			case 'add': {
				newName = nextAccessControlName(design, op.controlKind);
				return addDesignControl(
					design, op.controlKind, newName, hostPlatform().randomBytes(16),
					writer.designPrototypesFor(designName),
					{ left: twips(op.left), top: twips(op.top) },
				);
			}
		}
	});
	atomicContainerWrite(filePath, writer.toBuffer());
	return newName === undefined
		? { ok: true, signatureDropped: false }
		: { ok: true, signatureDropped: false, newName };
}

function twips(points: number): number {
	return Math.max(0, Math.round(points * 20));
}

function setAccessGeometry(
	design: AccessDesign,
	name: string,
	box: { left?: number; top?: number; width?: number; height?: number },
): AccessDesign {
	let next = design;
	for (const [property, value] of [
		['Left', box.left], ['Top', box.top], ['Width', box.width], ['Height', box.height],
	] as Array<[string, number | undefined]>) {
		if (value !== undefined) {
			next = setDesignProperty(next, name, property, twips(value));
		}
	}
	return next;
}

/** The next free `TextBox0`-style name for a control the canvas just dropped. */
function nextAccessControlName(design: AccessDesign, controlKind: string): string {
	const taken = new Set(design.objects
		.map(accessDesignObjectName)
		.filter((name): name is string => name !== undefined)
		.map((name) => name.toLowerCase()));
	for (let n = 0; ; n += 1) {
		const candidate = `${controlKind}${n}`;
		if (!taken.has(candidate.toLowerCase())) {
			return candidate;
		}
	}
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
		return hostPlatform().readFileIfPresent(frxPath) ?? Buffer.alloc(0);
	} catch (err) {
		throw new Error(`Cannot read the form's sidecar ${path.basename(frxPath)}: ${err instanceof Error ? err.message : String(err)}`);
	}
}

/** Sidecars as last read, by path, good while the file's stamp and size hold: a render per gesture must not reread a picture set. */
const sidecarCache = new Map<string, { mtimeMs: number; size: number; blob: Buffer }>();

/** The sidecar's bytes for a render, cached by the file's stamp; undefined when there is no file to read. */
function readVb6SidecarForRender(frxPath: string): Buffer | undefined {
	let stat;
	try {
		stat = hostPlatform().statIfPresent(frxPath);
	} catch {
		return undefined;
	}
	if (!stat) {
		return undefined;
	}
	const cached = sidecarCache.get(frxPath);
	if (cached && cached.mtimeMs === stat.mtimeMs && cached.size === stat.size) {
		return cached.blob;
	}
	try {
		const blob = hostPlatform().readFile(frxPath);
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
	if (entry.folder !== undefined) { out.folder = entry.folder; }
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
	// A file with no VBA in it lists no modules, which is the truth about it;
	// it is not a file that failed to load. See openContainerIfAnyVba.
	const open = openContainerIfAnyVba(filePath);
	if (!open) {
		return [];
	}
	const { container, cfb, project } = open;
	const entries = project.modules.map((module) => moduleEntryWithDesigner(cfb, project, module));
	return withHostDesigns(container, entries, project.codePage);
}

/**
 * The type libraries a project references, as `Tools > References` lists
 * them. A project that references another application's library can name its
 * types, so this is what says whether `Excel.Application` means anything in a
 * Word document.
 *
 * A VB6 project keeps its references in the .vbp manifest rather than a dir
 * stream, and none is read here yet, so it answers with an empty list rather
 * than a wrong one.
 */
export function listReferences(filePath: string): VbaProjectReference[] {
	if (isVb6ProjectPath(filePath)) {
		return [];
	}
	const open = openContainerIfAnyVba(filePath);
	return open ? readProjectReferences(open.project.dirStream) : [];
}

/**
 * Give a project a reference to another Office application's type library,
 * so its VBA can name that application's types early-bound.
 *
 * Late binding needs no reference at all - `CreateObject("Excel.Application")`
 * into an `Object` works from any project - so this is for code that names
 * the library, which is what the VBE's Tools > References dialog is for.
 * Adding one the project already has is a no-op rather than a duplicate.
 */
export function addReference(filePath: string, library: string): { ok: true; added: boolean; name: string } {
	const known = HOST_LIBRARIES[library.toLowerCase()];
	if (!known) {
		throw new Error(
			`'${library}' is not a library XLIDE can add; it knows ${Object.values(HOST_LIBRARIES).map((one) => one.name).join(', ')}.`,
		);
	}
	if (isVb6ProjectPath(filePath)) {
		throw new Error('A VB6 project keeps its references in the .vbp manifest, which XLIDE does not write.');
	}
	// A container's own host library is implicit: neither Excel nor Word
	// declares it in the project's records, and Tools > References shows it
	// checked and greyed. Writing one would be a second, redundant copy.
	if (hostTokenForFileName(filePath) === library.toLowerCase()) {
		throw new Error(`A ${known.name} file already has the ${known.name} object library; it is implicit in the project.`);
	}
	const wb = openContainerForWrite(filePath);
	const guid = known.guid.toLowerCase();
	if (readProjectReferences(wb.project.dirStream).some((one) => one.libid.toLowerCase().includes(guid))) {
		return { ok: true, added: false, name: known.name };
	}
	wb.project.addReferenceRecords(buildRegisteredReference(known));
	saveContainer(filePath, wb);
	return { ok: true, added: true, name: known.name };
}

/**
 * Take a reference away again, as unchecking it in `Tools > References` does.
 *
 * The library is named as the project knows it (`Excel`, `Office`, `stdole`)
 * or as a host token (`excel`), which are the same word for the four Office
 * applications. Removing one the project does not have is a no-op rather than
 * an error, the way adding one it already has is.
 *
 * Code that still names the library stops compiling, which is the VBE's
 * behaviour too and what `missing-library-reference` then reports. The one
 * refusal is Microsoft Forms while the project still has a UserForm: that
 * reference is what makes a form instantiable, and without it nothing in the
 * project compiles - so XLIDE, which adds it with the first form, will not
 * take it away underneath one.
 *
 * Measured against both hosts over COM: a workbook XLIDE added a reference to
 * and then removed opens in Excel with its four modules and its original
 * reference list, and a document XLIDE cut Excel out of opens in Word with
 * Normal - a project-kind reference sitting right beside the one that went -
 * and every other reference intact. Both hosts also list `VBA` and their own
 * application library, neither of which is in the dir stream: those are
 * implicit, which is why {@link addReference} refuses to write one.
 */
export function removeReference(filePath: string, library: string): { ok: true; removed: boolean; name: string } {
	if (isVb6ProjectPath(filePath)) {
		throw new Error('A VB6 project keeps its references in the .vbp manifest, which XLIDE does not write.');
	}
	const wb = openContainerForWrite(filePath);
	const wanted = library.trim().toLowerCase();
	const guid = HOST_LIBRARIES[wanted]?.guid.toLowerCase();
	const matches = (reference: VbaProjectReference): boolean =>
		reference.name.toLowerCase() === wanted
		|| (guid !== undefined && reference.libid.toLowerCase().includes(guid));
	const present = readProjectReferences(wb.project.dirStream).filter(matches);
	if (present.length === 0) {
		return { ok: true, removed: false, name: HOST_LIBRARIES[wanted]?.name ?? library };
	}
	if (present.some((reference) => reference.name.toLowerCase() === MSFORMS_REFERENCE_NAME.toLowerCase())) {
		const forms = wb.project.modules
			.filter((module) => moduleEntry(module).type === 'userform')
			.map((module) => module.name);
		if (forms.length > 0) {
			throw new Error(
				`${forms.length === 1 ? `Form ${forms[0]} needs` : `Forms ${forms.join(', ')} need`} `
				+ 'the Microsoft Forms library. Removing the reference would leave a form the host '
				+ 'cannot instantiate, and nothing in the project would compile.',
			);
		}
	}
	wb.project.replaceReferenceSection(removeReferenceRecords(wb.project.dirStream, matches));
	saveContainer(filePath, wb);
	return { ok: true, removed: true, name: present[0].name };
}

export function readModules(filePath: string, full = false): ModuleEntry[] {
	if (isVb6ProjectPath(filePath)) {
		return readVb6Modules(filePath, full).map(vb6ModuleEntry);
	}
	const open = openContainerIfAnyVba(filePath);
	return open ? readModulesFromContainer(open, full) : [];
}

/**
 * The modules of a macro container given its bytes rather than its path: a
 * workbook as git committed it, say. Same entries as {@link readModules}
 * gives for the file on disk, so the two can be compared module by module.
 */
export function readModulesFromBuffer(data: Buffer, full = false): ModuleEntry[] {
	const container = openMacroContainer(data);
	const cfb = container.vbaCfb();
	return readModulesFromContainer({ container, cfb, project: VbaProject.parse(cfb) }, full);
}

function readModulesFromContainer({ container, cfb, project }: OpenContainer, full: boolean): ModuleEntry[] {
	const out: ModuleEntry[] = [];
	const constants = project.conditionalConstantsRaw || undefined;
	// Read once and shared by every entry, the way the constants are.
	const references = readProjectReferences(project.dirStream);
	for (const module of project.modules) {
		try {
			const entry = moduleEntryWithDesigner(cfb, project, module);
			entry.source = full ? module.source : splitVbaSource(module.source).body;
			entry.projectConditionalConstants = constants;
			if (references.length > 0) { entry.projectReferences = references; }
			out.push(entry);
		} catch {
			// Keep project-wide reads best-effort at the module boundary.
			continue;
		}
	}
	return withHostDesigns(container, out, project.codePage, constants);
}

/**
 * The host's own forms and reports, folded into a module listing.
 *
 * Only Access has any. Its designs are objects in the database and the module
 * behind one is optional, so a design with code is the module Access names for
 * it - retyped, so the tree shows it as the form it is - and a design without
 * code is an entry of its own with no source behind it. Every other host keeps
 * its forms in the project, where they are already modules.
 */
function withHostDesigns(
	container: MacroContainer,
	entries: ModuleEntry[],
	projectCodePage: number,
	constants?: string,
): ModuleEntry[] {
	const designs = container.designs?.() ?? [];
	if (designs.length === 0) {
		return entries;
	}
	const byModule = new Map(designs.map((design) => [design.moduleName.toLowerCase(), design]));
	const out = entries.map((entry) => {
		const design = byModule.get(entry.name.toLowerCase());
		return design
			? { ...entry, type: designModuleType(design.kind), ...designFacts(design, projectCodePage) }
			: entry;
	});
	const covered = new Set(out.map((entry) => entry.name.toLowerCase()));
	for (const design of designs) {
		if (covered.has(design.moduleName.toLowerCase())) {
			continue;
		}
		// A design Access has never opened the code window for has no module
		// at all. It is still a design, and the designer is what reaches it.
		const entry: ModuleEntry = {
			name: design.moduleName,
			type: designModuleType(design.kind),
			...designFacts(design, projectCodePage),
		};
		if (constants) { entry.projectConditionalConstants = constants; }
		out.push(entry);
	}
	return out;
}

/**
 * What the design says about the module behind it that the module's text
 * never does: the class it is, which is what `Me` means there, and the
 * sections and controls that are members of that class.
 */
function designFacts(
	design: AccessContainerDesign,
	projectCodePage: number,
): Pick<ModuleEntry, 'designerClass' | 'implicitMembers'> {
	const members = design.members(projectCodePage);
	return {
		designerClass: ACCESS_DESIGN_CLASSES[design.kind],
		...(members ? { implicitMembers: members } : {}),
	};
}

function designModuleType(kind: 'form' | 'report'): ModuleType {
	return kind === 'form' ? 'accessform' : 'accessreport';
}

/**
 * A module entry, with a userform's controls read from its designer storage.
 * The storage is named after the module and its `f` stream is the MS-OFORMS
 * FormControl; the module's own text never mentions the controls at all.
 */
/**
 * Where a form's designer storage sits: beside the project's VBA storage.
 * That is the root of a vbaProject.bin, but a legacy file holds the project a
 * level down - in `_VBA_PROJECT_CUR` in an .xls, in `Macros` in a .doc - and
 * a designer created at the file's root instead belonged to no project: Excel
 * opened such an .xls without the form or any other code module (measured
 * 2026-09-19).
 */
function designerPath(cfb: Cfb, moduleName: string): string[] {
	if (!cfb.hasStoragePath(['VBA'])) {
		const holder = cfb.listStoragesAtPath([]).find((name) => cfb.hasStoragePath([name, 'VBA']));
		if (holder) {
			return [holder, moduleName];
		}
	}
	return [moduleName];
}

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
		const pkg = parseFormPackage(cfb, designerPath(cfb, module.name), oformsCodec(project.codePage));
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
	const { project } = requireVbaProject(filePath);
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
	const { cfb, project } = requireVbaProject(filePath);
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
		const designer = designerPath(wb.cfb, module.name);
		for (const child of wb.cfb.listChildrenAtPath(designer)) {
			if (child.kind === 'storage') { wb.cfb.removeStorageAtPath([...designer, child.name]); }
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
		plant([], designer);
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
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		return {
			markup: printAccessDesignMarkup(design.entry.design, design.entry.name, design.entry.kind),
		};
	}
	const { cfb, project } = requireVbaProject(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!cfb.hasStoragePath(designerPath(cfb, module.name))) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const pkg = parseFormPackage(cfb, designerPath(cfb, module.name), oformsCodec(project.codePage));
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
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		const { design: parsed, name, kind } = design.entry;
		const options: Parameters<typeof renderFormSceneHtml>[1] = {
			formName: name,
			properties: accessDesignProperties(parsed, name, kind),
			markup: markup ?? printAccessDesignMarkup(parsed, name, kind),
		};
		if (selected !== undefined) { options.selected = selected; }
		if (identityPath !== undefined) {
			options.identity = { project: identityPath, module: moduleName };
		}
		return { html: renderFormSceneHtml(sceneOfAccessDesign(parsed, name, kind), options) };
	}
	const { cfb, project } = requireVbaProject(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!cfb.hasStoragePath(designerPath(cfb, module.name))) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const pkg = parseFormPackage(cfb, designerPath(cfb, module.name), oformsCodec(project.codePage));
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
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		return applyAccessMarkup(filePath, design.entry.name, root);
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!wb.cfb.hasStoragePath(designerPath(wb.cfb, module.name))) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const codec = oformsCodec(wb.project.codePage);
	const pkg = parseFormPackage(wb.cfb, designerPath(wb.cfb, module.name), codec);
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
	writeFormPackage(wb.cfb, designerPath(wb.cfb, module.name), pkg, codec);
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
/** One gesture the designer canvas makes, whatever host the form belongs to. */
export type FormDesignerOp =
	| { kind: 'geometry'; name: string; left?: number; top?: number; width?: number; height?: number }
	| { kind: 'add'; container: string; controlKind: string; left: number; top: number }
	| { kind: 'remove'; name: string }
	| { kind: 'reparent'; name: string; container: string; left: number; top: number }
	| { kind: 'setProp'; name: string; prop: string; value: string }
	| { kind: 'formSize'; width: number; height: number }
	| { kind: 'geometryBatch'; items: readonly { name: string; left?: number; top?: number; width?: number; height?: number }[] }
	| { kind: 'zOrder'; name: string; toFront: boolean }
	| { kind: 'tabOrder'; container: string; names: readonly string[] };

export function applyFormDesignerOp(
	filePath: string,
	moduleName: string,
	op: FormDesignerOp,
): WriteResult & { newName?: string } {
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		return applyAccessDesignerOp(filePath, design.entry.name, op);
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const module = wb.project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	if (!wb.cfb.hasStoragePath(designerPath(wb.cfb, module.name))) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const codec = oformsCodec(wb.project.codePage);
	const pkg = parseFormPackage(wb.cfb, designerPath(wb.cfb, module.name), codec);
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
	writeFormPackage(wb.cfb, designerPath(wb.cfb, module.name), pkg, codec);
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
	const { cfb, project } = requireVbaProject(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	const designer = designerPath(cfb, module.name);
	if (!cfb.hasStoragePath(designer)) {
		throw new Error(`Module has no designer storage: ${moduleName}`);
	}
	const streams: Record<string, string> = {};
	const walk = (rel: string[]): void => {
		const at = [...designer, ...rel];
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
	const designer = designerPath(wb.cfb, module.name);
	if (wb.cfb.hasStoragePath(designer)) {
		wb.cfb.removeStorageAtPath(designer);
	}
	wb.cfb.addStorageAtPath(designer.slice(0, -1), module.name);
	const ensured = new Set<string>(['']);
	for (const key of Object.keys(streams).sort()) {
		const parts = key.split('/');
		const streamName = parts.pop()!;
		let parent: string[] = [];
		for (const part of parts) {
			const pathKey = [...parent, part].join('/');
			if (!ensured.has(pathKey)) {
				wb.cfb.addStorageAtPath([...designer, ...parent], part);
				ensured.add(pathKey);
			}
			parent = [...parent, part];
		}
		wb.cfb.setStreamAtPath([...designer, ...parts], streamName, Buffer.from(streams[key], 'base64'));
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
	kind: 'form' | 'report' = 'form',
): WriteResult & { moduleName: string } {
	// An Access database keeps its forms and reports as database objects, so a
	// new one is a design rather than a module with a designer storage. The
	// tree lists it as the module Access binds its code to, which is the name
	// the caller has to open it by.
	if (!isVb6ProjectPath(filePath) && containerOf(filePath).kind === 'access') {
		const named = accessDesignRename(kind, moduleName);
		const writer = new AccessVbaWriter(hostPlatform().readFile(filePath));
		writer.addDesign(named.design, kind);
		atomicContainerWrite(filePath, writer.toBuffer());
		return { ok: true, signatureDropped: false, moduleName: named.module };
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	if (wb.project.getModule(moduleName)) {
		throw new Error(`Module already exists: ${moduleName}`);
	}
	assertValidModuleName(wb.container, moduleName);
	assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, moduleName);
	// No module has the name, so a designer storage that does is an orphan:
	// earlier versions of XLIDE left one behind when they deleted or renamed a
	// form. Nothing reads it, and refusing the name over it left no way forward.
	const designer = designerPath(wb.cfb, moduleName);
	if (wb.cfb.hasStoragePath(designer)) {
		wb.cfb.removeStorageAtPath(designer);
	}
	// A project holding a form must reference the Microsoft Forms library, or
	// its host cannot instantiate the form and nothing in the project compiles,
	// while every reader here still finds the file healthy. This is not Excel's
	// alone: a UserForm in a Word or PowerPoint project needs the same library,
	// and none of the blank templates carries it. A project that gains its
	// first form through XLIDE had no such reference before.
	if (!hasMsFormsReference(wb.project.dirStream)) {
		wb.project.addReferenceRecords(buildMsFormsReference());
	}
	const streams = composeNewForm({ name: moduleName });
	wb.project.addModule(
		moduleName,
		joinVbaSource(streams.header, body),
		'other',
		{ projectKeyword: 'BaseClass' },
	);
	wb.cfb.addStorageAtPath(designer.slice(0, -1), moduleName);
	wb.cfb.setStreamAtPath(designer, 'f', streams.f);
	wb.cfb.setStreamAtPath(designer, 'o', streams.o);
	wb.cfb.setStreamAtPath(designer, VBFRAME_STREAM, encodeCodePage(streams.vbFrame, wb.project.codePage));
	wb.cfb.setStreamAtPath(designer, '\x01CompObj', streams.compObj);
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped, moduleName };
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
	const open = openContainerIfAnyVba(filePath);
	// No project is nothing to lock and nothing to sign either.
	if (!open) {
		return { isPasswordProtected: false, isSigned: false };
	}
	return { isPasswordProtected: open.project.hasPassword, isSigned: detectSignature(open.cfb).present };
}

export function getModulesAndProtectionInfo(filePath: string): ProtectionInfo & { modules: ModuleEntry[] } {
	if (isVb6ProjectPath(filePath)) {
		return { modules: listModules(filePath), isPasswordProtected: false, isSigned: false };
	}
	const open = openContainerIfAnyVba(filePath);
	if (!open) {
		return { modules: [], isPasswordProtected: false, isSigned: false };
	}
	return {
		modules: open.project.modules.map(moduleEntry),
		isPasswordProtected: open.project.hasPassword,
		isSigned: detectSignature(open.cfb).present,
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
	// The sheet surface of a workbook with no VBA in it reads perfectly well,
	// so it is answered whether or not there is a project behind it.
	const container = containerOf(filePath);
	const open = openContainerIfAnyVba(filePath);
	// Only the OOXML Excel container has a READABLE sheet surface (.xlsb
	// keeps a binary project part); for every other shape the modules and
	// protection facts still answer.
	const xlsx = container.kind === 'excel' && container.xlsx?.hasSheetSurface()
		? container.xlsx
		: undefined;
	return {
		sheets: xlsx ? xlsx.sheetSummaries() : [],
		namedRanges: xlsx ? xlsx.definedNames() : [],
		modules: open ? open.project.modules.map(moduleEntry) : [],
		isPasswordProtected: open?.project.hasPassword ?? false,
		isSigned: open ? detectSignature(open.cfb).present : false,
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
	let wb: OpenContainer | undefined;
	try {
		wb = openContainerIfAnyVba(filePath);
	} catch (err) {
		return { issues: [`VBA project could not be parsed: ${err instanceof Error ? err.message : String(err)}`] };
	}
	// Nothing to check and nothing wrong: a file with no VBA in it has no
	// structural problems, and reporting one would read as damage.
	if (!wb) {
		return { issues: [] };
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
	// A form needs the Microsoft Forms library declared, or its host cannot
	// instantiate it and the project will not compile - a failure every reader
	// here is blind to, because the modules and the designer storage are all
	// perfectly readable without it. True of every host that has UserForms,
	// not Excel alone.
	const forms = wb.project.modules
		.filter((module) => moduleEntry(module).type === 'userform')
		.map((module) => module.name);
	if (forms.length > 0 && !hasMsFormsReference(wb.project.dirStream)) {
		const named = forms.length === 1 ? `Form ${forms[0]}` : `Forms ${forms.join(', ')}`;
		issues.push(
			`${named} need the Microsoft Forms library, which this project does not reference. `
			+ 'The host application cannot instantiate the form or compile the project until it is '
			+ 'added (Tools > References > Microsoft Forms 2.0 Object Library).',
		);
	}
	return { issues };
}

// ----------------------------------------------------------------- write API

/** A name as the project's ANSI code page stores it ('?' for what it cannot express). */
function foldedModuleName(name: string, codePage: number): string {
	return decodeCodePage(encodeCodePage(name, codePage), codePage);
}

/**
 * A name the VBE would give a module: an identifier, not a reserved word, at
 * most 31 characters. The tree's prompts ask the same; an agent's tool call
 * and a folder import reach the engine without them, and wrote `Bad Name`,
 * `Sub` and `1Leading` into projects - which Excel opens, but no code can
 * name such a module and the VBE never makes one. Access is left as it was:
 * a module there is a database object, named by Access's rules.
 */
function assertValidModuleName(container: MacroContainer, name: string): void {
	if (container.kind === 'access') {
		return;
	}
	const problem = validateVbaModuleName(name);
	if (problem) {
		throw new Error(`"${name}" is not a valid module name. ${problem}.`);
	}
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
				`(${codePage}) stores both as "${foldedModuleName(name, codePage)}", and Office treats the ` +
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
	let attributeChanges: string[] = [];
	if (existing) {
		const { header } = splitVbaSource(existing.source);
		const written = withAnnotatedAttributes(joinVbaSource(header, body));
		attributeChanges = written.changes;
		wb.project.setModuleSource(existing.name, written.text);
	} else {
		assertValidModuleName(wb.container, moduleName);
		assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, moduleName);
		const header = kind === 'class'
			? synthesizeClassHeader(moduleName)
			: synthesizeStandardHeader(moduleName);
		const written = withAnnotatedAttributes(joinVbaSource(header, body));
		attributeChanges = written.changes;
		wb.project.addModule(moduleName, written.text, kind === 'class' ? 'other' : 'standard');
	}
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped, ...(attributeChanges.length > 0 ? { attributeChanges } : {}) };
}

/**
 * The hidden attributes an annotation in the code names, written into the
 * module on its way to the container.
 *
 * A VBA module carries attributes the code pane never shows and the editor
 * gives no way to set - `VB_PredeclaredId`, `VB_Description`, `VB_UserMemId`.
 * A comment naming one, `'@PredeclaredId` or `'@Description("...")`, is
 * reviewable in the code; this is where it becomes the attribute. Inert unless
 * the developer writes an annotation: no annotation, no change, byte for byte.
 */
function withAnnotatedAttributes(source: string): { text: string; changes: string[] } {
	const annotations = readAttributeAnnotations(source);
	if (annotations.annotations.length === 0) {
		return { text: source, changes: [] };
	}
	const applied = applyAttributeAnnotations(source, annotations);
	return {
		text: applied.text,
		changes: applied.changes.map((change) => (change.from === undefined
			? `${change.target}: ${change.attribute} = ${change.to}`
			: `${change.target}: ${change.attribute} ${change.from} -> ${change.to}`)),
	};
}

/**
 * Renames a module. The result names the module it became: an Access form or
 * report's module takes its `Form_` or `Report_` prefix whether the new name
 * carries it or not, so `Customers` makes `Form_Customers`.
 */
export function renameModule(
	filePath: string,
	moduleName: string,
	newName: string,
): WriteResult & { moduleName: string } {
	if (isVb6ProjectPath(filePath)) {
		throw new Error(`Renaming a module of a VB6 project is not supported yet; rename ${moduleName} in the .vbp and its file.`);
	}
	// An Access form or report is an object in the database, not a module in
	// the project: renaming it moves its catalog row, its container's listing
	// and the module behind it together.
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		const renamed = accessDesignRename(design.entry.kind, newName);
		const writer = new AccessVbaWriter(hostPlatform().readFile(filePath));
		writer.renameDesign(design.entry.name, renamed.design);
		atomicContainerWrite(filePath, writer.toBuffer());
		return { ok: true, signatureDropped: false, moduleName: renamed.module };
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	assertNotDocumentModule(wb.project, moduleName,
		'its name is also the code name its document keeps, which XLIDE does not change, '
		+ 'so the code would come loose from its document. Rename it in the VBE, in the Properties window');
	assertValidModuleName(wb.container, newName);
	assertFoldedNameDistinct(wb.project.modules, wb.project.codePage, newName, moduleName);
	const oldName = wb.project.getModule(moduleName)?.name;
	const designer = oldName === undefined ? undefined : designerPath(wb.cfb, oldName);
	wb.project.renameModule(moduleName, newName);
	if (designer && wb.cfb.hasStoragePath(designer)) {
		renameDesignerStorage(wb.cfb, designer, newName, wb.project.codePage);
	}
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped, moduleName: newName };
}

/**
 * A document module - a sheet's, a workbook's, a document's code - is not
 * the project's to rename or delete. Its name is the code name the document
 * itself stores, and measured in Excel (2026-09-19): renamed here, the code
 * stayed in a module no sheet owned while the sheet got an empty one; deleted,
 * the sheet got an empty one back. The VBE refuses both as well.
 */
function assertNotDocumentModule(project: VbaProject, moduleName: string, reason: string): void {
	const module = project.getModule(moduleName);
	if (module && moduleEntry(module).type === 'document') {
		throw new Error(`${module.name} is a document module: ${reason}.`);
	}
}

/**
 * A form's designer is a storage named after the form, and its VBFrame names
 * the form again on its `Begin` line; the VBE renames both with the form
 * (measured in Excel, 2026-09-19). Left under the old name, the designer
 * belonged to no module, and Office opened the project without the form or
 * any other code module.
 */
function renameDesignerStorage(cfb: Cfb, designer: readonly string[], newName: string, codePage: number): void {
	cfb.renameStorageAtPath(designer, newName);
	const renamedPath = [...designer.slice(0, -1), newName];
	if (!cfb.hasStreamAtPath(renamedPath, VBFRAME_STREAM)) {
		return;
	}
	const frame = decodeCodePage(cfb.getStreamAtPath(renamedPath, VBFRAME_STREAM), codePage);
	const renamed = frame.replace(/^(\s*Begin\s+\{[^}]*\}\s+)\S+/m, (_, head: string) => head + newName);
	if (renamed !== frame) {
		cfb.setStreamAtPath(renamedPath, VBFRAME_STREAM, encodeCodePage(renamed, codePage));
	}
}

export function deleteModule(filePath: string, moduleName: string): WriteResult {
	if (isVb6ProjectPath(filePath)) {
		throw new Error(`Deleting a module of a VB6 project is not supported yet; remove ${moduleName} from the .vbp and delete its file.`);
	}
	const design = accessDesignFor(filePath, moduleName);
	if (design) {
		const writer = new AccessVbaWriter(hostPlatform().readFile(filePath));
		writer.deleteDesign(design.entry.name);
		atomicContainerWrite(filePath, writer.toBuffer());
		return { ok: true, signatureDropped: false };
	}
	const wb = openContainerForWrite(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	assertNotDocumentModule(wb.project, moduleName,
		'it belongs to its document, which Office gives an empty one again. To remove its code, '
		+ 'write the module without it');
	const name = wb.project.getModule(moduleName)?.name;
	const designer = name === undefined ? undefined : designerPath(wb.cfb, name);
	wb.project.deleteModule(moduleName);
	// A form's designer storage goes with it, as the VBE removes it; left
	// behind, it stopped a new form taking the name.
	if (designer && wb.cfb.hasStoragePath(designer)) {
		wb.cfb.removeStorageAtPath(designer);
	}
	saveContainer(filePath, wb);
	return { ok: true, signatureDropped };
}

/**
 * The drawing surfaces of a container, whichever host owns it.
 *
 * All three OOXML hosts keep shapes in the same package, so one reader
 * opens it and the host decides which module reads the parts. The package
 * is a fresh copy, never the cached instance readers share, because the
 * caller is about to change it.
 */
function shapeSurfaces(
	container: MacroContainer,
	filePath: string,
	what: string,
): { host: 'excel' | 'word' | 'powerpoint'; xlsx: XlsxWorkbook; zip: ZipArchive } {
	if (!container.xlsx || container.kind === 'access') {
		throw new Error(
			`${path.basename(filePath)} is ${container.description}; ${what} need an OOXML Excel workbook, `
			+ 'Word document or PowerPoint presentation.',
		);
	}
	if (container.kind === 'excel' && !container.xlsx.hasSheetSurface()) {
		throw new Error(
			`${path.basename(filePath)} is a binary Excel workbook (.xlsb); its worksheet data is `
			+ 'stored in a binary format XLIDE does not write. VBA editing is unaffected - save the '
			+ 'workbook as .xlsm to use the sheet and shape tools.',
		);
	}
	return { host: container.kind, xlsx: container.xlsx, zip: container.xlsx.zipArchive() };
}

/** A fresh copy of the workbook's package to change: never the cached instance readers share. */
function writableSheetSurface(filePath: string, what: string): XlsxWorkbook {
	const container = openMacroContainer(hostPlatform().readFile(filePath));
	if (container.kind !== 'excel' || !container.xlsx) {
		throw new Error(
			`${path.basename(filePath)} is ${container.description}; ${what} need an OOXML Excel workbook.`,
		);
	}
	if (!container.xlsx.hasSheetSurface()) {
		throw new Error(
			`${path.basename(filePath)} is a binary Excel workbook (.xlsb); its worksheet data is `
			+ 'stored in a binary format XLIDE does not write. VBA editing is unaffected - save the '
			+ 'workbook as .xlsm to use the sheet and cell tools.',
		);
	}
	return container.xlsx;
}

export function writeCells(
	filePath: string,
	sheet: string,
	startCell: string,
	data: CellValue[][],
): { ok: true } {
	const xlsx = writableSheetSurface(filePath, 'cell writes');
	xlsx.writeCells(sheet, startCell, data);
	atomicContainerWrite(filePath, xlsx.toBytes());
	return { ok: true };
}

/** One surface of a file, and the shapes on it. */
export interface ShapeSurface {
	/** The worksheet, slide or Word story the shapes are on. */
	surface: string;
	shapes: ShapeInfo[];
}

/**
 * The shapes on every surface of a file, or on the one named. A surface is
 * a worksheet in Excel, a slide in PowerPoint, and a story in Word - its
 * body, a header, a footer - since that is what each host puts shapes on.
 */
export function listShapes(filePath: string, surface?: string): { surfaces: ShapeSurface[] } {
	const container = openMacroContainer(hostPlatform().readFile(filePath));
	const drawing = shapeSurfaces(container, filePath, 'listing shapes');
	if (drawing.host === 'excel') {
		const sheets = drawing.xlsx.shapes(surface);
		return { surfaces: sheets.map((s) => ({ surface: s.sheet, shapes: s.shapes })) };
	}
	if (drawing.host === 'powerpoint') {
		const zip = drawing.zip;
		const slides = presentationSlides(zip);
		const wanted = surface === undefined ? slides : [requireSlide(slides, surface)];
		return { surfaces: wanted.map((slide) => ({ surface: slide.name, shapes: listSlideShapes(zip, slide) })) };
	}
	const zip = drawing.zip;
	const stories = documentStories(zip);
	const wanted = surface === undefined ? stories : [requireStory(stories, surface)];
	return { surfaces: wanted.map((story) => ({ surface: story.name, shapes: listStoryShapes(zip, story) })) };
}

/**
 * Add, change or remove one shape. A macro must name what Excel's Assign
 * Macro offers - a Public Sub with no required parameters, in a standard
 * module or, qualified, in a sheet's or the workbook's module - since Excel
 * finds a missing one only when someone clicks the shape.
 */
export function editShape(filePath: string, surface: string, edit: ShapeEdit): { ok: true; name: string } {
	const container = openMacroContainer(hostPlatform().readFile(filePath));
	const drawing = shapeSurfaces(container, filePath, 'shape edits');
	if (drawing.host === 'word' && edit.macro !== undefined) {
		// Ahead of the macro check below, which would otherwise refuse a Word
		// macro for the wrong reason: that no such Sub exists.
		throw new Error(WORD_SHAPE_MACRO_REFUSAL);
	}
	if (!surface && drawing.host !== 'word') {
		// Only Word has one obvious surface; the others need to be told which.
		const noun = drawing.host === 'excel' ? 'worksheet' : 'slide';
		throw new Error(`Editing a shape needs the ${noun} it is on; call the list tool for the ${noun}s in this file.`);
	}
	const checked = edit.macro ? { ...edit, macro: checkedMacro(filePath, edit.macro, drawing.host) } : edit;
	if (drawing.host === 'excel') {
		const name = drawing.xlsx.editShape(surface, checked);
		atomicContainerWrite(filePath, drawing.xlsx.toBytes());
		return { ok: true, name };
	}
	if (drawing.host === 'powerpoint') {
		const name = editSlideShape(drawing.zip, requireSlide(presentationSlides(drawing.zip), surface), checked);
		atomicContainerWrite(filePath, drawing.xlsx.toBytes());
		return { ok: true, name };
	}
	const stories = documentStories(drawing.zip);
	const story = surface ? requireStory(stories, surface) : defaultStory(stories);
	const name = editStoryShape(drawing.zip, story, checked);
	atomicContainerWrite(filePath, drawing.xlsx.toBytes());
	return { ok: true, name };
}

/** What each host calls the file a shape's macro has to live in. */
const SHAPE_MACRO_HOST: Record<'excel' | 'powerpoint', { noun: string; fileNoun: string; assign: string; scope: string }> = {
	excel: {
		noun: 'Excel',
		fileNoun: 'workbook',
		assign: "Excel's Assign Macro",
		scope: "a standard module, or in a sheet's or the workbook's module",
	},
	powerpoint: {
		noun: 'PowerPoint',
		fileNoun: 'presentation',
		assign: "PowerPoint's Action Settings",
		scope: 'a standard module',
	},
};

/**
 * A shape's macro, as Proc or Module.Proc, checked against the project.
 *
 * Both hosts that can run one find a missing macro only when someone clicks
 * the shape, so the check happens here instead: the Sub has to exist, be
 * Public, and take no required parameter, which is what each host's own
 * dialog offers.
 */
function checkedMacro(filePath: string, macro: string, host: 'excel' | 'word' | 'powerpoint'): string {
	if (host === 'word') {
		// Unreachable through editShape, which refuses a Word macro earlier
		// with the format explanation; kept so the type stays total.
		throw new Error('A Word shape cannot run a macro.');
	}
	const words = SHAPE_MACRO_HOST[host];
	// Excel writes this file as [0]! or by its name; both mean the same.
	const m = /^(?:\[0\]!|'([^']*)'!|([^'!\s]+)!)?(?:([\p{L}_][\p{L}\p{N}_]*)\.)?([\p{L}_][\p{L}\p{N}_]*)$/u.exec(macro.trim());
	if (!m) {
		throw new Error(`'${macro}' is not a macro ${words.noun} can run from a shape; give a Sub as Name or Module.Name.`);
	}
	const [, quotedBook, book, moduleName, procName] = m;
	const container = quotedBook ?? book;
	if (container !== undefined && container.toLowerCase() !== path.basename(filePath).toLowerCase()) {
		throw new Error(`'${macro}' runs a macro in ${container}; XLIDE links shapes only to Subs in this ${words.fileNoun}.`);
	}
	const modules = listModules(filePath).filter((module) => (moduleName
		? module.name.toLowerCase() === moduleName.toLowerCase()
		: module.type === 'standard'));
	if (moduleName && modules.length === 0) {
		throw new Error(`The project has no module named ${moduleName}.`);
	}
	const found: Array<{ module: string; proc: string }> = [];
	const problems: string[] = [];
	for (const module of modules) {
		if (module.type === 'class' || module.type === 'userform') {
			throw new Error(`${module.name} is a ${module.type === 'class' ? 'class' : 'UserForm'} module; a shape runs a Sub in ${words.scope}.`);
		}
		const { source } = readModule(filePath, module.name, false);
		for (const member of parseModule(source).members) {
			if (member.kind !== 'Procedure' || member.name.toLowerCase() !== procName.toLowerCase()) { continue; }
			if (member.procKind !== 'Sub') {
				problems.push(`${module.name}.${member.name} is a ${member.procKind === 'Function' ? 'Function' : 'Property'}; a shape runs a Sub.`);
			} else if (member.modifiers.some((modifier) => modifier.toLowerCase() === 'private')) {
				problems.push(`${module.name}.${member.name} is Private; ${words.assign} offers only Public Subs.`);
			} else if (member.params.some((param) => !param.optional && !param.paramArray)) {
				problems.push(`${module.name}.${member.name} takes parameters; a click passes none.`);
			} else {
				found.push({ module: module.name, proc: member.name });
			}
		}
	}
	if (found.length > 1) {
		throw new Error(`${procName} is a Public Sub in ${found.map((f) => f.module).join(' and ')}; say which, as Module.${procName}.`);
	}
	if (found.length === 0) {
		throw new Error(problems[0] ?? `The project has no Public Sub named ${procName}${moduleName ? ` in ${moduleName}` : ' in a standard module'}.`);
	}
	return moduleName ? `${found[0].module}.${found[0].proc}` : found[0].proc;
}

/**
 * Create a new macro-enabled file by copying the bundled blank template for
 * its extension byte for byte. Overwrites `filePath` if it exists - callers
 * gate that: the New File command's save dialog confirms replacement
 * natively, and the agent tool refuses existing paths outright.
 */
export function createProject(filePath: string, templatePath: string): { ok: true; path: string } {
	const template = hostPlatform().readFile(templatePath);
	atomicContainerWrite(filePath, template);
	return { ok: true, path: filePath };
}

/**
 * Adds a VBA project to a macro-enabled file that has none, starting from
 * the project in the blank template of its format (addVbaProject.ts says
 * what goes in). Office Open XML files only: a legacy file keeps its project
 * inside records XLIDE does not write, and an Access database always has one.
 */
export function addVbaProject(filePath: string, templatePath: string): { ok: true; modules: string[] } {
	const data = hostPlatform().readFile(filePath);
	if (!data.subarray(0, 2).equals(Buffer.from('PK', 'latin1'))) {
		throw new Error(
			`${path.basename(filePath)} is not an Office Open XML file, and XLIDE adds a VBA project only to those `
			+ '(.xlsm, .docm, .pptm and their templates and add-ins).',
		);
	}
	const extension = path.extname(filePath).slice(1).toLowerCase();
	const added = addVbaProjectToPackage(data, extension, hostPlatform().readFile(templatePath));
	atomicContainerWrite(filePath, added.bytes);
	return { ok: true, modules: added.modules };
}
