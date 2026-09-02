// A VB6 project as the engine sees it: a `.vbp` manifest over loose text
// files, answering the same module questions a workbook does so the tree,
// the analyzer, and the agent tools treat it as one more container. The
// files stay the truth - there is no project stream to rebuild - so reading
// is a parse of the manifest plus the member files, and writing an existing
// module is a rewrite of its own file with its header kept.

import * as fs from 'fs';
import * as path from 'path';
import { decodeCodePage, encodeCodePage } from '../codePages';
import { splitFrmSource } from '../formDesigner';
import { atomicWrite } from '../atomicWrite';
import {
	attributeValue,
	blankDesignerHeader,
	joinVbaSource,
	listProcedures,
	splitVbaSource,
	type ProcedureEntry,
} from '../moduleSource';
import { hasAuthoritativeDesignerHeader, parseUserFormControls } from '../../vbaUserFormControls';
import { parseVbpManifest, type VbpManifest, type VbpModuleKind, type VbpModuleRef } from './vbpProject';

/**
 * VB6 saved source in the system ANSI code page; on the machines that wrote
 * the projects this engine reads that is Windows-1252. A UTF-8 byte-order
 * mark, which a later editor may have added, overrides it.
 */
export const VB6_CODE_PAGE = 1252;

/** The module kinds a project can hold. The last three are recognized and
 * listed, but their designers are opaque to the engine for now. */
export type Vb6ModuleType = 'standard' | 'class' | 'userform' | 'usercontrol' | 'propertypage' | 'designer';

export interface Vb6ModuleEntry {
	name: string;
	type: Vb6ModuleType;
	/** The absolute path of the module's own file. */
	filePath: string;
	/** The manifest line the module came from. */
	manifestKind: VbpModuleKind;
	/** Present only when the file was read; a listing never reads sources. */
	source?: string;
	implicitMembers?: { name: string; type: string }[];
	predeclaredId?: boolean;
}

export interface Vb6Project {
	vbpPath: string;
	dir: string;
	manifest: VbpManifest;
	modules: Vb6ModuleEntry[];
}

export class Vb6ProjectError extends Error {}

/** Whether a path names a VB6 project manifest, by extension. */
export function isVb6ProjectPath(filePath: string): boolean {
	return /\.vbp$/i.test(filePath);
}

const MODULE_TYPES: Record<VbpModuleKind, Vb6ModuleType | undefined> = {
	form: 'userform',
	module: 'standard',
	class: 'class',
	usercontrol: 'usercontrol',
	propertypage: 'propertypage',
	designer: 'designer',
	relateddoc: undefined,
};

/** Kinds whose file opens with a designer block before the attributes. */
const DESIGNER_KINDS = new Set<VbpModuleKind>(['form', 'usercontrol', 'propertypage', 'designer']);

interface DecodedFile {
	text: string;
	/** The file carried a UTF-8 byte-order mark; writes keep it. */
	bom: boolean;
	eol: string;
}

function readTextFile(filePath: string): DecodedFile {
	const bytes = fs.readFileSync(filePath);
	const bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
	const text = bom ? bytes.subarray(3).toString('utf8') : decodeCodePage(bytes, VB6_CODE_PAGE);
	return { text, bom, eol: text.includes('\r\n') ? '\r\n' : '\n' };
}

function encodeTextFile(text: string, shape: DecodedFile): Buffer {
	if (shape.bom) {
		return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, 'utf8')]);
	}
	return encodeCodePage(text, VB6_CODE_PAGE);
}

/**
 * A module file as its two halves: the designer block a form-like file opens
 * with (kept verbatim, never parsed here) and the module text - attributes
 * then code - which is what a workbook module's source is.
 */
function splitModuleFile(kind: VbpModuleKind, text: string): { designerBlock: string; moduleText: string } {
	if (DESIGNER_KINDS.has(kind)) {
		const split = splitFrmSource(text);
		if (split) {
			return split;
		}
	}
	return { designerBlock: '', moduleText: text };
}

interface CachedProject {
	mtimeMs: number;
	size: number;
	project: Vb6Project;
}

const projectCache = new Map<string, CachedProject>();

function projectCacheKey(vbpPath: string): string {
	return path.resolve(vbpPath).toLowerCase();
}

/**
 * The project a manifest describes. Modules are listed from the manifest
 * alone - a member file the manifest names but the disk lacks is still a
 * module, and reading it is what fails, with the file named.
 */
export function openVb6Project(vbpPath: string): Vb6Project {
	const resolved = path.resolve(vbpPath);
	let stat: fs.Stats;
	try {
		stat = fs.statSync(resolved);
	} catch (err) {
		throw new Vb6ProjectError(`Project file not found: ${resolved} (${err instanceof Error ? err.message : String(err)})`);
	}
	const key = projectCacheKey(resolved);
	const cached = projectCache.get(key);
	if (cached && cached.mtimeMs === stat.mtimeMs && cached.size === stat.size) {
		return cached.project;
	}
	const manifestFile = readTextFile(resolved);
	let manifest: VbpManifest;
	try {
		manifest = parseVbpManifest(manifestFile.text);
	} catch (err) {
		throw new Vb6ProjectError(`${path.basename(resolved)}: ${err instanceof Error ? err.message : String(err)}`);
	}
	const dir = path.dirname(resolved);
	const modules: Vb6ModuleEntry[] = [];
	for (const ref of manifest.modules) {
		const type = MODULE_TYPES[ref.kind];
		if (!type) {
			continue;
		}
		const filePath = path.resolve(dir, ref.file.replace(/\\/g, path.sep));
		modules.push({ name: moduleNameFor(ref, filePath), type, filePath, manifestKind: ref.kind });
	}
	const project: Vb6Project = { vbpPath: resolved, dir, manifest, modules };
	projectCache.set(key, { mtimeMs: stat.mtimeMs, size: stat.size, project });
	return project;
}

/**
 * A module's name is the `VB_Name` attribute its file carries - that is the
 * name the compiler uses and the name code refers to. The manifest's own
 * name (`Module=Name; File`) and the file's base name are the fallbacks when
 * the file cannot be read.
 */
function moduleNameFor(ref: VbpModuleRef, filePath: string): string {
	try {
		const file = readTextFile(filePath);
		const { moduleText } = splitModuleFile(ref.kind, file.text);
		const attribute = attributeValue(moduleText, 'VB_Name');
		if (attribute) {
			return attribute;
		}
	} catch {
		// Missing or unreadable: the manifest still names the module.
	}
	return ref.name ?? path.basename(filePath).replace(/\.[^.]+$/, '');
}

/** The listing the tree shows: names, kinds, files; no sources read beyond the name. */
export function listVb6Modules(vbpPath: string): Vb6ModuleEntry[] {
	return openVb6Project(vbpPath).modules.map((m) => ({ ...m }));
}

function findModule(project: Vb6Project, moduleName: string): Vb6ModuleEntry {
	const wanted = moduleName.toLowerCase();
	const found = project.modules.find((m) => m.name.toLowerCase() === wanted);
	if (!found) {
		throw new Vb6ProjectError(`Module not found: ${moduleName}`);
	}
	return found;
}

interface ReadModuleFile {
	entry: Vb6ModuleEntry;
	file: DecodedFile;
	designerBlock: string;
	moduleText: string;
}

function readModuleFile(project: Vb6Project, entry: Vb6ModuleEntry): ReadModuleFile {
	let file: DecodedFile;
	try {
		file = readTextFile(entry.filePath);
	} catch (err) {
		throw new Vb6ProjectError(
			`Module file not found: ${entry.filePath} (named by ${path.basename(project.vbpPath)}; ${err instanceof Error ? err.message : String(err)})`,
		);
	}
	const { designerBlock, moduleText } = splitModuleFile(entry.manifestKind, file.text);
	return { entry, file, designerBlock, moduleText };
}

/**
 * A module with its source read. The source is the FILE'S OWN TEXT with the
 * designer block blanked to whitespace, so every line number an analysis,
 * a tree row or an agent reports is a line of the file the editor shows -
 * a VB6 module has no virtual view to hide its header behind. `full`
 * answers the raw file, designer block included, for callers that want the
 * bytes as written. A form's controls are supplied only when its header can
 * be read authoritatively, and absent otherwise - "not known", never "none".
 */
export function readVb6Module(vbpPath: string, moduleName: string, full = false): Vb6ModuleEntry {
	const project = openVb6Project(vbpPath);
	const read = readModuleFile(project, findModule(project, moduleName));
	return moduleWithSource(read, full);
}

function moduleWithSource(read: ReadModuleFile, full: boolean): Vb6ModuleEntry {
	const { entry, moduleText, designerBlock, file } = read;
	const out: Vb6ModuleEntry = {
		...entry,
		source: full ? file.text : blankDesignerHeader(file.text),
	};
	const predeclared = attributeValue(moduleText, 'VB_PredeclaredId');
	if (predeclared) {
		out.predeclaredId = /^True$/i.test(predeclared);
	}
	if (entry.type === 'userform') {
		const header = designerBlock + moduleText;
		if (hasAuthoritativeDesignerHeader(header)) {
			out.implicitMembers = parseUserFormControls(header).map((c) => ({ name: c.name, type: c.type }));
		}
	}
	return out;
}

/** Every module with its source, skipping files that cannot be read. */
export function readVb6Modules(vbpPath: string, full = false): Vb6ModuleEntry[] {
	const project = openVb6Project(vbpPath);
	const out: Vb6ModuleEntry[] = [];
	for (const entry of project.modules) {
		try {
			out.push(moduleWithSource(readModuleFile(project, entry), full));
		} catch {
			continue;
		}
	}
	return out;
}

/** The procedures of one module, at the lines they hold in the file. */
export function listVb6Procedures(vbpPath: string, moduleName: string): ProcedureEntry[] {
	const project = openVb6Project(vbpPath);
	const entry = findModule(project, moduleName);
	const { file } = readModuleFile(project, entry);
	return listProcedures(blankDesignerHeader(file.text));
}

/**
 * Rewrites an existing module's body. The file's designer block and
 * attribute header stay exactly as they were, the body arrives with the
 * file's own line endings, and the bytes go back in the encoding they came
 * in. Adding a module means editing the manifest as well and is not offered
 * here yet.
 */
export function writeVb6Module(vbpPath: string, moduleName: string, source: string): void {
	const project = openVb6Project(vbpPath);
	const wanted = moduleName.toLowerCase();
	if (!project.modules.some((m) => m.name.toLowerCase() === wanted)) {
		throw new Vb6ProjectError(
			'Adding a module to a VB6 project is not supported yet: the manifest would need an entry. '
			+ 'Add ' + moduleName + ' to ' + path.basename(project.vbpPath) + ' and create its file.',
		);
	}
	const read = readModuleFile(project, findModule(project, moduleName));
	// A caller may hand back what it read - blanked header lines, then the
	// attributes, then the code - or a bare body; either way only the body
	// is new, and the file's own header is the one that persists.
	const { body } = splitVbaSource(source.replace(/^\s+/, ''));
	const { header } = splitVbaSource(read.moduleText);
	const normalizedBody = body.replace(/\r\n|\r|\n/g, read.file.eol);
	const text = read.designerBlock + joinVbaSource(header, normalizedBody);
	atomicWrite(read.entry.filePath, encodeTextFile(text, read.file));
}

/** What a project validation can say without a compiler: names and files. */
export function validateVb6Project(vbpPath: string): { issues: string[] } {
	let project: Vb6Project;
	try {
		project = openVb6Project(vbpPath);
	} catch (err) {
		return { issues: [err instanceof Error ? err.message : String(err)] };
	}
	const issues: string[] = [];
	const seen = new Set<string>();
	for (const entry of project.modules) {
		const key = entry.name.toLowerCase();
		if (seen.has(key)) {
			issues.push(`Duplicate module name: ${entry.name}`);
		}
		seen.add(key);
		if (!fs.existsSync(entry.filePath)) {
			issues.push(`Missing module file: ${path.relative(project.dir, entry.filePath)} (${entry.name})`);
		}
	}
	if (project.manifest.startup && !/^\(none\)$/i.test(project.manifest.startup)
		&& !/^sub main$/i.test(project.manifest.startup)
		&& !project.modules.some((m) => m.name.toLowerCase() === project.manifest.startup!.toLowerCase())) {
		issues.push(`Startup object not found: ${project.manifest.startup}`);
	}
	return { issues };
}

/** Test hook; product code never calls it. */
export function resetVb6ProjectCacheForTests(): void {
	projectCache.clear();
}
