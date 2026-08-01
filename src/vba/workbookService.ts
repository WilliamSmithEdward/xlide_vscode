// Native workbook service: every VBA/cell operation XLIDE needs, implemented
// directly against the OOXML package and the MS-OVBA VBA project. This replaces
// the Python bridge (pyOpenVBA + openpyxl) entirely.
//
// Writes are atomic: the updated package is written to a sibling temp file and
// renamed over the original, so a failure part-way through never leaves a
// truncated workbook.

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Cfb } from './cfb';
import {
	detectSignature,
	synthesizeClassHeader,
	synthesizeStandardHeader,
	VbaProject,
	type VbaModule,
} from './vbaProject';
import { XlsxWorkbook, type CellValue, type NamedRange, type SheetSummary } from './xlsx';

export type ModuleType = 'standard' | 'class' | 'document' | 'userform';
export type DocumentType = 'workbook' | 'worksheet' | 'chart';

export interface ModuleEntry {
	name: string;
	type: ModuleType;
	documentType?: DocumentType;
	source?: string;
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
const GUID_RE = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g;
const DOCUMENT_NAME_RE = /^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha)\d*$/i;
const ATTR_LINE_RE = /^Attribute\s+VB_/i;
const PROC_RE =
	/^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*[(\r\n]/gim;

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
		if (i < lines.length && lines[i].replace(/[\r\n]+$/, '').trim().toUpperCase() === 'BEGIN') {
			i++;
			while (i < lines.length && lines[i].replace(/[\r\n]+$/, '').trim().toUpperCase() !== 'END') {
				i++;
			}
			if (i < lines.length) { i++; }
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
		if (upper.includes(WORKBOOK_CLSID) || upper.includes(WORKSHEET_CLSID) || upper.includes(CHART_CLSID)) {
			return 'document';
		}
	}
	if (name === 'ThisWorkbook' || DOCUMENT_NAME_RE.test(name)) { return 'document'; }
	return 'standard';
}

export function classifyDocumentType(name: string, source: string): DocumentType | undefined {
	const vbBase = attributeValue(source, 'VB_Base').toUpperCase();
	if (vbBase.includes(WORKBOOK_CLSID)) { return 'workbook'; }
	if (vbBase.includes(WORKSHEET_CLSID)) { return 'worksheet'; }
	if (vbBase.includes(CHART_CLSID)) { return 'chart'; }
	if (name === 'ThisWorkbook') { return 'workbook'; }
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
		type = classifyModuleType(module.name, module.source);
		if (type === 'standard') { type = 'class'; }
	}
	const entry: ModuleEntry = { name: module.name, type };
	if (type === 'document') {
		const documentType = classifyDocumentType(module.name, module.source);
		if (documentType) { entry.documentType = documentType; }
	}
	return entry;
}

interface OpenWorkbook {
	xlsx: XlsxWorkbook;
	cfb: Cfb;
	project: VbaProject;
}

function openWorkbook(filePath: string): OpenWorkbook {
	const xlsx = XlsxWorkbook.fromBuffer(fs.readFileSync(filePath));
	const cfb = Cfb.fromBytes(xlsx.readVbaProject());
	return { xlsx, cfb, project: VbaProject.parse(cfb) };
}

/** Write the mutated VBA project back into the package, atomically. */
function saveWorkbook(filePath: string, wb: OpenWorkbook): void {
	wb.project.save(wb.cfb);
	wb.xlsx.writeVbaProject(wb.cfb.toBytes());
	atomicWrite(filePath, wb.xlsx.toBytes());
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
}

// ------------------------------------------------------------------ read API

export function listModules(filePath: string): ModuleEntry[] {
	return openWorkbook(filePath).project.modules.map(moduleEntry);
}

export function readModules(filePath: string, full = false): ModuleEntry[] {
	const { project } = openWorkbook(filePath);
	const out: ModuleEntry[] = [];
	for (const module of project.modules) {
		try {
			const entry = moduleEntry(module);
			entry.source = full ? module.source : splitVbaSource(module.source).body;
			out.push(entry);
		} catch {
			// Keep workbook-wide reads best-effort at the module boundary.
			continue;
		}
	}
	return out;
}

export function readModule(filePath: string, moduleName: string, full = false): { source: string } {
	const { project } = openWorkbook(filePath);
	const module = project.getModule(moduleName);
	if (!module) {
		throw new Error(`Module not found: ${moduleName}`);
	}
	return { source: full ? module.source : splitVbaSource(module.source).body };
}

export function listSubs(filePath: string, moduleName: string): ProcedureEntry[] {
	const { source } = readModule(filePath, moduleName, true);
	const { body } = splitVbaSource(source);
	const out: ProcedureEntry[] = [];
	PROC_RE.lastIndex = 0;
	let m: RegExpExecArray | null;
	while ((m = PROC_RE.exec(body)) !== null) {
		out.push({
			name: m[2],
			kind: m[1].replace(/\s+/g, ' ').trim(),
			line: body.slice(0, m.index).split('\n').length,
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
	return { sheets: XlsxWorkbook.fromBuffer(fs.readFileSync(filePath)).sheetSummaries() };
}

export function getWorkbookInfo(filePath: string): {
	sheets: SheetSummary[];
	namedRanges: NamedRange[];
	modules: ModuleEntry[];
	isPasswordProtected: boolean;
	isSigned: boolean;
} {
	const buffer = fs.readFileSync(filePath);
	const xlsx = XlsxWorkbook.fromBuffer(buffer);
	const cfb = Cfb.fromBytes(xlsx.readVbaProject());
	const project = VbaProject.parse(cfb);
	return {
		sheets: xlsx.sheetSummaries(),
		namedRanges: xlsx.definedNames(),
		modules: project.modules.map(moduleEntry),
		isPasswordProtected: project.hasPassword,
		isSigned: detectSignature(cfb).present,
	};
}

export function readCells(filePath: string, sheet: string, range: string): { data: CellValue[][] } {
	return { data: XlsxWorkbook.fromBuffer(fs.readFileSync(filePath)).readCells(sheet, range, true) };
}

export function readFormulas(filePath: string, sheet: string, range: string): { data: CellValue[][] } {
	return { data: XlsxWorkbook.fromBuffer(fs.readFileSync(filePath)).readCells(sheet, range, false) };
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
		if (!wb.cfb.hasStreamInStorage('VBA', streamName)) {
			issues.push(`Module ${module.name} references missing stream '${streamName}'.`);
		}
		if (module.source === '' && module.prefixBytes.length === 0) {
			issues.push(`Module ${module.name} has no readable source stream.`);
		}
	}
	return { issues };
}

// ----------------------------------------------------------------- write API

export function writeModule(
	filePath: string,
	moduleName: string,
	source: string,
	kind: 'standard' | 'class' = 'standard',
): WriteResult {
	// Callers may pass a bare body or a full export; strip any incoming header
	// so the workbook's own header is always the one that persists.
	const { body } = splitVbaSource(source);
	const wb = openWorkbook(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	const existing = wb.project.getModule(moduleName);
	if (existing) {
		const { header } = splitVbaSource(existing.source);
		wb.project.setModuleSource(existing.name, joinVbaSource(header, body));
	} else {
		const header = kind === 'class'
			? synthesizeClassHeader(moduleName)
			: synthesizeStandardHeader(moduleName);
		wb.project.addModule(moduleName, joinVbaSource(header, body), kind === 'class' ? 'other' : 'standard');
	}
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

export function renameModule(filePath: string, moduleName: string, newName: string): WriteResult {
	const wb = openWorkbook(filePath);
	const signatureDropped = detectSignature(wb.cfb).present;
	wb.project.renameModule(moduleName, newName);
	saveWorkbook(filePath, wb);
	return { ok: true, signatureDropped };
}

export function deleteModule(filePath: string, moduleName: string): WriteResult {
	const wb = openWorkbook(filePath);
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
	const xlsx = XlsxWorkbook.fromBuffer(fs.readFileSync(filePath));
	xlsx.writeCells(sheet, startCell, data);
	atomicWrite(filePath, xlsx.toBytes());
	return { ok: true };
}

/**
 * Create a new macro-enabled workbook from the bundled template (ThisWorkbook,
 * Sheet1 and an empty Module1), renaming the VBA project's declared name to
 * match. Overwrites `filePath` if it exists.
 */
export function createWorkbook(filePath: string, templatePath: string): { ok: true; path: string } {
	const template = fs.readFileSync(templatePath);
	atomicWrite(filePath, template);
	return { ok: true, path: filePath };
}

/** Temp-directory copy helper used by the test host staging path. */
export function copyWorkbookToTemp(filePath: string, prefix = 'xlide-'): string {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
	const target = path.join(dir, path.basename(filePath));
	fs.copyFileSync(filePath, target);
	return target;
}
