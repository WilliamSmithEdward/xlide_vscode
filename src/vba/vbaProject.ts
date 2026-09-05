// VBA project inside a vbaProject.bin CFB - [MS-OVBA].
//
// Layout: the VBA storage holds a compressed `dir` stream (project metadata +
// one record block per module) and one compressed stream per module. A module
// stream is [performance-cache prefix][compressed source]; MODULEOFFSET in the
// dir stream says where the source starts.
//
// Mutations preserve everything they do not touch: the dir stream's
// project-information/references prefix is reused verbatim and only the
// PROJECTMODULES section is regenerated, and each module stream keeps its
// original cache prefix byte-for-byte.

import { Cfb } from './cfb';
import { decodeCodePage, encodeCodePage } from './codePages';
import { compress, decompress } from './ovba';

export class VbaProjectError extends Error {}

export type VbaModuleKind = 'standard' | 'other';

/** [MS-OVBA] dir record ids used by the modules section. */
const REC_MODULENAME = 0x0019;
const REC_MODULENAME_UNICODE = 0x0047;
const REC_MODULESTREAMNAME = 0x001a;
const REC_MODULESTREAMNAME_UNICODE = 0x0032;
const REC_MODULEDOCSTRING = 0x001c;
const REC_MODULEDOCSTRING_UNICODE = 0x0048;
const REC_MODULEOFFSET = 0x0031;
const REC_MODULEHELPCONTEXT = 0x001e;
const REC_MODULECOOKIE = 0x002c;
const REC_MODULETYPE_STANDARD = 0x0021;
const REC_MODULETYPE_OTHER = 0x0022;
const REC_MODULEREADONLY = 0x0025;
const REC_MODULEPRIVATE = 0x0028;
const REC_MODULE_TERMINATOR = 0x002b;
export const REC_PROJECTMODULES = 0x000f;
const REC_PROJECTCOOKIE = 0x0013;
const REC_DIR_TERMINATOR = 0x0010;
const REC_PROJECTCODEPAGE = 0x0003;
const REC_PROJECTVERSION = 0x0009;
// MS-OVBA 2.3.4.2.1.12: the VBE "Conditional Compilation Arguments" project
// property, as `Name = Value : Name = Value`. The Unicode twin at 0x003C holds
// the same text, and is preferred because it needs no code page.
const REC_PROJECTCONSTANTS = 0x000c;
const REC_PROJECTCONSTANTS_UNICODE = 0x003c;

const SIGNATURE_STREAMS: Record<string, string> = {
	_VBA_PROJECT_SIGNATURE: 'legacy',
	_VBA_PROJECT_SIGNATURE_AGILE: 'agile',
	_VBA_PROJECT_SIGNATURE_V3: 'v3',
};

// Every ANSI string in the project - module sources, dir-stream names, the
// PROJECT stream - is encoded in PROJECTCODEPAGE, and a Russian or Japanese
// Excel writes cp1251/cp932, not cp1252. See codePages.ts; issue #6 is the
// mojibake that shipping cp1252-only produced.
const decodeAnsi = decodeCodePage;
const encodeAnsi = encodeCodePage;

export interface DirRecord {
	id: number;
	start: number;
	dataStart: number;
	dataEnd: number;
	end: number;
}

/** Tokenize a decompressed dir stream into flat records. */
export function readDirRecords(raw: Buffer): DirRecord[] {
	const out: DirRecord[] = [];
	let pos = 0;
	while (pos + 2 <= raw.length) {
		const start = pos;
		const id = raw.readUInt16LE(pos);
		pos += 2;
		if (id === REC_PROJECTVERSION) {
			// No real Size field: Reserved u32 + u32 + u16.
			if (pos + 10 > raw.length) { break; }
			pos += 10;
			out.push({ id, start, dataStart: start + 2, dataEnd: pos, end: pos });
			continue;
		}
		if (pos + 4 > raw.length) { break; }
		const size = raw.readUInt32LE(pos);
		pos += 4;
		const dataStart = pos;
		const dataEnd = Math.min(raw.length, pos + size);
		pos = dataEnd;
		out.push({ id, start, dataStart, dataEnd, end: pos });
		if (id === REC_DIR_TERMINATOR) {
			pos += 4; // trailing reserved u32
			break;
		}
	}
	return out;
}

export interface VbaModule {
	/** Logical module name (MODULENAME). */
	name: string;
	/** CFB stream name (MODULESTREAMNAME) - not always equal to `name`. */
	streamName: string;
	nameUnicode: string;
	streamNameUnicode: string;
	kind: VbaModuleKind;
	textOffset: number;
	docString: string;
	docStringUnicode: string;
	helpContext: number;
	cookie: number;
	isReadOnly: boolean;
	isPrivate: boolean;
	/**
	 * Full module source, including the hidden attribute header. Decompressed
	 * on first read (see `defineLazySource`), so callers that never look at a
	 * module's body never pay to inflate it.
	 */
	source: string;
	/**
	 * Just enough of `source` to carry the `Attribute VB_*` header, for callers
	 * classifying the module rather than reading it. Cheaper than `source` on a
	 * large module and identical to its prefix.
	 */
	readonly sourceHeader: string;
	/** Performance-cache bytes preceding the compressed source. */
	prefixBytes: Buffer;
}

/**
 * Decompressed bytes to inflate for `sourceHeader`. The attribute header sits
 * at the very start of a module, so one [MS-OVBA] chunk always covers it.
 */
const HEADER_BYTES = 4096;

export interface SignatureInfo {
	present: boolean;
	kinds: string[];
}

/**
 * Back `module.source` / `module.sourceHeader` with on-demand decompression of
 * `body` (the module stream from MODULEOFFSET onward).
 *
 * Parsing a project used to inflate every module up front, which cost more than
 * everything else in the parse combined and was wasted for the many calls that
 * only want names, types, protection state, or one module out of forty. The
 * accessors keep the eager shape - `source` still reads and assigns like a
 * plain string - so nothing downstream changes.
 */
function defineLazySource(
	module: VbaModule,
	body: Buffer,
	streamName: string,
	codePage: number,
): void {
	defineSourceAccessors(module, (maxBytes) =>
		decodeAnsi(decompress(body, `VBA/${streamName}`, maxBytes), codePage));
}

/**
 * Give a module the same source/sourceHeader pair as a parsed one, with its
 * body already in hand. Added modules take this path so `sourceHeader` tracks
 * later `source` assignments instead of freezing at the value it was born with.
 */
function defineEagerSource(module: VbaModule, initial: string): void {
	defineSourceAccessors(module, () => initial);
}

function defineSourceAccessors(module: VbaModule, inflate: (maxBytes: number) => string): void {
	let source: string | undefined;
	let header: string | undefined;

	Object.defineProperty(module, 'source', {
		configurable: true,
		enumerable: true,
		get(): string {
			if (source === undefined) { source = inflate(Infinity); }
			return source;
		},
		set(value: string): void {
			source = value;
			header = value;
		},
	});
	Object.defineProperty(module, 'sourceHeader', {
		configurable: true,
		enumerable: true,
		get(): string {
			if (source !== undefined) { return source; }
			if (header === undefined) { header = inflate(HEADER_BYTES); }
			return header;
		},
	});
}

export class VbaProject {
	/**
	 * The storage the project's streams sit in, or undefined when they sit at
	 * the root: a bare vbaProject.bin, and the synthetic file an Access
	 * database is read through, have no VBA storage to put them in.
	 */
	private streamStorage: string | undefined = 'VBA';

	codePage = 1252;
	projectCookie = 0;
	/**
	 * Raw text of the project's conditional compilation arguments, exactly as
	 * the VBE stores it. Empty when the project declares none.
	 */
	conditionalConstantsRaw = '';
	modules: VbaModule[] = [];
	hasPassword = false;

	private dirRaw: Buffer = Buffer.alloc(0);
	private dirModulesOffset = -1;
	private projectStreamRaw: Buffer | undefined;
	private readonly dirtySources = new Set<string>();
	private readonly renames = new Map<string, string>();
	private readonly added: Array<{ name: string; kind: VbaModuleKind; projectKeyword?: string }> = [];
	private readonly deleted = new Set<string>();
	private readonly removedStreams: string[] = [];
	private readonly renamedStreams: Array<[string, string]> = [];

	/**
	 * The decompressed dir stream as parsed, which carries the reference
	 * section this class otherwise passes through untouched.
	 */
	get dirStream(): Buffer {
		return this.dirRaw;
	}

	/**
	 * Adds reference records to the project's reference section, which sits
	 * immediately before the module section. Everything already declared is
	 * kept: the records go in at the end of the section, where the format
	 * puts a reference added last.
	 */
	addReferenceRecords(records: Buffer): void {
		if (this.dirModulesOffset < 0) {
			throw new VbaProjectError('dir stream has no PROJECTMODULES section; refusing to add a reference.');
		}
		this.dirRaw = Buffer.concat([
			this.dirRaw.subarray(0, this.dirModulesOffset),
			records,
			this.dirRaw.subarray(this.dirModulesOffset),
		]);
		this.dirModulesOffset += records.length;
	}

	static parse(cfb: Cfb): VbaProject {
		const project = new VbaProject();
		let dirCompressed: Buffer;
		try {
			dirCompressed = cfb.getStreamInStorage('VBA', 'dir');
		} catch {
			try {
				dirCompressed = cfb.getStream('dir');
				// Saving has to put the streams back where they were found.
				project.streamStorage = undefined;
			} catch {
				throw new VbaProjectError("No 'dir' stream found; not a valid VBA project.");
			}
		}
		const dirRaw = decompress(dirCompressed, 'VBA/dir');
		project.dirRaw = dirRaw;

		const records = readDirRecords(dirRaw);
		// Project code page must be known before decoding any ANSI record.
		for (const rec of records) {
			if (rec.id === REC_PROJECTCODEPAGE && rec.dataEnd - rec.dataStart >= 2) {
				project.codePage = dirRaw.readUInt16LE(rec.dataStart);
				break;
			}
		}
		for (const rec of records) {
			if (rec.id === REC_PROJECTMODULES) {
				project.dirModulesOffset = rec.start;
			} else if (rec.id === REC_PROJECTCOOKIE && rec.dataEnd - rec.dataStart >= 2) {
				project.projectCookie = dirRaw.readUInt16LE(rec.dataStart);
			}
		}
		// Conditional compilation arguments, Unicode first: it carries the same
		// text without needing the code page, and both records are written
		// together. Absent or empty in most projects, which is why the analyzer
		// treats a custom `#If` as undecidable unless one is declared here.
		for (const rec of records) {
			if (rec.id === REC_PROJECTCONSTANTS_UNICODE && rec.dataEnd > rec.dataStart) {
				project.conditionalConstantsRaw =
					dirRaw.subarray(rec.dataStart, rec.dataEnd).toString('utf16le');
				break;
			}
		}

		const cp = project.codePage;
		const ansi = (r: DirRecord): string => decodeAnsi(dirRaw.subarray(r.dataStart, r.dataEnd), cp);
		const uni = (r: DirRecord): string => dirRaw.subarray(r.dataStart, r.dataEnd).toString('utf16le');
		if (!project.conditionalConstantsRaw) {
			// Older writers emit only the MBCS record, which needs the code page.
			const ansiConstants = records.find(
				(r) => r.id === REC_PROJECTCONSTANTS && r.dataEnd > r.dataStart,
			);
			if (ansiConstants) {
				project.conditionalConstantsRaw = ansi(ansiConstants);
			}
		}

		let current: VbaModule | undefined;
		const flush = (): void => {
			if (current && current.name) {
				project.modules.push(current);
			}
			current = undefined;
		};
		for (const rec of records) {
			switch (rec.id) {
				case REC_MODULENAME:
					flush();
					current = {
						name: ansi(rec),
						streamName: '',
						nameUnicode: '',
						streamNameUnicode: '',
						kind: 'standard',
						textOffset: 0,
						docString: '',
						docStringUnicode: '',
						helpContext: 0,
						cookie: 0,
						isReadOnly: false,
						isPrivate: false,
						source: '',
						sourceHeader: '',
						prefixBytes: Buffer.alloc(0),
					};
					break;
				case REC_MODULENAME_UNICODE: if (current) { current.nameUnicode = uni(rec); } break;
				case REC_MODULESTREAMNAME: if (current) { current.streamName = ansi(rec); } break;
				case REC_MODULESTREAMNAME_UNICODE: if (current) { current.streamNameUnicode = uni(rec); } break;
				case REC_MODULEDOCSTRING: if (current) { current.docString = ansi(rec); } break;
				case REC_MODULEDOCSTRING_UNICODE: if (current) { current.docStringUnicode = uni(rec); } break;
				case REC_MODULEOFFSET:
					if (current && rec.dataEnd - rec.dataStart >= 4) {
						current.textOffset = dirRaw.readUInt32LE(rec.dataStart);
					}
					break;
				case REC_MODULEHELPCONTEXT:
					if (current && rec.dataEnd - rec.dataStart >= 4) {
						current.helpContext = dirRaw.readUInt32LE(rec.dataStart);
					}
					break;
				case REC_MODULECOOKIE:
					if (current && rec.dataEnd - rec.dataStart >= 2) {
						current.cookie = dirRaw.readUInt16LE(rec.dataStart);
					}
					break;
				case REC_MODULETYPE_STANDARD: if (current) { current.kind = 'standard'; } break;
				case REC_MODULETYPE_OTHER: if (current) { current.kind = 'other'; } break;
				case REC_MODULEREADONLY: if (current) { current.isReadOnly = true; } break;
				case REC_MODULEPRIVATE: if (current) { current.isPrivate = true; } break;
				case REC_MODULE_TERMINATOR: flush(); break;
				default: break;
			}
		}
		flush();

		// Locate each module's stream now - the read is cheap and the offset
		// check must still fail at parse time - but defer decompression, which
		// dominates the cost of parsing a project, until someone reads a body.
		for (const module of project.modules) {
			// The CFB directory names streams in UTF-16, so a module whose name
			// exceeds the project's ANSI code page has its real stream name only
			// in the unicode record - the ANSI record is the '?'-folded
			// projection and matches no stream. Try the unicode name first and
			// adopt whichever name resolved as the module's effective stream
			// name: every later stream operation then targets the real stream,
			// and the ANSI dir record re-serializes to the same folded
			// projection either way.
			const candidates: string[] = [];
			for (const candidate of [module.streamNameUnicode, module.streamName, module.name]) {
				if (candidate && !candidates.includes(candidate)) {
					candidates.push(candidate);
				}
			}
			let stream: Buffer | undefined;
			let resolvedName: string | undefined;
			for (const candidate of candidates) {
				try {
					stream = cfb.getStreamInStorage('VBA', candidate);
					resolvedName = candidate;
					break;
				} catch { /* try the next name */ }
			}
			if (stream === undefined) {
				for (const candidate of candidates) {
					try {
						stream = cfb.getStream(candidate);
						resolvedName = candidate;
						break;
					} catch { /* try the next name */ }
				}
			}
			if (stream === undefined || resolvedName === undefined) {
				continue;
			}
			module.streamName = resolvedName;
			if (module.textOffset > stream.length) {
				throw new VbaProjectError(
					`MODULEOFFSET ${module.textOffset} exceeds stream length ${stream.length} for module ${module.name}.`,
				);
			}
			module.prefixBytes = Buffer.from(stream.subarray(0, module.textOffset));
			defineLazySource(module, stream.subarray(module.textOffset), resolvedName, project.codePage);
		}

		// A module whose real name exceeds the ANSI code page carries it only
		// in the unicode record; the ANSI record holds the '?'-folded
		// projection. When the records show exactly that relationship, adopt
		// the unicode name as the module's outward name, so listings, lookups
		// and stream operations all speak the real name - the ANSI record
		// re-folds to the identical projection on save.
		for (const module of project.modules) {
			const uni = module.nameUnicode;
			if (uni && uni !== module.name
				&& decodeAnsi(encodeAnsi(uni, project.codePage), project.codePage) === module.name) {
				module.name = uni;
			}
		}

		try {
			project.projectStreamRaw = cfb.getStream('PROJECT');
			project.hasPassword = projectStreamHasPassword(project.projectStreamRaw, project.codePage);
		} catch {
			project.projectStreamRaw = undefined;
		}
		return project;
	}

	getModule(name: string): VbaModule | undefined {
		const needle = name.toLowerCase();
		return this.modules.find((m) => m.name.toLowerCase() === needle);
	}

	setModuleSource(name: string, source: string): void {
		const module = this.getModule(name);
		if (!module) {
			throw new VbaProjectError(`Module not found: ${name}`);
		}
		module.source = source;
		this.dirtySources.add(module.name.toLowerCase());
	}

	addModule(
		name: string,
		source: string,
		kind: VbaModuleKind,
		options: { projectKeyword?: string } = {},
	): VbaModule {
		if (this.getModule(name)) {
			throw new VbaProjectError(`Module already exists: ${name}`);
		}
		const module: VbaModule = {
			name,
			streamName: name,
			nameUnicode: name,
			streamNameUnicode: name,
			kind,
			textOffset: 0,
			docString: '',
			docStringUnicode: '',
			helpContext: 0,
			cookie: 0xffff,
			isReadOnly: false,
			isPrivate: false,
			source,
			sourceHeader: source,
			prefixBytes: Buffer.alloc(0),
		};
		defineEagerSource(module, source);
		this.modules.push(module);
		this.added.push({ name, kind, projectKeyword: options.projectKeyword });
		this.dirtySources.add(name.toLowerCase());
		return module;
	}

	renameModule(oldName: string, newName: string): void {
		const module = this.getModule(oldName);
		if (!module) {
			throw new VbaProjectError(`Module not found: ${oldName}`);
		}
		if (this.getModule(newName)) {
			throw new VbaProjectError(`Module already exists: ${newName}`);
		}
		const oldStream = module.streamName || module.name;
		this.renames.set(module.name, newName);
		// The hidden VB_Name attribute must follow the rename.
		module.source = module.source.replace(
			/^(\s*Attribute\s+VB_Name\s*=\s*")([^"]*)(")/im,
			`$1${newName}$3`,
		);
		module.name = newName;
		module.nameUnicode = newName;
		module.streamName = newName;
		module.streamNameUnicode = newName;
		this.renamedStreams.push([oldStream, newName]);
		this.dirtySources.add(newName.toLowerCase());
	}

	deleteModule(name: string): void {
		const idx = this.modules.findIndex((m) => m.name.toLowerCase() === name.toLowerCase());
		if (idx < 0) {
			throw new VbaProjectError(`Module not found: ${name}`);
		}
		const [module] = this.modules.splice(idx, 1);
		this.deleted.add(module.name);
		this.removedStreams.push(module.streamName || module.name);
		this.dirtySources.delete(module.name.toLowerCase());
	}

	/** Apply every pending change to the CFB (call cfb.toBytes() afterwards). */
	save(cfb: Cfb): void {
		for (const [oldStream, newStream] of this.renamedStreams) {
			if (oldStream !== newStream) {
				cfb.renameStreamIn(this.streamStorage, oldStream, newStream);
			}
		}
		for (const streamName of this.removedStreams) {
			try { cfb.removeStreamIn(this.streamStorage, streamName); } catch { /* already gone */ }
		}
		for (const module of this.modules) {
			if (!this.dirtySources.has(module.name.toLowerCase())) {
				continue;
			}
			const body = compress(encodeAnsi(module.source, this.codePage));
			const stream = Buffer.concat([module.prefixBytes, body]);
			cfb.setStreamIn(this.streamStorage, module.streamName || module.name, stream);
		}

		// dir: reuse the project-information/references prefix verbatim.
		if (this.dirModulesOffset < 0) {
			throw new VbaProjectError('dir stream has no PROJECTMODULES section; refusing to rewrite.');
		}
		const dir = Buffer.concat([
			this.dirRaw.subarray(0, this.dirModulesOffset),
			this.serializeModulesSection(),
		]);
		cfb.setStreamIn(this.streamStorage, 'dir', compress(dir));

		if (this.projectStreamRaw && (this.renames.size > 0 || this.added.length > 0 || this.deleted.size > 0)) {
			const updated = serializeProjectStream(this.projectStreamRaw, this.renames, {
				// A form registers as BaseClass, which is what MAKES it a form to
				// the VBE; classes and modules keep their own keywords.
				addModules: this.added.map((a) =>
					[a.name, a.projectKeyword ?? (a.kind === 'standard' ? 'Module' : 'Class')] as [string, string]),
				deleteNames: this.deleted,
				codePage: this.codePage,
			});
			cfb.writeStream('PROJECT', updated);
		}

		// The _VBA_PROJECT performance cache may reference module offsets that a
		// mutating save invalidates, so clear its body then - but never on a
		// non-mutating save, where the cache still matches the project exactly.
		const mutating = this.dirtySources.size > 0
			|| this.added.length > 0
			|| this.deleted.size > 0
			|| this.renames.size > 0;
		if (mutating) {
			invalidateVbaProjectCache(cfb, this.streamStorage);
		}
		// [MS-OVBA] writers MUST NOT emit performance-cache (__SRP_*) streams.
		// Leaving them behind hands Excel stale compiled p-code for a module set
		// that no longer matches, which it follows into a hard crash on open.
		cfb.dropStreamsIn(this.streamStorage, (name) => name.startsWith('__SRP_'));
	}

	private serializeModulesSection(): Buffer {
		const cp = this.codePage;
		const parts: Buffer[] = [];
		const rec = (id: number, data: Buffer): Buffer => {
			const head = Buffer.alloc(6);
			head.writeUInt16LE(id, 0);
			head.writeUInt32LE(data.length, 2);
			return Buffer.concat([head, data]);
		};
		const u16 = (v: number): Buffer => { const b = Buffer.alloc(2); b.writeUInt16LE(v & 0xffff, 0); return b; };
		const u32 = (v: number): Buffer => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0, 0); return b; };

		parts.push(rec(REC_PROJECTMODULES, u16(this.modules.length)));
		parts.push(rec(REC_PROJECTCOOKIE, u16(this.projectCookie)));
		for (const m of this.modules) {
			const streamName = m.streamName || m.name;
			parts.push(rec(REC_MODULENAME, encodeAnsi(m.name, cp)));
			parts.push(rec(REC_MODULENAME_UNICODE, Buffer.from(m.nameUnicode || m.name, 'utf16le')));
			parts.push(rec(REC_MODULESTREAMNAME, encodeAnsi(streamName, cp)));
			parts.push(rec(REC_MODULESTREAMNAME_UNICODE, Buffer.from(m.streamNameUnicode || streamName, 'utf16le')));
			parts.push(rec(REC_MODULEDOCSTRING, encodeAnsi(m.docString, cp)));
			parts.push(rec(REC_MODULEDOCSTRING_UNICODE, Buffer.from(m.docStringUnicode, 'utf16le')));
			parts.push(rec(REC_MODULEOFFSET, u32(m.textOffset)));
			parts.push(rec(REC_MODULEHELPCONTEXT, u32(m.helpContext)));
			parts.push(rec(REC_MODULECOOKIE, u16(m.cookie)));
			parts.push(rec(m.kind === 'standard' ? REC_MODULETYPE_STANDARD : REC_MODULETYPE_OTHER, Buffer.alloc(0)));
			if (m.isReadOnly) { parts.push(rec(REC_MODULEREADONLY, Buffer.alloc(0))); }
			if (m.isPrivate) { parts.push(rec(REC_MODULEPRIVATE, Buffer.alloc(0))); }
			parts.push(rec(REC_MODULE_TERMINATOR, Buffer.alloc(0)));
		}
		parts.push(rec(REC_DIR_TERMINATOR, Buffer.alloc(0)), Buffer.alloc(4));
		return Buffer.concat(parts);
	}
}

function projectStreamHasPassword(raw: Buffer, codePage: number): boolean {
	const text = decodeAnsi(raw, codePage);
	for (const line of text.split(/\r?\n/)) {
		const trimmed = line.trim();
		const eq = trimmed.indexOf('=');
		if (eq <= 0) { continue; }
		if (trimmed.slice(0, eq).trim() === 'DPB') {
			// A placeholder DPB decodes to ~30 hex chars; a real password pushes it past ~60.
			return trimmed.slice(eq + 1).trim().replace(/^"|"$/g, '').length >= 60;
		}
	}
	return false;
}

/**
 * Rewrite a PROJECT stream to apply renames/additions/deletions, preserving
 * every other line (ID, Name, CMG, DPB, GC, [Host Extender Info]) verbatim and
 * de-duplicating declarations (Excel treats duplicates as corruption).
 */
export function serializeProjectStream(
	raw: Buffer,
	renameMap: Map<string, string>,
	opts: { addModules?: Array<[string, string]>; deleteNames?: Set<string>; codePage?: number } = {},
): Buffer {
	const addModules = opts.addModules ?? [];
	const deleteNames = opts.deleteNames ?? new Set<string>();
	const codePage = opts.codePage ?? 1252;
	const text = decodeAnsi(raw, codePage);
	// The decoded PROJECT text can only contain names the code page stores,
	// so a unicode module name appears here as its '?'-folded projection.
	// Fold the lookup keys the same way or a rename/delete of a unicode-named
	// module would never match its own declaration line.
	const fold = (s: string): string => decodeAnsi(encodeAnsi(s, codePage), codePage);
	const renameCi = new Map<string, string>();
	for (const [k, v] of renameMap) { renameCi.set(fold(k).toLowerCase(), v); }
	const deleteCi = new Set([...deleteNames].map((n) => fold(n).toLowerCase()));

	const seenDecls = new Set<string>();
	const seenWorkspace = new Set<string>();
	const outLines: string[] = [];
	let inWorkspace = false;
	let lastDeclIdx = -1;
	let workspaceIdx = -1;
	let workspaceEndIdx = -1;

	for (const line of text.split(/\r?\n/)) {
		const stripped = line.trim();
		if (stripped.toLowerCase() === '[workspace]') {
			inWorkspace = true;
			workspaceIdx = outLines.length;
			outLines.push(line);
			continue;
		}
		if (stripped.startsWith('[') && stripped.endsWith(']')) {
			if (inWorkspace) { workspaceEndIdx = outLines.length; }
			inWorkspace = false;
			outLines.push(line);
			continue;
		}
		const eq = stripped.indexOf('=');
		if (eq < 0) {
			// Preserve blank/゙structural lines, but not the trailing empty split artifact.
			if (stripped.length > 0 || line.length > 0) { outLines.push(line); }
			continue;
		}
		const key = stripped.slice(0, eq).trim();
		const value = stripped.slice(eq + 1);
		if (inWorkspace) {
			if (deleteCi.has(key.toLowerCase())) { continue; }
			const newKey = renameCi.get(key.toLowerCase()) ?? key;
			if (seenWorkspace.has(fold(newKey).toLowerCase())) { continue; }
			seenWorkspace.add(fold(newKey).toLowerCase());
			outLines.push(`${newKey}=${value.trim()}`);
			continue;
		}
		if (key === 'Module' || key === 'Class' || key === 'BaseClass') {
			const v = value.trim();
			if (deleteCi.has(v.toLowerCase())) { continue; }
			const newVal = renameCi.get(v.toLowerCase()) ?? v;
			if (seenDecls.has(fold(newVal).toLowerCase())) { continue; }
			seenDecls.add(fold(newVal).toLowerCase());
			outLines.push(`${key}=${newVal}`);
			lastDeclIdx = outLines.length - 1;
			continue;
		}
		if (key === 'Document') {
			const v = value.trim();
			const slash = v.indexOf('/');
			const namePart = slash >= 0 ? v.slice(0, slash) : v;
			const idPart = slash >= 0 ? v.slice(slash) : '';
			if (deleteCi.has(namePart.toLowerCase())) { continue; }
			const newName = renameCi.get(namePart.toLowerCase()) ?? namePart;
			if (seenDecls.has(fold(newName).toLowerCase())) { continue; }
			seenDecls.add(fold(newName).toLowerCase());
			outLines.push(`Document=${newName}${idPart}`);
			lastDeclIdx = outLines.length - 1;
			continue;
		}
		outLines.push(line);
	}

	const freshAdds = addModules.filter(([name]) => !seenDecls.has(fold(name).toLowerCase()));
	if (freshAdds.length > 0) {
		let insertAt = lastDeclIdx >= 0 ? lastDeclIdx + 1 : outLines.length;
		if (lastDeclIdx < 0) {
			const headerIdx = outLines.findIndex((l) => l.trim().startsWith('[') && l.trim().endsWith(']'));
			insertAt = headerIdx >= 0 ? headerIdx : outLines.length;
		}
		const newDecls = freshAdds.map(([name, declKey]) => `${declKey}=${name}`);
		outLines.splice(insertAt, 0, ...newDecls);
		for (const [name] of freshAdds) { seenDecls.add(fold(name).toLowerCase()); }
		if (workspaceIdx >= insertAt) { workspaceIdx += newDecls.length; }
		if (workspaceEndIdx >= insertAt) { workspaceEndIdx += newDecls.length; }
		if (workspaceIdx >= 0) {
			const wsNew = freshAdds
				.filter(([name]) => !seenWorkspace.has(fold(name).toLowerCase()))
				.map(([name]) => `${name}=0, 0, 0, 0, C`);
			if (wsNew.length > 0) {
				const wsInsert = workspaceEndIdx > workspaceIdx ? workspaceEndIdx : outLines.length;
				outLines.splice(wsInsert, 0, ...wsNew);
			}
		}
	}

	while (outLines.length > 0 && outLines[outLines.length - 1].trim() === '') {
		outLines.pop();
	}
	return encodeAnsi(outLines.join('\r\n') + '\r\n', codePage);
}

/**
 * Zero the _VBA_PROJECT performance cache body while keeping its 5-byte header,
 * so Office regenerates p-code rather than trusting a stale cache after a
 * mutating save ([MS-OVBA] 2.3.4.1).
 */
export function invalidateVbaProjectCache(
	cfb: Cfb,
	storage: string | undefined = 'VBA',
): boolean {
	let stream: Buffer;
	try {
		stream = cfb.getStreamIn(storage, '_VBA_PROJECT');
	} catch {
		return false;
	}
	if (stream.length < 5 || stream[0] !== 0xcc || stream[1] !== 0x61 || stream[4] !== 0x00) {
		return false;
	}
	if (stream.length === 5) {
		return false;
	}
	cfb.setStreamIn(storage, '_VBA_PROJECT', Buffer.concat([
		stream.subarray(0, 5),
		Buffer.alloc(stream.length - 5),
	]));
	return true;
}

export function detectSignature(cfb: Cfb): SignatureInfo {
	const info: SignatureInfo = { present: false, kinds: [] };
	const candidates: string[] = [];
	try { candidates.push(...cfb.listStreamsInStorage('VBA')); } catch { /* no VBA storage */ }
	candidates.push(...cfb.listStreams());
	for (const name of candidates) {
		const kind = SIGNATURE_STREAMS[name];
		if (kind && !info.kinds.includes(kind)) {
			info.kinds.push(kind);
			info.present = true;
		}
	}
	return info;
}

/** Minimum attribute header VBE emits for a new standard module. */
export function synthesizeStandardHeader(name: string): string {
	return `Attribute VB_Name = "${name}"\r\n`;
}

/** Attribute header VBE emits for a new plain class module. */
export function synthesizeClassHeader(name: string): string {
	return (
		'VERSION 1.0 CLASS\r\n'
		+ 'BEGIN\r\n'
		+ '  MultiUse = -1  \'True\r\n'
		+ 'END\r\n'
		+ `Attribute VB_Name = "${name}"\r\n`
		+ 'Attribute VB_GlobalNameSpace = False\r\n'
		+ 'Attribute VB_Creatable = False\r\n'
		+ 'Attribute VB_PredeclaredId = False\r\n'
		+ 'Attribute VB_Exposed = False\r\n'
	);
}
