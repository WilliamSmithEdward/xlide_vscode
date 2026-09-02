// A VB6 project manifest, the `.vbp`: the text file that names a project's
// modules, references, components and settings. It is parsed into a typed
// view AND kept as its lines, so a manifest prints back byte for byte and a
// key this parser does not understand survives untouched. The key set here
// is the one observed across nine open-source manifests plus the shapes
// Microsoft's archived documentation names; anything else is carried, not
// interpreted.

/** The kinds of file a manifest can name. */
export type VbpModuleKind =
	| 'form'
	| 'module'
	| 'class'
	| 'usercontrol'
	| 'propertypage'
	| 'designer'
	| 'relateddoc';

/** One `Form=`, `Module=`, `Class=`, `UserControl=`, ... line. */
export interface VbpModuleRef {
	kind: VbpModuleKind;
	/** The name the manifest gives (`Module=Name; File`); forms carry none. */
	name?: string;
	/** The path as written, relative to the manifest's folder. */
	file: string;
	/** 0-based line index in the manifest. */
	line: number;
}

/** One `Reference=*\G{GUID}#major.minor#lcid#path#description` line. */
export interface VbpReference {
	guid: string;
	version: string;
	lcid: string;
	path: string;
	description: string;
	raw: string;
	line: number;
}

/** One `Object={GUID}#version#0; file.ocx` line. */
export interface VbpObject {
	guid: string;
	version: string;
	file: string;
	raw: string;
	line: number;
}

/** A manifest line as read: the raw text, and its key/value when it has one. */
export interface VbpLine {
	raw: string;
	/** The `[Section]` the line sits under, when any. */
	section?: string;
	key?: string;
	/** The value with surrounding double quotes removed. */
	value?: string;
}

export interface VbpManifest {
	/** The line ending the file uses; CRLF is what VB6 writes. */
	eol: string;
	trailingEol: boolean;
	lines: VbpLine[];
	/** `Type=`: Exe, OleDll, OleExe, or Control. */
	type?: string;
	/** `Name=`: the project name, and the default namespace of its members. */
	name?: string;
	/** `Startup=`: a form name, `Sub Main`, or `(None)`. */
	startup?: string;
	title?: string;
	exeName32?: string;
	/** `CondComp=`: conditional-compilation constants, `A = 1 : B = 0`. */
	condComp?: string;
	modules: VbpModuleRef[];
	references: VbpReference[];
	objects: VbpObject[];
}

export class VbpManifestError extends Error {}

const MODULE_KINDS: Record<string, VbpModuleKind> = {
	form: 'form',
	module: 'module',
	class: 'class',
	usercontrol: 'usercontrol',
	propertypage: 'propertypage',
	designer: 'designer',
	relateddoc: 'relateddoc',
};

/** Kinds whose line reads `Name; File`; the others name only the file. */
const NAMED_KINDS = new Set<VbpModuleKind>(['module', 'class']);

const SECTION_RE = /^\s*\[([^\]]*)\]\s*$/;
const KEY_VALUE_RE = /^([^=\[\]]+?)=(.*)$/;

function unquote(value: string): string {
	const trimmed = value.trim();
	if (trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"')) {
		return trimmed.slice(1, -1);
	}
	return trimmed;
}

function detectEol(text: string): string {
	return text.includes('\r\n') ? '\r\n' : '\n';
}

/**
 * Parses manifest text. Throws when the text is not a manifest at all: no
 * `Key=Value` line anywhere means a wrong file, and an empty module list is
 * a legitimate (if unusual) project, so it is not an error.
 */
export function parseVbpManifest(text: string): VbpManifest {
	const eol = detectEol(text);
	const trailingEol = text.endsWith('\n');
	const body = trailingEol ? text.slice(0, text.length - eol.length) : text;
	const rawLines = body.length === 0 ? [] : body.split(/\r?\n/);
	const manifest: VbpManifest = {
		eol,
		trailingEol,
		lines: [],
		modules: [],
		references: [],
		objects: [],
	};
	let section: string | undefined;
	let sawKey = false;
	rawLines.forEach((raw, index) => {
		const sectionMatch = SECTION_RE.exec(raw);
		if (sectionMatch) {
			section = sectionMatch[1];
			manifest.lines.push({ raw, section });
			return;
		}
		const match = KEY_VALUE_RE.exec(raw);
		if (!match) {
			manifest.lines.push({ raw, section });
			return;
		}
		sawKey = true;
		const key = match[1].trim();
		const value = unquote(match[2]);
		manifest.lines.push({ raw, section, key, value });
		if (section !== undefined) {
			// Sectioned keys (`[MS Transaction Server]`) are settings, never members.
			return;
		}
		const kind = MODULE_KINDS[key.toLowerCase()];
		if (kind) {
			manifest.modules.push(moduleRef(kind, value, index));
			return;
		}
		switch (key.toLowerCase()) {
			case 'type': manifest.type = value; break;
			case 'name': manifest.name = value; break;
			case 'startup': manifest.startup = value; break;
			case 'title': manifest.title = value; break;
			case 'exename32': manifest.exeName32 = value; break;
			case 'condcomp': manifest.condComp = value; break;
			case 'reference': manifest.references.push(reference(raw, match[2].trim(), index)); break;
			case 'object': manifest.objects.push(objectRef(raw, match[2].trim(), index)); break;
			default: break;
		}
	});
	if (!sawKey) {
		throw new VbpManifestError('Not a Visual Basic project file: no Key=Value line found.');
	}
	return manifest;
}

function moduleRef(kind: VbpModuleKind, value: string, line: number): VbpModuleRef {
	if (NAMED_KINDS.has(kind)) {
		const semicolon = value.indexOf(';');
		if (semicolon >= 0) {
			return {
				kind,
				name: value.slice(0, semicolon).trim(),
				file: value.slice(semicolon + 1).trim(),
				line,
			};
		}
	}
	return { kind, file: value.trim(), line };
}

/**
 * `*\G{GUID}#major.minor#lcid#path#description`. The path is the only field
 * free to contain a `#`, so the first three and the last fields are fixed and
 * whatever lies between is the path.
 */
function reference(raw: string, value: string, line: number): VbpReference {
	const stripped = value.replace(/^\*\\G/i, '');
	const parts = stripped.split('#');
	if (parts.length < 5) {
		return { guid: parts[0] ?? '', version: parts[1] ?? '', lcid: parts[2] ?? '', path: parts.slice(3, -1).join('#'), description: parts[parts.length - 1] ?? '', raw, line };
	}
	return {
		guid: parts[0],
		version: parts[1],
		lcid: parts[2],
		path: parts.slice(3, -1).join('#'),
		description: parts[parts.length - 1],
		raw,
		line,
	};
}

/** `{GUID}#version#0; file.ocx`. */
function objectRef(raw: string, value: string, line: number): VbpObject {
	const semicolon = value.indexOf(';');
	const head = semicolon >= 0 ? value.slice(0, semicolon) : value;
	const file = semicolon >= 0 ? value.slice(semicolon + 1).trim() : '';
	const parts = head.split('#');
	return { guid: parts[0] ?? '', version: parts[1] ?? '', file, raw, line };
}

/** The text of a manifest, byte for byte what was parsed when nothing changed. */
export function printVbpManifest(manifest: VbpManifest): string {
	const text = manifest.lines.map((line) => line.raw).join(manifest.eol);
	return manifest.trailingEol ? text + manifest.eol : text;
}

/** The first value of a top-level key, case-insensitively; undefined when absent. */
export function manifestValue(manifest: VbpManifest, key: string): string | undefined {
	const wanted = key.toLowerCase();
	const line = manifest.lines.find((l) => l.section === undefined && l.key?.toLowerCase() === wanted);
	return line?.value;
}
