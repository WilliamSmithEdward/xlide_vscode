#!/usr/bin/env node
// Extracts the member surface of twinBASIC's VB package from its own source.
//
// The twinBASIC IDE ships the VB compatibility package (App, Screen, Printer,
// Clipboard, Global, Form, MDIForm and the intrinsic controls) as a
// `.twinproj` of MIT-licensed twinBASIC source. Exported with the compiler's
// own CLI, that source is the exact surface twinBASIC implements, member by
// member, with `[Unimplemented]` on the VB6 members it declares but does not
// run and `[Hidden]` on the ones it keeps out of IntelliSense. This reads the
// export and writes reference/vb6/twinbasic-vb-surface.json, which the
// transcriber (scripts/transcribe-vb6-docs.mjs) cross-reads against the
// documentation-derived VB dumps (roadmap_vb6_support.md, Slice 4).
//
//   <ide>\bin\twinBASIC_win64.exe export "<ide>\packages\{F50B82D0-DCAB-43FE-9631-11959D4A4728}_VB\package.twinproj" "<folder>\" --overwrite
//   node scripts/twinbasic-vb-surface.mjs <folder> [--ide "twinBASIC BETA 983"]
//
// What is read: `Class`/`CoClass`/`Interface`/`Module` blocks, `Inherits`
// and `Extends` chains, class-level `Implements` (recorded, not expanded),
// public fields, properties (Get/Let/Set merged), subs, functions, events
// and constants, with the attributes stacked above each. `#If FEATURE_...`
// guards are recorded on the member and treated as active: the shipped
// package has every feature on, and no `#Else` branch is a public member.
// Procedure bodies are skipped, so nothing local is mistaken for a member.
// Names are matched case-insensitively downstream; the original spelling is
// kept.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ACCESS = /^(Public|Private|Friend|Protected)\b\s*/i;
const BLOCK_OPEN = /^(?:(Public|Private|Friend)\s+)?(Class|CoClass|Interface|Module)\s+([A-Za-z_]\w*)(?:\s*\(Of\s+[^)]*\))?(?:\s+Extends\s+([\w.]+))?/i;
const BLOCK_CLOSE = /^End\s+(Class|CoClass|Interface|Module|Enum|Type)\b/i;
const NESTED_OPEN = /^(?:(?:Public|Private|Friend|Protected)\s+)?(Enum|Type)\s+([A-Za-z_]\w*)/i;
// A member named after a keyword is written in brackets: `Public [ReadOnly]
// As Boolean` (the Data control), `Sub [Print]`.
const PROC_OPEN = /^(?:(Public|Private|Friend|Protected)\s+)?(?:Static\s+)?(Property\s+(?:Get|Let|Set)|Sub|Function)\s+\[?([A-Za-z_]\w*)\]?\s*(\(.*)?$/i;
const PROC_CLOSE = /^End\s+(Sub|Function|Property)\b/i;
const EVENT = /^(?:(Public|Private|Friend)\s+)?Event\s+\[?([A-Za-z_]\w*)\]?\s*(\(.*)?$/i;
// `Public ReadOnly ControlType As ...` is how the package declares most of
// its read-only properties; `ReadOnly` is a modifier, not the name.
// Non-greedy on the modifiers: the Data control's `Public ReadOnly ReadOnly
// As Boolean` names a property ReadOnly.
const FIELD = /^(Public|Private|Friend|Protected)\s+((?:ReadOnly\s+|WithEvents\s+|Shared\s+)*?)\[?([A-Za-z_]\w*)\]?(?:\s*\([^)]*\))?\s+As\s+(?:New\s+)?([\w.]+(?:\s*\(Of\s+[^)]*\))?)/i;
const CONSTANT = /^(?:(Public|Private|Friend|Protected)\s+)?Const\s+([A-Za-z_]\w*)/i;
const DECLARE = /^(?:(Public|Private|Friend|Protected)\s+)?Declare\s+/i;
const INHERITS = /^Inherits\s+([\w.]+)/i;
const IMPLEMENTS = /^Implements\s+([\w.]+)\s*$/i;
const DEFAULT_INTERFACE = /^Interface\s+([\w.]+)/i;
const ATTRIBUTE = /^\[([^\]]*)\]\s*/;
// The package spells the closer both `#End If` and `#EndIf`.
const CONDITIONAL = /^#(If|ElseIf|Else|End\s*If)\b(.*)$/i;

/** Splits a stacked attribute text into names, keeping the Description text. */
function parseAttributes(text, into) {
	for (const part of text.split(/,\s*(?=[A-Za-z_]\w*(?:\(|,|$))/)) {
		const m = part.match(/^([A-Za-z_]\w*)\s*(?:\((.*)\))?\s*$/s);
		if (!m) { continue; }
		const name = m[1];
		into.names.add(name);
		if (name === 'Description' && m[2] !== undefined) {
			const s = m[2].match(/^\s*"([\s\S]*)"\s*$/);
			into.description = s ? s[1].replace(/""/g, '"') : m[2];
		}
	}
	return into;
}

/** A line with its leading attributes removed, and the attributes collected. */
function stripAttributes(line, pending) {
	let rest = line;
	for (;;) {
		const m = rest.match(ATTRIBUTE);
		if (!m) { break; }
		parseAttributes(m[1], pending);
		rest = rest.slice(m[0].length);
	}
	return rest;
}

function newPending() {
	return { names: new Set(), description: undefined };
}

function declaredType(rest) {
	const m = rest.match(/\)\s*As\s+(?:New\s+)?([\w.]+)\s*$/i) ?? rest.match(/^\s*As\s+(?:New\s+)?([\w.]+)/i);
	return m ? m[1] : undefined;
}

function paramsOf(rest) {
	const open = rest.indexOf('(');
	if (open < 0) { return ''; }
	let depth = 0;
	for (let i = open; i < rest.length; i += 1) {
		if (rest[i] === '(') { depth += 1; } else if (rest[i] === ')') {
			depth -= 1;
			if (depth === 0) { return rest.slice(open + 1, i).trim(); }
		}
	}
	return rest.slice(open + 1).trim();
}

/** Strips a trailing comment that is not inside a string literal. */
function stripComment(line) {
	let inString = false;
	for (let i = 0; i < line.length; i += 1) {
		const c = line[i];
		if (c === '"') { inString = !inString; } else if (c === "'" && !inString) { return line.slice(0, i); }
	}
	return line;
}

/**
 * Parses one .twin file into its blocks. Each block: { kind, name, access,
 * extends, inherits: [], implements: [], defaultInterface, members: [] }.
 * Members: { name, kind, access, declaredType, params, unimplemented,
 * hidden, restricted, description, feature, attributes: [] }.
 */
export function parseTwinSource(text, file = '') {
	const blocks = [];
	const stack = [];           // open Class/CoClass/Interface/Module blocks
	const conditions = [];      // active #If guards (feature names)
	let nested = 0;             // depth inside Enum/Type
	let inProcedure = false;
	let pending = newPending();
	const lines = text.replace(/\r\n?/g, '\n').split('\n');
	for (let index = 0; index < lines.length; index += 1) {
		let line = stripComment(lines[index]).trim();
		if (!line) { continue; }
		const conditional = line.match(CONDITIONAL);
		if (conditional) {
			const word = conditional[1].toLowerCase();
			if (word === 'if') {
				conditions.push({ text: conditional[2].replace(/\bThen\s*$/i, '').trim(), active: true });
			} else if (word === 'elseif' || word === 'else') {
				if (conditions.length) { conditions[conditions.length - 1].active = false; }
			} else if (conditions.length) {
				conditions.pop();
			}
			continue;
		}
		if (conditions.some((c) => !c.active)) { continue; }
		if (line.startsWith('#')) { continue; }   // #Region and friends
		// Attributes may sit on their own lines above a declaration or inline.
		line = stripAttributes(line, pending);
		if (!line) { continue; }
		if (inProcedure) {
			if (PROC_CLOSE.test(line)) { inProcedure = false; }
			pending = newPending();
			continue;
		}
		if (nested > 0) {
			if (BLOCK_CLOSE.test(line)) { nested -= 1; }
			pending = newPending();
			continue;
		}
		const close = line.match(BLOCK_CLOSE);
		if (close) {
			if (stack.length) { stack.pop(); }
			pending = newPending();
			continue;
		}
		// Inside a CoClass, `[Default] Interface X` names the interface the
		// class exposes; it opens no block.
		const enclosing = stack[stack.length - 1];
		if (enclosing?.kind === 'coclass') {
			const iface = line.match(DEFAULT_INTERFACE);
			if (iface) {
				const role = pending.names.has('Source') ? 'source' : (pending.names.has('Default') ? 'default' : 'other');
				if (role === 'default' && !enclosing.defaultInterface) { enclosing.defaultInterface = iface[1]; }
				if (role === 'source') { enclosing.sourceInterface = enclosing.sourceInterface ?? iface[1]; }
			}
			pending = newPending();
			continue;
		}
		// `Public Class As String` is the OLE control's Class property, a
		// field whose name is a keyword; a block header never continues
		// with `As <type>`.
		const open = /\bAs\s+[\w.]+\s*(=.*)?$/i.test(line.replace(/\(Of\s+[^)]*\)/i, '')) ? null : line.match(BLOCK_OPEN);
		if (open) {
			const block = {
				kind: open[2].toLowerCase(),
				name: open[3],
				access: (open[1] ?? 'Public').toLowerCase(),
				extends: open[4],
				inherits: [],
				implements: [],
				defaultInterface: undefined,
				attributes: [...pending.names],
				description: pending.description,
				file,
				members: [],
			};
			blocks.push(block);
			stack.push(block);
			pending = newPending();
			continue;
		}
		const current = stack[stack.length - 1];
		if (!current) { pending = newPending(); continue; }
		if (current.kind === 'coclass') { pending = newPending(); continue; }
		if (NESTED_OPEN.test(line) && !/^(Public|Private|Friend|Protected)\s+Enum\s+\w+\s*=/.test(line)) {
			// An Enum/Type declared inside the block: skip to its End.
			const single = /\bEnd\s+(Enum|Type)\b/i.test(line);
			if (!single) { nested += 1; }
			pending = newPending();
			continue;
		}
		const inherits = line.match(INHERITS);
		if (inherits) { current.inherits.push(inherits[1]); pending = newPending(); continue; }
		const implementsAll = line.match(IMPLEMENTS);
		if (implementsAll) {
			current.implements.push({ name: implementsAll[1], forwarding: pending.names.has('WithDispatchForwarding') });
			pending = newPending();
			continue;
		}
		const feature = conditions.map((c) => c.text).join(' And ') || undefined;
		const flags = () => ({
			unimplemented: pending.names.has('Unimplemented'),
			hidden: pending.names.has('Hidden'),
			restricted: pending.names.has('Restricted'),
			description: pending.description,
			attributes: [...pending.names],
			feature,
		});
		const proc = line.match(PROC_OPEN);
		if (proc) {
			const access = (proc[1] ?? 'Public').toLowerCase();
			const head = proc[2].toLowerCase();
			const rest = proc[4] ?? '';
			const kind = head.startsWith('property') ? 'property' : 'method';
			const member = {
				name: proc[3], kind, access, params: paramsOf(rest), ...flags(),
			};
			if (head === 'property get' || head === 'function') { member.declaredType = declaredType(rest); }
			if (kind === 'property') { member.accessor = head.split(/\s+/)[1].toLowerCase(); }
			current.members.push(member);
			// Interfaces declare; classes and modules define, with a body to skip.
			if (current.kind !== 'interface' && !PROC_CLOSE.test(line.replace(/^.*?:\s*/, ''))) {
				inProcedure = !/\bEnd\s+(Sub|Function|Property)\b/i.test(line);
			}
			pending = newPending();
			continue;
		}
		const event = line.match(EVENT);
		if (event) {
			current.members.push({ name: event[2], kind: 'event', access: (event[1] ?? 'Public').toLowerCase(), params: paramsOf(event[3] ?? ''), ...flags() });
			pending = newPending();
			continue;
		}
		if (DECLARE.test(line)) { pending = newPending(); continue; }
		const constant = line.match(CONSTANT);
		if (constant) {
			current.members.push({ name: constant[2], kind: 'constant', access: (constant[1] ?? 'Public').toLowerCase(), ...flags() });
			pending = newPending();
			continue;
		}
		const field = line.match(FIELD);
		if (field) {
			current.members.push({
				name: field[3], kind: 'field', access: field[1].toLowerCase(), declaredType: field[4],
				readOnly: /\bReadOnly\b/i.test(field[2]), ...flags(),
			});
			pending = newPending();
			continue;
		}
		pending = newPending();
	}
	return blocks;
}

/**
 * The effective public surface of every block: own public members plus the
 * members of every `Inherits` base (classes) or `Extends` base
 * (interfaces), and for a CoClass the members of its default interface and
 * the events of its source interface. Property accessors merge into one
 * property with an access label. Returns a map name -> surface.
 */
export function resolveSurfaces(blocks) {
	const byName = new Map();
	for (const block of blocks) {
		if (!byName.has(block.name)) { byName.set(block.name, block); }
	}
	const cache = new Map();
	function membersOf(name, trail = []) {
		const key = name.replace(/^stdole\./i, 'stdole.');
		if (cache.has(key)) { return cache.get(key); }
		if (trail.includes(key)) { return new Map(); }
		const block = byName.get(name) ?? byName.get(name.split('.').pop());
		const out = new Map();
		if (!block) { cache.set(key, out); return out; }
		const add = (member, via) => {
			if (member.access !== 'public') { return; }
			const id = `${member.kind === 'event' ? 'event:' : ''}${member.name.toLowerCase()}`;
			const existing = out.get(id);
			if (existing) {
				if (member.kind === 'property' && existing.kind === 'property') {
					existing.accessors = [...new Set([...(existing.accessors ?? []), member.accessor])];
					existing.declaredType = existing.declaredType ?? member.declaredType;
					existing.unimplemented = existing.unimplemented && member.unimplemented;
					existing.description = existing.description || member.description;
				}
				return;
			}
			out.set(id, {
				name: member.name,
				kind: member.kind === 'field' ? 'property' : member.kind,
				// A member arriving from a base already carries its merged accessors.
				accessors: member.accessors ?? (member.kind === 'field' ? (member.readOnly ? ['get'] : ['get', 'let']) : (member.accessor ? [member.accessor] : undefined)),
				declaredType: member.declaredType,
				params: member.params,
				unimplemented: member.unimplemented,
				hidden: member.hidden,
				restricted: member.restricted,
				description: member.description,
				feature: member.feature,
				...(via ? { via } : {}),
			});
		};
		const bases = [];
		if (block.kind === 'coclass') {
			if (block.defaultInterface) { bases.push(block.defaultInterface); }
			if (block.sourceInterface) {
				for (const m of membersOf(block.sourceInterface, [...trail, key]).values()) {
					if (m.kind === 'method') { add({ ...m, kind: 'event', access: 'public' }, block.sourceInterface); }
				}
			}
		} else {
			for (const m of block.members) { add(m); }
			bases.push(...block.inherits);
			if (block.extends) { bases.push(block.extends); }
		}
		for (const base of bases) {
			if (/^stdole\.I(Unknown|Dispatch)$/i.test(base)) { continue; }
			for (const m of membersOf(base, [...trail, key]).values()) { add({ ...m, access: 'public' }, m.via ?? base); }
		}
		cache.set(key, out);
		return out;
	}
	// Interfaces a class implements, its bases' included: the public class
	// is usually a thin `Inherits` over a private base that carries them.
	function implementsOf(name, trail = []) {
		const block = byName.get(name);
		if (!block || trail.includes(name)) { return []; }
		const out = [...block.implements];
		for (const base of block.inherits) {
			for (const entry of implementsOf(base, [...trail, name])) {
				if (!out.some((e) => e.name === entry.name)) { out.push(entry); }
			}
		}
		return out;
	}
	const surfaces = new Map();
	for (const block of blocks) {
		if (surfaces.has(block.name)) { continue; }
		const members = [...membersOf(block.name).values()].map((m) => {
			const entry = { ...m };
			if (m.kind === 'property') {
				const a = new Set(m.accessors ?? []);
				entry.access = a.has('get') && (a.has('let') || a.has('set')) ? 'read/write' : (a.has('get') ? 'read-only' : 'write-only');
			}
			delete entry.accessors;
			return entry;
		});
		surfaces.set(block.name, {
			kind: block.kind,
			access: block.access,
			file: block.file,
			bases: block.kind === 'coclass' ? [block.defaultInterface, block.sourceInterface].filter(Boolean) : [...block.inherits, ...(block.extends ? [block.extends] : [])],
			implements: implementsOf(block.name),
			description: block.description,
			members,
		});
	}
	return surfaces;
}

function walk(dir, out = []) {
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name);
		if (entry.isDirectory()) { walk(full, out); } else if (entry.name.toLowerCase().endsWith('.twin')) { out.push(full); }
	}
	return out.sort();
}

export function extractSurface(exportFolder, ide) {
	const sourcesDir = path.join(exportFolder, 'Sources');
	const settingsPath = path.join(exportFolder, 'Settings');
	const settings = fs.existsSync(settingsPath) ? JSON.parse(fs.readFileSync(settingsPath, 'utf8')) : {};
	const blocks = [];
	for (const file of walk(sourcesDir)) {
		blocks.push(...parseTwinSource(fs.readFileSync(file, 'utf8'), path.relative(exportFolder, file).replace(/\\/g, '/')));
	}
	const surfaces = resolveSurfaces(blocks);
	// Every class, coclass and interface, private ones included: the Forms
	// collection is the private `_Forms` interface that VBGlobal.Forms
	// returns, and the cross-read needs it under that name.
	const classes = {};
	for (const [name, surface] of [...surfaces].sort(([a], [b]) => a.localeCompare(b))) {
		if (surface.kind === 'module') { continue; }
		classes[name] = surface;
	}
	return {
		source: {
			package: settings['project.name'] ?? 'VB',
			description: settings['project.description'] ?? '',
			version: [settings['project.versionMajor'], settings['project.versionMinor'], settings['project.versionRevision']].map((v) => v ?? 0).join('.'),
			licence: settings['project.licence'] ?? '',
			ide,
			exportedFrom: path.basename(exportFolder),
		},
		blocks: blocks.length,
		classes,
	};
}

const invokedDirectly = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (invokedDirectly) {
	const args = process.argv.slice(2);
	const folder = args.find((a) => !a.startsWith('--'));
	const ideIndex = args.indexOf('--ide');
	const ide = ideIndex >= 0 ? args[ideIndex + 1] : 'twinBASIC';
	if (!folder || !fs.existsSync(path.join(folder, 'Sources'))) {
		console.error('usage: twinbasic-vb-surface.mjs <exported VB package folder> [--ide "<label>"]');
		process.exit(1);
	}
	const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
	const outPath = path.join(root, 'reference', 'vb6', 'twinbasic-vb-surface.json');
	const surface = extractSurface(path.resolve(folder), ide);
	fs.mkdirSync(path.dirname(outPath), { recursive: true });
	fs.writeFileSync(outPath, `${JSON.stringify(surface, null, 2)}\n`, 'utf8');
	const classCount = Object.keys(surface.classes).length;
	const memberCount = Object.values(surface.classes).reduce((n, c) => n + c.members.length, 0);
	const unimplemented = Object.values(surface.classes).reduce((n, c) => n + c.members.filter((m) => m.unimplemented).length, 0);
	console.log(`Wrote ${path.relative(root, outPath)}: ${classCount} public classes, ${memberCount} members (${unimplemented} unimplemented) from ${surface.blocks} blocks; package ${surface.source.package} ${surface.source.version} (${surface.source.licence}).`);
}
