#!/usr/bin/env node
// Generates a complete HostObjectModel for a host from its reference dumps:
// node scripts/generate-host-object-model.mjs <host> (word, powerpoint,
// access, vb6). The Office hosts' dumps come from pyVBAReference; vb6's come
// from scripts/dump-vb6-typelib.py (VBRUN, the msvbvm60.dll type library)
// and scripts/transcribe-vb6-docs.mjs (VB, from twinBASIC's documentation),
// two libraries that each carry a `libraryId` mapped to its own namespace.
// The sixth-and-counting instance of the reference-generator pattern
// (issue #25).
//
// Deliberately mechanical where Excel's model is curated: every class becomes
// a type, every enum becomes constants, and NOTHING is exhaustive - these
// models exist to offer and to describe, never to prove a member absent. The
// hand-written wrapper (<host>ObjectModel.ts) adds the host's injected
// globals, which are the one part a type library cannot say.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
    CLASS_KINDS,
    classNamesIn,
    collapseWhitespace,
    createCurator,
    declaredType,
    descriptionIndex,
    localizeHostName,
    memberAccess,
    memberDoc,
    memberSignature,
    readDumps,
    typeDoc,
} from './reference-curation.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const host = process.argv[2];
if (!host || !/^[a-z][a-z0-9]*$/.test(host)) {
    console.error('usage: generate-host-object-model.mjs <host>   (word | powerpoint | access | vb6)');
    process.exit(1);
}
const PREFIXES = { word: 'Word', powerpoint: 'PowerPoint', access: 'Access', vb6: 'VB' };
// A host whose dumps span several libraries names each one: a dump's
// `libraryId` picks its namespace, and a library with no entry here is
// evidence only (VBA6 is dumped beside VBRUN but the analyzer's VBA runtime
// already models those names for every host).
const LIBRARY_PREFIXES = { vb6: { VB: 'VB', VBRUN: 'VBRUN' } };
const prefix = PREFIXES[host];
if (!prefix) {
    console.error(`unknown host ${host}; add it to PREFIXES deliberately.`);
    process.exit(1);
}
function prefixOf(dump) {
    if (dump.libraryId === undefined) {
        return prefix;
    }
    return LIBRARY_PREFIXES[host]?.[dump.libraryId];
}

const jsonDir = path.join(root, 'reference', host, 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', `${host}ObjectModelData.ts`);

const dumps = readDumps(jsonDir);
// Descriptions the reference cross-published from another host's page name that
// host as the actor; the index is what proves a sentence was copied.
const descriptions = descriptionIndex(path.join(root, 'reference'));

// The shared Office library's class names, so a member that returns one of its
// types stays a resolvable chain step. The host corpora do not carry those
// dumps, so without this a PowerPoint TextFrame2.TextRange (an Office
// TextRange2) dead-ends at the first hop.
// A VB6 project references no Office library, so its chains never leave its
// own two namespaces.
const foreignClasses = new Map(
    host === 'vb6'
        ? []
        : [...classNamesIn(path.join(root, 'reference', 'office', 'json'))].map((name) => [name, 'Office']),
);
const namespaces = new Map(
    [...dumps.values()]
        .filter((dump) => CLASS_KINDS.has(dump.kind) && prefixOf(dump))
        .map((dump) => [dump.name, prefixOf(dump)]),
);
const curator = createCurator({ dumps, prefix, foreignClasses, namespaces });

function memberOf(ownerName, raw, kind) {
    const member = { name: raw.name, kind };
    const { returns, returnsAnyOf } = curator.resolveReturn(ownerName, raw, kind);
    if (returns) { member.returns = returns; }
    if (returnsAnyOf) { member.returnsAnyOf = returnsAnyOf; }
    const declared = declaredType(raw, kind);
    if (declared) { member.declaredType = declared; }
    const access = memberAccess(raw, kind);
    if (access) { member.access = access; }
    const signature = memberSignature(raw, kind);
    if (signature) { member.signature = signature; }
    const doc = memberDoc(raw, 300, prefix, descriptions);
    if (doc) { member.doc = doc; }
    // A member the documentation marks reserved is real in VB6 and not run by
    // the oracle; the note travels with it so a hover says so.
    if (typeof raw.reservedNote === 'string' && raw.reservedNote) {
        member.doc = { ...(member.doc ?? { params: [], source: 'external' }), remarks: raw.reservedNote };
    }
    // The oracle's own word on the member (vb6: twinBASIC's VB package
    // source, cross-read by scripts/transcribe-vb6-docs.mjs).
    if (raw.oracle === 'implemented' || raw.oracle === 'unimplemented' || raw.oracle === 'absent') {
        member.oracle = raw.oracle;
    }
    return member;
}

const types = {};
const aliases = {};
const constants = {};
const enums = {};
let memberCount = 0;
let constantCount = 0;
let documented = 0;
let repaired = 0;
let evidenceOnly = 0;

for (const [name, dump] of dumps) {
    const libraryPrefix = prefixOf(dump);
    if (!libraryPrefix) {
        evidenceOnly += 1;
        continue;
    }
    if (CLASS_KINDS.has(dump.kind)) {
        const members = [];
        // Office hosts keep their events out of the object surfaces (the
        // analyzer's event-handler tables carry them); VB6's form and control
        // events are carried here, kind 'event', because the handler stubs a
        // VB6 form offers (Form_Load, Command1_Click ...) are derived from
        // them, and member completion filters events out by kind.
        const lists = [[dump.properties ?? [], 'property'], [dump.methods ?? [], 'method']];
        if (host === 'vb6') { lists.push([dump.events ?? [], 'event']); }
        for (const [list, kind] of lists) {
            for (const raw of list) {
                if (String(raw.name ?? '').startsWith('_')) { continue; }
                const member = memberOf(name, raw, kind);
                const declared = kind === 'property' ? raw.type : raw.returns;
                if (declared === 'Object' && member.returns && member.returns !== 'Object') { repaired += 1; }
                if (member.doc?.summary) { documented += 1; }
                members.push(member);
            }
        }
        if (members.length === 0) { continue; }
        members.sort((a, b) => a.name.localeCompare(b.name));
        memberCount += members.length;
        const qualified = `${libraryPrefix}.${name}`;
        const type = { displayName: name, members };
        if (typeof dump.source === 'string' && dump.source) { type.provenance = dump.source; }
        const doc = typeDoc(dump, 300, prefix, descriptions);
        if (doc) { type.doc = doc; }
        types[qualified] = type;
        aliases[name.toLowerCase()] = qualified;
    } else if (dump.kind === 'Module') {
        // A module's constants are plain names with no enum to belong to
        // (VBRUN's RecordsetTypeConstants); its functions are not modelled.
        for (const c of dump.constants ?? []) {
            if (String(c.name ?? '').startsWith('_') || constants[c.name]) { continue; }
            const summary = collapseWhitespace(c.description);
            constants[c.name] = {
                name: c.name,
                value: typeof c.value === 'number' ? c.value : String(c.value ?? ''),
                source: 'external',
                ...(summary ? { doc: { summary, params: [], source: 'external' } } : {}),
            };
            constantCount += 1;
        }
    } else if (dump.kind === 'Enumeration') {
        // VBA accepts the enum NAME as a declared type and as a qualifier, so
        // the model carries it alongside the members.
        if (!enums[name]) {
            const enumSummary = localizeHostName(collapseWhitespace(dump.description), prefix, descriptions);
            enums[name] = {
                displayName: name,
                ...(enumSummary ? { doc: { summary: enumSummary, params: [], source: 'external' } } : {}),
            };
        }
        for (const c of dump.constants ?? []) {
            if (String(c.name ?? '').startsWith('_') || constants[c.name]) { continue; }
            const summary = localizeHostName(collapseWhitespace(c.description), prefix, descriptions);
            constants[c.name] = {
                name: c.name,
                type: dump.name,
                value: typeof c.value === 'number' ? c.value : String(c.value ?? ''),
                source: 'external',
                ...(summary ? { doc: { summary, params: [], source: 'external' } } : {}),
            };
            constantCount += 1;
        }
    }
}

const lines = [];
lines.push(`// Generated from reference/${host}/json by generate-host-object-model.mjs.`);
lines.push('// Do not hand-edit: regenerate instead.');
lines.push('//');
if (host === 'vb6') {
    lines.push('// Types, aliases and constants of the VB6 runtime: VBRUN read from the');
    lines.push('// type library inside msvbvm60.dll (scripts/dump-vb6-typelib.py) and VB');
    lines.push('// transcribed from twinBASIC\'s documentation (scripts/transcribe-vb6-docs.mjs),');
    lines.push('// each type carrying its source as provenance. Every type is deliberately');
    lines.push('// NON-exhaustive: this metadata offers and describes, and must never prove');
    lines.push('// a member absent.');
} else {
    lines.push(`// Types, aliases and enum constants of the ${prefix} type`);
    lines.push('// library, introspected via pyVBAReference and enriched from Microsoft');
    lines.push('// Learn. Every type is deliberately NON-exhaustive: this metadata offers');
    lines.push('// and describes, and must never prove a member absent.');
}
lines.push('');
lines.push("import type { HostConstant, HostEnum, HostType } from './excelObjectModel';");
lines.push('');
lines.push('// The literals live inside a function body so V8 defers parsing and');
lines.push('// evaluating them until the host model is first requested: the extension');
lines.push('// bundle and the analysis worker both load this module at startup, and');
lines.push(`// an eagerly evaluated ${prefix} model would cost startup time and heap`);
lines.push('// in every session that never opens one of its files.');
lines.push(`export interface ${prefix}ReferenceData {`);
lines.push('\treadonly types: Readonly<Record<string, HostType>>;');
lines.push('\treadonly aliases: Readonly<Record<string, string>>;');
lines.push('\treadonly constants: Readonly<Record<string, HostConstant>>;');
lines.push('\treadonly enums: Readonly<Record<string, HostEnum>>;');
lines.push('}');
lines.push('');
lines.push(`let CACHE: ${prefix}ReferenceData | undefined;`);
lines.push('');
lines.push(`export function ${host}ReferenceData(): ${prefix}ReferenceData {`);
lines.push('\tCACHE ??= {');
lines.push('\t\ttypes: {');
for (const [qualified, type] of Object.entries(types).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t\t\t${JSON.stringify(qualified)}: ${JSON.stringify(type)},`);
}
lines.push('\t\t},');
lines.push(`\t\taliases: ${JSON.stringify(aliases)},`);
lines.push('\t\tconstants: {');
for (const [name, constant] of Object.entries(constants).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t\t\t${JSON.stringify(name)}: ${JSON.stringify(constant)},`);
}
lines.push('\t\t},');
lines.push('\t\tenums: {');
for (const [name, entry] of Object.entries(enums).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t\t\t${JSON.stringify(name)}: ${JSON.stringify(entry)},`);
}
lines.push('\t\t},');
lines.push('\t};');
lines.push('\treturn CACHE;');
lines.push('}');
lines.push('');

fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
console.log(
    `Wrote ${path.relative(root, outputPath)}: ${Object.keys(types).length} types, ${memberCount} members `
    + `(${documented} documented, ${repaired} generic returns repaired), `
    + `${constantCount} constants in ${Object.keys(enums).length} enumerations`
    + (evidenceOnly ? `; ${evidenceOnly} evidence-only dumps left out.` : '.'),
);
