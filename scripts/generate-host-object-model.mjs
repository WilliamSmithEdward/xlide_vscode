#!/usr/bin/env node
// Generates a complete HostObjectModel for an Office host from its
// pyVBAReference dump: node scripts/generate-host-object-model.mjs <host>
// (word, powerpoint, access). The sixth-and-counting instance of the
// reference-generator pattern (issue #25).
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
if (!host || !/^[a-z]+$/.test(host)) {
    console.error('usage: generate-host-object-model.mjs <host>   (word | powerpoint | access)');
    process.exit(1);
}
const PREFIXES = { word: 'Word', powerpoint: 'PowerPoint', access: 'Access' };
const prefix = PREFIXES[host];
if (!prefix) {
    console.error(`unknown host ${host}; add it to PREFIXES deliberately.`);
    process.exit(1);
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
const foreignClasses = new Map(
    [...classNamesIn(path.join(root, 'reference', 'office', 'json'))].map((name) => [name, 'Office']),
);
const curator = createCurator({ dumps, prefix, foreignClasses });

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

for (const [name, dump] of dumps) {
    if (CLASS_KINDS.has(dump.kind)) {
        const members = [];
        for (const [list, kind] of [[dump.properties ?? [], 'property'], [dump.methods ?? [], 'method']]) {
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
        const qualified = `${prefix}.${name}`;
        const type = { displayName: name, members };
        const doc = typeDoc(dump, 300, prefix, descriptions);
        if (doc) { type.doc = doc; }
        types[qualified] = type;
        aliases[name.toLowerCase()] = qualified;
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
lines.push(`// Types, aliases and enum constants of the ${prefix} type`);
lines.push('// library, introspected via pyVBAReference and enriched from Microsoft');
lines.push('// Learn. Every type is deliberately NON-exhaustive: this metadata offers');
lines.push('// and describes, and must never prove a member absent.');
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
    + `${constantCount} constants in ${Object.keys(enums).length} enumerations.`,
);
