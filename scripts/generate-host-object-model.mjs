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
const MAX_SUMMARY = 300;

const index = JSON.parse(fs.readFileSync(path.join(jsonDir, '_index.json'), 'utf8'));
const dumps = [];
for (const t of index.types ?? []) {
    const file = path.join(jsonDir, `${t.name.replace(/[^A-Za-z0-9_.-]/g, '_')}.json`);
    try {
        dumps.push(JSON.parse(fs.readFileSync(file, 'utf8')));
    } catch {
        // A type listed but not written stays out of the model.
    }
}

// The dump vocabulary spans 'Class' (coclasses), 'Dispatch Interface' and
// 'Interface' (the member-bearing shapes), and 'Enumeration'. Events-only
// interfaces fall out naturally: no properties or methods, no type emitted.
const CLASS_KINDS = new Set(['Class', 'Dispatch Interface', 'Interface']);
const classNames = new Set(dumps.filter((d) => CLASS_KINDS.has(d.kind)).map((d) => d.name));

/** Same-library class types qualify (`Document` -> `Word.Document`); the rest pass through. */
function qualifyReturn(type) {
    if (!type) { return undefined; }
    const bare = String(type).replace(/^\s+|\s+$/g, '');
    return classNames.has(bare) ? `${prefix}.${bare}` : bare;
}

function summarize(text) {
    if (!text) { return undefined; }
    const line = String(text).replace(/\s+/g, ' ').trim();
    if (!line) { return undefined; }
    return line.length > MAX_SUMMARY ? `${line.slice(0, MAX_SUMMARY - 1)}…` : line;
}

function memberOf(item, kind) {
    const member = { name: item.name, kind };
    const returns = qualifyReturn(kind === 'property' ? item.type : item.returns);
    if (returns && returns !== 'void') { member.returns = returns; }
    if (kind === 'method' && item.signature) { member.signature = item.signature; }
    const summary = summarize(item.description);
    if (summary) { member.doc = { summary, params: [], source: 'external' }; }
    return member;
}

const types = {};
const aliases = {};
const constants = {};
let memberCount = 0;
let constantCount = 0;

for (const dump of dumps) {
    if (CLASS_KINDS.has(dump.kind)) {
        const members = [];
        for (const p of dump.properties ?? []) {
            if (p.name.startsWith('_')) { continue; }
            members.push(memberOf(p, 'property'));
        }
        for (const m of dump.methods ?? []) {
            if (m.name.startsWith('_')) { continue; }
            members.push(memberOf(m, 'method'));
        }
        if (members.length === 0) { continue; }
        members.sort((a, b) => a.name.localeCompare(b.name));
        memberCount += members.length;
        const qualified = `${prefix}.${dump.name}`;
        const type = { displayName: dump.name, members };
        const summary = summarize(dump.description);
        if (summary) { type.provenance = summary; }
        types[qualified] = type;
        aliases[dump.name.toLowerCase()] = qualified;
    } else if (dump.kind === 'Enumeration') {
        for (const c of dump.constants ?? []) {
            if (c.name.startsWith('_') || constants[c.name]) { continue; }
            constants[c.name] = {
                name: c.name,
                type: dump.name,
                value: typeof c.value === 'number' ? c.value : String(c.value ?? ''),
                source: 'external',
            };
            constantCount += 1;
        }
    }
}

const lines = [];
lines.push(`// Generated from reference/${host}/json by generate-host-object-model.mjs.`);
lines.push('// Do not hand-edit: regenerate instead.');
lines.push('//');
lines.push(`// Types, aliases and enum constants of the ${index.library ?? prefix} type`);
lines.push('// library, introspected via pyVBAReference and enriched from Microsoft');
lines.push('// Learn. Every type is deliberately NON-exhaustive: this metadata offers');
lines.push('// and describes, and must never prove a member absent.');
lines.push('');
lines.push("import type { HostConstant, HostType } from './excelObjectModel';");
lines.push('');
lines.push(`export const ${host.toUpperCase()}_REFERENCE_TYPES: Readonly<Record<string, HostType>> = {`);
for (const [qualified, type] of Object.entries(types).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t${JSON.stringify(qualified)}: ${JSON.stringify(type)},`);
}
lines.push('};');
lines.push('');
lines.push(`export const ${host.toUpperCase()}_REFERENCE_ALIASES: Readonly<Record<string, string>> = ${JSON.stringify(aliases)};`);
lines.push('');
lines.push(`export const ${host.toUpperCase()}_REFERENCE_CONSTANTS: Readonly<Record<string, HostConstant>> = {`);
for (const [name, constant] of Object.entries(constants).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t${JSON.stringify(name)}: ${JSON.stringify(constant)},`);
}
lines.push('};');
lines.push('');

fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
console.log(`Wrote ${path.relative(root, outputPath)}: ${Object.keys(types).length} types, ${memberCount} members, ${constantCount} constants.`);
