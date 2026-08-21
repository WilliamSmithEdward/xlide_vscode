#!/usr/bin/env node
// Generates the shared Office library's TYPE metadata:
// src/analyzer/host/officeReferenceTypes.ts
//
// Every Office VBA project references the Office library, and its types are
// reachable from the host libraries: PowerPoint's TextFrame2.TextRange returns
// an Office TextRange2, Word's and Excel's shape formatting reach ColorFormat,
// ThreeDFormat and friends the same way. The host corpora do not carry those
// dumps, so a chain that lands on one used to dead-end and `As TextRange2`
// resolved to nothing (the shape of issue #32, one library over).
//
// Constants are generated separately by generate-office-reference-constants;
// this is the object surface only. Nothing here is exhaustive: it offers and
// describes, and must never prove a member absent.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const jsonDir = path.join(root, 'reference', 'office', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'officeReferenceTypes.ts');
const MAX_SUMMARY = 300;

const CLASS_KINDS = new Set(['Class', 'Dispatch Interface']);

function summarize(text) {
    if (!text) { return undefined; }
    const line = String(text).replace(/\s+/g, ' ').trim();
    if (!line) { return undefined; }
    return line.length > MAX_SUMMARY ? `${line.slice(0, MAX_SUMMARY - 1)}\u2026` : line;
}

const dumps = [];
for (const fileName of fs.readdirSync(jsonDir).sort()) {
    if (!fileName.endsWith('.json') || fileName.startsWith('_')) { continue; }
    try {
        dumps.push(JSON.parse(fs.readFileSync(path.join(jsonDir, fileName), 'utf8')));
    } catch {
        // A malformed dump stays out of the model rather than breaking it.
    }
}

const classNames = new Set(dumps.filter((d) => CLASS_KINDS.has(d.kind)).map((d) => d.name));

/** An Office type reference stays qualified; anything else is left bare. */
function qualifyReturn(type) {
    if (!type) { return undefined; }
    const bare = String(type).trim();
    return classNames.has(bare) ? `Office.${bare}` : bare;
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
let memberCount = 0;
for (const dump of dumps) {
    if (!CLASS_KINDS.has(dump.kind)) { continue; }
    const members = [];
    for (const property of dump.properties ?? []) {
        if (property.name.startsWith('_')) { continue; }
        members.push(memberOf(property, 'property'));
    }
    for (const method of dump.methods ?? []) {
        if (method.name.startsWith('_')) { continue; }
        members.push(memberOf(method, 'method'));
    }
    if (members.length === 0) { continue; }
    members.sort((a, b) => a.name.localeCompare(b.name));
    memberCount += members.length;
    const qualified = `Office.${dump.name}`;
    const type = { displayName: dump.name, members };
    const summary = summarize(dump.description);
    if (summary) { type.provenance = summary; }
    types[qualified] = type;
    aliases[dump.name.toLowerCase()] = qualified;
}

const lines = [];
lines.push('// Generated from reference/office/json by generate-office-reference-types.mjs.');
lines.push('// Do not hand-edit: regenerate instead.');
lines.push('//');
lines.push('// Types of the shared Microsoft Office object library, which every Office VBA');
lines.push('// project references. Host libraries return these from their own members -');
lines.push("// PowerPoint's TextFrame2.TextRange is an Office TextRange2 - so a host model");
lines.push('// without them dead-ends the chain there. Merged into every host model, with');
lines.push("// the host's own types winning any shared name. Deliberately NON-exhaustive:");
lines.push('// this metadata offers and describes, and must never prove a member absent.');
lines.push('');
lines.push("import type { HostType } from './excelObjectModel';");
lines.push('');
lines.push('// The literals live inside a function body so V8 defers parsing and');
lines.push('// evaluating them until a host model is first requested.');
lines.push('interface OfficeReferenceData {');
lines.push('\treadonly types: Readonly<Record<string, HostType>>;');
lines.push('\treadonly aliases: Readonly<Record<string, string>>;');
lines.push('}');
lines.push('');
lines.push('let CACHE: OfficeReferenceData | undefined;');
lines.push('');
lines.push('export function officeReferenceTypeData(): OfficeReferenceData {');
lines.push('\tCACHE ??= {');
lines.push('\t\ttypes: {');
for (const [name, type] of Object.entries(types).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t\t\t${JSON.stringify(name)}: ${JSON.stringify(type)},`);
}
lines.push('\t\t},');
lines.push('\t\taliases: {');
for (const [name, qualified] of Object.entries(aliases).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`\t\t\t${JSON.stringify(name)}: ${JSON.stringify(qualified)},`);
}
lines.push('\t\t},');
lines.push('\t};');
lines.push('\treturn CACHE;');
lines.push('}');
lines.push('');

fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
console.log(`Wrote ${path.relative(root, outputPath)}: ${Object.keys(types).length} Office types, ${memberCount} members.`);
