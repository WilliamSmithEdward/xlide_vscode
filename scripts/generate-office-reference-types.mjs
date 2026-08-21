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
import {
    CLASS_KINDS,
    createCurator,
    declaredType,
    memberAccess,
    memberDoc,
    memberSignature,
    readDumps,
    typeDoc,
} from './reference-curation.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const jsonDir = path.join(root, 'reference', 'office', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'officeReferenceTypes.ts');

const dumps = readDumps(jsonDir);
const curator = createCurator({ dumps, prefix: 'Office' });

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
    const doc = memberDoc(raw);
    if (doc) { member.doc = doc; }
    return member;
}

const types = {};
const aliases = {};
let memberCount = 0;
let documented = 0;
let repaired = 0;
for (const [name, dump] of dumps) {
    if (!CLASS_KINDS.has(dump.kind)) { continue; }
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
    const qualified = `Office.${name}`;
    const type = { displayName: name, members };
    const doc = typeDoc(dump);
    if (doc) { type.doc = doc; }
    types[qualified] = type;
    aliases[name.toLowerCase()] = qualified;
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
console.log(
    `Wrote ${path.relative(root, outputPath)}: ${Object.keys(types).length} Office types, `
    + `${memberCount} members (${documented} documented, ${repaired} generic returns repaired).`,
);
