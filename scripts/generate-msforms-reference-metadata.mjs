#!/usr/bin/env node
// Generates the MSForms member table used to type a UserForm's controls.
//
// A form's controls are members of the form's class (issue #17). Knowing their
// NAMES is enough to stop reporting them as undeclared; knowing their TYPES is
// what lets `RegionPick.` offer ComboBox members. The types come from the same
// repo-local reference dump the Excel model is built from.
//
// Deliberately narrower than the Excel generator: this emits member names for
// COMPLETION only, and nothing here feeds the hard-diagnostic surface. Promoting
// MSForms into `member-not-found` would risk a fresh false-positive class on
// form code, which is a decision to take on its own evidence rather than as a
// side effect of typing controls.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const jsonDir = path.join(root, 'reference', 'msforms', 'json');
const outputPath = path.join(root, 'src', 'analyzer', 'host', 'msformsReferenceMembers.ts');

/**
 * Members worth offering: properties and methods, not events - and not the
 * `_`-prefixed dispatch internals (`_GetGridX`, `_SetLeft`), which the VBE's
 * own completion hides.
 */
function membersOf(dump) {
    const out = [];
    for (const property of dump.properties ?? []) {
        if (property.name.startsWith('_')) { continue; }
        out.push({
            name: property.name,
            kind: 'property',
            returns: property.type || undefined,
            readOnly: property.access === 'read-only' ? true : undefined,
        });
    }
    for (const method of dump.methods ?? []) {
        if (method.name.startsWith('_')) { continue; }
        out.push({
            name: method.name,
            kind: 'method',
            returns: method.returns || undefined,
            // The call signature the dump carries verbatim, e.g.
            // `AddItem([pvargItem As Variant], [pvargIndex As Variant])`. Hover
            // and the call tip show it; without it a control method could only
            // be described as taking nothing (issue #19).
            signature: method.signature || undefined,
        });
    }
    out.sort((a, b) => a.name.localeCompare(b.name));
    return out;
}

const classDumps = [];
const interfaceDumps = new Map();
for (const fileName of fs.readdirSync(jsonDir).sort()) {
    if (!fileName.endsWith('.json')) { continue; }
    const dump = JSON.parse(fs.readFileSync(path.join(jsonDir, fileName), 'utf8'));
    if (dump.kind === 'Class') {
        classDumps.push({ dump, fileName });
    } else if (dump.kind === 'Dispatch Interface') {
        interfaceDumps.set(dump.name, { dump, fileName });
    }
}

// The library's enums (fmMatchEntry, fmAlignment, ...): their constants are
// legal bare names in any project that carries a form, and were absent from
// the model entirely, so `x = fmMatchEntryComplete` read as an undeclared
// variable (issue #41).
const constants = {};
for (const fileName of fs.readdirSync(jsonDir).sort()) {
    if (!fileName.endsWith('.json')) { continue; }
    const dump = JSON.parse(fs.readFileSync(path.join(jsonDir, fileName), 'utf8'));
    if (dump.kind !== 'Enumeration') { continue; }
    for (const c of dump.constants ?? []) {
        if (c.name.startsWith('_') || constants[c.name]) { continue; }
        constants[c.name] = {
            name: c.name,
            type: dump.name,
            value: typeof c.value === 'number' ? c.value : String(c.value ?? ''),
            source: 'external',
        };
    }
}

const types = [];
const emitted = new Set();
function emit(dump, fileName) {
    const members = membersOf(dump);
    if (members.length === 0 || emitted.has(dump.name)) { return; }
    emitted.add(dump.name);
    types.push({
        name: dump.name,
        guid: dump.guid,
        library: dump.library,
        fileName,
        members,
        // The type library's own line between a placeable control and a helper
        // type (DataObject, Page, Return*): controls source events, helpers do
        // not. Control and UserForm are the base surfaces themselves.
        isControl: (dump.events ?? []).length > 0
            && dump.name !== 'Control'
            && !dump.name.startsWith('UserForm'),
    });
    // A member's return can name a Dispatch Interface (TabStrip.SelectedItem
    // As Tab, everything's Font) - the surfaces chains land on (issue #32).
    // Pull exactly those in, closing over their own returns; the library's
    // unreferenced I*-twin and event-sink interfaces stay out.
    for (const member of members) {
        const target = member.returns && interfaceDumps.get(member.returns);
        if (target) {
            emit(target.dump, target.fileName);
        }
    }
}
for (const { dump, fileName } of classDumps) {
    emit(dump, fileName);
}
types.sort((a, b) => a.name.localeCompare(b.name));

const lines = [];
lines.push('// Generated from reference/msforms/json. Do not hand-edit member names here.');
lines.push('// Regenerate with `npm run generate:reference:msforms`.');
lines.push('//');
lines.push('// Members of the Microsoft Forms controls a UserForm can carry, used to type');
lines.push('// a form\'s implicit members (issue #17). Completion only - nothing here feeds');
lines.push('// the hard-diagnostic surface, so an unknown member on a control is not');
lines.push('// reported rather than risking a fresh false-positive class on form code.');
lines.push('');
lines.push("import type { HostConstant } from './excelObjectModel';");
lines.push('');
lines.push('export interface MsFormsMember {');
lines.push('\tname: string;');
lines.push("\tkind: 'property' | 'method';");
lines.push('\treturns?: string;');
lines.push('\treadOnly?: boolean;');
lines.push('\t/** Call signature as the dump wrote it, for methods that have one. */');
lines.push('\tsignature?: string;');
lines.push('}');
lines.push('');
lines.push('/** Control type name (without the MSForms prefix) -> its members. */');
lines.push('export const MSFORMS_REFERENCE_MEMBERS: Readonly<Record<string, readonly MsFormsMember[]>> = {');
for (const type of types) {
    lines.push(`\t${JSON.stringify(type.name)}: ${JSON.stringify(type.members)},`);
}
lines.push('};');
lines.push('');
lines.push('/**');
lines.push(' * Classes that are placeable controls, as the type library itself draws the');
lines.push(' * line: a control sources events, a helper type (DataObject, Page, Return*)');
lines.push(' * does not. Every control also carries the `Control` base surface - Left,');
lines.push(' * Top, Visible, Name, SetFocus, Move, ZOrder - which lives in Control.json');
lines.push(' * rather than being repeated in each per-type dump.');
lines.push(' */');
lines.push('export const MSFORMS_CONTROL_CLASS_NAMES: readonly string[] = [');
lines.push(`\t${types.filter((type) => type.isControl).map((type) => JSON.stringify(type.name)).join(', ')},`);
lines.push('];');
lines.push('');
lines.push('/**');
lines.push(' * Enum constants of the Microsoft Forms library, keyed by canonical name.');
lines.push(' * Legal bare names wherever the library is referenced, which is every');
lines.push(' * project carrying a UserForm (issue #41).');
lines.push(' */');
lines.push('export const MSFORMS_REFERENCE_ENUM_CONSTANTS: Readonly<Record<string, HostConstant>> = {');
for (const [name, constant] of Object.entries(constants).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`	${JSON.stringify(name)}: ${JSON.stringify(constant)},`);
}
lines.push('};');
lines.push('');
lines.push('/** Where each type came from, so a member list can be traced to its dump. */');
lines.push('export const MSFORMS_REFERENCE_PROVENANCE: Readonly<Record<string, string>> = {');
for (const type of types) {
    lines.push(`\t${JSON.stringify(type.name)}: ${JSON.stringify(`${type.library}; ${type.guid}; reference/msforms/json/${type.fileName}`)},`);
}
lines.push('};');
lines.push('');

fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
const memberCount = types.reduce((sum, type) => sum + type.members.length, 0);
console.log(`Wrote ${path.relative(root, outputPath)}: ${types.length} types, ${memberCount} members, ${Object.keys(constants).length} enum constants.`);
