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

/** Members worth offering: properties and methods, not events. */
function membersOf(dump) {
    const out = [];
    for (const property of dump.properties ?? []) {
        out.push({
            name: property.name,
            kind: 'property',
            returns: property.type || undefined,
            readOnly: property.access === 'read-only' ? true : undefined,
        });
    }
    for (const method of dump.methods ?? []) {
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

const types = [];
for (const fileName of fs.readdirSync(jsonDir).sort()) {
    if (!fileName.endsWith('.json')) { continue; }
    const dump = JSON.parse(fs.readFileSync(path.join(jsonDir, fileName), 'utf8'));
    if (dump.kind !== 'Class') { continue; }
    const members = membersOf(dump);
    if (members.length === 0) { continue; }
    types.push({
        name: dump.name,
        guid: dump.guid,
        library: dump.library,
        fileName,
        members,
    });
}

const lines = [];
lines.push('// Generated from reference/msforms/json. Do not hand-edit member names here.');
lines.push('// Regenerate with `npm run generate:reference:msforms`.');
lines.push('//');
lines.push('// Members of the Microsoft Forms controls a UserForm can carry, used to type');
lines.push('// a form\'s implicit members (issue #17). Completion only - nothing here feeds');
lines.push('// the hard-diagnostic surface, so an unknown member on a control is not');
lines.push('// reported rather than risking a fresh false-positive class on form code.');
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
lines.push('/** Where each type came from, so a member list can be traced to its dump. */');
lines.push('export const MSFORMS_REFERENCE_PROVENANCE: Readonly<Record<string, string>> = {');
for (const type of types) {
    lines.push(`\t${JSON.stringify(type.name)}: ${JSON.stringify(`${type.library}; ${type.guid}; reference/msforms/json/${type.fileName}`)},`);
}
lines.push('};');
lines.push('');

fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
const memberCount = types.reduce((sum, type) => sum + type.members.length, 0);
console.log(`Wrote ${path.relative(root, outputPath)}: ${types.length} types, ${memberCount} members.`);
