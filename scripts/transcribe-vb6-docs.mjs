#!/usr/bin/env node
// Transcribes the `VB` library - App, Screen, Printer, Clipboard, Global,
// Form and the intrinsic controls - from twinBASIC's package documentation
// into reference/vb6/json, the shape scripts/generate-host-object-model.mjs
// reads. VB6's own VB.OLB is not available to introspect (it ships only with
// the VB6 IDE), so the documentation is the source, and every dump says so.
//
//   node scripts/transcribe-vb6-docs.mjs <path-to-twinbasic-documentation-checkout>
//
// The checkout is github.com/twinbasic/documentation; the pages live under
// docs/Reference/Default/VB/<Class>/index.md. Filtered to VB6's real surface:
// twinBASIC's own controls (CheckMark, MultiFrame, QRCode, Report) are
// skipped, and a member the page marks "New in twinBASIC" (or a twinBASIC
// extension) is left out and listed under `excludedMembers` for the record.
// A member the page marks reserved for VB6 compatibility - real in VB6, not
// implemented in twinBASIC - is kept and flagged `reserved`, so the oracle
// harness knows twinBASIC cannot vouch for it.
//
// Idempotent: re-running rewrites the same files.

import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const checkout = process.argv[2];
if (!checkout || !fs.existsSync(path.join(checkout, 'docs', 'Reference', 'Default', 'VB'))) {
    console.error('usage: transcribe-vb6-docs.mjs <twinbasic/documentation checkout>');
    process.exit(1);
}
const pagesDir = path.join(checkout, 'docs', 'Reference', 'Default', 'VB');
const outDir = path.join(root, 'reference', 'vb6', 'json');
fs.mkdirSync(outDir, { recursive: true });

const TWINBASIC_ONLY_CLASSES = new Set(['CheckMark', 'MultiFrame', 'QRCode', 'Report']);
// A member-level marker is a sentence: "New in twinBASIC." after the
// description, or an explicit "twinBASIC-specific" / "no VB6 equivalent". A
// parenthetical "(new in twinBASIC)" beside one constant value marks that
// value, not the member (Form.BorderStyle lists two such constants).
const TWINBASIC_ONLY = /(^|[.!?]\s+)New in twinBASIC|twinBASIC-specific|twinBASIC extension|VB6 had no equivalent|no equivalent VB6|no VB6 equivalent/;
const RESERVED = /\b(Reserved|Declared) for (compatibility with )?VB6/i;
const SCALARS = new Set([
    'Boolean', 'Integer', 'Long', 'Single', 'Double', 'String', 'Variant', 'Currency', 'Date', 'Byte',
    'Object', 'OLE_COLOR', 'StdPicture', 'StdFont', 'IPictureDisp', 'IFontDisp', 'Control', 'Form',
]);

let commit = 'unknown';
try {
    commit = execSync('git rev-parse HEAD', { cwd: checkout, encoding: 'utf8' }).trim();
} catch {
    // The checkout may be a plain copy; the dump then names no commit.
}

// The name-level cross-read against Microsoft's archived VB6 Language
// Reference (scripts/fetch-vb6-reference-names.mjs): a member whose name has
// no page of its kind anywhere in that reference is not VB6, whatever
// twinBASIC's page says, and is left out on the record.
const referencePath = path.join(root, 'reference', 'vb6', 'vb6-reference-names.json');
const reference = fs.existsSync(referencePath) ? JSON.parse(fs.readFileSync(referencePath, 'utf8')) : undefined;
const referenceNames = reference
    ? Object.fromEntries(Object.entries(reference.names).map(([k, v]) => [k, new Set(v.map((n) => n.toLowerCase()))]))
    : undefined;
// A Global "method" is a VB6 statement (Load, Unload) or function
// (LoadPicture); the reference files them under those kinds. Events are not
// cross-read: the archived reference is missing event pages VB6 certainly
// had (Unload, the OLE drag-and-drop events, a scroll bar's Scroll), so
// their absence proves nothing, and events rest on twinBASIC's pages and
// their own markers.
const REFERENCE_KINDS = { property: ['properties'], method: ['methods', 'functions', 'statements'] };

// Real VB6 members whose pages the archive lost; each was checked against
// the VB6 IDE's own Object Browser vocabulary before being listed.
const KNOWN_VB6_MISSING_FROM_ARCHIVE = new Set(['DataMemberChanged']);

function inVb6Reference(kind, name) {
    if (!referenceNames || !REFERENCE_KINDS[kind] || KNOWN_VB6_MISSING_FROM_ARCHIVE.has(name)) { return true; }
    return REFERENCE_KINDS[kind].some((set) => referenceNames[set]?.has(name.toLowerCase()));
}

/** Markdown to plain text: bold, links, code, the docs' triple-dash. */
function plain(text) {
    return text
        .replace(/\[\*\*([^*\]]+)\*\*\]\([^)]*\)/g, '$1')
        .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
        .replace(/\*\*([^*]+)\*\*/g, '$1')
        .replace(/\*([^*]+)\*/g, '$1')
        .replace(/`([^`]*)`/g, '$1')
        .replace(/\s*---\s*/g, ' - ')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * The declared type a body states: "member of [**XxxConstants**]", the first
 * `**Scalar**`, the assigned kind in a `Syntax: object.Name [ = *string* ]`
 * line, or a link to another VB page - except for a collection member such
 * as Global.Forms, whose link names the element type, not its own.
 */
function declaredTypeOf(body, memberName = '') {
    const enumRef = body.match(/member of \[\*\*(\w+)\*\*\]/);
    if (enumRef) { return enumRef[1]; }
    for (const m of body.matchAll(/\*\*([A-Za-z_][\w.]*)\*\*/g)) {
        if (SCALARS.has(m[1])) { return m[1]; }
    }
    const assigned = body.match(/\[ = \*(string|boolean|number|color|value)\* \]/);
    if (assigned) {
        return { string: 'String', boolean: 'Boolean', color: 'OLE_COLOR' }[assigned[1]] ?? '';
    }
    const classRef = body.match(/\[\*\*(\w+)\*\*\]\(\.\.\/(\w+)\/?\)/);
    if (classRef && classRef[1] === classRef[2] && `${classRef[1]}s` !== memberName) { return classRef[1]; }
    return '';
}

function accessOf(body) {
    if (/write-only/i.test(body)) { return 'write-only'; }
    if (/read-only/i.test(body)) { return 'read-only'; }
    if (/readable and (writable|assignable)|read\/write/i.test(body)) { return 'read/write'; }
    return '';
}

/** Definition-list parameters: `*Name*` (or `*X*, *Y*`) then `: *optional* text`. */
function definitionParams(lines) {
    const params = [];
    for (let i = 0; i < lines.length - 1; i += 1) {
        const head = lines[i].match(/^\*(\w+)\*(?:,\s*\*(\w+)\*)*\s*$/);
        const def = lines[i + 1].match(/^:\s*(?:\*(optional|required)\*\s*)?(.*)$/);
        if (!head || !def) { continue; }
        const names = [...lines[i].matchAll(/\*(\w+)\*/g)].map((m) => m[1]);
        for (const name of names) {
            params.push({
                name,
                type: declaredTypeOf(def[2]),
                optional: def[1] === 'optional',
                description: plain(def[2]),
            });
        }
    }
    return params;
}

/** Event syntax: `*object*\_**Click**( *Source* **As Control**, *X* **As Single** )`. */
function syntaxParams(syntax) {
    const params = [];
    for (const m of syntax.matchAll(/\*(\w+)\*\s*\*\*As (\w+)\*\*/g)) {
        params.push({ name: m[1], type: m[2], optional: false, description: '' });
    }
    return params;
}

function signatureOf(name, params, returns) {
    const parts = params.map((p) => {
        const text = `${p.name}${p.type ? ` As ${p.type}` : ''}`;
        return p.optional ? `[${text}]` : text;
    });
    return `${name}(${parts.join(', ')})${returns ? ` As ${returns}` : ''}`;
}

/** Splits a member body into its parts: paragraphs, notes, syntax, definition lines. */
function parseBody(lines) {
    const paragraphs = [];
    const notes = [];
    let syntax = '';
    let current = [];
    let inFence = false;
    let inNote = false;
    const flush = () => {
        if (current.length) { paragraphs.push(current.join(' ')); current = []; }
    };
    for (const raw of lines) {
        const line = raw.trimEnd();
        if (line.startsWith('```')) { inFence = !inFence; flush(); continue; }
        if (inFence) { continue; }
        if (line.startsWith('> [!')) { inNote = true; flush(); continue; }
        if (inNote && line.startsWith('>')) { notes.push(line.replace(/^>\s?/, '')); continue; }
        inNote = false;
        if (line === '' || line.startsWith('{:')) { flush(); continue; }
        if (line.startsWith('Syntax:')) { flush(); syntax = line.slice('Syntax:'.length).trim(); continue; }
        if (/^\*\w+\*(?:,\s*\*\w+\*)*\s*$/.test(line) || line.startsWith(':')) { flush(); continue; }
        current.push(line.trim());
    }
    flush();
    return { paragraphs, note: plain(notes.join(' ')), syntax };
}

function parsePage(name, text) {
    const lines = text.replace(/\r\n/g, '\n').split('\n');
    // Class description: the first paragraph after the title.
    let description = '';
    let started = false;
    let inFence = false;
    for (const line of lines) {
        if (line.startsWith('```')) { inFence = !inFence; continue; }
        if (inFence) { continue; }
        if (/^# /.test(line)) { started = true; continue; }
        if (!started || line.trim() === '' || line.startsWith('{:') || line.startsWith('* TOC')) { continue; }
        description = plain(line);
        break;
    }
    const sections = { Properties: 'property', Methods: 'method', Events: 'event' };
    const members = { property: [], method: [], event: [] };
    const excluded = [];
    let kind;
    let member;
    let body = [];
    const finish = () => {
        if (!member) { return; }
        const parsed = parseBody(body);
        const all = `${parsed.paragraphs.join(' ')} ${parsed.note}`;
        if (TWINBASIC_ONLY.test(all)) {
            excluded.push({ name: member.name, kind: member.kind, reason: 'twinBASIC addition, absent from VB6' });
            member = undefined;
            return;
        }
        if (!inVb6Reference(member.kind, member.name)) {
            excluded.push({ name: member.name, kind: member.kind, reason: 'no page in the VB6 Language Reference' });
            member = undefined;
            return;
        }
        const out = { name: member.name, kind: member.kind, description: plain(parsed.paragraphs[0] ?? '') };
        if (RESERVED.test(parsed.note)) {
            out.reserved = true;
            out.reservedNote = parsed.note;
            if (!out.description) { out.description = parsed.note; }
        }
        const bodyText = body.join('\n');
        if (member.kind === 'property') {
            out.type = declaredTypeOf(bodyText, member.name);
            const access = accessOf(bodyText);
            if (access) { out.access = access; }
        } else if (member.kind === 'method') {
            const params = definitionParams(body.map((l) => l.trimEnd()));
            const ret = bodyText.match(/Returns (?:a |an |the )?\*\*([\w.]+)\*\*/);
            const returns = ret ? ret[1].replace(/^stdole\./, '') : '';
            if (returns) { out.returns = returns; }
            out.signature = signatureOf(member.name, params, returns);
            if (params.length) { out.parameters = params; }
        } else {
            const params = parsed.syntax ? syntaxParams(parsed.syntax) : [];
            out.signature = signatureOf(member.name, params, '');
            if (params.length) { out.parameters = params; }
        }
        members[member.kind].push(out);
        member = undefined;
    };
    for (const line of lines) {
        const section = line.match(/^## (\w+)\s*$/);
        if (section) {
            finish();
            kind = sections[section[1]];
            body = [];
            continue;
        }
        const head = line.match(/^### (\w+)\s*$/);
        if (head && kind) {
            finish();
            member = { name: head[1], kind };
            body = [];
            continue;
        }
        if (member) { body.push(line); }
    }
    finish();
    return { description, members, excluded };
}

function sourceOf(entry) {
    return `github.com/twinbasic/documentation@${commit.slice(0, 12)} docs/Reference/Default/VB/${entry}/index.md`;
}

/**
 * The Forms collection has no page of its own; the Global page describes it
 * in a section, and VB6's reference documents it as "Forms Collection" with
 * Count and Item. Carried as a class so `Forms` chains to it rather than to
 * a single Form.
 */
function formsCollection(globalText) {
    const section = globalText.replace(/\r\n/g, '\n').match(/^## Forms collection\n([\s\S]*?)(?=^## )/m);
    const description = section ? plain(section[1].split('\n').find((l) => l.trim() && !l.startsWith('{:')) ?? '') : '';
    return {
        name: 'Forms',
        kind: 'Class',
        guid: '',
        library: 'VB (twinBASIC documentation)',
        libraryId: 'VB',
        source: `${sourceOf('Global')} (section "Forms collection")`,
        description: description || 'The collection of currently loaded forms.',
        remarks: '',
        example: '',
        properties: [
            { name: 'Count', kind: 'property', description: 'The number of loaded forms.', type: 'Long', access: 'read-only' },
            { name: 'Item', kind: 'property', description: 'A loaded form, by zero-based position or by name.', type: 'Form', access: 'read-only' },
        ],
        methods: [],
        events: [],
    };
}

const written = [];
const dumps = [];
for (const entry of fs.readdirSync(pagesDir).sort()) {
    const page = path.join(pagesDir, entry, 'index.md');
    if (!fs.existsSync(page) || TWINBASIC_ONLY_CLASSES.has(entry)) { continue; }
    const text = fs.readFileSync(page, 'utf8');
    const { description, members, excluded } = parsePage(entry, text);
    if (entry === 'Global') {
        dumps.push(formsCollection(text));
        const forms = members.property.find((m) => m.name === 'Forms');
        if (forms) { forms.type = 'Forms'; }
    }
    const dump = {
        name: entry,
        kind: 'Class',
        guid: '',
        library: 'VB (twinBASIC documentation)',
        libraryId: 'VB',
        source: sourceOf(entry),
        description,
        remarks: entry === 'OLE' ? 'twinBASIC documents this class as a VB6 compatibility stub, mostly unimplemented; the member list is VB6\'s.' : '',
        example: '',
        properties: members.property,
        methods: members.method,
        events: members.event,
    };
    if (excluded.length) { dump.excludedMembers = excluded; }
    dumps.push(dump);
    written.push([entry, members.property.length, members.method.length, members.event.length, excluded.length,
        members.property.concat(members.method, members.event).filter((m) => m.reserved).length]);
}
for (const dump of dumps) {
    fs.writeFileSync(path.join(outDir, `${dump.name}.json`), `${JSON.stringify(dump, null, 2)}\n`, 'utf8');
}
if (reference) {
    const crossread = dumps
        .filter((d) => d.excludedMembers?.some((m) => m.reason.startsWith('no page')))
        .map((d) => `${d.name}: ${d.excludedMembers.filter((m) => m.reason.startsWith('no page')).map((m) => m.name).join(', ')}`);
    console.log(`Cross-read against ${reference.source} (${reference.fetchedAt}); left out for having no VB6 reference page:`);
    for (const line of crossread) { console.log(`  ${line}`); }
} else {
    console.log('No reference/vb6/vb6-reference-names.json: members were not cross-read against the VB6 Language Reference.');
}
for (const [name, p, m, e, x, r] of written) {
    console.log(`${name.padEnd(14)} ${String(p).padStart(3)} properties ${String(m).padStart(3)} methods ${String(e).padStart(3)} events  (${x} twinBASIC-only excluded, ${r} reserved)`);
}
console.log(`Wrote ${written.length} VB classes to ${path.relative(root, outDir)} from twinbasic/documentation@${commit.slice(0, 12)}.`);
