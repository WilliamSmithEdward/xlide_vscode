#!/usr/bin/env node
// Fetches the table of contents of Microsoft's archived Visual Basic 6.0
// Language Reference and records every property, method, event and object
// name it documents, as reference/vb6/vb6-reference-names.json. The
// transcriber (scripts/transcribe-vb6-docs.mjs) reads that file to cross-read
// twinBASIC's VB package pages: a member with no page anywhere in the VB6
// reference is not VB6, whatever the page says.
//
// This is a name-level check. The archive keeps the reference pages but not
// the "Applies To" object lists behind them, so which objects a documented
// name belongs to cannot be read from Microsoft; per-object membership rests
// on twinBASIC's pages, which state VB6 compatibility per control.
//
//   node scripts/fetch-vb6-reference-names.mjs [toc.json path]

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const TOC_URL = 'https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/toc.json';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outPath = path.join(root, 'reference', 'vb6', 'vb6-reference-names.json');

const local = process.argv[2];
const toc = local
    ? JSON.parse(fs.readFileSync(local, 'utf8'))
    : await (await fetch(TOC_URL)).json();
const items = Array.isArray(toc) ? toc : toc.items;

function* walk(nodes, trail) {
    for (const node of nodes) {
        yield [node, trail];
        if (Array.isArray(node.children)) {
            yield* walk(node.children, [...trail, node.toc_title ?? '']);
        }
    }
}

// A page's kind is its title's suffix, wherever the page sits: the Language
// Reference holds most, and the Controls Reference and Component Tools
// Guide hold the rest (the OLE drag-and-drop events, a scroll bar's Scroll,
// a UserControl's AccessKeyPress). "Left, Top Properties" names two members
// on one page; "OLE Container Control" is one object; "Action Property
// (CommonDialog)" carries a qualifier.
const KINDS = [
    [/\s+Propert(y|ies)\b/, 'properties'],
    [/\s+Methods?\b/, 'methods'],
    [/\s+Events?\b/, 'events'],
    [/\s+Functions?\b/, 'functions'],
    [/\s+Statements?\b/, 'statements'],
    [/\s+(Object|Container Control|Controls?|Collection)\b/, 'objects'],
];
const SUFFIX = /\s+(Properties|Property|Methods|Method|Events|Event|Object|Container Control|Controls|Control|Collection|Statements|Statement|Functions|Function)\b.*$/;
const names = Object.fromEntries(KINDS.map(([, kind]) => [kind, new Set()]));
const titles = Object.fromEntries(KINDS.map(([, kind]) => [kind, []]));
let nodeCount = 0;
for (const [node] of walk(items, [])) {
    nodeCount += 1;
    const title = String(node.toc_title ?? '').trim();
    const kind = KINDS.find(([pattern]) => pattern.test(title.replace(/\s*\(.*$/, '')))?.[1];
    if (!kind || !node.href || title.length <= 1) {
        continue;
    }
    titles[kind].push(title);
    const bare = title.replace(/\s*\(.*$/, '').replace(SUFFIX, '');
    for (const name of bare.split(/,\s*/)) {
        if (/^[A-Za-z_]\w*$/.test(name)) {
            names[kind].add(name);
        }
    }
}

const out = {
    source: local ? path.basename(local) : TOC_URL,
    fetchedAt: new Date().toISOString().slice(0, 10),
    tocNodes: nodeCount,
    pages: Object.fromEntries(Object.entries(titles).map(([k, v]) => [k, v.length])),
    names: Object.fromEntries(Object.entries(names).map(([k, v]) => [k, [...v].sort((a, b) => a.localeCompare(b))])),
};
fs.mkdirSync(path.dirname(outPath), { recursive: true });
fs.writeFileSync(outPath, `${JSON.stringify(out, null, 2)}\n`, 'utf8');
console.log(
    `Wrote ${path.relative(root, outPath)}: ${out.names.properties.length} property, ${out.names.methods.length} method, `
    + `${out.names.events.length} event, ${out.names.functions.length} function, ${out.names.statements.length} statement `
    + `and ${out.names.objects.length} object names from ${nodeCount} TOC nodes.`,
);
